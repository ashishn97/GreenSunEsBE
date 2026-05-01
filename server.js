require('dotenv').config();

const express = require('express');
const cors = require('cors');
const mongoose = require('mongoose');
const fs = require('fs');
const path = require('path');
const https = require('https');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const Handlebars = require('handlebars');
const puppeteer = require('puppeteer');

const PORT = process.env.PORT || 3000;
const MONGO_URI = process.env.MONGO_URI;
const GOOGLE_SHEETS_WEBHOOK_URL = process.env.GOOGLE_SHEETS_WEBHOOK_URL;
const ALLOWED_ORIGIN = process.env.ALLOWED_ORIGIN || '*';
const COUNTER_INITIAL_VALUE = Number.parseInt(process.env.COUNTER_INITIAL_VALUE || '1', 10);

const TEMPLATES_DIR = path.join(__dirname, 'templates');
const PDF_TEMPLATES_DIR = path.join(TEMPLATES_DIR, 'pdf');

if (!MONGO_URI) {
  throw new Error('MONGO_URI is required');
}

const app = express();
const allowedOrigins = ALLOWED_ORIGIN.split(',').map(origin => origin.trim()).filter(Boolean);

app.use(cors({
  origin(origin, callback) {
    if (ALLOWED_ORIGIN === '*' || !origin || allowedOrigins.includes(origin)) {
      callback(null, true);
      return;
    }
    callback(new Error('Not allowed by CORS'));
  }
}));
app.use(express.json({ limit: '10mb' }));

const QuoteSchema = new mongoose.Schema({
  ref_no: { type: String, unique: true, required: true },
  date: String,
  client_name: String,
  client_number: String,
  vendor_name: String,
  type: String,
  kw: String,
  base_cost: String,
  final_amount: String,
  sheet_logged: { type: Boolean, default: false }
}, { timestamps: true });

const CounterSchema = new mongoose.Schema({
  key: { type: String, unique: true, required: true },
  count: { type: Number, required: true, default: COUNTER_INITIAL_VALUE }
}, { timestamps: true });

const Quote = mongoose.model('Quote', QuoteSchema);
const Counter = mongoose.model('Counter', CounterSchema);

let browserPromise = null;
let pdfQueue = Promise.resolve();

Handlebars.registerHelper('inc', value => Number(value) + 1);

function postJson(url, payload) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify(payload);
    const urlObj = new URL(url);
    const req = https.request({
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(body)
      }
    }, (res) => {
      let responseBody = '';
      res.on('data', chunk => { responseBody += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 400) {
          reject(new Error(`Google Sheets returned ${res.statusCode}: ${responseBody}`));
          return;
        }
        resolve(responseBody);
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

async function getCurrentCount() {
  const initialCount = Number.isFinite(COUNTER_INITIAL_VALUE) ? COUNTER_INITIAL_VALUE : 1;
  const counter = await Counter.findOneAndUpdate(
    { key: 'quotation' },
    { $setOnInsert: { count: initialCount } },
    { upsert: true, new: true }
  );
  return counter.count;
}

async function incrementCurrentCount() {
  const initialCount = Number.isFinite(COUNTER_INITIAL_VALUE) ? COUNTER_INITIAL_VALUE : 1;
  const counter = await Counter.findOneAndUpdate(
    { key: 'quotation' },
    { $inc: { count: 1 }, $setOnInsert: { count: initialCount } },
    { upsert: true, new: true }
  );
  return counter.count;
}

function getDocxTemplatePath(type) {
  if (type !== 'bank' && type !== 'client') throw new Error('Invalid quotation type');
  return path.join(TEMPLATES_DIR, type === 'bank' ? 'bank-quotation.docx' : 'client-quotation.docx');
}

function getPdfTemplatePath(type) {
  if (type !== 'bank' && type !== 'client') throw new Error('Invalid quotation type');
  return path.join(PDF_TEMPLATES_DIR, type === 'bank' ? 'bank.html' : 'client.html');
}

function getLogoDataUri() {
  const logoPath = path.join(PDF_TEMPLATES_DIR, 'logo_new.png');
  if (!fs.existsSync(logoPath)) return '';
  return `data:image/png;base64,${fs.readFileSync(logoPath).toString('base64')}`;
}

function generateDocxBuffer(type, data) {
  const content = fs.readFileSync(getDocxTemplatePath(type), 'binary');
  const zip = new PizZip(content);
  let xml = zip.files['word/document.xml'].asText();

  xml = xml.replace(/<a:ln([^>/]*)(\/?)>/g, (match, attrs, slash) => {
    let newAttrs = attrs;
    if (/w="\d+"/.test(newAttrs)) newAttrs = newAttrs.replace(/w="\d+"/, 'w="0"');
    else newAttrs += ' w="0"';
    return `<a:ln${newAttrs}${slash}>`;
  });

  xml = xml.replace(/<a:ln([^>]*)>([\s\S]*?)<\/a:ln>/g, (match, attrs, inner) => {
    if (inner.includes('<a:ln')) return match;
    const fixedInner = inner
      .replace(/<a:solidFill[\s\S]*?<\/a:solidFill>/g, '')
      .replace(/<a:solidFill[^>]*\/>/g, '');
    return `<a:ln${attrs}>${fixedInner}</a:ln>`;
  });

  zip.file('word/document.xml', xml);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: '{{', end: '}}' },
    nullGetter: () => ''
  });

  try {
    doc.render(data);
  } catch (err) {
    const message = err.properties && Array.isArray(err.properties.errors)
      ? err.properties.errors.map(e => e.message).join('; ')
      : err.message;
    throw new Error('Template render failed: ' + message);
  }

  return doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });
}

function getSafeFilename(clientName, ext) {
  const safeName = (clientName || '')
    .replace(/[^a-zA-Z0-9 .\-]/g, '')
    .trim()
    .replace(/ /g, '_');
  return safeName ? `${safeName}_GSE_Quotation.${ext}` : `GSE_Quotation_Draft.${ext}`;
}

function validateDownloadPayload(type, data) {
  if (type !== 'bank' && type !== 'client') throw new Error('Invalid quotation type');
  if (!data || typeof data !== 'object') throw new Error('Missing quotation data');
  if (!data.ref_no || !data.client_name || !data.kw || !data.base_cost) {
    throw new Error('Missing required quotation fields');
  }
}

function buildQuoteDocument(type, data) {
  return {
    ref_no: data.ref_no,
    date: data.date,
    client_name: data.client_name,
    client_number: data.client_number,
    vendor_name: data.vendor_name,
    type,
    kw: data.kw,
    base_cost: data.base_cost,
    final_amount: data.final_amount
  };
}

async function saveQuoteOnce(type, data) {
  try {
    const quote = await Quote.create(buildQuoteDocument(type, data));
    await incrementCurrentCount();
    return { quote, created: true };
  } catch (err) {
    if (err && err.code === 11000) {
      const quote = await Quote.findOne({ ref_no: data.ref_no });
      return { quote, created: false };
    }
    throw err;
  }
}

function buildSheetPayload(type, data) {
  return {
    ref_no: data.ref_no,
    client_name: data.client_name,
    client_number: data.client_number,
    vendor_name: data.vendor_name,
    kw: data.kw,
    base_cost: data.base_cost,
    final_amount: data.final_amount,
    type
  };
}

async function ensureSheetLogged(quote, type, data) {
  if (!GOOGLE_SHEETS_WEBHOOK_URL) {
    console.warn('GOOGLE_SHEETS_WEBHOOK_URL is not configured');
    return;
  }
  if (!quote || quote.sheet_logged) return;
  await postJson(GOOGLE_SHEETS_WEBHOOK_URL, buildSheetPayload(type, data));
  await Quote.updateOne({ _id: quote._id }, { $set: { sheet_logged: true } });
}

async function saveAndLogQuote(type, data) {
  const { quote } = await saveQuoteOnce(type, data);
  await ensureSheetLogged(quote, type, data);
}

function buildMaterials(data) {
  const components = [
    'PV Modules',
    'Grid-Tie Inverter',
    'Structure',
    'Meter',
    'DC Wire',
    'AC Wire',
    'Earthing Wire',
    'Earthing Material',
    'Lighting Arrestor',
    'Fitting Accessories',
    'ACDB / DCDB / MCB Box / Busbar / Panel Box'
  ];
  const warranties = ['25', '8 - 10', '10', '1', '2', '2', 'N/A', '5', 'N/A', 'N/A', 'As Per Standard'];
  const qtyDefaults = ['10 No.', '1 Nos.', 'As Required', '1 set', 'As Required', 'As Required', 'As Required', '1 Set', '1 No.', 'As Required', '1 Set'];

  return components.map((component, index) => {
    const row = index + 1;
    return {
      no: row,
      component,
      spec: data[`spec_${row}`] || (row === 2 ? `Capacity of ${data.kw || ''} kW` : ''),
      company: data[`company_${row}`] || '',
      warranty: warranties[index],
      qty: data[`qty_${row}`] || qtyDefaults[index]
    };
  });
}

function buildPdfViewModel(type, data) {
  return {
    ...data,
    quotationTypeLabel: type === 'bank' ? 'Bank Quotation' : 'Client Quotation',
    isClient: type === 'client',
    logoDataUri: getLogoDataUri(),
    materials: buildMaterials(data),
    discount: data.Discount_amount || '0',
    central_subsidy: data.central_subsidy || '78,000',
    state_subsidy: data.state_subsidy || '17,000'
  };
}

async function getBrowser() {
  if (!browserPromise) {
    browserPromise = puppeteer.launch({
      headless: 'new',
      args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--font-render-hinting=none']
    });
  }
  return browserPromise;
}

function enqueuePdfJob(task) {
  const job = pdfQueue.then(task, task);
  pdfQueue = job.catch(() => {});
  return job;
}

async function generatePdfBuffer(type, data) {
  return enqueuePdfJob(async () => {
    const template = Handlebars.compile(fs.readFileSync(getPdfTemplatePath(type), 'utf-8'));
    const html = template({
      ...buildPdfViewModel(type, data),
      css: fs.readFileSync(path.join(PDF_TEMPLATES_DIR, 'quotation.css'), 'utf-8')
    });

    const browser = await getBrowser();
    const page = await browser.newPage();
    try {
      await page.setContent(html, { waitUntil: 'networkidle0' });
      await page.emulateMediaType('screen');
      const pdf = await page.pdf({
        format: 'A4',
        printBackground: true,
        preferCSSPageSize: true,
        margin: { top: '0mm', right: '0mm', bottom: '0mm', left: '0mm' }
      });
      return Buffer.from(pdf);
    } finally {
      await page.close();
    }
  });
}

app.get('/health', (req, res) => {
  res.status(200).json({ ok: true });
});

app.get('/api/next-count', async (req, res) => {
  try {
    const count = await getCurrentCount();
    res.json({ count });
  } catch (err) {
    console.error('Next count failed:', err.message);
    res.status(500).json({ error: 'Unable to get next count' });
  }
});

app.post('/api/increment-count', async (req, res) => {
  try {
    const count = await incrementCurrentCount();
    res.json({ count });
  } catch (err) {
    console.error('Increment count failed:', err.message);
    res.status(500).json({ error: 'Unable to increment count' });
  }
});

app.post('/api/download/docx', async (req, res) => {
  try {
    const { type, data } = req.body;
    validateDownloadPayload(type, data);
    await saveAndLogQuote(type, data);
    const docxBuffer = generateDocxBuffer(type, data);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${getSafeFilename(data.client_name, 'docx')}"`);
    res.send(docxBuffer);
  } catch (err) {
    console.error('DOCX generation failed:', err.message);
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/download/pdf', async (req, res) => {
  try {
    const { type, data } = req.body;
    validateDownloadPayload(type, data);
    await saveAndLogQuote(type, data);
    const pdfBuffer = await generatePdfBuffer(type, data);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${getSafeFilename(data.client_name, 'pdf')}"`);
    res.send(pdfBuffer);
  } catch (err) {
    console.error('PDF generation failed:', err.message);
    res.status(500).json({ error: err.message });
  }
});

async function shutdown() {
  try {
    if (browserPromise) {
      const browser = await browserPromise;
      await browser.close();
    }
    await mongoose.connection.close();
  } finally {
    process.exit(0);
  }
}

process.on('SIGINT', shutdown);
process.on('SIGTERM', shutdown);

mongoose.connect(MONGO_URI, { serverSelectionTimeoutMS: 10000 })
  .then(() => {
    console.log('MongoDB connected');
    app.listen(PORT, '0.0.0.0', () => {
      console.log(`Server running on port ${PORT}`);
    });
  })
  .catch((err) => {
    console.error('MongoDB connection failed:', err.message);
    process.exit(1);
  });