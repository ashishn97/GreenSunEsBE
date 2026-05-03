require('dotenv').config();
const crypto = require('crypto');

const express = require('express');
const cors = require('cors');
const mongoose = require('mongoose');
const fs = require('fs');
const fsp = require('fs/promises');
const path = require('path');
const os = require('os');
const https = require('https');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const { execFile } = require('child_process');
const { promisify } = require('util');

const execFileAsync = promisify(execFile);

const PORT = process.env.PORT || 10000;
const MONGO_URI = process.env.MONGO_URI;
const GOOGLE_SHEETS_WEBHOOK_URL = process.env.GOOGLE_SHEETS_WEBHOOK_URL;
const ALLOWED_ORIGIN = process.env.ALLOWED_ORIGIN || 'https://greensunenergyservices.co.in,https://www.greensunenergyservices.co.in';
const allowedOrigins = ALLOWED_ORIGIN.split(',').map(origin => origin.trim()).filter(Boolean);
const COUNTER_INITIAL_VALUE = Number.parseInt(process.env.COUNTER_INITIAL_VALUE || '1', 10);
const SOFFICE_PATH = process.env.LIBREOFFICE_PATH || 'soffice';
const GOOGLE_SHEETS_TIMEOUT_MS = Number.parseInt(process.env.GOOGLE_SHEETS_TIMEOUT_MS || '5000', 10);

const TEMPLATES_DIR = path.join(__dirname, 'templates');

if (!MONGO_URI) {
  throw new Error('MONGO_URI is required');
}

const app = express();
const corsOptions = {
  origin(origin, callback) {
    if (!origin) { callback(null, true); return; }

    // Always allow localhost / 127.0.0.1 for local development
    const isLocal = /^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/.test(origin);
    if (isLocal || allowedOrigins.includes(origin)) {
      callback(null, true);
      return;
    }

    console.warn(`CORS blocked origin: ${origin}`);
    callback(null, false);
  },
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'x-auth-token'],
  credentials: false,
  optionsSuccessStatus: 204
};

app.use(cors(corsOptions));
app.options('*', cors(corsOptions));
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
  sheet_logged: { type: Boolean, default: false },
  kw_number: { type: Number },
  final_amount_number: { type: Number },
  deleted_at: { type: Date, default: null }
}, { timestamps: true });

const HISTORY_ACCESS_CODE = process.env.HISTORY_ACCESS_CODE || '';

function hashToken(raw) {
  return crypto.createHash('sha256').update(raw).digest('hex');
}

const HistorySessionSchema = new mongoose.Schema({
  token_hash: { type: String, unique: true, required: true },
  name: { type: String, required: true },
  created_at: { type: Date, default: Date.now }
});
const HistorySession = mongoose.model('HistorySession', HistorySessionSchema);

const CounterSchema = new mongoose.Schema({
  key: { type: String, unique: true, required: true },
  count: { type: Number, required: true, default: COUNTER_INITIAL_VALUE }
}, { timestamps: true });

const Quote = mongoose.model('Quote', QuoteSchema);
const Counter = mongoose.model('Counter', CounterSchema);

let pdfQueue = Promise.resolve();

function postJson(url, payload) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify(payload);
    const urlObj = new URL(url);
    const req = https.request({
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      method: 'POST',
      timeout: GOOGLE_SHEETS_TIMEOUT_MS,
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
    req.on('timeout', () => req.destroy(new Error(`Google Sheets timeout after ${GOOGLE_SHEETS_TIMEOUT_MS}ms`)));
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
    { upsert: true, returnDocument: 'after' }
  );
  return counter.count;
}

async function incrementCurrentCount() {
  await getCurrentCount();

  const counter = await Counter.findOneAndUpdate(
    { key: 'quotation' },
    { $inc: { count: 1 } },
    { returnDocument: 'after' }
  );

  return counter.count;
}

function getDocxTemplatePath(type) {
  if (type !== 'bank' && type !== 'client') throw new Error('Invalid quotation type');
  return path.join(TEMPLATES_DIR, type === 'bank' ? 'bank-quotation.docx' : 'client-quotation.docx');
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

  xml = xml.replace(/<v:stroke[^>]*>/g, '');
  xml = xml.replace(/<\/v:stroke>/g, '');
  xml = xml.replace(/stroked="true"/g, 'stroked="false"');
  xml = xml.replace(/strokeweight="[^"]*"/g, 'strokeweight="0"');
  xml = xml.replace(/strokecolor="[^"]*"/g, '');
  xml = xml.replace(/style="[^"]*stroke:[^"]*"/g, 'style=""');

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

function enqueuePdfJob(task) {
  const job = pdfQueue.then(task, task);
  pdfQueue = job.catch(() => {});
  return job;
}

async function convertDocxToPdf(docxBuffer) {
  return enqueuePdfJob(async () => {
    const tempDir = await fsp.mkdtemp(path.join(os.tmpdir(), 'gse-pdf-'));
    const inputPath = path.join(tempDir, 'input.docx');
    const outputPath = path.join(tempDir, 'input.pdf');
    const profileDir = path.join(tempDir, 'lo-profile');

    try {
      await fsp.writeFile(inputPath, docxBuffer);
      await fsp.mkdir(profileDir, { recursive: true });

      console.log(`Starting LibreOffice PDF conversion: ${inputPath}`);
      const startedAt = Date.now();

      await execFileAsync(SOFFICE_PATH, [
        '--headless',
        '--nologo',
        '--nofirststartwizard',
        '--nodefault',
        '--nolockcheck',
        `-env:UserInstallation=file://${profileDir.replace(/\\/g, '/')}`,
        '--convert-to',
        'pdf',
        '--outdir',
        tempDir,
        inputPath
      ], {
        timeout: 90000,
        maxBuffer: 1024 * 1024 * 10
      });

      if (!fs.existsSync(outputPath)) {
        throw new Error('LibreOffice did not create PDF output');
      }

      const pdfBuffer = await fsp.readFile(outputPath);
      if (!pdfBuffer.length || pdfBuffer.slice(0, 4).toString() !== '%PDF') {
        throw new Error('LibreOffice produced an invalid PDF');
      }

      console.log(`LibreOffice PDF conversion completed in ${Date.now() - startedAt}ms`);
      return pdfBuffer;
    } catch (err) {
      console.error('LibreOffice PDF conversion failed:', err.message);
      throw new Error('PDF conversion failed');
    } finally {
      await fsp.rm(tempDir, { recursive: true, force: true }).catch(cleanupErr => {
        console.warn('PDF temp cleanup failed:', cleanupErr.message);
      });
    }
  });
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

function safeSheetValue(value) {
  return value === undefined || value === null ? '' : String(value).trim();
}

function buildSheetPayload(type, data) {
  return {
    ref_no: safeSheetValue(data.ref_no),
    client_name: safeSheetValue(data.client_name),
    client_number: safeSheetValue(data.client_number),
    vendor_name: safeSheetValue(data.vendor_name),
    kw: safeSheetValue(data.kw),
    base_cost: safeSheetValue(data.base_cost),
    final_amount: safeSheetValue(data.final_amount),
    type: safeSheetValue(type)
  };
}

async function ensureSheetLogged(quote, type, data) {
  if (!GOOGLE_SHEETS_WEBHOOK_URL) {
    console.warn('GOOGLE_SHEETS_WEBHOOK_URL is not configured');
    return;
  }
  if (!quote || quote.sheet_logged) return;

  try {
    const payload = buildSheetPayload(type, data);
    if (!payload.ref_no || !payload.client_name) throw new Error('Invalid Google Sheets payload');
    await postJson(GOOGLE_SHEETS_WEBHOOK_URL, payload);
    await Quote.updateOne({ _id: quote._id }, { $set: { sheet_logged: true } });
  } catch (err) {
    console.error('Google Sheets logging failed:', {
      ref_no: data && data.ref_no,
      message: err.message
    });
  }
}

async function saveAndLogQuote(type, data) {
  const { quote } = await saveQuoteOnce(type, data);
  setImmediate(() => {
    ensureSheetLogged(quote, type, data).catch(err => {
      console.error('Google Sheets background log failed:', {
        ref_no: data && data.ref_no,
        message: err.message
      });
    });
  });
}

// ── AUTH MIDDLEWARE ────────────────────────────────────────────────────────────
async function authHistory(req, res, next) {
  const raw = req.headers['x-auth-token'];
  if (!raw) return res.status(401).json({ error: 'Authentication required' });
  const hash = hashToken(raw);
  const session = await HistorySession.findOne({ token_hash: hash });
  if (!session) return res.status(401).json({ error: 'Invalid or expired token' });
  req.histUser = session.name;
  next();
}

// ── SYNC SHEETS HELPER ──────────────────────────────────────────────────────────
function syncSheetsAsync(payload) {
  if (!GOOGLE_SHEETS_WEBHOOK_URL) return;
  setImmediate(() => {
    postJson(GOOGLE_SHEETS_WEBHOOK_URL, payload).catch(err => {
      console.error('Sheets sync failed:', err.message);
    });
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
    const docxBuffer = generateDocxBuffer(type, data);
    const pdfBuffer = await convertDocxToPdf(docxBuffer);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${getSafeFilename(data.client_name, 'pdf')}"`);
    res.send(pdfBuffer);
  } catch (err) {
    console.error('PDF generation failed:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── POST /api/auth/login ───────────────────────────────────────────────────────
app.post('/api/auth/login', async (req, res) => {
  try {
    const { name, code } = req.body;
    if (!name || !name.trim()) return res.status(400).json({ error: 'Name is required' });
    if (!HISTORY_ACCESS_CODE) return res.status(500).json({ error: 'Server not configured' });
    if (!code || code.trim() !== HISTORY_ACCESS_CODE) {
      return res.status(401).json({ error: 'Invalid access code' });
    }
    const rawToken = crypto.randomBytes(32).toString('hex');
    const hash = hashToken(rawToken);
    await HistorySession.create({ token_hash: hash, name: name.trim() });
    res.json({ token: rawToken, name: name.trim() });
  } catch (err) {
    console.error('Login failed:', err.message);
    res.status(500).json({ error: 'Login failed' });
  }
});

// ── GET /api/quotations ────────────────────────────────────────────────────────
app.get('/api/quotations', authHistory, async (req, res) => {
  try {
    const page  = Math.max(1, parseInt(req.query.page  || '1', 10));
    const limit = Math.min(50, Math.max(1, parseInt(req.query.limit || '10', 10)));
    const { q, vendor, kw, type, sortBy, sortOrder } = req.query;

    const filter = { deleted_at: null };
    if (q)      filter.$or = [{ client_name: new RegExp(q, 'i') }, { ref_no: new RegExp(q, 'i') }, { client_number: new RegExp(q, 'i') }, { vendor_name: new RegExp(q, 'i') }];
    if (vendor) filter.vendor_name = new RegExp(vendor, 'i');
    if (kw)     filter.kw = new RegExp(`^${parseFloat(kw)}`, 'i');
    if (type)   filter.type = type;

    const allowedSortFields = ['createdAt', 'kw_number', 'final_amount_number'];
    const sortField = allowedSortFields.includes(sortBy) ? sortBy : 'createdAt';
    const order     = sortOrder === 'asc' ? 1 : -1;

    const [data, total] = await Promise.all([
      Quote.find(filter).sort({ [sortField]: order }).skip((page - 1) * limit).limit(limit).select('-__v'),
      Quote.countDocuments(filter)
    ]);

    res.json({ data, pagination: { page, limit, total, pages: Math.ceil(total / limit) } });
  } catch (err) {
    console.error('List quotations failed:', err.message);
    res.status(500).json({ error: 'Failed to retrieve quotations' });
  }
});

// ── PUT /api/quotations/:id ────────────────────────────────────────────────────
app.put('/api/quotations/:id', authHistory, async (req, res) => {
  try {
    const { id } = req.params;
    const allowed = ['client_name', 'vendor_name', 'kw', 'base_cost', 'final_amount', 'type'];
    const updates = {};
    for (const field of allowed) {
      if (req.body[field] !== undefined) updates[field] = String(req.body[field]).trim();
    }
    if (req.body.kw !== undefined) {
      const n = parseFloat(req.body.kw);
      if (!isNaN(n)) updates.kw_number = n;
    }
    if (req.body.final_amount !== undefined) {
      const n = parseFloat(String(req.body.final_amount).replace(/,/g, ''));
      if (!isNaN(n)) updates.final_amount_number = n;
    }
    if (Object.keys(updates).length === 0) return res.status(400).json({ error: 'No valid fields to update' });

    const quote = await Quote.findByIdAndUpdate(id, { $set: updates }, { new: true, runValidators: true });
    if (!quote) return res.status(404).json({ error: 'Quotation not found' });

    syncSheetsAsync({ action: 'update', ref_no: quote.ref_no, ...updates });
    res.json({ ok: true, data: quote });
  } catch (err) {
    console.error('Update quotation failed:', err.message);
    res.status(500).json({ error: 'Update failed' });
  }
});

// ── DELETE /api/quotations/:id ─────────────────────────────────────────────────
app.delete('/api/quotations/:id', authHistory, async (req, res) => {
  try {
    const quote = await Quote.findByIdAndUpdate(
      req.params.id,
      { $set: { deleted_at: new Date() } },
      { new: true }
    );
    if (!quote) return res.status(404).json({ error: 'Quotation not found' });
    syncSheetsAsync({ action: 'delete', ref_no: quote.ref_no });
    res.json({ ok: true });
  } catch (err) {
    console.error('Delete quotation failed:', err.message);
    res.status(500).json({ error: 'Delete failed' });
  }
});

async function shutdown() {
  try {
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
