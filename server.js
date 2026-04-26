
mongoose.connect(process.env.MONGO_URI)
  .then(() => console.log("MongoDB Connected"))
  .catch(err => console.error("MongoDB Error:", err));
const express = require('express');
const cors = require('cors');
const mongoose = require('mongoose');
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const { spawnSync } = require('child_process');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const COUNTER_FILE = path.join(__dirname, 'counter.json');
const TEMPLATES_DIR = path.join(__dirname, 'templates');
const OUTPUT_DIR = path.join(__dirname, 'output'); // Temporary dir for PDF conversion

if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR);
}

const QuoteSchema = new mongoose.Schema({
  ref_no: String,
  date: String,
  client_name: String,
  client_number: String,
  vendor_name: String,
  type: String,
  kw: String,
  base_cost: String,
  final_amount: String,
}, { timestamps: true });

const Quote = mongoose.model("Quote", QuoteSchema);

const sofficePath = process.env.LIBREOFFICE_PATH || null

// Ensure counter file exists
if (!fs.existsSync(COUNTER_FILE)) {
  fs.writeFileSync(COUNTER_FILE, JSON.stringify({ count: 1 }));
}

// Helpers
function getCount() {
  const data = JSON.parse(fs.readFileSync(COUNTER_FILE, 'utf-8'));
  return data.count;
}

function generateDocxBlob(type, data) {
  const templatePath = path.join(
    TEMPLATES_DIR,
    type === 'bank' ? 'bank-quotation.docx' : 'client-quotation.docx'
  );
  const content = fs.readFileSync(templatePath, 'binary');
  const zip = new PizZip(content);

  // Fix background image border in XML before rendering
  let xml = zip.files['word/document.xml'].asText();

  // Zero out a:ln width to remove image borders.
  // Handles both open tags <a:ln ...> and self-closing <a:ln .../>
  xml = xml.replace(/<a:ln([^>/]*)(\/?)>/g, (match, attrs, slash) => {
    let newAttrs = attrs;
    if (/w="\d+"/.test(newAttrs)) {
      newAttrs = newAttrs.replace(/w="\d+"/, 'w="0"');
    } else {
      newAttrs += ' w="0"';
    }
    return `<a:ln${newAttrs}${slash}>`;
  });

  // Remove <a:solidFill> inside <a:ln> blocks (coloured borders)
  xml = xml.replace(/<a:ln([^>]*)>([\s\S]*?)<\/a:ln>/g, (match, attrs, inner) => {
    // Don't recurse into nested a:ln
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
    nullGetter: () => '' // render missing keys as empty string instead of throwing
  });

  try {
    doc.render(data);
  } catch (err) {
    // Extract structured error messages from docxtemplater
    const message =
      err.properties && Array.isArray(err.properties.errors)
        ? err.properties.errors.map(e => e.message).join('; ')
        : err.message;
    console.error('Docxtemplater render error:', message);
    throw new Error('Template render failed: ' + message);
  }

  return doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });
}

function getSafeFilename(clientName, ext) {
  const safeName = (clientName || '')
    .replace(/[^a-zA-Z0-9 .\-]/g, '')
    .trim()
    .replace(/ /g, '_');
  if (!safeName) return `GSE_Quotation_Draft.${ext}`;
  return `${safeName}_GSE_Quotation.${ext}`;
}

// Endpoints
app.get('/api/next-count', (req, res) => {
  res.json({ count: getCount() });
});

app.post('/api/increment-count', (req, res) => {
  let count = getCount();
  count++;
  fs.writeFileSync(COUNTER_FILE, JSON.stringify({ count }));
  res.json({ count });
});

app.post('/api/download/docx', async (req, res) => {
  try {
    const { type, data } = req.body;
    await Quote.create({
      ref_no: data.ref_no,
      client_name: data.client_name,
      client_number: data.client_number,
      vendor_name: data.vendor_name,
      type: data.type,
      kw: data.kw,
      base_cost: data.base_cost,
      final_amount: data.final_amount,
      date: data.date,
    });
    const buf = generateDocxBlob(type, data);
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    );
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="${getSafeFilename(data.client_name, 'docx')}"`
    );
    res.send(buf);
  } catch (error) {
    console.error('Error generating DOCX:', error.message);
    res.status(500).json({ error: error.message });
  }
});

app.post('/api/download/pdf', async (req, res) => {
  let tempDocxPath = null;
  let pdfPath = null;
  try {
    const { type, data } = req.body;
    await Quote.create({
      ref_no: data.ref_no,
      client_name: data.client_name,
      client_number: data.client_number,
      vendor_name: data.vendor_name,
      type: data.type,
      kw: data.kw,
      base_cost: data.base_cost,
      final_amount: data.final_amount,
      date: data.date,
    });
    const docxBuf = generateDocxBlob(type, data);

    // Save temp docx
    tempDocxPath = path.join(OUTPUT_DIR, `temp_${Date.now()}.docx`);
    fs.writeFileSync(tempDocxPath, docxBuf);

    // Convert to PDF using LibreOffice
    const result = spawnSync(
      sofficePath,
      ['--headless', '--convert-to', 'pdf', '--outdir', OUTPUT_DIR, tempDocxPath],
      { timeout: 60000 }
    );

    const stderr = result.stderr ? result.stderr.toString().trim() : '';
    const stdout = result.stdout ? result.stdout.toString().trim() : '';
    if (stdout) console.log('LibreOffice stdout:', stdout);
    if (stderr) console.error('LibreOffice stderr:', stderr);

    if (result.status !== 0) {
      return res.status(500).json({
        error: 'LibreOffice conversion failed.',
        detail: stderr || 'Non-zero exit code: ' + result.status
      });
    }

    pdfPath = tempDocxPath.replace(/\.docx$/, '.pdf');
    if (!fs.existsSync(pdfPath)) {
      console.error('LibreOffice did not produce a PDF file.');
      return res.status(500).json({ error: 'PDF file not created by LibreOffice.' });
    }

    const pdfBuf = fs.readFileSync(pdfPath);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="${getSafeFilename(data.client_name, 'pdf')}"`
    );
    res.send(pdfBuf);
  } catch (error) {
    console.error('Error generating PDF:', error.message);
    res.status(500).json({ error: error.message });
  } finally {
    // Cleanup temp files regardless of success/failure
    try { if (tempDocxPath && fs.existsSync(tempDocxPath)) fs.unlinkSync(tempDocxPath); } catch (e) {}
    try { if (pdfPath && fs.existsSync(pdfPath)) fs.unlinkSync(pdfPath); } catch (e) {}
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
