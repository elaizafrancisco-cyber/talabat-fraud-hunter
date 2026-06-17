const dns = require('dns');
dns.setDefaultResultOrder('ipv4first');

const express = require('express');
const multer = require('multer');
const nodemailer = require('nodemailer');
const path = require('path');
const fs = require('fs');
const { runAnalysis } = require('./fraud_engine');

const app = express();
const PORT = process.env.PORT || 3000;

const GDRIVE_ROOT_DEFAULT = 'G:/Shared drives/AR Team (Talabat UAE Finance) - Shared Drive/Claude/Fraud';
const DATA_ROOT = process.env.DATA_ROOT || (fs.existsSync(GDRIVE_ROOT_DEFAULT) ? GDRIVE_ROOT_DEFAULT : path.join(__dirname, 'data'));
const INPUT_DIR = path.join(DATA_ROOT, 'Input');
const OUTPUT_DIR = path.join(DATA_ROOT, 'Output');
const APP_DIR = __dirname;

[INPUT_DIR, OUTPUT_DIR].forEach(dir => {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});
['AX365', 'Branch Inquiry', 'Checkout', 'Detailed Reports'].forEach(folder => {
  const dir = path.join(INPUT_DIR, folder);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

let latestResults = null;
let latestOutputPath = null;
let analysisRunning = false;
let analysisLog = [];

const uploadMetaPath = path.join(DATA_ROOT, 'upload_meta.json');
function loadUploadMeta() {
  try { return JSON.parse(fs.readFileSync(uploadMetaPath, 'utf-8')); }
  catch { return {}; }
}
function saveUploadMeta(meta) {
  fs.writeFileSync(uploadMetaPath, JSON.stringify(meta, null, 2));
}

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const folder = req.params.folder;
    const validFolders = { 'ax365': 'AX365', 'branch-inquiry': 'Branch Inquiry', 'checkout': 'Checkout', 'detailed-reports': 'Detailed Reports' };
    const target = validFolders[folder];
    if (!target) return cb(new Error('Invalid folder'));
    const dir = path.join(INPUT_DIR, target);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    cb(null, dir);
  },
  filename: (req, file, cb) => cb(null, file.originalname),
});
const upload = multer({ storage, limits: { fileSize: 500 * 1024 * 1024 } });

app.post('/api/upload/:folder', upload.array('files', 20), (req, res) => {
  const uploaderName = req.body.uploaderName || 'Unknown';
  const meta = loadUploadMeta();
  const now = new Date().toISOString();
  const uploaded = req.files.map(f => {
    const key = `${req.params.folder}/${f.originalname}`;
    meta[key] = { dateUploaded: now, dateModified: now, uploaderName, size: f.size };
    return f.originalname;
  });
  saveUploadMeta(meta);
  res.json({ success: true, files: uploaded, message: `${uploaded.length} file(s) uploaded` });
});

app.get('/api/files', (req, res) => {
  const folders = ['AX365', 'Branch Inquiry', 'Checkout', 'Detailed Reports'];
  const folderKeyMap = { 'AX365': 'ax365', 'Branch Inquiry': 'branch-inquiry', 'Checkout': 'checkout', 'Detailed Reports': 'detailed-reports' };
  const result = {};
  const meta = loadUploadMeta();
  for (const folder of folders) {
    const dir = path.join(INPUT_DIR, folder);
    if (fs.existsSync(dir)) {
      result[folder] = fs.readdirSync(dir)
        .filter(f => f.endsWith('.xlsx') || f.endsWith('.csv'))
        .map(f => {
          const stat = fs.statSync(path.join(dir, f));
          const metaKey = `${folderKeyMap[folder]}/${f}`;
          const fileMeta = meta[metaKey] || {};
          return {
            name: f, size: stat.size,
            dateUploaded: fileMeta.dateUploaded || stat.birthtime.toISOString(),
            dateModified: fileMeta.dateModified || stat.mtime.toISOString(),
            uploaderName: fileMeta.uploaderName || '-',
          };
        });
    } else {
      result[folder] = [];
    }
  }
  result._source = fs.existsSync(GDRIVE_ROOT_DEFAULT) ? 'gdrive' : 'cloud';
  res.json(result);
});

app.post('/api/run', async (req, res) => {
  if (analysisRunning) return res.status(409).json({ error: 'Analysis already running' });
  analysisRunning = true;
  analysisLog = [];
  const origEmit = console.log;
  console.log = (...args) => { const msg = args.join(' '); analysisLog.push(msg); origEmit(...args); };

  try {
    const ts = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
    const outputPath = path.join(OUTPUT_DIR, `Fraud_Analysis_${ts}.xlsx`);
    latestOutputPath = outputPath;
    latestResults = await runAnalysis(INPUT_DIR, outputPath);
    fs.writeFileSync(path.join(OUTPUT_DIR, 'dashboard_data.json'), JSON.stringify(latestResults, null, 2));
    const warnings = latestResults.fileErrors || null;
    res.json({ success: true, data: latestResults, log: analysisLog, warnings });
  } catch (err) {
    res.status(500).json({ error: err.message, log: analysisLog });
  } finally {
    console.log = origEmit;
    analysisRunning = false;
  }
});

app.get('/api/results', (req, res) => {
  if (latestResults) return res.json(latestResults);
  const gdriveJson = path.join(OUTPUT_DIR, 'dashboard_data.json');
  const localJson = path.join(APP_DIR, 'dashboard_data.json');
  const jsonPath = fs.existsSync(gdriveJson) ? gdriveJson : (fs.existsSync(localJson) ? localJson : null);
  if (jsonPath) {
    latestResults = JSON.parse(fs.readFileSync(jsonPath, 'utf-8'));
    return res.json(latestResults);
  }
  res.status(404).json({ error: 'No analysis results yet. Click Run Analysis.' });
});

app.get('/api/status', (req, res) => {
  res.json({ running: analysisRunning, log: analysisLog });
});

app.get('/api/download/excel', (req, res) => {
  let file = latestOutputPath;
  if (!file || !fs.existsSync(file)) {
    const searchDirs = [OUTPUT_DIR, APP_DIR];
    for (const dir of searchDirs) {
      if (!fs.existsSync(dir)) continue;
      const files = fs.readdirSync(dir)
        .filter(f => f.startsWith('Fraud_Analysis_') && f.endsWith('.xlsx'))
        .sort().reverse();
      if (files.length) { file = path.join(dir, files[0]); break; }
    }
  }
  if (!file || !fs.existsSync(file)) return res.status(404).json({ error: 'No output file. Run analysis first.' });
  res.download(file, 'Fraud_Analysis_Output.xlsx');
});

app.post('/api/share/slack', async (req, res) => {
  const { channelId, message, token } = req.body;
  if (!token) return res.status(400).json({ error: 'Slack Bot Token is required. Set it in the dashboard settings.' });
  if (!channelId) return res.status(400).json({ error: 'Channel ID required' });

  try {
    const { WebClient } = require('@slack/web-api');
    const slack = new WebClient(token);
    let filePath = latestOutputPath;
    if (!filePath || !fs.existsSync(filePath)) {
      for (const dir of [OUTPUT_DIR, APP_DIR]) {
        if (!fs.existsSync(dir)) continue;
        const files = fs.readdirSync(dir).filter(f => f.startsWith('Fraud_Analysis_') && f.endsWith('.xlsx')).sort().reverse();
        if (files.length) { filePath = path.join(dir, files[0]); break; }
      }
    }
    if (!filePath || !fs.existsSync(filePath)) return res.status(404).json({ error: 'No output file to share' });

    if (message) {
      await slack.chat.postMessage({ channel: channelId, text: message });
    }
    const fileStream = fs.createReadStream(filePath);
    await slack.filesUploadV2({
      channel_id: channelId,
      file: fileStream,
      filename: 'Fraud_Analysis_Output.xlsx',
      title: 'Fraud Analysis Report',
      initial_comment: message || 'Fraud Analysis Report attached.',
    });
    res.json({ success: true, message: 'Shared to Slack successfully' });
  } catch (err) {
    res.status(500).json({ error: `Slack error: ${err.message}` });
  }
});

app.post('/api/share/email', async (req, res) => {
  const { from, to, cc, subject, body, smtpHost, smtpPort, smtpUser, smtpPass } = req.body;
  if (!to) return res.status(400).json({ error: 'TO field is required' });
  if (!smtpHost) return res.status(400).json({ error: 'SMTP settings required. Configure in dashboard settings.' });

  try {
    const transporter = nodemailer.createTransport({
      host: smtpHost, port: parseInt(smtpPort) || 587,
      secure: parseInt(smtpPort) === 465,
      auth: smtpUser ? { user: smtpUser, pass: smtpPass } : undefined,
    });

    const emailDataPath = path.join(OUTPUT_DIR, 'dashboard_data.json');
    let htmlBody = body || 'Dear Team,\n\nPlease verify the emails below that have suspicious transactions.\n\nRegards,\nUAE Talabat Finance';
    htmlBody = htmlBody.replace(/\n/g, '<br>');
    if (latestResults && latestResults.checkoutEmailTop20) {
      htmlBody += '<br><br><h3>Top 20 Suspicious Customer Emails</h3><table border="1" cellpadding="5" style="border-collapse:collapse;font-family:Arial;font-size:12px;">';
      htmlBody += '<tr style="background:#FF5900;color:white;"><th>Rank</th><th>Customer Email</th></tr>';
      latestResults.checkoutEmailTop20.forEach((e, i) => {
        htmlBody += `<tr><td style="text-align:center;">${i + 1}</td><td>${e.email}</td></tr>`;
      });
      htmlBody += '</table>';
    }

    await transporter.sendMail({
      from: from || smtpUser,
      to, cc: cc || undefined,
      subject: subject || 'Fraud Analysis - Top 20 Suspicious Emails',
      html: htmlBody,
    });
    res.json({ success: true, message: 'Email sent successfully' });
  } catch (err) {
    res.status(500).json({ error: `Email error: ${err.message}` });
  }
});

// Google Drive URLs for reference
const GDRIVE_OUTPUT_URL = 'https://drive.google.com/drive/folders/1ZWyk9K86N81G4Xe1iyB8Y6yvrbVz-6IX';
const GDRIVE_INPUT_URL = 'https://drive.google.com/drive/folders/13WyEDftGEHZjhbCRaXF298gRWbpuUJX8';

app.post('/api/export/gsheets', (req, res) => {
  let file = latestOutputPath;
  if (!file || !fs.existsSync(file)) {
    for (const dir of [OUTPUT_DIR, APP_DIR]) {
      if (!fs.existsSync(dir)) continue;
      const files = fs.readdirSync(dir)
        .filter(f => f.startsWith('Fraud_Analysis_') && f.endsWith('.xlsx'))
        .sort().reverse();
      if (files.length) { file = path.join(dir, files[0]); break; }
    }
  }
  if (!file || !fs.existsSync(file)) {
    return res.status(404).json({ error: 'No output file. Run analysis first.' });
  }
  const destName = path.basename(file);
  const destPath = path.join(OUTPUT_DIR, destName);
  try {
    if (path.resolve(file) !== path.resolve(destPath)) {
      fs.copyFileSync(file, destPath);
    }
    res.json({ success: true, message: `${destName} is available in Google Drive Output folder`, driveUrl: GDRIVE_OUTPUT_URL });
  } catch (err) {
    res.status(500).json({ error: `Drive export error: ${err.message}` });
  }
});

// Public / Guest access
const guestTokensPath = path.join(DATA_ROOT, 'guest_tokens.json');
function loadGuestTokens() {
  try { return JSON.parse(fs.readFileSync(guestTokensPath, 'utf-8')); }
  catch { return {}; }
}
function saveGuestTokens(tokens) {
  fs.writeFileSync(guestTokensPath, JSON.stringify(tokens, null, 2));
}

app.post('/api/public/generate', (req, res) => {
  const crypto = require('crypto');
  const token = crypto.randomBytes(16).toString('hex');
  const tokens = loadGuestTokens();
  tokens[token] = { createdAt: new Date().toISOString(), role: 'viewer' };
  saveGuestTokens(tokens);
  const proto = req.headers['x-forwarded-proto'] || req.protocol;
  const host = req.headers['x-forwarded-host'] || req.get('host');
  const url = `${proto}://${host}/guest/${token}`;
  res.json({ success: true, token, url });
});

app.get('/guest/:token', (req, res) => {
  const tokens = loadGuestTokens();
  if (!tokens[req.params.token]) {
    return res.status(403).send('<h2>Invalid or expired link</h2><p>This guest access link is not valid.</p>');
  }
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.get('/api/guest/verify/:token', (req, res) => {
  const tokens = loadGuestTokens();
  if (tokens[req.params.token]) {
    return res.json({ valid: true, role: 'viewer' });
  }
  res.json({ valid: false });
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`\n  Talabat Fraud Hunter Dashboard`);
  console.log(`  http://localhost:${PORT}`);
  console.log(`  Data root: ${DATA_ROOT}`);
  console.log(`  Press Ctrl+C to stop\n`);
});
