const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const session = require('express-session');
const archiver = require('archiver');

const config = JSON.parse(
  fs.readFileSync(path.join(__dirname, 'config.json'), 'utf8')
);
const TEST_SHEETS_PATH = config.testSheetsPath || path.join(__dirname, 'Testsheets');
const PORT = config.port || 3456;
const USERS = config.auth?.users || { 'admin@example.com': 'admin123' };
const AUTH_DISABLED = !!config.auth?.disabled;

if (!fs.existsSync(TEST_SHEETS_PATH)) {
  fs.mkdirSync(TEST_SHEETS_PATH, { recursive: true });
}

const app = express();
app.use(cors({ origin: true, credentials: true }));
app.use(express.json());
app.use(session({
  secret: config.auth?.sessionSecret || 'test-sheet-manager-secret',
  resave: false,
  saveUninitialized: false,
  cookie: { httpOnly: true, secure: false },
}));

app.get('/login', (req, res) => {
  if (AUTH_DISABLED) return res.redirect('/');
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

app.use(express.static('public'));

const requireAuth = (req, res, next) => {
  if (AUTH_DISABLED || req.session?.email) return next();
  res.status(401).json({ error: 'Unauthorized' });
};

app.post('/api/login', (req, res) => {
  const { email, password } = req.body || {};
  const normalizedEmail = email?.trim().toLowerCase();
  if (!normalizedEmail || !password) {
    return res.status(400).json({ error: 'Email and password required' });
  }
  if (USERS[normalizedEmail] !== password) {
    return res.status(401).json({ error: 'Invalid email or password' });
  }
  req.session.email = normalizedEmail;
  res.json({ success: true, email: normalizedEmail });
});

app.post('/api/logout', (req, res) => {
  req.session.destroy(() => {});
  res.json({ success: true });
});

app.get('/api/me', (req, res) => {
  if (AUTH_DISABLED) return res.json({ email: 'guest@local' });
  if (req.session?.email) return res.json({ email: req.session.email });
  res.status(401).json({ error: 'Not authenticated' });
});

const storage = multer.diskStorage({
  destination: (_, __, cb) => cb(null, TEST_SHEETS_PATH),
  filename: (_, file, cb) => cb(null, file.originalname),
});
const upload = multer({ storage });

app.get('/api/sheets', requireAuth, (req, res) => {
  try {
    const files = fs.readdirSync(TEST_SHEETS_PATH)
      .filter(f => f.endsWith('.xlsx') || f.endsWith('.xls'))
      .map(filename => {
        const fullPath = path.join(TEST_SHEETS_PATH, filename);
        const stats = fs.statSync(fullPath);
        const ticketMatch = filename.match(/^([A-Z]+-\d+)/);
        return {
          filename,
          ticketKey: ticketMatch ? ticketMatch[1] : null,
          size: stats.size,
          modified: stats.mtime.toISOString(),
        };
      })
      .sort((a, b) => new Date(b.modified) - new Date(a.modified));
    res.json(files);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/sheets/:filename/download', requireAuth, (req, res) => {
  const filename = path.basename(req.params.filename);
  if (filename !== req.params.filename) return res.status(400).send('Invalid filename');
  const filePath = path.join(TEST_SHEETS_PATH, filename);
  if (!fs.existsSync(filePath)) return res.status(404).send('File not found');
  res.download(filePath, filename);
});

app.post('/api/sheets/download-batch', requireAuth, (req, res) => {
  const { files } = req.body || {};
  if (!Array.isArray(files) || files.length === 0) return res.status(400).json({ error: 'No files specified' });
  const safeFiles = files.filter((f) => typeof f === 'string' && path.basename(f) === f && (f.endsWith('.xlsx') || f.endsWith('.xls')));
  if (safeFiles.length === 0) return res.status(400).json({ error: 'No valid files' });
  res.setHeader('Content-Type', 'application/zip');
  res.setHeader('Content-Disposition', 'attachment; filename="test-sheets.zip"');
  const archive = archiver('zip', { zlib: { level: 9 } });
  archive.pipe(res);
  for (const f of safeFiles) {
    const filePath = path.join(TEST_SHEETS_PATH, f);
    if (fs.existsSync(filePath)) archive.file(filePath, { name: f });
  }
  archive.finalize();
});

app.get('/api/sheets/open', requireAuth, (req, res) => {
  const filename = req.query.file;
  if (!filename) return res.status(400).send('Missing file parameter');
  const safeName = path.basename(filename);
  const filePath = path.join(TEST_SHEETS_PATH, safeName);
  if (!fs.existsSync(filePath)) return res.status(404).send('File not found');
  const mime = safeName.endsWith('.xlsx') ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'application/vnd.ms-excel';
  res.setHeader('Content-Type', mime);
  res.setHeader('Content-Disposition', `inline; filename="${safeName.replace(/"/g, '\\"')}"`);
  res.sendFile(path.resolve(filePath));
});

app.post('/api/sheets/upload', requireAuth, upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  res.json({ filename: req.file.filename, success: true });
});

app.delete('/api/sheets/:filename', requireAuth, (req, res) => {
  const filename = path.basename(req.params.filename);
  if (filename !== req.params.filename) return res.status(400).json({ error: 'Invalid filename' });
  const filePath = path.join(TEST_SHEETS_PATH, filename);
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'File not found' });
  fs.unlinkSync(filePath);
  res.json({ success: true });
});

app.listen(PORT, () => {
  console.log(`Test Sheet Manager running at http://localhost:${PORT}`);
  console.log(`Sheets folder: ${TEST_SHEETS_PATH}`);
  if (AUTH_DISABLED) console.log('Auth is DISABLED — login bypassed');
});
