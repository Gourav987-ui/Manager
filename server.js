const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const session = require('express-session');
const archiver = require('archiver');
const ExcelJS = require('exceljs');

let config = {};
try {
  config = JSON.parse(fs.readFileSync(path.join(__dirname, 'config.json'), 'utf8'));
} catch {
  config = {};
}
const TEST_SHEETS_PATH = config.testSheetsPath || path.join(__dirname, 'Testsheets');
const METADATA_PATH = path.join(__dirname, 'sheets-metadata.json');
const USE_NETLIFY_BLOBS = Boolean(process.env.NETLIFY || process.env.NETLIFY_DEV || process.env.NETLIFY_BLOBS_CONTEXT);
const netlifyBlobs = USE_NETLIFY_BLOBS ? require('@netlify/blobs') : null;
let blobStore = null;

function getBlobStore() {
  if (!USE_NETLIFY_BLOBS) return null;
  if (!blobStore) {
    try {
      blobStore = netlifyBlobs.getStore('testsheets');
    } catch (err) {
      throw new Error(`Netlify Blobs not initialized: ${err.message || err}`);
    }
  }
  return blobStore;
}

function loadMetadata() {
  try {
    const data = fs.readFileSync(METADATA_PATH, 'utf8');
    return JSON.parse(data);
  } catch {
    return {};
  }
}

function saveMetadata(data) {
  fs.writeFileSync(METADATA_PATH, JSON.stringify(data, null, 2), 'utf8');
}

function emailToName(email) {
  if (!email || typeof email !== 'string') return null;
  const local = email.split('@')[0] || '';
  const parts = local
    .split(/[._\s-]+/)
    .filter(Boolean)
    .map(p => p.charAt(0).toUpperCase() + p.slice(1).toLowerCase());
  if (parts.length === 0) return null;
  const seen = new Set();
  const unique = [];
  parts.forEach((part) => {
    const key = part.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    unique.push(part);
  });
  return unique.join(' ');
}
const PORT = config.port || 3456;
const USERS = config.auth?.users || { 'admin@example.com': 'admin123' };
const AUTH_DISABLED = !!config.auth?.disabled;

if (!USE_NETLIFY_BLOBS && !fs.existsSync(TEST_SHEETS_PATH)) {
  fs.mkdirSync(TEST_SHEETS_PATH, { recursive: true });
}

function listLocalFilenames() {
  return fs.readdirSync(TEST_SHEETS_PATH).filter((f) => f.endsWith('.xlsx') || f.endsWith('.xls'));
}

async function listStoredFilenames() {
  if (!USE_NETLIFY_BLOBS) return listLocalFilenames();
  const store = getBlobStore();
  const result = await store.list();
  return (result?.blobs || [])
    .map((b) => b.key)
    .filter((key) => key.endsWith('.xlsx') || key.endsWith('.xls'));
}

async function listStoredFiles() {
  if (!USE_NETLIFY_BLOBS) {
    const meta = loadMetadata();
    return listLocalFilenames().map((filename) => {
      const fullPath = path.join(TEST_SHEETS_PATH, filename);
      const stats = fs.statSync(fullPath);
      return {
        filename,
        size: stats.size,
        modified: stats.mtime.toISOString(),
        ownerEmail: meta[filename] || null,
      };
    });
  }
  const store = getBlobStore();
  const list = await store.list();
  const filenames = (list?.blobs || [])
    .map((b) => b.key)
    .filter((key) => key.endsWith('.xlsx') || key.endsWith('.xls'));
  const entries = await Promise.all(
    filenames.map(async (filename) => {
      const metaResult = await store.getMetadata(filename).catch(() => null);
      const metadata = metaResult?.metadata || {};
      return {
        filename,
        size: typeof metadata.size === 'number' ? metadata.size : 0,
        modified: metadata.updatedAt || new Date().toISOString(),
        ownerEmail: metadata.ownerEmail || null,
      };
    })
  );
  return entries;
}

async function getStoredFileBuffer(filename) {
  if (!USE_NETLIFY_BLOBS) {
    const filePath = path.join(TEST_SHEETS_PATH, filename);
    if (!fs.existsSync(filePath)) return null;
    return fs.readFileSync(filePath);
  }
  const store = getBlobStore();
  const data = await store.get(filename, { type: 'arrayBuffer' });
  if (!data) return null;
  return Buffer.from(data);
}

async function storedFileExists(filename) {
  if (!USE_NETLIFY_BLOBS) {
    const filePath = path.join(TEST_SHEETS_PATH, filename);
    return fs.existsSync(filePath);
  }
  const store = getBlobStore();
  const metaResult = await store.getMetadata(filename).catch(() => null);
  return Boolean(metaResult);
}

async function saveStoredFile(filename, buffer, ownerEmail) {
  if (!USE_NETLIFY_BLOBS) {
    const filePath = path.join(TEST_SHEETS_PATH, filename);
    fs.writeFileSync(filePath, buffer);
    const meta = loadMetadata();
    meta[filename] = ownerEmail || 'guest@local';
    saveMetadata(meta);
    return;
  }
  const store = getBlobStore();
  await store.set(filename, buffer, {
    metadata: {
      ownerEmail: ownerEmail || 'guest@local',
      size: buffer.length,
      updatedAt: new Date().toISOString(),
    },
  });
}

async function deleteStoredFile(filename) {
  if (!USE_NETLIFY_BLOBS) {
    const filePath = path.join(TEST_SHEETS_PATH, filename);
    if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
    const meta = loadMetadata();
    delete meta[filename];
    saveMetadata(meta);
    return;
  }
  const store = getBlobStore();
  await store.delete(filename);
}

async function renameStoredFile(oldName, newName, ownerEmail) {
  if (!USE_NETLIFY_BLOBS) {
    const oldPath = path.join(TEST_SHEETS_PATH, oldName);
    const newPath = path.join(TEST_SHEETS_PATH, newName);
    fs.renameSync(oldPath, newPath);
    const meta = loadMetadata();
    const nextOwner = meta[oldName] || ownerEmail || 'guest@local';
    delete meta[oldName];
    meta[newName] = nextOwner;
    saveMetadata(meta);
    return { ownerEmail: nextOwner };
  }
  const store = getBlobStore();
  const data = await store.get(oldName, { type: 'arrayBuffer' });
  if (!data) return { ownerEmail: ownerEmail || 'guest@local' };
  const metaResult = await store.getMetadata(oldName).catch(() => null);
  const metadata = metaResult?.metadata || {};
  const buffer = Buffer.from(data);
  const nextOwner = metadata.ownerEmail || ownerEmail || 'guest@local';
  await store.set(newName, buffer, {
    metadata: {
      ...metadata,
      ownerEmail: nextOwner,
      size: typeof metadata.size === 'number' ? metadata.size : buffer.length,
      updatedAt: new Date().toISOString(),
    },
  });
  await store.delete(oldName);
  return { ownerEmail: nextOwner };
}

async function getStoredOwnerEmail(filename) {
  if (!USE_NETLIFY_BLOBS) {
    const meta = loadMetadata();
    return meta[filename] || null;
  }
  const store = getBlobStore();
  const metaResult = await store.getMetadata(filename).catch(() => null);
  return metaResult?.metadata?.ownerEmail || null;
}

/** If a file with the same name already exists (case-insensitive), returns that name; otherwise null. */
async function findExistingFilenameConflict(desiredName) {
  const files = await listStoredFilenames();
  const lower = desiredName.toLowerCase();
  return files.find((f) => f.toLowerCase() === lower) || null;
}

function duplicateFileError(conflictName) {
  return `Duplicate file present: "${conflictName}" already exists. Delete or rename the existing file first.`;
}

const QA_NAME = config.qaName || 'Gourav Singh';
const SHEET_HEADERS = ['Test case', 'Expected Result', 'Actual Result', 'Status', 'Message', "Developer's Remark", 'Bug Rating'];

function normalizeDisplayName(name) {
  const parts = String(name || '')
    .split(/[\s.]+/)
    .map(p => p.trim())
    .filter(Boolean);
  if (parts.length === 0) return '';
  const seen = new Set();
  const unique = [];
  parts.forEach((part) => {
    const key = part.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    unique.push(part);
  });
  return unique.join(' ');
}

function getJiraSettings() {
  const jira = config.jira || {};
  return {
    domain: process.env.JIRA_DOMAIN || jira.domain,
    email: process.env.JIRA_EMAIL || jira.email,
    apiToken: process.env.JIRA_API_TOKEN || jira.apiToken,
  };
}

function sanitizeFilenameComponent(value) {
  if (!value) return 'Manual_Test_Cases';
  const cleaned = value.replace(/[<>:"/\\|?*]/g, '').replace(/\s+/g, ' ').trim();
  return cleaned || 'Manual_Test_Cases';
}

function htmlToText(html) {
  if (!html) return '';
  let text = String(html);
  text = text.replace(/<br\s*\/?>/gi, '\n');
  text = text.replace(/<\/p>/gi, '\n');
  text = text.replace(/<\/li>/gi, '\n');
  text = text.replace(/<li>/gi, '- ');
  text = text.replace(/<\/h\d>/gi, '\n');
  text = text.replace(/<[^>]*>/g, '');
  const entities = {
    '&nbsp;': ' ',
    '&amp;': '&',
    '&lt;': '<',
    '&gt;': '>',
    '&quot;': '"',
    '&#39;': "'",
  };
  Object.entries(entities).forEach(([k, v]) => {
    text = text.replace(new RegExp(k, 'g'), v);
  });
  return text;
}

function adfToText(node) {
  if (!node) return '';
  if (Array.isArray(node)) return node.map(adfToText).join('');
  if (typeof node === 'string') return node;
  const content = node.content ? node.content.map(adfToText).join('') : '';
  switch (node.type) {
    case 'text':
      return node.text || '';
    case 'hardBreak':
      return '\n';
    case 'paragraph':
    case 'heading':
      return `${content}\n`;
    case 'listItem':
      return `- ${content}\n`;
    case 'bulletList':
    case 'orderedList':
      return `${content}\n`;
    default:
      return content;
  }
}

function normalizeText(text) {
  return String(text || '')
    .replace(/\r\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function extractAcceptanceCriteria(text) {
  if (!text) return [];
  const lines = String(text).split(/\r?\n/);
  const startIndex = lines.findIndex((l) => /acceptance criteria|^ac\b/i.test(l.trim()));
  if (startIndex === -1) return [];
  const results = [];
  for (let i = startIndex; i < lines.length; i += 1) {
    const line = lines[i].trim();
    if (i === startIndex) {
      const remainder = line.replace(/acceptance criteria/i, '').replace(/^[:\-\s]+/, '').trim();
      if (remainder && !/^[#]+$/.test(remainder)) results.push(remainder);
      continue;
    }
    if (!line && results.length > 0) break;
    if (/^(#+\s+|[A-Z][A-Za-z\s]+:)$/.test(line) && results.length > 0) break;
    if (!line) continue;
    results.push(line);
  }
  return results
    .map((l) => l.replace(/^[-*•\d.\)\s]+/, '').trim())
    .filter(Boolean);
}

function detectFeatureFlag(text) {
  const raw = String(text || '');
  if (!/(feature flag|feature-flag|flag on|flag off|when enabled|when disabled|flagged)/i.test(raw)) {
    return { hasFlag: false, name: null };
  }
  const ignore = new Set(['on', 'off', 'enabled', 'disabled', 'feature', 'flag']);
  const candidates = [];
  const regexes = [
    /feature flag[:\s]*`?([a-z0-9_-]{3,})`?/gi,
    /flag[:\s]*`?([a-z0-9_-]{3,})`?/gi,
    /`([a-z0-9_-]{3,})`/gi,
  ];
  regexes.forEach((rgx) => {
    let match;
    while ((match = rgx.exec(raw)) !== null) {
      const candidate = match[1];
      if (candidate && !ignore.has(candidate.toLowerCase())) {
        candidates.push(candidate);
      }
    }
  });
  return { hasFlag: true, name: candidates[0] || null };
}

function caseTitleFromText(text) {
  const cleaned = String(text || '')
    .replace(/^[^a-zA-Z0-9]+/, '')
    .replace(/\.+$/, '')
    .trim();
  if (!cleaned) return 'Test case';
  const words = cleaned.split(/\s+/).slice(0, 8);
  return words.join(' ');
}

function expectedFromText(text) {
  const trimmed = String(text || '').trim();
  if (!trimmed) return 'Action should complete successfully.';
  if (/\bshould\b/i.test(trimmed) || /\bwould\b/i.test(trimmed)) return trimmed;
  if (/\bmust\b/i.test(trimmed)) return trimmed.replace(/\bmust\b/i, 'should');
  if (/\bwill\b/i.test(trimmed)) return trimmed.replace(/\bwill\b/i, 'should');
  return `Result should be: ${trimmed}`;
}

function dedupeCases(cases) {
  const seen = new Set();
  return cases.filter((c) => {
    const key = String(c.name || '').toLowerCase();
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function coverageCases(summary, flagLabel, state) {
  const suffix = flagLabel ? ` when ${flagLabel} is ${state}` : '';
  const label = summary || 'feature';
  return [
    {
      name: `Reject invalid input for ${label}`,
      expected: `Invalid input should be rejected with a clear error message${suffix}.`,
    },
    {
      name: `Handle missing data for ${label}`,
      expected: `Missing data should be handled gracefully${suffix}.`,
    },
    {
      name: `Verify logging for ${label}`,
      expected: `Relevant logs should be recorded${suffix}.`,
    },
  ];
}

function defaultCases(summary) {
  const label = summary || 'feature';
  return [
    {
      name: `Verify ${label} happy path`,
      expected: `${label} should work for valid inputs.`,
    },
    {
      name: `Reject invalid input for ${label}`,
      expected: 'Invalid input should be rejected with a clear error message.',
    },
    {
      name: `Handle missing data for ${label}`,
      expected: 'Missing data should be handled gracefully.',
    },
    {
      name: `Process multiple items for ${label}`,
      expected: 'Multiple items should be handled without errors.',
    },
    {
      name: `Verify logging for ${label}`,
      expected: `Relevant logs should be recorded for ${label} actions.`,
    },
    {
      name: `Retry ${label} after failure`,
      expected: `${label} should recover or allow retry after failure.`,
    },
  ];
}

function buildTestCases(summary, acceptanceCriteria, featureFlag) {
  const acCases = acceptanceCriteria.map((ac) => ({
    name: caseTitleFromText(ac),
    expected: expectedFromText(ac),
  }));
  if (!featureFlag?.hasFlag) {
    const base = acCases.length > 0 ? acCases.concat(coverageCases(summary)) : defaultCases(summary);
    return { cases: dedupeCases(base) };
  }
  const flagLabel = featureFlag.name ? `the "${featureFlag.name}" flag` : 'the feature flag';
  const onCases = acCases.length > 0
    ? acCases.concat(coverageCases(summary, flagLabel, 'ON'))
    : [
      {
        name: `Verify new behavior for ${summary || 'feature'}`,
        expected: `New behavior should apply when ${flagLabel} is ON.`,
      },
      ...coverageCases(summary, flagLabel, 'ON'),
    ];
  const offCases = [
    {
      name: `Verify legacy behavior for ${summary || 'feature'}`,
      expected: `Legacy behavior should remain unchanged when ${flagLabel} is OFF.`,
    },
    ...coverageCases(summary, flagLabel, 'OFF'),
  ];
  return {
    flagOffCases: dedupeCases(offCases),
    flagOnCases: dedupeCases(onCases),
  };
}

function applyRowStyle(row, isBold) {
  for (let col = 1; col <= 7; col += 1) {
    const cell = row.getCell(col);
    cell.alignment = { wrapText: true, horizontal: 'left', vertical: 'center' };
    if (isBold) cell.font = { bold: true };
  }
}

function addHeaderRow(sheet, rowNumber) {
  const row = sheet.getRow(rowNumber);
  SHEET_HEADERS.forEach((header, index) => {
    row.getCell(index + 1).value = header;
  });
  applyRowStyle(row, true);
}

function addTestCaseRow(sheet, rowNumber, testCase) {
  const row = sheet.getRow(rowNumber);
  row.getCell(1).value = testCase.name;
  row.getCell(2).value = testCase.expected;
  applyRowStyle(row, false);
}

async function fetchJiraIssue(key) {
  const jira = getJiraSettings();
  if (!jira.email || !jira.apiToken || !jira.domain) {
    return {
      ok: false,
      error: 'Jira configuration is required. Set jira.email, jira.apiToken, and jira.domain in config.json or provide JIRA_EMAIL, JIRA_API_TOKEN, and JIRA_DOMAIN env vars.',
    };
  }
  const auth = Buffer.from(`${jira.email}:${jira.apiToken}`).toString('base64');
  const url = `https://${jira.domain}/rest/api/3/issue/${encodeURIComponent(
    key
  )}?fields=summary,description,assignee,issuetype&expand=renderedFields`;
  try {
    const r = await fetch(url, {
      headers: {
        Authorization: `Basic ${auth}`,
        Accept: 'application/json',
      },
    });
    if (r.status === 404) return { ok: false, error: `Ticket ${key} does not exist` };
    if (r.status === 403) return { ok: false, error: `No access to ticket ${key}` };
    if (!r.ok) {
      const errBody = await r.json().catch(() => ({}));
      return { ok: false, error: errBody.errorMessages?.[0] || `Jira error: ${r.status}` };
    }
    const data = await r.json();
    return { ok: true, issue: data };
  } catch (err) {
    return { ok: false, error: err.message || 'Failed to reach Jira' };
  }
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

app.get('/', (_, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

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
  if (AUTH_DISABLED) return res.json({ email: 'guest@local', name: null });
  if (req.session?.email) {
    const name = emailToName(req.session.email);
    return res.json({ email: req.session.email, name });
  }
  res.status(401).json({ error: 'Not authenticated' });
});

const upload = multer({ storage: multer.memoryStorage() });

app.get('/api/sheets', requireAuth, async (req, res) => {
  try {
    const files = await listStoredFiles();
    const results = files
      .map((file) => {
        const ticketMatch = file.filename.match(/^([A-Z]+-\d+)/);
        const owner = file.ownerEmail ? emailToName(file.ownerEmail) : null;
        const ownedByMe = !file.ownerEmail || file.ownerEmail === req.session?.email;
        return {
          filename: file.filename,
          ticketKey: ticketMatch ? ticketMatch[1] : null,
          size: file.size,
          modified: file.modified,
          owner,
          ownedByMe,
        };
      })
      .sort((a, b) => new Date(b.modified) - new Date(a.modified));
    res.json(results);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/sheets/:filename/download', requireAuth, (req, res) => {
  const filename = path.basename(req.params.filename);
  if (filename !== req.params.filename) return res.status(400).send('Invalid filename');
  (async () => {
    const buffer = await getStoredFileBuffer(filename);
    if (!buffer) return res.status(404).send('File not found');
    const mime = filename.endsWith('.xlsx') ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'application/vnd.ms-excel';
    res.setHeader('Content-Type', mime);
    res.setHeader('Content-Disposition', `attachment; filename="${filename.replace(/"/g, '\\"')}"`);
    res.send(buffer);
  })().catch((err) => res.status(500).json({ error: err.message }));
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
  (async () => {
    for (const f of safeFiles) {
      const buffer = await getStoredFileBuffer(f);
      if (buffer) archive.append(buffer, { name: f });
    }
    archive.finalize();
  })().catch((err) => res.status(500).json({ error: err.message }));
});

app.get('/api/sheets/open', requireAuth, (req, res) => {
  const filename = req.query.file;
  if (!filename) return res.status(400).send('Missing file parameter');
  const safeName = path.basename(filename);
  (async () => {
    const buffer = await getStoredFileBuffer(safeName);
    if (!buffer) return res.status(404).send('File not found');
    const mime = safeName.endsWith('.xlsx') ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'application/vnd.ms-excel';
    res.setHeader('Content-Type', mime);
    res.setHeader('Content-Disposition', `inline; filename="${safeName.replace(/"/g, '\\"')}"`);
    res.send(buffer);
  })().catch((err) => res.status(500).json({ error: err.message }));
});

app.post('/api/sheets/upload', requireAuth, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  const safeName = path.basename(req.file.originalname);
  if (safeName !== req.file.originalname) return res.status(400).json({ error: 'Invalid filename' });
  const conflictName = await findExistingFilenameConflict(safeName);
  if (conflictName) {
    return res.status(409).json({
      error: duplicateFileError(conflictName),
    });
  }
  await saveStoredFile(safeName, req.file.buffer, req.session?.email || 'guest@local');
  res.json({ filename: safeName, success: true });
});

app.post('/api/sheets/create-from-ticket', requireAuth, async (req, res) => {
  const { ticketKey } = req.body || {};
  const key = String(ticketKey || '').trim().toUpperCase();
  if (!key || !/^[A-Z]+-\d+$/.test(key)) {
    return res.status(400).json({ error: 'Valid Jira ticket key required (e.g. INVST-123)' });
  }
  const issueResult = await fetchJiraIssue(key);
  if (!issueResult.ok) {
    return res.status(400).json({ error: issueResult.error });
  }
  const issue = issueResult.issue || {};
  const summary = issue.fields?.summary?.trim() || key;
  const safeSummary = sanitizeFilenameComponent(summary);
  const filename = `${key}_${safeSummary}.xlsx`;
  const conflictName = await findExistingFilenameConflict(filename);
  if (conflictName) {
    return res.status(409).json({
      error: duplicateFileError(conflictName),
    });
  }
  try {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('ManualTests', { views: [{ state: 'frozen', ySplit: 3 }] });
    sheet.columns = [
      { width: 60 },
      { width: 55 },
      { width: 25 },
      { width: 15 },
      { width: 30 },
      { width: 25 },
      { width: 15 },
    ];
    const jiraDomain = getJiraSettings().domain || 'orion-advisor.atlassian.net';
    const ticketUrl = `https://${jiraDomain}/browse/${key}`;
    const developerName = issue.fields?.assignee?.displayName || '—';
    const descriptionText = normalizeText(
      issue.renderedFields?.description
        ? htmlToText(issue.renderedFields.description)
        : adfToText(issue.fields?.description)
    );
    const acceptanceCriteria = extractAcceptanceCriteria(descriptionText);
    const featureFlag = detectFeatureFlag(descriptionText);
    const cases = buildTestCases(summary, acceptanceCriteria, featureFlag);

    const cellA1 = sheet.getCell('A1');
    cellA1.value = { text: ticketUrl, hyperlink: ticketUrl };
    cellA1.font = { bold: true };
    cellA1.alignment = { wrapText: true, horizontal: 'left', vertical: 'center' };
    const cellC1 = sheet.getCell('C1');
    cellC1.value = `Developer: ${developerName}`;
    cellC1.font = { bold: true };
    cellC1.alignment = { wrapText: true, horizontal: 'left', vertical: 'center' };
    const qaName = normalizeDisplayName(emailToName(req.session?.email)) || normalizeDisplayName(QA_NAME) || 'QA';
    const cellE1 = sheet.getCell('E1');
    cellE1.value = `QA - ${qaName}`;
    cellE1.font = { bold: true };
    cellE1.alignment = { wrapText: true, horizontal: 'left', vertical: 'center' };

    let rowIndex = 3;
    let caseNumber = 1;
    applyRowStyle(sheet.getRow(2), false);

    if (featureFlag.hasFlag) {
      const flagOffRow = sheet.getRow(rowIndex);
      flagOffRow.getCell(1).value = 'FLAG OFF';
      applyRowStyle(flagOffRow, true);
      rowIndex += 1;
      addHeaderRow(sheet, rowIndex);
      rowIndex += 1;
      cases.flagOffCases.forEach((testCase) => {
        addTestCaseRow(sheet, rowIndex, {
          name: `TC${String(caseNumber).padStart(2, '0')} - ${testCase.name}`,
          expected: testCase.expected,
        });
        caseNumber += 1;
        rowIndex += 1;
        applyRowStyle(sheet.getRow(rowIndex), false);
        rowIndex += 1;
      });

      const flagOnRow = sheet.getRow(rowIndex);
      flagOnRow.getCell(1).value = 'FLAG ON';
      applyRowStyle(flagOnRow, true);
      rowIndex += 1;
      addHeaderRow(sheet, rowIndex);
      rowIndex += 1;
      cases.flagOnCases.forEach((testCase) => {
        addTestCaseRow(sheet, rowIndex, {
          name: `TC${String(caseNumber).padStart(2, '0')} - ${testCase.name}`,
          expected: testCase.expected,
        });
        caseNumber += 1;
        rowIndex += 1;
        applyRowStyle(sheet.getRow(rowIndex), false);
        rowIndex += 1;
      });
    } else {
      addHeaderRow(sheet, rowIndex);
      rowIndex += 1;
      cases.cases.forEach((testCase) => {
        addTestCaseRow(sheet, rowIndex, {
          name: `TC${String(caseNumber).padStart(2, '0')} - ${testCase.name}`,
          expected: testCase.expected,
        });
        caseNumber += 1;
        rowIndex += 1;
        applyRowStyle(sheet.getRow(rowIndex), false);
        rowIndex += 1;
      });
    }
    const buffer = await workbook.xlsx.writeBuffer();
    await saveStoredFile(filename, buffer, req.session?.email || 'guest@local');
    res.json({ filename, success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/sheets/rename', requireAuth, (req, res) => {
  const { oldFilename, newFilename } = req.body || {};
  if (!oldFilename || !newFilename) {
    return res.status(400).json({ error: 'Old and new filenames are required' });
  }
  const oldName = path.basename(String(oldFilename));
  const newName = path.basename(String(newFilename));
  if (oldName !== oldFilename || newName !== newFilename) {
    return res.status(400).json({ error: 'Invalid filename' });
  }
  if (oldName === newName) {
    return res.status(400).json({ error: 'New filename must be different' });
  }
  if (!newName.endsWith('.xlsx') && !newName.endsWith('.xls')) {
    return res.status(400).json({ error: 'Filename must end with .xlsx or .xls' });
  }
  (async () => {
    const exists = await storedFileExists(oldName);
    if (!exists) return res.status(404).json({ error: 'File not found' });
    const ownerEmail = await getStoredOwnerEmail(oldName);
    if (ownerEmail && ownerEmail !== req.session?.email) {
      return res.status(403).json({ error: 'You can only rename files you uploaded' });
    }
    const conflictName = await findExistingFilenameConflict(newName);
    if (conflictName && conflictName.toLowerCase() !== oldName.toLowerCase()) {
      return res.status(409).json({ error: duplicateFileError(conflictName) });
    }
    if (conflictName && conflictName.toLowerCase() === oldName.toLowerCase()) {
      return res.status(400).json({ error: 'New filename differs only by case' });
    }
    await renameStoredFile(oldName, newName, ownerEmail || req.session?.email || 'guest@local');
    res.json({ success: true, filename: newName });
  })().catch((err) => res.status(500).json({ error: err.message }));
});

app.delete('/api/sheets/:filename', requireAuth, (req, res) => {
  const filename = path.basename(req.params.filename);
  if (filename !== req.params.filename) return res.status(400).json({ error: 'Invalid filename' });
  (async () => {
    const ownerEmail = await getStoredOwnerEmail(filename);
    if (ownerEmail && ownerEmail !== req.session?.email) {
      return res.status(403).json({ error: 'You can only delete files you uploaded' });
    }
    const exists = await storedFileExists(filename);
    if (!exists) return res.status(404).json({ error: 'File not found' });
    await deleteStoredFile(filename);
    res.json({ success: true });
  })().catch((err) => res.status(500).json({ error: err.message }));
});

if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`Test Sheet Manager running at http://localhost:${PORT}`);
    console.log(`Sheets folder: ${TEST_SHEETS_PATH}`);
    if (AUTH_DISABLED) console.log('Auth is DISABLED — login bypassed');
  });
}

module.exports = { app };
