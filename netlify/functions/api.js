const { getStore } = require('@netlify/blobs');
const Busboy = require('busboy');
const archiver = require('archiver');

const STORE_NAME = 'test-sheets';
const AUTH_DISABLED = true;

function json(body, status = 200) {
  return {
    statusCode: status,
    headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
    body: JSON.stringify(body),
  };
}

function text(body, status = 200, contentType = 'text/plain') {
  return {
    statusCode: status,
    headers: { 'Content-Type': contentType, 'Access-Control-Allow-Origin': '*' },
    body,
  };
}

function err(message, status = 500) {
  return json({ error: message }, status);
}

function parseMultipart(event) {
  return new Promise((resolve, reject) => {
    const fields = {};
    const body = event.isBase64Encoded ? Buffer.from(event.body, 'base64') : event.body;
    const headers = {};
    for (const [k, v] of Object.entries(event.headers || {})) headers[k.toLowerCase()] = v;
    const busboy = Busboy({ headers: { ...headers, 'content-type': headers['content-type'] || event.headers?.['Content-Type'] } });

    busboy.on('file', (fieldname, stream, info) => {
      const { filename } = info;
      const chunks = [];
      stream.on('data', (d) => chunks.push(d));
      stream.on('end', () => {
        fields[fieldname] = { filename: filename || 'upload.xlsx', content: Buffer.concat(chunks) };
      });
      stream.resume();
    });
    busboy.on('field', (name, value) => { fields[name] = value; });
    busboy.on('finish', () => resolve(fields));
    busboy.on('error', reject);
    if (body) busboy.write(body);
    busboy.end();
  });
}

exports.handler = async (event, context) => {
  const path = '/api' + (event.path.replace(/^\/\.netlify\/functions\/api/, '') || '');
  const method = event.httpMethod;
  const store = getStore(STORE_NAME);

  if (path === '/api/me' && method === 'GET') {
    return json(AUTH_DISABLED ? { email: 'guest@local' } : { error: 'Not authenticated' }, AUTH_DISABLED ? 200 : 401);
  }

  if (path === '/api/sheets' && method === 'GET') {
    try {
      const indexRaw = await store.get('__index');
      const index = indexRaw ? JSON.parse(indexRaw) : [];
      const files = index
        .filter((f) => f.filename && (f.filename.endsWith('.xlsx') || f.filename.endsWith('.xls')))
        .map((f) => {
          const ticketMatch = f.filename.match(/^([A-Z]+-\d+)/);
          return {
            filename: f.filename,
            ticketKey: ticketMatch ? ticketMatch[1] : null,
            size: f.size || 0,
            modified: f.modified || new Date().toISOString(),
          };
        })
        .sort((a, b) => new Date(b.modified) - new Date(a.modified));
      return json(files);
    } catch (e) {
      return err(e.message, 500);
    }
  }

  if (path.startsWith('/api/sheets/') && path.endsWith('/download') && method === 'GET') {
    const filename = decodeURIComponent(path.replace(/^\/api\/sheets\//, '').replace(/\/download$/, ''));
    if (!filename || filename.includes('/') || filename.includes('..')) return text('Invalid filename', 400);
    const valid = ['.xlsx', '.xls'].some((e) => filename.endsWith(e));
    if (!valid) return text('File not found', 404);
    const data = await store.get(filename, { type: 'arrayBuffer' });
    if (!data) return text('File not found', 404);
    return {
      statusCode: 200,
      headers: {
        'Content-Type': filename.endsWith('.xlsx') ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'application/vnd.ms-excel',
        'Content-Disposition': `attachment; filename="${filename.replace(/"/g, '\\"')}"`,
        'Access-Control-Allow-Origin': '*',
      },
      body: Buffer.from(data).toString('base64'),
      isBase64Encoded: true,
    };
  }

  if (path === '/api/sheets/open' && method === 'GET') {
    const filename = event.queryStringParameters?.file;
    if (!filename) return text('Missing file parameter', 400);
    const safeName = filename.replace(/\.\./g, '').split('/').pop();
    const valid = ['.xlsx', '.xls'].some((e) => safeName.endsWith(e));
    if (!valid) return text('File not found', 404);
    const data = await store.get(safeName, { type: 'arrayBuffer' });
    if (!data) return text('File not found', 404);
    return {
      statusCode: 200,
      headers: {
        'Content-Type': safeName.endsWith('.xlsx') ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'application/vnd.ms-excel',
        'Content-Disposition': `inline; filename="${safeName.replace(/"/g, '\\"')}"`,
        'Access-Control-Allow-Origin': '*',
      },
      body: Buffer.from(data).toString('base64'),
      isBase64Encoded: true,
    };
  }

  if (path === '/api/sheets/upload' && method === 'POST') {
    try {
      const fields = await parseMultipart(event);
      const file = fields.file;
      if (!file || !file.content) return err('No file uploaded', 400);
      const filename = file.filename || 'unnamed.xlsx';
      if (!['.xlsx', '.xls'].some((e) => filename.toLowerCase().endsWith(e))) return err('Invalid file type', 400);
      const modified = new Date().toISOString();
      await store.set(filename, file.content, { metadata: { size: file.content.length, modified } });
      const indexRaw = await store.get('__index');
      const index = indexRaw ? JSON.parse(indexRaw) : [];
      const existing = index.findIndex((e) => e.filename === filename);
      const entry = { filename, size: file.content.length, modified };
      if (existing >= 0) index[existing] = entry;
      else index.push(entry);
      await store.set('__index', JSON.stringify(index));
      return json({ filename, success: true });
    } catch (e) {
      return err(e.message || 'Upload failed', 500);
    }
  }

  if (path.startsWith('/api/sheets/') && !path.includes('/download') && method === 'DELETE') {
    const filename = decodeURIComponent(path.replace(/^\/api\/sheets\//, ''));
    if (!filename || filename.includes('/') || filename.includes('..')) return err('Invalid filename', 400);
    const valid = ['.xlsx', '.xls'].some((e) => filename.endsWith(e));
    if (!valid) return err('File not found', 404);
    const exists = await store.get(filename);
    if (!exists) return err('File not found', 404);
    await store.delete(filename);
    const indexRaw = await store.get('__index');
    if (indexRaw) {
      const index = JSON.parse(indexRaw).filter((e) => e.filename !== filename);
      await store.set('__index', JSON.stringify(index));
    }
    return json({ success: true });
  }

  if (path === '/api/sheets/download-batch' && method === 'POST') {
    let files = [];
    try {
      files = JSON.parse(event.body || '{}').files || [];
    } catch {
      return err('Invalid body', 400);
    }
    if (!Array.isArray(files) || files.length === 0) return err('No files specified', 400);
    const safeFiles = files.filter((f) => typeof f === 'string' && /\.(xlsx|xls)$/i.test(f) && !f.includes('/') && !f.includes('..'));
    if (safeFiles.length === 0) return err('No valid files', 400);

    const buffers = [];
    for (const f of safeFiles) {
      const data = await store.get(f, { type: 'arrayBuffer' });
      if (data) buffers.push({ name: f, data: Buffer.from(data) });
    }
    if (buffers.length === 0) return err('No files found', 404);

    return new Promise((resolve, reject) => {
      const chunks = [];
      const archive = archiver('zip', { zlib: { level: 9 } });
      archive.on('data', (c) => chunks.push(c));
      archive.on('end', () => {
        resolve({
          statusCode: 200,
          headers: {
            'Content-Type': 'application/zip',
            'Content-Disposition': 'attachment; filename="test-sheets.zip"',
            'Access-Control-Allow-Origin': '*',
          },
          body: Buffer.concat(chunks).toString('base64'),
          isBase64Encoded: true,
        });
      });
      archive.on('error', reject);
      for (const b of buffers) archive.append(b.data, { name: b.name });
      archive.finalize();
    });
  }

  return err('Not found', 404);
};
