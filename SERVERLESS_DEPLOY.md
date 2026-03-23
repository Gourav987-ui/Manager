# Serverless Deployment (Netlify)

The app now supports **serverless deployment** on Netlify. Files are stored in **Netlify Blob** instead of the local filesystem.

## Deploy to Netlify

1. **Push to Git** and connect your repo to Netlify.

2. **Build settings** (usually auto-detected):
   - Build command: `npm run build`
   - Publish directory: `public`
   - Functions directory: `netlify/functions` (default)

3. **Install dependencies**: Netlify runs `npm install` by default. Ensure `@netlify/blobs`, `busboy`, and `archiver` are in `package.json` (they are).

4. **Deploy** – Netlify will:
   - Build (no-op for static)
   - Publish the `public` folder
   - Deploy the `api` function
   - Redirect `/api/*` to the function

## How It Works

| Route | Function |
|-------|----------|
| `/api/me` | Returns guest user (auth disabled) |
| `GET /api/sheets` | List files from Blob store |
| `POST /api/sheets/upload` | Upload file to Blob |
| `GET /api/sheets/:file/download` | Download file |
| `GET /api/sheets/open?file=` | Open file inline |
| `DELETE /api/sheets/:file` | Delete file |
| `POST /api/sheets/download-batch` | Zip & download multiple |

## Local Development

**Option A – Netlify CLI** (recommended for testing serverless):

```bash
npm install -g netlify-cli
netlify dev
```

Opens the app and emulates functions + Blob locally.

**Option B – Classic server** (local filesystem):

```bash
npm start
```

Uses `server.js` and the `Testsheets` folder. Auth can be toggled in `config.json`.

## Storage

- **Serverless**: Netlify Blob store `test-sheets`
- **Local**: `config.testSheetsPath` or `./Testsheets`

Data does not migrate between them. Local files stay local; Blob files stay in Blob.
