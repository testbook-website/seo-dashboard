const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');
const sheetsService = require('./sheets_service');

const PORT = process.env.PORT || 4173;

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon'
};

function sendJson(res, statusCode, data) {
  res.writeHead(statusCode, {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  });
  res.end(JSON.stringify(data));
}

const server = http.createServer(async (req, res) => {
  // Handle CORS Preflight
  if (req.method === 'OPTIONS') {
    res.writeHead(204, {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    });
    return res.end();
  }

  const parsedUrl = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
  const pathname = parsedUrl.pathname;

  // --- API: Health & Connection Status ---
  if (pathname === '/api/status' && req.method === 'GET') {
    try {
      const status = await sheetsService.checkConnection();
      return sendJson(res, 200, {
        success: true,
        ...status
      });
    } catch (err) {
      return sendJson(res, 500, {
        success: false,
        error: err.message
      });
    }
  }

  // --- API: Dispatch New Assignment to Google Sheet (Inserts at Row 2) ---
  if (pathname === '/api/assign' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => { body += chunk; });
    req.on('end', async () => {
      try {
        const payload = JSON.parse(body);

        // Validation
        if (!payload.topic || !payload.category) {
          return sendJson(res, 400, {
            success: false,
            error: "Both 'topic' and 'category' are mandatory."
          });
        }

        console.log(`[API /assign] Processing assignment for "${payload.projectType || 'SEO Optimization'}": ${payload.topic}`);

        const result = await sheetsService.assignToSheet(payload);
        return sendJson(res, 200, result);

      } catch (err) {
        console.error("[API /assign Error]:", err.message);
        const statusCode = err.isDuplicate ? 409 : 500;
        return sendJson(res, statusCode, {
          success: false,
          duplicate: !!err.isDuplicate,
          error: err.message,
          details: err.duplicateDetails || null
        });
      }
    });
    return;
  }

  // --- API: 60-Day URL & Page Registry (Duplicate Prevention) ---
  if (pathname === '/api/url-registry' && req.method === 'GET') {
    try {
      const force = parsedUrl.searchParams.get('refresh') === 'true';
      const registry = await sheetsService.getRecentUrlRegistry(force);
      return sendJson(res, 200, {
        success: true,
        count: registry.length,
        days: 60,
        lastUpdated: new Date().toISOString(),
        registry
      });
    } catch (err) {
      console.error("[API /api/url-registry Error]:", err);
      return sendJson(res, 500, { success: false, error: err.message });
    }
  }

  // --- API: Real-time Live Duplicate Check (URL, Page, Topic within 30/60 days) ---
  if (pathname === '/api/check-duplicate') {
    try {
      if (req.method === 'GET') {
        const queryParams = {
          url: parsedUrl.searchParams.get('url') || '',
          page: parsedUrl.searchParams.get('page') || '',
          topic: parsedUrl.searchParams.get('topic') || '',
          date: parsedUrl.searchParams.get('date') || ''
        };
        const check = await sheetsService.checkDuplicateAssignment(queryParams);
        return sendJson(res, 200, { success: true, ...check });
      } else if (req.method === 'POST') {
        let body = '';
        req.on('data', chunk => { body += chunk; });
        req.on('end', async () => {
          try {
            const payload = JSON.parse(body || '{}');
            const check = await sheetsService.checkDuplicateAssignment(payload);
            return sendJson(res, 200, { success: true, ...check });
          } catch (e) {
            return sendJson(res, 400, { success: false, error: e.message });
          }
        });
        return;
      }
    } catch (err) {
      console.error("[API /api/check-duplicate Error]:", err);
      return sendJson(res, 500, { success: false, error: err.message });
    }
  }

  // --- API: Get All OND (Oct, Nov, Dec) 2026 Assignments (Live from Sheets) ---
  if ((pathname === '/api/ond-assignments' || pathname === '/api/october-assignments') && req.method === 'GET') {
    try {
      const data = await sheetsService.getONDAssignments();
      return sendJson(res, 200, data);
    } catch (err) {
      console.error("[API /api/ond-assignments Error]:", err);
      return sendJson(res, 500, { success: false, error: err.message });
    }
  }

  // --- API: Get Recent Sheet Rows ---
  if (pathname === '/api/recent' && req.method === 'GET') {
    try {
      const tab = parsedUrl.searchParams.get('tab') || 'SEO Optimization';
      const limit = parseInt(parsedUrl.searchParams.get('limit'), 10) || 10;
      const rows = await sheetsService.getRecentRows(tab, limit);
      return sendJson(res, 200, { success: true, tab, rows });
    } catch (err) {
      return sendJson(res, 500, { success: false, error: err.message });
    }
  }

  // --- Static File Serving ---
  let filePath = path.join(__dirname, pathname === '/' ? 'index.html' : pathname);
  
  // Prevent directory traversal
  if (!filePath.startsWith(__dirname)) {
    res.writeHead(403);
    return res.end("Forbidden");
  }

  // Block access to sensitive files (credentials, env files, keys)
  const baseName = path.basename(filePath).toLowerCase();
  if (
    baseName.startsWith('.env') ||
    baseName.startsWith('credentials') ||
    baseName.endsWith('.key') ||
    baseName.endsWith('.pem') ||
    baseName.endsWith('.token') ||
    baseName.includes('service-account')
  ) {
    res.writeHead(403, { 'Content-Type': 'text/plain; charset=utf-8' });
    return res.end("403 Forbidden: Access to sensitive files is restricted");
  }

  fs.stat(filePath, (err, stats) => {
    if (err || !stats.isFile()) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      return res.end("404 Not Found");
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || 'application/octet-stream';

    res.writeHead(200, { 'Content-Type': contentType });
    const stream = fs.createReadStream(filePath);
    stream.pipe(res);
  });
});

server.listen(PORT, () => {
  console.log(`================================================================`);
  console.log(`🚀 RankAssign PRO Server running on http://127.0.0.1:${PORT}`);
  console.log(`📊 Connected to Sheet: 1ihLeB9ZOJdaF841qGLoTWBULSRNsF9BjtxUXRxKuK2A`);
  console.log(`🎯 Inserting at Row 2 into 'SEO Optimization' and 'New Content'`);
  console.log(`================================================================`);
});
