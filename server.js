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
        console.error("[API /assign Error]:", err);
        return sendJson(res, 500, {
          success: false,
          error: err.message
        });
      }
    });
    return;
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
