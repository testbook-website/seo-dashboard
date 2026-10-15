const sheetsService = require('../sheets_service');
const { URL } = require('url');

function sendJson(res, statusCode, data) {
  res.writeHead(statusCode, {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  });
  res.end(JSON.stringify(data));
}

function parseBody(req) {
  return new Promise((resolve, reject) => {
    if (req.body && typeof req.body === 'object') {
      return resolve(req.body);
    }
    if (typeof req.body === 'string') {
      try {
        return resolve(JSON.parse(req.body));
      } catch (e) {
        return reject(e);
      }
    }
    let body = '';
    req.on('data', chunk => { body += chunk; });
    req.on('end', () => {
      if (!body) return resolve({});
      try {
        resolve(JSON.parse(body));
      } catch (err) {
        reject(err);
      }
    });
    req.on('error', err => reject(err));
  });
}

module.exports = async (req, res) => {
  // CORS Preflight
  if (req.method === 'OPTIONS') {
    res.writeHead(204, {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    });
    return res.end();
  }

  const parsedUrl = new URL(req.url, `https://${req.headers.host || 'localhost'}`);
  const pathname = parsedUrl.pathname.replace(/\/+$/, '') || '/';

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

  // --- API: Dispatch New Assignment (Row 2 Insertion) ---
  if (pathname === '/api/assign' && req.method === 'POST') {
    try {
      const data = await parseBody(req);
      if (!data || (!data.topic && !data.page)) {
        return sendJson(res, 400, {
          success: false,
          error: "Missing required topic/page or payload"
        });
      }

      const result = await sheetsService.assignToSheet(data);
      return sendJson(res, 200, result);
    } catch (err) {
      console.error("[Vercel /api/assign Error]:", err);
      return sendJson(res, 500, {
        success: false,
        error: err.message
      });
    }
  }

  // --- API: OND 2026 Assignments ---
  if ((pathname === '/api/ond-assignments' || pathname === '/api/october-assignments') && req.method === 'GET') {
    try {
      const data = await sheetsService.getONDAssignments();
      return sendJson(res, 200, data);
    } catch (err) {
      console.error("[Vercel /api/ond-assignments Error]:", err);
      return sendJson(res, 500, { success: false, error: err.message });
    }
  }

  // --- API: Recent Sheet Rows ---
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

  return sendJson(res, 404, { success: false, error: "API endpoint not found: " + pathname });
};
