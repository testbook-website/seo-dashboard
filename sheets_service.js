const fs = require('fs');
const crypto = require('crypto');
const path = require('path');

// Locate credentials file
let credPath = path.join(__dirname, 'credentials.json');
if (!fs.existsSync(credPath)) {
  credPath = path.join(__dirname, 'credentials.json.json');
}

if (!fs.existsSync(credPath)) {
  console.error("WARNING: No credentials.json or credentials.json.json found!");
}

const SPREADSHEET_ID = '1ihLeB9ZOJdaF841qGLoTWBULSRNsF9BjtxUXRxKuK2A';

const SHEET_METADATA = {
  "SEO Optimization": {
    sheetId: 1073261968,
    gid: "1073261968",
    url: "https://docs.google.com/spreadsheets/d/1ihLeB9ZOJdaF841qGLoTWBULSRNsF9BjtxUXRxKuK2A/edit?gid=1073261968#gid=1073261968"
  },
  "New Content": {
    sheetId: 1107724151,
    gid: "1107724151",
    url: "https://docs.google.com/spreadsheets/d/1ihLeB9ZOJdaF841qGLoTWBULSRNsF9BjtxUXRxKuK2A/edit?gid=1107724151#gid=1107724151"
  }
};

let cachedToken = null;
let tokenExpiresAt = 0;

/**
 * Generate OAuth2 JWT and fetch access token for the Google Service Account
 */
async function getAccessToken() {
  const now = Math.floor(Date.now() / 1000);
  if (cachedToken && tokenExpiresAt > now + 60) {
    return cachedToken;
  }

  const creds = JSON.parse(fs.readFileSync(credPath, 'utf8'));

  const header = { alg: 'RS256', typ: 'JWT' };
  const payload = {
    iss: creds.client_email,
    scope: 'https://www.googleapis.com/auth/spreadsheets',
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600,
    iat: now
  };

  const b64Header = Buffer.from(JSON.stringify(header)).toString('base64url');
  const b64Payload = Buffer.from(JSON.stringify(payload)).toString('base64url');
  const unsignedToken = `${b64Header}.${b64Payload}`;

  const signer = crypto.createSign('RSA-SHA256');
  signer.update(unsignedToken);
  signer.end();
  const signature = signer.sign(creds.private_key, 'base64url');

  const jwt = `${unsignedToken}.${signature}`;

  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt
    })
  });

  const data = await res.json();
  if (data.error) {
    throw new Error(data.error_description || data.error);
  }

  cachedToken = data.access_token;
  tokenExpiresAt = now + (data.expires_in || 3600);
  return cachedToken;
}

/**
 * Format date string into "14th October 2026"
 */
function formatDateWithSuffix(dateStr) {
  if (!dateStr) dateStr = new Date().toISOString().split('T')[0];
  const parts = dateStr.split('-');
  let d;
  if (parts.length === 3) {
    d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  } else {
    d = new Date(dateStr);
  }

  const day = d.getDate();
  const monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  const month = monthNames[d.getMonth()];
  const year = d.getFullYear();

  let suffix = "th";
  if (day === 1 || day === 21 || day === 31) suffix = "st";
  else if (day === 2 || day === 22) suffix = "nd";
  else if (day === 3 || day === 23) suffix = "rd";

  return `${day}${suffix} ${month} ${year}`;
}

/**
 * Build column values for "SEO Optimization" tab
 */
function buildSeoOptimizationRow(data) {
  const formattedDate = formatDateWithSuffix(data.date);
  const strategyGoal = `[${data.goalType || 'Goal'}]: ${data.targetGoal || ''} | Why Now: ${data.urgencyReason || ''}`;

  // Validate Search Intent against strict dropdown
  let validSearchIntent = "";
  if (data.searchIntent) {
    if (data.searchIntent.includes("Mixed")) validSearchIntent = "Mixed → balance info + action";
    else if (data.searchIntent.includes("Transactional")) validSearchIntent = "Transactional → highlight benefits, CTAs";
    else if (data.searchIntent.includes("Informational")) validSearchIntent = "Informational → explain, guide, clarify";
    else validSearchIntent = data.searchIntent;
  } else {
    validSearchIntent = "Informational → explain, guide, clarify";
  }

  // Validate Area of Optimisation against strict dropdown (Col 17)
  const allowedAreas = [
    "Content Outdated",
    "Factual Information Missing",
    "Missing Clarity",
    "Need To Work on Intro",
    "Page Structure",
    "Too Generic",
    "Update as per the expected notification."
  ];
  const validArea = allowedAreas.includes(data.areaOfOptimisation) ? data.areaOfOptimisation : "";

  // Validate Content Gap against strict dropdown (Col 21)
  const allowedGaps = [
    "Missing Examples",
    "Missing Explanations",
    "Missing FAQs",
    "Missing Sections",
    "Missing Tables / Lists"
  ];
  const validGap = allowedGaps.includes(data.contentGap) ? data.contentGap : "";

  // Priority clamp for SEO Optimization (Col 7 only accepts P0, P1, P2)
  let priority = data.contentPriority || "P1";
  if (priority === "P3") priority = "P2";

  // Prepend topic and strategic goal to SEO Suggestion so all context is preserved
  const combinedSuggestion = (data.topic ? `📌 TOPIC / SCOPE: ${data.topic}\n` : '') +
    `🎯 GOAL [${data.goalType || 'SEO Optimization'}]: ${data.targetGoal || 'N/A'}\n` +
    `⚡ WHY NOW: ${data.urgencyReason || 'N/A'}\n\n` +
    `📝 SEO SUGGESTION:\n${data.seoSuggestion || ''}`;

  return [
    formattedDate,                                  // Col 1: Date
    "",                                             // Col 2: Date Picked
    data.assignedBy || "SEO Team",                  // Col 3: Assigned By
    "",                                             // Col 4: Picked By
    "",                                             // Col 5: SEO Status (Complies with sheet validation)
    data.category || "",                            // Col 6: Category
    priority,                                       // Col 7: Content Priority (P0, P1, P2)
    data.page || data.topic || "Overview",          // Col 8: Page
    "",                                             // Col 9: Updated Docs
    data.url || data.page || "",                    // Col 10: Url
    data.primaryKeywords || "",                     // Col 11: Primary Keywords
    data.beforeUpdateRank || "",                    // Col 12: Before Update Rank
    "",                                             // Col 13: 17th Jan 2026 Rank
    data.secondaryKeywords || "",                   // Col 14: Seconday Keywords
    data.keywords || "",                            // Col 15: LSI Keywords
    combinedSuggestion,                             // Col 16: SEO Suggestion
    validArea,                                      // Col 17: Area of Optimisation (Strict dropdown)
    data.competitorReference || "",                 // Col 18: Competitors References
    data.keywords || "",                            // Col 19:  Keywords
    "",                                             // Col 20: Suggestion Doc (If an)
    validGap,                                       // Col 21: Content Gap vs Competitors
    validSearchIntent,                              // Col 22: Search Intent (Strict dropdown)
    strategyGoal,                                   // Col 23: Strategy (Goal & Outcome)
    data.urgencyReason || "",                       // Col 24: Comment
    "",                                             // Col 25: Under Review Message ID
    "",                                             // Col 26: Under Review Sent
    ""                                              // Col 27: Under Review Debug
  ];
}

/**
 * Build column values for "New Content" tab
 */
function buildNewContentRow(data) {
  const formattedDate = formatDateWithSuffix(data.date);
  const fullKeywords = [data.primaryKeywords, data.secondaryKeywords, data.keywords]
    .filter(Boolean)
    .join(", ");

  const combinedSuggestion = `🎯 GOAL [${data.goalType || 'New Content'}]: ${data.targetGoal || 'N/A'}\n` +
    `⚡ WHY NOW: ${data.urgencyReason || 'N/A'}\n\n` +
    `📝 SEO SUGGESTIONS:\n${data.seoSuggestion || ''}`;

  return [
    formattedDate,                                  // Col 1: Date
    data.assignedBy || "SEO Team",                  // Col 2: Assigned By
    "",                                             // Col 3: Date Picked
    "",                                             // Col 4: Picked by
    data.contentPriority || "P1",                   // Col 5: Content Priority
    data.category || "",                            // Col 6: Category
    data.topic || data.page || "",                  // Col 7: Page / Topic headline
    data.contentType || "Article",                  // Col 8: Type (News, Article, Notification)
    "",                                             // Col 9: [Blank]
    "",                                             // Col 10: Suggestions Doc.
    data.url || data.page || "",                    // Col 11: Url
    fullKeywords,                                   // Col 12: Keywords
    combinedSuggestion,                             // Col 13: SEO Suggestion
    "Assigned",                                     // Col 14: Status
    "",                                             // Col 15: Approval
    "",                                             // Col 16: DOCS LINK
    data.url || "",                                 // Col 17: URL
    "",                                             // Col 18: Under Review Message ID
    "",                                             // Col 19: Under Review Sent
    "",                                             // Col 20: Under Review Debug
    data.expectedEventDate || ""                    // Col 21: Expected Event Date
  ];
}

/**
 * Insert new row at Row 2 (index 1) in the designated sheet tab
 */
async function assignToSheet(assignmentData) {
  const projectType = assignmentData.projectType || "SEO Optimization";
  const tabMeta = SHEET_METADATA[projectType];
  if (!tabMeta) {
    throw new Error(`Invalid project type: "${projectType}". Must be 'SEO Optimization' or 'New Content'.`);
  }

  const token = await getAccessToken();
  const rowData = projectType === "SEO Optimization"
    ? buildSeoOptimizationRow(assignmentData)
    : buildNewContentRow(assignmentData);

  // 1. Insert blank row at Row 2 (index 1)
  const insertUrl = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}:batchUpdate`;
  const insertRes = await fetch(insertUrl, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      requests: [
        {
          insertDimension: {
            range: {
              sheetId: tabMeta.sheetId,
              dimension: 'ROWS',
              startIndex: 1, // index 1 is Row 2 in Google Sheets
              endIndex: 2
            },
            inheritFromBefore: false
          }
        }
      ]
    })
  });

  const insertJson = await insertRes.json();
  if (insertJson.error) {
    throw new Error(`Google Sheets insertDimension failed: ${insertJson.error.message || JSON.stringify(insertJson.error)}`);
  }

  // 2. Write values into Row 2
  // Compute column letter for end column
  const numCols = rowData.length;
  const colLetter = numCols <= 26 
    ? String.fromCharCode(64 + numCols) 
    : 'A' + String.fromCharCode(64 + (numCols - 26));

  const writeRange = `${projectType}!A2:${colLetter}2`;
  const writeUrl = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(writeRange)}?valueInputOption=USER_ENTERED`;

  const writeRes = await fetch(writeUrl, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      values: [rowData]
    })
  });

  const writeJson = await writeRes.json();
  if (writeJson.error) {
    throw new Error(`Google Sheets write failed: ${writeJson.error.message || JSON.stringify(writeJson.error)}`);
  }

  return {
    success: true,
    message: `Successfully inserted into "${projectType}" at Row 2!`,
    projectType,
    sheetId: tabMeta.sheetId,
    sheetUrl: tabMeta.url,
    rowNumber: 2,
    updatedCells: writeJson.updatedCells,
    rowData
  };
}

/**
 * Health check & status verification
 */
async function checkConnection() {
  const token = await getAccessToken();
  const res = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}?fields=properties.title,sheets.properties(sheetId,title)`, {
    headers: { Authorization: `Bearer ${token}` }
  });

  const data = await res.json();
  if (data.error) {
    throw new Error(data.error.message || JSON.stringify(data.error));
  }

  return {
    connected: true,
    spreadsheetId: SPREADSHEET_ID,
    title: data.properties.title,
    sheets: data.sheets.map(s => ({
      id: s.properties.sheetId,
      title: s.properties.title
    }))
  };
}

/**
 * Get top recent rows from sheet
 */
async function getRecentRows(tabName, limit = 10) {
  const token = await getAccessToken();
  const endRow = limit + 1;
  const range = `${tabName}!A2:Z${endRow}`;
  const res = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(range)}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  const data = await res.json();
  return data.values || [];
}

/**
 * Detect month name (October, November, December) if date falls in OND 2026
 * Returns 'October' | 'November' | 'December' | null
 */
function getONDMonth2026(dateStr) {
  if (!dateStr) return null;
  const s = String(dateStr).trim().toLowerCase();
  const is2026 = s.includes('2026') || s.includes('26') || !s.match(/202[0-57-9]/);
  if (!is2026) return null;

  if (s.includes('oct') || s.includes('-10-') || s.includes('/10/') || s.includes('2026-10') || s.includes('2026/10')) {
    return 'October';
  }
  if (s.includes('nov') || s.includes('-11-') || s.includes('/11/') || s.includes('2026-11') || s.includes('2026/11')) {
    return 'November';
  }
  if (s.includes('dec') || s.includes('-12-') || s.includes('/12/') || s.includes('2026-12') || s.includes('2026/12')) {
    return 'December';
  }
  return null;
}

function isOND2026(dateStr) {
  return getONDMonth2026(dateStr) !== null;
}

function isOctober2026(dateStr) {
  return getONDMonth2026(dateStr) === 'October';
}

/**
 * Fetch and normalize all assignments from OND 2026 (October, November, December)
 * from both 'New Content' and 'SEO Optimization' tabs
 */
async function getONDAssignments() {
  const token = await getAccessToken();

  // Batch get top 300 rows from both tabs simultaneously
  const range1 = encodeURIComponent('SEO Optimization!A1:Z300');
  const range2 = encodeURIComponent('New Content!A1:Z300');
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values:batchGet?ranges=${range1}&ranges=${range2}`;

  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` }
  });
  const data = await res.json();
  if (data.error) {
    throw new Error(data.error.message || JSON.stringify(data.error));
  }

  const seoRows = (data.valueRanges && data.valueRanges[0] && data.valueRanges[0].values) || [];
  const newRows = (data.valueRanges && data.valueRanges[1] && data.valueRanges[1].values) || [];

  const items = [];

  // 1. Process New Content rows
  newRows.slice(1).forEach((r, idx) => {
    const rowNum = idx + 2;
    const dateVal = r[0] || '';
    const month = getONDMonth2026(dateVal);
    if (month) {
      const topic = r[6] || '';
      const page = r[6] || '';
      const url = r[10] || r[16] || '';
      const assignedBy = r[1] || '';
      const pickedBy = r[3] || '';
      const priority = r[4] || 'P1';
      const category = r[5] || '';
      const contentType = r[7] || 'Article';
      const status = r[13] || (pickedBy ? 'Picked' : 'Assigned');
      const gid = SHEET_METADATA["New Content"].gid;

      items.push({
        id: `NEW-ROW-${rowNum}`,
        sheetTab: 'New Content',
        sheetRow: rowNum,
        sheetUrl: `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit?gid=${gid}#range=A${rowNum}`,
        date: dateVal,
        month: month,
        quarter: 'OND 2026',
        datePicked: r[2] || '',
        assignedBy: assignedBy,
        pickedBy: pickedBy,
        priority: priority,
        category: category,
        topic: topic,
        page: page,
        contentType: contentType,
        url: url,
        primaryKeywords: r[11] || '',
        secondaryKeywords: '',
        seoSuggestion: r[12] || '',
        status: status,
        approval: r[14] || '',
        docsLink: r[15] || '',
        expectedEventDate: r[20] || ''
      });
    }
  });

  // 2. Process SEO Optimization rows
  seoRows.slice(1).forEach((r, idx) => {
    const rowNum = idx + 2;
    const dateVal = r[0] || '';
    const month = getONDMonth2026(dateVal);
    if (month) {
      const page = r[7] || '';
      const topic = r[10] || page || 'SEO Optimization Task';
      const url = r[9] || '';
      const assignedBy = r[2] || '';
      const pickedBy = r[3] || '';
      const priority = r[6] || 'P1';
      const category = r[5] || '';
      const status = r[4] || (pickedBy ? 'Picked' : 'Assigned');
      const gid = SHEET_METADATA["SEO Optimization"].gid;

      items.push({
        id: `SEO-ROW-${rowNum}`,
        sheetTab: 'SEO Optimization',
        sheetRow: rowNum,
        sheetUrl: `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit?gid=${gid}#range=A${rowNum}`,
        date: dateVal,
        month: month,
        quarter: 'OND 2026',
        datePicked: r[1] || '',
        assignedBy: assignedBy,
        pickedBy: pickedBy,
        priority: priority,
        category: category,
        topic: topic,
        page: page,
        contentType: 'SEO Optimization',
        url: url,
        primaryKeywords: r[10] || '',
        beforeRank: r[11] || '',
        secondaryKeywords: r[13] || '',
        lsiKeywords: r[14] || '',
        seoSuggestion: r[15] || '',
        areaOfOptimisation: r[16] || '',
        competitorReference: r[17] || '',
        contentGap: r[20] || '',
        searchIntent: r[21] || '',
        strategy: r[22] || '',
        status: status
      });
    }
  });

  // Sort items: lower sheetRow first (since row 2 is the most recent assignment)
  items.sort((a, b) => a.sheetRow - b.sheetRow);

  // Compute summary stats
  const total = items.length;
  const newContentCount = items.filter(i => i.sheetTab === 'New Content').length;
  const seoOptCount = items.filter(i => i.sheetTab === 'SEO Optimization').length;
  const urgentCount = items.filter(i => i.priority === 'P0' || i.priority === 'P1').length;
  const assignedCount = items.filter(i => !i.pickedBy && (i.status === 'Assigned' || !i.status)).length;
  const pickedCount = items.filter(i => i.pickedBy || i.status === 'Picked' || i.status === 'Good to Go' || i.status === 'Live').length;

  const octCount = items.filter(i => i.month === 'October').length;
  const novCount = items.filter(i => i.month === 'November').length;
  const decCount = items.filter(i => i.month === 'December').length;

  const categories = Array.from(new Set(items.map(i => i.category).filter(Boolean))).sort();
  const assignees = Array.from(new Set(items.map(i => i.assignedBy).filter(Boolean))).sort();
  const contentTypes = Array.from(new Set(items.map(i => i.contentType).filter(Boolean))).sort();

  return {
    success: true,
    quarter: 'OND 2026',
    total,
    stats: {
      total,
      newContentCount,
      seoOptCount,
      urgentCount,
      assignedCount,
      pickedCount,
      octCount,
      novCount,
      decCount
    },
    meta: {
      categories,
      assignees,
      contentTypes
    },
    lastUpdated: new Date().toISOString(),
    assignments: items
  };
}

module.exports = {
  SPREADSHEET_ID,
  SHEET_METADATA,
  assignToSheet,
  checkConnection,
  getRecentRows,
  getONDAssignments,
  getOctoberAssignments: getONDAssignments,
  formatDateWithSuffix,
  isOND2026,
  isOctober2026,
  getONDMonth2026
};
