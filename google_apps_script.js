/**
 * ============================================================================
 * RANKASSIGN PRO - STANDALONE GOOGLE APPS SCRIPT WEBHOOK (ROW 2 INSERTION)
 * ============================================================================
 * 
 * Target Google Spreadsheet:
 * https://docs.google.com/spreadsheets/d/1ihLeB9ZOJdaF841qGLoTWBULSRNsF9BjtxUXRxKuK2A/edit
 * 
 * Target Tabs:
 * 1. "SEO Optimization" (gid: 1073261968)
 * 2. "New Content" (gid: 1107724151)
 * 
 * INSTRUCTIONS FOR NEW STANDALONE APPS SCRIPT PROJECT:
 * 1. Go to https://script.google.com/
 * 2. Click "+ New project" (top left)
 * 3. Give your project a title: "SEO Assignment Webhook"
 * 4. Delete everything in Code.gs and paste this entire code.
 * 5. Click "Deploy" (blue button, top right) -> "New deployment" (or "Manage deployments" -> Edit -> "New version" if updating)
 * 6. Click the gear icon next to "Select type" -> choose "Web app"
 * 7. Set Description: "SEO Webhook Row 2 Insertion"
 * 8. Set "Execute as": "Me"
 * 9. Set "Who has access": "Anyone" (CRITICAL for dashboard access)
 * 10. Click "Deploy" -> "Authorize access" (Advanced -> Go to ... -> Allow)
 * 11. To test immediately inside Apps Script editor: select "testInsertNewContent" in toolbar and click "Run"!
 * ============================================================================
 */

// Target Spreadsheet ID
var SPREADSHEET_ID = "1ihLeB9ZOJdaF841qGLoTWBULSRNsF9BjtxUXRxKuK2A";

function getSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.length > 5) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function doGet(e) {
  try {
    var ss = getSpreadsheet();
    var action = (e && e.parameter && e.parameter.action) || "";
    
    // Support pulling OND 2026 (October, November, December) assignments directly via Webhook
    if (action === "getONDAssignments" || action === "ond" || action === "getOctoberAssignments" || action === "october") {
      var items = [];

      function getMonthOND(dateVal) {
        var s = String(dateVal || "").toLowerCase();
        var is2026 = s.indexOf("2026") !== -1 || s.indexOf("26") !== -1 || !s.match(/202[0-57-9]/);
        if (!is2026) return null;
        if (s.indexOf("oct") !== -1 || s.indexOf("2026-10") !== -1 || s.indexOf("/10/2026") !== -1 || s.indexOf("-10-2026") !== -1 || s.indexOf("/10/") !== -1) return "October";
        if (s.indexOf("nov") !== -1 || s.indexOf("2026-11") !== -1 || s.indexOf("/11/2026") !== -1 || s.indexOf("-11-2026") !== -1 || s.indexOf("/11/") !== -1) return "November";
        if (s.indexOf("dec") !== -1 || s.indexOf("2026-12") !== -1 || s.indexOf("/12/2026") !== -1 || s.indexOf("-12-2026") !== -1 || s.indexOf("/12/") !== -1) return "December";
        return null;
      }

      var sheetNew = ss.getSheetByName("New Content");
      if (sheetNew && sheetNew.getLastRow() > 1) {
        var numRows = Math.min(sheetNew.getLastRow() - 1, 300);
        var newValues = sheetNew.getRange(2, 1, numRows, 21).getValues();
        newValues.forEach(function(r, idx) {
          var dateVal = String(r[0] || "");
          var month = getMonthOND(dateVal);
          if (month) {
            items.push({
              id: "NEW-ROW-" + (idx + 2),
              sheetTab: "New Content",
              sheetRow: idx + 2,
              date: dateVal,
              month: month,
              quarter: "OND 2026",
              assignedBy: r[1] || "",
              datePicked: r[2] || "",
              pickedBy: r[3] || "",
              priority: r[4] || "P1",
              category: r[5] || "",
              topic: r[6] || "",
              page: r[6] || "",
              contentType: r[7] || "Article",
              url: r[10] || r[16] || "",
              primaryKeywords: r[11] || "",
              seoSuggestion: r[12] || "",
              status: r[13] || (r[3] ? "Picked" : "Assigned"),
              expectedEventDate: r[20] || ""
            });
          }
        });
      }

      var sheetSeo = ss.getSheetByName("SEO Optimization");
      if (sheetSeo && sheetSeo.getLastRow() > 1) {
        var numSeoRows = Math.min(sheetSeo.getLastRow() - 1, 300);
        var seoValues = sheetSeo.getRange(2, 1, numSeoRows, 26).getValues();
        seoValues.forEach(function(r, idx) {
          var dateVal = String(r[0] || "");
          var month = getMonthOND(dateVal);
          if (month) {
            items.push({
              id: "SEO-ROW-" + (idx + 2),
              sheetTab: "SEO Optimization",
              sheetRow: idx + 2,
              date: dateVal,
              month: month,
              quarter: "OND 2026",
              datePicked: r[1] || "",
              assignedBy: r[2] || "",
              pickedBy: r[3] || "",
              status: r[4] || (r[3] ? "Picked" : "Assigned"),
              category: r[5] || "",
              priority: r[6] || "P1",
              topic: r[10] || r[7] || "SEO Optimization",
              page: r[7] || "",
              contentType: "SEO Optimization",
              url: r[9] || "",
              primaryKeywords: r[10] || "",
              beforeRank: r[11] || "",
              secondaryKeywords: r[13] || "",
              seoSuggestion: r[15] || "",
              strategy: r[22] || ""
            });
          }
        });
      }

      var octCount = items.filter(function(i){ return i.month === 'October'; }).length;
      var novCount = items.filter(function(i){ return i.month === 'November'; }).length;
      var decCount = items.filter(function(i){ return i.month === 'December'; }).length;

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        quarter: "OND 2026",
        total: items.length,
        stats: {
          total: items.length,
          octCount: octCount,
          novCount: novCount,
          decCount: decCount,
          newContentCount: items.filter(function(i){ return i.sheetTab === 'New Content'; }).length,
          seoOptCount: items.filter(function(i){ return i.sheetTab === 'SEO Optimization'; }).length,
          urgentCount: items.filter(function(i){ return i.priority === 'P0' || i.priority === 'P1'; }).length,
          assignedCount: items.filter(function(i){ return !i.pickedBy && (i.status === 'Assigned' || !i.status); }).length,
          pickedCount: items.filter(function(i){ return i.pickedBy || i.status === 'Picked' || i.status === 'Good to Go' || i.status === 'Live'; }).length
        },
        assignments: items
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Check Duplicate API endpoint (Remote live check)
    if (action === "checkDuplicate" || action === "check") {
      var dupCheck = checkDuplicateInSheetAppsScript(ss, {
        url: (e && e.parameter && e.parameter.url) || "",
        page: (e && e.parameter && e.parameter.page) || "",
        topic: (e && e.parameter && e.parameter.topic) || "",
        date: (e && e.parameter && e.parameter.date) || ""
      });
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        isDuplicate: dupCheck.isDuplicate,
        isWarning: dupCheck.isWarning,
        allowed: dupCheck.allowed,
        severity: dupCheck.severity,
        daysDiff: dupCheck.daysDiff,
        message: dupCheck.message,
        existing: dupCheck.existing || null
      })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "SEO Assignment Webhook is active and connected to '" + ss.getName() + "'!",
      spreadsheetId: SPREADSHEET_ID,
      tabs: ["SEO Optimization", "New Content"],
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function formatDateWithSuffix(dateStr) {
  if (!dateStr) dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var parts = dateStr.split('-');
  var d;
  if (parts.length === 3) {
    d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  } else {
    d = new Date(dateStr);
  }

  var day = d.getDate();
  var monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  var month = monthNames[d.getMonth()];
  var year = d.getFullYear();

  var suffix = "th";
  if (day === 1 || day === 21 || day === 31) suffix = "st";
  else if (day === 2 || day === 22) suffix = "nd";
  else if (day === 3 || day === 23) suffix = "rd";

  return day + suffix + " " + month + " " + year;
}

// ============================================================================
// 60-DAY DUPLICATE PREVENTION & ANTI-CANNIBALIZATION SHIELD
// ============================================================================

var GENERIC_SLUGS_SET = {
  "overview": 1, "home": 1, "index": 1, "page": 1, "article": 1, "news": 1, "default": 1,
  "syllabus": 1, "exam-pattern": 1, "notification": 1, "admit-card": 1, "result": 1,
  "answer-key": 1, "cut-off": 1, "eligibility": 1, "salary": 1, "selection-process": 1
};

function normalizeUrlOrSlugAppsScript(raw) {
  if (!raw) return "";
  var str = String(raw).trim().toLowerCase();
  str = str.replace(/^(https?:\/\/)?(www\.)?testbook\.com\/?/i, "");
  str = str.replace(/^(https?:\/\/)/i, "");
  str = str.replace(/[?#].*$/, "");
  str = str.replace(/^\/+|\/+$/g, "");
  return str;
}

function normalizeTitleAppsScript(raw) {
  if (!raw) return "";
  return String(raw)
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function parseSheetDateAppsScript(str) {
  if (!str) return null;
  if (Object.prototype.toString.call(str) === '[object Date]') {
    return isNaN(str.getTime()) ? null : str;
  }
  var s = String(str).trim();
  var match = s.match(/^(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)\s+(\d{4})$/i);
  if (match) {
    var day = parseInt(match[1], 10);
    var monthName = match[2].toLowerCase();
    var year = parseInt(match[3], 10);
    var monthNames = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];
    var mIdx = monthNames.indexOf(monthName);
    if (mIdx !== -1) {
      return new Date(year, mIdx, day);
    }
  }
  var iso = new Date(s);
  if (!isNaN(iso.getTime())) return iso;
  return null;
}

function checkDuplicateInSheetAppsScript(ss, data) {
  var targetUrl = normalizeUrlOrSlugAppsScript(data.url || (data.page && String(data.page).indexOf("http") !== -1 ? data.page : ""));
  var targetPage = normalizeUrlOrSlugAppsScript(data.page || "");
  var targetTopic = normalizeTitleAppsScript(data.topic || "");

  if (!targetUrl && !targetPage && !targetTopic) {
    return { isDuplicate: false, isWarning: false, allowed: true, severity: "CLEAN", message: "No URL or topic provided." };
  }

  var targetDate = data.date ? (parseSheetDateAppsScript(data.date) || new Date()) : new Date();
  var sixtyDaysMs = 60 * 24 * 60 * 60 * 1000;

  var tabs = [
    { name: "SEO Optimization", urlCol: 10, pageCol: 8, topicCol: 8, dateCol: 1, userCol: 3, maxCols: 16 },
    { name: "New Content", urlCol: 11, pageCol: 7, topicCol: 7, dateCol: 1, userCol: 2, maxCols: 15 }
  ];

  var matched = null;

  for (var t = 0; t < tabs.length; t++) {
    var tabInfo = tabs[t];
    var sheet = ss.getSheetByName(tabInfo.name);
    if (!sheet || sheet.getLastRow() <= 1) continue;

    var rowsToCheck = Math.min(sheet.getLastRow() - 1, 350);
    var values = sheet.getRange(2, 1, rowsToCheck, tabInfo.maxCols).getValues();

    for (var r = 0; r < values.length; r++) {
      var row = values[r];
      var rawDate = row[tabInfo.dateCol - 1];
      var d = parseSheetDateAppsScript(rawDate);
      if (!d) continue;

      var ageMs = Math.abs(targetDate.getTime() - d.getTime());
      if (ageMs > sixtyDaysMs) continue; // Scan past 60 days only

      var rowUrl = normalizeUrlOrSlugAppsScript(row[tabInfo.urlCol - 1]);
      var rowPage = normalizeUrlOrSlugAppsScript(row[tabInfo.pageCol - 1]);
      var rowTopic = normalizeTitleAppsScript(row[tabInfo.topicCol - 1]);

      // If both have explicit URLs and they are different, skip
      if (targetUrl && rowUrl && targetUrl !== rowUrl) {
        continue;
      }

      var isMatch = false;
      var reason = "";

      if (targetUrl && rowUrl && targetUrl === rowUrl && !GENERIC_SLUGS_SET[targetUrl] && targetUrl.length >= 4) {
        isMatch = true;
        reason = "Exact URL match";
      } else if (targetUrl && !GENERIC_SLUGS_SET[targetUrl] && targetUrl.length >= 4 && rowPage && !GENERIC_SLUGS_SET[rowPage] && rowPage.length >= 4 && targetUrl === rowPage) {
        isMatch = true;
        reason = "URL matches previously assigned page";
      } else if (targetPage && !GENERIC_SLUGS_SET[targetPage] && targetPage.length >= 4 && rowPage && !GENERIC_SLUGS_SET[rowPage] && rowPage.length >= 4 && targetPage === rowPage) {
        isMatch = true;
        reason = "Page slug match";
      } else if (targetTopic && targetTopic.length > 12 && rowTopic && targetTopic === rowTopic) {
        isMatch = true;
        reason = "Identical topic headline";
      }

      if (isMatch) {
        var daysDiff = Math.max(0, Math.round(ageMs / (24 * 60 * 60 * 1000)));
        matched = {
          sheetTab: tabInfo.name,
          sheetRow: r + 2,
          date: String(rawDate),
          daysDiff: daysDiff,
          assignedBy: String(row[tabInfo.userCol - 1] || "SEO Team"),
          topic: String(row[tabInfo.topicCol - 1] || ""),
          reason: reason
        };
        break;
      }
    }
    if (matched) break;
  }

  if (!matched) {
    return {
      isDuplicate: false,
      isWarning: false,
      allowed: true,
      severity: "CLEAN",
      message: "Verified Clean: No duplicate found in the past 60 days."
    };
  }

  if (matched.daysDiff <= 30) {
    return {
      isDuplicate: true,
      isWarning: false,
      allowed: false,
      severity: "BLOCKED",
      daysDiff: matched.daysDiff,
      existing: matched,
      message: "🚫 Duplicate Assignment Blocked: This page was already assigned " + (matched.daysDiff === 0 ? "today" : matched.daysDiff + " day" + (matched.daysDiff === 1 ? "" : "s") + " ago") + " (on " + matched.date + ") in \"" + matched.sheetTab + "\" by " + matched.assignedBy + " for topic \"" + matched.topic + "\". Re-assigning within 30 days is strictly prohibited to prevent content cannibalization."
    };
  }

  return {
    isDuplicate: false,
    isWarning: true,
    allowed: true,
    severity: "WARNING",
    daysDiff: matched.daysDiff,
    existing: matched,
    message: "⚠️ Re-optimization Notice: This page was previously assigned " + matched.daysDiff + " days ago (on " + matched.date + ") in \"" + matched.sheetTab + "\". Cooldown period (30 days) has passed, but ensure new goals are defined."
  };
}

function doPost(e) {
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);

    var data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter && Object.keys(e.parameter).length > 0) {
      data = e.parameter;
    } else {
      // Default sample for manual "Run" test inside Apps Script Editor
      data = {
        projectType: "New Content",
        date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
        topic: "SBI PO 2026 Notification & Syllabus Guide",
        category: "Banking Exams",
        contentPriority: "P1",
        contentType: "News",
        page: "https://testbook.com/sbi-po-2026",
        primaryKeywords: "sbi po 2026 notification",
        seoSuggestion: "Verified Row 2 Insertion from Apps Script Console",
        assignedBy: "SEO Team"
      };
    }

    var projectType = data.projectType || "SEO Optimization";
    var ss = getSpreadsheet();

    // ========================================================================
    // STRICT 30-DAY DUPLICATE PREVENTION & ANTI-CANNIBALIZATION SHIELD
    // ========================================================================
    var dupCheck = checkDuplicateInSheetAppsScript(ss, data);
    if (!dupCheck.allowed || dupCheck.isDuplicate) {
      lock.releaseLock();
      return ContentService.createTextOutput(JSON.stringify({
        status: "error",
        duplicate: true,
        allowed: false,
        severity: dupCheck.severity,
        daysDiff: dupCheck.daysDiff,
        message: dupCheck.message,
        details: dupCheck
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var sheet = ss.getSheetByName(projectType);

    if (!sheet) {
      sheet = ss.insertSheet(projectType);
    }

    var formattedDate = formatDateWithSuffix(data.date);
    var row = [];

    if (projectType === "SEO Optimization") {
      var strategyGoal = "[" + (data.goalType || "Goal") + "]: " + (data.targetGoal || "") + " | Why Now: " + (data.urgencyReason || "");

      // Search Intent strict validation (Col 22)
      var validSearchIntent = "Informational → explain, guide, clarify";
      var rawIntent = data.searchIntent || "";
      if (rawIntent.indexOf("Mixed") !== -1) {
        validSearchIntent = "Mixed → balance info + action";
      } else if (rawIntent.indexOf("Transactional") !== -1) {
        validSearchIntent = "Transactional → highlight benefits, CTAs";
      } else if (rawIntent.indexOf("Informational") !== -1) {
        validSearchIntent = "Informational → explain, guide, clarify";
      }

      // Area of Optimisation strict validation (Col 17)
      var allowedAreas = [
        "Content Outdated",
        "Factual Information Missing",
        "Missing Clarity",
        "Need To Work on Intro",
        "Page Structure",
        "Too Generic",
        "Update as per the expected notification."
      ];
      var validArea = (data.areaOfOptimisation && allowedAreas.indexOf(data.areaOfOptimisation) !== -1) ? data.areaOfOptimisation : "";

      // Content Gap strict validation (Col 21)
      var allowedGaps = [
        "Missing Examples",
        "Missing Explanations",
        "Missing FAQs",
        "Missing Sections",
        "Missing Tables / Lists"
      ];
      var validGap = (data.contentGap && allowedGaps.indexOf(data.contentGap) !== -1) ? data.contentGap : "";

      // Priority clamp for SEO Optimization (Col 7: P0, P1, P2 only)
      var priority = data.contentPriority || "P1";
      if (priority === "P3") priority = "P2";

      // Prepend topic and goal to SEO Suggestion so all information is preserved
      var combinedSuggestion = (data.topic ? "📌 TOPIC / SCOPE: " + data.topic + "\n" : "") +
        "🎯 GOAL [" + (data.goalType || "SEO Optimization") + "]: " + (data.targetGoal || "N/A") + "\n" +
        "⚡ WHY NOW: " + (data.urgencyReason || "N/A") + "\n\n" +
        "📝 SEO SUGGESTION:\n" + (data.seoSuggestion || "");

      row = [
        formattedDate,                          // Col 1: Date
        "",                                     // Col 2: Date Picked
        data.assignedBy || "SEO Team",          // Col 3: Assigned By
        "",                                     // Col 4: Picked By
        "",                                     // Col 5: SEO Status (Complies with sheet validation: '', 'Pending', etc.)
        data.category || "",                    // Col 6: Category
        priority,                               // Col 7: Content Priority (P0, P1, P2)
        data.page || data.topic || "Overview",  // Col 8: Page
        "",                                     // Col 9: Updated Docs
        data.url || data.page || "",            // Col 10: Url
        data.primaryKeywords || "",             // Col 11: Primary Keywords
        data.beforeUpdateRank || "",            // Col 12: Before Update Rank
        "",                                     // Col 13: 17th Jan 2026 Rank
        data.secondaryKeywords || "",           // Col 14: Seconday Keywords
        data.keywords || "",                    // Col 15: LSI Keywords
        combinedSuggestion,                     // Col 16: SEO Suggestion
        validArea,                              // Col 17: Area of Optimisation (Strict validation)
        data.competitorReference || "",         // Col 18: Competitors References
        data.keywords || "",                    // Col 19:  Keywords
        "",                                     // Col 20: Suggestion Doc (If an)
        validGap,                               // Col 21: Content Gap vs Competitors
        validSearchIntent,                      // Col 22: Search Intent (Strict validation)
        strategyGoal,                           // Col 23: Strategy
        data.urgencyReason || "",               // Col 24: Comment
        "",                                     // Col 25: Under Review Message ID
        "",                                     // Col 26: Under Review Sent
        ""                                      // Col 27: Under Review Debug
      ];
    } else {
      // New Content
      var fullKeywords = [data.primaryKeywords, data.secondaryKeywords, data.keywords].filter(Boolean).join(", ");
      var combinedSuggestion = "🎯 GOAL [" + (data.goalType || "New Content") + "]: " + (data.targetGoal || "N/A") + "\n" +
        "⚡ WHY NOW: " + (data.urgencyReason || "N/A") + "\n\n" +
        "📝 SEO SUGGESTIONS:\n" + (data.seoSuggestion || "");

      // Priority clamp for New Content (Col 5: P0, P1, P2 only)
      var newPriority = data.contentPriority || "P1";
      if (newPriority === "P3") newPriority = "P2";

      // Type strict validation for New Content (Col 8: "Child Page", "Blog", "News", "SEO Group", "Discover")
      var allowedTypes = ["Child Page", "Blog", "News", "SEO Group", "Discover"];
      var validContentType = "News";
      if (data.contentType && allowedTypes.indexOf(data.contentType) !== -1) {
        validContentType = data.contentType;
      } else if (data.contentType === "Article") {
        validContentType = "News";
      }

      row = [
        formattedDate,                          // Col 1: Date
        data.assignedBy || "SEO Team",          // Col 2: Assigned By
        "",                                     // Col 3: Date Picked
        "",                                     // Col 4: Picked by
        newPriority,                            // Col 5: Content Priority (P0, P1, P2)
        data.category || "",                    // Col 6: Category
        data.topic || data.page || "",          // Col 7: Page / Headline
        validContentType,                       // Col 8: Type (Strictly News, Blog, Child Page, SEO Group, Discover)
        "",                                     // Col 9: [Blank]
        "",                                     // Col 10: Suggestions Doc.
        data.url || data.page || "",            // Col 11: Url
        fullKeywords,                           // Col 12: Keywords
        combinedSuggestion,                     // Col 13: SEO Suggestion
        "Assigned",                             // Col 14: Status
        "",                                     // Col 15: Approval
        "",                                     // Col 16: DOCS LINK
        data.url || "",                         // Col 17: URL
        "",                                     // Col 18: Under Review Message ID
        "",                                     // Col 19: Under Review Sent
        "",                                     // Col 20: Under Review Debug
        data.expectedEventDate || ""            // Col 21: Expected Event Date
      ];
    }

    // Insert at Row 2 (Immediately beneath header row 1)
    sheet.insertRowBefore(2);
    var newRowRange = sheet.getRange(2, 1, 1, row.length);
    newRowRange.setValues([row]);
    newRowRange.setVerticalAlignment("middle");
    newRowRange.setWrap(true);

    lock.releaseLock();

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Successfully inserted into '" + projectType + "' at Row 2!",
      rowNumber: 2,
      tab: projectType,
      insertedData: {
        date: formattedDate,
        topic: data.topic,
        category: data.category,
        priority: projectType === "SEO Optimization" ? priority : newPriority,
        type: projectType === "SEO Optimization" ? "SEO Optimization" : validContentType
      }
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ============================================================================
 * ONE-CLICK TEST FUNCTIONS (Select function from dropdown and click "Run")
 * ============================================================================
 */
function testInsertNewContent() {
  var result = doPost({
    postData: {
      contents: JSON.stringify({
        projectType: "New Content",
        date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
        topic: "CSIR NET Physical Sciences 2026 Notification & Syllabus",
        category: "CSIR NET",
        contentPriority: "P0",
        contentType: "News",
        page: "CSIR NET Physical Sciences 2026",
        url: "https://testbook.com/csir-net-physical-sciences",
        primaryKeywords: "csir net physical sciences 2026 syllabus",
        secondaryKeywords: "csir net exam date, eligibility criteria",
        keywords: "nta csir net 2026 pdf download",
        seoSuggestion: "1. Urgent news release\n2. Add official PDF link\n3. Table with exam schedule",
        assignedBy: "SEO Team",
        goalType: "New Exam Notification",
        targetGoal: "Top 3 SERP rank within 24h",
        urgencyReason: "Notification announced today"
      })
    }
  });
  Logger.log("Result: " + result.getContent());
  return result.getContent();
}

function testInsertSEOOptimization() {
  var result = doPost({
    postData: {
      contents: JSON.stringify({
        projectType: "SEO Optimization",
        date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
        topic: "Bihar Police Constable Syllabus & Exam Pattern 2026",
        category: "Bihar Gov. Exam",
        contentPriority: "P1",
        page: "/bihar-gov-exams/police-constable-syllabus",
        url: "https://testbook.com/bihar-police-constable-syllabus",
        primaryKeywords: "bihar police constable syllabus 2026",
        beforeUpdateRank: "#12",
        secondaryKeywords: "bihar police exam pattern, subject wise marks",
        keywords: "csbc bihar police syllabus pdf",
        seoSuggestion: "1. Update 2026 revised syllabus\n2. Add Subject-wise weightage table\n3. 5 FAQs with schema",
        assignedBy: "SEO Team",
        goalType: "Rank Jump (Top 3)",
        targetGoal: "Move from #12 to Top 3 within 21 days",
        urgencyReason: "Competitors updated syllabus table"
      })
    }
  });
  Logger.log("Result: " + result.getContent());
  return result.getContent();
}

function testDuplicatePrevention() {
  // Test with known recently assigned page (IOB SO Admit Card)
  var result = doPost({
    postData: {
      contents: JSON.stringify({
        projectType: "SEO Optimization",
        date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
        topic: "IOB SO Admit Card 2026 Test",
        category: "Banking Exams",
        contentPriority: "P0",
        page: "/iob-so/admit-card",
        url: "https://testbook.com/iob-so/admit-card",
        primaryKeywords: "iob so admit card 2026",
        seoSuggestion: "Test duplicate prevention",
        assignedBy: "Test User"
      })
    }
  });
  Logger.log("Duplicate Test Result (Should be blocked): " + result.getContent());
  return result.getContent();
}

