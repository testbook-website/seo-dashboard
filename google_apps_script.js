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
 * 5. Click "Deploy" (blue button, top right) -> "New deployment"
 * 6. Click the gear icon next to "Select type" -> choose "Web app"
 * 7. Set Description: "SEO Webhook"
 * 8. Set "Execute as": "Me"
 * 9. Set "Who has access": "Anyone" (CRITICAL for dashboard access)
 * 10. Click "Deploy" -> "Authorize access" (Advanced -> Go to ... -> Allow)
 * 11. Copy the Web app URL (ends with /exec) into the dashboard if using webhook!
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

function doPost(e) {
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);

    var data = {};
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      data = e.parameter;
    }

    var projectType = data.projectType || "SEO Optimization";
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(projectType);

    if (!sheet) {
      sheet = ss.insertSheet(projectType);
    }

    var formattedDate = formatDateWithSuffix(data.date);
    var row = [];

    if (projectType === "SEO Optimization") {
      var strategyGoal = "[" + (data.goalType || "Goal") + "]: " + (data.targetGoal || "") + " | Why Now: " + (data.urgencyReason || "");

      // Search Intent strict validation (Col 22)
      var validSearchIntent = "";
      var rawIntent = data.searchIntent || "";
      if (rawIntent.indexOf("Mixed") !== -1) {
        validSearchIntent = "Mixed → balance info + action";
      } else if (rawIntent.indexOf("Transactional") !== -1) {
        validSearchIntent = "Transactional → highlight benefits, CTAs";
      } else if (rawIntent.indexOf("Informational") !== -1) {
        validSearchIntent = "Informational → explain, guide, clarify";
      } else if (rawIntent === "") {
        validSearchIntent = "Informational → explain, guide, clarify";
      } else {
        validSearchIntent = rawIntent;
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

      row = [
        formattedDate,                          // Col 1: Date
        data.assignedBy || "SEO Team",          // Col 2: Assigned By
        "",                                     // Col 3: Date Picked
        "",                                     // Col 4: Picked by
        data.contentPriority || "P1",           // Col 5: Content Priority
        data.category || "",                    // Col 6: Category
        data.topic || data.page || "",          // Col 7: Page / Headline
        data.contentType || "Article",          // Col 8: Type
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
      tab: projectType
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
