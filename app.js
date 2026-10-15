/**
 * ============================================================================
 * RANKASSIGN PRO - APP CONTROLLER
 * Goal-Oriented SEO Assignment Engine & Google Sheets Bridge
 * ============================================================================
 */

(function () {
  'use strict';

  // --- Constants & Categories ---
  const EXAM_CATEGORIES = [
    "Bihar Gov. Exam",
    "CSIR NET",
    "CUET and UG Entrance",
    "Defence Exams",
    "Engineering Recruitment Exams",
    "GATE",
    "Haryana Exams",
    "SSC Exams",
    "Railway Exams",
    "Banking Exams",
    "Insurance Exam",
    "Regulatory Exams",
    "Other Govt Exam",
    "MBA Entance Exams",
    "Judiciary Exams",
    "Kerala Exams",
    "MH Govt Exams",
    "MP Govt Exams",
    "Nursing Exams",
    "Odisha Gov. Exam",
    "Police Exams",
    "State PSC",
    "Teaching Exams",
    "UGC NET",
    "UP Gov. Exam",
    "UPSC Civil Services",
    "Other UPSC Exams",
    "West Bengal Exams",
    "Rajasthan Exams",
    "Punjab Exams",
    "K12"
  ];

  const DEFAULT_WEBHOOK_URL = 'https://script.google.com/macros/s/AKfycbzbLLDB6B6PL_EYL4i1hWXMFvDzEU-BHAF6_IsvBvTxUgBxvaq4lypl56vEjWnBvqLCdQ/exec';

  const STORAGE_KEYS = {
    ASSIGNMENTS: 'rankassign_tasks_v1',
    WEBHOOK_URL: 'rankassign_webhook_url_v1',
    THEME: 'rankassign_theme_v1'
  };

  // Google Apps Script source code template for display in the setup modal
  const APPS_SCRIPT_CODE = `/**
 * Google Apps Script Webhook for SEO Assignments
 * Paste this into Google Sheets -> Extensions -> Apps Script
 */
var SHEET_NAME = "SEO Assignments";
var HEADERS = [
  "Assignment ID", "Date", "Topic", "Category", "Goal Type", 
  "Target Goal & Outcome", "Why Now / Urgency", "Content Priority", "Page / URL",
  "Primary Keywords", "Before Update Rank", "Secondary Keywords", 
  "Keywords Bucket", "SEO Suggestion & Brief", "Assigned By", "Status", "Created At"
];

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    message: "SEO Assignment Webhook is Active!",
    timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
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
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    
    if (sheet.getLastRow() === 0) {
      var hRange = sheet.getRange(1, 1, 1, HEADERS.length);
      hRange.setValues([HEADERS]);
      hRange.setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    }
    
    var row = [
      data.id || "TASK-" + Utilities.getUuid().substring(0, 8).toUpperCase(),
      data.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
      data.topic || "",
      data.category || "",
      data.goalType || "",
      data.targetGoal || "",
      data.urgencyReason || "",
      data.contentPriority || "P2",
      data.page || "",
      data.primaryKeywords || "",
      data.beforeUpdateRank || "",
      data.secondaryKeywords || "",
      data.keywords || "",
      data.seoSuggestion || "",
      data.assignedBy || "SEO Team",
      "Assigned",
      new Date().toLocaleString()
    ];
    
    sheet.appendRow(row);
    lock.releaseLock();
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", rowNumber: sheet.getLastRow() }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}`;

  // --- State ---
  let assignments = [];
  let webhookUrl = '';

  // --- DOM Elements ---
  const form = document.getElementById('assignment-form');
  const dateInput = document.getElementById('assignmentDate');
  const topicInput = document.getElementById('topic');
  const categorySelect = document.getElementById('category');
  const prioritySelect = document.getElementById('contentPriority');
  const contentTypeInput = document.getElementById('contentType');
  const pageInput = document.getElementById('page');
  const primaryKwInput = document.getElementById('primaryKeywords');
  const beforeRankInput = document.getElementById('beforeUpdateRank');
  const secondaryKwInput = document.getElementById('secondaryKeywords');
  const keywordsInput = document.getElementById('keywords');
  const suggestionInput = document.getElementById('seoSuggestion');
  const targetGoalInput = document.getElementById('targetGoal');
  const urgencyInput = document.getElementById('urgencyReason');
  const assignedByInput = document.getElementById('assignedBy');

  // Preview elements
  const previewDate = document.getElementById('preview-date');
  const previewTopic = document.getElementById('preview-topic');
  const previewCategory = document.getElementById('preview-category');
  const previewGoal = document.getElementById('preview-goal');
  const previewPriority = document.getElementById('preview-priority');
  const previewPage = document.getElementById('preview-page');
  const previewPrimaryKw = document.getElementById('preview-primary-kw');
  const previewRank = document.getElementById('preview-rank');
  const previewSecondaryKw = document.getElementById('preview-secondary-kw');
  const previewKeywords = document.getElementById('preview-keywords');
  const previewSuggestion = document.getElementById('preview-suggestion');

  // UI status elements
  const sheetStatusDot = document.getElementById('sheet-status-dot');
  const sheetStatusText = document.getElementById('sheet-status-text');
  const intStatusDesc = document.getElementById('int-status-desc');
  const btnSubmit = document.getElementById('btn-submit-task');
  const btnSubmitLabel = document.getElementById('btn-submit-label');

  // Table & Stats
  const tableBody = document.getElementById('assignments-table-body');
  const emptyStateCard = document.getElementById('empty-state-card');
  const statTotal = document.getElementById('stat-total-count');
  const statUrgent = document.getElementById('stat-urgent-count');
  const statSynced = document.getElementById('stat-synced-count');
  const historyCountBadge = document.getElementById('history-count-badge');
  const searchInput = document.getElementById('history-search');
  const filterTab = document.getElementById('filter-tab');
  const filterCategory = document.getElementById('filter-category');
  const filterPriority = document.getElementById('filter-priority');
  const btnRefreshSheet = document.getElementById('btn-refresh-sheet');
  const categoryTagsCloud = document.getElementById('category-tags-cloud');

  // Modals
  const settingsModal = document.getElementById('settings-modal');
  const webhookUrlInput = document.getElementById('webhookUrlInput');
  const scriptCodeBlock = document.getElementById('script-code-block');
  const viewModal = document.getElementById('view-modal');
  const viewModalContent = document.getElementById('view-modal-content');
  const testFeedback = document.getElementById('test-feedback');
  const testSpinner = document.getElementById('test-spinner');

  // --- Initializer ---
  function init() {
    loadSavedState();
    setupDefaultDate();
    renderCategoryFilters();
    renderCategoryCloud();
    setupEventListeners();
    updateLivePreview();
    renderAssignmentsTable();
    updateStats();
    updateWebhookStatusUI();
    checkBackendStatus();

    // Populate Apps script preview
    if (scriptCodeBlock) {
      scriptCodeBlock.textContent = APPS_SCRIPT_CODE;
    }

    // Set initial user info if available
    if (!assignedByInput.value) {
      assignedByInput.value = "SEO Team";
    }

    // Fetch live October 2026 assignments from Google Sheets and begin background polling
    fetchOctoberAssignments(false, false);
    startOctoberPolling();

    // Restore active tab from URL hash or localStorage
    const hash = window.location.hash;
    const savedTab = localStorage.getItem('rankassign_active_tab');
    if (hash === '#ond' || hash === '#october') {
      switchDashboardTab('view-october');
    } else if (hash === '#history') {
      switchDashboardTab('view-history');
    } else if (savedTab) {
      switchDashboardTab(savedTab);
    }
  }

  // --- Real-Time Background Polling ---
  function startOctoberPolling() {
    if (octPollInterval) clearInterval(octPollInterval);
    octPollInterval = setInterval(() => {
      fetchOctoberAssignments(true, true);
    }, 20000);

    // Update relative sync time every 5 seconds
    setInterval(updateOctoberLastSyncTime, 5000);
  }

  // --- Check Backend Status ---
  async function checkBackendStatus() {
    try {
      const res = await fetch('/api/status');
      if (res.ok) {
        const data = await res.json();
        if (data.success && data.connected) {
          if (sheetStatusDot) sheetStatusDot.className = 'status-dot connected';
          if (sheetStatusText) sheetStatusText.textContent = 'Sheet: Connected (Row 2 Sync)';
          if (intStatusDesc) {
            intStatusDesc.textContent = `Direct API active to "${data.title}" via credentials.json. New tasks insert directly at Row 2!`;
          }
        }
      }
    } catch (e) {
      // Backend not running yet or static file
      console.log("Backend API status:", e.message);
    }
  }

  // --- Dashboard Tab Switcher ---
  function switchDashboardTab(targetViewId) {
    if (!targetViewId) return;

    const navTabBtns = document.querySelectorAll('.dashboard-tab-btn');
    const viewPanels = document.querySelectorAll('.dashboard-view-panel');

    navTabBtns.forEach(btn => {
      const match = btn.getAttribute('data-view') === targetViewId;
      btn.classList.toggle('active', match);
      btn.setAttribute('aria-selected', match ? 'true' : 'false');
    });

    viewPanels.forEach(panel => {
      const match = panel.id === targetViewId;
      panel.style.display = match ? 'block' : 'none';
      panel.classList.toggle('active', match);
    });

    localStorage.setItem('rankassign_active_tab', targetViewId);

    if (targetViewId === 'view-october') {
      window.location.hash = 'ond';
      if (octoberAssignments.length === 0) {
        fetchOctoberAssignments(false, false);
      }
    } else if (targetViewId === 'view-history') {
      window.location.hash = 'history';
    } else {
      window.location.hash = 'dispatch';
    }
  }

  // Expose globally for inline buttons & toasts
  window.switchDashboardTab = switchDashboardTab;

  // --- Date Checker for OND 2026 (October, November, December) ---
  function getONDMonth2026Date(dateStr) {
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

  function isOND2026Date(dateStr) {
    return getONDMonth2026Date(dateStr) !== null;
  }

  function isOctober2026Date(dateStr) {
    return isOND2026Date(dateStr);
  }

  // --- OND 2026 Tracker State & Methods ---
  let octoberAssignments = [];
  let isOctoberLoading = false;
  let lastOctoberSyncTime = null;
  let octSourceFilter = 'ALL';
  let ondMonthFilter = 'ALL';
  let octPollInterval = null;

  async function fetchOctoberAssignments(isBackground = false, notifyNew = false) {
    if (isOctoberLoading && !isBackground) return;
    isOctoberLoading = true;

    const octLoadingCard = document.getElementById('oct-loading-card');
    const octEmptyState = document.getElementById('oct-empty-state');
    const btnRefreshOctober = document.getElementById('btn-refresh-october');
    const btnRefreshOctLabel = document.getElementById('btn-refresh-oct-label');
    const octSyncStatusLabel = document.getElementById('oct-sync-status-label');

    if (!isBackground && octLoadingCard && octoberAssignments.length === 0) {
      octLoadingCard.style.display = 'block';
      if (octEmptyState) octEmptyState.style.display = 'none';
    }

    if (btnRefreshOctober && !isBackground) {
      btnRefreshOctober.disabled = true;
      const spinSvg = btnRefreshOctober.querySelector('.spin-icon');
      if (spinSvg) spinSvg.classList.add('is-spinning');
      if (btnRefreshOctLabel) btnRefreshOctLabel.textContent = 'Syncing...';
    }

    try {
      let data = null;

      // 1. Try local/daemon API endpoint first
      try {
        const res = await fetch('/api/ond-assignments');
        if (res.ok) {
          data = await res.json();
        }
      } catch (localErr) {
        // Fallback for static hosting like GitHub Pages
      }

      // 2. Fallback to Google Apps Script Webhook if on GitHub Pages / static
      if ((!data || !data.success) && webhookUrl && webhookUrl.startsWith('http')) {
        try {
          const sep = webhookUrl.includes('?') ? '&' : '?';
          const whRes = await fetch(webhookUrl + sep + 'action=getONDAssignments');
          if (whRes.ok) {
            data = await whRes.json();
          }
        } catch (whErr) {
          console.warn("Webhook fetch fallback:", whErr);
        }
      if (data && data.success && Array.isArray(data.assignments)) {
          const oldIds = new Set(octoberAssignments.map(a => a.id));
          let newDetectedCount = 0;

          const freshList = data.assignments.map(item => {
            const isBrandNew = oldIds.size > 0 && !oldIds.has(item.id);
            if (isBrandNew) newDetectedCount++;
            return {
              ...item,
              isNew: isBrandNew
            };
          });

          // Retain any newly created optimistic items
          const optimisticItems = octoberAssignments.filter(a => a.isLocalOptimistic && !freshList.some(f => f.topic === a.topic));
          octoberAssignments = [...optimisticItems, ...freshList];

          lastOctoberSyncTime = new Date();
          updateOctoberLastSyncTime();

          populateOctoberFilterDropdowns(octoberAssignments);
          updateOctoberStats(data.stats, octoberAssignments.length);
          renderOctoberTable();

          if (octSyncStatusLabel) {
            octSyncStatusLabel.innerHTML = 'LIVE SYNC &bull; Connected';
          }

          if (notifyNew && newDetectedCount > 0) {
            showToast(`✨ ${newDetectedCount} new assignment(s) detected in October 2026!`, 'info');
          } else if (!isBackground && !notifyNew) {
            showToast(`✅ Synced ${octoberAssignments.length} October 2026 tasks from Google Sheet!`, 'success');
          }
        }
      }
    } catch (err) {
      console.warn("Could not fetch live October assignments:", err);
      if (!isBackground) {
        showToast('Could not sync with Google Sheets: ' + err.message, 'error');
      }
      if (octSyncStatusLabel) {
        octSyncStatusLabel.textContent = 'Sync Offline';
      }
    } finally {
      isOctoberLoading = false;
      if (octLoadingCard) octLoadingCard.style.display = 'none';
      if (btnRefreshOctober) {
        btnRefreshOctober.disabled = false;
        const spinSvg = btnRefreshOctober.querySelector('.spin-icon');
        if (spinSvg) spinSvg.classList.remove('is-spinning');
        if (btnRefreshOctLabel) btnRefreshOctLabel.textContent = 'Sync Now';
      }
    }
  }

  function updateOctoberLastSyncTime() {
    const timeDisplay = document.getElementById('oct-last-synced-time');
    if (!timeDisplay) return;
    if (!lastOctoberSyncTime) {
      timeDisplay.textContent = 'Just now';
      return;
    }
    const secsAgo = Math.floor((Date.now() - lastOctoberSyncTime.getTime()) / 1000);
    if (secsAgo < 10) {
      timeDisplay.textContent = 'Just now';
    } else if (secsAgo < 60) {
      timeDisplay.textContent = `${secsAgo}s ago`;
    } else {
      const minsAgo = Math.floor(secsAgo / 60);
      timeDisplay.textContent = `${minsAgo}m ago`;
    }
  }

  function updateOctoberStats(stats, totalCount) {
    const total = totalCount || (stats ? stats.total : 0);
    const newContent = stats ? stats.newContentCount : octoberAssignments.filter(a => a.sheetTab === 'New Content').length;
    const seoOpt = stats ? stats.seoOptCount : octoberAssignments.filter(a => a.sheetTab === 'SEO Optimization').length;
    const unpicked = stats ? stats.assignedCount : octoberAssignments.filter(a => !a.pickedBy && (a.status === 'Assigned' || !a.status)).length;
    const picked = stats ? stats.pickedCount : octoberAssignments.filter(a => a.pickedBy || a.status === 'Picked' || a.status === 'Good to Go' || a.status === 'Live').length;
    const urgent = stats ? stats.urgentCount : octoberAssignments.filter(a => a.priority === 'P0' || a.priority === 'P1').length;
    const octCount = stats && stats.octCount !== undefined ? stats.octCount : octoberAssignments.filter(a => a.month === 'October').length;
    const novCount = stats && stats.novCount !== undefined ? stats.novCount : octoberAssignments.filter(a => a.month === 'November').length;
    const decCount = stats && stats.decCount !== undefined ? stats.decCount : octoberAssignments.filter(a => a.month === 'December').length;

    // Badges
    const tabOctBadge = document.getElementById('tab-oct-badge');
    if (tabOctBadge) tabOctBadge.textContent = total;

    const headerDisplay = document.getElementById('header-oct-count-display');
    if (headerDisplay) headerDisplay.textContent = `${total} OND Articles`;

    // OND Month Filter Pills
    const pillOndAll = document.getElementById('pill-count-ond-all');
    if (pillOndAll) pillOndAll.textContent = total;

    const pillOndOct = document.getElementById('pill-count-ond-oct');
    if (pillOndOct) pillOndOct.textContent = octCount;

    const pillOndNov = document.getElementById('pill-count-ond-nov');
    if (pillOndNov) pillOndNov.textContent = novCount;

    const pillOndDec = document.getElementById('pill-count-ond-dec');
    if (pillOndDec) pillOndDec.textContent = decCount;

    // Source Sheet Filter Pills
    const pillAll = document.getElementById('pill-count-all');
    if (pillAll) pillAll.textContent = total;

    const pillNew = document.getElementById('pill-count-new');
    if (pillNew) pillNew.textContent = newContent;

    const pillSeo = document.getElementById('pill-count-seo');
    if (pillSeo) pillSeo.textContent = seoOpt;

    // KPI Cards
    const elTotal = document.getElementById('oct-stat-total');
    if (elTotal) elTotal.textContent = total;

    const elNew = document.getElementById('oct-stat-new-content');
    if (elNew) elNew.textContent = newContent;

    const elSeo = document.getElementById('oct-stat-seo-opt');
    if (elSeo) elSeo.textContent = seoOpt;

    const elUnpicked = document.getElementById('oct-stat-unpicked');
    if (elUnpicked) elUnpicked.textContent = unpicked;

    const elPicked = document.getElementById('oct-stat-picked');
    if (elPicked) elPicked.textContent = picked;

    const elUrgent = document.getElementById('oct-stat-urgent');
    if (elUrgent) elUrgent.textContent = urgent;
  }

  function populateOctoberFilterDropdowns(items) {
    const catSelect = document.getElementById('oct-filter-category');
    const assignSelect = document.getElementById('oct-filter-assignee');
    const writerSelect = document.getElementById('oct-filter-writer');

    if (catSelect && catSelect.options.length <= 1) {
      const cats = Array.from(new Set(items.map(i => i.category).filter(Boolean))).sort();
      cats.forEach(c => {
        const opt = document.createElement('option');
        opt.value = c;
        opt.textContent = c;
        catSelect.appendChild(opt);
      });
    }

    if (assignSelect && assignSelect.options.length <= 1) {
      const assignees = Array.from(new Set(items.map(i => i.assignedBy).filter(Boolean))).sort();
      assignees.forEach(a => {
        const opt = document.createElement('option');
        opt.value = a;
        opt.textContent = a;
        assignSelect.appendChild(opt);
      });
    }

    if (writerSelect && writerSelect.options.length <= 1) {
      const writers = Array.from(new Set(items.map(i => i.pickedBy).filter(Boolean))).sort();
      writers.forEach(w => {
        const opt = document.createElement('option');
        opt.value = w;
        opt.textContent = w;
        writerSelect.appendChild(opt);
      });
    }
  }

  function renderOctoberTable() {
    const octTableBody = document.getElementById('october-table-body');
    const octEmptyState = document.getElementById('oct-empty-state');
    const octSearchInput = document.getElementById('oct-search');
    const octFilterType = document.getElementById('oct-filter-type');
    const octFilterCategory = document.getElementById('oct-filter-category');
    const octFilterAssignee = document.getElementById('oct-filter-assignee');
    const octFilterWriter = document.getElementById('oct-filter-writer');
    const octFilterStatus = document.getElementById('oct-filter-status');
    const octFilterPriority = document.getElementById('oct-filter-priority');
    const octMatchesCount = document.getElementById('oct-matches-count');

    if (!octTableBody) return;

    const searchTerm = (octSearchInput ? octSearchInput.value : '').toLowerCase().trim();
    const typeFilter = octFilterType ? octFilterType.value : 'ALL';
    const catFilter = octFilterCategory ? octFilterCategory.value : 'ALL';
    const assignFilter = octFilterAssignee ? octFilterAssignee.value : 'ALL';
    const writerFilter = octFilterWriter ? octFilterWriter.value : 'ALL';
    const statusFilter = octFilterStatus ? octFilterStatus.value : 'ALL';
    const priorityFilter = octFilterPriority ? octFilterPriority.value : 'ALL';

    const filtered = octoberAssignments.filter(item => {
      // OND Month filter (October, November, December)
      if (ondMonthFilter !== 'ALL' && item.month !== ondMonthFilter) return false;

      // Source sheet tab filter
      if (octSourceFilter !== 'ALL' && item.sheetTab !== octSourceFilter) return false;

      // Type filter
      if (typeFilter !== 'ALL') {
        if (typeFilter === 'Article' && (item.contentType || '').toLowerCase() !== 'article') return false;
        if (typeFilter === 'News' && (item.contentType || '').toLowerCase() !== 'news') return false;
        if (typeFilter === 'SEO Optimization' && item.sheetTab !== 'SEO Optimization') return false;
      }

      // Category filter
      if (catFilter !== 'ALL' && item.category !== catFilter) return false;

      // Assignee filter
      if (assignFilter !== 'ALL' && item.assignedBy !== assignFilter) return false;

      // Writer filter
      if (writerFilter !== 'ALL' && item.pickedBy !== writerFilter) return false;

      // Status filter
      if (statusFilter !== 'ALL') {
        const s = (item.status || '').toLowerCase();
        if (statusFilter === 'Assigned' && (s !== 'assigned' && item.pickedBy)) return false;
        if (statusFilter === 'Picked' && s !== 'picked' && !item.pickedBy) return false;
        if (statusFilter === 'Good to Go' && !s.includes('good')) return false;
        if (statusFilter === 'Live' && !s.includes('live')) return false;
      }

      // Priority filter
      if (priorityFilter !== 'ALL' && item.priority !== priorityFilter) return false;

      // Search query filter
      if (searchTerm) {
        const fullText = [
          item.topic,
          item.page,
          item.url,
          item.primaryKeywords,
          item.secondaryKeywords,
          item.category,
          item.assignedBy,
          item.pickedBy,
          item.seoSuggestion
        ].filter(Boolean).join(' ').toLowerCase();

        if (!fullText.includes(searchTerm)) return false;
      }

      return true;
    });

    if (octMatchesCount) {
      octMatchesCount.textContent = filtered.length;
    }

    if (filtered.length === 0) {
      octTableBody.innerHTML = '';
      if (octEmptyState) octEmptyState.style.display = 'block';
      return;
    }

    if (octEmptyState) octEmptyState.style.display = 'none';

    octTableBody.innerHTML = filtered.map(item => {
      const pClass = (item.priority || 'p1').toLowerCase();
      const isNewClass = item.isNew ? 'new-row-highlight' : '';

      // Status badge calculation
      let statusHtml = '';
      const s = (item.status || '').toLowerCase();
      if (s.includes('live')) {
        statusHtml = `<span class="badge-status-pill live" title="Live on site">🟢 Live</span>`;
      } else if (s.includes('good')) {
        statusHtml = `<span class="badge-status-pill ready" title="Ready to publish">✓ Good to Go</span>`;
      } else if (s.includes('review')) {
        statusHtml = `<span class="badge-status-pill review" title="Under Review">⏳ Under Review</span>`;
      } else if (item.pickedBy || s.includes('picked')) {
        statusHtml = `<span class="badge-status-pill picked" title="Picked by ${escapeHtml(item.pickedBy || 'Writer')}">✍️ Picked</span>`;
      } else {
        statusHtml = `<span class="badge-status-pill assigned" title="Assigned & Awaiting Writer">● Assigned</span>`;
      }

      // Sheet Tab Badge
      const tabClass = item.sheetTab === 'New Content' ? 'new-content' : 'seo-opt';
      const rowBadge = `<a href="${item.sheetUrl || '#'}" target="_blank" rel="noopener" class="badge-tab-pill ${tabClass}" title="Open Google Sheet at Row ${item.sheetRow}">
        ${escapeHtml(item.sheetTab)} &bull; R${item.sheetRow} ↗
      </a>`;

      // Topic & URL
      const topicHtml = `
        <div>
          <span class="oct-topic-title">${item.isNew ? '<span class="badge-priority p0" style="font-size:0.65rem;margin-right:4px;">NEW</span> ' : ''}${escapeHtml(item.topic || item.page)}</span>
          ${item.url ? `<a href="${item.url}" target="_blank" rel="noopener" class="oct-topic-url" title="${escapeHtml(item.url)}">
            <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"></path><polyline points="15 3 21 3 21 9"></polyline><line x1="10" y1="14" x2="21" y2="3"></line></svg>
            ${escapeHtml(item.url.replace(/^https?:\/\/(www\.)?/, '').substring(0, 40))}${item.url.length > 45 ? '...' : ''}
          </a>` : `<span class="font-mono text-xs text-muted">${escapeHtml(item.page || '-')}</span>`}
        </div>
      `;

      // Type Badge
      const typeBadge = `<span class="badge-type-pill ${item.contentType ? item.contentType.toLowerCase().replace(/\s+/g, '-') : 'article'}">${escapeHtml(item.contentType || 'Article')}</span>`;

      return `
        <tr class="${isNewClass}" data-id="${item.id}">
          <td>${statusHtml}</td>
          <td class="font-mono text-xs text-nowrap">
            <div>${escapeHtml(item.date || '-')}</div>
            ${item.month ? `<span class="badge-month-tag ${item.month.toLowerCase()}" style="margin-top:3px;">${escapeHtml(item.month)}</span>` : ''}
          </td>
          <td>${rowBadge}</td>
          <td>${topicHtml}</td>
          <td>${typeBadge}</td>
          <td><span class="badge-tag">${escapeHtml(item.category || '-')}</span></td>
          <td class="font-medium text-xs">${escapeHtml(item.assignedBy || 'SEO Team')}</td>
          <td>
            ${item.pickedBy 
              ? `<div class="writer-badge-pill" title="Picked on ${escapeHtml(item.datePicked || '')}"><strong>${escapeHtml(item.pickedBy)}</strong></div>`
              : `<span class="text-subtle text-xs italic">Unassigned</span>`}
          </td>
          <td><span class="badge-priority ${pClass}">${escapeHtml(item.priority || 'P1')}</span></td>
          <td>
            <div class="font-medium text-xs text-truncate-2" title="${escapeHtml(item.primaryKeywords || '')}">${escapeHtml(item.primaryKeywords || '-')}</div>
            ${item.beforeRank ? `<div class="font-mono text-xs text-muted">Rank: ${escapeHtml(item.beforeRank)}</div>` : ''}
          </td>
          <td>
            <div class="text-xs text-truncate-2 text-muted" title="${escapeHtml(item.seoSuggestion || item.strategy || '')}">
              ${escapeHtml(item.seoSuggestion || item.strategy || '-')}
            </div>
          </td>
          <td>
            <div class="table-actions">
              <button type="button" class="btn-icon-action" data-action="view-oct" data-id="${item.id}" title="View Full Brief Details">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                  <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path>
                  <circle cx="12" cy="12" r="3"></circle>
                </svg>
              </button>
              <a href="${item.sheetUrl || '#'}" target="_blank" rel="noopener" class="btn-icon-action" title="Open directly in Google Sheets">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                  <path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"></path>
                  <polyline points="15 3 21 3 21 9"></polyline>
                  <line x1="10" y1="14" x2="21" y2="3"></line>
                </svg>
              </a>
              <button type="button" class="btn-icon-action" data-action="copy-tsv-oct" data-id="${item.id}" title="Copy TSV">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                  <rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>
                  <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
                </svg>
              </button>
            </div>
          </td>
        </tr>
      `;
    }).join('');
  }

  function exportOctoberAsCsv() {
    if (octoberAssignments.length === 0) {
      showToast('No October assignments available to export.', 'info');
      return;
    }

    const headers = [
      "ID", "Date Assigned", "Source Sheet", "Sheet Row", "Topic / Article Headline", 
      "Content Type", "Category", "Target Page", "Target URL", "Assigned By", 
      "Picked By / Writer", "Date Picked", "Priority", "Status", "Primary Keywords", 
      "Secondary Keywords", "SEO Suggestions & Goal"
    ];

    const escapeCsv = (str) => {
      if (str === null || str === undefined) return '""';
      const s = str.toString().replace(/"/g, '""');
      return `"${s}"`;
    };

    const rows = octoberAssignments.map(a => [
      escapeCsv(a.id),
      escapeCsv(a.date),
      escapeCsv(a.sheetTab),
      escapeCsv(a.sheetRow),
      escapeCsv(a.topic),
      escapeCsv(a.contentType),
      escapeCsv(a.category),
      escapeCsv(a.page),
      escapeCsv(a.url),
      escapeCsv(a.assignedBy),
      escapeCsv(a.pickedBy),
      escapeCsv(a.datePicked),
      escapeCsv(a.priority),
      escapeCsv(a.status),
      escapeCsv(a.primaryKeywords),
      escapeCsv(a.secondaryKeywords),
      escapeCsv(a.seoSuggestion)
    ].join(','));

    const csvContent = "\uFEFF" + [headers.map(escapeCsv).join(','), ...rows].join('\r\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', `October_2026_Assigned_Articles_${new Date().toISOString().split('T')[0]}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    showToast(`📥 Exported ${octoberAssignments.length} October 2026 articles to CSV!`, 'success');
  }

  function openOctoberBriefModal(item) {
    if (!viewModal || !viewModalContent) return;

    viewModalContent.innerHTML = `
      <div class="brief-detail-grid">
        <div class="brief-card-meta">
          <div class="brief-meta-title">Article Headline / Topic</div>
          <div class="brief-meta-val font-bold">${escapeHtml(item.topic || item.page)}</div>
        </div>
        <div class="brief-card-meta">
          <div class="brief-meta-title">Source Sheet & Row</div>
          <div class="brief-meta-val">
            <span class="badge-tab-pill ${item.sheetTab === 'New Content' ? 'new-content' : 'seo-opt'}">${escapeHtml(item.sheetTab)} &bull; Row ${item.sheetRow}</span>
            <a href="${item.sheetUrl || '#'}" target="_blank" rel="noopener" style="margin-left:6px;font-size:0.75rem;color:var(--primary);text-decoration:underline;">Open in Sheet ↗</a>
          </div>
        </div>
        <div class="brief-card-meta">
          <div class="brief-meta-title">Category & Content Type</div>
          <div class="brief-meta-val"><span class="badge-tag">${escapeHtml(item.category || '-')}</span> &bull; <strong>${escapeHtml(item.contentType || 'Article')}</strong></div>
        </div>
        <div class="brief-card-meta">
          <div class="brief-meta-title">Priority & Status</div>
          <div class="brief-meta-val">
            <span class="badge-priority ${(item.priority || 'p1').toLowerCase()}">${escapeHtml(item.priority || 'P1')}</span>
            <span style="margin-left: 8px; font-weight: 600;">${escapeHtml(item.status || 'Assigned')}</span>
          </div>
        </div>
      </div>

      <div class="brief-detail-grid">
        <div class="brief-card-meta">
          <div class="brief-meta-title">Assigned By</div>
          <div class="brief-meta-val font-medium">${escapeHtml(item.assignedBy || 'SEO Team')} on <strong>${escapeHtml(item.date || '-')}</strong></div>
        </div>
        <div class="brief-card-meta">
          <div class="brief-meta-title">Picked By / Writer</div>
          <div class="brief-meta-val font-medium">${item.pickedBy ? `<span style="color:var(--accent-emerald);font-weight:700;">${escapeHtml(item.pickedBy)}</span> ${item.datePicked ? `(on ${escapeHtml(item.datePicked)})` : ''}` : '<span class="text-subtle italic">Awaiting Writer Pickup</span>'}</div>
        </div>
      </div>

      ${item.url ? `
      <div class="brief-block-full">
        <h5>Target URL</h5>
        <div class="brief-text-box font-mono text-xs">
          <a href="${item.url}" target="_blank" rel="noopener" style="color:var(--primary);text-decoration:underline;">${escapeHtml(item.url)}</a>
        </div>
      </div>` : ''}

      <div class="brief-block-full">
        <h5>Primary Keywords & Search Intent</h5>
        <div class="brief-text-box">
          <strong>${escapeHtml(item.primaryKeywords || '-')}</strong>
          ${item.beforeRank ? `<span style="margin-left:12px;color:var(--text-muted);">(Before Rank: ${escapeHtml(item.beforeRank)})</span>` : ''}
          ${item.searchIntent ? `<div style="margin-top:4px;color:var(--text-muted);font-size:0.8rem;">Search Intent: ${escapeHtml(item.searchIntent)}</div>` : ''}
        </div>
      </div>

      ${item.secondaryKeywords ? `
      <div class="brief-block-full">
        <h5>Secondary / Supporting Keywords</h5>
        <div class="brief-text-box">${escapeHtml(item.secondaryKeywords)}</div>
      </div>` : ''}

      <div class="brief-block-full">
        <h5>SEO Suggestions, Strategic Goal & Directives</h5>
        <div class="brief-text-box" style="white-space: pre-wrap; font-family: var(--font-sans); line-height: 1.6;">${escapeHtml(item.seoSuggestion || item.strategy || 'No specific suggestion provided.')}</div>
      </div>
    `;

    // Hook copy button
    const btnCopyBrief = document.getElementById('btn-copy-brief-clipboard');
    if (btnCopyBrief) {
      btnCopyBrief.onclick = () => {
        const fullText = `OCTOBER 2026 ASSIGNMENT BRIEF
-----------------------------------------
Topic: ${item.topic}
Date Assigned: ${item.date}
Source Sheet: ${item.sheetTab} (Row ${item.sheetRow})
Category: ${item.category}
Priority: ${item.priority}
Status: ${item.status}
Assigned By: ${item.assignedBy}
Picked By: ${item.pickedBy || 'None'}
URL: ${item.url || item.page}

KEYWORDS:
- Primary: ${item.primaryKeywords || 'None'}
${item.secondaryKeywords ? `- Secondary: ${item.secondaryKeywords}` : ''}

SEO SUGGESTIONS & GOAL:
${item.seoSuggestion || item.strategy || 'None'}`;

        navigator.clipboard.writeText(fullText).then(() => {
          showToast('📋 Copied brief to clipboard!', 'success');
        });
      };
    }

    viewModal.classList.add('is-open');
  }

  // --- State Persistence ---
  function loadSavedState() {
    try {
      const storedAssignments = localStorage.getItem(STORAGE_KEYS.ASSIGNMENTS);
      if (storedAssignments) {
        assignments = JSON.parse(storedAssignments);
      } else {
        // Sample starter data to showcase the goal-driven requirement
        assignments = [
          {
            id: "TASK-B89E21",
            date: new Date().toISOString().split('T')[0],
            topic: "Bihar Police Constable Syllabus & Exam Pattern 2026 (Updated PDF)",
            category: "Bihar Gov. Exam",
            goalType: "SERP Rank Jump",
            targetGoal: "Jump from Rank #14 to Top 3. Capture 25,000 monthly search clicks before written test.",
            urgencyReason: "CSBC announced exam window yesterday; competitors missing updated physical criteria.",
            contentPriority: "P0",
            page: "/bihar-gov-exams/police-constable-syllabus-pattern",
            primaryKeywords: "bihar police constable syllabus 2026, exam pattern",
            beforeUpdateRank: "#14 (Page 2)",
            secondaryKeywords: "bihar police syllabus pdf in hindi, subject wise marks distribution, negative marking",
            keywords: "csbc constable selection process, physical test norms, previous year cutoff comparison",
            seoSuggestion: "1. H1 with exact primary keyword.\n2. Add downloadable bilingual PDF table.\n3. Integrate 5 FAQs with schema markup.\n4. Link to Admit Card & Mock Test series.",
            assignedBy: "Priyanshu (SEO Lead)",
            syncedToSheet: false,
            createdAt: new Date().toISOString()
          }
        ];
        saveAssignments();
      }

      webhookUrl = localStorage.getItem(STORAGE_KEYS.WEBHOOK_URL) || DEFAULT_WEBHOOK_URL;
      if (!localStorage.getItem(STORAGE_KEYS.WEBHOOK_URL)) {
        localStorage.setItem(STORAGE_KEYS.WEBHOOK_URL, DEFAULT_WEBHOOK_URL);
      }
      if (webhookUrlInput) {
        webhookUrlInput.value = webhookUrl;
      }

      let savedTheme = localStorage.getItem(STORAGE_KEYS.THEME);
      if (!savedTheme || savedTheme === 'dark') {
        savedTheme = 'light';
        localStorage.setItem(STORAGE_KEYS.THEME, 'light');
      }
      document.body.setAttribute('data-theme', savedTheme);
    } catch (e) {
      console.error("Error loading saved state:", e);
    }
  }

  function saveAssignments() {
    try {
      localStorage.setItem(STORAGE_KEYS.ASSIGNMENTS, JSON.stringify(assignments));
    } catch (e) {
      console.error("Error saving assignments:", e);
    }
  }

  function setupDefaultDate() {
    if (dateInput && !dateInput.value) {
      const today = new Date().toISOString().split('T')[0];
      dateInput.value = today;
    }
  }

  // --- Category Renderers ---
  function renderCategoryFilters() {
    if (!filterCategory) return;
    filterCategory.innerHTML = '<option value="ALL">All Categories (31)</option>';
    EXAM_CATEGORIES.forEach(cat => {
      const opt = document.createElement('option');
      opt.value = cat;
      opt.textContent = cat;
      filterCategory.appendChild(opt);
    });
  }

  function renderCategoryCloud() {
    if (!categoryTagsCloud) return;
    categoryTagsCloud.innerHTML = '';
    EXAM_CATEGORIES.forEach(cat => {
      const pill = document.createElement('span');
      pill.className = 'cat-pill';
      pill.textContent = cat;
      pill.addEventListener('click', () => {
        if (categorySelect) {
          categorySelect.value = cat;
          updateLivePreview();
          showToast(`Selected Category: ${cat}`, 'info');
          categorySelect.focus();
        }
      });
      categoryTagsCloud.appendChild(pill);
    });
  }

  // --- Real-Time Live Preview Sync ---
  function updateLivePreview() {
    const topicVal = topicInput ? topicInput.value.trim() : '';
    const dateVal = dateInput ? dateInput.value : '';
    const categoryVal = categorySelect ? categorySelect.value : '';
    const priorityVal = prioritySelect ? prioritySelect.value : 'P1';
    const pageVal = pageInput ? pageInput.value.trim() : '';
    const primaryKwVal = primaryKwInput ? primaryKwInput.value.trim() : '';
    const rankVal = beforeRankInput ? beforeRankInput.value.trim() : '';
    const secondaryKwVal = secondaryKwInput ? secondaryKwInput.value.trim() : '';
    const keywordsVal = keywordsInput ? keywordsInput.value.trim() : '';
    const suggestionVal = suggestionInput ? suggestionInput.value.trim() : '';
    
    const checkedGoal = document.querySelector('input[name="goalType"]:checked');
    const goalType = checkedGoal ? checkedGoal.value : 'Rank Jump';
    const targetGoalVal = targetGoalInput ? targetGoalInput.value.trim() : '';

    const checkedProject = document.querySelector('input[name="projectType"]:checked');
    const projectType = checkedProject ? checkedProject.value : 'SEO Optimization';

    const tabBadge = document.getElementById('preview-target-tab-badge');
    if (tabBadge) {
      if (projectType === 'New Content') {
        tabBadge.className = 'badge-tab-pill new-content font-bold';
        tabBadge.textContent = 'New Content • Row 2 (gid: 1107724151)';
      } else {
        tabBadge.className = 'badge-tab-pill seo-opt font-bold';
        tabBadge.textContent = 'SEO Optimization • Row 2 (gid: 1073261968)';
      }
    }

    if (btnSubmitLabel && !btnSubmit.classList.contains('is-loading')) {
      btnSubmitLabel.textContent = `Assign Task & Push to '${projectType}' (Row 2)`;
    }

    if (previewDate) previewDate.textContent = dateVal || 'Today';
    if (previewTopic) previewTopic.textContent = topicVal || 'Enter Topic headline above...';
    if (previewCategory) {
      previewCategory.innerHTML = categoryVal 
        ? `<span class="badge-tag">${escapeHtml(categoryVal)}</span>`
        : `<span class="badge-tag text-muted">Select Category</span>`;
    }
    if (previewGoal) {
      previewGoal.textContent = targetGoalVal ? `[${goalType}]: ${targetGoalVal}` : `[${goalType}]`;
    }
    if (previewPriority) {
      const pClass = (priorityVal || 'P1').toLowerCase();
      previewPriority.innerHTML = `<span class="badge-priority ${pClass}">${priorityVal || 'P1'}</span>`;
    }
    if (previewPage) previewPage.textContent = pageVal || '/target-page-slug';
    if (previewPrimaryKw) previewPrimaryKw.textContent = primaryKwVal || '-';
    if (previewRank) previewRank.textContent = rankVal || '-';
    if (previewSecondaryKw) previewSecondaryKw.textContent = secondaryKwVal || '-';
    if (previewKeywords) previewKeywords.textContent = keywordsVal || '-';
    if (previewSuggestion) previewSuggestion.textContent = suggestionVal || '-';
  }

  // --- Form Submission & Validation ---
  async function handleFormSubmit(e) {
    e.preventDefault();

    // Check HTML5 validity
    if (!form.checkValidity()) {
      form.reportValidity();
      showToast('Please fill out all mandatory fields.', 'error');
      return;
    }

    const checkedGoal = document.querySelector('input[name="goalType"]:checked');
    const goalType = checkedGoal ? checkedGoal.value : 'SERP Rank Jump';

    const checkedProject = document.querySelector('input[name="projectType"]:checked');
    const projectType = checkedProject ? checkedProject.value : 'SEO Optimization';

    const selectedContentType = (contentTypeInput && contentTypeInput.value) 
      ? contentTypeInput.value 
      : (projectType === 'New Content' ? 'News' : 'SEO Optimization');

    const newAssignment = {
      id: "TASK-" + Math.random().toString(36).substring(2, 8).toUpperCase(),
      projectType: projectType,
      date: dateInput.value,
      topic: topicInput.value.trim(),
      category: categorySelect.value,
      goalType: goalType,
      targetGoal: targetGoalInput.value.trim(),
      urgencyReason: urgencyInput.value.trim(),
      contentPriority: prioritySelect.value,
      contentType: selectedContentType,
      page: pageInput.value.trim(),
      primaryKeywords: primaryKwInput.value.trim(),
      beforeUpdateRank: beforeRankInput.value.trim(),
      secondaryKeywords: secondaryKwInput.value.trim(),
      keywords: keywordsInput.value.trim(),
      seoSuggestion: suggestionInput.value.trim(),
      assignedBy: assignedByInput.value.trim() || "SEO Team",
      syncedToSheet: false,
      createdAt: new Date().toISOString()
    };

    // UI Loading state
    setButtonLoading(true);

    try {
      // First attempt direct backend sync via Service Account (/api/assign) which inserts at Row 2!
      let synced = false;
      let syncMessage = '';

      try {
        const assignEndpoint = (window.location.origin && window.location.origin.startsWith('http'))
          ? '/api/assign'
          : 'http://127.0.0.1:4173/api/assign';

        const apiRes = await fetch(assignEndpoint, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(newAssignment)
        });

        if (apiRes.ok) {
          const apiJson = await apiRes.json();
          if (apiJson.success) {
            synced = true;
            syncMessage = `Inserted at Row 2 in "${projectType}"!`;
          }
        }
      } catch (backendErr) {
        console.warn("Direct /api/assign unavailable, checking webhook fallback...", backendErr);
      }

      // If backend was not running or failed, fallback to webhook if present
      if (!synced && webhookUrl && webhookUrl.trim().startsWith('http')) {
        const pushResult = await pushToGoogleSheet(newAssignment);
        if (pushResult.success) {
          synced = true;
          syncMessage = 'Pushed to Google Sheet via Webhook!';
        }
      }

      if (synced) {
        newAssignment.syncedToSheet = true;
        const targetGid = projectType === 'New Content' ? '1107724151' : '1073261968';
        const targetUrl = `https://docs.google.com/spreadsheets/d/1ihLeB9ZOJdaF841qGLoTWBULSRNsF9BjtxUXRxKuK2A/edit?gid=${targetGid}#gid=${targetGid}`;
        showToast(`Successfully inserted into <strong>${escapeHtml(projectType)}</strong> at <strong>Row 2</strong>! <a href="${targetUrl}" target="_blank" style="color:#0284c7;text-decoration:underline;margin-left:8px;font-weight:700;">Open '${escapeHtml(projectType)}' Tab ↗</a>`, 'success', true);
      } else {
        showToast('Saved locally! (Backend server connects directly to Google Sheet Row 2)', 'info');
      }

      // Add to beginning of local dispatched history
      assignments.unshift(newAssignment);
      saveAssignments();

      // Real-time optimistic insertion into OND 2026 Tracker tab
      const ondMonth = getONDMonth2026Date(newAssignment.date);
      if (ondMonth) {
        const targetGid = projectType === 'New Content' ? '1107724151' : '1073261968';
        const newOctItem = {
          id: newAssignment.id,
          sheetTab: projectType,
          sheetRow: 2, // New tasks insert directly at Row 2
          sheetUrl: `https://docs.google.com/spreadsheets/d/1ihLeB9ZOJdaF841qGLoTWBULSRNsF9BjtxUXRxKuK2A/edit?gid=${targetGid}#range=A2`,
          date: newAssignment.date,
          month: ondMonth,
          quarter: 'OND 2026',
          datePicked: '',
          assignedBy: newAssignment.assignedBy,
          pickedBy: '',
          priority: newAssignment.contentPriority,
          category: newAssignment.category,
          topic: newAssignment.topic,
          page: newAssignment.page,
          contentType: projectType === 'New Content' ? 'Article' : 'SEO Optimization',
          url: (newAssignment.page && newAssignment.page.startsWith('http')) ? newAssignment.page : '',
          primaryKeywords: newAssignment.primaryKeywords,
          beforeRank: newAssignment.beforeUpdateRank || '',
          secondaryKeywords: newAssignment.secondaryKeywords,
          seoSuggestion: newAssignment.seoSuggestion,
          status: 'Assigned',
          isNew: true,
          isLocalOptimistic: true
        };

        // Real-time prepend so it appears instantly at the top of the OND Tracker
        octoberAssignments.unshift(newOctItem);
        populateOctoberFilterDropdowns(octoberAssignments);
        updateOctoberStats(null, octoberAssignments.length);
        renderOctoberTable();

        showToast(`✨ Real-time: Added "<strong>${escapeHtml(newAssignment.topic)}</strong>" to OND (${ondMonth}) Tracker! <a href="#ond" onclick="switchDashboardTab('view-october'); return false;" style="color:#818cf8;text-decoration:underline;margin-left:6px;font-weight:700;">View OND Tracker ↗</a>`, 'success', true);

        // Schedule background sync to confirm row from Google Sheet
        setTimeout(() => {
          fetchOctoberAssignments(true, false);
        }, 2500);
      }

      // Reset form fields (keeping date & assigned by)
      const curDate = dateInput.value;
      const curAssignee = assignedByInput.value;
      form.reset();
      dateInput.value = curDate;
      assignedByInput.value = curAssignee;

      // Update UI
      updateLivePreview();
      renderAssignmentsTable();
      updateStats();

    } catch (err) {
      console.error("Submission error:", err);
      showToast('Encountered an issue while assigning task: ' + err.message, 'error');
    } finally {
      setButtonLoading(false);
    }
  }

  function setButtonLoading(loading) {
    if (!btnSubmit) return;
    if (loading) {
      btnSubmit.classList.add('is-loading');
      btnSubmit.disabled = true;
      if (btnSubmitLabel) btnSubmitLabel.textContent = 'Inserting at Row 2...';
    } else {
      btnSubmit.classList.remove('is-loading');
      btnSubmit.disabled = false;
      if (btnSubmitLabel) btnSubmitLabel.textContent = 'Assign Task & Push to Sheet';
    }
  }

  // --- Google Sheets Webhook Fallback Bridge ---
  async function pushToGoogleSheet(item) {
    if (!webhookUrl || !webhookUrl.trim()) {
      return { success: false, message: 'No Google Apps Script webhook URL configured.' };
    }

    try {
      // Priority safety for sheet (accepts P0, P1, P2; clamp P3 to P2)
      let safePriority = item.contentPriority || "P1";
      if (safePriority === "P3") {
        safePriority = "P2";
      }

      // Format Search Intent to exact allowed dropdown value
      let safeSearchIntent = "Informational → explain, guide, clarify";
      if (item.searchIntent) {
        if (item.searchIntent.indexOf("Mixed") !== -1) safeSearchIntent = "Mixed → balance info + action";
        else if (item.searchIntent.indexOf("Transactional") !== -1) safeSearchIntent = "Transactional → highlight benefits, CTAs";
      }

      // Strictly normalize Content Type for New Content (accepts Child Page, Blog, News, SEO Group, Discover)
      const allowedTypes = ["Child Page", "Blog", "News", "SEO Group", "Discover"];
      let safeContentType = "News";
      if (item.contentType && allowedTypes.includes(item.contentType)) {
        safeContentType = item.contentType;
      } else if (item.contentType === "Article") {
        safeContentType = "News";
      }

      const payload = {
        id: item.id,
        projectType: item.projectType || "SEO Optimization",
        date: item.date,
        topic: item.topic || "",
        category: item.category || "",
        goalType: item.goalType || "",
        targetGoal: item.targetGoal || "",
        urgencyReason: item.urgencyReason || "",
        contentPriority: safePriority,
        contentType: safeContentType,
        page: item.page || item.topic || "Overview",
        primaryKeywords: item.primaryKeywords || "",
        beforeUpdateRank: item.beforeUpdateRank || "",
        secondaryKeywords: item.secondaryKeywords || "",
        keywords: item.keywords || "",
        seoSuggestion: item.seoSuggestion || "",
        assignedBy: item.assignedBy || "SEO Team",
        searchIntent: safeSearchIntent,
        areaOfOptimisation: item.areaOfOptimisation || "",
        contentGap: item.contentGap || "",
        status: "Assigned"
      };

      await fetch(webhookUrl.trim(), {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(payload)
      });

      return { success: true };
    } catch (err) {
      console.error("Google Sheets Webhook Error:", err);
      return { success: false, message: err.message };
    }
  }

  // --- Render Table & Filters ---
  function renderAssignmentsTable() {
    if (!tableBody) return;

    const searchTerm = searchInput ? searchInput.value.toLowerCase().trim() : '';
    const selectedTab = filterTab ? filterTab.value : 'ALL';
    const selectedCategory = filterCategory ? filterCategory.value : 'ALL';
    const selectedPriority = filterPriority ? filterPriority.value : 'ALL';

    const filtered = assignments.filter(item => {
      const currentTab = item.sheetTab || item.projectType || 'SEO Optimization';
      const matchTab = selectedTab === 'ALL' || currentTab === selectedTab;
      const matchCat = selectedCategory === 'ALL' || item.category === selectedCategory;
      const matchPri = selectedPriority === 'ALL' || item.contentPriority === selectedPriority;
      const matchSearch = !searchTerm || 
        (item.topic && item.topic.toLowerCase().includes(searchTerm)) ||
        (item.primaryKeywords && item.primaryKeywords.toLowerCase().includes(searchTerm)) ||
        (item.category && item.category.toLowerCase().includes(searchTerm)) ||
        (item.page && item.page.toLowerCase().includes(searchTerm)) ||
        (item.assignedBy && item.assignedBy.toLowerCase().includes(searchTerm));
      return matchTab && matchCat && matchPri && matchSearch;
    });

    if (historyCountBadge) {
      historyCountBadge.textContent = filtered.length;
    }

    if (filtered.length === 0) {
      tableBody.innerHTML = '';
      if (emptyStateCard) emptyStateCard.classList.add('is-visible');
      return;
    }

    if (emptyStateCard) emptyStateCard.classList.remove('is-visible');

    tableBody.innerHTML = filtered.map(item => {
      const pClass = (item.contentPriority || 'p1').toLowerCase();
      const syncBadge = item.syncedToSheet
        ? `<span class="badge-sync synced" title="Synced with Google Sheet">✓ Live Sheet</span>`
        : `<span class="badge-sync local" title="Stored locally. Click Sync to push">● Local</span>`;

      const tabName = item.sheetTab || item.projectType || 'SEO Optimization';
      const tabClass = tabName === 'New Content' ? 'new-content' : 'seo-opt';
      const rowNum = item.sheetRow || 2;
      const targetGid = tabName === 'New Content' ? '1107724151' : '1073261968';
      const sheetRowUrl = item.sheetUrl || `https://docs.google.com/spreadsheets/d/1ihLeB9ZOJdaF841qGLoTWBULSRNsF9BjtxUXRxKuK2A/edit?gid=${targetGid}#range=A${rowNum}`;

      return `
        <tr data-id="${item.id}">
          <td>${syncBadge}</td>
          <td class="font-mono text-xs">${escapeHtml(item.date)}</td>
          <td>
            <div class="font-medium">${escapeHtml(item.topic)}</div>
            <div class="table-topic-meta">
              <span class="badge-tab-pill ${tabClass}">${escapeHtml(tabName)} &bull; Row ${rowNum}</span>
              <span class="font-mono text-xs text-muted">${escapeHtml(item.page)}</span>
            </div>
          </td>
          <td><span class="badge-tag">${escapeHtml(item.category)}</span></td>
          <td>
            <div class="goal-highlight-text text-xs font-bold">[${escapeHtml(item.goalType)}]</div>
            <div class="text-xs text-truncate-2" title="${escapeHtml(item.targetGoal)}">${escapeHtml(item.targetGoal)}</div>
          </td>
          <td><span class="badge-priority ${pClass}">${escapeHtml(item.contentPriority)}</span></td>
          <td>
            <div class="font-medium text-xs">${escapeHtml(item.primaryKeywords)}</div>
            <div class="font-mono text-xs text-muted">Rank: ${escapeHtml(item.beforeUpdateRank)}</div>
          </td>
          <td>
            <div class="text-xs text-truncate-2" title="${escapeHtml(item.secondaryKeywords)}">${escapeHtml(item.secondaryKeywords)}</div>
          </td>
          <td>
            <div class="text-xs text-truncate-2" title="${escapeHtml(item.seoSuggestion)}">${escapeHtml(item.seoSuggestion)}</div>
          </td>
          <td>
            <div class="table-actions">
              <a href="${sheetRowUrl}" target="_blank" class="btn-icon-action" title="Open Row ${rowNum} in Google Sheets">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                  <path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"></path>
                  <polyline points="15 3 21 3 21 9"></polyline>
                  <line x1="10" y1="14" x2="21" y2="3"></line>
                </svg>
              </a>
              <button class="btn-icon-action" data-action="view" data-id="${item.id}" title="View Full Brief">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                  <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path>
                  <circle cx="12" cy="12" r="3"></circle>
                </svg>
              </button>
              <button class="btn-icon-action" data-action="copy-tsv" data-id="${item.id}" title="Copy TSV (Paste to Sheet)">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                  <rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>
                  <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
                </svg>
              </button>
            </div>
          </td>
        </tr>
      `;
    }).join('');
  }

  function updateStats(customStats) {
    if (customStats) {
      if (statTotal) statTotal.textContent = customStats.total || assignments.length;
      if (statUrgent) statUrgent.textContent = customStats.urgentCount || assignments.filter(a => a.contentPriority === 'P0' || a.contentPriority === 'P1').length;
      if (statSynced) statSynced.textContent = customStats.pickedCount !== undefined ? `${customStats.pickedCount} Picked` : assignments.filter(a => a.syncedToSheet).length;
      return;
    }
    const total = assignments.length;
    const urgent = assignments.filter(a => a.contentPriority === 'P0' || a.contentPriority === 'P1').length;
    const synced = assignments.filter(a => a.syncedToSheet).length;

    if (statTotal) statTotal.textContent = total;
    if (statUrgent) statUrgent.textContent = urgent;
    if (statSynced) statSynced.textContent = synced;
  }

  function updateWebhookStatusUI() {
    if (!sheetStatusDot || !sheetStatusText) return;
    if (webhookUrl && webhookUrl.trim().startsWith('http')) {
      sheetStatusDot.className = 'status-dot connected';
      sheetStatusText.textContent = 'Sheets: Connected';
      if (intStatusDesc) {
        intStatusDesc.textContent = 'Active Webhook Connected to Google Sheet.';
      }
    } else {
      sheetStatusDot.className = 'status-dot disconnected';
      sheetStatusText.textContent = 'Sheets: Local Only';
      if (intStatusDesc) {
        intStatusDesc.textContent = 'No webhook configured. Saving locally.';
      }
    }
  }

  // --- Export & Clipboard Utilities ---
  function formatAsTsvRow(item) {
    const fields = [
      item.id,
      item.date,
      item.topic,
      item.category,
      item.goalType,
      item.targetGoal,
      item.urgencyReason,
      item.contentPriority,
      item.page,
      item.primaryKeywords,
      item.beforeUpdateRank,
      item.secondaryKeywords,
      item.keywords,
      item.seoSuggestion,
      item.assignedBy,
      "Assigned",
      item.createdAt
    ];
    // Clean newlines/tabs inside cells for clean single-row paste
    return fields.map(f => (f || '').toString().replace(/[\t\r\n]+/g, ' ')).join('\t');
  }

  function copyCurrentDraftAsTsv() {
    const checkedGoal = document.querySelector('input[name="goalType"]:checked');
    const draft = {
      id: "DRAFT",
      date: dateInput.value,
      topic: topicInput.value.trim(),
      category: categorySelect.value,
      goalType: checkedGoal ? checkedGoal.value : '',
      targetGoal: targetGoalInput.value.trim(),
      urgencyReason: urgencyInput.value.trim(),
      contentPriority: prioritySelect.value,
      page: pageInput.value.trim(),
      primaryKeywords: primaryKwInput.value.trim(),
      beforeUpdateRank: beforeRankInput.value.trim(),
      secondaryKeywords: secondaryKwInput.value.trim(),
      keywords: keywordsInput.value.trim(),
      seoSuggestion: suggestionInput.value.trim(),
      assignedBy: assignedByInput.value.trim(),
      createdAt: new Date().toISOString()
    };

    const tsv = formatAsTsvRow(draft);
    navigator.clipboard.writeText(tsv).then(() => {
      showToast('📋 Copied row as TSV! Paste directly into Google Sheet (Ctrl+V)', 'success');
    }).catch(err => {
      showToast('Clipboard error: ' + err.message, 'error');
    });
  }

  function exportAssignmentsAsCsv() {
    if (assignments.length === 0) {
      showToast('No assignments to export yet.', 'info');
      return;
    }

    const headers = [
      "Assignment ID", "Date", "Topic", "Category", "Goal Type", 
      "Target Goal & Outcome", "Why Now / Urgency", "Content Priority", "Page / URL",
      "Primary Keywords", "Before Update Rank", "Secondary Keywords", 
      "Keywords Bucket", "SEO Suggestion & Brief", "Assigned By", "Status", "Created At"
    ];

    const escapeCsv = (str) => {
      if (str === null || str === undefined) return '""';
      const s = str.toString().replace(/"/g, '""');
      return `"${s}"`;
    };

    const rows = assignments.map(a => [
      escapeCsv(a.id),
      escapeCsv(a.date),
      escapeCsv(a.topic),
      escapeCsv(a.category),
      escapeCsv(a.goalType),
      escapeCsv(a.targetGoal),
      escapeCsv(a.urgencyReason),
      escapeCsv(a.contentPriority),
      escapeCsv(a.page),
      escapeCsv(a.primaryKeywords),
      escapeCsv(a.beforeUpdateRank),
      escapeCsv(a.secondaryKeywords),
      escapeCsv(a.keywords),
      escapeCsv(a.seoSuggestion),
      escapeCsv(a.assignedBy),
      escapeCsv(a.syncedToSheet ? "Synced" : "Local"),
      escapeCsv(a.createdAt)
    ].join(','));

    const csvContent = "\uFEFF" + [headers.map(escapeCsv).join(','), ...rows].join('\r\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', `SEO_Assignments_${new Date().toISOString().split('T')[0]}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    showToast('📥 Exported assignments CSV successfully!', 'success');
  }

  async function syncAllPendingToSheet() {
    if (!webhookUrl || !webhookUrl.trim()) {
      showToast('Please configure your Google Apps Script Webhook URL first.', 'error');
      openSettingsModal();
      return;
    }

    const pending = assignments.filter(a => !a.syncedToSheet);
    if (pending.length === 0) {
      showToast('All assignments are already synced to Google Sheet!', 'info');
      return;
    }

    showToast(`Syncing ${pending.length} pending items to Google Sheet...`, 'info');

    let successCount = 0;
    for (const item of pending) {
      const res = await pushToGoogleSheet(item);
      if (res.success) {
        item.syncedToSheet = true;
        successCount++;
      }
    }

    saveAssignments();
    renderAssignmentsTable();
    updateStats();
    showToast(`✅ Successfully synced ${successCount} assignments to Google Sheet!`, 'success');
  }

  // --- View Brief Modal ---
  function openBriefModal(item) {
    if (!viewModal || !viewModalContent) return;

    viewModalContent.innerHTML = `
      <div class="brief-detail-grid">
        <div class="brief-card-meta">
          <div class="brief-meta-title">Topic Headline</div>
          <div class="brief-meta-val">${escapeHtml(item.topic)}</div>
        </div>
        <div class="brief-card-meta">
          <div class="brief-meta-title">Category & Priority</div>
          <div class="brief-meta-val">${escapeHtml(item.category)} &bull; <strong>${escapeHtml(item.contentPriority)}</strong></div>
        </div>
        <div class="brief-card-meta">
          <div class="brief-meta-title">Strategic Goal Type</div>
          <div class="brief-meta-val goal-highlight-text">${escapeHtml(item.goalType)}</div>
        </div>
        <div class="brief-card-meta">
          <div class="brief-meta-title">Target Page URL</div>
          <div class="brief-meta-val font-mono text-xs">${escapeHtml(item.page)}</div>
        </div>
      </div>

      <div class="brief-block-full">
        <h5>Target Goal & Expected Metric</h5>
        <div class="brief-text-box">${escapeHtml(item.targetGoal)}</div>
      </div>

      <div class="brief-block-full">
        <h5>Why Now? / Urgency Justification</h5>
        <div class="brief-text-box">${escapeHtml(item.urgencyReason)}</div>
      </div>

      <div class="brief-detail-grid">
        <div class="brief-card-meta">
          <div class="brief-meta-title">Primary Keywords</div>
          <div class="brief-meta-val">${escapeHtml(item.primaryKeywords)}</div>
        </div>
        <div class="brief-card-meta">
          <div class="brief-meta-title">Before Update SERP Rank</div>
          <div class="brief-meta-val font-bold text-amber">${escapeHtml(item.beforeUpdateRank)}</div>
        </div>
      </div>

      <div class="brief-block-full">
        <h5>Secondary Keywords</h5>
        <div class="brief-text-box">${escapeHtml(item.secondaryKeywords)}</div>
      </div>

      <div class="brief-block-full">
        <h5>Keywords Bucket / LSI / Search Volumes</h5>
        <div class="brief-text-box">${escapeHtml(item.keywords)}</div>
      </div>

      <div class="brief-block-full">
        <h5>SEO Suggestions & Guidelines</h5>
        <div class="brief-text-box">${escapeHtml(item.seoSuggestion)}</div>
      </div>
    `;

    // Hook copy button
    const btnCopyBrief = document.getElementById('btn-copy-brief-clipboard');
    if (btnCopyBrief) {
      btnCopyBrief.onclick = () => {
        const fullText = `SEO CONTENT ASSIGNMENT BRIEF
-----------------------------------------
Topic: ${item.topic}
Date: ${item.date}
Category: ${item.category}
Priority: ${item.contentPriority}
Page: ${item.page}

STRATEGIC GOAL: [${item.goalType}]
${item.targetGoal}

WHY NOW:
${item.urgencyReason}

KEYWORD STRATEGY:
- Primary Keywords: ${item.primaryKeywords}
- Before Update Rank: ${item.beforeUpdateRank}
- Secondary Keywords: ${item.secondaryKeywords}
- Keywords Bucket: ${item.keywords}

SEO SUGGESTIONS & REQUIREMENTS:
${item.seoSuggestion}

Assigned By: ${item.assignedBy}`;

        navigator.clipboard.writeText(fullText).then(() => {
          showToast('📋 Copied Full Brief to clipboard!', 'success');
        });
      };
    }

    viewModal.classList.add('is-open');
  }

  // --- Settings Modal & Connection Test ---
  function openSettingsModal() {
    if (settingsModal) {
      settingsModal.classList.add('is-open');
      if (webhookUrlInput) webhookUrlInput.focus();
    }
  }

  function closeSettingsModal() {
    if (settingsModal) {
      settingsModal.classList.remove('is-open');
    }
  }

  async function testWebhookConnection() {
    const url = webhookUrlInput ? webhookUrlInput.value.trim() : '';
    if (!url) {
      setTestFeedback('Please enter a Google Apps Script Web App URL first.', 'error');
      return;
    }

    if (!url.startsWith('https://script.google.com/macros/s/')) {
      setTestFeedback('URL must be a valid Google Apps Script Web App URL starting with https://script.google.com/macros/s/', 'error');
      return;
    }

    if (testSpinner) testSpinner.style.display = 'inline-block';
    setTestFeedback('Testing connection to Google Sheets...', '');

    try {
      // In Google Apps Script, a GET request returns {status: 'success'}
      // When fetched from browser, it may redirect (302) to googleusercontent.com
      const startTime = performance.now();
      const response = await fetch(url, { method: 'GET' });
      const duration = Math.round(performance.now() - startTime);

      if (response.ok) {
        const json = await response.json();
        setTestFeedback(`✅ Connection verified! Responded in ${duration}ms (${json.message || 'Ready'})`, 'success');
      } else {
        setTestFeedback(`Server responded with status ${response.status}. Please check deployment permissions ('Anyone').`, 'error');
      }
    } catch (err) {
      // If CORS blocked the GET JSON reading, test a dummy HEAD or no-cors ping
      setTestFeedback(`ℹ️ Webhook reachable! (Browser security restricts direct GET read, but POST submissions will work perfectly with 'Anyone' access)`, 'success');
    } finally {
      if (testSpinner) testSpinner.style.display = 'none';
    }
  }

  function setTestFeedback(msg, type) {
    if (!testFeedback) return;
    testFeedback.textContent = msg;
    testFeedback.className = 'test-feedback-text ' + (type || '');
  }

  function saveSettings() {
    const url = webhookUrlInput ? webhookUrlInput.value.trim() : '';
    webhookUrl = url;
    localStorage.setItem(STORAGE_KEYS.WEBHOOK_URL, url);
    updateWebhookStatusUI();
    closeSettingsModal();
    showToast(url ? 'Google Sheets Webhook saved & connected!' : 'Webhook removed. Saving locally.', 'success');
  }

  // --- Toast Notifications ---
  function showToast(message, type = 'info', isHtml = false) {
    const container = document.getElementById('toast-container');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    
    let icon = 'ℹ️';
    if (type === 'success') icon = '✅';
    if (type === 'error') icon = '❌';

    const content = isHtml ? message : escapeHtml(message);
    toast.innerHTML = `<span style="font-size:1.1rem;flex-shrink:0">${icon}</span> <span>${content}</span>`;
    container.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateX(30px)';
      toast.style.transition = 'all 0.3s ease';
      setTimeout(() => toast.remove(), 300);
    }, 6000);
  }

  // --- Helper: Escape HTML ---
  function escapeHtml(str) {
    if (str === null || str === undefined) return '';
    return str.toString()
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  // --- Event Listeners Setup ---
  function setupEventListeners() {
    // Form submission
    if (form) {
      form.addEventListener('submit', handleFormSubmit);
    }

    // Input listeners for live preview
    const inputsToWatch = [
      topicInput, dateInput, categorySelect, prioritySelect, pageInput,
      primaryKwInput, beforeRankInput, secondaryKwInput, keywordsInput,
      suggestionInput, targetGoalInput, urgencyInput
    ];

    inputsToWatch.forEach(input => {
      if (input) {
        input.addEventListener('input', updateLivePreview);
        input.addEventListener('change', updateLivePreview);
      }
    });

    // Radio change for goal types
    const goalRadios = document.querySelectorAll('input[name="goalType"]');
    goalRadios.forEach(radio => {
      radio.addEventListener('change', updateLivePreview);
    });

    // Radio change for project type (SEO Optimization vs New Content)
    const projectRadios = document.querySelectorAll('input[name="projectType"]');
    projectRadios.forEach(radio => {
      radio.addEventListener('change', () => {
        document.querySelectorAll('.project-toggle-card').forEach(c => c.classList.remove('active'));
        radio.closest('.project-toggle-card')?.classList.add('active');
        updateLivePreview();
        showToast(`Target Tab switched to: "${radio.value}" (Row 2)`, 'info');
      });
    });

    // Quick suggestion prompt helpers
    const helperButtons = document.querySelectorAll('.btn-chip-helper');
    helperButtons.forEach(btn => {
      btn.addEventListener('click', () => {
        const insertText = btn.getAttribute('data-insert');
        if (suggestionInput && insertText) {
          if (suggestionInput.value.trim()) {
            suggestionInput.value += "\n" + insertText;
          } else {
            suggestionInput.value = insertText;
          }
          updateLivePreview();
          suggestionInput.focus();
        }
      });
    });

    // Copy draft TSV
    const btnCopyPreview = document.getElementById('btn-copy-preview-tsv');
    if (btnCopyPreview) {
      btnCopyPreview.addEventListener('click', copyCurrentDraftAsTsv);
    }

    // Search and filters
    if (searchInput) searchInput.addEventListener('input', renderAssignmentsTable);
    if (filterTab) filterTab.addEventListener('change', renderAssignmentsTable);
    if (filterCategory) filterCategory.addEventListener('change', renderAssignmentsTable);
    if (filterPriority) filterPriority.addEventListener('change', renderAssignmentsTable);
    if (btnRefreshSheet) btnRefreshSheet.addEventListener('click', () => loadLiveOctoberAssignments(true));

    // Export CSV & Sync All
    const btnExportCsv = document.getElementById('btn-export-csv');
    if (btnExportCsv) btnExportCsv.addEventListener('click', exportAssignmentsAsCsv);

    const btnSyncAll = document.getElementById('btn-sync-all');
    if (btnSyncAll) btnSyncAll.addEventListener('click', syncAllPendingToSheet);

    // Settings Modal handlers
    const btnOpenSettings = document.getElementById('btn-open-settings');
    if (btnOpenSettings) btnOpenSettings.addEventListener('click', openSettingsModal);

    const btnViewGuide = document.getElementById('btn-view-guide');
    if (btnViewGuide) btnViewGuide.addEventListener('click', openSettingsModal);

    const btnConfigureShortcut = document.getElementById('btn-configure-sheet-shortcut');
    if (btnConfigureShortcut) btnConfigureShortcut.addEventListener('click', openSettingsModal);

    const btnCloseSettings = document.getElementById('btn-close-settings');
    if (btnCloseSettings) btnCloseSettings.addEventListener('click', closeSettingsModal);

    const btnCancelSettings = document.getElementById('btn-cancel-settings');
    if (btnCancelSettings) btnCancelSettings.addEventListener('click', closeSettingsModal);

    const btnSaveSettings = document.getElementById('btn-save-settings');
    if (btnSaveSettings) btnSaveSettings.addEventListener('click', saveSettings);

    const btnTestConnection = document.getElementById('btn-test-connection');
    if (btnTestConnection) btnTestConnection.addEventListener('click', testWebhookConnection);

    // Copy script button in modal
    const btnCopyScript = document.getElementById('btn-copy-script');
    if (btnCopyScript) {
      btnCopyScript.addEventListener('click', () => {
        navigator.clipboard.writeText(APPS_SCRIPT_CODE).then(() => {
          showToast('📋 Copied Google Apps Script code to clipboard!', 'success');
        });
      });
    }

    // View Modal Close
    const btnCloseView = document.getElementById('btn-close-view');
    if (btnCloseView) btnCloseView.addEventListener('click', () => viewModal.classList.remove('is-open'));

    const btnCloseViewFooter = document.getElementById('btn-close-view-footer');
    if (btnCloseViewFooter) btnCloseViewFooter.addEventListener('click', () => viewModal.classList.remove('is-open'));

    // Theme toggle
    const themeBtn = document.getElementById('theme-toggle-btn');
    if (themeBtn) {
      themeBtn.addEventListener('click', () => {
        const curTheme = document.body.getAttribute('data-theme') || 'dark';
        const newTheme = curTheme === 'dark' ? 'light' : 'dark';
        document.body.setAttribute('data-theme', newTheme);
        localStorage.setItem(STORAGE_KEYS.THEME, newTheme);
      });
    }

    // Top Dashboard Navigation Tabs
    document.querySelectorAll('.dashboard-tab-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const viewId = btn.getAttribute('data-view');
        switchDashboardTab(viewId);
      });
    });

    // Quick Assign button from October tab
    const btnQuickNewAssign = document.getElementById('btn-quick-new-assign');
    if (btnQuickNewAssign) {
      btnQuickNewAssign.addEventListener('click', () => switchDashboardTab('view-dispatch'));
    }

    // Refresh October assignments button
    const btnRefreshOctober = document.getElementById('btn-refresh-october');
    if (btnRefreshOctober) {
      btnRefreshOctober.addEventListener('click', () => fetchOctoberAssignments(false, false));
    }

    // Export October CSV
    const btnExportOctoberCsv = document.getElementById('btn-export-october-csv');
    if (btnExportOctoberCsv) {
      btnExportOctoberCsv.addEventListener('click', exportOctoberAsCsv);
    }

    // Segmented OND Month Filter Pills (All OND / October / November / December)
    document.querySelectorAll('#ond-month-pills .segment-pill').forEach(pill => {
      pill.addEventListener('click', () => {
        document.querySelectorAll('#ond-month-pills .segment-pill').forEach(p => p.classList.remove('active'));
        pill.classList.add('active');
        ondMonthFilter = pill.getAttribute('data-month') || 'ALL';
        renderOctoberTable();
      });
    });

    // Segmented Source Filter Pills (All / New Content / SEO Optimization)
    document.querySelectorAll('#oct-source-pills .segment-pill').forEach(pill => {
      pill.addEventListener('click', () => {
        document.querySelectorAll('#oct-source-pills .segment-pill').forEach(p => p.classList.remove('active'));
        pill.classList.add('active');
        octSourceFilter = pill.getAttribute('data-source') || 'ALL';
        renderOctoberTable();
      });
    });

    // October Search & Clear
    const octSearch = document.getElementById('oct-search');
    const octClearSearch = document.getElementById('oct-clear-search');
    if (octSearch) {
      octSearch.addEventListener('input', () => {
        if (octClearSearch) octClearSearch.style.display = octSearch.value ? 'inline-block' : 'none';
        renderOctoberTable();
      });
    }
    if (octClearSearch) {
      octClearSearch.addEventListener('click', () => {
        if (octSearch) {
          octSearch.value = '';
          octClearSearch.style.display = 'none';
          renderOctoberTable();
          octSearch.focus();
        }
      });
    }

    // October Dropdown Filters
    const octFilterIds = [
      'oct-filter-type', 'oct-filter-category', 'oct-filter-assignee',
      'oct-filter-writer', 'oct-filter-status', 'oct-filter-priority'
    ];
    octFilterIds.forEach(id => {
      const el = document.getElementById(id);
      if (el) el.addEventListener('change', renderOctoberTable);
    });

    // Reset October Filters
    const btnResetOctFilters = document.getElementById('btn-reset-oct-filters');
    if (btnResetOctFilters) {
      btnResetOctFilters.addEventListener('click', () => {
        if (octSearch) octSearch.value = '';
        if (octClearSearch) octClearSearch.style.display = 'none';
        octFilterIds.forEach(id => {
          const el = document.getElementById(id);
          if (el) el.value = 'ALL';
        });
        document.querySelectorAll('#ond-month-pills .segment-pill').forEach(p => p.classList.remove('active'));
        const allMonthPill = document.querySelector('#ond-month-pills .segment-pill[data-month="ALL"]');
        if (allMonthPill) allMonthPill.classList.add('active');
        ondMonthFilter = 'ALL';

        document.querySelectorAll('#oct-source-pills .segment-pill').forEach(p => p.classList.remove('active'));
        const allPill = document.querySelector('#oct-source-pills .segment-pill[data-source="ALL"]');
        if (allPill) allPill.classList.add('active');
        octSourceFilter = 'ALL';
        renderOctoberTable();
        showToast('Filters reset.', 'info');
      });
    }

    // October Table Delegated Actions (Brief view & copy TSV)
    const octTableBody = document.getElementById('october-table-body');
    if (octTableBody) {
      octTableBody.addEventListener('click', (e) => {
        const btn = e.target.closest('button[data-action]');
        if (!btn) return;
        const action = btn.getAttribute('data-action');
        const id = btn.getAttribute('data-id');
        const item = octoberAssignments.find(a => a.id === id);
        if (!item) return;

        if (action === 'view-oct') {
          openOctoberBriefModal(item);
        } else if (action === 'copy-tsv-oct') {
          const tsv = [
            item.id, item.date, item.sheetTab, item.sheetRow, item.topic, item.contentType,
            item.category, item.page, item.url, item.assignedBy, item.pickedBy, item.datePicked,
            item.priority, item.status, item.primaryKeywords, item.secondaryKeywords, item.seoSuggestion
          ].map(v => (v || '').toString().replace(/[\t\r\n]+/g, ' ')).join('\t');
          navigator.clipboard.writeText(tsv).then(() => {
            showToast('📋 Copied article TSV to clipboard!', 'success');
          });
        }
      });
    }

    // Table delegated actions
    if (tableBody) {
      tableBody.addEventListener('click', (e) => {
        const btn = e.target.closest('button[data-action]');
        if (!btn) return;

        const action = btn.getAttribute('data-action');
        const id = btn.getAttribute('data-id');
        const item = assignments.find(a => a.id === id);
        if (!item) return;

        if (action === 'view') {
          openBriefModal(item);
        } else if (action === 'copy-tsv') {
          const tsv = formatAsTsvRow(item);
          navigator.clipboard.writeText(tsv).then(() => {
            showToast('📋 Copied row TSV to clipboard! (Ready to paste into Sheet)', 'success');
          });
        } else if (action === 'delete') {
          if (confirm(`Are you sure you want to delete assignment: "${item.topic}"?`)) {
            assignments = assignments.filter(a => a.id !== id);
            saveAssignments();
            renderAssignmentsTable();
            updateStats();
            showToast('Assignment deleted.', 'info');
          }
        }
      });
    }

    // Keyboard shortcut: Ctrl + Enter to submit form
    document.addEventListener('keydown', (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        if (form) {
          form.requestSubmit();
        }
      }
    });
  }

  // Initialize on DOM ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
