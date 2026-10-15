# RankAssign PRO - Goal-Driven SEO to Content Assignment Hub

An executive-level, goal-oriented assignment dashboard engineered specifically for **SEO Teams** to dispatch structured content tasks to the **Content Writing Team**, with automated real-time synchronization directly into **Google Sheets**.

---

## 🎯 Target Google Spreadsheet Configuration
- **Spreadsheet URL**: [CM JFM - JAS 2026](https://docs.google.com/spreadsheets/d/1ihLeB9ZOJdaF841qGLoTWBULSRNsF9BjtxUXRxKuK2A/edit)
- **Spreadsheet ID**: `1ihLeB9ZOJdaF841qGLoTWBULSRNsF9BjtxUXRxKuK2A`
- **Target Tabs**:
  1. 🚀 **SEO Optimization** (`gid: 1073261968`) - For updating and refreshing existing content
  2. ✍️ **New Content** (`gid: 1107724151`) - For fresh topics, new exams, and new articles
- **Row 2 Insertion**: Every newly dispatched assignment is automatically inserted right at **Row 2** (immediately below the header row), pushing older tasks down so the latest assignment is always at the top!

---

## 🔑 Security & Authentication (Zero-Secret Leakage Guarantee)
- **Security Policy**: Credentials files (`credentials.json`, `credentials.json.json`, `*.env`) and private keys are **strictly excluded** from version control via `.gitignore` and blocked from static HTTP serving in `server.js`.
- **Environment Variables Support**: The backend seamlessly reads `GOOGLE_SERVICE_ACCOUNT_KEY` or `GOOGLE_CREDENTIALS` from environment variables, or falls back to local `credentials.json` when developing on a workstation.
- **GitHub Pages Live Link**:
  👉 **https://testbook-website.github.io/seo-dashboard/**
  *(On static GitHub Pages, the application safely delegates row insertions to Google Apps Script Webhook without exposing any service account keys).*

---

## 🎯 The Core Problem Solved: Goal-Driven Content Tasks
> *"SEO Team assigns things to Content Team, but they assign sometimes randomly, too randomly just to assign, so every assignment needs a Goal."*

RankAssign PRO enforces accountability before any task can be dispatched:
1. **Strategic Goal Selection**: SEO assignees must categorize the strategic intent (`SERP Rank Jump (Top 3)`, `Traffic Decay Recovery`, `New Exam Notification Alert`, `Intent Gap / Syllabus Refresh`, `Featured Snippet Win`, or `High-Intent Conversion / Leads`).
2. **Explicit Target Outcome**: Assignees must quantify the expected result (e.g. *"Move from Rank #14 to Top 3 within 21 days; capture 25,000 monthly clicks"*).
3. **Urgency Justification ("Why Now?")**: Clarifies why the task is prioritized today to prevent arbitrary task dumping.

---

## 📋 Complete Field Specification

| Field Name | Type | Mandatory? | Target Column in Google Sheet |
| :--- | :--- | :---: | :--- |
| **Project Type** | Radio Toggle | ✅ Yes | Directs to tab: `SEO Optimization` or `New Content` |
| **Date** | Date Picker | ✅ Yes | Col A (`Date`, formatted e.g. `14th October 2026`) |
| **Assigned By** | Text Input | ✅ Yes | `Assigned By` column |
| **Topic** | Text Input | ✅ Yes | `Area of Optimisation` (SEO Opt) / `Page` (New Content) |
| **Category** | Dropdown (31) | ✅ Yes | Col `Category` |
| **Goal & Expected Outcome** | Select & Textarea | ✅ Yes | `Strategy` column in SEO Opt / `SEO Suggestion` in New Content |
| **Urgency Reason ("Why Now?")** | Text Input | ✅ Yes | `Comment` / `Strategy` |
| **Content Priority** | Select Dropdown | ✅ Yes | Col `Content Priority` (`P0`, `P1`, `P2` — Sheet dropdown valid) |
| **Content Type** | Select Dropdown | ✅ Yes | Col `Type` (`News`, `Blog`, `Child Page`, `SEO Group`, `Discover`) |
| **Page / Target URL** | Text Input | ✅ Yes | Col `Page` and Col `Url` |
| **Primary Keywords** | Text Input | ✅ Yes | Col `Primary Keywords` / `Keywords` |
| **Before Update Rank** | Text Input | ✅ Yes | Col `Before Update Rank` |
| **Secondary Keywords** | Textarea | ✅ Yes | Col `Seconday Keywords` |
| **Keywords Bucket** | Textarea | ✅ Yes | Col `Keywords` / `LSI Keywords` |
| **SEO Suggestion** | Textarea | ✅ Yes | Col `SEO Suggestion` |

---

## 🏷️ Supported Exam Categories (All 31 in Dropdown & Quick Selector)
1. Bihar Gov. Exam
2. CSIR NET
3. CUET and UG Entrance
4. Defence Exams
5. Engineering Recruitment Exams
6. GATE
7. Haryana Exams
8. SSC Exams
9. Railway Exams
10. Banking Exams
11. Insurance Exam
12. Regulatory Exams
13. Other Govt Exam
14. MBA Entance Exams
15. Judiciary Exams
16. Kerala Exams
17. MH Govt Exams
18. MP Govt Exams
19. Nursing Exams
20. Odisha Gov. Exam
21. Police Exams
22. State PSC
23. Teaching Exams
24. UGC NET
25. UP Gov. Exam
26. UPSC Civil Services
27. Other UPSC Exams
28. West Bengal Exams
29. Rajasthan Exams
30. Punjab Exams
31. K12

---

## 🚀 How to Run the Server

Start the local Node backend:
```bash
node server.js
```
The server will start on:
👉 **`http://127.0.0.1:4173/`**

### What happens when you click "Assign Task & Push to Sheet":
1. The dashboard calls `POST /api/assign` (or Google Apps Script webhook on static deployments).
2. The server signs a secure JWT using `credentials.json` (kept strictly local/hidden).
3. It calls Google Sheets API `insertDimension` to insert a fresh row at **Row 2** in your chosen tab (`SEO Optimization` or `New Content`).
4. It maps all 27/21 columns into Row 2 and updates the sheet in under 1 second.
5. The UI shows an instant confirmation with the row number and a link to open the sheet!
6. If the assignment date falls in **OND 2026 (October, November, December)**, it is automatically and optimistically prepended to the **OND 2026 Tracker** tab with live animations!

---

## 📅 OND 2026 Tracker (October, November, December)

RankAssign PRO includes a dedicated real-time tracker for **OND 2026 (Q4 Content Execution Pipeline)**:
- **Comprehensive Coverage**: Pulls all tasks assigned in October, November, and December 2026 across both `New Content` and `SEO Optimization` tabs.
- **Real-Time Polling & Optimistic Updates**: New assignments appear at the top in real time; background auto-polling (every 20s) keeps all teammate changes synchronized without page refresh.
- **OND Month Selector**: Single-click filtering by **All OND**, **October (163)**, **November**, or **December**.
- **Interactive Multi-Filter Bar**: Filter by Source Sheet, Content Type, Exam Category, Assigner, Writer, Status, and Priority.
- **One-Click Export**: Download filtered or full OND 2026 records as clean CSV.
- **Direct Row Navigation**: Click any row's badge to jump straight to that row in Google Sheets (`#range=A{row}`).
