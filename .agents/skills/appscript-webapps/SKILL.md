---
name: appscript-webapps
description: |
  Comprehensive instructions, code snippets, and patterns for building robust, modern, and performant Single Page Web Applications (SPAs) on top of Google Apps Script and Google Sheets.
  
  Trigger when:
  - Designing or editing Google Apps Script web apps (Code.gs).
  - Building HTML templates, client-side CSS (styles.html), or client-side JavaScript (script.html) for Google Apps Script.
  - Implementing data persistence, relational mapping, or transactional locking with Google Sheets databases.
  - Adding features like inactivity logouts, History API routing, dynamic charts, or search/filtering to an AppScript webapp.
---

# Google Apps Script Web Application Blueprint

This document defines the architectural patterns, security controls, performance practices, and UI blueprints for building highly performant, responsive Single Page Applications (SPAs) using Google Apps Script (GAS) and Google Sheets.

---

## 1. Core Architecture (The 4-File SPA Model)
To avoid the complexity of modern build chains while keeping a clean separation of concerns, utilize the **4-File SPA model**:
1. `Code.gs`: Server-side controller. Handles HTTP requests, spreadsheet operations, authentication, and routing context.
2. `Index.html`: Main HTML entry point. Includes Google Fonts, icons, CDN libraries, and inlines style/script blocks.
3. `styles.html`: Contiguous CSS block containing CSS custom properties (variables), responsive sidebar designs, and component styling.
4. `script.html`: Client-side JavaScript. Handles routing, data fetching, rendering, user state, and session timeouts.

### Code.gs Boilerplate
```javascript
/**
 * Serves the Web App UI.
 */
function doGet() {
  // 1. Initial migrations or sanity checks can run here
  
  // 2. Evaluate template
  var template = HtmlService.createTemplateFromFile('Index');
  var evaluated = template.evaluate();
  
  // 3. Configure meta properties
  return evaluated
    .setTitle('Apps Script Web App')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, shrink-to-fit=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // If embedding is required
}

/**
 * Inlines file contents into the template.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
```

### Index.html Boilerplate
```html
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <!-- Inline styles template -->
    <?!= include('styles'); ?>
    <!-- External typography and icons -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
  </head>
  <body>
    <!-- App Layout Containers -->
    <div id="login-screen" class="login-container">...</div>
    <div id="app-container" class="app-container" style="display: none;">...</div>
    
    <!-- Inline script template -->
    <?!= include('script'); ?>
  </body>
</html>
```

---

## 2. Server-Side Data Access & Performance Patterns
Google Sheets acts as the database. However, direct row-by-row reading is an expensive RPC operation. Follow these performance optimization patterns:

### Read Optimizations (In-Memory Processing)
> [!IMPORTANT]
> Never query cells iteratively with `.getValue()`. Always fetch the entire sheet range in one RPC call using `.getDataRange().getValues()`, and then filter/map the array in-memory.

```javascript
function getFilteredData(params) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DataSheet');
  if (!sheet) return [];
  
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var data = [];
  
  // Map rows to objects for easy JSON manipulation
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var item = {};
    headers.forEach(function(header, index) {
      item[header] = row[index];
    });
    item.rowIndex = i + 1; // Store 1-indexed row number for direct edits
    
    // Apply in-memory filtering
    if (params.filterKey && item.status !== params.filterKey) continue;
    
    data.push(item);
  }
  return data;
}
```

### Write Optimizations (Transactional Concurrency & Locking)
Google Apps Script runs concurrently for different users. To prevent write conflicts and database corruption, always acquire a **Script Lock** before writing to Google Sheets.

```javascript
function addRecord(data) {
  var lock = LockService.getScriptLock();
  try {
    // Wait up to 30 seconds to acquire lock
    lock.waitLock(30000);
  } catch (e) {
    throw new Error('Database is busy. Please try again in a few moments.');
  }
  
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DataSheet');
    var newRow = [
      Utilities.getUuid(), // Generate primary key
      data.name,
      data.phone,
      new Date()
    ];
    sheet.appendRow(newRow);
    return { success: true, message: 'Record added successfully' };
  } finally {
    // Release the lock so other threads can write
    lock.releaseLock();
  }
}
```

### Relational Schema Emulation in Google Sheets
Simulate relationships using UUID foreign keys:
* **Table `Clients`**: `clientId` (UUID, primary key), `name`, `phone`, `city`
* **Table `Leads`**: `leadId` (UUID, primary key), `clientId` (Foreign Key matching `Clients.clientId`), `status`, `notes`, `enquiryDate`
* Use in-memory hash maps to join datasets quickly without nested loops:
  ```javascript
  var clientsMap = {};
  clientsRows.forEach(row => clientsMap[row.clientId] = row);
  ```

---

## 3. Client-Side SPA Design Patterns
To create a high-quality SPA experience inside a Google Apps Script environment, implement standard client routing and session state controls.

### SPA Navigation with History API (Back/Forward Buttons Support)
SPAs inside Apps Script default to losing user state on page refresh or when clicking the browser's back button. Bind layout state directly to the History API:

```javascript
let appHistoryReady = false;

function showSection(sectionId, shouldPushState = true) {
  // Hide all sections and show the target one
  document.querySelectorAll('.content-section').forEach(sec => sec.classList.remove('active'));
  document.getElementById(sectionId).classList.add('active');
  
  // Highlight active sidebar navigation links
  document.querySelectorAll('.nav-links li').forEach(li => {
    li.classList.toggle('active', li.getAttribute('data-section') === sectionId);
  });
  
  if (shouldPushState && appHistoryReady) {
    history.pushState({ section: sectionId }, '', location.href);
  }
}

function initAppHistory() {
  history.replaceState({ section: 'dashboard-section' }, '', location.href);
  if (!appHistoryReady) {
    appHistoryReady = true;
    window.addEventListener('popstate', function(event) {
      if (event.state && event.state.section) {
        showSection(event.state.section, false);
      }
    });
  }
}
```

### User Session Inactivity Timer
Apps Script webapps run inside user-scoped sessions. Manage session timeouts gracefully to protect sensitive CRM or business data:

```javascript
const WARNING_TIMEOUT = 25 * 60 * 1000; // 25 minutes
const LOGOUT_TIMEOUT  = 30 * 60 * 1000; // 30 minutes
let warnTimer, logoutTimer;

function startInactivityTimer() {
  clearInactivityTimers();
  warnTimer = setTimeout(showTimeoutWarning, WARNING_TIMEOUT);
  logoutTimer = setTimeout(forceLogout, LOGOUT_TIMEOUT);
}

function resetInactivityTimer() {
  // Run on user actions (clicks, keypresses)
  startInactivityTimer();
  hideTimeoutWarning();
}

function showTimeoutWarning() {
  document.getElementById('session-warning-toast').style.display = 'flex';
}

function keepSessionAlive() {
  resetInactivityTimer();
  // Call a lightweight server ping to extend Apps Script session cookies
  google.script.run.ping(); 
}

function forceLogout() {
  // Clear local variables and toggle back to login screen
  currentUser = null;
  document.getElementById('app-container').style.display = 'none';
  document.getElementById('login-screen').style.display = 'flex';
}
```

---

## 4. UI/UX Design System Tokens (styles.html)
To achieve premium visuals, declare consistent design tokens using CSS variables inside `styles.html`:

```css
:root {
  --primary: #006A7F;
  --primary-dark: #004d5e;
  --secondary: #e6f3f6;
  --text-dark: #1d1d1f;
  --text-light: #86868b;
  --bg-light: #f5f5f7;
  --border: #d2d2d7;
  --success: #34c759;
  --danger: #ff3b30;
  --shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
  --border-radius: 12px;
}

body {
  font-family: 'Inter', -apple-system, sans-serif;
  color: var(--text-dark);
  background-color: var(--bg-light);
}

/* Micro-Animations for Buttons & Cards */
.btn {
  padding: 10px 20px;
  border-radius: var(--border-radius);
  border: 1px solid transparent;
  font-weight: 500;
  cursor: pointer;
  transition: transform 0.15s ease, background-color 0.2s ease;
}
.btn:active {
  transform: scale(0.97);
}

.card {
  background: white;
  border: 1px solid var(--border);
  border-radius: var(--border-radius);
  box-shadow: var(--shadow);
  transition: transform 0.2s ease, box-shadow 0.2s ease;
}
.card:hover {
  transform: translateY(-2px);
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.08);
}
```

---

## 5. Security & Validation Practices
1. **Prevent Cross-Site Scripting (XSS)**: Always escape dynamic text injected into client HTML.
   ```javascript
   function escapeHtml(str) {
     return String(str || '').replace(/[&<>"']/g, function(ch) {
       return ({
         '&': '&amp;',
         '<': '&lt;',
         '>': '&gt;',
         '"': '&quot;',
         "'": '&#39;'
       })[ch];
     });
   }
   ```
2. **Normalize Inputs**: Convert input phone numbers, names, and emails to a strict standard format client-side before sending them to `google.script.run` to prevent data pollution.
3. **Role Enforcement**: Never trust client-side admin roles. When performing write/delete operations in `Code.gs`, verify the calling user's admin credentials server-side.
