# Student Movement Tracker

A complete, production-ready web application for tracking student movements outside the classroom using **Google Sheets** as the primary database and **Google Apps Script** as the backend API.

---

## 1. Project Structure

```text
student-movement-tracker/
├── index.html                        # A. Teacher Dashboard (Classroom Checkout/Return Console)
├── principal-dashboard.html          # B. Principal Dashboard (Schoolwide Live Monitor)
├── leading-teacher-dashboard.html    # C. Leading Teacher Dashboard (Assigned Class Supervision)
├── admin-dashboard.html              # D. Admin Dashboard (10-Tab Management, Security & Reports)
├── css/
│   └── styles.css                    # Shared responsive design system & print layout
├── js/
│   ├── config.js                     # Externalized Web App URL & threshold configuration
│   ├── utils.js                      # Safe DOM rendering, 1899-safe timestamp engine, CSV parser/exporter
│   ├── api.js                        # Shared Google Apps Script API client & role session manager
│   ├── teacher-dashboard.js          # Teacher Dashboard logic
│   ├── principal-dashboard.js        # Principal Dashboard logic
│   ├── leading-teacher-dashboard.js  # Leading Teacher Dashboard logic
│   └── admin-dashboard.js            # Admin Dashboard logic
├── google-apps-script/
│   ├── Code.gs                       # Complete Google Apps Script backend API
│   └── appsscript.json               # Apps Script manifest
├── scripts/
│   └── syntax-check.js               # Automated JS/GS syntax & HTML structure verifier
├── tests/
│   ├── gas-mock.js                   # In-memory Google Apps Script & Spreadsheet mock runtime
│   └── run-tests.js                  # Automated end-to-end backend & frontend test suite
├── config.json                       # Reference configuration file
├── .clasp.json.example               # Template for Google clasp CLI deployment
├── server.js                         # Zero-dependency Node.js local preview server
└── package.json                      # NPM scripts for start, test, and syntax check
```

---

## 2. Local Development & Testing

### Start Local Preview Server
```bash
npm start
```
Then open any of the four role interfaces:
- **Teacher Dashboard:** `http://localhost:3000/index.html`
- **Principal Dashboard:** `http://localhost:3000/principal-dashboard.html`
- **Leading Teacher Dashboard:** `http://localhost:3000/leading-teacher-dashboard.html`
- **Admin Dashboard:** `http://localhost:3000/admin-dashboard.html`

### Run Syntax Checks & Automated Tests
```bash
npm run check
npm test
```

---

## 3. Google Sheets & Google Apps Script Setup Instructions

### Step 1: Create or Select Your Google Spreadsheet
1. Open [Google Sheets](https://sheets.google.com) and create a new spreadsheet (or open your existing school movement spreadsheet).
2. Copy the **Spreadsheet ID** from its URL:
   `https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit`
3. You do **not** need to manually create tabs—`Code.gs` automatically and non-destructively initializes all 8 required sheets (`Students`, `Classes`, `Movement_Log`, `Settings`, `Teachers`, `Teaching_Assignments`, `Leading_Teachers`, `Leading_Teacher_Assignments`) and their column headers on the first request while preserving any existing rows.

### Step 2: Deploy `google-apps-script/Code.gs`
#### Option A: Using the Google Apps Script Editor
1. In your Google Spreadsheet, click **Extensions → Apps Script** (or create a standalone project at [script.google.com](https://script.google.com)).
2. Copy the contents of `google-apps-script/Code.gs` into the script editor's `Code.gs` file.
3. If using a standalone script, go to **Project Settings (Gear Icon) → Script Properties** and add:
   - Property: `SPREADSHEET_ID`
   - Value: `<Your Google Spreadsheet ID>`
4. Provision your initial role passcodes and Teacher App PIN once (stored as salted SHA-256 hashes; plaintext is never stored):
   - In the Apps Script editor, add a temporary runner function or execute `setupInitialCredentials` from the console:
     ```javascript
     function initializeMySchoolCredentials() {
       setupInitialCredentials({
         adminPasscode: 'REPLACE_WITH_STRONG_ADMIN_PASSCODE',
         principalPasscode: 'REPLACE_WITH_PRINCIPAL_PASSCODE',
         leadingTeacherPasscode: 'REPLACE_WITH_LEADING_TEACHER_PASSCODE',
         teacherPin: 'REPLACE_WITH_TEACHER_PIN'
       });
     }
     ```
   - Select `initializeMySchoolCredentials` in the toolbar, click **Run**, authorize the script, and then remove the temporary function.
5. Click **Deploy → New deployment** (or **Manage deployments → Edit** to update an existing deployment):
   - **Select type:** Web app
   - **Execute as:** Me
   - **Who has access:** Anyone (required for the browser frontend to make cross-origin requests; all sensitive actions are strictly protected by server-side HMAC-SHA256 role/teacher tokens).
6. Copy the generated **Web App URL** (`https://script.google.com/macros/s/.../exec`).

#### Option B: Using Google `clasp` CLI (Optional, Requires Explicit Approval)
1. Copy `.clasp.json.example` to `.clasp.json` and set your `scriptId`.
2. Authenticate with `npx clasp login`.
3. When ready and approved, push code with `npx clasp push`.

### Step 3: Central Endpoint Configuration
The application uses one shared central configuration in `js/config.js` (`PRODUCTION_WEB_APP_URL`):
- **Production Endpoint (Default):**
  `https://script.google.com/macros/s/AKfycbygZM_HFzRnom0e3AFk4pSyvHufSLyF0hCpPba9oSr9wkyGXQrT-8SsMTU3MoLI-vtB/exec`
- All four dashboards (Teacher, Principal, Leading Teacher, Admin) connect to this exact same backend automatically without requiring manual URL entry.
- **Local Development Override:**
  Administrators can toggle between the production Google Apps Script backend and the local mock test proxy (`http://localhost:3000/api`) directly from the Admin Dashboard, or by using URL parameter `?local=true`.
