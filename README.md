# QR Time Clock App

A mobile-friendly QR time clock that runs on **GitHub Pages** and stores data in **Google Sheets** through **Google Apps Script**.

## What it does
- Employees scan a posted QR code to open the app
- Employees scan their own QR badge or type their employee code
- Tracks:
  - Clock In
  - Start Lunch
  - End Lunch
  - Clock Out
- Manager login can review daily timesheets and sign each one
- Includes a printable poster QR and employee badge QR generator

## Files
- `index.html` — app UI
- `style.css` — styling
- `app.js` — front-end logic
- `google-apps-script/Code.gs` — backend logic for Google Sheets

## Setup

### 1) Create the Google Sheet backend
Create a new Google Sheet and open **Extensions > Apps Script**.
Paste the contents of `google-apps-script/Code.gs` into the script editor.
Set the script timezone to **America/New_York**.
Run `setupSheets()` one time.

That creates these tabs:
- `ApprovedEmployees`
- `TimeLogs`
- `ManagerUsers`
- `Timesheets`

### 2) Add employees and managers
In `ApprovedEmployees`, use:
- EmployeeCode
- EmployeeName
- Active

Example:
- EMP001 | John Doe | TRUE

In `ManagerUsers`, use:
- Username
- Password
- ManagerName
- Active

### 3) Deploy Apps Script
In Apps Script:
- Click **Deploy > New deployment**
- Type: **Web app**
- Execute as: **Me**
- Who has access: **Anyone**

Copy the web app URL.

### 4) Connect the front end
Open `app.js` and replace:
- `PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE`

with your actual Apps Script web app URL.

### 5) Upload to GitHub Pages
Upload `index.html`, `style.css`, and `app.js` to a GitHub repo.
Enable **GitHub Pages** in repo settings.

### 6) Print the wall QR poster
Open the app, go to **QR Poster**, and print it.
That QR opens the app on employee phones.

### 7) Generate employee QR badges
In the QR Poster section, type the employee code and employee name.
It generates a badge QR that contains the employee code.
Print it and give it to the employee.

## Notes
- Phone camera scanning works best over HTTPS, which GitHub Pages already uses.
- Manager login in this version uses username/password stored in the Google Sheet.
- For tighter security later, you can replace that with Firebase Auth or Google login.
