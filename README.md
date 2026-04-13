# QR TimeClock Pro (Firebase Version)

This version is built for **Firebase Hosting + Firebase Authentication + Cloud Firestore**.

## What it does

- Employee signs in on phone
- Employee scans personal QR badge or types employee code
- Employee can:
  - Clock In
  - Start Lunch
  - End Lunch
  - Clock Out
- Every punch saves instantly to Firestore
- Manager sees live punches automatically
- Manager can open weekly timesheets and sign each one
- Signed timesheets can be reopened by a manager if needed
- Admin page lets you manage the Firestore `users` profiles
- QR tools page creates:
  - Company QR poster
  - Employee badge QR

## Folder files

- `index.html` — main app UI
- `style.css` — styling
- `app.js` — Firebase logic, punches, timesheets, QR tools
- `firebase-config.js` — your Firebase config goes here
- `firestore.rules` — starter Firestore security rules
- `firebase.json` — Firebase Hosting config

## Firebase setup

### 1) Create a Firebase project
In Firebase Console:
- Create project
- Add a **Web App**
- Enable **Authentication > Email/Password**
- Enable **Cloud Firestore**
- Enable **Hosting**

### 2) Paste config into `firebase-config.js`
Replace the placeholder values with your real Firebase web config.

### 3) Create Auth users
Create your managers and employees in **Authentication**.

### 4) Create matching Firestore user profiles
For every person, create a doc in:

`users/{uid}`

Example:

```json
{
  "name": "Brandon Evanshine",
  "email": "manager@company.com",
  "employeeId": "EMP-1001",
  "role": "manager",
  "active": true
}
```

Roles supported:
- `employee`
- `manager`
- `admin`

## Firestore collections used

### `users`
Stores profile and role info.

### `punches`
One document per punch.

Example:

```json
{
  "uid": "firebase-auth-uid",
  "employeeId": "EMP-1001",
  "name": "Brandon Evanshine",
  "action": "clock_in",
  "dateKey": "2026-04-13",
  "weekKey": "2026-04-13",
  "timestampMs": 1776110400000
}
```

### `timesheets`
One document per employee per week.

Doc id format:

`EMP-1001_2026-04-13`

Example fields:

```json
{
  "employeeId": "EMP-1001",
  "name": "Brandon Evanshine",
  "weekKey": "2026-04-13",
  "weeklyHours": 38.5,
  "status": "open",
  "managerSignedBy": "Jane Smith"
}
```

## How weekly signoff works

- Every punch updates that employee’s weekly timesheet doc
- Manager opens **Weekly Signoff**
- Manager clicks **Sign** on each person’s row
- The app stores:
  - `status = signed`
  - `managerSignedBy`
  - `managerSignedAt`

## How QR works

### Company QR poster
Generates a QR that opens your deployed app URL.
Post this in the warehouse so employees can open the punch page fast.

### Employee badge QR
Generates a QR containing only the employee code.
When scanned on the punch page, it fills the employee code box.

## Deploy with Firebase Hosting

Install Firebase CLI:

```bash
npm install -g firebase-tools
```

Log in:

```bash
firebase login
```

From this app folder:

```bash
firebase init hosting
```

Use these choices:
- existing project
- public directory: `.`
- single-page app: `No`
- do not overwrite files

Deploy rules:

```bash
firebase deploy --only firestore:rules
```

Deploy hosting:

```bash
firebase deploy --only hosting
```

## Good next upgrades

- export payroll CSV
- overtime calculation
- employee acknowledgment signature
- edit/correction request flow
- shift schedules and late flags
- department or location tags
- photo capture on clock in/out
- geofencing
- Cloud Functions for stricter server-side payroll calculations

## Important note

This is a practical working starter app.
For payroll-grade compliance, you may eventually want:
- Cloud Functions for server-side validation
- audit trail logs
- stronger rules around edits after signoff
- payroll export approval flow
