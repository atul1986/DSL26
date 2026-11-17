# DSL Lab Project Submission Portal

**Ajeenkya D Y Patil University · School of Engineering · Second Year B.Tech · Data Structures Laboratory (DSL)**

| Page | Who | Access |
|---|---|---|
| `index.html` | Students | Open: no login. Submit a project, track its status |
| `admin.html` | Faculty | Secure login: review, download, export, manage accounts |

Website: hosted free on **GitHub Pages**. Data: stored in **your own Google Sheet + Google Drive** through a small Google Apps Script backend (`backend/Code.gs`).

```
Student browser ──► GitHub Pages (index.html) ──► Apps Script API ──► Google Sheet (data)
Faculty browser ──► GitHub Pages (admin.html) ──►   (Code.gs)     ──► Google Drive (files, per division)
```

---

## What students fill in
- **Program** (drop-down) and **Division A–M** (drop-down): once per team
- **Project title** (required) + optional data structure, abstract, GitHub link
- **Team of up to 4**: Student 1 (leader) name, URN, contact **required**; Students 2–4 **optional**
- **Upload**: PDF / DOC / DOCX, max 10 MB

Extra features: step-by-step wizard drawn as a linked list, team drawn as an array (`team[0..3]`), live **Division race** leaderboard, deadline countdown, draft auto-save, duplicate-URN protection, a submission ticket with QR code and **queue position**, Track Status page, and dark mode.

## Email notifications
- Students give an email (team leader compulsory, others optional).
- **On submission** the whole team gets a confirmation email: Submission ID, details, track link.
- **On review** (status or remarks changed) the team gets an update email with the faculty remarks.
- Sent from the Google account that deployed the script; replies go to `REPLY_TO` in `CONFIG`.
- Free Gmail limit: about **100 emails/day** (Workspace: 1,500). The dashboard shows how many are left today; if the limit is reached the submission is still saved and the sheet's *Email Log* column records it.

## Faculty security
- Username + password, **salted & hashed** (SHA-256 × 1000), never stored in plain text
- **Lockout**: 5 wrong attempts → account locked for 15 min
- **Sessions** expire after 2 h on the server and 20 min of inactivity in the browser
- **Roles**: `admin` (everything) and `faculty` (review only)
- **Division scoping**: e.g. a faculty account for `A,B,C` sees only those divisions, enforced on the server
- **Contacts masked** (98XXXXX210) until "Reveal" is clicked, and every reveal is logged
- **Files stay private** in your Drive and are downloaded only through an authenticated request
- **Audit Log** sheet: logins, failed logins, reviews, downloads, exports, account changes
- No passwords or secrets in this repo. The admin password is generated randomly at setup.

---

## One-time setup (≈ 10 minutes)

### 1 · Backend (Google Apps Script)
1. Open <https://script.google.com> → **New project** → name it `DSL Lab Portal API`.
2. Delete the sample code and paste the full contents of [`backend/Code.gs`](backend/Code.gs). Save.
3. In the toolbar choose function **`setup`** → **Run** → allow permissions.
4. Open **Execution log**. **Copy the admin password shown there.** It is shown only once.
   (It also prints the links to the Google Sheet and Drive folder it created.)
5. **Deploy → New deployment →** ⚙ type **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
   
   Click **Deploy** and copy the **Web app URL** (ends with `/exec`).

### 2 · Connect the website
Edit [`assets/config.js`](assets/config.js) and paste the URL:
```js
API_URL: 'https://script.google.com/macros/s/XXXXXXXX/exec',
```
Commit. GitHub Pages updates within a minute.

### 3 · First login
Open `admin.html` → username **`admin`** + the password from step 1.4 → **My account → change password** → **Faculty accounts → add** your lab faculty with their divisions.

---

## Changing settings later
Deadline, program list, max file size, late-submission rule etc. are in the `CONFIG` block at the top of `backend/Code.gs`.
After editing: **Deploy → Manage deployments → ✏️ → Version: New version → Deploy** (the URL stays the same).

**Forgot the admin password?** In the Apps Script editor run `resetAdminPassword` and read the new one from the log.

**Logo**: the official logo is loaded from adypu.edu.in. To use your own copy, add it as `assets/logo.png`.
