# StuDy v2 — Setup Guide (do these in order)

The app works in two modes:
- **DEMO mode** (default): everything runs on the device (localStorage). Good for testing the full flow — registration shows the OTP in a popup instead of emailing it.
- **LIVE mode**: after Step 2 below, data lives in your Google Sheet, OTPs go to real Gmail inboxes, and users can log in from any device.

## Step 1 — Create the GitHub repo (frontend)
1. Create a new GitHub repo (e.g. `StuDy`). Upload everything from this package **except** the `backend/` folder and this file:
   `index.html`, `questions.js`, `manifest.webmanifest`, `sw.js`, `assets/` (whole folder).
2. Repo → Settings → Pages → Source: `main` branch, root. Your app URL will be
   `https://<your-username>.github.io/<repo>/`.
3. Tell Muse the repo URL (it gets saved to `REPO.txt`) — then the Current Affairs auto-refresh (every 3 days) can push new questions to the repo automatically.

## Step 2 — Backend (Google Sheets + Apps Script)
1. Create a new Google Sheet named **StuDy DB**.
2. In the Sheet: **Extensions → Apps Script**. Delete the sample code, paste all of `backend/Code.gs`.
3. In the editor, select the function **setup** and click **Run** once (approve permissions). This creates the tabs: Users, Attempts, Payments, Settings, Notifications, OTPs.
4. **Deploy → New deployment → Web app**: Execute as **Me**, Who has access: **Anyone**. Copy the Web App URL.
5. Open `index.html`, find `const BACKEND_URL = "";` at the top of the script and paste your URL between the quotes. Commit/push.
6. MailApp (Gmail OTP) sends from YOUR Google account — free quota is ~100 mails/day, plenty to start.

## Step 3 — Admin account
1. Open the app and **Register the admin number (9588293680)** normally with your chosen password — that account automatically gets the Admin Panel (Profile page → Admin Panel button, plus approve buttons work server-side).
   - Do NOT share this login. The admin option never appears for any other number.
2. In Admin Panel you can: approve/reject deposits, change prices, post/delete notifications, edit Help and Contact pages.

## Step 4 — Notes PDFs (Drive)
1. Upload each notes PDF to your Google Drive (Muse prepares them with the StuDy watermark — a sample is in `notes/`).
2. For each file: Share → "Anyone with the link" (viewer).
3. Copy each file's ID from its link and add ONE line in `index.html` right after the `<script src="questions.js"></script>` tag:
   `<script>window.NOTES_FILES = { "rpsc_raj":"FILE_ID", "rpsc_bundle":"FILE_ID" };</script>`
   Keys: `<category>_<subject>` for subject notes, `<category>_bundle` for the all-subjects PDF.

## Step 5 — Payments
- The wallet QR codes in `assets/qr/` are already generated for your UPI ID `dineshjangirdjsp-1@oksbi` (₹100/150/200/250/300/500).
- Flow: user pays → submits 12-digit UTR → you approve in Admin Panel → stars credit instantly.
- True automatic payment detection is not possible with a plain UPI ID; the approval step is the safety net (also protects against fake UTRs — always match the UTR in your bank/UPI app before approving).

## Known limits (by design)
- Landscape lock during tests works on Android Chrome; on iPhone the test shows a "rotate your phone" strip instead.
- Push notifications to phones need a push server — in-app notifications work now; the service worker is push-ready for phase 2.
- Demo mode data is per-device. Go live (Step 2) before real users register.

## Content growth
- Question banks live in `questions.js` — one row per question (EN + HI). Muse adds fresh Current Affairs automatically every 3 days; new subject questions can be added the same way, then pushed to the repo.

## UI (v2 light redesign)
- Light, minimalist theme throughout — white screens, black line-style icons, compact type for phones.
- Registration collects the **student's name**; the profile/header avatar shows the **first initial** of the name.
- Category cards use the **official commission logos** (in `assets/cats/` — UPSC, RPSC, SSC).
- Contact page has a **Telegram button** (handle read from your Contact text, default @dj3680).
- A floating **Instagram button** (bottom-left, every screen except during a test) opens **@dinesh.jangir_**.

## New in this update
- **Header**: notification/profile live only in the bottom bar now. Top-right has two featured buttons — 🏆 Rewards (top-rankers leaderboard) and 🎁 Offers.
- **Payment notifications**: approving or rejecting a deposit automatically sends that user a personal notification (Alerts tab). Rejected Instagram-discount requests include a "DM on Instagram" button in the notification.
- **Admin → Add stars to a user**: enter phone + stars (+ optional rank & exam — the reward message auto-writes itself, e.g. "You stood at Rank #1 in RPSC Full Mock and won 50 stars!"). No payment needed; the user is notified instantly.
- **Instagram 10% discount**: on the Wallet page users can tick "Follow the owner & get 10% OFF" and pay 10% less. Such requests show a 📸 badge in your approvals with two choices: **Approve full ★** (follow verified) or **Approve ★(paid amount) only**, plus Reject.
- **Test screen**: the Submit popup now opens above the test screen (bug fixed); closing fullscreen mid-test asks Exit (back to the category page, attempt not counted) or Resume (fullscreen re-applied, timer paused during the popup). Question palette is pinned to the right edge.
- Backend (live mode) gained a `UserNotifs` sheet tab + `adminAddStars` and `leaderboard` actions — re-paste the new `Code.gs` and run `setup()` once more when you deploy.

## Update notes (v4)
- **Instagram tick reworked**: payment is always the FULL amount (QR unchanged). Ticking the follow box means the user gets **10% EXTRA stars on approval** (₹200 → 220★) — automatic when you tap Approve on a 📸-badged request.
- **Header buttons (left→right after the balance chip)**: Telegram (opens your chat with a preset help message), Help, 🏆 Rank, 🔍 Search (searches all tests/mocks/notes).
- **Rank list privacy**: names shown with masked phones — `XXXXX` + last 5 digits only.
- **Verified Report Card**: Profile → "My Report Card" generates a StuDy Verified card (name, points, tests, rank, average, accuracy + a unique `SDY-XXXX-XXXX-XXXX` code) downloadable as PDF via Print. Anyone can verify it in the app: Help → "Verify a Report Card" → enter the code. Live mode: re-paste `Code.gs` and run `setup()` once (adds the `Cards` tab).

## v5 — Repo flow (current)
- The app is now **self-contained**: the logo, commission logos and UPI QRs are embedded inside `index.html` itself. The repo only needs: `index.html`, `questions.js`, `manifest.webmanifest`, `sw.js`, the `notes/` folder (PDFs) and optionally `assets/icon-192.png` + `assets/icon-512.png` (PWA install icons).
- Muse pushes all code files (text) directly to `dj3680/StuDy` on request. **PDFs and PNGs cannot be pushed by Muse** (connector accepts text only) — upload the `notes/` PDFs once via GitHub's web UI: repo → `notes` folder (create via Add file → Create new file → name `notes/keep.txt`, then Add file → Upload files → select all 20 PDFs).
- Updates: just tell Muse in chat ("update current affairs", "add questions", "change price X") — code changes are pushed automatically; regenerated notes PDFs are delivered in chat for re-upload.
- GitHub Pages: repo Settings → Pages → Source: `main` / root. App URL: `https://dj3680.github.io/StuDy/`.
