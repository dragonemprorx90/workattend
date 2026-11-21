# دوام · WorkAttend

Employee attendance app — Arabic + English.

- **5 statuses:** حضور Present · غياب Absent · مرضية Sick leave · إجازة Leave · سنوية Annual leave
- **Employees:** add, edit, delete, or import a list from **Excel (.xlsx/.xls/.ods), CSV/TSV, contacts (.vcf), text or JSON** — or paste names.
- **WhatsApp:** write a message in Arabic *or* English → the app makes it **Arabic + English** (Google Translate) → opens WhatsApp ready to send to your group or manager. Optional daily attendance summary.
- **Reports:** monthly totals per status, per employee, per day; export to Excel (CSV).
- **Shared database:** a Google Sheet owned by whoever runs it. Works offline; syncs automatically.
- **Android app** (`WorkAttend.apk`) and **web app** (this site — on iPhone use Safari → Share → *Add to Home Screen*).

---

## 1 · Server setup (the owner, once — about 3 minutes)

1. Create a Google Sheet (any name, e.g. *WorkAttend Database*).
2. **Extensions → Apps Script** → delete what's there → paste all of [`server/Code.gs`](server/Code.gs) → **Save**.
3. In the toolbar choose the function **`setup`** → **▶ Run** → allow the permissions (Advanced → Go to project).
   The tabs *Employees, Attendance, Messages, Settings* appear in the Sheet.
4. **Deploy → New deployment** → ⚙ type **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
   → **Deploy** → copy the **Web app URL** (ends in `/exec`).
5. The PIN is in the **Settings** tab, cell **B2** (default `WORK`). Change it any time.

## 2 · Phones

1. Open the app (APK or website) → paste the **Web app URL** and the **PIN** → **Connect**.
2. **Settings → Share setup link** — send it to the other phones; one tap configures them.

## 3 · Handing it over to a new owner

Nothing in the app is tied to one person.

1. **Google Drive** → the Sheet → Share → make the new person **Owner** (the script moves with the Sheet).
2. The new owner opens the Sheet → Extensions → Apps Script → **Deploy → New deployment** (same settings as above) → gets a **new URL**.
3. In the app: **Settings** → paste the new URL (+ PIN) → **Connect** → **Share setup link** to everyone.
4. Optional: transfer this GitHub repository (Settings → Transfer). The website address changes to the new owner's GitHub name — update the one line `window.WEB_URL` at the top of the script in `index.html`.

The data lives only in the Sheet and on the phones — there is no other server, account or cost.
