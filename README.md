# qlik-report-burst

Automates the weekly manual job: for each employee, filter the Qlik Sense
dashboard to that person, screenshot it, and email the image to them.

Written for **Qlik Sense Enterprise on Windows (on-prem / client-managed)** —
a hub reached at something like `http://10.0.2.5/hub/`. (It also works against
Qlik Sense Cloud; just point `TENANT_URL` at your `*.qlikcloud.com` tenant.)

- **Filter** is applied via the Qlik Single Integration API URL (`&select=Field,Value`) — no fragile clicking in the filter pane.
- **Login** is never scripted. You log in **once** in the browser the script opens; the session is remembered in a local profile folder. No password is stored.
- **Email** goes through your already-signed-in **desktop Outlook** (COM automation) — no SMTP, app passwords, or OAuth.

No admin rights required.

---

## 1. One-time setup (portable Python)

From the folder containing your portable `python.exe`:

```powershell
python -m pip install playwright pywin32 openpyxl
python -m playwright install chromium
```

`playwright install` downloads Chromium into your user cache (`%USERPROFILE%\AppData\Local\ms-playwright`) — **no admin needed**.

If `import win32com` later errors, run once:
```powershell
python <portable-python-path>\Scripts\pywin32_postinstall.py -install
```

## 2. Make your recipients file

Create `recipients.xlsx` (or `.csv`) next to the script with two columns. The
header names must match `NAME_COLUMN` / `EMAIL_COLUMN` in the script:

| Employee        | Email                   |
|-----------------|-------------------------|
| Jane Smith      | jane@company.com        |
| John Doe        | john@company.com        |

`Employee` values must match the Qlik field values **exactly**.

## 3. Find your Qlik IDs and fill in CONFIG

Open the script's `CONFIG` block and set:

- **TENANT_URL** — the host part of your Qlik URL, no trailing slash, e.g.
  `http://10.0.2.5`. If your hub is behind a virtual proxy **prefix** (hub URL
  looks like `http://10.0.2.5/sales/hub/…`), include it: `http://10.0.2.5/sales`
  — the script builds `.../sales/single/…` from it. With no prefix (your case),
  it builds `http://10.0.2.5/single/…`.
- **APP_ID** — open the app from the hub and read the URL:
  `http://10.0.2.5/sense/app/<APP_ID>/sheet/<SHEET_ID>/state/analysis`. Paste the
  GUID between `/app/` and `/sheet/`.
- **OBJECT_ID** — the chart to capture. On-prem, use the **Single Configurator**
  in Dev Hub: `http://10.0.2.5/dev-hub/single-configurator`. Pick app → sheet →
  object; it live-previews and builds the exact `/single/` URL and shows the
  object id. (If Dev Hub is disabled by your admin, leave `OBJECT_ID = ""` and set
  **SHEET_ID** — the GUID after `/sheet/` in the app URL — to screenshot the whole
  sheet instead.)
- **FILTER_FIELD** — the field you change today, spelled exactly as Qlik shows
  it (case-sensitive), e.g. `Employee Name`.

The script then requests, per employee:
`http://10.0.2.5/single/?appid=<APP_ID>&obj=<OBJECT_ID>&select=<FILTER_FIELD>,<Employee>&opt=nointeraction`

## 4. Test on demand (safe defaults are already set)

The script ships in **test mode**:
- `REVIEW_MODE = True` → emails open as **drafts** (nothing sends).
- `TEST_REDIRECT_EMAIL = "marcelo@pennrosefarms.com"` → all mail addressed to you.
- `MAX_EMPLOYEES = 2` → only the first 2 rows are processed.

Run it:
```powershell
python qlik_report_burst.py
```

First run: a browser opens → **log into Qlik once** → return to the terminal and
press **Enter**. It then screenshots each employee and opens draft emails for you
to inspect. Check the screenshots in `screenshots\` and the drafts in Outlook.

## 5. Go live

When the drafts look right:
1. `MAX_EMPLOYEES = None` (all employees)
2. `TEST_REDIRECT_EMAIL = ""` (real recipients)
3. `REVIEW_MODE = False` (actually send)

## 6. (Later) run it weekly, unattended

Use **Windows Task Scheduler** → new task → run `python.exe` with argument
`qlik_report_burst.py`, weekly. Note: an unattended run needs a still-valid Qlik
session in the profile folder; if the session has expired it will pause for a
manual login, so for a hands-off weekly job keep an eye on the first automated
run and re-login when prompted.

---

## Notes / gotchas

- If the chart screenshots before it finishes drawing, increase
  `RENDER_SETTLE_SECONDS`, or set `READY_SELECTOR` to a CSS selector that only
  appears once your specific chart is fully rendered.
- Don't delete `.browser-profile\` — that's what keeps you logged in.
- This path does **not** depend on any extra Qlik reporting product. On on-prem,
  the native server-side equivalent of a "burst report" is **Qlik NPrinting** (a
  separately licensed/installed add-on that can filter per recipient and email
  PDFs/images on a schedule). If your site already runs NPrinting, that would be
  lower maintenance than this script — worth checking with your admin.
