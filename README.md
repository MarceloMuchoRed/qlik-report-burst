# qlik-report-burst

Automates a weekly manual job: for each employee, filter the Qlik Sense
dashboard to that person, screenshot the chart, and email the image to them.

Written for **Qlik Sense Enterprise on Windows (on-prem / client-managed)** —
a hub reached at something like `http://10.0.2.5/hub/`. (It also works against
Qlik Sense Cloud; just point `TENANT_URL` at your `*.qlikcloud.com` tenant.)

## What it is (and is not)

This is plain, deterministic browser automation. **There is no AI / ML / LLM
anywhere in it.** It does exactly three things, in order, and nothing else:

1. **Log in** to Qlik once, using a service account (a username and password) —
   the same credentials a person would enter, sent by the script.
2. **Export** one chart per person: it requests a Qlik chart URL with that
   person's filter applied and saves the rendered chart as a PNG.
3. **Email** that PNG to the person through the desktop Outlook app.

It reads no other data, makes no decisions, and changes nothing inside Qlik.
It's a login-export-email bot.

- **Filter** is applied via the Qlik Single Integration API URL
  (`&select=Field,Value`) — no fragile clicking in the filter pane.
- **Login**: the Qlik virtual proxy uses **Windows authentication**
  (`/internal_windows_authentication/`) — an NTLM/Negotiate challenge, i.e. the
  browser credential popup, not a web form. The script answers that challenge
  with the service account via the browser's HTTP credentials. This matters:
  without it the browser would silently sign in as the machine's own Windows
  login, which typically has no Qlik access pass. Credentials come from
  environment variables or a local, gitignored file — **never from the code**.
- **Browser** runs in a fresh throwaway profile every time. It never touches
  your real Chrome profile, so your normal Chrome can stay open, and it runs
  headless for silent scheduled runs.
- **Email** goes through your already-signed-in **desktop Outlook** (COM
  automation) — no SMTP, app passwords, or OAuth.

No admin rights required.

---

## 1. One-time setup (portable Python)

From the folder containing your portable `python.exe`:

```powershell
python -m pip install playwright pywin32 openpyxl
```

The script drives your already-installed Chrome (`BROWSER_CHANNEL = "chrome"`),
so you normally do **not** need Playwright's bundled Chromium. It's only used as
a fallback if Chrome isn't found; to have it ready, run once:

```powershell
python -m playwright install chromium
```

If `import win32com` later errors, run once:
```powershell
python <portable-python-path>\Scripts\pywin32_postinstall.py -install
```

## 2. Provide the Qlik service-account credentials

The credentials are **not** stored in the script. Set them one of two ways:

Environment variables (good for a logged-in user / scheduled task):
```powershell
setx QLIK_USERNAME "the-service-account"
setx QLIK_PASSWORD "the-password"
```

...or a local file `qlik_credentials.txt` next to the script (it's gitignored,
so it never gets committed):
```
QLIK_USERNAME=the-service-account
QLIK_PASSWORD=the-password
```

Use the account your admin licensed for this (the one that works when you log in
by hand), **not** the machine's own Windows login. If the account is
domain-joined and a bare username is rejected, use `DOMAIN\user` or
`user@domain`. Credentials are required — the script exits if none are set.

## 3. Make your recipients file

Create `recipients.xlsx` (or `.csv`) next to the script with two columns. The
header names must match `NAME_COLUMN` / `EMAIL_COLUMN` in the script:

| Employee        | Email                   |
|-----------------|-------------------------|
| Jane Smith      | jane@company.com        |
| John Doe        | john@company.com        |

`Employee` values must match the Qlik field values **exactly**.

## 4. Find your Qlik IDs and fill in CONFIG

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

## 5. Test on demand (safe defaults are already set)

The script ships in **test mode**:
- `REVIEW_MODE = True` → emails open as **drafts** (nothing sends).
- `TEST_REDIRECT_EMAIL` is set → all mail is addressed to that test inbox.
- `MAX_EMPLOYEES = 2` → only the first 2 rows are processed.

Run it:
```powershell
python qlik_report_burst.py
```

With credentials set, it logs in automatically, screenshots each employee, and
opens draft emails for you to inspect. Check the screenshots in `screenshots\`
and the drafts in Outlook. With `HEADLESS = False` you can watch it happen.

## 6. Go live

When the drafts look right:
1. `MAX_EMPLOYEES = None` (all employees)
2. `TEST_REDIRECT_EMAIL = ""` (real recipients)
3. `REVIEW_MODE = False` (actually send)

## 7. Run it weekly, unattended

Use **Windows Task Scheduler** → new task → run `python.exe` with argument
`qlik_report_burst.py`, weekly. For a fully silent run set `HEADLESS = True`.
The bot sends the service account on every run, so there's no saved session to
expire and no prompt to answer. (Make sure the credentials are available to the
scheduled task's account, via env vars or `qlik_credentials.txt`.)

---

## Notes / gotchas

- If the chart screenshots before it finishes drawing, increase
  `RENDER_SETTLE_SECONDS`, or set `READY_SELECTOR` to a CSS selector that only
  appears once your specific chart is fully rendered.
- If you get **"You cannot access Qlik Sense because you have no access pass"**,
  the bot authenticated as the wrong (unlicensed) Windows identity. Make sure
  `QLIK_USERNAME`/`QLIK_PASSWORD` hold the licensed service account, and try a
  `DOMAIN\user` / `user@domain` form if a bare username is rejected.
- If the run stays on `/internal_windows_authentication/`, the credentials were
  rejected — same fix as above.
- This path does **not** depend on any extra Qlik reporting product. On on-prem,
  the native server-side equivalent of a "burst report" is **Qlik NPrinting** (a
  separately licensed/installed add-on that can filter per recipient and email
  PDFs/images on a schedule). If your site already runs NPrinting, that would be
  lower maintenance than this script — worth checking with your admin.
