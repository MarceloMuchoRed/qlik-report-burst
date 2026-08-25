"""
qlik_report_burst.py
---------------------
Weekly "burst" job: for each employee in a spreadsheet, open the Qlik Sense
dashboard filtered to that employee, screenshot the chart, and email the image
to that person via the local (already-signed-in) Outlook desktop app.

What this is (and is NOT):
  This is plain, deterministic browser automation. There is NO AI / ML / LLM
  anywhere in it. Step by step it does exactly this and nothing else:
    1. Open a headless browser and log into Qlik once, using a Qlik service
       account (a username + password), the same way a person would type them
       into the login page.
    2. For each row in the recipient list, request one Qlik chart URL with that
       person's filter applied, and save the rendered chart as a PNG.
    3. Attach that PNG to an email and send it through the desktop Outlook app.
  It reads no other data, makes no decisions, and changes nothing in Qlik. It is
  a login-export-email bot.

Why this design:
  * The per-employee filter is applied through the Qlik Single Integration API
    URL (...&select=Field,Value) instead of clicking around the filter pane, so
    it doesn't break when the UI changes.
  * Login is a scripted Qlik forms login: the script fills the username/password
    on Qlik's own login page (/internal_forms_authentication/) and submits. The
    credentials come from environment variables (or a local, gitignored file),
    never from the code, so nothing secret is ever committed.
  * Email goes through your desktop Outlook via COM automation, so there are no
    SMTP servers, app passwords, or OAuth apps to set up.

No admin rights required. Install into your portable Python with:
    python -m pip install playwright pywin32 openpyxl
The script drives your installed Chrome (BROWSER_CHANNEL = "chrome"); it only
needs Playwright's bundled Chromium as a fallback, so this is optional:
    python -m playwright install chromium

Provide the Qlik service-account credentials (do NOT put them in this file):
    setx QLIK_USERNAME "the-service-account"
    setx QLIK_PASSWORD "the-password"
  ...or create a gitignored qlik_credentials.txt next to this script:
    QLIK_USERNAME=the-service-account
    QLIK_PASSWORD=the-password

Then edit the CONFIG block below and run:
    python qlik_report_burst.py
"""

import csv
import os
import sys
import time
from urllib.parse import quote

# Folder this script lives in; all default paths below are relative to it so
# the project stays portable (works wherever the repo is cloned).
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ==========================================================================
# CONFIG  ---  edit everything in this block, then run the script.
# ==========================================================================

# --- Your Qlik server + app --------------------------------------------------
# The host portion of your Qlik URL, no trailing slash. This is set up for
# Qlik Sense Enterprise on Windows (on-prem / client-managed), e.g.
#   http://10.0.2.5
# If your hub is reached through a virtual proxy PREFIX (e.g. your hub URL is
#   http://10.0.2.5/sales/hub/...  ), include that prefix here too:
#   TENANT_URL = "http://10.0.2.5/sales"
# so the script builds .../sales/single/... instead of /single/.
# (Qlik Sense Cloud users would instead use https://yourtenant.region.qlikcloud.com)
TENANT_URL = "http://10.0.2.5"

# The app (document) GUID. Open the app from the hub and read the URL:
#   http://10.0.2.5/sense/app/76daa5f7-c3b3-40a7-a8a1-8c1453b2acf6/sheet/<SHEET_ID>/state/analysis
# Paste the part between /app/ and /sheet/ here.
APP_ID = "48ab9fa2-5dc9-48f1-8b0e-25041e6313bd"

# What to capture. Use EITHER a single object OR a whole sheet:
#   OBJECT_ID -> cleanest: a single chart/table. On-prem, get it from the
#                Single Configurator in Dev Hub:
#                  http://10.0.2.5/dev-hub/single-configurator
#                Pick app -> sheet -> object; it previews and builds the exact
#                /single/ URL, and shows the object id.
#   SHEET_ID  -> screenshots the whole sheet page instead. Leave OBJECT_ID = ""
#                and set SHEET_ID (the GUID after /sheet/ in the app URL) to
#                capture the full dashboard sheet.
OBJECT_ID = ""  # empty -> capture the whole sheet (SHEET_ID below) instead of one chart
SHEET_ID = "1266bc38-8212-4401-aa26-b3652bb6483d"  # the sheet to screenshot

# The field you currently change in the filter, EXACTLY as Qlik names it
# (case-sensitive), e.g. "Employee Name" or "EmployeeID".
FILTER_FIELD = "SALESPERSON_ORDER"

# --- Your recipient spreadsheet ---------------------------------------------
# Path to your employee -> email list. .xlsx (needs openpyxl) or .csv both work.
# Defaults to recipients.xlsx sitting next to this script.
RECIPIENTS_FILE = os.path.join(SCRIPT_DIR, "recipients.xlsx")
# The column headers in that file:
NAME_COLUMN = "Employee"      # the value that gets put into the Qlik filter
EMAIL_COLUMN = "Email"        # where the screenshot is sent

# --- Email content ----------------------------------------------------------
EMAIL_SUBJECT = "Your weekly dashboard - {name}"
# {name} is substituted per employee. The screenshot is embedded inline below.
EMAIL_HTML_BODY = """
<p>Hi {name},</p>
<p>Here is your dashboard for this week:</p>
<p><img src="cid:dashboard_image"></p>
<p>Regards</p>
"""

# --- Test / safety switches (use these for your first few runs) -------------
# REVIEW_MODE = True opens each email as a DRAFT (does not send) so you can
# eyeball it. Set to False only when you're happy and ready to actually send.
REVIEW_MODE = True
# If set to an address, ALL emails go there instead of the real recipients.
# Great for a first end-to-end test to yourself. Set to "" to use real emails.
TEST_REDIRECT_EMAIL = "gerson@pennrosefarms.com"
# Process at most this many rows (None = all). Keep it small while testing.
MAX_EMPLOYEES = 2

# --- Rendering / browser ----------------------------------------------------
# Where the screenshots are written (next to this script by default).
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "screenshots")

# The browser to drive. "chrome" = your installed Google Chrome (no download
# needed). If that isn't present, the script falls back to Playwright's bundled
# Chromium (run `python -m playwright install chromium` once to have it ready).
# The browser runs in a fresh, throwaway profile every time — it never touches
# your real Chrome profile, so your normal Chrome can stay open and nothing has
# to be closed or copied.
BROWSER_CHANNEL = "chrome"
# headless=False shows the browser window (useful for a first run / live demo,
# and required if you want to log in by hand when no credentials are set).
# headless=True runs it invisibly for silent scheduled runs.
HEADLESS = False
# Seconds to wait after the chart appears, for animations/data to settle.
RENDER_SETTLE_SECONDS = 4
# Optional: a CSS selector that only appears once the chart is fully drawn.
# Leave as "" to use the generic Qlik object container.
READY_SELECTOR = ""

# --- Qlik login (service account) -------------------------------------------
# This server uses Qlik Sense INTERNAL FORMS authentication: opening the hub
# redirects to a web login page (/internal_forms_authentication/) with a
# username and password field. (Confirmed by curl: /hub/ 302s to that page, and
# Windows/NTLM SSO does nothing here.) The script logs in by filling that form.
#
# Credentials are read from, in order:
#   1. environment variables QLIK_USERNAME / QLIK_PASSWORD
#   2. a local file `qlik_credentials.txt` next to this script (KEY=VALUE lines)
# They are intentionally NOT stored in this file, so the code can be committed
# and shared without leaking the password. If neither is set and HEADLESS=False,
# the script pauses so you can log in by hand in the browser window (handy for a
# demo before the service-account password has been issued).
CREDENTIALS_FILE = os.path.join(SCRIPT_DIR, "qlik_credentials.txt")
# CSS selectors for the Qlik login form. These are the Qlik Sense defaults
# (field names `username` and `pwd`); .first is used so a slightly customized
# login template still matches. Only touch these if your login page is heavily
# customized and the defaults don't find the fields.
LOGIN_USERNAME_SELECTOR = 'input[name="username"], input#username, input[type="text"]'
LOGIN_PASSWORD_SELECTOR = 'input[name="pwd"], input[name="password"], input[type="password"]'
LOGIN_SUBMIT_SELECTOR = 'button[type="submit"], input[type="submit"], #loginbtn, .submit-button'

# ==========================================================================
# END CONFIG.  You normally don't need to edit below this line.
# ==========================================================================


def load_recipients(path):
    """Return a list of {name, email} dicts from an .xlsx or .csv file."""
    if not os.path.exists(path):
        sys.exit(f"Recipients file not found: {path}")

    ext = os.path.splitext(path)[1].lower()
    rows = []
    if ext == ".csv":
        with open(path, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                rows.append(r)
    elif ext in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        header = [str(c.value).strip() if c.value is not None else "" for c in next(ws.iter_rows(max_row=1))]
        for row in ws.iter_rows(min_row=2, values_only=True):
            rows.append({header[i]: row[i] for i in range(len(header))})
    else:
        sys.exit(f"Unsupported recipients file type: {ext} (use .csv or .xlsx)")

    people = []
    for r in rows:
        name = (r.get(NAME_COLUMN) or "")
        email = (r.get(EMAIL_COLUMN) or "")
        name = str(name).strip()
        email = str(email).strip()
        if name and email:
            people.append({"name": name, "email": email})
    if not people:
        sys.exit(
            f"No rows found. Check that '{NAME_COLUMN}' and '{EMAIL_COLUMN}' "
            f"match the column headers in {path}."
        )
    return people


def read_credentials():
    """Return (username, password) for the Qlik service account.

    Looked up from environment variables first, then a local gitignored
    qlik_credentials.txt (KEY=VALUE lines). Never stored in this script. Returns
    ("", "") if nothing is configured, in which case the caller falls back to a
    manual browser login."""
    user = os.environ.get("QLIK_USERNAME", "").strip()
    pwd = os.environ.get("QLIK_PASSWORD", "")
    if user and pwd:
        return user, pwd
    if os.path.exists(CREDENTIALS_FILE):
        try:
            with open(CREDENTIALS_FILE, encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith("#") or "=" not in line:
                        continue
                    key, val = line.split("=", 1)
                    key = key.strip().upper()
                    val = val.strip()
                    if key in ("QLIK_USERNAME", "USERNAME") and not user:
                        user = val
                    elif key in ("QLIK_PASSWORD", "PASSWORD") and not pwd:
                        pwd = val
        except OSError:
            pass
    return user, pwd


def build_single_url(employee_name):
    """Build a Single Integration API URL with the employee selection applied."""
    base = TENANT_URL.rstrip("/") + "/single/"
    params = [f"appid={quote(APP_ID, safe='')}"]
    if OBJECT_ID:
        params.append(f"obj={quote(OBJECT_ID, safe='')}")
    elif SHEET_ID:
        params.append(f"sheet={quote(SHEET_ID, safe='')}")
    else:
        sys.exit("Set either OBJECT_ID or SHEET_ID in CONFIG.")
    # select=Field,Value  (each part URL-encoded, comma kept as the delimiter)
    sel = f"{quote(FILTER_FIELD, safe='')},{quote(str(employee_name), safe='')}"
    params.append(f"select={sel}")
    # opt=nointeraction hides selection bars/toolbars for a cleaner shot.
    params.append("opt=nointeraction")
    return base + "?" + "&".join(params)


def on_login_page(page):
    """True if we're sitting on Qlik's forms-login page rather than the app."""
    if "internal_forms_authentication" in (page.url or "").lower():
        return True
    # A visible password field also means we're on a login form. If we're on the
    # hub instead, this simply times out and returns False.
    try:
        page.wait_for_selector('input[type="password"]', timeout=2500)
        return True
    except Exception:
        return False


def submit_login(page, username, password):
    """Fill Qlik's forms-login page with the service account and submit."""
    print("   login page detected; signing in with the Qlik service account...")
    page.locator(LOGIN_USERNAME_SELECTOR).first.fill(username)
    page.locator(LOGIN_PASSWORD_SELECTOR).first.fill(password)
    try:
        page.locator(LOGIN_SUBMIT_SELECTOR).first.click(timeout=3000)
    except Exception:
        # No obvious submit button on this template: submit by pressing Enter.
        page.locator(LOGIN_PASSWORD_SELECTOR).first.press("Enter")
    try:
        page.wait_for_load_state("networkidle", timeout=30000)
    except Exception:
        pass


def manual_login_pause():
    """Headed-only fallback: let a human log in by hand, then continue. Used for
    a live demo before the service-account password has been issued."""
    print("\n" + "=" * 70)
    print(" Qlik is showing its LOGIN page in the browser window.")
    print(" Log in there by hand, then come back to this terminal and")
    input(" press ENTER once you can see your Qlik hub/app... ")
    print("=" * 70 + "\n")


def ensure_logged_in(page, username, password):
    """Open the hub and, if Qlik shows its forms-login page, authenticate. Uses
    the configured service-account credentials when present; otherwise (headed
    only) pauses for a manual login."""
    target = TENANT_URL.rstrip("/") + "/hub/"
    print(f"Opening Qlik: {target}")
    try:
        page.goto(target, wait_until="domcontentloaded", timeout=60000)
        time.sleep(2)
        print(f"   landed at: {page.url}")
    except Exception as e:
        print(f"   ! couldn't load the page ({e}); will check for a login form.")

    if not on_login_page(page):
        print("   already logged in (no login form shown).")
        return

    if username and password:
        try:
            submit_login(page, username, password)
        except Exception as e:
            print(f"   ! scripted login failed ({e}).")
        if on_login_page(page):
            print("   ! still on the login page after submitting "
                  "(check the credentials / field selectors).")
            if not HEADLESS:
                manual_login_pause()
        else:
            print(f"   logged in; now at: {page.url}")
        return

    # No credentials configured.
    if HEADLESS:
        sys.exit(
            "Qlik requires a login but no QLIK_USERNAME/QLIK_PASSWORD is set and "
            "the run is HEADLESS (can't prompt). Set the credentials (env vars or "
            "qlik_credentials.txt) or run with HEADLESS = False to log in by hand."
        )
    print("   no credentials configured; log in by hand in the browser window.")
    manual_login_pause()


def capture(page, employee_name, out_path):
    """Load the filtered single view and screenshot it."""
    url = build_single_url(employee_name)
    page.goto(url, wait_until="networkidle")
    # Wait for the Qlik object container to exist, then let it settle.
    selector = READY_SELECTOR or ".qv-object"
    try:
        page.wait_for_selector(selector, timeout=30000)
    except Exception:
        print(f"   ! '{selector}' never appeared for {employee_name}; "
              f"screenshotting whatever is on screen.")
    time.sleep(RENDER_SETTLE_SECONDS)

    element = None
    if OBJECT_ID:
        # Screenshot just the chart element for a tight, clean image.
        element = page.query_selector(".qv-object") or page.query_selector(selector)
    if element:
        element.screenshot(path=out_path)
    else:
        page.screenshot(path=out_path, full_page=True)


def send_email(name, to_address, image_path):
    """Create (and optionally send) an email through the desktop Outlook app."""
    import win32com.client  # from pywin32
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0 = olMailItem
    mail.To = to_address
    mail.Subject = EMAIL_SUBJECT.format(name=name)

    # Attach the image and mark it inline via a Content-ID so it shows in-body.
    attachment = mail.Attachments.Add(os.path.abspath(image_path))
    try:
        # PR_ATTACH_CONTENT_ID -> lets HTMLBody reference it with cid:
        attachment.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
            "dashboard_image",
        )
    except Exception:
        pass  # if inline fails, the image still rides along as an attachment
    mail.HTMLBody = EMAIL_HTML_BODY.format(name=name)

    if REVIEW_MODE:
        mail.Display(False)  # open as a draft for review; does NOT send
        return "drafted"
    mail.Send()
    return "sent"


def launch_browser(p):
    """Launch the configured browser in a fresh throwaway profile. Falls back to
    Playwright's bundled Chromium if the installed Chrome channel isn't found."""
    try:
        return p.chromium.launch(channel=BROWSER_CHANNEL, headless=HEADLESS)
    except Exception as e:
        print(f"   ! couldn't launch '{BROWSER_CHANNEL}' ({e}); "
              f"falling back to Playwright's bundled Chromium.")
        return p.chromium.launch(headless=HEADLESS)


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    people = load_recipients(RECIPIENTS_FILE)
    if MAX_EMPLOYEES is not None:
        people = people[:MAX_EMPLOYEES]

    username, password = read_credentials()

    print(f"Processing {len(people)} employee(s). "
          f"REVIEW_MODE={REVIEW_MODE}  "
          f"redirect={'ON -> ' + TEST_REDIRECT_EMAIL if TEST_REDIRECT_EMAIL else 'off'}")
    print(f"Qlik login: "
          f"{'service account ' + username if username else 'NONE set (will prompt in the browser)'}")

    print("Starting the browser engine...", flush=True)
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        browser = launch_browser(p)
        context = browser.new_context()
        page = context.new_page()

        ensure_logged_in(page, username, password)

        for i, person in enumerate(people, 1):
            name = person["name"]
            recipient = TEST_REDIRECT_EMAIL or person["email"]
            safe = "".join(c for c in name if c.isalnum() or c in " _-").strip()
            img = os.path.join(OUTPUT_DIR, f"{i:02d}_{safe or 'employee'}.png")

            print(f"[{i}/{len(people)}] {name} -> {recipient}")
            try:
                capture(page, name, img)
                status = send_email(name, recipient, img)
                print(f"   screenshot saved: {img}")
                print(f"   email {status}")
            except Exception as e:
                print(f"   ! FAILED for {name}: {e}")

        context.close()
        browser.close()

    print("\nDone. Review the drafts in Outlook (REVIEW_MODE) or check Sent.")


if __name__ == "__main__":
    main()
