"""
Weekly per-employee Qlik dashboard burst: for each row in a recipient
spreadsheet, render the Qlik Sense chart filtered to that person and email the
image via the local Outlook desktop app.

Design notes:
  * The per-employee filter is applied through the Qlik Single Integration API
    URL (&select=Field,Value) rather than by driving the filter pane.
  * Login is a scripted Qlik forms login against /internal_forms_authentication/.
    Credentials come from QLIK_USERNAME/QLIK_PASSWORD or a gitignored
    qlik_credentials.txt, never from this file.
  * Email is sent through the desktop Outlook via COM (no SMTP/OAuth setup).

See README.md for setup and configuration.
"""

import csv
import os
import sys
import time
from urllib.parse import quote, urlparse

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ==========================================================================
# CONFIG
# ==========================================================================

# Qlik host, no trailing slash (e.g. http://10.0.2.5). Include a virtual-proxy
# prefix here if your hub is behind one (e.g. http://10.0.2.5/sales).
TENANT_URL = "http://10.0.2.5"

# App (document) GUID, from the app URL between /app/ and /sheet/.
APP_ID = "48ab9fa2-5dc9-48f1-8b0e-25041e6313bd"

# Capture a single object (OBJECT_ID); if empty, capture the whole sheet
# (SHEET_ID). See README for how to find the ids.
OBJECT_ID = ""
SHEET_ID = "6527c8b7-f73a-4a7c-962c-6f347d52a009"

# The filter field, exactly as Qlik names it (case-sensitive).
FILTER_FIELD = "SALESPERSON_ORDER"

# Recipient list (.xlsx or .csv) and its column headers.
RECIPIENTS_FILE = os.path.join(SCRIPT_DIR, "recipients.xlsx")
NAME_COLUMN = "Employee"
EMAIL_COLUMN = "Email"

# Email content. {name} is substituted per employee; the image embeds inline
# via the cid referenced in the body.
EMAIL_SUBJECT = "Your weekly dashboard - {name}"
EMAIL_HTML_BODY = """
<p>Hi {name},</p>
<p>Here is your dashboard for this week:</p>
<p><img src="cid:dashboard_image"></p>
<p>Regards</p>
"""

# --- Test / safety switches -------------------------------------------------
REVIEW_MODE = True                                # True = open drafts, don't send
TEST_REDIRECT_EMAIL = "gerson@pennrosefarms.com"  # send all mail here; "" = real recipients
MAX_EMPLOYEES = 2                                 # cap rows processed; None = all

# Where screenshots are written (next to this script by default).
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "screenshots")

# --- Browser ----------------------------------------------------------------
# "chrome" uses installed Chrome, falling back to Playwright's bundled Chromium.
# Runs in a throwaway profile, so it never touches your real Chrome profile.
BROWSER_CHANNEL = "chrome"
HEADLESS = False              # True for silent scheduled runs
RENDER_SETTLE_SECONDS = 4     # settle time after the chart appears
READY_SELECTOR = ""           # CSS selector that appears once drawn; "" = generic

# --- Qlik login -------------------------------------------------------------
# Credentials: env vars QLIK_USERNAME/QLIK_PASSWORD, else a gitignored
# qlik_credentials.txt (KEY=VALUE). If unset and headed, prompts for a manual
# browser login.
CREDENTIALS_FILE = os.path.join(SCRIPT_DIR, "qlik_credentials.txt")
# Qlik Sense default login-form fields; override if your login page is customized.
LOGIN_USERNAME_SELECTOR = 'input[name="username"], input#username, input[type="text"]'
LOGIN_PASSWORD_SELECTOR = 'input[name="pwd"], input[name="password"], input[type="password"]'
LOGIN_SUBMIT_SELECTOR = 'button[type="submit"], input[type="submit"], #loginbtn, .submit-button'

# The Qlik host usually sits behind Windows Integrated Auth (NTLM/Negotiate).
# The automated browser answers that challenge silently with the VM's current
# Windows login, the same SSO your normal Chrome does before the Qlik form even
# shows. Without it a fresh profile fails with net::ERR_INVALID_AUTH_CREDENTIALS.
# Set to "" if your server doesn't use integrated auth.
AUTH_SERVER_ALLOWLIST = urlparse(TENANT_URL).hostname or ""

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
        name = str(r.get(NAME_COLUMN) or "").strip()
        email = str(r.get(EMAIL_COLUMN) or "").strip()
        if name and email:
            people.append({"name": name, "email": email})
    if not people:
        sys.exit(
            f"No rows found. Check that '{NAME_COLUMN}' and '{EMAIL_COLUMN}' "
            f"match the column headers in {path}."
        )
    return people


def read_credentials():
    """Return (username, password) from env vars, then qlik_credentials.txt.
    Returns ("", "") if neither is configured."""
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
    sel = f"{quote(FILTER_FIELD, safe='')},{quote(str(employee_name), safe='')}"
    params.append(f"select={sel}")
    params.append("opt=nointeraction")  # hide selection bars/toolbars for a clean shot
    return base + "?" + "&".join(params)


def on_login_page(page):
    """True if we're on Qlik's forms-login page rather than the app."""
    if "internal_forms_authentication" in (page.url or "").lower():
        return True
    # Otherwise fall back to detecting a password field; times out on the hub.
    try:
        page.wait_for_selector('input[type="password"]', timeout=2500)
        return True
    except Exception:
        return False


def submit_login(page, username, password):
    """Fill Qlik's forms-login page and submit."""
    print("   login page detected; signing in with the Qlik service account...")
    page.locator(LOGIN_USERNAME_SELECTOR).first.fill(username)
    page.locator(LOGIN_PASSWORD_SELECTOR).first.fill(password)
    try:
        page.locator(LOGIN_SUBMIT_SELECTOR).first.click(timeout=3000)
    except Exception:
        page.locator(LOGIN_PASSWORD_SELECTOR).first.press("Enter")  # no submit button
    try:
        page.wait_for_load_state("networkidle", timeout=30000)
    except Exception:
        pass


def manual_login_pause():
    """Headed-only fallback: wait for a human to log in by hand, then continue."""
    print("\n" + "=" * 70)
    print(" Qlik is showing its LOGIN page in the browser window.")
    print(" Log in there by hand, then come back to this terminal and")
    input(" press ENTER once you can see your Qlik hub/app... ")
    print("=" * 70 + "\n")


def ensure_logged_in(page, username, password):
    """Open the hub and, if Qlik shows its forms-login page, authenticate with
    the configured credentials (or, headed and without them, pause for a manual
    login)."""
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
    selector = READY_SELECTOR or ".qv-object"
    try:
        page.wait_for_selector(selector, timeout=30000)
    except Exception:
        print(f"   ! '{selector}' never appeared for {employee_name}; "
              f"screenshotting whatever is on screen.")
    time.sleep(RENDER_SETTLE_SECONDS)

    element = None
    if OBJECT_ID:
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

    attachment = mail.Attachments.Add(os.path.abspath(image_path))
    try:
        # PR_ATTACH_CONTENT_ID: lets HTMLBody reference the image with cid:
        attachment.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
            "dashboard_image",
        )
    except Exception:
        pass  # if the cid fails, the image still rides along as an attachment
    mail.HTMLBody = EMAIL_HTML_BODY.format(name=name)

    if REVIEW_MODE:
        mail.Display(False)  # open as a draft; does not send
        return "drafted"
    mail.Send()
    return "sent"


def launch_browser(p):
    """Launch the configured browser in a throwaway profile, falling back to
    Playwright's bundled Chromium if the installed Chrome channel isn't found."""
    args = []
    if AUTH_SERVER_ALLOWLIST:
        args += [
            f"--auth-server-allowlist={AUTH_SERVER_ALLOWLIST}",
            f"--auth-negotiate-delegate-allowlist={AUTH_SERVER_ALLOWLIST}",
        ]
    try:
        return p.chromium.launch(channel=BROWSER_CHANNEL, headless=HEADLESS, args=args)
    except Exception as e:
        print(f"   ! couldn't launch '{BROWSER_CHANNEL}' ({e}); "
              f"falling back to Playwright's bundled Chromium.")
        return p.chromium.launch(headless=HEADLESS, args=args)


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
