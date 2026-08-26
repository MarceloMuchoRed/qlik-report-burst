"""
Weekly per-employee Qlik dashboard burst: for each row in a recipient
spreadsheet, render the Qlik Sense chart filtered to that person and email the
image via the local Outlook desktop app.

Design notes:
  * The per-employee filter is applied through the Qlik Single Integration API
    URL (&select=Field,Value) rather than by driving the filter pane.
  * The Qlik virtual proxy uses Windows authentication
    (/internal_windows_authentication/) — an NTLM/Negotiate challenge (the
    browser credential popup), not a web form. The bot answers that challenge
    with a specific service account via Playwright's http_credentials, so it
    authenticates as that licensed account instead of silently single-signing-on
    as the VM's own (unlicensed) Windows login. Credentials come from
    QLIK_USERNAME/QLIK_PASSWORD or a gitignored qlik_credentials.txt, never from
    this file.
  * Email is sent through the desktop Outlook via COM (no SMTP/OAuth setup).

See README.md for setup and configuration.
"""

import csv
import os
import sys
import time
from urllib.parse import quote

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
# Screenshot canvas. A larger viewport makes Qlik render its full desktop layout
# (a small viewport triggers a cramped/rescaled "small screen" layout); bump it
# if your dashboard is bigger. DEVICE_SCALE = 2 doubles pixel density for crisp
# images (a 1920x1080 viewport then yields a 3840x2160 shot).
VIEWPORT_WIDTH = 1920
VIEWPORT_HEIGHT = 1080
DEVICE_SCALE = 2
RENDER_SETTLE_SECONDS = 6     # extra settle after objects load, for animations
READY_SELECTOR = ""           # CSS selector that appears once drawn; "" = generic

# --- Qlik login (Windows authentication) ------------------------------------
# The Qlik virtual proxy challenges with Windows auth (NTLM/Negotiate). The bot
# answers that challenge with the credentials below (a LICENSED service account,
# e.g. gmrqlik). Do NOT rely on the VM's own Windows login: that signs in
# silently as the VM account, which has no Qlik access pass. If the account is
# domain-joined and plain "gmrqlik" is rejected, try DOMAIN\gmrqlik or
# gmrqlik@domain.
# Credentials are read from env vars QLIK_USERNAME/QLIK_PASSWORD, else a
# gitignored qlik_credentials.txt (KEY=VALUE). They are never stored in this file.
CREDENTIALS_FILE = os.path.join(SCRIPT_DIR, "qlik_credentials.txt")

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


def ensure_logged_in(page):
    """Open the hub. The Windows-auth challenge is answered at the HTTP layer by
    the context's http_credentials, so this just navigates and reports where we
    landed. Staying on the /internal_windows_authentication/ endpoint means the
    service-account credentials were rejected."""
    target = TENANT_URL.rstrip("/") + "/hub/"
    print(f"Opening Qlik: {target}")
    try:
        page.goto(target, wait_until="domcontentloaded", timeout=60000)
        time.sleep(2)
    except Exception as e:
        print(f"   ! couldn't load the hub ({e}).")
    print(f"   landed at: {page.url}")
    if "authentication" in (page.url or "").lower():
        print("   ! still on the Windows-auth endpoint — the service-account "
              "credentials look rejected. Check QLIK_USERNAME/QLIK_PASSWORD "
              "(a domain account may need DOMAIN\\user or user@domain form).")


def capture(page, employee_name, out_path):
    """Load the filtered single view, wait for the Qlik objects to finish
    rendering, then screenshot it."""
    url = build_single_url(employee_name)
    page.goto(url, wait_until="networkidle")
    selector = READY_SELECTOR or ".qv-object"
    try:
        page.wait_for_selector(selector, timeout=30000)
    except Exception:
        print(f"   ! '{selector}' never appeared for {employee_name}; "
              f"screenshotting whatever is on screen.")

    # Qlik lazy-loads its charts, so the object count climbs as they appear.
    # Wait until it stops growing before the fixed settle, so we don't shoot a
    # half-rendered dashboard.
    prev = -1
    for _ in range(30):  # up to ~15s
        count = page.locator(selector).count()
        if count > 0 and count == prev:
            break
        prev = count
        page.wait_for_timeout(500)
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
    if not (username and password):
        sys.exit(
            "No Qlik credentials set. The server uses Windows authentication, so "
            "the bot must send a service account. Set QLIK_USERNAME/QLIK_PASSWORD "
            "(env vars) or fill qlik_credentials.txt."
        )

    print(f"Processing {len(people)} employee(s). "
          f"REVIEW_MODE={REVIEW_MODE}  "
          f"redirect={'ON -> ' + TEST_REDIRECT_EMAIL if TEST_REDIRECT_EMAIL else 'off'}")
    print(f"Qlik login (Windows auth): {username}")

    print("Starting the browser engine...", flush=True)
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        browser = launch_browser(p)
        # Answer the Qlik proxy's Windows-auth (NTLM/Negotiate) challenge with the
        # service account, rather than the VM's ambient login.
        context = browser.new_context(
            http_credentials={"username": username, "password": password},
            viewport={"width": VIEWPORT_WIDTH, "height": VIEWPORT_HEIGHT},
            device_scale_factor=DEVICE_SCALE,
        )
        page = context.new_page()

        ensure_logged_in(page)

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
