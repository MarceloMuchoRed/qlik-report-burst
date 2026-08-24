"""
qlik_report_burst.py
---------------------
Weekly "burst" job: for each employee in a spreadsheet, open the Qlik Sense
Cloud dashboard filtered to that employee, screenshot it, and email the image
to that person via the local (already-signed-in) Outlook desktop app.

Why this design:
  * The Qlik filter is applied through the Single Integration API URL
    (...&select=Field,Value) instead of clicking around the filter pane, so it
    doesn't break when the UI changes.
  * Login is NOT scripted and no password is needed. The script drives your
    INSTALLED Chrome using your real Chrome profile, so your saved Qlik login
    is already there — Chrome autofills it and you just press ENTER once on the
    first run. Chrome must be CLOSED while the script runs (it locks the
    profile). The script never needs (or stores) your Qlik password.
  * Email goes through your desktop Outlook via COM automation, so there are no
    SMTP servers, app passwords, or OAuth apps to set up.

No admin rights required. Install into your portable Python with:
    python -m pip install playwright pywin32 openpyxl
    python -m playwright install chromium

Then edit the CONFIG block below and run:
    python qlik_report_burst.py
"""

import csv
import os
import subprocess
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

# This script drives your INSTALLED Chrome with your real Chrome profile, so
# your saved Qlik login (and any live session) is already present — you never
# type or need to know the password. On the first run the Qlik login page
# loads, Chrome autofills your saved username/password, and you press ENTER
# once; the session then persists for future runs.
BROWSER_CHANNEL = "chrome"  # use installed Chrome, not Playwright's own Chromium
# Your Chrome "User Data" root (auto-detected from %LOCALAPPDATA%). Override
# only if Chrome is installed somewhere non-standard.
CHROME_USER_DATA_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "User Data"
)
# The profile folder INSIDE User Data that holds your Qlik login. On this
# machine it's "Profile 1" (not "Default"). To confirm: open chrome://version
# and read "Profile Path" — the last path segment is this value.
CHROME_PROFILE = "Profile 1"
# Chrome locks a profile while it's open, so Chrome must be CLOSED during a run.
#   False -> the script pauses and asks you to close Chrome yourself (friendlier
#            for manual runs; you keep your tabs).
#   True  -> the script force-closes Chrome for you (needed for unattended /
#            scheduled runs; you lose any open tabs).
CLOSE_CHROME = False
# headless=False lets you see the one-time login and lets Chrome autofill work.
# Once your session persists, you can set this True for silent scheduled runs.
HEADLESS = False
# Seconds to wait after the chart appears, for animations/data to settle.
RENDER_SETTLE_SECONDS = 4
# Optional: a CSS selector that only appears once the chart is fully drawn.
# Leave as "" to use the generic Qlik object container.
READY_SELECTOR = ""

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


def looks_like_login(page):
    """Heuristic: are we sitting on a login / SSO page rather than the app?"""
    url = page.url.lower()
    if "single" in url and "qlikcloud" in url:
        return False
    login_markers = ["login", "signin", "sign-in", "auth", "oauth", "sso", "idp"]
    return any(m in url for m in login_markers)


def chrome_is_running():
    """True if any chrome.exe process is running (it locks the user profile)."""
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq chrome.exe", "/NH"],
            capture_output=True, text=True,
        ).stdout.lower()
        return "chrome.exe" in out
    except Exception:
        return False


def ensure_chrome_closed():
    """Chrome locks its profile while open, so make sure it's closed before we
    launch Playwright against CHROME_PROFILE."""
    if not chrome_is_running():
        return
    if CLOSE_CHROME:
        print("Chrome is open; closing it so its profile can be used...")
        subprocess.run(["taskkill", "/IM", "chrome.exe", "/F"],
                       capture_output=True, text=True)
        time.sleep(2)
        return
    print("\n" + "=" * 70)
    print(" Chrome is currently OPEN. It locks the profile this script needs.")
    print(" Please CLOSE all Chrome windows now (check the system tray too),")
    input(" then come back here and press ENTER to continue... ")
    print("=" * 70 + "\n")
    if chrome_is_running():
        print(" ! Chrome still looks like it's running. If the next step fails,"
              " close Chrome fully and re-run (or set CLOSE_CHROME = True).")


def ensure_logged_in(page):
    """Navigate to the tenant and, if needed, pause for a one-time manual login."""
    page.goto(TENANT_URL, wait_until="domcontentloaded")
    time.sleep(3)
    if looks_like_login(page) or "hub" not in page.url.lower():
        print("\n" + "=" * 70)
        print(" A browser window is open. Please LOG IN to Qlik there now.")
        print(" (You only have to do this once; the session is remembered.)")
        print(" When you can see your Qlik hub/app, come back here and")
        input(" press ENTER to continue... ")
        print("=" * 70 + "\n")


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


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    profile_path = os.path.join(CHROME_USER_DATA_DIR, CHROME_PROFILE)
    if not os.path.isdir(profile_path):
        sys.exit(
            f"Chrome profile not found: {profile_path}\n"
            f"Fix CHROME_USER_DATA_DIR / CHROME_PROFILE in CONFIG. "
            f"(Open chrome://version and read 'Profile Path'.)"
        )

    people = load_recipients(RECIPIENTS_FILE)
    if MAX_EMPLOYEES is not None:
        people = people[:MAX_EMPLOYEES]

    print(f"Processing {len(people)} employee(s). "
          f"REVIEW_MODE={REVIEW_MODE}  "
          f"redirect={'ON -> ' + TEST_REDIRECT_EMAIL if TEST_REDIRECT_EMAIL else 'off'}")

    ensure_chrome_closed()

    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        # Drive your INSTALLED Chrome (channel="chrome") using your real profile
        # so the saved Qlik login / live session is already present. Passing the
        # User Data root plus --profile-directory selects CHROME_PROFILE.
        ctx = p.chromium.launch_persistent_context(
            CHROME_USER_DATA_DIR,
            channel=BROWSER_CHANNEL,
            headless=HEADLESS,
            accept_downloads=False,
            args=[f"--profile-directory={CHROME_PROFILE}"],
        )
        # Use a fresh tab (don't grab a session-restored one from the profile).
        page = ctx.new_page()

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

        ctx.close()

    print("\nDone. Review the drafts in Outlook (REVIEW_MODE) or check Sent.")


if __name__ == "__main__":
    main()
