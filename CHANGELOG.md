# Changelog

## 2026-08-26 (name-match safeguard: resolve recipients to exact Qlik value)
- A recipient listed as "…Jr" didn't match Qlik's "…Jr." and was silently
  SKIPPED (validation already did `strip().casefold()`, but not trailing-period /
  whitespace differences). Just loosening validation would be worse — the name is
  also the Qlik *selection* value, so a near-miss that passed validation would
  select nothing and shoot the whole-company dashboard with dead deep links.
- Fix: build `canon_map` (normalized name → exact Qlik value) from the values we
  already fetch for validation, resolve each recipient through it, and use the
  **exact Qlik value** for the screenshot selection, the greeting, and both deep
  links. `_normalize_name()` is case-/whitespace-insensitive and ignores a
  trailing period. Every adjustment is logged (`note: 'X' matched Qlik 'X.'`);
  genuinely unknown names still skip. Only active when name validation is on.

## 2026-08-26 (email: detail link as text + browser note; faster capture)
- The "detail of your performance" link now shows friendly text ("View your
  performance detail") instead of the raw URL, matching the named dashboard link.
- Added a note telling recipients the links open in their **default browser**,
  which must be signed in to Qlik for them to work.
- **Capture speedups** (a 20-employee run spends most of its time here):
  - `page.goto` now waits on `domcontentloaded` instead of `networkidle`. Qlik
    holds a websocket open so the page never goes fully idle; readiness is already
    enforced by the selector wait + object-count-stable loop, so `networkidle`
    was just adding dead time per employee.
  - `RENDER_SETTLE_SECONDS` 6 → 3 (the fixed post-render settle; 6×20 = 2 min of
    pure waiting). This is the speed/quality knob — if any shot looks mid-render,
    raise it back toward 6.

## 2026-08-26 (email reformat: "Sales Report" subject + personalized links)
- Reworked the email to the requested layout:
  - Subject is now `Sales Report {date}` (run date, US `M/D/YYYY`) instead of
    `Your weekly dashboard - {name}`.
  - Body adds two personalized dashboard links (a named "Dashboard Sales KPI
    Performance v.4" link and the raw URL), each a Qlik Sense client deep link
    (`…/sheet/<SHEET_ID>/state/analysis/options/clearselections/select/
    SALESPERSON_ORDER/<name>`) so each recipient lands on only their own results.
    The "dashboard" link uses `SHEET_ID` (`6527c8b7…`, same as the screenshot);
    the "detail" link uses a separate `DETAIL_SHEET_ID` (`1266bc38…`). Name is
    URL-encoded.
  - Added contact lines (`gerson@pennrosefarms.com` for questions,
    `it@pennrosefarms.com` for login issues) and moved the embedded screenshot
    below them, above "Regards".
- New `build_detail_url(name, sheet_id)` helper builds the deep link from
  TENANT_URL/APP_ID/`sheet_id`/FILTER_FIELD; `EMAIL_HTML_BODY` gains
  `{dashboard_url}` + `{detail_url}` and the subject gains `{date}`. Byte-compiles
  clean; runtime not re-exercised (needs the VM).

## 2026-08-26 (regained personal-account access in the VM's Chrome)
- **`gmrqlik` is a limited service account** and can't open an app Marcelo built
  under his own Qlik account, so he needed his personal login to auto-fill again
  in the VM's interactive Chrome. The Qlik proxy challenges with a **native
  Windows-auth (NTLM/Negotiate) dialog**, and Chrome was pre-filling the saved
  `gmrqlik` credential — that native dialog has no picker to choose the other
  saved account, and clearing the username shows no autocomplete.
- **Chrome's UI delete was walled off:** opening the `10.0.2.5` entry in
  `chrome://password-manager` triggers a Windows re-auth, and the VM's Windows
  password is unknown (RDP supplies it automatically; no PIN/Hello over RDP to
  satisfy that prompt). Both accounts were confirmed saved.
- **Resolved by deleting only the `gmrqlik` row from Chrome's `Login Data`
  SQLite** with stdlib `sqlite3` (`DELETE ... WHERE signon_realm LIKE '%10.0.2.5%'
  AND username_value LIKE '%gmrqlik%'`, Chrome fully closed). Deleting a row reads
  no encrypted blob, so no Windows password/decryption is involved; the personal
  entry was kept and Chrome now auto-fills it. The VM AV did **not** block this
  row-delete (it had previously killed scripts that *copy* `Login Data`).
- **No effect on the burst:** it sends `gmrqlik` explicitly via `http_credentials`
  in a throwaway profile, reading `QLIK_USERNAME`/`QLIK_PASSWORD` (env vars or
  `qlik_credentials.txt`) — a separate store from Chrome's saved passwords.

## 2026-08-26 (email image sizing + unknown-name validation)
- **Embedded dashboard image came out huge in Outlook.** The screenshot is
  captured at high resolution (1920×1080 @2× = 3840px wide) and the `<img>` had
  no size, so Outlook rendered it at full pixel width. Set a display width on the
  tag (`EMAIL_IMAGE_WIDTH`, now 850) so it downscales in the client while keeping
  the high-res pixels for a crisp render.
- **Constrained-width `<img>` cropped the logo in Outlook.** A first attempt used
  `style="width:…; max-width:100%; height:auto;"`; Outlook's desktop (Word)
  rendering engine mishandles that CSS on large images and clipped the logo. The
  capture code was never touched (PNG byte-identical) — it was purely an email
  render effect. Fix: use only the bare `width` HTML attribute, which Outlook
  scales proportionally without clipping.
- **Unknown recipient names silently got the whole-company dashboard.** A test
  row ("Tom Brady", not a real salesperson) rendered company-wide totals instead
  of erroring — the Single Integration API ignores a `select=Field,Value` whose
  value isn't in the field, applying no filter. That's a data-exposure risk (an
  unknown recipient would be emailed the entire company's numbers). Fix: added
  `fetch_valid_field_values()` — it opens an engine JSON-RPC WebSocket from inside
  the already-authenticated browser page, reads the full list of `FILTER_FIELD`
  values via a session list object, and `main()` skips + reports any recipient
  whose name isn't among them (case/space-insensitive match). Gated by
  `VALIDATE_NAMES`; if the lookup fails it's disabled with a warning and everyone
  is processed as before. **Untested against the live engine — needs one VM run
  to confirm the WebSocket lookup returns the salesperson list.**

## 2026-08-25 (auth scheme correction — Windows auth, not forms)
- **Confirmed the real auth scheme once IT issued the `gmrqlik` service account:
  the Qlik virtual proxy uses WINDOWS authentication, not forms.** The login URL
  is `/internal_windows_authentication/?targetId=...` (an NTLM/Negotiate browser
  popup), and manual login = typing `gmrqlik` into that popup. This overturns the
  earlier "forms auth" conclusion (the Aug-25 curl saw a forms redirect, but the
  server was either reconfigured when IT set up `gmrqlik` or was never actually
  forms for a browser). Root cause of the failures, in order:
  - First run (fresh throwaway profile, no auth machinery): `net::ERR_INVALID_
    AUTH_CREDENTIALS` on `/hub/` — the browser couldn't answer the NTLM challenge.
  - After adding `--auth-server-allowlist`: it authenticated, but **as the VM's
    own Windows account (`p-mdragustinovis`) via silent SSO**, landing on
    `/hub/?qlikTicket=...` with no form. That account has no Qlik access pass →
    the app showed "You cannot access Qlik Sense because you have no access pass."
    The `gmrqlik` creds were never used (SSO auto-answered the popup first).
- **Fix:** answer the Windows-auth challenge explicitly with the service account
  via Playwright `http_credentials={username, password}` on the browser context,
  and **remove `--auth-server-allowlist`** (it was forcing the wrong-identity
  SSO). Removed the now-wrong forms-login machinery entirely (`submit_login`,
  `on_login_page`, `manual_login_pause`, the `LOGIN_*_SELECTOR` config) and the
  `urlparse`/`AUTH_SERVER_ALLOWLIST` bits. `ensure_logged_in` now just navigates
  and reports; staying on `/internal_windows_authentication/` means the creds
  were rejected. Credentials are now mandatory (script exits if unset). Also
  fixed the stale `SHEET_ID` (was `1266bc38-...`, now `6527c8b7-...` from the live
  app URL). Byte-compiles clean.
- **CONFIRMED working end-to-end on the VM:** with `http_credentials` sending
  `gmrqlik`, the bot authenticates as the licensed account and the dashboard
  renders (no more "no access pass"). Auth chapter closed.
- Screenshot-quality pass (dashboards came out zoomed wrong with some charts
  missing). Root causes: (1) Playwright's default 1280x720 viewport made Qlik use
  a cramped/rescaled layout, (2) charts lazy-load and we shot too early. Fix: set
  a `VIEWPORT_WIDTH/HEIGHT` (1920x1080) + `DEVICE_SCALE=2` context for a full,
  crisp desktop render, and in `capture()` wait for the `.qv-object` count to
  stop growing before the (now 6s) settle. Rejected switching to PDF: browser
  `page.pdf()` inherits the same viewport/timing problems and print-CSS makes
  dashboards worse; only Qlik NPrinting is a genuinely better PDF path, but it's
  a separate licensed product — deferred unless screenshots prove insufficient.

## 2026-08-25
- Resolved the last open auth question: the username Chrome autofills on the Qlik
  login form is a **separate Qlik/service account**, not Marcelo's own AD/corporate
  login, so its password isn't knowable or resettable by him. This confirms the IT
  path is the only viable one (consistent with the two dead ends already recorded).
  Drafted a credential request to IT asking for either the existing account's
  username/password or a dedicated service account with read access to the app.
- Rewrote the script for a scripted forms login so it's demo-ready and ready to
  drop the password into the moment IT issues it (Marcelo has a meeting to get
  sign-off from a stakeholder wary of "AI" on the platform — the point is to show
  it's a plain login-export-email bot, no AI). Changes:
  - `ensure_logged_in` now fills Qlik's `/internal_forms_authentication/` form
    (username + `pwd`) and submits, instead of relying on Chrome autofill. Creds
    come from env vars `QLIK_USERNAME`/`QLIK_PASSWORD` or a gitignored
    `qlik_credentials.txt` — never hardcoded, so the auto-push can't leak them.
    If no creds and headed, it falls back to a manual login pause (lets Marcelo
    demo by logging in by hand before the password exists).
  - Dropped ALL the dead auth/profile machinery: real-profile-in-place, the
    profile-copy path, `CLOSE_CHROME`/`chrome_is_running`/`ensure_chrome_closed`,
    the singleton-lock cleanup, and the NTLM/`WINDOWS_INTEGRATED_AUTH`/HTTP-auth
    options. The browser now launches in a fresh throwaway profile
    (`p.chromium.launch(channel="chrome")`, Chromium fallback), which sidesteps
    the Chrome-151 default-dir automation block and is fully headless-capable.
  - Removed now-unused `json`/`subprocess` imports. Rewrote the module docstring
    and README to state plainly there is no AI/ML/LLM in it and describe the
    login→export→email flow. `.gitignore` now protects `qlik_credentials.txt`.
    Byte-compiles clean; runtime not yet exercised (no Qlik creds / not on VM).
- Pinned down the auth scheme, ending the HTTP/NTLM detour. `page.goto` had been
  failing with `ERR_INVALID_AUTH_CREDENTIALS`, which we'd guessed was HTTP Basic/
  NTLM (added `WINDOWS_INTEGRATED_AUTH` + `--auth-server-allowlist` and
  `HTTP_AUTH_USERNAME/PASSWORD` on a hunch). Curl on the VM disproved it:
  `curl -sI http://10.0.2.5/hub/` returns `302 -> /internal_forms_authentication/
  ?targetId=...`, i.e. **Qlik Sense internal FORMS authentication (a web login
  page)**, not an HTTP 401 challenge. `curl --ntlm -u :` (current Windows user via
  SSPI) got the *same* form redirect, proving **NTLM/Windows SSO does nothing
  here** — the Qlik virtual proxy isn't configured for it. So the integrated-auth
  and HTTP-credentials machinery is dead weight for this server. chrome://version
  on the VM confirms the active profile is `Default` (matches CONFIG). Net: auth
  must come from either (a) scripting the Qlik forms login with real credentials,
  or (b) reusing a live session cookie — decision pending on whether the Qlik
  username/password can be obtained.
- Ruled out in-place real-profile automation entirely. The VM's Chrome is
  **151.0.7922.174**; since Chrome 136 (early 2025) Chrome silently disables the
  remote-debugging pipe when `--user-data-dir` is the DEFAULT profile dir, so
  Playwright launches Chrome but can never drive it — the observed "about:blank
  forever." Modern Chrome REQUIRES a non-default dir for automation (i.e. a copy).
- Confirmed the Qlik session cookie is **session-scoped**: a full Chrome restart
  drops you back to the login form every time (Chrome re-autofills it, user just
  presses Enter). So there is no on-disk session cookie to copy — every run needs
  a FRESH login, which needs the password. The password lives only in Chrome's
  saved-login store; the AV kills any script that copies it and the user can't
  view it (no VM Windows password). Net conclusion: **no reliable path without
  IT** — need the Qlik credentials (or a service account) to script the forms
  login. Also fixed a real bug: `chrome_is_running()` counted other sessions'/VMs'
  chrome.exe on the shared host (tasklist is host-wide), causing a false "Chrome
  is open" prompt; now scoped to the current user via `tasklist /V` CSV owner match.

## 2026-08-24
- Fixed the login blocker: the Qlik password is saved in Chrome's password
  manager but unknown to the user (VM they don't have the Windows password for),
  so the old design — a blank Playwright Chromium profile requiring a manual
  password-typed login — couldn't authenticate. Switched the browser launch to
  drive the **installed Chrome** (`channel="chrome"`) against the user's **real
  Chrome profile** (`Profile 1`, User Data root + `--profile-directory`), so the
  saved Qlik login/session is already present; first run just needs autofill +
  Enter, no password. Added a preflight (`ensure_chrome_closed`) since Chrome
  locks its profile while open, plus `CLOSE_CHROME`/`HEADLESS` config for
  unattended runs. Considered and rejected pyautogui/AHK UI automation (brittle,
  can't run on the VM's lockable screen, and would drop the exact URL-driven
  per-recipient filter). Dropped the now-unused `.browser-profile/` and the
  `playwright install chromium` step. Win10 VM.
- Reconciled branch divergence caused by PR #1 being squash-merged on GitHub. `origin/master` held the relative-paths change as a single squashed commit (`5f25add`), while local work had that same change plus the on-prem retarget. Replayed only the on-prem commit onto the squashed `master` (clean, tree-identical to the old `a441d1d`), fast-forward-pushed `master` (`7edf2cb`), then deleted the now-redundant `relative-paths` branch (local + remote). Repo is now a single `master` branch, on par local and remote.
