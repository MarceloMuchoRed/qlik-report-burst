# Changelog

## 2026-08-25
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
