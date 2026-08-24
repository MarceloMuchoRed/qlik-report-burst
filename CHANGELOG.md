# Changelog

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
