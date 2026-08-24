# Changelog

## 2026-08-24
- Reconciled branch divergence caused by PR #1 being squash-merged on GitHub. `origin/master` held the relative-paths change as a single squashed commit (`5f25add`), while local work had that same change plus the on-prem retarget. Replayed only the on-prem commit onto the squashed `master` (clean, tree-identical to the old `a441d1d`), fast-forward-pushed `master` (`7edf2cb`), then deleted the now-redundant `relative-paths` branch (local + remote). Repo is now a single `master` branch, on par local and remote.
