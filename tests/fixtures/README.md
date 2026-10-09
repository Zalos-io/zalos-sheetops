# Test Fixtures

This folder stores real spreadsheet and CSV fixtures used to prove SheetOps
behavior. Generated fixtures are created inside tests. Upstream fixtures are
copied only when they cover behavior that is hard to represent with tiny
synthetic files.

Every upstream fixture must have an entry in `manifest.json` with source repo,
commit, original path, license, destination, SHA256, and behavior.
