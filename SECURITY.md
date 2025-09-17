# Security Policy

- Reporting: Please report vulnerabilities privately via email to the repository owner. Do not open public issues for security reports.
- Scope: This project ships a client-only SPA and a Google Apps Script. There are no backend secrets or tokens in the repository.
- Data handling: The SPA processes data locally in the browser and does not transmit files or contents over the network. If future cloud features are added, they must be opt-in and documented.
- Dependencies: Builds are pinned via `package-lock.json` and GitHub Actions use `npm ci` for reproducible installs.

