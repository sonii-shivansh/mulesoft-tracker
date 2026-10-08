# MuleSoft Job Tracker

Automated job tracker for MuleSoft Developer / Integration Engineer roles.

## What it does
- Runs on GitHub Actions every 5 minutes (GitHub may delay scheduled runs under load).
- Also runs on every push and can be started manually from the Actions tab.
- Searches public job-board/search-engine results plus targeted company career pages.
- Filters for MuleSoft / integration / API-development roles and targets roughly 2–4 years.
- Rejects obvious 6+ year, architect, manager, lead, and principal roles.
- Deduplicates by canonical application URL.
- Appends new matches to MuleSoft_Job_Tracker.xlsx.
- Keeps structured source data in data/jobs.json.
- Creates a GitHub Actions job summary when new jobs are found.

## Important limitation
LinkedIn, Indeed, Glassdoor and some corporate ATS sites restrict automated access. This tracker does not bypass login, anti-bot controls, CAPTCHAs, or private APIs. It uses publicly accessible search results/pages where available.

## Profile matching
Mule 4, Anypoint Platform/Studio/API Manager/Runtime Manager/CloudHub, RAML, APIKit, DataWeave 2.0, API-led connectivity, Salesforce, SQL/MySQL, Anypoint MQ, SFTP, SMTP, HTTP, SOAP, Batch, Object Store, OAuth 2.0, JWT, Client ID Enforcement, MUnit and Postman.

## Files
- tracker.py — collector, matcher, deduplicator and Excel writer
- data/jobs.json — persistent job database
- MuleSoft_Job_Tracker.xlsx — generated master Excel workbook
- .github/workflows/mulesoft-tracker.yml — scheduled workflow

The workbook columns are:
Date & Time Found | Company Name | Job Title | Location / Remote Status | Match Score / Relevance Note | Direct Application Link
