# MuleSoft Job Tracker

Automated job tracker for MuleSoft Developer / Integration Engineer roles.

## What it does

- Runs hourly through GitHub Actions.
- Filters toward approximately 2–4 year MuleSoft/integration roles.
- Rejects explicit 6+ year requirements and senior architect/principal/director roles.
- Deduplicates postings.
- Maintains `jobs.json` as the source of truth.
- Rebuilds `MuleSoft_Job_Tracker.xlsx` on every run.
- Uploads the Excel workbook as a workflow artifact.
- Commits tracker changes back to the repository.
- Creates a GitHub Issue when new jobs are discovered.

## Discovery providers

The tracker supports Greenhouse public job-board APIs, Lever public postings APIs, and configured RSS feeds. LinkedIn/Indeed/Glassdoor are not blindly scraped because a GitHub runner cannot reliably bypass their anti-bot/access controls; provider/API configuration should be used where permitted.

## Configuration

Repository secrets accept comma-separated entries.

Greenhouse: `board_token|Company Name`

Lever: `company_slug|Company Name`

RSS: `feed_url|Company Name`

## Limitation

GitHub Actions scheduled workflows run approximately hourly and may be delayed by GitHub. This is not guaranteed real-time monitoring.
