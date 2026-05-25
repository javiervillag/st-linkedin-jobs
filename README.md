# LinkedIn ST Job Discovery

Daily automated scraper that finds companies using ServiceTitan by analyzing LinkedIn job postings.

## How it works

1. Scrapes LinkedIn Ad Library for jobs mentioning "servicetitan"
2. Filters out blacklisted companies (staffing agencies, ServiceTitan itself, etc.)
3. Opens each job detail page to read the full description
4. Uses DeepSeek AI to classify whether the company is a confirmed ServiceTitan user
5. Generates a daily XLSX report with confirmed users and blacklisted companies

## Setup

```bash
npm install
npx playwright install chromium
```

Set environment variable:
```bash
export DEEPSEEK_API_KEY=sk-...
```

## Run

```bash
node scripts/scrape-daily.js
```

## Railway Deployment

Configured for daily cron via `railway.json`. Just set `DEEPSEEK_API_KEY` env var.

## Output

- `data/linkedin_companies.json` — persistent database of confirmed ST users
- `data/linkedin_blacklist.json` — auto-growing blacklist
- `data/linkedin_report_YYYY-MM-DD.xlsx` — shareable Excel report
