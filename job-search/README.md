# Germany entry / junior job search (Oct 2026)

`Germany_Entry_Junior_Jobs.xlsx` has 1,307 full-time entry-level and junior tech jobs in Germany, scored against `Yashik_Khunt_CV.tex`.

Sheets: Summary, All Jobs, Top Matches (60+), English-Friendly, AI-ML-LLM Roles, Near Paderborn - NRW.

Sources: LinkedIn public job search (Entry level + Associate, full-time, last 30 days), Arbeitnow API, Indeed.

Re-run:

```
python3 scripts/scrape.py raw.json
python3 scripts/build.py raw.json scripts/indeed.tsv Germany_Entry_Junior_Jobs.xlsx
```
