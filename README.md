# Hawaii Weekly Scraper

Automated weekly scraper for the Hawaii Campaign Finance **2026 Candidate Report**.

**Source:** https://olvr.hawaii.gov/Controls/CandidateFiling.aspx?elid=94

## What it does

- Paginates through all 24 pages (348 candidates) of the 2026 Candidate Report
- Extracts 11 columns per candidate: Contest, Party, Ballot Name, Legal Name, Mailing Address, Phone, Email, Website, Issued, Filed, Status
- Saves a formatted Excel spreadsheet to `hawaii_reports/hawaii_candidates_YYYYMMDD.xlsx`
- Automatically stops running after September 1, 2026

## Schedule

Runs every **Monday at 12:00 UTC** via GitHub Actions.
Can also be triggered manually via the Actions tab → "Run workflow".

## Output

Each run uploads `hawaii_reports/` as a GitHub Actions artifact (retained 90 days).

### Excel structure
- **Sheet 1 – Hawaii 2026 Candidates**: All 348 rows with 11 columns, frozen header, auto-filter
- **Sheet 2 – Summary**: Party breakdown, Status breakdown, Contest breakdown

## Columns

| Column | Description |
|---|---|
| Contest | Office being sought (e.g. U.S. REPRESENTATIVE, DIST I) |
| Party | Political party affiliation |
| Ballot Name | Name as it appears on the ballot |
| Legal Name | Full legal name |
| Mailing Address | Candidate's mailing address |
| Phone | Contact phone number |
| Email | Contact email address |
| Website | Campaign website |
| Issued | Date candidacy papers were issued |
| Filed | Date filed (if applicable) |
| Status | Current filing status (Issued / Filed) |

## Local usage

```bash
pip install -r requirements.txt
playwright install chromium --with-deps
python hawaii_weekly_scraper.py
```
