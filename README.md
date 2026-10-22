# UK Racing Analysis — browser-hosted prototype

This is a Streamlit web app that can be deployed online and opened in Chrome on a Chromebook. See [DEPLOY_TO_CHROMEBOOK.md](DEPLOY_TO_CHROMEBOOK.md) for deployment steps. This package is prepared for deployment but is not yet a live website; deployment must be completed through your own hosting account.

A browser-hosted prototype for personal UK horse-racing research. It fetches basic racecards from The Racing API Free endpoint, stores data locally, calculates transparent race-relative indicators, and automatically exports Excel workbooks plus website-ready JSON.

**This is an experimental screening tool, not a validated betting model.** It does not produce win probabilities or betting advice. The app never invents missing speed/performance ratings. The included demo horses are fictional.

## Quick start

1. Create an account at https://www.theracingapi.com/ and obtain your API username/password from the account dashboard.
2. Download and extract the application ZIP.
3. On Windows, double-click `Start_Windows.bat`. On macOS, right-click `Start_Mac.command` and choose Open if prompted.
4. In the browser app's sidebar, enter your API username/password, choose Today or Tomorrow, and click **Fetch racecards + calculate all reports**.
5. The app fetches the basic racecards, imports them into its local SQLite database, calculates available indicators, and creates Excel and JSON files for each returned meeting. Download the files from the Generated Files section.

The normal workflow does not require spreadsheet imports. To build historical context, click **Sync today's results into historical database** after results are published. Repeating this daily builds a local results archive over time. An optional manual CSV importer remains available for testing and alternate sources.

## What is calculated

- Percentile ranks for sourced speed figures, sourced performance ratings, official ratings, and a simple recent-form proxy, where data is available.
- A within-race screening index using these factor weights: speed figure 40%, performance rating 35%, official rating 15%, recent-form proxy 10%. Weights are renormalized over available factors.
- Rating coverage and overall data-completeness indicators.
- A separate, clearly labelled recency heuristic based on days since last run.
- Relative carried-weight context, displayed but not included in the score.
- Pedigree, draw, headgear, age, sex, trainer and jockey fields when returned by the API.
- Flags for long breaks, quick turnarounds and missing core data.

A provisional index can appear at 20% factor-weight coverage so the free tier can be explored, but low-coverage scores are explicitly labelled **indicative only**. Do not compare scores across different races as if they were absolute ability ratings.

## Important limitations

- The Free racecard endpoint is basic. The API may return no speed or performance figures; the app uses those fields only if present in the response.
- The recency heuristic is not empirically validated. Weight context and pedigree are displayed but are not treated as proven advantages.
- The app can sync today's basic results from The Racing API Free endpoint into a persistent local historical archive. Use **Sync today's results into historical database** after results are published; repeated syncs update existing records rather than duplicating them.
- Historical context now includes the locally stored horse run/win/top-3 record, course record, distance-similar record, going-family record, local trainer/jockey strike rates, course race counts, and counts of same-named race editions. These are calculated only from prior results stored locally and show their sample sizes.
- The free API does not provide a bulk historical backfill: its full historical results and horse-results endpoints are plan-restricted. Therefore, the local archive starts sparse and grows from the dates you sync; it is not a complete historical course database.
- No robust speed figures are created from race times in this version. Reliable speed ratings need actual race-time data, distance/course/going adjustments, and calibration. Pace, jumping, stamina and breeding outcomes also need richer historical data.
- Public racing websites may show historic results, but public visibility does not automatically grant permission to bulk scrape or reuse their databases. This app does not scrape paywalled/proprietary ratings or bulk-copy sites whose terms do not permit it.
- The app has been syntax-checked and tested against local fictional sample data, but a live API login has not been tested with your account.

## Data and privacy

- API requests use HTTP Basic Authentication. Credentials are entered in the local UI and are not intentionally saved to the database or source files.
- The app stores racecard records locally in `data/racing.db` and writes generated reports to `output/`.
- The free API is limited to one request per second; the app makes a single racecard request per fetch.
- Only store, analyse or redistribute data in ways allowed by the provider's current terms. API access does not automatically grant permission to republish provider data or derived outputs on a public website.

## Technical details

- Python 3.10+
- Streamlit UI
- SQLite local database
- Excel reports with Meeting Summary, Runner Database, Historical Context, Source Audit and Scoring Method sheets
- JSON exports for a future website
- CSV schema in `data/racecard_template.csv`

For troubleshooting, check that Python is installed, your credentials are correct, and the app has internet access. Do not share your API password when asking for help.

## Hosting note

The default SQLite database is stored on the host's filesystem. On free ephemeral hosting, historical data may be lost after restarts or redeployments. For dependable long-term history, connect persistent hosted storage before relying on the site as your permanent archive.
