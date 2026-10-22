# Put UK Racing Analysis online so it opens in Chrome

This project is a Streamlit web app. You do **not** need to install Python on your Chromebook once it is deployed: you open the web address in Chrome.

## Easiest route: Streamlit Community Cloud

You will need a GitHub account and a Streamlit Community Cloud account. This package is prepared for deployment, but it is not published yet; deployment must be completed in your accounts.

1. Download and unzip this project.
2. Sign in to GitHub and create a **private repository** (recommended for a personal project).
3. Upload the contents of the `uk_racing_analysis_mvp` folder to the repository. Include `app.py`, `racing_app.py`, `requirements.txt`, `providers/`, `data/`, `tests/`, and `.streamlit/config.toml`. Do **not** upload API passwords, `data/racing.db`, or generated files.
4. Sign in to Streamlit Community Cloud and choose **Create app**.
5. Select your repository, the branch you uploaded, and `app.py` as the main file.
6. Deploy. Streamlit will install the packages listed in `requirements.txt` and give you a web URL.
7. Open that URL in Chrome on your Chromebook. Bookmark it for quick access.
8. In the app's sidebar, enter your own The Racing API username and password when you use the fetch/sync buttons. Do not put API credentials in source code or upload them to GitHub.

## Important: historical database persistence

The app currently uses SQLite at `data/racing.db`. Streamlit Community Cloud's local filesystem is not a durable database service; local data can be lost when the app is restarted, rebuilt, or redeployed. This means a hosted app may not reliably retain the historical results archive over time. Excel/JSON downloads should be downloaded to your device when you need to keep them.

For reliable long-term historical tracking, the next step is to connect the app to persistent hosted storage (for example, a managed database or persistent-volume host). That requires an additional setup and should be done before relying on the hosted site as your permanent history database.

## Privacy and access

- Keep the GitHub repository private.
- Check the hosting platform's current sharing/privacy options before using the app; do not assume an app URL is private just because the source repository is private.
- The API username/password fields are used to make requests and are not intentionally written to the database. Still, only use the app over the hosting provider's HTTPS URL and don't share the app with other people if you want it for personal use.
- The Racing API's plan, endpoint access and usage terms apply. The free tier does not provide a bulk historical backfill, so historical stats begin sparse and grow only as results are synced.

## If deployment fails

- Confirm `app.py` is selected as the main file.
- Check the deployment logs for package installation or Python errors.
- Ensure `requirements.txt` is present at the repository root.
- Do not commit `secrets.toml`, API passwords, the SQLite database, or generated reports.
