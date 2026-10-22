from __future__ import annotations
import csv
import io
import sys
from datetime import datetime
from pathlib import Path

import streamlit as st
import requests
from requests.auth import HTTPBasicAuth

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
import racing_app as core

st.set_page_config(page_title="UK Racing Analysis", page_icon="🏇", layout="wide")

st.title("🏇 UK Racing Analysis")
st.caption("One-click racecard retrieval, transparent race-relative analysis, Excel workbooks and website-ready JSON.")

st.info(
    "Fetch today's or tomorrow's racecards, sync today's available results into a local historical archive, and calculate available runner/course context. "
    "Missing speed/performance ratings are not invented; historical statistics are labelled with their stored sample size."
)

API_BASE = "https://api.theracingapi.com/v1/racecards/free"
API_DOCS = "https://api.theracingapi.com/documentation"
RESULTS_API_BASE = "https://api.theracingapi.com/v1/results/today/free"


def _first(item, *keys, default=""):
    for key in keys:
        value = item.get(key)
        if value not in (None, ""):
            return value
    return default


def fetch_free_racecards(username: str, password: str, day: str, include_ireland: bool = False) -> tuple[int, int, list[str]]:
    """Fetch basic racecards from The Racing API Free endpoint and import them locally."""
    region_codes = "gb,ire" if include_ireland else "gb"
    response = requests.get(
        API_BASE,
        params={"day": day, "region_codes": region_codes, "limit": 500, "skip": 0},
        auth=HTTPBasicAuth(username, password),
        timeout=45,
    )
    if response.status_code == 401:
        raise ValueError("Authentication failed or your account does not have access. Check the API username/password in My Account.")
    if response.status_code == 403:
        raise ValueError("Access was denied. Check your plan and account status in The Racing API dashboard.")
    response.raise_for_status()
    payload = response.json()
    races = payload.get("racecards", []) if isinstance(payload, dict) else payload
    if not isinstance(races, list):
        raise ValueError("The API response did not contain a racecards list. Check the provider documentation.")

    stamp = datetime.now().astimezone().isoformat(timespec="seconds")
    output = io.StringIO(newline="")
    writer = csv.DictWriter(output, fieldnames=core.CSV_FIELDS)
    writer.writeheader()
    runner_count = 0
    race_count = 0
    fetched_dates = set()
    for race in races:
        if not isinstance(race, dict):
            continue
        date_value = str(_first(race, "date", default=datetime.now().date().isoformat()))[:10]
        course = str(_first(race, "course", default="Unknown course"))
        race_time = str(_first(race, "off_time", "off", default="Unknown time"))
        race_name = str(_first(race, "race_name", default="Unnamed race"))
        fetched_dates.add(date_value)
        runners = race.get("runners", []) or []
        if not isinstance(runners, list):
            continue
        race_count += 1
        for runner in runners:
            if not isinstance(runner, dict):
                continue
            status = str(_first(runner, "status", "runner_status", default="")).lower()
            non_runner = status in {"non-runner", "non runner", "nr", "withdrawn"} or bool(runner.get("non_runner", False))
            writer.writerow({
                "date": date_value,
                "meeting": course,
                "course": course,
                "country": str(_first(race, "region", default="GB")),
                "race_time": race_time,
                "race_name": race_name,
                "race_type": _first(race, "type", "race_type", default=""),
                "distance_f": _first(race, "distance_f", "dist_f", default=""),
                "going": _first(race, "going_detailed", "going", default=""),
                "cloth": _first(runner, "number", "cloth", default=""),
                "runner_name": _first(runner, "horse", "runner_name", default="Unknown runner"),
                "age": _first(runner, "age", default=""),
                "weight_lbs": _first(runner, "lbs", "weight_lbs", default=""),
                "official_rating": _first(runner, "ofr", "official_rating", default=""),
                "speed_figure": _first(runner, "speed_rating", "speed_figure", "ts", default=""),
                "performance_rating": _first(runner, "performance_rating", "rpr", default=""),
                "recent_form": _first(runner, "form", default=""),
                "trainer": _first(runner, "trainer", default=""),
                "jockey": _first(runner, "jockey", default=""),
                "non_runner": "yes" if non_runner else "no",
                "source_name": "The Racing API Free",
                "source_url": API_DOCS,
                "source_updated_at": stamp,
                "race_class": _first(race, "race_class", "class", default=""),
                "field_size": _first(race, "field_size", default=len(runners)),
                "surface": _first(race, "surface", default=""),
                "prize": _first(race, "prize", default=""),
                "rating_band": _first(race, "rating_band", default=""),
                "age_band": _first(race, "age_band", default=""),
                "sex_restriction": _first(race, "sex_restriction", default=""),
                "race_id": _first(race, "race_id", default=""),
                "course_id": _first(race, "course_id", default=""),
                "horse_id": _first(runner, "horse_id", default=""),
                "sex": _first(runner, "sex", default=""),
                "colour": _first(runner, "colour", default=""),
                "sire": _first(runner, "sire", default=""),
                "dam": _first(runner, "dam", default=""),
                "damsire": _first(runner, "damsire", default=""),
                "sire_id": _first(runner, "sire_id", default=""),
                "dam_id": _first(runner, "dam_id", default=""),
                "damsire_id": _first(runner, "damsire_id", default=""),
                "draw": _first(runner, "draw", default=""),
                "last_run_days": _first(runner, "last_run", "last_run_days", default=""),
                "headgear": _first(runner, "headgear", default=""),
                "owner": _first(runner, "owner", default=""),
                "jockey_id": _first(runner, "jockey_id", default=""),
                "trainer_id": _first(runner, "trainer_id", default=""),
            })
            runner_count += 1

    if runner_count == 0:
        raise ValueError("The API returned no runners for that selection. Try the other day or check the provider dashboard.")
    import_path = ROOT / "data" / f"racing_api_{day}_racecards.csv"
    import_path.parent.mkdir(parents=True, exist_ok=True)
    import_path.write_text(output.getvalue(), encoding="utf-8-sig")
    core.import_csv(import_path)
    return race_count, runner_count, sorted(fetched_dates)

def sync_free_results(username: str, password: str, include_ireland: bool = False) -> tuple[int, int]:
    """Fetch today's basic results and append/update the local historical archive."""
    region_codes = "gb,ire" if include_ireland else "gb"
    response = requests.get(
        RESULTS_API_BASE,
        params={"region": region_codes, "limit": 500, "skip": 0},
        auth=HTTPBasicAuth(username, password),
        timeout=45,
    )
    if response.status_code == 401:
        raise ValueError("Authentication failed or your account does not have access to today's free results.")
    if response.status_code == 403:
        raise ValueError("Your account does not have access to today's free results endpoint. Check your plan.")
    response.raise_for_status()
    return core.store_historical_results(response.json(), RESULTS_API_BASE)

def generate_reports_for_dates(dates: list[str]) -> tuple[int, list[str]]:
    """Automatically calculate all reports for the dates just fetched."""
    core.init_db()
    generated = []
    meeting_count = 0
    with core.connect() as con:
        for date in dates:
            meetings = con.execute(
                "SELECT meeting_name FROM meetings WHERE meeting_date=? ORDER BY meeting_name", (date,)
            ).fetchall()
            for row in meetings:
                meeting_name = row["meeting_name"]
                core.build_report(date, meeting_name)
                core.export_json(date, meeting_name)
                safe_meeting = "".join(c if c.isalnum() or c in "-_" else "_" for c in meeting_name.replace(" ", "_"))
                generated.extend([
                    str(core.OUTPUT_DIR / f"{date}_{safe_meeting}_analysis.xlsx"),
                    str(core.OUTPUT_DIR / f"{date}_{safe_meeting}_website.json"),
                ])
                meeting_count += 1
    return meeting_count, generated

with st.sidebar:
    st.header("1. Fetch and analyse")
    st.caption("Requires a free account at The Racing API. Credentials are used for this request and are not saved by this app.")
    api_username = st.text_input("API username", key="racing_api_username")
    api_password = st.text_input("API password", type="password", key="racing_api_password")
    api_day = st.selectbox("Racecard date", ["today", "tomorrow"], index=0)
    include_ireland = st.checkbox("Include Irish meetings", value=False)
    if st.button("Fetch racecards + calculate all reports", type="primary", use_container_width=True):
        if not api_username.strip() or not api_password:
            st.error("Enter your API username and password first.")
        else:
            try:
                with st.spinner("Fetching racecards from The Racing API..."):
                    race_count, runner_count, fetched_dates = fetch_free_racecards(api_username.strip(), api_password, api_day, include_ireland)
                    meeting_count, generated_files = generate_reports_for_dates(fetched_dates)
                st.success(f"Fetched {runner_count} runners across {race_count} races and generated Excel + JSON reports for {meeting_count} meeting(s).")
                st.session_state["last_pipeline_files"] = generated_files
                st.session_state["last_pipeline_dates"] = fetched_dates
                st.session_state["last_pipeline_summary"] = f"Last run complete: {runner_count} runners across {race_count} races; reports generated for {meeting_count} meeting(s)."
                st.rerun()
            except requests.RequestException as e:
                st.error(f"Could not reach The Racing API: {e}")
            except Exception as e:
                st.error(f"Racecard fetch failed: {e}")

    if st.button("Sync today's results into historical database", use_container_width=True):
        if not api_username.strip() or not api_password:
            st.error("Enter your API username and password first.")
        else:
            try:
                with st.spinner("Syncing today's available results into your local history..."):
                    race_count, runner_count = sync_free_results(api_username.strip(), api_password, include_ireland)
                st.session_state["last_history_summary"] = f"Historical sync complete: {runner_count} runner results across {race_count} races stored/updated."
                st.success(st.session_state["last_history_summary"])
                st.rerun()
            except requests.RequestException as e:
                st.error(f"Could not reach the results API: {e}")
            except Exception as e:
                st.error(f"Historical results sync failed: {e}")

    with st.expander("Optional: manual CSV import"):
        st.caption("The normal workflow does not require spreadsheets. This is only for advanced testing or a future alternate data source.")
        uploaded = st.file_uploader("Racecard CSV", type=["csv"])
        if uploaded is not None and st.button("Import uploaded CSV", use_container_width=True):
            try:
                data_dir = ROOT / "data"
                data_dir.mkdir(parents=True, exist_ok=True)
                import_path = data_dir / "latest_uploaded_racecards.csv"
                import_path.write_bytes(uploaded.getvalue())
                core.import_csv(import_path)
                st.success("Racecards imported/updated.")
                st.rerun()
            except Exception as e:
                st.error(f"Import failed: {e}")

    st.divider()
    st.header("Demo / testing")
    st.caption("Use fictional horses to test the workflow.")
    if st.button("Load fictional demo data", use_container_width=True):
        try:
            core.import_csv(ROOT / "data" / "demo_racecard_fictional.csv")
            st.success("Fictional demo data loaded.")
            st.rerun()
        except Exception as e:
            st.error(f"Demo import failed: {e}")

core.init_db()
if st.session_state.get("last_pipeline_summary"):
    st.success(st.session_state["last_pipeline_summary"])
if st.session_state.get("last_history_summary"):
    st.info(st.session_state["last_history_summary"])
with core.connect() as con:
    meetings = con.execute(
        "SELECT meeting_date, meeting_name, course, COUNT(DISTINCT races.id) AS race_count, "
        "COUNT(runners.id) AS runner_rows "
        "FROM meetings LEFT JOIN races ON races.meeting_id=meetings.id "
        "LEFT JOIN runners ON runners.race_id=races.id "
        "GROUP BY meetings.id ORDER BY meeting_date DESC, meeting_name"
    ).fetchall()

with core.connect() as con:
    historical_race_count = con.execute("SELECT COUNT(*) FROM historical_races").fetchone()[0]
    historical_runner_count = con.execute("SELECT COUNT(*) FROM historical_runners").fetchone()[0]
col1, col2, col3, col4 = st.columns(4)
col1.metric("Meetings in database", len(meetings))
col2.metric("Current race rows", sum(m["race_count"] or 0 for m in meetings))
col3.metric("Historical races stored", historical_race_count)
col4.metric("Historical runner results", historical_runner_count)

st.header("3. Generate reports")
if not meetings:
    st.warning("No racecards in the local database yet. Enter your free API credentials in the sidebar and click Fetch racecards + calculate all reports.")
else:
    meeting_dates = sorted({m["meeting_date"] for m in meetings}, reverse=True)
    selected_date = st.selectbox("Choose race date", meeting_dates)
    date_meetings = [m for m in meetings if m["meeting_date"] == selected_date]
    selected_names = st.multiselect(
        "Choose meetings (leave all selected to generate every meeting on that date)",
        [m["meeting_name"] for m in date_meetings],
        default=[m["meeting_name"] for m in date_meetings],
    )
    make_xlsx = st.checkbox("Generate Excel workbooks", value=True)
    make_json = st.checkbox("Generate website-ready JSON files", value=True)

    if st.button("Generate selected reports", type="primary", disabled=not selected_names):
        errors = []
        progress = st.progress(0)
        for i, meeting_name in enumerate(selected_names):
            try:
                if make_xlsx:
                    core.build_report(selected_date, meeting_name)
                if make_json:
                    core.export_json(selected_date, meeting_name)
            except Exception as e:
                errors.append(f"{meeting_name}: {e}")
            progress.progress((i + 1) / len(selected_names))
        if errors:
            for error in errors:
                st.error(error)
        else:
            st.success(f"Finished generating reports for {len(selected_names)} meeting(s).")
        st.rerun()

st.header("4. Race-by-race analysis")
if meetings:
    analysis_dates = sorted({m["meeting_date"] for m in meetings}, reverse=True)
    analysis_date = st.selectbox("Analysis date", analysis_dates, key="analysis_date")
    analysis_meetings = [m["meeting_name"] for m in meetings if m["meeting_date"] == analysis_date]
    analysis_meeting = st.selectbox("Analysis meeting", analysis_meetings, key="analysis_meeting")
    with core.connect() as con:
        meeting_row = core.get_meeting(con, analysis_date, analysis_meeting)
        race_pairs = core.get_races(con, meeting_row["id"])
    if race_pairs:
        race_labels = [f"{race['race_time']} — {race['race_name']}" for race, _ in race_pairs]
        selected_race_label = st.selectbox("Race", race_labels, key="analysis_race")
        selected_race_index = race_labels.index(selected_race_label)
        race_row, race_runners = race_pairs[selected_race_index]
        analysis_rows = core.enrich_race(race_runners)
        analysis_rows = core.historical_context({**dict(race_row), "course": meeting_row["course"], "race_date": analysis_date}, analysis_rows)
        if analysis_rows:
            display_rows = [{
                "Rank": r.get("screening_rank"),
                "Runner": r.get("runner_name"),
                "Screening index": r.get("screening_index"),
                "Rating coverage %": r.get("coverage_pct"),
                "Confidence": r.get("confidence"),
                "Official rating": r.get("official_rating"),
                "Recent form": r.get("recent_form"),
                "Days since run": r.get("last_run_days"),
                "Recency profile": r.get("recency_label"),
                "Weight (lb)": r.get("weight_lbs"),
                "Trainer": r.get("trainer"),
                "Jockey": r.get("jockey"),
                "Sire": r.get("sire"),
                "Dam": r.get("dam"),
                "Prior runs stored": r.get("history_runs"),
                "Wins stored": r.get("history_wins"),
                "Historical win %": r.get("history_win_rate_pct"),
                "Course runs stored": r.get("course_runs"),
                "Course wins": r.get("course_wins"),
                "Distance-similar runs": r.get("distance_runs"),
                "Going-family runs": r.get("going_runs"),
                "Trainer runs stored": r.get("trainer_history_runs"),
                "Trainer win % (local)": r.get("trainer_win_rate_pct"),
                "Jockey runs stored": r.get("jockey_history_runs"),
                "Jockey win % (local)": r.get("jockey_win_rate_pct"),
                "Historical sample": r.get("history_sample_note"),
                "Data completeness %": r.get("data_completeness_pct"),
                "Flags": r.get("analysis_flags"),
            } for r in analysis_rows]
            st.caption("Historical stats only reflect results stored in your local archive, not the full history of British racing. They exclude results on or after the target race date. Small samples are unreliable.")
            st.dataframe(display_rows, use_container_width=True, hide_index=True)
            with st.expander("Course and race historical context"):
                st.write(f"Course: **{meeting_row['course']}**. Historical races stored at this course: **{analysis_rows[0].get('course_race_history_count', 0)}**.")
                st.write(f"Races stored at this course with similar race type and distance (within 1.5 furlongs): **{analysis_rows[0].get('course_type_distance_races', 0)}**.")
                st.write(f"Past editions of this exact race name stored: **{analysis_rows[0].get('named_race_history_count', 0)}**.")
                st.caption("These counts grow as you sync results after racing. The free API does not provide a bulk historical backfill, so this is a cumulative local archive rather than a complete historic course database.")
        else:
            st.caption("No active runners available for this race.")
else:
    st.caption("Fetch a racecard to see the analysis here.")

st.header("5. Generated files")
output_dir = ROOT / "output"
if output_dir.exists():
    outputs = sorted(
        [p for p in output_dir.iterdir() if p.is_file()],
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not outputs:
        st.caption("No reports generated yet.")
    else:
        for path in outputs[:50]:
            col_a, col_b, col_c = st.columns([5, 2, 2])
            col_a.write(path.name)
            col_b.caption(f"{path.stat().st_size / 1024:.1f} KB")
            mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if path.suffix == ".xlsx" else "application/json"
            col_c.download_button(
                "Download",
                data=path.read_bytes(),
                file_name=path.name,
                mime=mime,
                key=f"download_{path.name}",
            )
else:
    st.caption("Output folder will be created when the first report is generated.")

with st.expander("What this version does and does not do"):
    st.markdown("""
    **Available now**
    - Fetches racecards for today or tomorrow from The Racing API Free endpoint.
    - Automatically generates Excel and website-ready JSON reports immediately after fetching.
    - Stores racecards locally and preserves source metadata.
    - Syncs today's basic results from The Racing API Free endpoint into a persistent local historical database.
    - Calculates horse career/course/distance/going context and local trainer/jockey strike rates from results stored in that database, plus counts for course/type/distance history and same-named race editions.
    - Adds context columns for last-run recency, carried weight, pedigree, draw, headgear and data completeness.
    - Keeps missing ratings blank and provides a source audit.

    **Limitations**
    - The free feed may not provide speed or performance ratings; the app only uses those fields if actually returned.
    - Recency and form proxies are heuristics, not proven predictors. Weight context and pedigree are shown but not treated as proven advantages.
    - The free API does not provide a historical results backfill. The local archive starts growing when you sync results, so early course/trainer/jockey samples will be sparse. A complete pre-existing archive would require a legitimately licensed dataset or an open dataset with suitable rights.
    - No proprietary speed figures or sectionals are created; robust speed ratings require reliable race times and calibration.

    **Important:** this is a transparent research prototype, not a probability-of-winning model or betting advice. Only use data you are permitted to store and analyse.
    """)
