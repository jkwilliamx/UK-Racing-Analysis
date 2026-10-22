        ["Official rating", "15% weight when available; percentile-ranked among active runners."],
        ["Recent-form proxy", "10% weight; transparent heuristic from recent finishing positions, not a validated form model."],
        ["Missing data", "No imputation. Weights are renormalized over available rating factors."],
        ["Coverage threshold", "An indicative score is shown from 20% rating-weight coverage; below that it is withheld."],
        ["Confidence", "Low coverage means the score is especially fragile; do not treat ranks as probabilities."],
        ["Recency heuristic", "Displayed separately from the main score; broad heuristic, not empirically validated."],
        ["Weight context", "Relative carried-weight rank is displayed but not scored as lower weight is not automatically better."],
        ["Pedigree", "Sire, dam and damsire are displayed when available; no unsupported suitability claim is inferred."],
        ["Not inferred from racecard alone", "Course/distance/going record, stamina, pace, jumping, fitness and trainer/jockey strike rates require historical datasets."],
        ["Speed modelling", "Do not estimate speed from racecard alone. Reliable time, distance, course and going data plus calibration are required."],
        ["Selection Assessment sheet", "Separate heuristic shortlist using course, similar-distance, going-family, trainer, jockey and recent-form evidence when available."],
        ["Candidate Score", "When at least 3 of 6 evidence categories are available: 60% Screening Index + 40% evidence score. Otherwise the score falls back to Screening Index only."],
        ["Evidence points", "Each available category is scored 0-2 using transparent rules; missing history is left blank and does not count against a runner."],
        ["Historical archive limitation", "Course, distance, going, trainer and jockey stats only reflect results stored in this app; they are not complete external career records."],
        ["Validation", "Backtest with time-ordered historical data before relying on any ranking."],
        ["Data rights", "Confirm source licence covers storage, derived outputs, exports and any onward display."],
    ]
    add_table_sheet(wb, "Scoring Method", ["Item","Definition"], method_rows)
    for sheet in wb.worksheets:
        if sheet.title != "Meeting Summary":
            sheet.sheet_view.showGridLines = False
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    safe_meeting = "".join(c if c.isalnum() or c in "-_" else "_" for c in meeting_name.replace(" ", "_"))
    path = OUTPUT_DIR / f"{date}_{safe_meeting}_analysis.xlsx"
    wb.save(path)
    print(f"Workbook created: {path}")

def export_json(date, meeting_name):
    init_db()
    with connect() as con:
        meeting = get_meeting(con, date, meeting_name)
        races = get_races(con, meeting["id"])
    races = [(dict(race), historical_context({**dict(race), "course": meeting["course"], "race_date": date}, enrich_race(runners))) for race, runners in races]
    payload = {
        "schema_version": "1.0",
        "generated_at_utc": now_utc(),
        "meeting": {
            "date": meeting["meeting_date"], "name": meeting["meeting_name"],
            "course": meeting["course"], "country": meeting["country"]
        },
        "disclaimer": "Screening scores are not win probabilities or betting advice.",
        "races": []
    }
    for race, runners in races:
        enriched = enrich_race(runners)
        payload["races"].append({
            "time": race["race_time"], "name": race["race_name"], "type": race["race_type"],
            "class": race["race_class"], "distance_f": race["distance_f"], "going": race["going"],
            "surface": race["surface"], "field_size": race["field_size"], "prize": race["prize"],
            "runners": [{
                "cloth": r.get("cloth"), "name": r["runner_name"], "age": r.get("age"), "sex": r.get("sex"),
                "weight_lbs": r.get("weight_lbs"), "official_rating": r.get("official_rating"),
                "speed_figure": r.get("speed_figure"), "performance_rating": r.get("performance_rating"),
                "recent_form": r.get("recent_form"), "form_proxy": r.get("form_proxy"),
                "last_run_days": r.get("last_run_days"), "recency_profile": r.get("recency_label"),
                "trainer": r.get("trainer"), "jockey": r.get("jockey"), "draw": r.get("draw"),
                "headgear": r.get("headgear"), "sire": r.get("sire"), "dam": r.get("dam"), "damsire": r.get("damsire"),
                "non_runner": bool(r.get("non_runner")), "screening_index": r.get("screening_index"),
                "screening_rank": r.get("screening_rank"), "rating_coverage_pct": r.get("coverage_pct"),
                "data_completeness_pct": r.get("data_completeness_pct"), "recency_heuristic": r.get("recency_proxy"),
                "weight_context_percentile": r.get("weight_context_percentile"), "analysis_flags": r.get("analysis_flags"),
                "confidence": r.get("confidence"), "historical": {
                    "runs_stored": r.get("history_runs"), "wins_stored": r.get("history_wins"),
                    "win_rate_pct": r.get("history_win_rate_pct"), "top3_rate_pct": r.get("history_top3_rate_pct"),
                    "course_runs": r.get("course_runs"), "course_wins": r.get("course_wins"),
                    "distance_similar_runs": r.get("distance_runs"), "distance_similar_wins": r.get("distance_wins"),
                    "going_family_runs": r.get("going_runs"), "going_family_wins": r.get("going_wins"),
                    "trainer_runs_stored": r.get("trainer_history_runs"), "trainer_win_rate_pct": r.get("trainer_win_rate_pct"),
                    "jockey_runs_stored": r.get("jockey_history_runs"), "jockey_win_rate_pct": r.get("jockey_win_rate_pct"),
                    "course_races_stored": r.get("course_race_history_count"), "course_type_distance_races_stored": r.get("course_type_distance_races"),
                    "same_named_race_editions_stored": r.get("named_race_history_count"), "sample_note": r.get("history_sample_note")
                }, "source": r.get("source_name"),
                "source_url": r.get("source_url"), "source_updated_at": r.get("source_updated_at")
            } for r in enriched]
        })
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    safe_meeting = "".join(c if c.isalnum() or c in "-_" else "_" for c in meeting_name.replace(" ", "_"))
    path = OUTPUT_DIR / f"{date}_{safe_meeting}_website.json"
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Website-ready JSON created: {path}")

def build_day(date, csv_path=None):
    """Optional daily orchestration for an already-normalized provider CSV."""
    if csv_path:
        import_csv(csv_path)
    init_db()
    with connect() as con:
        meetings = con.execute(
            "SELECT meeting_name FROM meetings WHERE meeting_date=? ORDER BY meeting_name", (date,)
        ).fetchall()
    if not meetings:
        raise SystemExit(f"No meetings found for {date}. Import racecards first or connect a provider adapter.")
    for meeting in meetings:
        build_report(date, meeting["meeting_name"])
        export_json(date, meeting["meeting_name"])
    print(f"Daily pipeline complete for {date}: {len(meetings)} meeting(s).")

def main():
    parser = argparse.ArgumentParser(description="UK Racing Analysis MVP")
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("init-db", help="Create/update the local SQLite schema")
    imp = sub.add_parser("import-csv", help="Import normalized racecard CSV")
    imp.add_argument("path")
    rep = sub.add_parser("build-report", help="Build a meeting workbook")
    rep.add_argument("--date", required=True)
    rep.add_argument("--meeting", required=True)
    exp = sub.add_parser("export-json", help="Export meeting data for a future website")
    exp.add_argument("--date", required=True)
    exp.add_argument("--meeting", required=True)
    day = sub.add_parser("build-day", help="Import an optional normalized CSV and build all meetings for a date")
    day.add_argument("--date", required=True)
    day.add_argument("--csv", help="Optional normalized racecard CSV to import first")
    args = parser.parse_args()
    if args.command == "init-db":
        init_db()
        print(f"Database initialized: {DB_PATH}")
    elif args.command == "import-csv":
        import_csv(args.path)
    elif args.command == "build-report":
        build_report(args.date, args.meeting)
    elif args.command == "export-json":
        export_json(args.date, args.meeting)
    elif args.command == "build-day":
        build_day(args.date, args.csv)

if __name__ == "__main__":
    main()
