from __future__ import annotations
import argparse
import csv
import json
import math
import os
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
DB_PATH = Path(os.getenv("RACING_DATABASE_PATH", ROOT / "data" / "racing.db"))
OUTPUT_DIR = ROOT / "output"

CSV_FIELDS = [
    "date","meeting","course","country","race_time","race_name","race_type",
    "distance_f","going","cloth","runner_name","age","weight_lbs",
    "official_rating","speed_figure","performance_rating","recent_form",
    "trainer","jockey","non_runner","source_name","source_url","source_updated_at",
    "race_class","field_size","surface","prize","rating_band","age_band","sex_restriction",
    "race_id","course_id","horse_id","sex","colour","sire","dam","damsire",
    "sire_id","dam_id","damsire_id","draw","last_run_days","headgear","owner",
    "jockey_id","trainer_id"
]
# Only sourced ratings and recent form affect the primary within-race index. Contextual
# factors are calculated and displayed separately until validated with historical results.
FACTOR_WEIGHTS = {
    "speed_figure": 0.40,
    "performance_rating": 0.35,
    "official_rating": 0.15,
    "form_proxy": 0.10,
}

def connect():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys=ON")
    return con

def init_db():
    with connect() as con:
        con.executescript("""
        CREATE TABLE IF NOT EXISTS meetings (
            id INTEGER PRIMARY KEY,
            meeting_date TEXT NOT NULL,
            meeting_name TEXT NOT NULL,
            course TEXT,
            country TEXT DEFAULT 'GB',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(meeting_date, meeting_name)
        );
        CREATE TABLE IF NOT EXISTS races (
            id INTEGER PRIMARY KEY,
            meeting_id INTEGER NOT NULL REFERENCES meetings(id) ON DELETE CASCADE,
            race_time TEXT NOT NULL,
            race_name TEXT NOT NULL,
            race_type TEXT,
            distance_f REAL,
            going TEXT,
            race_class TEXT,
            field_size INTEGER,
            surface TEXT,
            prize TEXT,
            rating_band TEXT,
            age_band TEXT,
            sex_restriction TEXT,
            race_id TEXT,
            course_id TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(meeting_id, race_time, race_name)
        );
        CREATE TABLE IF NOT EXISTS runners (
            id INTEGER PRIMARY KEY,
            race_id INTEGER NOT NULL REFERENCES races(id) ON DELETE CASCADE,
            cloth TEXT,
            runner_name TEXT NOT NULL,
            age INTEGER,
            weight_lbs REAL,
            official_rating REAL,
            speed_figure REAL,
            performance_rating REAL,
            recent_form TEXT,
            trainer TEXT,
            jockey TEXT,
            horse_id TEXT,
            sex TEXT,
            colour TEXT,
            sire TEXT,
            dam TEXT,
            damsire TEXT,
            sire_id TEXT,
            dam_id TEXT,
            damsire_id TEXT,
            draw INTEGER,
            last_run_days REAL,
            headgear TEXT,
            owner TEXT,
            jockey_id TEXT,
            trainer_id TEXT,
            non_runner INTEGER NOT NULL DEFAULT 0,
            source_name TEXT,
            source_url TEXT,
            source_updated_at TEXT,
            imported_at TEXT NOT NULL,
            UNIQUE(race_id, runner_name)
        );
        CREATE TABLE IF NOT EXISTS source_records (
            id INTEGER PRIMARY KEY,
            meeting_id INTEGER REFERENCES meetings(id) ON DELETE CASCADE,
            race_id INTEGER REFERENCES races(id) ON DELETE CASCADE,
            runner_id INTEGER REFERENCES runners(id) ON DELETE CASCADE,
            source_name TEXT NOT NULL,
            source_url TEXT,
            source_updated_at TEXT,
            captured_at TEXT NOT NULL,
            record_type TEXT NOT NULL,
            note TEXT
        );
        CREATE TABLE IF NOT EXISTS historical_races (
            id INTEGER PRIMARY KEY,
            race_key TEXT NOT NULL UNIQUE,
            race_date TEXT NOT NULL,
            course TEXT,
            country TEXT,
            off_time TEXT,
            race_name TEXT,
            race_type TEXT,
            distance_f REAL,
            going TEXT,
            race_class TEXT,
            field_size INTEGER,
            source_name TEXT NOT NULL,
            source_url TEXT,
            captured_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS historical_runners (
            id INTEGER PRIMARY KEY,
            race_key TEXT NOT NULL,
            horse_key TEXT NOT NULL,
            horse_id TEXT,
            horse_name TEXT NOT NULL,
            finishing_position INTEGER,
            result_status TEXT,
            official_rating REAL,
            weight_lbs REAL,
            trainer TEXT,
            trainer_id TEXT,
            jockey TEXT,
            jockey_id TEXT,
            starting_price TEXT,
            beaten_distance TEXT,
            comments TEXT,
            source_name TEXT NOT NULL,
            captured_at TEXT NOT NULL,
            UNIQUE(race_key, horse_key)
        );
        CREATE INDEX IF NOT EXISTS idx_hist_races_course ON historical_races(course, race_date);
        CREATE INDEX IF NOT EXISTS idx_hist_runners_horse ON historical_runners(horse_key);
        CREATE INDEX IF NOT EXISTS idx_hist_runners_trainer ON historical_runners(trainer);
        CREATE INDEX IF NOT EXISTS idx_hist_runners_jockey ON historical_runners(jockey);
        CREATE INDEX IF NOT EXISTS idx_meetings_date ON meetings(meeting_date);
        CREATE INDEX IF NOT EXISTS idx_races_meeting ON races(meeting_id);
        CREATE INDEX IF NOT EXISTS idx_runners_race ON runners(race_id);
        CREATE INDEX IF NOT EXISTS idx_sources_runner ON source_records(runner_id);
        CREATE TABLE IF NOT EXISTS horse_history_sync (
            horse_id TEXT PRIMARY KEY,
            last_attempt_at TEXT NOT NULL,
            result_count INTEGER NOT NULL DEFAULT 0,
            status TEXT NOT NULL DEFAULT 'ok',
            note TEXT
        );
        """)
        # Lightweight migrations let existing local databases upgrade in place.
        _ensure_columns(con, "races", {
            "race_class": "TEXT", "field_size": "INTEGER", "surface": "TEXT", "prize": "TEXT",
            "rating_band": "TEXT", "age_band": "TEXT", "sex_restriction": "TEXT",
            "race_id": "TEXT", "course_id": "TEXT"
        })
        _ensure_columns(con, "runners", {
            "horse_id": "TEXT", "sex": "TEXT", "colour": "TEXT", "sire": "TEXT", "dam": "TEXT",
            "damsire": "TEXT", "sire_id": "TEXT", "dam_id": "TEXT", "damsire_id": "TEXT",
            "draw": "INTEGER", "last_run_days": "REAL", "headgear": "TEXT", "owner": "TEXT",
            "jockey_id": "TEXT", "trainer_id": "TEXT"
        })

def _ensure_columns(con, table, columns):
    existing = {row[1] for row in con.execute(f"PRAGMA table_info({table})").fetchall()}
    for name, sql_type in columns.items():
        if name not in existing:
            con.execute(f"ALTER TABLE {table} ADD COLUMN {name} {sql_type}")

def now_utc():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")

def nullable_str(value):
    if value is None:
        return None
    value = str(value).strip()
    return value if value else None

def as_float(value):
    value = nullable_str(value)
    if value is None:
        return None
    try:
        return float(value)
    except ValueError:
        return None

def as_int(value):
    x = as_float(value)
    return int(x) if x is not None else None

def as_bool(value):
    return str(value or "").strip().lower() in {"1","true","yes","y","non-runner","nr"}

def validate_row(row, line_no):
    missing = [f for f in ("date","meeting","race_time","race_name","runner_name") if not nullable_str(row.get(f))]
    if missing:
        raise ValueError(f"CSV line {line_no}: missing required fields: {', '.join(missing)}")
    try:
        datetime.strptime(row["date"].strip(), "%Y-%m-%d")
    except ValueError:
        raise ValueError(f"CSV line {line_no}: date must be YYYY-MM-DD")
    return

def import_csv(path):
    init_db()
    path = Path(path)
    imported = 0
    with path.open("r", encoding="utf-8-sig", newline="") as f, connect() as con:
        reader = csv.DictReader(f)
        if not reader.fieldnames:
            raise ValueError("CSV is empty or has no header row.")
        required_fields = ("date", "meeting", "race_time", "race_name", "runner_name")
        absent = [x for x in required_fields if x not in reader.fieldnames]
        if absent:
            raise ValueError("CSV header is missing required columns: " + ", ".join(absent))
        for line_no, row in enumerate(reader, start=2):
            if not any(nullable_str(v) for v in row.values()):
                continue
            validate_row(row, line_no)
            stamp = now_utc()
            meeting_name = row["meeting"].strip()
            course = nullable_str(row.get("course"))
            country = nullable_str(row.get("country")) or "GB"
            con.execute("""
                INSERT INTO meetings(meeting_date,meeting_name,course,country,created_at,updated_at)
                VALUES(?,?,?,?,?,?)
                ON CONFLICT(meeting_date,meeting_name) DO UPDATE SET
                course=COALESCE(excluded.course,meetings.course),
                country=COALESCE(excluded.country,meetings.country),
                updated_at=excluded.updated_at
            """, (row["date"].strip(), meeting_name, course, country, stamp, stamp))
            meeting = con.execute("SELECT * FROM meetings WHERE meeting_date=? AND meeting_name=?",
                                  (row["date"].strip(), meeting_name)).fetchone()
            distance = as_float(row.get("distance_f"))
            con.execute("""
                INSERT INTO races(meeting_id,race_time,race_name,race_type,distance_f,going,race_class,field_size,
                    surface,prize,rating_band,age_band,sex_restriction,race_id,course_id,created_at,updated_at)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ON CONFLICT(meeting_id,race_time,race_name) DO UPDATE SET
                race_type=COALESCE(excluded.race_type,races.race_type),
                distance_f=COALESCE(excluded.distance_f,races.distance_f),
                going=COALESCE(excluded.going,races.going),
                race_class=COALESCE(excluded.race_class,races.race_class),
                field_size=COALESCE(excluded.field_size,races.field_size),
                surface=COALESCE(excluded.surface,races.surface), prize=COALESCE(excluded.prize,races.prize),
                rating_band=COALESCE(excluded.rating_band,races.rating_band), age_band=COALESCE(excluded.age_band,races.age_band),
                sex_restriction=COALESCE(excluded.sex_restriction,races.sex_restriction),
                race_id=COALESCE(excluded.race_id,races.race_id), course_id=COALESCE(excluded.course_id,races.course_id),
                updated_at=excluded.updated_at
            """, (meeting["id"], row["race_time"].strip(), row["race_name"].strip(),
                  nullable_str(row.get("race_type")), distance, nullable_str(row.get("going")),
                  nullable_str(row.get("race_class")), as_int(row.get("field_size")), nullable_str(row.get("surface")),
                  nullable_str(row.get("prize")), nullable_str(row.get("rating_band")), nullable_str(row.get("age_band")),
                  nullable_str(row.get("sex_restriction")), nullable_str(row.get("race_id")), nullable_str(row.get("course_id")), stamp, stamp))
            race = con.execute("SELECT * FROM races WHERE meeting_id=? AND race_time=? AND race_name=?",
                               (meeting["id"], row["race_time"].strip(), row["race_name"].strip())).fetchone()
            source_name = nullable_str(row.get("source_name")) or "Unspecified CSV source"
            source_url = nullable_str(row.get("source_url"))
            source_updated_at = nullable_str(row.get("source_updated_at"))
            con.execute("""
                INSERT INTO runners(race_id,cloth,runner_name,age,weight_lbs,official_rating,speed_figure,
                    performance_rating,recent_form,trainer,jockey,horse_id,sex,colour,sire,dam,damsire,sire_id,dam_id,
                    damsire_id,draw,last_run_days,headgear,owner,jockey_id,trainer_id,non_runner,source_name,source_url,
                    source_updated_at,imported_at)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ON CONFLICT(race_id,runner_name) DO UPDATE SET
                cloth=excluded.cloth, age=excluded.age, weight_lbs=excluded.weight_lbs,
                official_rating=excluded.official_rating, speed_figure=excluded.speed_figure,
                performance_rating=excluded.performance_rating, recent_form=excluded.recent_form,
                trainer=excluded.trainer, jockey=excluded.jockey, horse_id=excluded.horse_id, sex=excluded.sex,
                colour=excluded.colour, sire=excluded.sire, dam=excluded.dam, damsire=excluded.damsire,
                sire_id=excluded.sire_id, dam_id=excluded.dam_id, damsire_id=excluded.damsire_id,
                draw=excluded.draw, last_run_days=excluded.last_run_days, headgear=excluded.headgear, owner=excluded.owner,
                jockey_id=excluded.jockey_id, trainer_id=excluded.trainer_id, non_runner=excluded.non_runner,
                source_name=excluded.source_name, source_url=excluded.source_url,
                source_updated_at=excluded.source_updated_at, imported_at=excluded.imported_at
            """, (race["id"], nullable_str(row.get("cloth")), row["runner_name"].strip(),
                  as_int(row.get("age")), as_float(row.get("weight_lbs")), as_float(row.get("official_rating")),
                  as_float(row.get("speed_figure")), as_float(row.get("performance_rating")),
                  nullable_str(row.get("recent_form")), nullable_str(row.get("trainer")), nullable_str(row.get("jockey")),
                  nullable_str(row.get("horse_id")), nullable_str(row.get("sex")), nullable_str(row.get("colour")),
                  nullable_str(row.get("sire")), nullable_str(row.get("dam")), nullable_str(row.get("damsire")),
                  nullable_str(row.get("sire_id")), nullable_str(row.get("dam_id")), nullable_str(row.get("damsire_id")),
                  as_int(row.get("draw")), as_float(row.get("last_run_days")), nullable_str(row.get("headgear")),
                  nullable_str(row.get("owner")), nullable_str(row.get("jockey_id")), nullable_str(row.get("trainer_id")),
                  1 if as_bool(row.get("non_runner")) else 0, source_name, source_url, source_updated_at, stamp))
            runner = con.execute("SELECT id FROM runners WHERE race_id=? AND runner_name=?",
                                 (race["id"], row["runner_name"].strip())).fetchone()
            con.execute("""INSERT INTO source_records(meeting_id,race_id,runner_id,source_name,source_url,
                source_updated_at,captured_at,record_type,note) VALUES(?,?,?,?,?,?,?,?,?)""",
                (meeting["id"], race["id"], runner["id"], source_name, source_url, source_updated_at,
                 stamp, "csv_import", "Imported from " + path.name))
            imported += 1
    print(f"Imported/updated {imported} runner rows from {path}. Database: {DB_PATH}")

def normalise_key(value):
    import re, unicodedata
    text = unicodedata.normalize("NFKD", str(value or "")).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()

def _result_position(runner):
    for key in ("position", "pos", "finish_position", "finishing_position", "place"):
        val = nullable_str(runner.get(key))
        if val:
            try:
                return int(float(val))
            except (ValueError, TypeError):
                if val.upper() in {"1ST", "WON"}: return 1
    return None

def store_historical_results(payload, source_url="https://api.theracingapi.com/v1/results/today"):
    """Store race results returned by a supported Racing API results endpoint.

    Repeated imports update existing race/runner rows instead of duplicating them.
    """
    init_db()
    races = payload.get("results", []) if isinstance(payload, dict) else payload
    if not isinstance(races, list):
        raise ValueError("The results response did not contain a results list.")
    stamp = now_utc()
    race_count = runner_count = 0
    with connect() as con:
        for race in races:
            if not isinstance(race, dict):
                continue
            race_date = str(_first_value(race, "date", default=""))[:10]
            course = str(_first_value(race, "course", default="Unknown course"))
            off = str(_first_value(race, "off", "off_time", default=""))
            name = str(_first_value(race, "race_name", "name", default="Unnamed race"))
            if not race_date:
                continue
            raw_id = _first_value(race, "race_id", "id", default="")
            race_key = str(raw_id) if raw_id else "|".join([race_date, normalise_key(course), normalise_key(off), normalise_key(name)])
            runners = race.get("runners", []) or []
            try:
                distance = as_float(race.get("distance_f", race.get("dist_f")))
            except Exception:
                distance = None
            con.execute("""INSERT INTO historical_races(race_key,race_date,course,country,off_time,race_name,race_type,distance_f,going,race_class,field_size,source_name,source_url,captured_at)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(race_key) DO UPDATE SET
                race_date=excluded.race_date,course=excluded.course,country=excluded.country,off_time=excluded.off_time,
                race_name=excluded.race_name,race_type=excluded.race_type,distance_f=COALESCE(excluded.distance_f,historical_races.distance_f),
                going=COALESCE(excluded.going,historical_races.going),race_class=COALESCE(excluded.race_class,historical_races.race_class),
                field_size=COALESCE(excluded.field_size,historical_races.field_size),captured_at=excluded.captured_at""",
                (race_key,race_date,course,str(_first_value(race,"region",default="GB")),off,name,
                 nullable_str(_first_value(race,"type","race_type",default="")),distance,
                 nullable_str(_first_value(race,"going","going_detailed",default="")),
                 nullable_str(_first_value(race,"class","race_class",default="")),as_int(race.get("field_size") or len(runners)),
                 "The Racing API results",source_url,stamp))
            race_count += 1
            for runner in runners:
                if not isinstance(runner, dict): continue
                horse_name = str(_first_value(runner,"horse","horse_name","runner_name","name",default="Unknown runner"))
                horse_id = nullable_str(_first_value(runner,"horse_id","id",default=""))
                horse_key = horse_id or normalise_key(horse_name)
                pos = _result_position(runner)
                status = nullable_str(_first_value(runner,"status","result_status",default=""))
                con.execute("""INSERT INTO historical_runners(race_key,horse_key,horse_id,horse_name,finishing_position,result_status,official_rating,weight_lbs,trainer,trainer_id,jockey,jockey_id,starting_price,beaten_distance,comments,source_name,captured_at)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(race_key,horse_key) DO UPDATE SET
                    horse_id=COALESCE(excluded.horse_id,historical_runners.horse_id),horse_name=excluded.horse_name,
                    finishing_position=COALESCE(excluded.finishing_position,historical_runners.finishing_position),
                    result_status=COALESCE(excluded.result_status,historical_runners.result_status),
                    official_rating=COALESCE(excluded.official_rating,historical_runners.official_rating),
                    weight_lbs=COALESCE(excluded.weight_lbs,historical_runners.weight_lbs),trainer=COALESCE(excluded.trainer,historical_runners.trainer),
                    trainer_id=COALESCE(excluded.trainer_id,historical_runners.trainer_id),jockey=COALESCE(excluded.jockey,historical_runners.jockey),
                    jockey_id=COALESCE(excluded.jockey_id,historical_runners.jockey_id),starting_price=COALESCE(excluded.starting_price,historical_runners.starting_price),
                    beaten_distance=COALESCE(excluded.beaten_distance,historical_runners.beaten_distance),comments=COALESCE(excluded.comments,historical_runners.comments),captured_at=excluded.captured_at""",
                    (race_key,horse_key,horse_id,horse_name,pos,status,as_float(_first_value(runner,"ofr","official_rating",default="")),
                     as_float(_first_value(runner,"lbs","weight_lbs",default="")),nullable_str(_first_value(runner,"trainer",default="")),
                     nullable_str(_first_value(runner,"trainer_id",default="")),nullable_str(_first_value(runner,"jockey",default="")),
                     nullable_str(_first_value(runner,"jockey_id",default="")),nullable_str(_first_value(runner,"sp","starting_price",default="")),
                     nullable_str(_first_value(runner,"btn","beaten_distance","distance_beaten",default="")),
                     nullable_str(_first_value(runner,"comments","comment",default="")),"The Racing API results",stamp))
                runner_count += 1
    return race_count, runner_count

def _first_value(obj, *keys, default=None):
    for key in keys:
        val = obj.get(key) if isinstance(obj, dict) else None
        if val not in (None, ""):
            return val
    return default

def get_horse_history_sync_status(horse_id):
    """Return cached history-sync metadata for one horse ID, if present."""
    init_db()
    with connect() as con:
        row = con.execute("SELECT * FROM horse_history_sync WHERE horse_id=?", (str(horse_id),)).fetchone()
        return dict(row) if row else None

def record_horse_history_sync(horse_id, result_count, status="ok", note=None):
    """Remember a horse-history request so repeated app runs avoid needless API calls."""
    init_db()
    with connect() as con:
        con.execute("""INSERT INTO horse_history_sync(horse_id,last_attempt_at,result_count,status,note)
            VALUES(?,?,?,?,?) ON CONFLICT(horse_id) DO UPDATE SET
            last_attempt_at=excluded.last_attempt_at,result_count=excluded.result_count,
            status=excluded.status,note=excluded.note""",
            (str(horse_id), now_utc(), int(result_count or 0), str(status), nullable_str(note)))

def shortlist_reasons(r):
    """Produce plain-language reasons and cautions using only available fields."""
    reasons = []
    concerns = []
    if as_float(r.get("screening_index")) is not None:
        rank = r.get("screening_rank")
        if rank == 1:
            reasons.append("Top race-relative screening score")
        elif rank in (2, 3):
            reasons.append(f"Strong race-relative screening score (rank {rank})")
    if as_float(r.get("speed_figure")) is not None:
        reasons.append("Sourced speed figure available")
    if as_float(r.get("performance_rating")) is not None:
        reasons.append("Sourced performance rating available")
    if as_float(r.get("official_rating")) is not None:
        reasons.append("Official rating available")
    if r.get("recent_form"):
        form = str(r.get("recent_form"))
        proxy = as_float(r.get("form_proxy"))
        if proxy is not None and proxy >= 0.6:
            reasons.append(f"Recent form looks encouraging ({form})")
        else:
            concerns.append(f"Recent form needs scrutiny ({form})")
    if as_int(r.get("history_runs")):
        runs = as_int(r.get("history_runs")) or 0
        wins = as_int(r.get("history_wins")) or 0
        if wins:
            reasons.append(f"Historical record: {wins} win(s) from {runs} stored run(s)")
        else:
            concerns.append(f"No wins in {runs} stored historical run(s)")
    else:
        concerns.append("No historical runs currently stored for this horse")
    if as_int(r.get("course_runs")):
        if (as_int(r.get("course_wins")) or 0) > 0:
            reasons.append("Has a recorded win at this course")
        else:
            concerns.append("Course experience stored, but no course win recorded")
    if as_int(r.get("distance_runs")) and (as_int(r.get("distance_wins")) or 0) > 0:
        reasons.append("Has a recorded win at a similar distance")
    if as_int(r.get("going_runs")) and (as_int(r.get("going_wins")) or 0) > 0:
        reasons.append("Has a recorded win on similar going")
    if not r.get("speed_figure") and not r.get("performance_rating"):
        concerns.append("No sourced speed or performance rating")
    if r.get("confidence", "").startswith("Low") or r.get("confidence", "").startswith("Insufficient"):
        concerns.append("Low rating-data coverage; ranking is especially uncertain")
    return ("; ".join(reasons) if reasons else "Limited evidence available",
            "; ".join(concerns) if concerns else "No major automated caution; review race manually")

def _going_family(value):
    text = normalise_key(value)
    for family, tokens in (("heavy",("heavy",)),("soft",("soft","very soft","good to soft","soft to heavy")),
                           ("good",("good","good to firm","good to yielding","yielding")),
                           ("firm",("firm","good to firm")),("standard",("standard",))):
        if any(t in text for t in tokens): return family
    return text or "unknown"

def historical_context(race, runners):
    """Attach evidence-based context from locally stored result history; never invent missing data."""
    init_db()
    with connect() as con:
        hist_races = [dict(x) for x in con.execute("SELECT * FROM historical_races").fetchall()]
        hist_runners = [dict(x) for x in con.execute("SELECT * FROM historical_runners").fetchall()]
    course_key = normalise_key(race.get("course") or "")
    race_type = normalise_key(race.get("race_type") or "")
    distance = as_float(race.get("distance_f"))
    going_family = _going_family(race.get("going"))
    race_name_key = normalise_key(race.get("race_name"))
    target_date = str(race.get("race_date") or race.get("date") or "9999")
    # Prevent same-day and future results leaking into pre-race analysis.
    hist_races = [hr for hr in hist_races if str(hr.get("race_date") or "9999") < target_date]
    race_meta = {x["race_key"]: x for x in hist_races}
    relevant_course_races = []
    relevant_course_type_distance = []
    relevant_course_type_distance_going = []
    same_named_race = []
    for hr in hist_races:
        if normalise_key(hr.get("course")) != course_key: continue
        relevant_course_races.append(hr)
        same_type = not race_type or normalise_key(hr.get("race_type")) == race_type
        d = as_float(hr.get("distance_f"))
        same_distance = distance is None or d is None or abs(d-distance) <= 1.5
        if same_type and same_distance:
            relevant_course_type_distance.append(hr)
            if _going_family(hr.get("going")) == going_family and going_family != "unknown":
                relevant_course_type_distance_going.append(hr)
        if race_name_key and normalise_key(hr.get("race_name")) == race_name_key and hr.get("race_date", "") < str(race.get("race_date") or "9999"):
            same_named_race.append(hr)
    # Course/race context, based only on results stored locally.
    result_by_key = {}
    for rr in hist_runners: result_by_key.setdefault(rr["race_key"], []).append(rr)
    course_starters = [rr for hr in relevant_course_races for rr in result_by_key.get(hr["race_key"], [])]
    context_starters = [rr for hr in relevant_course_type_distance for rr in result_by_key.get(hr["race_key"], [])]
    for item in runners:
        horse_name = item.get("runner_name") or item.get("horse_name") or ""
        horse_id = item.get("horse_id")
        horse_key = horse_id or normalise_key(horse_name)
        horse_runs = [rr for rr in hist_runners if rr.get("horse_key") == horse_key or (not horse_id and normalise_key(rr.get("horse_name")) == normalise_key(horse_name))]
        # Only use past runs for this runner; future or same-day results are excluded.
        race_date = str(race.get("race_date") or race.get("date") or "9999")
        horse_runs = [rr for rr in horse_runs if (result_by_key.get(rr["race_key"]) and
                      (race_meta.get(rr["race_key"], {}).get("race_date", "9999") < race_date))]
        horse_race_meta = [race_meta.get(rr["race_key"], {}) for rr in horse_runs]
        wins = sum(1 for rr in horse_runs if rr.get("finishing_position") == 1)
        top3 = sum(1 for rr in horse_runs if rr.get("finishing_position") is not None and rr["finishing_position"] <= 3)
        course_runs = [rr for rr, meta in zip(horse_runs, horse_race_meta) if normalise_key(meta.get("course")) == course_key]
        course_wins = sum(1 for rr in course_runs if rr.get("finishing_position") == 1)
        distance_runs = [rr for rr, meta in zip(horse_runs, horse_race_meta) if distance is not None and as_float(meta.get("distance_f")) is not None and abs(as_float(meta.get("distance_f"))-distance) <= 1.5]
        going_runs = [rr for rr, meta in zip(horse_runs, horse_race_meta) if _going_family(meta.get("going")) == going_family and going_family != "unknown"]
        item["history_runs"] = len(horse_runs)
        item["history_wins"] = wins
        item["history_win_rate_pct"] = round(100*wins/len(horse_runs),1) if horse_runs else None
        item["history_top3_rate_pct"] = round(100*top3/len(horse_runs),1) if horse_runs else None
        item["course_runs"] = len(course_runs)
        item["course_wins"] = course_wins
        item["course_win_rate_pct"] = round(100*course_wins/len(course_runs),1) if course_runs else None
        item["distance_runs"] = len(distance_runs)
        item["distance_wins"] = sum(1 for rr in distance_runs if rr.get("finishing_position") == 1)
        item["going_runs"] = len(going_runs)
        item["going_wins"] = sum(1 for rr in going_runs if rr.get("finishing_position") == 1)
        item["history_sample_note"] = "Local archive only; grows as results are synced" if horse_runs else "No prior results stored locally"
        trainer = normalise_key(item.get("trainer"))
        jockey = normalise_key(item.get("jockey"))
        trainer_runs = [rr for rr in hist_runners if trainer and normalise_key(rr.get("trainer")) == trainer and rr.get("race_key") in race_meta and race_meta[rr["race_key"]].get("race_date", "9999") < race_date]
        jockey_runs = [rr for rr in hist_runners if jockey and normalise_key(rr.get("jockey")) == jockey and rr.get("race_key") in race_meta and race_meta[rr["race_key"]].get("race_date", "9999") < race_date]
        item["trainer_history_runs"] = len(trainer_runs)
        item["trainer_history_wins"] = sum(1 for rr in trainer_runs if rr.get("finishing_position") == 1)
        item["trainer_win_rate_pct"] = round(100*item["trainer_history_wins"]/len(trainer_runs),1) if trainer_runs else None
        item["jockey_history_runs"] = len(jockey_runs)
        item["jockey_history_wins"] = sum(1 for rr in jockey_runs if rr.get("finishing_position") == 1)
        item["jockey_win_rate_pct"] = round(100*item["jockey_history_wins"]/len(jockey_runs),1) if jockey_runs else None
        item["course_race_history_count"] = len(relevant_course_races)
        item["course_type_distance_races"] = len(relevant_course_type_distance)
        item["course_type_distance_going_races"] = len(relevant_course_type_distance_going)
        item["named_race_history_count"] = len(same_named_race)
        item["course_context_winner_rate_pct"] = round(100*sum(1 for rr in context_starters if rr.get("finishing_position")==1)/len(context_starters),1) if context_starters else None
    return runners

def form_proxy(form):
    """Simple, explainable recency proxy from a form string; not a handicapping rating."""
    if not form:
        return None
    values = []
    for ch in str(form).upper():
        if ch in "123456789":
            # 1st = 1.0; 2nd = .85; scale down by finishing position.
            values.append(max(0.0, 1.0 - (int(ch)-1)*0.12))
        elif ch in {"0","P","F","U","R","B","D","C"}:
            values.append(0.0)
    if not values:
        return None
    # Recent characters usually appear at the right; use the last five known results.
    vals = values[-5:]
    return sum(vals[i] * (0.85 ** (len(vals)-1-i)) for i in range(len(vals))) / sum(0.85 ** (len(vals)-1-i) for i in range(len(vals)))

def percentile_ranks(values):
    """Map available numeric values to [0, 100], tied values receive average rank."""
    valid = [(i, v) for i, v in enumerate(values) if v is not None]
    if not valid:
        return [None] * len(values)
    if len(valid) == 1:
        out = [None] * len(values)
        out[valid[0][0]] = 50.0
        return out
    sorted_vals = sorted(v for _, v in valid)
    out = [None] * len(values)
    for i, v in valid:
        less = sum(1 for x in sorted_vals if x < v)
        equal = sum(1 for x in sorted_vals if x == v)
        rank = less + (equal - 1) / 2
        out[i] = 100.0 * rank / (len(sorted_vals) - 1)
    return out

def recency_proxy(days):
    """Transparent descriptive recency score; not a proven performance predictor."""
    if days is None:
        return None
    try:
        days = max(0.0, float(days))
    except (TypeError, ValueError):
        return None
    # Broad heuristic: 14-45 days gets the highest score; very quick returns and long
    # layoffs are flagged for review rather than treated as automatic negatives.
    if 14 <= days <= 45:
        return 100.0
    if days < 14:
        return max(35.0, 100.0 - (14 - days) * 4.0)
    if days <= 90:
        return max(65.0, 100.0 - (days - 45) * 0.6)
    if days <= 180:
        return max(25.0, 73.0 - (days - 90) * 0.65)
    return max(0.0, 14.5 - (days - 180) * 0.08)

def recency_label(days):
    if days is None:
        return "Unknown"
    try:
        d = float(days)
    except (TypeError, ValueError):
        return "Unknown"
    if d < 7:
        return "Very quick return"
    if d <= 45:
        return "Recent run"
    if d <= 90:
        return "Moderate gap"
    if d <= 180:
        return "Extended gap"
    return "Long layoff"

def data_completeness(r):
    fields = ("official_rating", "recent_form", "age", "weight_lbs", "last_run_days", "trainer", "jockey", "sire", "dam")
    present = sum(1 for f in fields if r.get(f) not in (None, ""))
    return round(100.0 * present / len(fields), 1)

def enrich_race(runners):
    """Calculate race-relative screening and descriptive context from available fields."""
    active = [dict(r) for r in runners if not r.get("non_runner")]
    fields = {
        "speed_figure": [as_float(r.get("speed_figure")) for r in active],
        "performance_rating": [as_float(r.get("performance_rating")) for r in active],
        "official_rating": [as_float(r.get("official_rating")) for r in active],
        "form_proxy": [form_proxy(r.get("recent_form")) for r in active],
    }
    ranks = {key: percentile_ranks(vals) for key, vals in fields.items()}
    # Lighter carried weight is shown as context, but is deliberately not scored as
    # automatically better: the handicapper has already allocated weight partly by ability.
    weights = [as_float(r.get("weight_lbs")) for r in active]
    weight_ranks = percentile_ranks([-w if w is not None else None for w in weights])
    recencies = [recency_proxy(as_float(r.get("last_run_days"))) for r in active]
    recency_ranks = percentile_ranks(recencies)
    result = []
    for idx, r in enumerate(active):
        present_weight = sum(FACTOR_WEIGHTS[k] for k in FACTOR_WEIGHTS if ranks[k][idx] is not None)
        score = None
        # Free racecards often contain only official rating + form. Permit an indicative
        # score at >=20% coverage, but label sparse scores clearly and never call them odds.
        if present_weight >= 0.20:
            score = sum(FACTOR_WEIGHTS[k] * ranks[k][idx] for k in FACTOR_WEIGHTS
                        if ranks[k][idx] is not None) / present_weight
        item = dict(r)
        item["form_proxy"] = fields["form_proxy"][idx]
        item["speed_percentile"] = ranks["speed_figure"][idx]
        item["performance_percentile"] = ranks["performance_rating"][idx]
        item["official_rating_percentile"] = ranks["official_rating"][idx]
        item["form_percentile"] = ranks["form_proxy"][idx]
        item["recency_proxy"] = recencies[idx]
        item["recency_percentile"] = recency_ranks[idx]
        item["recency_label"] = recency_label(as_float(r.get("last_run_days")))
        item["weight_context_percentile"] = weight_ranks[idx]
        item["data_completeness_pct"] = data_completeness(r)
        item["coverage_pct"] = round(present_weight * 100, 1)
        item["screening_index"] = round(score, 1) if score is not None else None
        item["confidence"] = ("Higher rating coverage" if present_weight >= 0.85 else
                               "Partial rating coverage" if present_weight >= 0.50 else
                               "Low coverage — indicative only" if present_weight >= 0.20 else
                               "Insufficient rating data")
        item["analysis_flags"] = analysis_flags(item)
        result.append(item)
    result.sort(key=lambda x: (x["screening_index"] is None, -(x["screening_index"] or 0), x["runner_name"].lower()))
    for rank, item in enumerate(result, start=1):
        item["screening_rank"] = rank if item["screening_index"] is not None else None
    return result

def analysis_flags(r):
    flags = []
    if r.get("non_runner"):
        flags.append("Non-runner")
    if r.get("last_run_days") not in (None, ""):
        try:
            days = float(r["last_run_days"])
            if days < 7: flags.append("Quick turnaround")
            elif days > 90: flags.append("Extended break")
        except (TypeError, ValueError):
            pass
    if r.get("official_rating") in (None, ""):
        flags.append("No official rating")
    if r.get("recent_form") in (None, ""):
        flags.append("No recent form string")
    if r.get("sire") or r.get("dam") or r.get("damsire"):
        flags.append("Pedigree available; suitability not inferred")
    if not flags:
        flags.append("No automatic flags")
    return "; ".join(flags)

def get_meeting(con, date, meeting_name):
    row = con.execute("SELECT * FROM meetings WHERE meeting_date=? AND meeting_name=?", (date, meeting_name)).fetchone()
    if not row:
        raise SystemExit(f"No meeting found for date={date}, meeting={meeting_name!r}")
    return row

def get_races(con, meeting_id):
    races = con.execute("SELECT * FROM races WHERE meeting_id=? ORDER BY race_time, race_name", (meeting_id,)).fetchall()
    packed = []
    for race in races:
        runners = con.execute("SELECT * FROM runners WHERE race_id=? ORDER BY CAST(cloth AS INTEGER), runner_name",
                              (race["id"],)).fetchall()
        packed.append((race, [dict(r) for r in runners]))
    return packed

def source_audit_rows(races):
    rows = []
    for race, runners in races:
        for r in runners:
            missing = []
            for field in ("official_rating","speed_figure","performance_rating","recent_form","trainer","jockey"):
                if r.get(field) is None or r.get(field) == "":
                    missing.append(field)
            rows.append([
                race["race_time"], race["race_name"], r["runner_name"],
                r.get("source_name"), r.get("source_url"), r.get("source_updated_at"),
                r.get("imported_at"), "Non-runner" if r.get("non_runner") else "Declared/active",
                ", ".join(missing) if missing else "None of tracked fields missing",
                "Check source permissions and freshness before use"
            ])
    return rows

def style_sheet(ws):
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    for col in ws.columns:
        max_len = 10
        for cell in list(col)[:200]:
            val = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, min(44, len(val) + 2))
        ws.column_dimensions[get_column_letter(col[0].column)].width = max_len
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

def add_table_sheet(wb, title, headers, rows):
    ws = wb.create_sheet(title)
    ws.append(headers)
    for row in rows:
        ws.append(row)
    style_sheet(ws)
    return ws


def selection_assessment(r):
    """Create a transparent evidence checklist from the locally stored history.

    This is a heuristic shortlist aid, not a trained prediction model. Missing
    history is left unknown and does not count against a runner.
    """
    components = {}

    def count(name, value):
        try:
            return int(value or 0)
        except (TypeError, ValueError):
            return 0

    # Horse-specific evidence: reward demonstrated success; otherwise treat
    # limited exposure as unknown rather than assuming a negative.
    course_runs, course_wins = count("course_runs", r.get("course_runs")), count("course_wins", r.get("course_wins"))
    if course_runs > 0:
        components["Cheltenham/course evidence"] = 2 if course_wins > 0 else (1 if course_runs >= 2 else 0)

    distance_runs, distance_wins = count("distance_runs", r.get("distance_runs")), count("distance_wins", r.get("distance_wins"))
    if distance_runs > 0:
        components["Distance evidence"] = 2 if distance_wins > 0 else (1 if distance_runs >= 2 else 0)

    going_runs, going_wins = count("going_runs", r.get("going_runs")), count("going_wins", r.get("going_wins"))
    if going_runs > 0:
        components["Going-family evidence"] = 2 if going_wins > 0 else (1 if going_runs >= 2 else 0)

    # Trainer/jockey samples are from this app's locally stored archive only.
    for label, runs_key, rate_key in [
        ("Trainer record", "trainer_history_runs", "trainer_win_rate_pct"),
        ("Jockey record", "jockey_history_runs", "jockey_win_rate_pct"),
    ]:
        runs = count(runs_key, r.get(runs_key))
        rate = as_float(r.get(rate_key))
        if runs >= 20 and rate is not None:
            components[label] = 2 if rate >= 15 else (1 if rate >= 8 else 0)
        elif runs >= 8 and rate is not None:
            components[label] = 1 if rate >= 12 else 0

    # The existing recent-form proxy is deliberately simple and not validated.
    form = as_float(r.get("form_proxy"))
    if form is not None:
        components["Recent-form proxy"] = 2 if form >= 0.75 else (1 if form >= 0.50 else 0)

    possible = len(components) * 2
    earned = sum(components.values())
    coverage = round(100 * len(components) / 6, 1)
    evidence_score = round(100 * earned / possible, 1) if possible else None

    # Only blend evidence into a candidate score when at least half of the six
    # evidence categories have data. Otherwise retain the screening index alone.
    screening = as_float(r.get("screening_index"))
    if screening is not None and coverage >= 50 and evidence_score is not None:
        candidate_score = round(0.60 * screening + 0.40 * evidence_score, 1)
        method = "60% Screening Index + 40% evidence score"
    elif screening is not None:
        candidate_score = round(screening, 1)
        method = "Screening Index only; insufficient context coverage"
    else:
        candidate_score = None
        method = "Not scored; insufficient rating data"

    return {
        "components": components, "evidence_score": evidence_score,
        "evidence_coverage_pct": coverage, "candidate_score": candidate_score,
        "method": method, "earned_points": earned, "possible_points": possible,
    }

def build_report(date, meeting_name):
    init_db()
    with connect() as con:
        meeting = get_meeting(con, date, meeting_name)
        races = get_races(con, meeting["id"])
    races = [(dict(race), historical_context({**dict(race), "course": meeting["course"], "race_date": date}, enrich_race(runners))) for race, runners in races]
    wb = Workbook()
    ws = wb.active
    ws.title = "Meeting Summary"
    ws.append(["UK Racing Analysis MVP — Meeting Summary"])
    ws["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="1F4E78")
    ws.append(["Date", meeting["meeting_date"]])
    ws.append(["Meeting", meeting["meeting_name"]])
    ws.append(["Course", meeting["course"]])
    ws.append(["Country", meeting["country"]])
    ws.append(["Generated (UTC)", now_utc()])
    ws.append(["Important", "Screening scores are within-race rankings, not win probabilities or betting advice."])
    ws.append([])
    ws.append(["Time","Race","Type","Class","Distance (f)","Going","Surface","Active runners","Non-runners","Top screening score"])
    for race, runners in races:
        enriched = enrich_race(runners)
        active = [r for r in runners if not r["non_runner"]]
        scored = [r["screening_index"] for r in enriched if r["screening_index"] is not None]
        ws.append([race["race_time"],race["race_name"],race["race_type"],race["race_class"],race["distance_f"],race["going"],race["surface"],
                   len(active),sum(1 for r in runners if r["non_runner"]),max(scored) if scored else None])
    for cell in ws[8]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    ws.freeze_panes = "A9"
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["G"].width = 20
    ws.column_dimensions["J"].width = 22

    runner_rows = []
    for race, runners in races:
        enriched = enrich_race(runners)
        for r in enriched:
            runner_rows.append([
                race["race_time"], race["race_name"], race["race_type"], race["race_class"], race["distance_f"], race["going"], race["surface"], race["field_size"],
                r.get("cloth"), r["runner_name"], r.get("age"), r.get("sex"), r.get("weight_lbs"),
                r.get("official_rating"), r.get("speed_figure"), r.get("performance_rating"), r.get("recent_form"),
                r.get("last_run_days"), r.get("recency_label"), r.get("trainer"), r.get("jockey"), r.get("draw"),
                r.get("headgear"), r.get("sire"), r.get("dam"), r.get("damsire"), r.get("owner"),
                r.get("speed_percentile"), r.get("performance_percentile"), r.get("official_rating_percentile"), r.get("form_percentile"),
                r.get("recency_proxy"), r.get("weight_context_percentile"), r.get("data_completeness_pct"),
                r.get("coverage_pct"), r.get("screening_index"), r.get("screening_rank"), r.get("confidence"), r.get("analysis_flags"),
                r.get("history_runs"), r.get("history_wins"), r.get("history_win_rate_pct"), r.get("history_top3_rate_pct"),
                r.get("course_runs"), r.get("course_wins"), r.get("distance_runs"), r.get("distance_wins"), r.get("going_runs"), r.get("going_wins"),
                r.get("trainer_history_runs"), r.get("trainer_history_wins"), r.get("trainer_win_rate_pct"),
                r.get("jockey_history_runs"), r.get("jockey_history_wins"), r.get("jockey_win_rate_pct"),
                r.get("course_race_history_count"), r.get("course_type_distance_races"), r.get("named_race_history_count"), r.get("history_sample_note"),
                r.get("source_name"), r.get("source_url"), r.get("source_updated_at"),
                "Fictional demo data" if (r.get("source_name") or "").startswith("DEMO") else ""
            ])
    add_table_sheet(wb, "Runner Database", [
        "Race Time","Race","Type","Class","Distance (f)","Going","Surface","Field Size","Cloth","Runner","Age","Sex","Weight (lb)",
        "Official Rating","Speed Figure (sourced)","Performance Rating (sourced)","Recent Form","Days Since Last Run","Recency Profile",
        "Trainer","Jockey","Draw","Headgear","Sire","Dam","Damsire","Owner",
        "Speed Percentile","Performance Percentile","OR Percentile","Form Percentile","Recency Heuristic (not validated)",
        "Weight Context Percentile (not scored)","Data Completeness (%)","Rating Coverage (%)","Screening Index","Screening Rank","Confidence","Analysis Flags",
        "Historical Runs Stored","Historical Wins","Historical Win Rate %","Historical Top-3 Rate %","Course Runs","Course Wins","Distance-Similar Runs","Distance-Similar Wins","Going-Family Runs","Going-Family Wins",
        "Trainer Runs Stored","Trainer Wins Stored","Trainer Win Rate % (local sample)","Jockey Runs Stored","Jockey Wins Stored","Jockey Win Rate % (local sample)",
        "Course Historical Races Stored","Course+Type+Distance Races Stored","Same-Named Race Editions Stored","Historical Sample Note",
        "Source","Source URL","Source Updated At","Notes"
    ], runner_rows)

    # Separate ranked shortlist sheet: uses only evidence available in the local archive.
    assessment_rows = []
    for race, runners in races:
        enriched = historical_context({**dict(race), "course": meeting["course"], "race_date": date}, enrich_race(runners))
        race_assessments = []
        for r in enriched:
            assessment = selection_assessment(r)
            components = assessment["components"]
            race_assessments.append([
                race["race_time"], race["race_name"], r.get("runner_name"),
                r.get("screening_index"), r.get("screening_rank"),
                assessment["candidate_score"], assessment["evidence_score"], assessment["evidence_coverage_pct"],
                components.get("Cheltenham/course evidence"), components.get("Distance evidence"),
                components.get("Going-family evidence"), components.get("Trainer record"),
                components.get("Jockey record"), components.get("Recent-form proxy"),
                r.get("course_runs", 0), r.get("course_wins", 0),
                r.get("distance_runs", 0), r.get("distance_wins", 0),
                r.get("going_runs", 0), r.get("going_wins", 0),
                r.get("trainer_history_runs", 0), r.get("trainer_win_rate_pct"),
                r.get("jockey_history_runs", 0), r.get("jockey_win_rate_pct"),
                assessment["method"], r.get("history_sample_note", "")
            ])
        race_assessments.sort(key=lambda row: (row[5] is None, -(row[5] or 0), row[3] is None, -(row[3] or 0)))
        assessment_rows.extend(race_assessments)
    add_table_sheet(wb, "Selection Assessment", [
        "Race Time", "Race", "Runner", "Screening Index", "Screening Rank",
        "Candidate Score (heuristic)", "Evidence Score (%)", "Evidence Coverage (%)",
        "Course Points (0-2)", "Distance Points (0-2)", "Going Points (0-2)",
        "Trainer Points (0-2)", "Jockey Points (0-2)", "Recent Form Points (0-2)",
        "Course Runs Stored", "Course Wins Stored", "Similar-Distance Runs", "Similar-Distance Wins",
        "Going-Family Runs", "Going-Family Wins", "Trainer Runs Stored", "Trainer Win Rate % (local)",
        "Jockey Runs Stored", "Jockey Win Rate % (local)", "Scoring Method", "Historical Data Note"
    ], assessment_rows)

    # Compact top-three shortlist, one row per selected runner per race.
    shortlist_rows = []
    for race, runners in races:
        scored = []
        for r in runners:
            assessment = selection_assessment(r)
            score = assessment.get("candidate_score")
            if score is None:
                score = as_float(r.get("screening_index"))
            reasons, concerns = shortlist_reasons(r)
            scored.append((r, assessment, score, reasons, concerns))
        scored.sort(key=lambda item: (item[2] is None, -(item[2] or 0), item[0].get("runner_name", "").lower()))
        ranked = [item for item in scored if item[2] is not None][:3]
        if not ranked:
            ranked = scored[:3]
        for rank, (r, assessment, score, reasons, concerns) in enumerate(ranked, start=1):
            shortlist_rows.append([
                race.get("race_time"), race.get("race_name"), rank, r.get("runner_name"),
                round(score, 1) if score is not None else None,
                r.get("screening_index"), r.get("official_rating"), r.get("speed_figure"),
                r.get("performance_rating"), r.get("recent_form"), reasons, concerns,
                assessment.get("evidence_coverage_pct"), r.get("history_runs", 0),
                r.get("course_runs", 0), r.get("distance_runs", 0), r.get("going_runs", 0),
                "Heuristic shortlist; not a win probability"
            ])
    add_table_sheet(wb, "Selection Shortlist", [
        "Race Time", "Race", "Shortlist Rank", "Horse", "Candidate Score (heuristic)",
        "Screening Index", "Official Rating", "Sourced Speed Figure", "Sourced Performance Rating",
        "Recent Form", "Why shortlisted", "Concerns / checks", "Historical Evidence Coverage %",
        "Stored Horse Runs", "Course Runs Stored", "Similar-Distance Runs Stored", "Similar-Going Runs Stored",
        "Interpretation"
    ], shortlist_rows)

    history_rows = []
    for race, runners in races:
        if runners:
            first = runners[0]
            history_rows.append([race["race_time"], race["race_name"], meeting["course"], race["race_type"], race["distance_f"], race["going"],
                                 first.get("course_race_history_count", 0), first.get("course_type_distance_races", 0),
                                 f"Same-going course/type/distance races stored: {first.get('course_type_distance_going_races', 0)}; past named-race editions: {first.get('named_race_history_count', 0)}",
                                 "Local result archive only; not a complete historical record"])
            for r in runners:
                history_rows.append([race["race_time"], "  Runner: " + r["runner_name"], meeting["course"], race["race_type"], race["distance_f"], race["going"],
                                     f"Horse runs: {r.get('history_runs', 0)}; wins: {r.get('history_wins', 0)}; win rate: {r.get('history_win_rate_pct')}",
                                     f"Course runs: {r.get('course_runs', 0)}; course wins: {r.get('course_wins', 0)}",
                                     f"Distance-similar runs: {r.get('distance_runs', 0)}; going-family runs: {r.get('going_runs', 0)}",
                                     f"Trainer: {r.get('trainer_history_runs', 0)} runs / {r.get('trainer_history_wins', 0)} wins; jockey: {r.get('jockey_history_runs', 0)} / {r.get('jockey_history_wins', 0)}; {r.get('history_sample_note', '')}"])
    add_table_sheet(wb, "Historical Context", ["Race Time", "Race / Runner", "Course", "Race Type", "Distance (f)", "Going", "Course Races Stored / Horse Record", "Course+Type+Distance / Course Record", "Same Race / Distance+Going Context", "Coverage Note"], history_rows)

    audit = source_audit_rows(races)
    add_table_sheet(wb, "Source Audit", [
        "Race Time","Race","Runner","Source Name","Source URL","Source Updated At",
        "Imported At","Status","Missing tracked fields","Audit Note"
    ], audit)

    method_rows = [
        ["Purpose", "Within-race screening aid only; not a probability of winning, tip, or betting advice."],
        ["Sourced speed figure", "40% weight when available; the free feed may leave it blank."],
        ["Sourced performance rating", "35% weight when available; the free feed may leave it blank."],
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
        ["Selection Assessment sheet", "Separate heuristic assessment using course, similar-distance, going-family, trainer, jockey and recent-form evidence when available."],
        ["Selection Shortlist sheet", "Shows up to three contenders per race, with plain-language supporting evidence and cautions. It is not a probability-of-winning model."],
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

