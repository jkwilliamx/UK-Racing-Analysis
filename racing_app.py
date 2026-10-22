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
# Transparent starting weights, not yet statistically calibrated. The score is a
# within-race ranking index, NOT a probability. Missing factors are omitted and the
# remaining weights are renormalised. Historical factors are sample-size shrunk.
FACTOR_WEIGHTS = {
    "speed_figure": 0.29,
    "performance_rating": 0.24,
    "official_rating": 0.15,
    "form_proxy": 0.15,
    "horse_history": 0.05,
    "course_history": 0.025,
    "distance_history": 0.025,
    "going_history": 0.015,
    "trainer_history": 0.02,
    "jockey_history": 0.015,
    "recency": 0.02,
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
    """Store only the free API's returned results locally; idempotent on repeated syncs."""
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
                 "The Racing API Basic horse results",source_url,stamp))
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
                     nullable_str(_first_value(runner,"comments","comment",default="")),"The Racing API Basic horse results",stamp))
                runner_count += 1
    return race_count, runner_count

def _first_value(obj, *keys, default=None):
    for key in keys:
        val = obj.get(key) if isinstance(obj, dict) else None
        if val not in (None, ""):
            return val
    return default

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

def _smoothed_rate_score(wins, runs, baseline, prior_strength):
    """Convert a win rate to a cautious 0-100 score, shrinking small samples to baseline."""
    try:
        wins, runs = float(wins or 0), int(runs or 0)
    except (TypeError, ValueError):
        return None
    if runs <= 0 or wins < 0 or wins > runs:
        return None
    smoothed_rate = (wins + baseline * prior_strength) / (runs + prior_strength)
    return max(0.0, min(100.0, 50.0 + (smoothed_rate - baseline) * 250.0))


def _horse_history_score(r):
    runs = int(r.get("history_runs") or 0)
    if runs < 2:
        return None
    win_score = _smoothed_rate_score(r.get("history_wins"), runs, 0.10, 8)
    top3 = r.get("history_top3_rate_pct")
    if top3 is None:
        return win_score
    top3_rate = max(0.0, min(1.0, float(top3) / 100.0))
    top3_smoothed = (top3_rate * runs + 0.30 * 8) / (runs + 8)
    top3_score = max(0.0, min(100.0, 50.0 + (top3_smoothed - 0.30) * 125.0))
    return 0.65 * win_score + 0.35 * top3_score if win_score is not None else top3_score


def _context_win_score(r, runs_key, wins_key, minimum_runs=2, prior_strength=8):
    runs = int(r.get(runs_key) or 0)
    if runs < minimum_runs:
        return None
    return _smoothed_rate_score(r.get(wins_key), runs, 0.10, prior_strength)


def _connection_score(r, runs_key, wins_key, minimum_runs=20):
    runs = int(r.get(runs_key) or 0)
    if runs < minimum_runs:
        return None
    return _smoothed_rate_score(r.get(wins_key), runs, 0.10, 40)


def enrich_race(runners, race=None):
    """Score active runners using available ratings plus cautiously shrunk local history."""
    active = [dict(r) for r in runners if not r.get("non_runner")]
    fields = {
        "speed_figure": [as_float(r.get("speed_figure")) for r in active],
        "performance_rating": [as_float(r.get("performance_rating")) for r in active],
        "official_rating": [as_float(r.get("official_rating")) for r in active],
        "form_proxy": [form_proxy(r.get("recent_form")) for r in active],
    }
    ranks = {key: percentile_ranks(vals) for key, vals in fields.items()}
    weights = [as_float(r.get("weight_lbs")) for r in active]
    weight_ranks = percentile_ranks([-w if w is not None else None for w in weights])
    recencies = [recency_proxy(as_float(r.get("last_run_days"))) for r in active]
    recency_ranks = percentile_ranks(recencies)
    result = []
    for idx, r in enumerate(active):
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
        context_values = {
            "horse_history": _horse_history_score(item),
            "course_history": _context_win_score(item, "course_runs", "course_wins"),
            "distance_history": _context_win_score(item, "distance_runs", "distance_wins"),
            "going_history": _context_win_score(item, "going_runs", "going_wins"),
            "trainer_history": _connection_score(item, "trainer_history_runs", "trainer_history_wins"),
            "jockey_history": _connection_score(item, "jockey_history_runs", "jockey_history_wins"),
            "recency": recency_ranks[idx],
        }
        component_ranks = {
            "speed_figure": ranks["speed_figure"][idx],
            "performance_rating": ranks["performance_rating"][idx],
            "official_rating": ranks["official_rating"][idx],
            "form_proxy": ranks["form_proxy"][idx],
            **context_values,
        }
        present_weight = sum(FACTOR_WEIGHTS[k] for k, value in component_ranks.items() if value is not None)
        score = None
        if present_weight >= 0.20:
            score = sum(FACTOR_WEIGHTS[k] * value for k, value in component_ranks.items()
                        if value is not None) / present_weight
        item.update(context_values)
        item["component_ranks"] = component_ranks
        item["coverage_pct"] = round(present_weight * 100, 1)
        item["screening_index"] = round(score, 1) if score is not None else None
        item["confidence"] = (
            "Higher factor coverage" if present_weight >= 0.80 else
            "Moderate factor coverage" if present_weight >= 0.55 else
            "Low coverage — indicative only" if present_weight >= 0.20 else
            "Insufficient scoring data"
        )
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

def build_report(date, meeting_name):
    init_db()
    with connect() as con:
        meeting = get_meeting(con, date, meeting_name)
        races = get_races(con, meeting["id"])
    races = [
        (dict(race), historical_context({**dict(race), "course": meeting["course"], "race_date": date}, runners))
        for race, runners in races
    ]
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
        enriched = enrich_race(runners, race)
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
        enriched = enrich_race(runners, race)
        for r in enriched:
            runner_rows.append([
                race["race_time"], race["race_name"], race["race_type"], race["race_class"], race["distance_f"], race["going"], race["surface"], race["field_size"],
                r.get("cloth"), r["runner_name"], r.get("age"), r.get("sex"), r.get("weight_lbs"),
                r.get("official_rating"), r.get("speed_figure"), r.get("performance_rating"), r.get("recent_form"),
                r.get("last_run_days"), r.get("recency_label"), r.get("trainer"), r.get("jockey"), r.get("draw"),
                r.get("headgear"), r.get("sire"), r.get("dam"), r.get("damsire"), r.get("owner"),
                r.get("speed_percentile"), r.get("performance_percentile"), r.get("official_rating_percentile"), r.get("form_percentile"),
                r.get("recency_proxy"), r.get("weight_context_percentile"), r.get("data_completeness_pct"),
                r.get("coverage_pct"), r.get("screening_index"), r.get("screening_rank"), r.get("confidence"),
                r.get("horse_history"), r.get("course_history"), r.get("distance_history"), r.get("going_history"),
                r.get("trainer_history"), r.get("jockey_history"), r.get("recency"), r.get("analysis_flags"),
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
        "Weight Context Percentile (context only)","Data Completeness (%)","Factor Coverage (%)","Screening Index (not probability)","Screening Rank","Evidence Coverage",
        "Horse History Score","Course History Score","Distance History Score","Going History Score","Trainer History Score","Jockey History Score","Recency Score","Analysis Flags",
        "Historical Runs Stored","Historical Wins","Historical Win Rate %","Historical Top-3 Rate %","Course Runs","Course Wins","Distance-Similar Runs","Distance-Similar Wins","Going-Family Runs","Going-Family Wins",
        "Trainer Runs Stored","Trainer Wins Stored","Trainer Win Rate % (local sample)","Jockey Runs Stored","Jockey Wins Stored","Jockey Win Rate % (local sample)",
        "Course Historical Races Stored","Course+Type+Distance Races Stored","Same-Named Race Editions Stored","Historical Sample Note",
        "Source","Source URL","Source Updated At","Notes"
    ], runner_rows)

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
        ["Purpose", "Within-race screening index only; not a win probability or betting advice."],
        ["Sourced speed figure", "29% nominal weight when available; percentile-ranked within this field."],
        ["Sourced performance rating", "24% nominal weight when available; percentile-ranked within this field."],
        ["Official rating", "15% nominal weight when available; percentile-ranked within this field."],
        ["Recent-form proxy", "15% nominal weight; simple finishing-position heuristic, not yet calibrated."],
        ["Horse history", "5% nominal weight; historical win/top-three record shrunk towards a baseline; needs at least two stored runs."],
        ["Course/distance/going history", "Small nominal weights; uses stored wins/runs, requires at least two comparable runs and shrinks small samples."],
        ["Trainer/jockey history", "Small nominal weights; uses local stored wins/runs only with at least 20 stored runs; shrunk towards baseline."],
        ["Recency", "2% nominal weight; heuristic only and not empirically validated."],
        ["Missing data", "No missing values are invented. Weights are renormalised over available components."],
        ["Coverage threshold", "An indicative score is shown from 20% factor-weight coverage; below that it is withheld."],
        ["Evidence coverage", "Share of scoring weight supported by available factors; not calibrated prediction confidence."],
        ["Weight context", "Carried weight is shown but not automatically scored; lower weight is not automatically better."],
        ["Pedigree", "Displayed when available; no unsupported stamina or going suitability is inferred."],
        ["Important limitation", "Historical data is only the local archive and may be incomplete or biased towards horses already fetched."],
        ["Validation", "Weights are provisional. Run a chronological backtest against past race results before treating ranks as predictive."],
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
    races = [
        (dict(race), historical_context({**dict(race), "course": meeting["course"], "race_date": date}, runners))
        for race, runners in races
    ]
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
        enriched = enrich_race(runners, race)
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
                "weight_context_percentile": r.get("weight_context_percentile"),
                "factor_scores": {
                    "horse_history": r.get("horse_history"), "course_history": r.get("course_history"),
                    "distance_history": r.get("distance_history"), "going_history": r.get("going_history"),
                    "trainer_history": r.get("trainer_history"), "jockey_history": r.get("jockey_history"),
                    "recency": r.get("recency")
                }, "analysis_flags": r.get("analysis_flags"),
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

