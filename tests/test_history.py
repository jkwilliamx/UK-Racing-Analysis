import sys
import tempfile
import unittest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import racing_app as core

class HistoricalContextTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.old_db = core.DB_PATH
        core.DB_PATH = Path(self.tmp.name) / "test.db"
        core.init_db()

    def tearDown(self):
        core.DB_PATH = self.old_db
        self.tmp.cleanup()

    def test_results_are_idempotent_and_historical_context_is_prior_only(self):
        payload = {"results": [{
            "race_id": "rac_old", "date": "2026-10-20", "course": "Cheltenham", "off": "14:00",
            "race_name": "Test Hurdle", "type": "hurdle", "dist_f": "16", "going": "Good to Soft",
            "runners": [{"horse": "Example Horse", "position": "1", "trainer": "Trainer A", "jockey": "Jockey A"},
                         {"horse": "Other Horse", "position": "2", "trainer": "Trainer B", "jockey": "Jockey B"}]
        }]}
        races, runners = core.store_historical_results(payload)
        self.assertEqual((races, runners), (1, 2))
        core.store_historical_results(payload)
        with core.connect() as con:
            self.assertEqual(con.execute("SELECT COUNT(*) FROM historical_races").fetchone()[0], 1)
            self.assertEqual(con.execute("SELECT COUNT(*) FROM historical_runners").fetchone()[0], 2)
        race = {"course": "Cheltenham", "race_date": "2026-10-22", "race_type": "hurdle", "distance_f": 16, "going": "Good to Soft", "race_name": "Test Hurdle"}
        enriched = core.historical_context(race, [{"runner_name": "Example Horse", "trainer": "Trainer A", "jockey": "Jockey A"}])
        self.assertEqual(enriched[0]["history_runs"], 1)
        self.assertEqual(enriched[0]["history_wins"], 1)
        self.assertEqual(enriched[0]["course_runs"], 1)
        self.assertEqual(enriched[0]["course_type_distance_going_races"], 1)
        same_day = core.historical_context({**race, "race_date": "2026-10-20"}, [{"runner_name": "Example Horse", "trainer": "Trainer A", "jockey": "Jockey A"}])
        self.assertEqual(same_day[0]["history_runs"], 0)
        self.assertEqual(same_day[0]["course_race_history_count"], 0)

if __name__ == "__main__":
    unittest.main()
