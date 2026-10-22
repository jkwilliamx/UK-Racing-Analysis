import sys
import unittest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from racing_app import percentile_ranks, form_proxy, enrich_race, recency_proxy, recency_label

class ScoringTests(unittest.TestCase):
    def test_percentile_ranks(self):
        ranks = percentile_ranks([10, 20, 30])
        self.assertEqual(ranks, [0, 50, 100])

    def test_form_proxy(self):
        self.assertGreater(form_proxy("111"), form_proxy("321"))
        self.assertIsNone(form_proxy(""))

    def test_enrich_with_missing_values(self):
        runners = [
            {"runner_name":"A","non_runner":0,"speed_figure":100,"performance_rating":120,"official_rating":110,"recent_form":"111"},
            {"runner_name":"B","non_runner":0,"speed_figure":90,"performance_rating":110,"official_rating":100,"recent_form":"321"},
        ]
        out = enrich_race(runners)
        self.assertGreaterEqual(out[0]["screening_index"], out[1]["screening_index"])
        self.assertEqual(out[0]["coverage_pct"], 100.0)

    def test_free_tier_data_gets_low_coverage_label(self):
        runners = [
            {"runner_name":"A","non_runner":0,"official_rating":110,"recent_form":"111"},
            {"runner_name":"B","non_runner":0,"official_rating":100,"recent_form":"321"},
        ]
        out = enrich_race(runners)
        self.assertIsNotNone(out[0]["screening_index"])
        self.assertEqual(out[0]["coverage_pct"], 25.0)
        self.assertIn("indicative only", out[0]["confidence"])

    def test_recency_is_separate_descriptive_heuristic(self):
        self.assertEqual(recency_proxy(30), 100.0)
        self.assertEqual(recency_label(120), "Extended gap")

if __name__ == "__main__":
    unittest.main()
