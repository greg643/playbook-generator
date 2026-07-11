import copy
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "pipeline"))

from play_geometry import merge_block_caps  # noqa: E402


class MergeBlockCapsTests(unittest.TestCase):
    def setUp(self):
        self.stub = {
            "points": [[0.1, 0.1], [0.5, 0.5]],
            "end": "none",
            "dash": False,
            "color": "#000000",
        }
        self.cap = {
            "points": [[0.47, 0.53], [0.53, 0.47]],
            "end": "none",
            "dash": False,
            "color": "#000000",
        }

    def test_merges_true_stub_and_cap(self):
        routes, remaining = merge_block_caps([copy.deepcopy(self.stub)], [copy.deepcopy(self.cap)])
        self.assertEqual(routes[0]["end"], "block")
        self.assertEqual(remaining, [])

    def test_does_not_replace_arrow(self):
        route = {**copy.deepcopy(self.stub), "end": "arrow"}
        routes, remaining = merge_block_caps([route], [copy.deepcopy(self.cap)])
        self.assertEqual(routes[0]["end"], "arrow")
        self.assertEqual(len(remaining), 1)

    def test_requires_matching_solid_style(self):
        wrong_color = {**copy.deepcopy(self.cap), "color": "#FF0000"}
        routes, remaining = merge_block_caps([copy.deepcopy(self.stub)], [wrong_color])
        self.assertEqual(routes[0]["end"], "none")
        self.assertEqual(len(remaining), 1)

        dashed = {**copy.deepcopy(self.cap), "dash": True}
        routes, remaining = merge_block_caps([copy.deepcopy(self.stub)], [dashed])
        self.assertEqual(routes[0]["end"], "none")
        self.assertEqual(len(remaining), 1)

    def test_chooses_nearest_matching_route(self):
        farther = copy.deepcopy(self.stub)
        farther["points"][-1] = [0.48, 0.50]
        nearer = copy.deepcopy(self.stub)
        routes, remaining = merge_block_caps([farther, nearer], [copy.deepcopy(self.cap)])
        self.assertEqual(routes[0]["end"], "none")
        self.assertEqual(routes[1]["end"], "block")
        self.assertEqual(remaining, [])


if __name__ == "__main__":
    unittest.main()
