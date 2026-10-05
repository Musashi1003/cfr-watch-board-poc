import unittest

from act_postgres import (
  ActConflictError,
  _normalized_updates,
  required_previous_week_code,
  rows_checksum,
  week_code,
)


class ActPostgresHelpersTest(unittest.TestCase):
  def test_week_code_and_continuity(self):
    self.assertEqual(2639, week_code("W2639"))
    self.assertEqual(2638, required_previous_week_code(2639))
    self.assertIsNone(required_previous_week_code(2701))

  def test_updates_are_rounded_and_deduplicated(self):
    rows = _normalized_updates(
      [
        {
          "source_type": "Gaming NB",
          "launch_year": "2026",
          "model": "FA401EA",
          "week": "w2639",
          "cumulative_activation": 5985.2,
        },
        {
          "source_type": "Gaming NB",
          "launch_year": "2026",
          "model": "FA401EA",
          "week": "W2639",
          "cumulative_activation": 5985,
        },
      ]
    )
    self.assertEqual(1, len(rows))
    self.assertEqual(5985, rows[0]["act_qty"])

  def test_conflicting_duplicate_is_rejected(self):
    with self.assertRaises(ActConflictError):
      _normalized_updates(
        [
          {
            "source_type": "PC NB",
            "launch_year": "2026",
            "model": "UX3407NA",
            "week": "W2639",
            "cumulative_activation": 2683,
          },
          {
            "source_type": "PC NB",
            "launch_year": "2026",
            "model": "UX3407NA",
            "week": "W2639",
            "cumulative_activation": 2684,
          },
        ]
      )

  def test_checksum_is_order_independent(self):
    left = [
      {"source_type": "PC NB", "launch_year": 2026, "model": "B", "week": "W2639", "act_qty": 2},
      {"source_type": "PC NB", "launch_year": 2026, "model": "A", "week": "W2639", "act_qty": 1},
    ]
    self.assertEqual(rows_checksum(left), rows_checksum(reversed(left)))


if __name__ == "__main__":
  unittest.main()
