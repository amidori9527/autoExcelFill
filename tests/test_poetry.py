from __future__ import annotations

from datetime import datetime
import unittest

from autoexcel.poetry import POETRY_BY_TAG, poetry_context, poetry_line_for


class PoetryTest(unittest.TestCase):
    def test_solar_term_takes_priority_over_time_of_day(self) -> None:
        moment = datetime(2026, 3, 5, 22, 30)

        self.assertEqual(poetry_context(moment), "惊蛰")
        self.assertEqual(poetry_line_for(moment), POETRY_BY_TAG["惊蛰"][0])

    def test_time_of_day_contexts(self) -> None:
        self.assertEqual(poetry_context(datetime(2026, 7, 28, 6)), "清晨")
        self.assertEqual(poetry_context(datetime(2026, 7, 28, 18)), "黄昏")
        self.assertEqual(poetry_context(datetime(2026, 7, 28, 22)), "夜晚")

    def test_daytime_uses_the_current_season(self) -> None:
        self.assertEqual(poetry_context(datetime(2026, 4, 20, 12)), "春天")
        self.assertEqual(poetry_context(datetime(2026, 7, 28, 12)), "夏天")
        self.assertEqual(poetry_context(datetime(2026, 10, 15, 12)), "秋天")
        self.assertEqual(poetry_context(datetime(2026, 1, 15, 12)), "冬天")

    def test_same_context_and_date_returns_the_same_line(self) -> None:
        moment = datetime(2026, 7, 28, 12)

        self.assertEqual(poetry_line_for(moment), poetry_line_for(moment))
        self.assertIn(poetry_line_for(moment), POETRY_BY_TAG["夏天"])


if __name__ == "__main__":
    unittest.main()
