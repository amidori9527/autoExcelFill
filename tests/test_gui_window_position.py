from __future__ import annotations

import unittest

from PySide6.QtCore import QRect, QSize

from autoexcel.gui import centered_window_geometry


class GuiWindowPositionTest(unittest.TestCase):
    def test_centers_window_inside_primary_screen_available_area(self) -> None:
        target = centered_window_geometry(
            QSize(1280, 820),
            QRect(0, 25, 1920, 1055),
        )

        self.assertEqual(target, QRect(320, 142, 1280, 820))

    def test_centers_window_on_screen_with_negative_coordinates(self) -> None:
        target = centered_window_geometry(
            QSize(1280, 820),
            QRect(-1920, 23, 1920, 1057),
        )

        self.assertEqual(target, QRect(-1600, 141, 1280, 820))

    def test_shrinks_window_to_small_available_area(self) -> None:
        target = centered_window_geometry(
            QSize(1280, 820),
            QRect(1920, 40, 1024, 720),
        )

        self.assertEqual(target, QRect(1920, 40, 1024, 720))


if __name__ == "__main__":
    unittest.main()
