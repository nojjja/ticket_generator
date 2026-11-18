import unittest

from ui_theme import FONT, SIZE, SPACE, _scale_layout


class UiThemeScalingTests(unittest.TestCase):
    def test_layout_dimensions_scale_with_dpi_and_reset_from_base_values(self):
        base_size = SIZE.copy()
        base_space = SPACE.copy()
        base_font = FONT.copy()
        try:
            _scale_layout(1.25)
            self.assertEqual(SIZE["main_width"], round(base_size["main_width"] * 1.25))
            self.assertEqual(SIZE["main_height"], round(base_size["main_height"] * 1.25))
            self.assertEqual(SPACE["md"], round(base_space["md"] * 1.25))
            self.assertEqual(FONT, base_font)

            _scale_layout(1.25)
            self.assertEqual(SIZE["main_width"], round(base_size["main_width"] * 1.25))

            _scale_layout(1.0)
            self.assertEqual(SIZE, base_size)
            self.assertEqual(SPACE, base_space)
        finally:
            SIZE.update(base_size)
            SPACE.update(base_space)


if __name__ == "__main__":
    unittest.main()
