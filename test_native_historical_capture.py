import math
import unittest
from scripts.validate_native_historical_capture import canonical_native_formula, point_geometry, value


class HistoricalNativeGeometryTests(unittest.TestCase):
    plot = dict(left=20., top=30., width=100., height=100.)

    def test_genuine_zero_center_not_missing(self):
        geometry = dict(left=70., top=80., width=0., height=0.)
        point_geometry(geometry, self.plot, 0., 1, 1, 1, 32, 0., 100.)
        with self.assertRaises(ValueError):
            point_geometry(geometry, self.plot, None, 1, 1, 1, 32, 0., 100.)
        with self.assertRaises(ValueError):
            value(0., None)

    def test_missing_positive_and_negative_geometry(self):
        for left, top in ((70000., 47000.), (-70000., 47000.), (70., -84000.), (70., 84000.)):
            point_geometry(dict(left=left, top=top, width=0., height=0.), self.plot, None, 1, 1, 12, 32, 0., 100.)

    def test_shifted_direction_or_wrong_radius_fails(self):
        north = dict(left=70., top=55., width=0., height=0.)
        point_geometry(north, self.plot, 50., 1, 1, 1, 32, 0., 100.)
        for index, distance in ((9, 50.), (1, 25.)):
            with self.assertRaises(ValueError):
                point_geometry(north, self.plot, distance, 1, 1, index, 32, 0., 100.)

    def test_bar_and_line_value_geometry(self):
        for series, geometry in ((1, dict(left=30., top=80., width=5., height=50.)),
                                 (3, dict(left=30., top=80., width=0., height=0.))):
            point_geometry(geometry, self.plot, 50., 2, series, 1, 32, 0., 100.)
            with self.assertRaises(ValueError):
                point_geometry(geometry, self.plot, 25., 2, series, 1, 32, 0., 100.)

    def test_nonfinite_missing_geometry_rejected(self):
        for bad in (None, math.nan, math.inf, True, "0"):
            with self.assertRaises(ValueError):
                point_geometry(dict(left=bad, top=84000., width=0., height=0.), self.plot, None, 1, 1, 12, 32, 0., 100.)

    def test_primary_empty_is_not_a_native_missing_series_point(self):
        value("", None, primary=True)
        with self.assertRaises(ValueError):
            value("", None)

    def test_native_known_function_prefix_only(self):
        for function in ("MAXIFS", "STDEV.S"):
            expected = f'=IF(COUNT(B2:B32)<1,"",_xlfn.{function}(B2:B32))'
            actual = expected.replace("_xlfn.", "")
            self.assertEqual(canonical_native_formula(expected), canonical_native_formula(actual))
            for wrong in (actual.replace("B2:B32", "C2:C32"), actual.replace(function, "SUM")):
                self.assertNotEqual(canonical_native_formula(expected), canonical_native_formula(wrong))

    def test_prefix_literals_and_unknown_functions_preserved(self):
        for formula in ('="_xlfn.MAXIFS(B2)"', "=_xlfn.UNKNOWN(B2)"):
            self.assertEqual(canonical_native_formula(formula), formula)

    def test_nonfinite_axis_and_plot_cannot_pass(self):
        point = dict(left=30., top=80., width=0., height=0.)
        for invalid in (math.nan, math.inf, None, True, "100"):
            with self.assertRaises(ValueError):
                point_geometry(point, self.plot, 50., 2, 3, 1, 32, 0., invalid)

    def test_native_null_bar_phantom_height_is_not_a_zero(self):
        point = dict(left=60., top=-84390., width=4., height=84520.)
        point_geometry(point, self.plot, None, 2, 1, 29, 32, 0., 100.)
        with self.assertRaises(ValueError):
            point_geometry(point, self.plot, 0., 2, 1, 29, 32, 0., 100.)

    def test_exact_zero_class_is_symmetric_for_tiny_positive(self):
        from scripts.compare_native_regression_pair import _native_value
        for checker in (value, _native_value):
            checker(0., 0.)
            checker(1e-14, 1e-14)
            for actual, expected in ((0., 1e-14), (1e-14, 0.)):
                with self.assertRaises(ValueError):
                    checker(actual, expected)


if __name__ == "__main__":
    unittest.main()
