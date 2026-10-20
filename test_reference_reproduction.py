import unittest

from scripts.reproduce_june_reference import compare_grid


class ReferenceReproductionTests(unittest.TestCase):
    def test_partial_or_missing_grid_cannot_pass_through_zip_truncation(self):
        complete = [[None] * 32 for _ in range(30)]
        for invalid in ([], complete[:29], [[None] * 31 for _ in range(30)]):
            with self.assertRaises(ValueError):
                compare_grid(complete, invalid)
            with self.assertRaises(ValueError):
                compare_grid(invalid, complete)

    def test_all_cells_and_blank_versus_zero_are_distinct(self):
        expected = [[None] * 32 for _ in range(30)]
        actual = [[None] * 32 for _ in range(30)]
        self.assertEqual(compare_grid(expected, actual)['cells'], 960)
        self.assertTrue(compare_grid(expected, actual)['pass'])
        actual[29][31] = 0.
        result = compare_grid(expected, actual)
        self.assertFalse(result['pass'])
        self.assertEqual(result['mismatch_count'], 1)
        self.assertEqual(result['mismatch_sample'][0]['direction_index'], 31)


if __name__ == '__main__':
    unittest.main()
