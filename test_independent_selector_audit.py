import math
import random
import unittest
import hashlib
import json
from pathlib import Path
import tempfile

from scripts.independent_selector_audit import (
    COUNTS, LABELS, ROW, audit_task, oracle, candidate_identity, equal_field, displayed_candidates,
    SCHEMA, retained_results,
)


class IndependentSelectorTests(unittest.TestCase):
    def test_missing_is_not_zero_and_cap_is_not_plot_limit(self):
        empty = oracle([(501., 1, 0.)])
        self.assertIsNone(empty['selected'])
        self.assertEqual(empty['over_cap_count'], 1)
        self.assertEqual(empty['over_cap_max'], 501.)
        zero = oracle([(0., 1, 0.), (0., 2, 0.), (0., 3, 0.)])
        self.assertEqual((zero['selected'], zero['cluster_count']), (0., 3))

    def test_isolated_max_deep_cluster_and_stable_ties(self):
        rows = [(500., 1, 0.), (300., 2, 11.), (100., 3, 22.), (90., 4, 33.), (91., 5, 44.)]
        result = oracle(rows)
        self.assertEqual((result['selected'], result['selected_rank'], result['cluster_count']), (100., 3, 3))
        ties = oracle([(100., 10, 20.), (100., 2, 30.), (100., 5, 40.)])
        self.assertEqual(ties['selected_bearing'], 30.)

    def test_fallback_and_complete_window_count(self):
        self.assertEqual(oracle([(100., 1, 0.), (1., 2, 0.)])['cluster_count'], 1)
        result = oracle([(100., 1, 0.), (91., 2, 0.), (90., 3, 0.), (90., 4, 0.), (89., 5, 0.)])
        self.assertEqual(result['cluster_count'], 4)

    def test_reject_nonfinite_and_negative(self):
        for value in (math.nan, math.inf, -1):
            with self.assertRaises(ValueError):
                oracle([(value, 1, 0.)])

    def test_against_slow_exhaustive_definition_including_epsilon_edges(self):
        rng = random.Random(32)
        for _ in range(200):
            values = [rng.uniform(0, 700) for _ in range(rng.randrange(80))]
            values += [100., 90., 90. - 1e-12, 90. - 2e-12, 500.]
            rows = [(value, index, float(index)) for index, value in enumerate(values)]
            descending = sorted((row for row in rows if row[0] <= 500), key=lambda row: (-row[0], row[1]))
            expected = None
            for rank, candidate in enumerate(descending, 1):
                count = sum(row[0] + 1e-12 >= candidate[0] * .9 for row in descending[rank - 1:])
                if count >= 3:
                    expected = (candidate[0], candidate[2], rank, count)
                    break
            if expected is None:
                expected = (descending[0][0], descending[0][2], 1, 1)
            actual = oracle(rows)
            self.assertEqual(tuple(actual[key] for key in ('selected', 'selected_bearing', 'selected_rank', 'cluster_count')), expected)

    def test_all_96_selectors_and_tampered_numeric_field(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, cache, spool = root / 'source.xlsx', root / 'cache.json', root / 'day.bin'
            source.write_bytes(b'ISOLATED_TEST_INPUT')
            data = b'AISNRM1\0' + COUNTS.pack(3, *([0] * 31))
            data += b''.join(ROW.pack(distance, 0., index, index, 1, 0, -1, b'A')
                             for index, distance in enumerate((100., 95., 90.), 2))
            spool.write_bytes(data)
            fragment = {'path': str(source), 'sheet': 'AIS', 'rows_scanned': 3,
                        'sha256': hashlib.sha256(source.read_bytes()).hexdigest()}
            empty = {'selected': None, 'selected_bearing': None, 'selected_rank': None,
                     'selected_candidate_id': None,
                     'cluster_count': 0, 'over_cap_count': 0, 'over_cap_max': None, 'candidates': []}
            nonempty = {**empty, 'selected': 100., 'selected_bearing': 0.,
                        'selected_candidate_id': candidate_identity(fragment, (100., 2, 0., 2, 1, -1, 'A')),
                        'selected_rank': 1, 'cluster_count': 3}
            full_records = [(distance, index, 0., index, 1, -1, 'A')
                            for index, distance in enumerate((100., 95., 90.), 2)]
            nonempty['candidates'] = displayed_candidates(full_records, nonempty, [fragment], '2026-01-01')
            modern = {label: dict(nonempty if index == 0 else empty) for index, label in enumerate(LABELS)}
            period_b = {label: dict(empty) for label in LABELS}
            payload = {'signature': {'message_types': [1, 2, 3, 18, 19], 'max_distance': 500.,
                                    'tolerance': .1, 'cluster_size': 3, 'top_candidates': 50,
                                    'date': '2026-01-01', 'port': 'HWLN'},
                       'result': {'day': '2026-01-01', 'directions': modern,
                                  'rows_normalized': 3,
                                  'fragment_info': [fragment],
                                  'period_directions': {'period_a': modern, 'period_b': period_b},
                                  'spool_sha256': hashlib.sha256(data).hexdigest()}}
            cache.write_text(json.dumps(payload), encoding='utf-8')
            task = {'port': 'HWLN', 'day': '2026-01-01', 'paths': [str(source)],
                    'cache': str(cache), 'spool': str(spool)}
            result = audit_task(task)
            self.assertTrue(result['pass'])
            self.assertEqual((result['selectors_checked'], result['numeric_fields_checked']), (96, 576))
            self.assertEqual(result['identity_fields_checked'], 96)
            self.assertEqual(result['candidate_rows_checked'], 6)
            report = {'schema': SCHEMA, 'source_root': str(root.resolve()),
                      'run_root': str(root.resolve()), 'planned_days': 1,
                      'completed_days': 1, 'results': [result]}
            self.assertEqual(len(retained_results(report, [task], root, root)), 1)
            wrong = json.loads(json.dumps(report))
            wrong['results'][0]['rows_scanned'] = 2
            with self.assertRaisesRegex(ValueError, 'seal/input stamps'):
                retained_results(wrong, [task], root, root)
            wrong = json.loads(json.dumps(report))
            wrong['results'].append(wrong['results'][0])
            wrong['completed_days'] = 2
            with self.assertRaisesRegex(ValueError, 'duplicate'):
                retained_results(wrong, [task], root, root)
            wrong = {**report, 'schema': 'AIS_INDEPENDENT_SELECTOR_AUDIT_1'}
            with self.assertRaisesRegex(ValueError, 'schema/root/plan'):
                retained_results(wrong, [task], root, root)
            # A coordinated cache/ledger wrong ID was invisible to the former
            # six-field numerical oracle. All numerical fields still agree.
            payload['result']['directions']['北']['selected_candidate_id'] = 'wrong-id'
            cache.write_text(json.dumps(payload), encoding='utf-8')
            result = audit_task(task)
            self.assertFalse(result['pass'])
            self.assertTrue(any(issue['field'] == 'selected_candidate_id' for issue in result['issues']))
            payload['result']['directions']['北']['selected_candidate_id'] = nonempty['selected_candidate_id']
            payload['result']['directions']['北']['candidates'][0]['source_row'] = 999
            cache.write_text(json.dumps(payload), encoding='utf-8')
            result = audit_task(task)
            self.assertFalse(result['pass'])
            self.assertTrue(any(issue['field'] == 'candidates' for issue in result['issues']))
            payload['result']['directions']['北']['candidates'][0]['source_row'] = 2
            payload['result']['directions']['北']['selected'] = 99.
            cache.write_text(json.dumps(payload), encoding='utf-8')
            result = audit_task(task)
            self.assertFalse(result['pass'])
            self.assertTrue(any(issue['field'] == 'selected' for issue in result['issues']))
            payload['result']['fragment_info'][0]['sha256'] = '0' * 64
            cache.write_text(json.dumps(payload), encoding='utf-8')
            with self.assertRaisesRegex(ValueError, 'SHA differs'):
                audit_task(task)

    def test_numerical_booleans_and_missing_identity_are_not_equivalent(self):
        self.assertFalse(equal_field(1, True))
        self.assertFalse(equal_field(0, False))
        self.assertFalse(equal_field(None, ''))
        self.assertTrue(equal_field(1., 1))
        self.assertEqual(oracle([(100., 10, 20.), (100., 2, 30.), (100., 5, 40.)])['selected_source_key'], 2)

    def test_deep_selected_display_cluster_does_not_truncate_selector_search(self):
        fragment = {'path': 'source.xlsx', 'sheet': 'AIS', 'sha256': '0' * 64}
        records = [(float(500 - index * 4), index + 2, 0., 0, 1, -1, 'A') for index in range(80)]
        choice = {'selected_rank': 61, 'cluster_count': 3}
        candidates = displayed_candidates(records, choice, [fragment], '2026-01-01')
        self.assertEqual([candidate['rank'] for candidate in candidates], list(range(1, 51)) + [61, 62, 63])

    def test_out_of_contract_configuration_is_not_silently_used_as_expected(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, cache, spool = root / 'source.xlsx', root / 'cache.json', root / 'day.bin'
            source.write_bytes(b'ISOLATED_TEST_INPUT')
            spool.write_bytes(b'ISOLATED_TEST_SPOOL')
            cache.write_text(json.dumps({'signature': {'message_types': [1, 2, 3, 18, 19],
                                                       'max_distance': 500, 'tolerance': .15},
                                         'result': {}}), encoding='utf-8')
            task = {'port': 'HWLN', 'day': '2026-01-01', 'paths': [str(source)],
                    'cache': str(cache), 'spool': str(spool)}
            with self.assertRaisesRegex(ValueError, 'documented default-rule'):
                audit_task(task)


if __name__ == '__main__':
    unittest.main()
