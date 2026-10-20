import importlib.util
import tempfile
import unittest
from pathlib import Path
from xml.etree import ElementTree as ET

SPEC = importlib.util.spec_from_file_location('role_geometry', Path(__file__).parent / 'scripts' / 'audit_reference_chart_geometry.py')
AUDIT = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(AUDIT)


class IndependentGeometryAuditTests(unittest.TestCase):
    def test_chart_group_spacing_is_observed_without_inventing_defaults(self):
        class Archive:
            def read(self, part):
                return ('<c:chartSpace xmlns:c="' + AUDIT.NS['c'] + '"><c:chart><c:plotArea>'
                        '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>'
                        '<c:gapWidth val="219"/><c:overlap val="-27"/></c:barChart>'
                        '<c:lineChart><c:grouping val="standard"/></c:lineChart>'
                        '</c:plotArea></c:chart></c:chartSpace>').encode()
        groups = AUDIT.chart_metadata(Archive(), 'chart.xml')['group_settings']
        self.assertEqual(groups[0]['gapWidth']['attributes'], {'val': '219'})
        self.assertEqual(groups[0]['overlap']['attributes'], {'val': '-27'})
        self.assertIsNone(groups[1]['gapWidth'])

    def test_exact_emu_marker_is_not_rounded_to_writer_pixels(self):
        xml = '<xdr:twoCellAnchor xmlns:xdr="' + AUDIT.NS['xdr'] + '"><xdr:from><xdr:col>11</xdr:col><xdr:colOff>33336</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>9524</xdr:rowOff></xdr:from></xdr:twoCellAnchor>'
        self.assertEqual(AUDIT.marker(ET.fromstring(xml), 'from'), {'col': 11, 'colOff': 33336, 'row': 1, 'rowOff': 9524})

    def test_703_writer_pixels_are_not_527_25_native_points(self):
        columns = [{'index': index, 'metadata': {'Left': (index - 1) * 43.2}} for index in range(1, 24)]
        grid = {'columns': columns, 'rows': [{'index': 2, 'metadata': {'Top': 15}}]}
        chart = {'from': {'col': 11, 'colOff': 19 * 9525, 'row': 1, 'rowOff': 0}, 'to': {'col': 22, 'colOff': 18 * 9525, 'row': 1, 'rowOff': 0}}
        observed_grid_prediction = AUDIT.resolve_anchor(chart, grid)['Width']
        self.assertAlmostEqual(observed_grid_prediction, 474.45)
        self.assertNotEqual(observed_grid_prediction, 703 * .75)

    def test_source_alias_guard_preserves_bytes(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary) / 'sources'; root.mkdir()
            source = root / 'reference.xlsx'; source.write_bytes(b'not-for-write')
            for target in (source, root / 'new.json'):
                with self.assertRaises(ValueError):
                    AUDIT.assert_output_safe(target, [root], [source])
            self.assertEqual(source.read_bytes(), b'not-for-write')
            self.assertFalse((root / 'new.json').exists())

    def test_existing_evidence_cannot_be_overwritten(self):
        with tempfile.TemporaryDirectory() as temporary:
            evidence = Path(temporary) / 'evidence.json'; evidence.write_bytes(b'history')
            with self.assertRaises(ValueError):
                AUDIT.assert_output_safe(evidence, [], [])
            self.assertEqual(evidence.read_bytes(), b'history')

    def test_empty_png_does_not_pass_export_equality(self):
        with tempfile.TemporaryDirectory() as temporary:
            bad = Path(temporary) / 'empty.png'; bad.touch()
            self.assertEqual(AUDIT.validate_png(bad)['status'], 'UNVERIFIED_INVALID_PNG')
            self.assertEqual(AUDIT.compare_png(bad, bad)['status'], 'UNVERIFIED_INVALID_PNG')

    def test_native_scalar_guards_and_formula_failure_selftest(self):
        import subprocess
        script = Path(__file__).parent / 'scripts' / 'excel_native_full_qa.ps1'
        result = subprocess.run(['pwsh', '-NoProfile', '-File', str(script), '-SelfTest'], capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn('SELFTEST PASS', result.stdout)


if __name__ == '__main__':
    unittest.main()
