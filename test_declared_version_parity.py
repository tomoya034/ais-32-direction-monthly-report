"""A version declaration must never hide numeric, formula or other metadata changes."""
from pathlib import Path
import tempfile
import unittest
import xlsxwriter
from scripts.validate_full_regression import semantic_package_fingerprint


class DeclaredVersionParityTests(unittest.TestCase):
    def test_only_exact_declared_application_version_is_normalized(self):
        with tempfile.TemporaryDirectory(prefix='ais-version-parity-') as directory:
            root = Path(directory)
            def build(name,version,value=42,extra='unchanged'):
                path = root/name
                workbook = xlsxwriter.Workbook(path)
                workbook.set_custom_property('Application version',version)
                workbook.set_custom_property('Other metadata',extra)
                sheet = workbook.add_worksheet('Plain')
                sheet.write_number(0,0,value)
                workbook.close()
                return path
            first = build('first.xlsx','1.5.5')
            current = build('current.xlsx','1.5.6')
            def digest(path,version=None):
                return semantic_package_fingerprint(path,application_version=version)['sha256']
            self.assertNotEqual(digest(first),digest(current))
            self.assertEqual(digest(first,'1.5.5'),digest(current,'1.5.6'))
            with self.assertRaisesRegex(ValueError,'explicit comparison declaration'):
                digest(current,'1.5.5')
            wrong_number = build('wrong-number.xlsx','1.5.6',43)
            self.assertNotEqual(digest(first,'1.5.5'),digest(wrong_number,'1.5.6'))
            wrong_metadata = build('wrong-metadata.xlsx','1.5.6',extra='changed')
            self.assertNotEqual(digest(first,'1.5.5'),digest(wrong_metadata,'1.5.6'))


if __name__=='__main__':
    unittest.main()
