from datetime import datetime
from io import StringIO
import json
import os
from pathlib import Path
import shutil
import tempfile
import threading
import unittest
import zipfile
from unittest import mock

import xlsxwriter
import ais_monthly_app as app
from ais_source_preflight import verify_source_identity, sha256


HEADERS = ['Year','Month','Day','Hour','Minute','Second','channel','msg_type','mmsi',
           'LONGITUDE_DESC','bearing','distance in nautical miles']


def write_source(path, observations, *, gps=False, reordered=False, missing=False, duplicate=False):
    """Disposable raw fixtures, never copies or alterations of research inputs."""
    headers = HEADERS.copy()
    if gps:
        headers[:6] = ['GPS_'+h.lower() for h in headers[:6]]
    if missing:
        headers[-1] = 'UNRECOGNIZED_DISTANCE_UNIT'
    order = list(reversed(range(len(headers)))) if reordered else list(range(len(headers)))
    workbook = xlsxwriter.Workbook(path)
    sheet = workbook.add_worksheet('AIS')
    sheet.write_row(0,0,[headers[i] for i in order]+(['mmsi'] if duplicate else []))
    for position,(timestamp,distance,mmsi) in enumerate(observations,1):
        values = [timestamp.year,timestamp.month,timestamp.day,timestamp.hour,timestamp.minute,timestamp.second,
                  'B',1,mmsi,'east',15.0,distance]
        sheet.write_row(position,0,[values[i] for i in order]+([mmsi] if duplicate else []))
    workbook.close()


class SourcePreflightTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='ais-preflight-test-')
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.input = self.root/'input'
        self.input.mkdir()
        self.file = self.input/'D&TMOK HWLN_20240201_23.xlsx'
        self.rows = [(datetime(2024,2,1,0,0,0),100.,111111111),
                     (datetime(2024,2,1,12,0,0),99.,111111111),
                     (datetime(2024,2,1,23,59,59),98.,222222222)]

    def inspect(self):
        return app.inspect_source_preflight(self.input)

    def test_standard_reordered_schema_is_name_mapped_not_index_mapped(self):
        write_source(self.file,self.rows,reordered=True)
        before = sha256(self.file)
        result = self.inspect()
        self.assertEqual(result['status'],'SAFE_WITH_WARNINGS')
        self.assertEqual(result['files'][0]['schema_profile'],'STANDARD_AIS')
        self.assertEqual(result['files'][0]['field_positions_1based']['distance'],1)
        self.assertEqual(result['rows_scanned'],3)
        self.assertEqual(result['files'][0]['normalized_core_rows'],3)
        self.assertEqual(result['calendar_days'],29)
        self.assertEqual(result['source_days'],1)
        self.assertEqual(sha256(self.file),before)

    def test_gps_known_but_never_authorized_alias(self):
        write_source(self.file,self.rows,gps=True)
        report = self.inspect()
        self.assertEqual(report['status'],'BLOCKED')
        self.assertIn('GPS_SCHEMA_DETECTED',report['files'][0]['issues'])
        self.assertEqual(report['files'][0]['rows_scanned'],3)
        self.assertEqual(report['files'][0]['time_basis'],'GPS_FIELDS_FORENSIC_ONLY_NOT_AUTHORIZED')
        self.assertEqual(report['policy']['gps_adapter'],'NOT_AUTHORIZED')

    def test_mixed_date_full_scan_blocks_whole_month_without_reassignment(self):
        rows = self.rows+[(datetime(2024,1,25,1),88.,333333333)]
        write_source(self.file,rows)
        report = self.inspect()
        self.assertEqual(report['rows_scanned'],4)
        self.assertEqual(report['mixed_date_rows'],1)
        self.assertEqual(report['blocked_months'],['HWLN_202402'])
        self.assertIn('2024-01-25',report['files'][0]['timestamp_dates'])
        self.assertFalse(report['policy']['mixed_date_reassignment'])

    def test_unknown_missing_or_duplicate_columns_fail_closed(self):
        for flags in ({'missing':True},{'duplicate':True}):
            with self.subTest(flags=flags):
                write_source(self.file,self.rows,**flags)
                report = self.inspect()
                self.assertEqual(report['status'],'BLOCKED')
                self.assertIn('SCHEMA_MISMATCH',report['files'][0]['issues'])
                self.assertEqual(report['files'][0]['row_coverage'],'UNVERIFIED')

    def test_subset_requires_explicit_exclusion_not_silent_union(self):
        short = self.input/'D&TMOK HWLN_20240201_4.xlsx'
        write_source(short,[self.rows[0],self.rows[0]])
        write_source(self.file,[self.rows[0],self.rows[0],self.rows[1]])
        original = {p:sha256(p) for p in self.input.iterdir()}
        report = self.inspect()
        pair = report['fragment_pairs'][0]
        self.assertIn('SUBSET_CANDIDATE',pair['classifications'])
        self.assertEqual(pair['exact_normalized_core_rows_shared'],2)
        self.assertEqual(pair['subset_path'],str(short.resolve()))
        self.assertTrue(pair['blocked'])
        self.assertFalse(pair['automatic_exclusion'])
        self.assertEqual(report['applicable_source_files'],2)
        self.assertEqual(len(report['blocked_files']),2)
        self.assertEqual({p:sha256(p) for p in self.input.iterdir()},original)

    def test_numeric_and_text_channel_identity_matches_production_reader(self):
        other = self.input/'D&TMOK HWLN_20240201_other.xlsx'
        for path,channel in ((self.file,1),(other,'1')):
            with xlsxwriter.Workbook(path) as workbook:
                sheet = workbook.add_worksheet('AIS')
                sheet.write_row(0,0,HEADERS)
                sheet.write_row(1,0,[2024,2,1,1,0,0,channel,1,111111111,'east',15.,50.])
        report = self.inspect()
        pair = report['fragment_pairs'][0]
        self.assertEqual(report['status'],'BLOCKED')
        self.assertIn('CORE_ROWS_EXACT_DUPLICATE',pair['classifications'])
        self.assertEqual(pair['exact_normalized_core_rows_shared'],1)

    def test_numeric_xml_integer_float_and_text_lexemes_preserved(self):
        from ais_source_preflight import _parse
        import io
        values = []
        class Archive:
            def open(self,target):
                return io.BytesIO(b'<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>1.0</v></c><c r="C1"><v>1e0</v></c><c r="D1" t="inlineStr"><is><t>1</t></is></c></row></sheetData></worksheet>')
        _parse(Archive(),'unused',[],lambda number,row:values.append(row))
        self.assertEqual([type(values[0][i]) for i in range(4)],[int,float,float,str])
        self.assertEqual([str(values[0][i]) for i in range(4)],['1','1.0','1.0','1'])

    def test_bad_numeric_xml_blocks_instead_of_coercing_channel_to_blank(self):
        write_source(self.file,self.rows)
        with zipfile.ZipFile(self.file) as archive:
            parts = {name:archive.read(name) for name in archive.namelist()}
        content = parts['xl/worksheets/sheet1.xml']
        import re
        content,count = re.subn(rb'<c r="G2"[^>]*>.*?</c>',b'<c r="G2"><v>not_a_number</v></c>',content,count=1)
        self.assertEqual(count,1)
        parts['xl/worksheets/sheet1.xml'] = content
        with zipfile.ZipFile(self.file,'w') as archive:
            for name,value in parts.items():
                archive.writestr(name,value)
        report = self.inspect()
        self.assertEqual(report['status'],'BLOCKED')
        self.assertIn('WORKBOOK_READ_FAILURE',report['files'][0]['issues'])

    def test_byte_identical_duplicate_is_disclosed_and_blocked(self):
        write_source(self.file,self.rows)
        shutil.copyfile(self.file,self.input/'D&TMOK HWLN_20240201_copy.xlsx')
        report = self.inspect()
        self.assertIn('EXACT_FILE_DUPLICATE',report['fragment_pairs'][0]['classifications'])
        self.assertEqual(report['status'],'BLOCKED')

    def test_same_second_different_messages_not_silently_deduplicated(self):
        write_source(self.file,self.rows)
        other = self.input/'D&TMOK HWLN_20240201_other.xlsx'
        write_source(other,[(d,v,999999999) for d,v,m in self.rows])
        report = self.inspect()
        pair = report['fragment_pairs'][0]
        self.assertFalse(pair['blocked'])
        self.assertEqual(pair['exact_normalized_core_rows_shared'],0)
        self.assertIn('HIGH_OCCUPIED_SECOND_OVERLAP_INDICATOR',pair['classifications'])
        self.assertEqual(report['applicable_source_files'],2)

    def test_complementary_fragments_and_partial_day_envelope(self):
        write_source(self.file,self.rows[1:])
        write_source(self.input/'D&TMOK HWLN_20240201_11.xlsx',self.rows[:1])
        report = self.inspect()
        self.assertIn('COMPLEMENTARY_FRAGMENTS',report['fragment_pairs'][0]['classifications'])
        self.assertNotEqual(report['status'],'BLOCKED')
        self.assertTrue(report['months'][0]['partial_day_files'])

    def test_corrupt_workbook_and_malformed_filename_disclosed(self):
        self.file.write_bytes(b'not an xlsx')
        report = self.inspect()
        self.assertIn('WORKBOOK_READ_FAILURE',report['files'][0]['issues'])
        self.file.rename(self.input/'D&TMOK HWLN_20240299_23.xlsx')
        report = self.inspect()
        self.assertEqual(report['discovery_issues'][0]['classification'],'UNPARSEABLE_AIS_FILENAME')
        self.assertEqual(report['status'],'BLOCKED')

    def test_excel_temp_not_formal_input_and_report_not_in_raw(self):
        write_source(self.file,self.rows)
        (self.input/'~$D&TMOK HWLN_20240201_23.xlsx').write_bytes(b'lock')
        report = self.inspect()
        self.assertEqual(report['physical_xlsx_count'],2)
        self.assertEqual(report['applicable_source_files'],1)
        with self.assertRaisesRegex(ValueError,'不可寫入'):
            app.save_source_preflight(report,self.input/'report')
        path = app.save_source_preflight(report,self.root/'report')
        self.assertTrue(path.is_file())
        self.assertTrue((path.parent/'SOURCE_PREFLIGHT_SUMMARY.md').is_file())

    def test_same_size_same_mtime_content_change_is_rejected(self):
        write_source(self.file,self.rows)
        report = self.inspect()
        stamp = self.file.stat()
        data = bytearray(self.file.read_bytes())
        data[-1] ^= 1
        self.file.write_bytes(data)
        os.utime(self.file,ns=(stamp.st_atime_ns,stamp.st_mtime_ns))
        with self.assertRaisesRegex(ValueError,'Source changed'):
            verify_source_identity(report)

    def test_inflated_dimension_is_rejected_before_production_reader(self):
        write_source(self.file,self.rows)
        with zipfile.ZipFile(self.file) as archive:
            parts = {name:archive.read(name) for name in archive.namelist()}
        parts['xl/worksheets/sheet1.xml'] = parts['xl/worksheets/sheet1.xml'].replace(b'ref="A1:L4"',b'ref="A1:L99"')
        with zipfile.ZipFile(self.file,'w') as archive:
            for name,data in parts.items():
                archive.writestr(name,data)
        report = self.inspect()
        self.assertIn('WORKSHEET_DIMENSION_MISMATCH',report['files'][0]['issues'])
        self.assertEqual(report['status'],'BLOCKED')

    def test_cancel_before_preflight_does_not_process_sources(self):
        write_source(self.file,self.rows)
        event = threading.Event()
        event.set()
        with self.assertRaises(app.CancelledError):
            app.inspect_source_preflight(self.input,cancel_event=event)

    def test_pipeline_rejection_creates_no_cache_or_official_xlsx(self):
        write_source(self.file,self.rows,gps=True)
        output = self.root/'output'/'modern.xlsx'
        config = app.AppConfig(self.input,output,'HWLN',2024,2)
        with self.assertRaises(app.SourcePreflightError):
            app.run_pipeline(config)
        self.assertFalse(list(output.parent.rglob('*.xlsx')))
        self.assertFalse(list(output.parent.rglob('*_cache')))
        self.assertEqual(len(list(output.parent.rglob('SOURCE_PREFLIGHT_REPORT.json'))),1)


class CliFailureTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='ais-cli-failure-test-')
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def test_source_exception_is_nonzero_and_persisted_with_shared_classification(self):
        error = app.SourceFileError('fixture source read failed')
        stderr = StringIO()
        with mock.patch.object(app,'_main',side_effect=error),mock.patch.object(app.sys,'stderr',stderr):
            code = app.main(['--output',str(self.root/'modern.xlsx')])
        self.assertEqual(code,2)
        self.assertIn('SOURCE_FILE_ERROR',stderr.getvalue())
        payload = json.loads(next(self.root.glob('AIS32_CLI_FAILURE_*.json')).read_text(encoding='utf-8'))
        self.assertEqual(payload['status'],'FAIL')
        self.assertEqual(payload['category'],app.error_category(error))
        self.assertEqual(payload['reason'],app.friendly_error_message(error))
        self.assertFalse(list(self.root.glob('*.xlsx')))

    def test_windowed_stdio_none_does_not_rethrow_or_open_dialog(self):
        with mock.patch.object(app,'_main',side_effect=app.SourceFileError('read failed')), \
             mock.patch.object(app.sys,'stderr',None),mock.patch.object(app.sys,'stdout',None):
            code = app.main(['--output',str(self.root/'modern.xlsx')])
        self.assertEqual(code,2)
        self.assertEqual(len(list(self.root.glob('AIS32_CLI_FAILURE_*.json'))),1)

    def test_logging_failure_cannot_rethrow_original_failure(self):
        with mock.patch.object(app,'_main',side_effect=app.SourceFileError('read failed')), \
             mock.patch.object(Path,'mkdir',side_effect=PermissionError('no log directory')), \
             mock.patch.object(app.sys,'stderr',StringIO()):
            self.assertEqual(app.main([]),2)


if __name__=='__main__':
    unittest.main()
