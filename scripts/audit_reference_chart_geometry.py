"""Read-only OOXML/native/pixel evidence, with explicit artifact roles.

Never imports the producer or changes input XLSX/PNG. Metadata equality is
separate from pixel equality and is never inferred from an exported PNG/hash.
Column widths are font-dependent, not universal pixels (Microsoft Column API:
https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.column).
We use the observed native grid, not an assumed font digit width, to resolve
twoCellAnchor coordinates. All offsets are preserved as exact integers.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import posixpath
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {
    's': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
}
EMU_PER_POINT = 12700


def local(tag):
    return tag.rsplit('}', 1)[-1]


def xml_data(node):
    if node is None:
        return None
    result = {'tag': local(node.tag), 'attributes': dict(node.attrib)}
    if node.text and node.text.strip():
        result['text'] = node.text
    if len(node):
        result['children'] = [xml_data(child) for child in node]
    return result


def sha256(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def relationship_part(part):
    return posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')


def relationships(archive, part):
    rel = relationship_part(part)
    if rel not in archive.namelist():
        return {}
    result = {}
    for node in ET.fromstring(archive.read(rel)):
        if node.get('TargetMode') == 'External':
            continue
        target = node.get('Target', '')
        resolved = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(part), target))
        result[node.get('Id')] = {'target': resolved, 'type': node.get('Type')}
    return result


def normal_style(archive):
    root = ET.fromstring(archive.read('xl/styles.xml'))
    styles = root.findall('s:cellStyles/s:cellStyle', NS)
    normal = next((item for item in styles if item.get('builtinId') == '0'), None)
    if normal is None:
        return {'status': 'UNVERIFIED', 'reason': 'No builtin Normal cellStyle'}
    xf = root.findall('s:cellStyleXfs/s:xf', NS)[int(normal.get('xfId'))]
    font_id = int(xf.get('fontId', '0'))
    font = root.findall('s:fonts/s:font', NS)[font_id]
    return {'cell_style': dict(normal.attrib), 'style_xf': dict(xf.attrib), 'font_id': font_id, 'font': xml_data(font)}


def sheet_grid(archive, part, max_row):
    """Read only the grid prefix needed by every chart anchor, not AIS detail."""
    result = {'part': part, 'sheet_format': None, 'columns': [], 'rows': [], 'row_scope': f'1..{max_row}'}
    with archive.open(part) as stream:
        for event, node in ET.iterparse(stream, events=('start', 'end')):
            tag = local(node.tag)
            if event == 'start' and tag == 'row' and int(node.get('r', '0')) > max_row:
                break
            if event == 'start' and tag == 'sheetFormatPr':
                result['sheet_format'] = dict(node.attrib)
            elif event == 'start' and tag == 'col':
                result['columns'].append(dict(node.attrib))
            elif event == 'start' and tag == 'row':
                result['rows'].append(dict(node.attrib))
            if event == 'end':
                node.clear()
    return result


def marker(anchor, name):
    node = anchor.find('xdr:' + name, NS)
    return None if node is None else {local(child.tag): int(child.text) for child in node}


def chart_metadata(archive, part):
    root = ET.fromstring(archive.read(part))
    chart = root.find('c:chart', NS)
    plot = chart.find('c:plotArea', NS)
    series = []
    for group in plot:
        for node in group.findall('c:ser', NS):
            series.append({'chart_kind': local(group.tag), 'order': xml_data(node.find('c:order', NS)),
                           'name': xml_data(node.find('c:tx', NS)), 'shape': xml_data(node.find('c:spPr', NS)),
                           'marker': xml_data(node.find('c:marker', NS)),
                           'category_formulas': [item.text for item in node.findall('.//c:cat//c:f', NS)],
                           'value_formulas': [item.text for item in node.findall('.//c:val//c:f', NS)],
                           'point_shapes': [xml_data(item) for item in node.findall('c:dPt', NS)]})
    group_settings = [
        {'chart_kind': local(group.tag), **{
            name: xml_data(group.find('c:' + name, NS))
            for name in ('grouping', 'barDir', 'gapWidth', 'overlap')
        }}
        for group in plot if local(group.tag).endswith('Chart')
    ]
    return {'part': part, 'group_settings': group_settings,
            'chart_area_shape': xml_data(root.find('c:spPr', NS)), 'chart_area_text': xml_data(root.find('c:txPr', NS)),
            'title': xml_data(chart.find('c:title', NS)), 'legend': xml_data(chart.find('c:legend', NS)),
            'plot_layout': xml_data(plot.find('c:layout', NS)), 'plot_shape': xml_data(plot.find('c:spPr', NS)),
            'axes': [xml_data(item) for item in plot if local(item.tag).endswith('Ax')], 'series': series}


def extract_workbook(path, role):
    before = sha256(path)
    result = {'path': str(path.resolve()), 'role': role, 'sha256_before': before, 'size_bytes': path.stat().st_size, 'sheets': []}
    with zipfile.ZipFile(path) as archive:
        result['normal_style'] = normal_style(archive)
        workbook = ET.fromstring(archive.read('xl/workbook.xml'))
        wb_rels = relationships(archive, 'xl/workbook.xml')
        for sheet in workbook.findall('s:sheets/s:sheet', NS):
            part = wb_rels[sheet.get('{' + NS['r'] + '}id')]['target']
            drawing_targets = [rel['target'] for rel in relationships(archive, part).values() if rel['type'].endswith('/drawing')]
            if not drawing_targets:
                continue
            charts = []
            for drawing_part in drawing_targets:
                drawing = ET.fromstring(archive.read(drawing_part))
                draw_rels = relationships(archive, drawing_part)
                for anchor in drawing:
                    chart = anchor.find('.//c:chart', NS)
                    if chart is None:
                        continue
                    chart_part = draw_rels[chart.get('{' + NS['r'] + '}id')]['target']
                    frame = anchor.find('xdr:graphicFrame', NS)
                    charts.append({'drawing_part': drawing_part, 'anchor_kind': local(anchor.tag), 'anchor_attributes': dict(anchor.attrib),
                                   'from': marker(anchor, 'from'), 'to': marker(anchor, 'to'),
                                   'absolute_position': xml_data(anchor.find('xdr:pos', NS)), 'extent': xml_data(anchor.find('xdr:ext', NS)),
                                   'frame_transform': xml_data(frame.find('xdr:xfrm', NS)) if frame is not None else None,
                                   'ooxml': chart_metadata(archive, chart_part)})
            last_row = max((item.get('to') or item.get('from') or {}).get('row', 0) for item in charts) + 2
            grid = sheet_grid(archive, part, last_row)
            grid['name'] = sheet.get('name'); grid['charts'] = charts
            result['sheets'].append(grid)
    result['sha256_after'] = sha256(path)
    result['input_unchanged'] = result['sha256_before'] == result['sha256_after']
    if not result['input_unchanged']:
        raise RuntimeError('Input changed during read: ' + str(path))
    return result


def resolve_anchor(chart, native_grid):
    columns = {item['index']: item['metadata'] for item in native_grid['columns']}
    rows = {item['index']: item['metadata'] for item in native_grid['rows']}
    def position(mark):
        col = columns[mark['col'] + 1]; row = rows[mark['row'] + 1]
        return col['Left'] + mark['colOff'] / EMU_PER_POINT, row['Top'] + mark['rowOff'] / EMU_PER_POINT
    start = position(chart['from']); end = position(chart['to'])
    return dict(zip(('Left', 'Top', 'Width', 'Height'), (start[0], start[1], end[0] - start[0], end[1] - start[1])))


def attach_native(books, report_paths):
    for report_path in report_paths:
        report = json.loads(report_path.read_text(encoding='utf-8-sig'))
        index = {os.path.normcase(str(Path(item['path']).resolve())): item for item in report['workbooks']}
        for book in books:
            native = index.get(os.path.normcase(book['path']))
            if native is None:
                continue
            if native.get('sha256_before', '').lower() != book['sha256_before'].lower() or native.get('sha256_after', '').lower() != book['sha256_after'].lower():
                raise ValueError('Native report workbook hash differs: ' + book['path'])
            book.setdefault('native_reports', []).append({'path': str(report_path.resolve()), 'capture_scope': report.get('capture_scope', 'DETAILED_STYLE_WITH_ALL_POINT_STYLES'),
                                                        'read_only': native['opened_read_only'], 'file_unchanged': native['file_unchanged']})
            for sheet in book['sheets']:
                grid = next((item for item in native.get('sheet_geometry', []) if item['Name'] == sheet['name']), None)
                if grid:
                    sheet['native_grid'] = grid
                native_charts = [item for item in native['charts'] if item['sheet'] == sheet['name']]
                if len(native_charts) != len(sheet['charts']):
                    raise ValueError('Native/OOXML chart count mismatch: ' + book['path'])
                for ordinal, (chart, actual) in enumerate(zip(sheet['charts'], native_charts), 1):
                    keep = ('chart_object', 'chart', 'chart_groups', 'chart_area', 'plot_area', 'title', 'legend', 'axes', 'series_count', 'png', 'png_sha256', 'png_exported', 'metadata_errors', 'fatal_error')
                    chart['native'] = {key: actual[key] for key in keep if key in actual}
                    chart['native']['report'] = str(report_path.resolve()); chart['native']['chart_index'] = ordinal
                    if grid and chart['from'] and chart['to']:
                        resolved = resolve_anchor(chart, grid)
                        chart['anchor_resolved_points'] = resolved
                        chart['native_minus_anchor_points'] = {key: actual['chart_object'][key] - value for key, value in resolved.items()}


def validate_png(path):
    try:
        size = path.stat().st_size
        if size < 33:
            return {'status': 'UNVERIFIED_INVALID_PNG', 'error': 'Empty or truncated PNG', 'bytes': size}
        with path.open('rb') as stream:
            if stream.read(8) != bytes.fromhex('89504e470d0a1a0a'):
                return {'status': 'UNVERIFIED_INVALID_PNG', 'error': 'Invalid PNG signature', 'bytes': size}
        try:
            from PIL import Image
        except ImportError as error:
            return {'status': 'UNVERIFIED_DECODER_UNAVAILABLE', 'error': str(error), 'bytes': size}
        with Image.open(path) as picture:
            picture.verify()
            return {'status': 'VALID_PNG', 'size': list(picture.size), 'sha256': sha256(path)}
    except Exception as error:
        return {'status': 'UNVERIFIED_INVALID_PNG', 'error': str(error), 'bytes': path.stat().st_size if path.exists() else None}


def compare_png(left, right):
    validations = [validate_png(left), validate_png(right)]
    if any(item['status'] != 'VALID_PNG' for item in validations):
        return {'status': 'UNVERIFIED_INVALID_PNG', 'left': str(left), 'right': str(right), 'validations': validations}
    from PIL import Image, ImageChops, ImageStat
    with Image.open(left) as first, Image.open(right) as second:
        # Composite alpha onto Excel's white export background before RGB diff.
        first_rgba = first.convert('RGBA'); second_rgba = second.convert('RGBA')
        first = Image.alpha_composite(Image.new('RGBA', first.size, 'white'), first_rgba).convert('RGB')
        second = Image.alpha_composite(Image.new('RGBA', second.size, 'white'), second_rgba).convert('RGB')
        result = {'left': str(left), 'right': str(right), 'left_size': list(first.size), 'right_size': list(second.size), 'resized': False}
        if first.size != second.size:
            result.update(status='DIFFERENT_DIMENSIONS', exact_pixel_equal=False)
            return result
        difference = ImageChops.difference(first, second)
        data = difference.get_flattened_data() if hasattr(difference, 'get_flattened_data') else difference.getdata()
        changed = sum(1 for pixel in data if any(pixel))
        result.update(status='EXACT_EQUAL' if changed == 0 else 'DIFFERENT_PIXELS', exact_pixel_equal=changed == 0,
                      changed_pixels=changed, total_pixels=first.width * first.height, changed_fraction=changed / (first.width * first.height),
                      mean_absolute_rgb_difference=ImageStat.Stat(difference).mean, difference_bbox=difference.getbbox())
        return result


def assert_output_safe(output, roots, protected):
    output = output.resolve()
    for root in roots:
        if output == root.resolve() or output.is_relative_to(root.resolve()):
            raise ValueError('Output must be outside all workbook input roots')
    for path in protected:
        if output == path.resolve() or (output.exists() and os.path.samefile(output, path)):
            raise ValueError('Output aliases protected input')
    if output.exists():
        raise ValueError('New evidence file required; do not overwrite prior evidence')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--root', action='append', required=True, help='reference_full=DIR, reference_summary=DIR, generated_summary=DIR')
    parser.add_argument('--native-report', type=Path, action='append', default=[])
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    specs = [item.split('=', 1) for item in args.root]
    if any(len(item) != 2 or item[0] not in {'reference_full', 'reference_summary', 'generated_full', 'generated_summary', 'generated_mixed'} for item in specs):
        raise ValueError('Every --root must declare a valid explicit reference/generated artifact role')
    roots = [Path(path).resolve() for _, path in specs]
    files = [((root_role.replace('_mixed', '_summary' if '總表' in path.name else '_full')) + ':' + ('integrated' if '整合' in path.name else 'A' if '_2' in path.name else 'B'), path)
             for (root_role, _), root in zip(specs, roots) for path in sorted(root.glob('*.xlsx')) if not path.name.startswith('~$')]
    if not files:
        raise ValueError('No XLSX inputs')
    assert_output_safe(args.output, roots, [path for _, path in files] + args.native_report)
    books = [extract_workbook(path, role) for role, path in files]
    attach_native(books, args.native_report)
    for book in books:
        for sheet in book['sheets']:
            for chart in sheet['charts']:
                if 'native' in chart:
                    chart['native']['actual_png_validation'] = validate_png(Path(chart['native']['png']))
    pairs = []
    references = {book['role']: book for book in books if book['role'].startswith('reference_')}
    for book in books:
        if not book['role'].startswith('generated_'):
            continue
        ref = references.get(book['role'].replace('generated_', 'reference_', 1))
        if ref is None:
            continue
        if len(book['sheets']) != len(ref['sheets']):
            raise ValueError('Role pair sheet count differs')
        for sheet, ref_sheet in zip(book['sheets'], ref['sheets']):
            if len(sheet['charts']) != len(ref_sheet['charts']):
                raise ValueError('Role pair chart count differs')
            for index, (chart, ref_chart) in enumerate(zip(sheet['charts'], ref_sheet['charts']), 1):
                pair = {'generated_role': book['role'], 'reference_role': ref['role'], 'chart_index': index,
                        'ooxml_exact_anchor_equal': (chart['from'], chart['to']) == (ref_chart['from'], ref_chart['to'])}
                if 'native' in chart and 'native' in ref_chart:
                    pair['native_geometry_difference_points'] = {key: chart['native']['chart_object'][key] - ref_chart['native']['chart_object'][key]
                                                                for key in ('Left', 'Top', 'Width', 'Height')}
                    pair['png'] = compare_png(Path(chart['native']['png']), Path(ref_chart['native']['png']))
                else:
                    pair['png'] = {'status': 'UNVERIFIED_NATIVE_EXPORT_MISSING'}
                pairs.append(pair)
    report = {'schema': 'AIS_REFERENCE_CHART_ROLE_AUDIT_1', 'workbooks': books, 'role_comparisons': pairs,
              'scope': 'Explicit copied references and supplied generated replay only; not all 60 outputs or selector correctness',
              'verdict': 'OBSERVATION_ONLY_NO_REFERENCE_EXPECTED_MUTATION', 'all_input_hashes_unchanged': all(book['input_unchanged'] for book in books)}
    args.output.parent.mkdir(parents=True, exist_ok=True)
    # Exclusive create also protects against a late output-name collision.
    with args.output.open('x', encoding='utf-8') as stream:
        json.dump(report, stream, ensure_ascii=False, indent=2)
        stream.write('\n')
    print(json.dumps({'workbooks': len(books), 'charts': sum(len(sheet['charts']) for book in books for sheet in book['sheets']),
                      'comparisons': len(pairs), 'all_input_hashes_unchanged': report['all_input_hashes_unchanged'], 'output': str(args.output)}, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    raise SystemExit(main())
