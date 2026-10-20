from __future__ import annotations

import argparse
import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


def _value(element: ET.Element | None, attribute: str = "val") -> str | None:
    return None if element is None else element.get(attribute)


def _colors(parent: ET.Element | None) -> list[dict[str, object]]:
    if parent is None:
        return []
    result: list[dict[str, object]] = []
    for element in parent.findall(".//a:srgbClr", NS) + parent.findall(".//a:schemeClr", NS):
        result.append(
            {
                "kind": element.tag.rsplit("}", 1)[-1],
                "value": element.get("val"),
                "alpha": [child.get("val") for child in element.findall("a:alpha", NS)],
                "satMod": [child.get("val") for child in element.findall("a:satMod", NS)],
                "lumMod": [child.get("val") for child in element.findall("a:lumMod", NS)],
                "lumOff": [child.get("val") for child in element.findall("a:lumOff", NS)],
            }
        )
    return result


def inspect_chart(name: str, payload: bytes) -> dict[str, object]:
    root = ET.fromstring(payload)
    plot = root.find(".//c:plotArea", NS)
    chart_kinds = []
    if plot is not None:
        chart_kinds = [child.tag.rsplit("}", 1)[-1] for child in plot if child.tag.startswith("{" + NS["c"] + "}")]
    series = []
    for item in root.findall(".//c:ser", NS):
        title = item.find("c:tx/c:v", NS)
        line = item.find("c:spPr/a:ln", NS)
        fill = item.find("c:spPr", NS)
        marker = item.find("c:marker/c:symbol", NS)
        series.append(
            {
                "title": None if title is None else title.text,
                "line_width": None if line is None else line.get("w"),
                "line_dash": _value(None if line is None else line.find("a:prstDash", NS)),
                "line_colors": _colors(line),
                "fill_colors": _colors(fill),
                "marker": _value(marker),
                "category_formula": next(
                    (node.text for node in item.findall(".//c:cat//c:f", NS) if node.text), None
                ),
                "value_formula": next(
                    (node.text for node in item.findall(".//c:val//c:f", NS) if node.text), None
                ),
                "point_count": _value(item.find(".//c:val//c:ptCount", NS)),
            }
        )
    return {
        "part": name,
        "chart_kinds": chart_kinds,
        "style": _value(root.find("c:style", NS)),
        "radar_style": _value(root.find(".//c:radarStyle", NS)),
        "legend_position": _value(root.find(".//c:legendPos", NS)),
        "display_blanks_as": _value(root.find(".//c:dispBlanksAs", NS)),
        "plot_area_colors": _colors(plot),
        "chart_area_colors": _colors(root.find(".//c:chartSpace/c:spPr", NS)),
        "series": series,
    }


def inspect_workbook(path: Path) -> dict[str, object]:
    with zipfile.ZipFile(path) as archive:
        charts = [
            inspect_chart(name, archive.read(name))
            for name in sorted(archive.namelist())
            if name.startswith("xl/charts/chart") and name.endswith(".xml")
        ]
    return {"path": str(path), "size": path.stat().st_size, "charts": charts}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("paths", nargs="+", type=Path)
    parser.add_argument("--output", type=Path)
    arguments = parser.parse_args()
    payload = [inspect_workbook(path) for path in arguments.paths]
    encoded = json.dumps(payload, ensure_ascii=False, indent=2)
    if arguments.output:
        arguments.output.parent.mkdir(parents=True, exist_ok=True)
        arguments.output.write_text(encoded + "\n", encoding="utf-8")
    else:
        print(encoded)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
