#!/usr/bin/env python3
"""Generate canonical completed-box pages from an explicit workbook."""

from collections import OrderedDict
from datetime import date
from html import escape
from pathlib import Path
import argparse
import re
from string import Template
import sys
import xml.etree.ElementTree as ET
import zipfile


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

ACCENTS = {
    "orange": {"accent": "#9A4F00", "soft": "#FFF3E6"},
    "green": {"accent": "#08733F", "soft": "#EAF8F0"},
    "blue": {"accent": "#1956C8", "soft": "#EDF3FF"},
    "purple": {"accent": "#6B2AA6", "soft": "#F5EEFC"},
    "red": {"accent": "#B42318", "soft": "#FFF0EF"},
}


def _column_index(reference):
    letters = re.match(r"[A-Z]+", reference).group(0)
    index = 0
    for letter in letters:
        index = index * 26 + ord(letter) - 64
    return index - 1


def _shared_strings(archive):
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return [
        "".join(node.text or "" for node in item.findall(f".//{{{MAIN_NS}}}t"))
        for item in root.findall(f"{{{MAIN_NS}}}si")
    ]


def _cell_value(cell, shared_strings):
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(
            node.text or "" for node in cell.findall(f".//{{{MAIN_NS}}}t")
        )
    value_node = cell.find(f"{{{MAIN_NS}}}v")
    if value_node is None or value_node.text is None:
        return None
    value = value_node.text
    if cell_type == "s":
        return shared_strings[int(value)]
    if cell_type == "b":
        return value == "1"
    try:
        number = float(value)
        return int(number) if number.is_integer() else number
    except ValueError:
        return value


def _worksheet_rows(archive, target, shared_strings):
    target = target.lstrip("/")
    if not target.startswith("xl/"):
        target = f"xl/{target}"
    root = ET.fromstring(archive.read(target))
    rows = []
    for row in root.findall(f".//{{{MAIN_NS}}}row"):
        values = []
        for cell in row.findall(f"{{{MAIN_NS}}}c"):
            index = _column_index(cell.attrib["r"])
            while len(values) <= index:
                values.append(None)
            values[index] = _cell_value(cell, shared_strings)
        rows.append(values)
    return rows


def read_workbook(path):
    """Return worksheet rows keyed by worksheet name without modifying the file."""
    with zipfile.ZipFile(path) as archive:
        shared_strings = _shared_strings(archive)
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {
            relation.attrib["Id"]: relation.attrib["Target"]
            for relation in relationships.findall(f"{{{PACKAGE_REL_NS}}}Relationship")
        }
        sheets = OrderedDict()
        for sheet in workbook.findall(f".//{{{MAIN_NS}}}sheet"):
            relationship_id = sheet.attrib[f"{{{REL_NS}}}id"]
            sheets[sheet.attrib["name"]] = _worksheet_rows(
                archive, targets[relationship_id], shared_strings
            )
        return sheets


def _records(rows, required_columns):
    if not rows:
        raise ValueError("Worksheet is empty")
    headers = [str(value).strip() if value is not None else "" for value in rows[0]]
    missing = [column for column in required_columns if column not in headers]
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(missing)}")
    records = []
    for row in rows[1:]:
        padded = list(row) + [None] * (len(headers) - len(row))
        records.append(dict(zip(headers, padded)))
    return records


def _quantity_label(item, quantity):
    if quantity in (None, "", 1, 1.0):
        return str(item)
    quantity_text = str(int(quantity)) if isinstance(quantity, float) and quantity.is_integer() else str(quantity)
    return f"{quantity_text} × {item}"


def build_box_models(
    summary_rows,
    inventory_rows,
    verified_on,
    max_categories=12,
    box_ids=None,
):
    """Build completed-box models while preserving workbook row order."""
    summaries = _records(
        summary_rows,
        ["Box ID", "Title", "Color", "Status", "Fragility", "GitHub URL"],
    )
    inventory = _records(inventory_rows, ["Box ID", "Category", "Item", "Qty"])
    selected_ids = set(box_ids) if box_ids is not None else None
    boxes = []
    for summary in summaries:
        status = str(summary.get("Status") or "").strip()
        if status.casefold() != "100% full":
            continue
        box_id = str(summary["Box ID"]).strip()
        if selected_ids is not None and box_id not in selected_ids:
            continue
        color_name = str(summary["Color"]).strip()
        color = ACCENTS.get(color_name.casefold())
        if color is None:
            raise ValueError(f"No approved accent for color: {color_name}")
        categories = OrderedDict()
        for row in inventory:
            if str(row.get("Box ID") or "").strip() != box_id:
                continue
            category_name = str(row.get("Category") or "Uncategorized").strip()
            categories.setdefault(category_name, []).append(
                {
                    "name": str(row.get("Item") or "").strip(),
                    "label": _quantity_label(row.get("Item") or "", row.get("Qty")),
                }
            )
        category_models = [
            {"name": name, "count": len(items), "items": items}
            for name, items in categories.items()
        ]
        if len(category_models) > max_categories:
            raise ValueError(
                f"{box_id} has {len(category_models)} categories; "
                f"maximum is {max_categories}"
            )
        handling = str(summary.get("Fragility") or "").strip()
        boxes.append(
            {
                "id": box_id,
                "title": str(summary["Title"]).strip(),
                "color_name": color_name,
                "accent": color["accent"],
                "accent_soft": color["soft"],
                "status": status,
                "handling": handling if handling.casefold() not in {"", "no", "tbd"} else None,
                "url": str(summary.get("GitHub URL") or "").strip(),
                "categories": category_models,
                "item_count": sum(category["count"] for category in category_models),
                "category_count": len(category_models),
                "verified_on": verified_on,
            }
        )
    return boxes


def _format_date(value):
    parsed = date.fromisoformat(value)
    return f"{parsed.strftime('%B')} {parsed.day}, {parsed.year}"


def render_box_page(box, template_text):
    """Render a box model into the canonical semantic page template."""
    metadata = [
        f'<li>Status: {escape(box["status"])}</li>',
        f'<li>{box["item_count"]} inventory lines</li>',
        f'<li>{box["category_count"]} categories</li>',
    ]
    if box["handling"]:
        metadata.append(f'<li>Handling: {escape(box["handling"])}</li>')

    categories = []
    for index, category in enumerate(box["categories"], start=1):
        items = "\n".join(
            '<li class="inventory-item" '
            f'data-item-search="{escape(item["name"].casefold(), quote=True)}">'
            f'{escape(item["label"])}</li>'
            for item in category["items"]
        )
        categories.append(
            '<section class="inventory-category" '
            f'aria-labelledby="category-{index}" '
            f'data-category-search="{escape(category["name"].casefold(), quote=True)}">\n'
            f'  <h2 id="category-{index}" class="category-heading">'
            f'{escape(category["name"])} <span>{category["count"]} '
            f'{"item" if category["count"] == 1 else "items"}</span></h2>\n'
            f'  <ul class="inventory-list">\n{items}\n  </ul>\n'
            '</section>'
        )

    values = {
        "accent": escape(box["accent"], quote=True),
        "accent_soft": escape(box["accent_soft"], quote=True),
        "page_title": escape(f'{box["id"]} — {box["title"]} inventory'),
        "box_id": escape(box["id"]),
        "color_name": escape(box["color_name"]),
        "title": escape(box["title"]),
        "metadata_html": "\n      ".join(metadata),
        "item_count": str(box["item_count"]),
        "categories_html": "\n      ".join(categories),
        "verified_date": escape(_format_date(box["verified_on"])),
    }
    return Template(template_text).substitute(values)


def generate_pages(
    workbook_path,
    output_dir,
    template_path,
    verified_on,
    box_ids=None,
    max_categories=12,
):
    """Generate deterministic preview pages from an explicit workbook path."""
    workbook_path = Path(workbook_path)
    output_dir = Path(output_dir)
    template_path = Path(template_path)
    workbook = read_workbook(workbook_path)
    boxes = build_box_models(
        workbook["Boxes Summary"],
        workbook["Inventory"],
        verified_on,
        max_categories=max_categories,
        box_ids=box_ids,
    )
    if box_ids is not None:
        found_ids = {box["id"] for box in boxes}
        missing_ids = [box_id for box_id in box_ids if box_id not in found_ids]
        if missing_ids:
            raise ValueError(
                "Completed box not found: " + ", ".join(missing_ids)
            )

    template_text = template_path.read_text(encoding="utf-8")
    rendered_pages = [
        (box["id"], render_box_page(box, template_text)) for box in boxes
    ]
    output_dir.mkdir(parents=True, exist_ok=True)
    generated_paths = []
    for box_id, page in rendered_pages:
        output_path = output_dir / f"{box_id}.html"
        output_path.write_text(page, encoding="utf-8", newline="\n")
        generated_paths.append(output_path)
    return generated_paths


def create_parser():
    parser = argparse.ArgumentParser(
        description="Generate canonical box-page previews from a workbook."
    )
    parser.add_argument("--workbook", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument(
        "--template",
        type=Path,
        default=Path(__file__).resolve().parents[1] / "templates" / "box.html",
    )
    parser.add_argument("--verified-on", required=True)
    parser.add_argument("--box-id", action="append", dest="box_ids")
    parser.add_argument("--max-categories", type=int, default=12)
    return parser


def main(argv=None):
    parser = create_parser()
    args = parser.parse_args(argv)
    try:
        generated_paths = generate_pages(
            args.workbook,
            args.output,
            args.template,
            args.verified_on,
            box_ids=args.box_ids,
            max_categories=args.max_categories,
        )
    except (KeyError, OSError, ValueError, zipfile.BadZipFile) as error:
        parser.error(str(error))
    for path in generated_paths:
        print(path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
