import sys
import tempfile
import unittest
import zipfile
from contextlib import redirect_stderr
from io import StringIO
from pathlib import Path
from xml.sax.saxutils import escape


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import generate_site
from generate_site import build_box_models, read_workbook, render_box_page


def _column_name(index):
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def _sheet_xml(rows, shared_strings):
    xml_rows = []
    for row_number, row in enumerate(rows, start=1):
        cells = []
        for column_number, value in enumerate(row, start=1):
            if value is None:
                continue
            reference = f"{_column_name(column_number)}{row_number}"
            if isinstance(value, (int, float)):
                cells.append(f'<c r="{reference}"><v>{value}</v></c>')
            else:
                if value not in shared_strings:
                    shared_strings.append(value)
                index = shared_strings.index(value)
                cells.append(f'<c r="{reference}" t="s"><v>{index}</v></c>')
        xml_rows.append(f'<row r="{row_number}">{"".join(cells)}</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(xml_rows)}</sheetData></worksheet>'
    )


def make_workbook(path, summary_rows=None, inventory_rows=None):
    summary_rows = summary_rows or [
        ["Box ID", "Title", "Category", "Color", "Status", "Fragility", "GitHub URL", "Label Status", "Notes"],
        ["ZS-100", "Garage & Utility", "Tools", "Orange", "100% Full", "Handle with care", "https://example.test/ZS-100.html", "Ready", "Sealed"],
        ["ZS-101", "Open Box", "Misc", "Green", "In Progress", "No", "https://example.test/ZS-101.html", "Not generated", "Open"],
    ]
    inventory_rows = inventory_rows or [
        ["Box ID", "Category", "Item", "Qty", "Notes"],
        ["ZS-100", "Fasteners", "Bolt set", 2, None],
        ["ZS-100", "Measuring", "Tape measure", 1, None],
        ["ZS-100", "Fasteners", "Washer tin", 1, None],
        ["ZS-101", "Misc", "Loose item", 1, None],
    ]
    shared_strings = []
    summary_xml = _sheet_xml(summary_rows, shared_strings)
    inventory_xml = _sheet_xml(inventory_rows, shared_strings)
    shared_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'count="{len(shared_strings)}" uniqueCount="{len(shared_strings)}">'
        + "".join(f"<si><t>{escape(value)}</t></si>" for value in shared_strings)
        + "</sst>"
    )
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
            '</Types>',
        )
        archive.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>',
        )
        archive.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Boxes Summary" sheetId="1" r:id="rId1"/>'
            '<sheet name="Inventory" sheetId="2" r:id="rId2"/></sheets></workbook>',
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
            '</Relationships>',
        )
        archive.writestr("xl/worksheets/sheet1.xml", summary_xml)
        archive.writestr("xl/worksheets/sheet2.xml", inventory_xml)
        archive.writestr("xl/sharedStrings.xml", shared_xml)


class WorkbookModelTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.workbook_path = Path(self.temp_dir.name) / "inventory.xlsx"
        make_workbook(self.workbook_path)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_reads_named_worksheets_and_cell_types(self):
        workbook = read_workbook(self.workbook_path)

        self.assertEqual(workbook["Boxes Summary"][1][0], "ZS-100")
        self.assertEqual(workbook["Inventory"][1][3], 2)

    def test_builds_only_completed_boxes_in_source_order(self):
        workbook = read_workbook(self.workbook_path)

        boxes = build_box_models(
            workbook["Boxes Summary"], workbook["Inventory"], "2026-07-02"
        )

        self.assertEqual([box["id"] for box in boxes], ["ZS-100"])
        self.assertEqual(boxes[0]["verified_on"], "2026-07-02")

    def test_preserves_first_seen_category_and_item_order(self):
        workbook = read_workbook(self.workbook_path)

        box = build_box_models(
            workbook["Boxes Summary"], workbook["Inventory"], "2026-07-02"
        )[0]

        self.assertEqual(
            [category["name"] for category in box["categories"]],
            ["Fasteners", "Measuring"],
        )
        self.assertEqual(
            [item["label"] for item in box["categories"][0]["items"]],
            ["2 × Bolt set", "Washer tin"],
        )
        self.assertEqual(box["item_count"], 3)
        self.assertEqual(box["category_count"], 2)

    def test_rejects_completed_box_with_excessive_categories(self):
        workbook = read_workbook(self.workbook_path)
        inventory_rows = [workbook["Inventory"][0]] + [
            ["ZS-100", f"Category {index}", f"Item {index}", 1, None]
            for index in range(1, 14)
        ]

        with self.assertRaisesRegex(
            ValueError,
            r"ZS-100 has 13 categories; maximum is 12",
        ):
            build_box_models(
                workbook["Boxes Summary"], inventory_rows, "2026-07-02"
            )

    def test_renders_semantic_searchable_inventory_with_escaped_content(self):
        workbook = read_workbook(self.workbook_path)
        box = build_box_models(
            workbook["Boxes Summary"], workbook["Inventory"], "2026-07-02"
        )[0]
        template = (ROOT / "templates" / "box.html").read_text(encoding="utf-8")

        page = render_box_page(box, template)

        self.assertIn("<main", page)
        self.assertIn("Garage &amp; Utility", page)
        self.assertIn('aria-labelledby="category-1"', page)
        self.assertIn('data-category-search="fasteners"', page)
        self.assertIn('data-item-search="bolt set"', page)
        self.assertEqual(page.count('<li class="inventory-item"'), 3)
        self.assertIn('id="box-search"', page)
        self.assertIn('aria-live="polite"', page)
        self.assertIn("3 items shown", page)
        self.assertIn("No matching items in this box", page)
        self.assertIn("Verified against the sealed-box inventory on July 2, 2026.", page)
        self.assertIn('href="../assets/box.css"', page)
        self.assertIn('src="../assets/box-search.js"', page)


class GeneratorCliTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.workbook_path = self.root / "inventory.xlsx"
        make_workbook(self.workbook_path)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_parser_requires_explicit_workbook_argument(self):
        self.assertTrue(hasattr(generate_site, "create_parser"))
        parser = generate_site.create_parser()

        with redirect_stderr(StringIO()), self.assertRaises(SystemExit) as raised:
            parser.parse_args(
                [
                    "--output",
                    str(self.root / "preview"),
                    "--verified-on",
                    "2026-07-02",
                ]
            )

        self.assertEqual(raised.exception.code, 2)

    def test_generates_only_selected_box_deterministically(self):
        self.assertTrue(hasattr(generate_site, "generate_pages"))
        template_path = ROOT / "templates" / "box.html"
        first_output = self.root / "preview-one"
        second_output = self.root / "preview-two"

        first_paths = generate_site.generate_pages(
            self.workbook_path,
            first_output,
            template_path,
            "2026-07-02",
            box_ids=["ZS-100"],
        )
        second_paths = generate_site.generate_pages(
            self.workbook_path,
            second_output,
            template_path,
            "2026-07-02",
            box_ids=["ZS-100"],
        )

        self.assertEqual([path.name for path in first_paths], ["ZS-100.html"])
        self.assertEqual(
            (first_output / "ZS-100.html").read_bytes(),
            (second_output / "ZS-100.html").read_bytes(),
        )
        self.assertEqual(
            [path.name for path in first_output.iterdir()], ["ZS-100.html"]
        )

    def test_excessive_categories_block_output(self):
        self.assertTrue(hasattr(generate_site, "generate_pages"))
        summary_rows = [
            [
                "Box ID",
                "Title",
                "Color",
                "Status",
                "Fragility",
                "GitHub URL",
            ],
            ["ZS-200", "Fragmented", "Blue", "100% Full", "No", ""],
        ]
        inventory_rows = [["Box ID", "Category", "Item", "Qty"]] + [
            ["ZS-200", f"Category {index}", f"Item {index}", 1]
            for index in range(1, 14)
        ]
        make_workbook(self.workbook_path, summary_rows, inventory_rows)
        output = self.root / "blocked-preview"

        with self.assertRaisesRegex(
            ValueError,
            r"ZS-200 has 13 categories; maximum is 12",
        ):
            generate_site.generate_pages(
                self.workbook_path,
                output,
                ROOT / "templates" / "box.html",
                "2026-07-02",
                box_ids=["ZS-200"],
            )

        self.assertFalse(output.exists())


if __name__ == "__main__":
    unittest.main()
