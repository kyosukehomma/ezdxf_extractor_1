import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

import ezdxf
from openpyxl import load_workbook

import oudan_red_yellow_checker as app


PROJECT_ROOT = Path(__file__).resolve().parents[1]
LOCAL_CASE_DIR = PROJECT_ROOT / "input" / "used_260723_162452"
VERIFIED_OUTPUT_DIR = PROJECT_ROOT / "output" / "exec_260723_162452"


def text_token(text, x, y, layer, height=0.5):
    return {
        "text": text,
        "x": x,
        "y": y,
        "anchor_x": x,
        "anchor_y": y,
        "layer": layer,
        "height": height,
        "handle": None,
    }


class TableDetectionTests(unittest.TestCase):
    def test_detects_two_block_header_signature(self):
        texts = [
            text_token(word, index * 3.0, 100.0, "D-MTR-TXT")
            for index, word in enumerate(app.EARTHWORK_HEADER_SEQUENCE)
        ]

        tables = app.detect_earthwork_tables(texts)

        self.assertEqual(1, len(tables))
        self.assertEqual("D-MTR-TXT", tables[0]["layer"])

    def test_center_marker_match_is_exact(self):
        texts = [
            text_token("CL", 0.0, 0.0, "D-BMK"),
            text_token("C.L.", 0.0, 1.0, "D-BMK"),
            # 土工表の記号値であり、中心線アンカーではない。
            text_token("CL1-1", 0.0, 2.0, "D-BMK"),
            text_token("CL1-2", 0.0, 3.0, "D-BMK"),
            text_token("CL", 0.0, 4.0, "D-MTR-TXT"),
        ]

        markers = app.extract_center_markers(texts)

        self.assertEqual(["CL", "C.L."], [marker["text"] for marker in markers])

    def test_rejects_data_group_far_from_its_only_table(self):
        tables = [{"x": 0.0, "y": 0.0}]
        data_groups = [[{"x": 100.0, "y": 0.0}]]

        with self.assertRaisesRegex(app.ExtractionError, "安全上限"):
            app.assign_data_groups_to_tables(tables, data_groups)


class StationResolutionTests(unittest.TestCase):
    def test_uses_single_station_assigned_to_each_table_when_cl_is_absent(self):
        tables = [
            {"x": 0.0, "y": 0.0},
            {"x": 100.0, "y": 0.0},
        ]
        stations = [
            {"text": "1", "plus": None, "x": 10.0, "y": 0.0},
            {"text": "2", "plus": None, "x": 90.0, "y": 0.0},
        ]

        resolved = app.resolve_table_stations(tables, stations, [])

        self.assertEqual(["1", "2"], [station["text"] for station in resolved])

    def test_uses_cl_before_direct_table_to_station_distance(self):
        tables = [{"x": 60.0, "y": 0.0}]
        center_markers = [{"text": "CL", "x": 0.0, "y": 0.0}]
        stations = [
            {"text": "10", "plus": Decimal("20"), "x": 0.0, "y": 5.0},
            {"text": "99", "plus": None, "x": 55.0, "y": 0.0},
        ]

        resolved = app.resolve_table_stations(tables, stations, center_markers)

        self.assertEqual("10", resolved[0]["text"])

    def test_rejects_multiple_local_stations_without_cl(self):
        tables = [{"x": 0.0, "y": 0.0}]
        stations = [
            {"text": "1", "plus": None, "x": 1.0, "y": 0.0},
            {"text": "2", "plus": None, "x": 2.0, "y": 0.0},
        ]

        with self.assertRaises(app.ExtractionError):
            app.resolve_table_stations(tables, stations, [])

    def test_rejects_nearly_equal_cl_station_candidates(self):
        tables = [{"x": 60.0, "y": 0.0}]
        center_markers = [{"text": "CL", "x": 0.0, "y": 0.0}]
        stations = [
            {"text": "1", "plus": None, "x": -5.0, "y": 0.0},
            {"text": "2", "plus": None, "x": 5.0, "y": 0.0},
        ]

        with self.assertRaisesRegex(app.ExtractionError, "測点候補"):
            app.resolve_table_stations(tables, stations, center_markers)

    def test_rejects_single_far_station_without_cl(self):
        tables = [{"x": 0.0, "y": 0.0}]
        stations = [
            {"text": "1", "plus": None, "x": 1000.0, "y": 0.0},
        ]

        with self.assertRaisesRegex(app.ExtractionError, "安全上限"):
            app.resolve_table_stations(tables, stations, [])


class VerifiedLocalRegressionTests(unittest.TestCase):
    good_dxf = LOCAL_CASE_DIR / "05横断図283～306.dxf"
    verified_workbook = (
        VERIFIED_OUTPUT_DIR / "01-01-03土工_05横断図283～306.xlsx"
    )
    problem_dxf = LOCAL_CASE_DIR / "05_本線横断図.dxf"

    def test_verified_good_dxf_matches_workbook_cell_values(self):
        if not self.good_dxf.exists() or not self.verified_workbook.exists():
            self.skipTest("ローカルの検証用DXFまたは確認済みExcelがありません。")

        doc = ezdxf.readfile(self.good_dxf)
        texts = app.collect_all_texts(doc.modelspace())
        self.assertEqual(25, len(app.detect_earthwork_tables(texts)))
        self.assertEqual(25, len(app.extract_no_texts(texts)))
        self.assertEqual(0, len(app.extract_center_markers(texts)))

        records = app.extract_output_records(texts)
        self.assertEqual(25, len(records))

        with tempfile.TemporaryDirectory() as temporary_directory:
            generated_path = Path(temporary_directory) / "generated.xlsx"
            app.write_output_workbook(
                PROJECT_ROOT / "template" / app.TEMPLATE_NAME,
                generated_path,
                records,
            )
            expected = load_workbook(self.verified_workbook, data_only=False)
            actual = load_workbook(generated_path, data_only=False)
            try:
                expected_sheet = expected[app.SHEET_NAME]
                actual_sheet = actual[app.SHEET_NAME]
                max_row = max(expected_sheet.max_row, actual_sheet.max_row)
                max_column = max(
                    expected_sheet.max_column,
                    actual_sheet.max_column,
                )
                mismatches = [
                    (
                        row,
                        column,
                        expected_sheet.cell(row, column).value,
                        actual_sheet.cell(row, column).value,
                    )
                    for row in range(1, max_row + 1)
                    for column in range(1, max_column + 1)
                    if expected_sheet.cell(row, column).value
                    != actual_sheet.cell(row, column).value
                ]
            finally:
                expected.close()
                actual.close()

        self.assertEqual([], mismatches)

    def test_problem_dxf_is_safely_rejected_in_phase_one(self):
        if not self.problem_dxf.exists():
            self.skipTest("ローカルの検証用DXFがありません。")

        doc = ezdxf.readfile(self.problem_dxf)
        texts = app.collect_all_texts(doc.modelspace())
        tables = app.detect_earthwork_tables(texts)
        stations = app.extract_no_texts(texts)
        center_markers = app.extract_center_markers(texts)
        self.assertEqual(77, len(tables))
        self.assertEqual(148, len(stations))
        self.assertEqual(141, len(center_markers))

        resolved_stations = app.resolve_table_stations(
            tables,
            stations,
            center_markers,
        )
        self.assertEqual(77, len(resolved_stations))
        self.assertEqual(
            77,
            len({station["handle"] for station in resolved_stations}),
        )
        first_low_table_index = min(
            range(len(tables)),
            key=lambda index: abs(tables[index]["y"] - 1260.654),
        )
        self.assertEqual(
            ("7", Decimal("40")),
            (
                resolved_stations[first_low_table_index]["text"],
                resolved_stations[first_low_table_index]["plus"],
            ),
        )

        with self.assertRaisesRegex(
            app.ExtractionError,
            "土工表77件にデータ群96件",
        ):
            app.extract_output_records(texts)


if __name__ == "__main__":
    unittest.main()
