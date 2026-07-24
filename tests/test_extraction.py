import hashlib
import json
import shutil
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

BAD_EXPECTED_STATS = {
    "オープン掘削": (17, Decimal("838.0"), 60),
    "路体盛土(W＜2.5m)": (3, Decimal("0.8"), 74),
    "路体盛土(2.5m≦W＜4.0m)": (2, Decimal("1.0"), 75),
    "路体盛土(4.0m≦W)": (56, Decimal("11944.3"), 21),
    "路床盛土(W＜2.5m)": (1, Decimal("0.2"), 76),
    "路床盛土(2.5m≦W＜4.0m)": (1, Decimal("0.7"), 76),
    "路床盛土(4.0m≦W)": (58, Decimal("1391.4"), 19),
    "路肩盛土": (61, Decimal("98.2"), 16),
    "路体外盛土(W＜2.5m)": (11, Decimal("6.7"), 66),
    "路体外盛土(2.5m≦W＜4.0m)": (1, Decimal("1.9"), 76),
    "路体外盛土(4.0m≦W)": (11, Decimal("38.0"), 66),
    "畦畔盛土": (23, Decimal("8.6"), 54),
    "切土法面整形(左)": (13, Decimal("33.4"), 64),
    "切土法面整形(右)": (13, Decimal("44.9"), 64),
    "盛土法面整形(左)": (47, Decimal("511.3"), 30),
    "盛土法面整形(右)": (47, Decimal("502.7"), 30),
}
MAINLINE_WORK_EXPECTED_STATS = {
    "床掘": (55, Decimal("125.7"), 22),
    "埋戻(C)": (11, Decimal("34.9"), 66),
    "埋戻(D)": (46, Decimal("45.5"), 31),
}
MAINLINE_EXPECTED_DIGEST = (
    "7297e6e9d4664d47c0437e72a907576"
    "b65f7db0e35d691d8e5d6479ff4a10d74"
)
SYMBOL_TABLE_EXPECTED_DIGEST = (
    "81a228956d3a08c5001493d3e9ff5e79"
    "827d78f6fe2ae9de9216928dec00fbf9"
)


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

    def test_detects_symbol_quantity_header_with_small_y_offsets(self):
        x_positions = (
            0.0,
            2.54,
            6.54,
            10.48,
            13.12,
            15.62,
            19.45,
            22.54,
        )
        texts = [
            text_token(
                word,
                x,
                100.0 + (0.03 if index % 2 else 0.0),
                "D-MTR-TXT",
            )
            for index, (word, x) in enumerate(
                zip(app.SYMBOL_TABLE_HEADER_SEQUENCE, x_positions)
            )
        ]

        tables = app.detect_earthwork_tables(texts)

        self.assertEqual(1, len(tables))
        self.assertEqual(app.SYMBOL_TABLE_FORMAT, tables[0]["format"])
        self.assertEqual(app.SYMBOL_TABLE_LINE_LAYER, tables[0]["line_layer"])

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

    def test_combines_upper_and_lower_roadbed_quantities(self):
        self.assertEqual(
            Decimal("25.2"),
            app._combine_table_quantities(
                Decimal("7.1"),
                Decimal("18.1"),
            ),
        )
        self.assertEqual("-", app._combine_table_quantities("-", "-"))

    def test_rejects_an_unexpected_hierarchical_source_schema(self):
        damaged_schema = list(app.HIERARCHICAL_SOURCE_SCHEMA)
        damaged_schema[9] = ("上部路床", "W≧4.0m")

        with self.assertRaisesRegex(app.ExtractionError, "種別・区分構成"):
            app._validate_hierarchical_source_schema(damaged_schema)


class StationResolutionTests(unittest.TestCase):
    def test_infers_20_or_100_meter_station_interval(self):
        standard = [
            {"測点": Decimal("283"), "追加距離": None},
            {"測点": Decimal("284"), "追加距離": None},
        ]
        hundred_meter = [
            {"測点": Decimal("7"), "追加距離": Decimal("80")},
            {"測点": Decimal("8"), "追加距離": None},
        ]

        self.assertEqual(
            Decimal("20"),
            app.infer_station_interval(standard),
        )
        self.assertEqual(
            Decimal("100"),
            app.infer_station_interval(hundred_meter),
        )

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
    symbol_table_dxf = PROJECT_ROOT / "input" / "004_横断図240514.dxf"

    def _first_good_table_texts(self):
        if not self.good_dxf.exists():
            self.skipTest("ローカルの検証用DXFがありません。")

        doc = ezdxf.readfile(self.good_dxf)
        modelspace = doc.modelspace()
        texts = app.collect_all_texts(modelspace)
        lines = app.collect_all_lines(modelspace)
        tables = app.detect_earthwork_tables(texts)
        regions = app.detect_table_regions(tables, lines, texts)
        return app._texts_in_table_region(
            texts,
            tables[0],
            regions[0],
        )

    def test_legacy_table_rejects_a_missing_unit_row(self):
        table_texts = self._first_good_table_texts()
        unit = next(
            text for text in table_texts
            if app.UNIT_PATTERN.fullmatch(text["text"])
        )
        damaged_texts = [
            text for text in table_texts
            if text is not unit
        ]

        with self.assertRaises(app.ExtractionError):
            app._build_legacy_table_group(damaged_texts)

    def test_legacy_table_rejects_a_missing_quantity(self):
        table_texts = self._first_good_table_texts()
        quantities = [
            text for text in table_texts
            if text["text"] == "9.2"
        ]
        self.assertEqual(1, len(quantities))
        damaged_texts = [
            text for text in table_texts
            if text is not quantities[0]
        ]

        with self.assertRaisesRegex(app.ExtractionError, "数値化"):
            app._build_legacy_table_group(damaged_texts)

    def test_verified_good_dxf_matches_workbook_cell_values(self):
        if not self.good_dxf.exists() or not self.verified_workbook.exists():
            self.skipTest("ローカルの検証用DXFまたは確認済みExcelがありません。")

        doc = ezdxf.readfile(self.good_dxf)
        modelspace = doc.modelspace()
        texts = app.collect_all_texts(modelspace)
        lines = app.collect_all_lines(modelspace)
        self.assertEqual(25, len(app.detect_earthwork_tables(texts)))
        self.assertEqual(25, len(app.extract_no_texts(texts)))
        self.assertEqual(0, len(app.extract_center_markers(texts)))

        tables = app.detect_earthwork_tables(texts)
        regions = app.detect_table_regions(tables, lines, texts)
        self.assertEqual({37}, {region["line_count"] for region in regions})
        self.assertEqual({1}, {region["attached_count"] for region in regions})

        records = app.extract_output_records(texts, lines)
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

    def test_mainline_dxf_produces_77_complete_records(self):
        if not self.problem_dxf.exists():
            self.skipTest("ローカルの検証用DXFがありません。")

        doc = ezdxf.readfile(self.problem_dxf)
        modelspace = doc.modelspace()
        texts = app.collect_all_texts(modelspace)
        lines = app.collect_all_lines(modelspace)
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

        regions = app.detect_table_regions(tables, lines, texts)
        self.assertEqual(77, len(regions))
        self.assertEqual({45}, {region["line_count"] for region in regions})
        self.assertEqual({0}, {region["attached_count"] for region in regions})

        records = app.extract_output_records(texts, lines)
        self.assertEqual(77, len(records))
        self.assertEqual(
            77,
            len({(record["測点"], record["追加距離"]) for record in records}),
        )
        self.assertTrue(
            all(
                [data["種別"] for data in record["データ"]]
                == list(app.HIERARCHICAL_ALL_OUTPUT_KINDS)
                for record in records
            )
        )

        actual_stats = {}
        for kind in app.HIERARCHICAL_OUTPUT_KINDS:
            values = [
                data["数量"]
                for record in records
                for data in record["データ"]
                if data["種別"] == kind
            ]
            numeric = [
                value for value in values
                if isinstance(value, Decimal)
            ]
            actual_stats[kind] = (
                len(numeric),
                sum(numeric, Decimal("0")),
                values.count("-"),
            )
        self.assertEqual(BAD_EXPECTED_STATS, actual_stats)

        actual_work_stats = {}
        for kind in app.HIERARCHICAL_WORK_OUTPUT_KINDS:
            values = [
                data["数量"]
                for record in records
                for data in record["データ"]
                if data["種別"] == kind
            ]
            numeric = [
                value for value in values
                if isinstance(value, Decimal)
            ]
            actual_work_stats[kind] = (
                len(numeric),
                sum(numeric, Decimal("0")),
                values.count("-"),
            )
        self.assertEqual(
            MAINLINE_WORK_EXPECTED_STATS,
            actual_work_stats,
        )

        digest_payload = [
            [
                str(record["測点"]),
                (
                    ""
                    if record["追加距離"] is None
                    else str(record["追加距離"])
                ),
                [
                    [data["種別"], str(data["数量"])]
                    for data in record["データ"]
                ],
            ]
            for record in records
        ]
        digest = hashlib.sha256(
            json.dumps(
                digest_payload,
                ensure_ascii=False,
                separators=(",", ":"),
            ).encode()
        ).hexdigest()
        self.assertEqual(MAINLINE_EXPECTED_DIGEST, digest)

        sample = next(
            record for record in records
            if record["測点"] == Decimal("7")
            and record["追加距離"] == Decimal("40")
        )
        sample_values = {
            data["種別"]: data["数量"]
            for data in sample["データ"]
        }
        self.assertEqual(Decimal("236.3"), sample_values["オープン掘削"])
        self.assertEqual(Decimal("0.4"), sample_values["路肩盛土"])
        self.assertEqual(
            Decimal("10.7"),
            sample_values["切土法面整形(左)"],
        )
        self.assertEqual(
            Decimal("11.2"),
            sample_values["切土法面整形(右)"],
        )
        self.assertEqual(Decimal("2.7"), sample_values["床掘"])
        self.assertEqual("-", sample_values["埋戻(C)"])
        self.assertEqual(Decimal("2.1"), sample_values["埋戻(D)"])

        with tempfile.TemporaryDirectory() as temporary_directory:
            generated_path = (
                Path(temporary_directory)
                / "01-01-03土工_05_本線横断図.xlsx"
            )
            app.write_output_workbook(
                PROJECT_ROOT / "template" / app.TEMPLATE_NAME,
                generated_path,
                records,
            )
            workbook = load_workbook(generated_path, data_only=False)
            try:
                input_sheet = workbook[app.SHEET_NAME]
                self.assertEqual(100, input_sheet["E4"].value)
                self.assertEqual("=SUM(AA81:AB81)", input_sheet["AH81"].value)

                positions = [
                    (
                        record["測点"] * Decimal("100")
                        + (record["追加距離"] or Decimal("0"))
                    )
                    for record in records
                ]
                distances = [
                    current - previous
                    for previous, current in zip(
                        positions,
                        positions[1:],
                    )
                ]
                self.assertTrue(all(distance > 0 for distance in distances))
                self.assertEqual(Decimal("2400"), sum(distances))

                for sheet_name in (
                    "掘削",
                    "路体",
                    "路床",
                    "路体外",
                    "作業土工1",
                ):
                    sheet = workbook[sheet_name]
                    self.assertEqual("=入力!C81", sheet["C82"].value)
                    self.assertEqual(84, sheet.max_row)
                    self.assertIn("$84", sheet.print_area)

                slope_sheet = workbook["法面整形"]
                self.assertEqual("=入力!C81", slope_sheet["C82"].value)
                self.assertEqual("=I84+L84", slope_sheet["L85"].value)
                self.assertIn("$85", slope_sheet.print_area)

                roadbed_sheet = workbook["路床"]
                self.assertEqual("=入力!T81", roadbed_sheet["P82"].value)
                self.assertEqual("=R83", roadbed_sheet["R84"].value)

                work_sheet = workbook["作業土工1"]
                for index, record in enumerate(records, start=6):
                    quantities = {
                        data["種別"]: data["数量"]
                        for data in record["データ"]
                    }
                    for column, kind in (
                        (7, "床掘"),
                        (10, "埋戻(D)"),
                        (13, "埋戻(C)"),
                    ):
                        expected = quantities[kind]
                        actual = work_sheet.cell(index, column).value
                        if isinstance(expected, Decimal):
                            actual = Decimal(str(actual))
                        self.assertEqual(expected, actual)

                self.assertEqual(
                    "=路床!R84",
                    workbook["集計表"]["F17"].value,
                )
                self.assertEqual(
                    "=作業土工1!O84",
                    workbook["作業土工2"]["H5"].value,
                )

                ref_formulas = [
                    (sheet.title, cell.coordinate, cell.value)
                    for sheet in workbook.worksheets
                    for row in sheet.iter_rows()
                    for cell in row
                    if (
                        isinstance(cell.value, str)
                        and cell.value.startswith("=")
                        and "#REF!" in cell.value
                    )
                ]
                self.assertEqual([], ref_formulas)

                check_sheet = workbook["チェック"]
                self.assertEqual(
                    "=INDEX(入力!$AB$5:$AB$81,$C$2)",
                    check_sheet["D24"].value,
                )
                self.assertEqual(
                    "掘削!$C$6:$C$82",
                    workbook["掘削"].defined_names["_1B"].attr_text,
                )
            finally:
                workbook.close()

    def test_symbol_quantity_dxf_splits_five_alignments_and_blanks_right_slopes(self):
        if not self.symbol_table_dxf.exists():
            self.skipTest("ローカルの記号型検証用DXFがありません。")

        doc = ezdxf.readfile(self.symbol_table_dxf)
        modelspace = doc.modelspace()
        texts = app.collect_all_texts(modelspace)
        lines = app.collect_all_lines(modelspace)
        tables = app.detect_earthwork_tables(texts)
        self.assertEqual(103, len(tables))
        self.assertEqual(
            {app.SYMBOL_TABLE_FORMAT},
            {table["format"] for table in tables},
        )

        regions = app.detect_table_regions(tables, lines, texts)
        self.assertEqual({27}, {region["line_count"] for region in regions})
        self.assertEqual({0}, {region["attached_count"] for region in regions})

        record_sets = app.extract_output_record_sets(texts, lines)
        self.assertEqual(
            {
                "本線": 52,
                "Aランプ": 13,
                "Bランプ": 12,
                "Cランプ": 13,
                "Dランプ": 13,
            },
            {
                alignment: len(records)
                for alignment, records in record_sets.items()
            },
        )
        self.assertEqual(
            {
                "本線": (Decimal("214"), Decimal("265")),
                "Aランプ": (Decimal("0"), Decimal("12")),
                "Bランプ": (Decimal("1"), Decimal("12")),
                "Cランプ": (Decimal("0"), Decimal("12")),
                "Dランプ": (Decimal("1"), Decimal("13")),
            },
            {
                alignment: (
                    records[0]["測点"],
                    records[-1]["測点"],
                )
                for alignment, records in record_sets.items()
            },
        )

        for alignment, records in record_sets.items():
            self.assertTrue(
                all(record["路線"] == alignment for record in records)
            )
            self.assertTrue(
                all(
                    [data["種別"] for data in record["データ"]]
                    == list(app.HIERARCHICAL_OUTPUT_KINDS)
                    for record in records
                )
            )
            for record in records:
                quantities = {
                    data["種別"]: data["数量"]
                    for data in record["データ"]
                }
                self.assertIsNone(quantities["切土法面整形(右)"])
                self.assertIsNone(quantities["盛土法面整形(右)"])

        digest_payload = {
            alignment: [
                [
                    str(record["測点"]),
                    (
                        ""
                        if record["追加距離"] is None
                        else str(record["追加距離"])
                    ),
                    [
                        [
                            data["種別"],
                            (
                                ""
                                if data["数量"] is None
                                else str(data["数量"])
                            ),
                        ]
                        for data in record["データ"]
                    ],
                ]
                for record in records
            ]
            for alignment, records in record_sets.items()
        }
        digest = hashlib.sha256(
            json.dumps(
                digest_payload,
                ensure_ascii=False,
                separators=(",", ":"),
            ).encode()
        ).hexdigest()
        self.assertEqual(SYMBOL_TABLE_EXPECTED_DIGEST, digest)

        with tempfile.TemporaryDirectory() as temporary_directory:
            temporary_root = Path(temporary_directory)
            input_directory = temporary_root / "input"
            template_directory = temporary_root / "template"
            input_directory.mkdir()
            template_directory.mkdir()
            shutil.copy2(
                self.symbol_table_dxf,
                input_directory / self.symbol_table_dxf.name,
            )
            shutil.copy2(
                PROJECT_ROOT / "template" / app.TEMPLATE_NAME,
                template_directory / app.TEMPLATE_NAME,
            )

            original_get_base_path = app.get_base_path
            app.get_base_path = lambda: str(temporary_root)
            try:
                processed, failures = app.proc([])
            finally:
                app.get_base_path = original_get_base_path

            self.assertEqual([], failures)
            self.assertEqual([self.symbol_table_dxf.name], processed)
            output_directories = list(
                (temporary_root / "output").glob("exec_*")
            )
            self.assertEqual(1, len(output_directories))
            output_paths = {
                path.stem.rsplit("_", 1)[-1]: path
                for path in output_directories[0].glob("*.xlsx")
            }
            self.assertEqual(set(record_sets), set(output_paths))

            for alignment, records in record_sets.items():
                workbook = load_workbook(
                    output_paths[alignment],
                    data_only=False,
                )
                try:
                    input_sheet = workbook[app.SHEET_NAME]
                    self.assertEqual(20, input_sheet["E4"].value)
                    for index, record in enumerate(records, start=app.ROW_START):
                        quantities = {
                            data["種別"]: data["数量"]
                            for data in record["データ"]
                        }
                        cut_left = input_sheet.cell(index, 25).value
                        fill_left = input_sheet.cell(index, 27).value
                        if isinstance(
                            quantities["切土法面整形(左)"],
                            Decimal,
                        ):
                            cut_left = Decimal(str(cut_left))
                        if isinstance(
                            quantities["盛土法面整形(左)"],
                            Decimal,
                        ):
                            fill_left = Decimal(str(fill_left))
                        self.assertEqual(
                            quantities["切土法面整形(左)"],
                            cut_left,
                        )
                        self.assertIsNone(input_sheet.cell(index, 26).value)
                        self.assertEqual(
                            quantities["盛土法面整形(左)"],
                            fill_left,
                        )
                        self.assertIsNone(input_sheet.cell(index, 28).value)

                    ref_formulas = [
                        cell.value
                        for sheet in workbook.worksheets
                        for row in sheet.iter_rows()
                        for cell in row
                        if (
                            isinstance(cell.value, str)
                            and cell.value.startswith("=")
                            and "#REF!" in cell.value
                        )
                    ]
                    self.assertEqual([], ref_formulas)

                    if alignment == "本線":
                        work_sheet = workbook["作業土工1"]
                        self.assertTrue(
                            all(
                                work_sheet.cell(row, column).value is None
                                for row in range(6, 6 + len(records))
                                for column in (7, 10, 13)
                            )
                        )
                finally:
                    workbook.close()

            self.assertEqual(
                [self.symbol_table_dxf.name],
                [
                    path.name
                    for path in input_directory.glob("used_*/*.dxf")
                ],
            )

    def test_proc_creates_both_workbooks_and_archives_both_dxf_files(self):
        if (
            not self.good_dxf.exists()
            or not self.problem_dxf.exists()
            or not (PROJECT_ROOT / "template" / app.TEMPLATE_NAME).exists()
        ):
            self.skipTest("ローカルの検証用DXFまたはテンプレートがありません。")

        with tempfile.TemporaryDirectory() as temporary_directory:
            temporary_root = Path(temporary_directory)
            input_directory = temporary_root / "input"
            template_directory = temporary_root / "template"
            input_directory.mkdir()
            template_directory.mkdir()
            shutil.copy2(self.good_dxf, input_directory / self.good_dxf.name)
            shutil.copy2(
                self.problem_dxf,
                input_directory / self.problem_dxf.name,
            )
            shutil.copy2(
                PROJECT_ROOT / "template" / app.TEMPLATE_NAME,
                template_directory / app.TEMPLATE_NAME,
            )

            original_get_base_path = app.get_base_path
            app.get_base_path = lambda: str(temporary_root)
            try:
                processed, failures = app.proc([])
            finally:
                app.get_base_path = original_get_base_path

            self.assertEqual([], failures)
            self.assertEqual(
                {self.good_dxf.name, self.problem_dxf.name},
                set(processed),
            )

            output_directories = list(
                (temporary_root / "output").glob("exec_*")
            )
            self.assertEqual(1, len(output_directories))
            output_names = {
                path.name
                for path in output_directories[0].glob("*.xlsx")
            }
            self.assertEqual(
                {
                    "01-01-03土工_05横断図283～306.xlsx",
                    "01-01-03土工_05_本線横断図.xlsx",
                },
                output_names,
            )
            self.assertEqual([], list(input_directory.glob("*.dxf")))
            archived_names = {
                path.name
                for path in input_directory.glob("used_*/*.dxf")
            }
            self.assertEqual(
                {self.good_dxf.name, self.problem_dxf.name},
                archived_names,
            )

    def test_proc_keeps_processing_errors_in_the_input_folder(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            temporary_root = Path(temporary_directory)
            input_directory = temporary_root / "input"
            template_directory = temporary_root / "template"
            input_directory.mkdir()
            template_directory.mkdir()
            broken_dxf = input_directory / "broken.dxf"
            broken_dxf.write_text("DXFではありません", encoding="utf-8")
            shutil.copy2(
                PROJECT_ROOT / "template" / app.TEMPLATE_NAME,
                template_directory / app.TEMPLATE_NAME,
            )

            original_get_base_path = app.get_base_path
            app.get_base_path = lambda: str(temporary_root)
            try:
                processed, failures = app.proc([])
            finally:
                app.get_base_path = original_get_base_path

            self.assertEqual([], processed)
            self.assertEqual(1, len(failures))
            self.assertIn("broken.dxf", failures[0])
            self.assertTrue(broken_dxf.exists())


if __name__ == "__main__":
    unittest.main()
