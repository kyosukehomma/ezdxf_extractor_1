import os
import sys
import re
import glob
import shutil
import unicodedata
from collections import Counter
from decimal import Decimal, InvalidOperation
from datetime import datetime
from math import hypot
import tkinter as tk
from tkinter import messagebox

import ezdxf
from openpyxl import load_workbook
from openpyxl.styles import Alignment

from tools import polyline_to_line_keep_others

# ==================================================
# 定数定義
# ==================================================

UNIT_PATTERN = re.compile(r"^(m2|㎡|m)$")
STATION_PATTERN = re.compile(
    r"^(?:No\.|NO\.)(\d+)(?:\+(\d+(?:\.\d+)?))?$",
    re.IGNORECASE,
)
PARENTHESIZED_STATION_PATTERN = re.compile(
    r"\((?:No\.|NO\.)(\d+)(?:\+(\d+(?:\.\d+)?))?\)",
    re.IGNORECASE,
)

TOLERANCE_Y_ROW   =  0.5
TOLERANCE_X_GROUP =  0.5
TOLERANCE_Y_GROUP = 10.0
MERGE_Y_TOLERANCE =  1.0

EARTHWORK_TEXT_LAYER_PREFIX = "D-MTR-TXT"
EARTHWORK_HEADER_SEQUENCE = ("種別", "単位", "数量", "種別", "単位", "数量")
TABLE_HEADER_Y_TOLERANCE = 0.15
TABLE_HEADER_X_GAP = 50.0
MAX_DATA_GROUP_TO_TABLE_DISTANCE = 10.0

STATION_TEXT_LAYER = "D-BMK-HTXT"
CENTER_MARKER_LAYER = "D-BMK"
CENTER_MARKER_TEXTS = {"CL", "C.L", "C.L."}
MAX_TABLE_TO_CENTER_DISTANCE = 75.0
MAX_TABLE_TO_CENTER_Y_DISTANCE = 15.0
MAX_CENTER_TO_STATION_DISTANCE = 15.0
MAX_TABLE_TO_STATION_DISTANCE = 70.0
MAX_TABLE_TO_STATION_Y_DISTANCE = 15.0
MIN_NEAREST_DISTANCE_GAP = 1.0

SHEET_NAME = "入力"
ROW_START = 5
TEMPLATE_BASE = "01-01-03土工"
TEMPLATE_NAME = TEMPLATE_BASE + ".xlsx"


class ExtractionError(RuntimeError):
    """誤った転記を避けるために抽出処理を中断するエラー。"""

# ==================================================
# ユーティリティ
# ==================================================

def get_base_path():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def normalize_text(s: str) -> str:
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", s))

def replace_angle_brackets(s: str) -> str:
    return s.replace("<", "＜").replace(">", "＞")

def sort_key(item):
    plus = item["追加距離"]
    return (item["測点"], Decimal("-1") if plus is None else plus)

# ==================================================
# DXF → TEXT 抽出
# ==================================================

def collect_all_texts(msp):
    texts = []
    for t in msp.query("TEXT"):
        s = replace_angle_brackets(normalize_text(t.dxf.text))
        alignment_point = t.get_placement()[1]
        texts.append({
            "text": s,
            "x": t.dxf.insert[0],
            "y": t.dxf.insert[1],
            "anchor_x": alignment_point.x,
            "anchor_y": alignment_point.y,
            "layer": t.dxf.layer,
            "height": t.dxf.height,
            "handle": t.dxf.handle,
        })
    return texts

def extract_no_texts(all_texts):
    results = []
    for t in all_texts:
        if t.get("layer") != STATION_TEXT_LAYER:
            continue

        norm = unicodedata.normalize("NFKC", t["text"])
        match = STATION_PATTERN.fullmatch(norm)
        if match is None:
            match = PARENTHESIZED_STATION_PATTERN.search(norm)
        if match is None:
            continue

        try:
            plus = Decimal(match.group(2)) if match.group(2) is not None else None
        except InvalidOperation:
            continue

        results.append({
            "text": match.group(1),
            "plus": plus,
            "x": t["x"],
            "y": t["y"],
            "anchor_x": t["anchor_x"],
            "anchor_y": t["anchor_y"],
            "layer": t["layer"],
            "handle": t.get("handle"),
        })
    return results


def extract_center_markers(all_texts):
    return [
        t for t in all_texts
        if t.get("layer") == CENTER_MARKER_LAYER
        and t["text"].upper() in CENTER_MARKER_TEXTS
    ]

# ==================================================
# 行構築
# ==================================================

def build_ba_records(all_texts):
    records = []

    for item in all_texts:
        if not UNIT_PATTERN.fullmatch(item["text"]):
            continue

        same_row = [
            t for t in all_texts
            if abs(t["y"] - item["y"]) <= TOLERANCE_Y_ROW
        ]
        same_row.sort(key=lambda x: x["x"])

        idx = [t["x"] for t in same_row].index(item["x"])
        kind = same_row[idx - 1]["text"] if idx - 1 >= 0 else ""
        qty  = same_row[idx + 1]["text"] if idx + 1 < len(same_row) else ""

        try:
            qty = Decimal(qty)
        except Exception:
            pass

        records.append({
            "x": item["x"],
            "y": item["y"],
            "種別": kind,
            "単位": item["text"],
            "数量": qty
        })

    return records

# ==================================================
# グループ化
# ==================================================

def group_by_x(ba_records):
    groups = []
    for r in ba_records:
        for g in groups:
            if abs(g["x_ref"] - r["x"]) <= TOLERANCE_X_GROUP:
                g["items"].append(r)
                break
        else:
            groups.append({"x_ref": r["x"], "items": [r]})
    return groups

def group_by_y(x_groups):
    final = []
    for g in x_groups:
        items = sorted(g["items"], key=lambda i: i["y"], reverse=True)
        subgroups, buf, last_y = [], [], None
        for it in items:
            if last_y is None or abs(last_y - it["y"]) <= TOLERANCE_Y_GROUP:
                buf.append(it)
            else:
                subgroups.append(buf)
                buf = [it]
            last_y = it["y"]
        if buf:
            subgroups.append(buf)
        final.append(subgroups)
    return final

# ==================================================
# 種別補正
# ==================================================

WIDTH_MAP = {
    "路体盛土(W＜2.5m)": "路体盛土",
    "路床盛土(W＜2.5m)": "路床盛土",
    "路体外盛土(W＜2.5m)": "路体外盛土",
}

SIDE_RIGHT_MAP = {
    "切土法面整形(左)": "切土法面整形",
    "盛土法面整形(左)": "盛土法面整形",
    "路肩表層(左)": "路肩表層",
    "路肩下層路盤(左)": "路肩下層路盤",
}

def normalize_kinds(final_groups):
    for x_group in final_groups:
        for sg in x_group:
            if len(sg) < 2:
                continue

            base = sg[0]["種別"]

            # 幅員系
            if base == "オープン掘削":
                for i in range(1, len(sg)):
                    k = sg[i]["種別"]
                    if k in WIDTH_MAP:
                        prefix = WIDTH_MAP[k]
                        if i + 1 < len(sg) and sg[i + 1]["種別"].startswith("("):
                            sg[i + 1]["種別"] = f"{prefix}{sg[i + 1]['種別']}"
                        if i + 2 < len(sg) and sg[i + 2]["種別"].startswith("("):
                            sg[i + 2]["種別"] = f"{prefix}{sg[i + 2]['種別']}"

            # 左右系
            elif base == "切土法面整形(左)":
                for i in range(1, len(sg)):
                    prev, curr = sg[i - 1]["種別"], sg[i]["種別"]
                    if curr == "(右)" and prev in SIDE_RIGHT_MAP:
                        sg[i]["種別"] = f"{SIDE_RIGHT_MAP[prev]}(右)"

# ==================================================
# merge
# ==================================================

def merge_x_groups(final_groups):
    result, prev = [], None

    for g in final_groups:
        if prev is None:
            prev = g
            continue

        len_ok  = len(prev) == len(g)
        kind_ok = prev[0][0]["種別"] != g[0][0]["種別"]
        y_ok    = abs(prev[0][0]["y"] - g[0][0]["y"]) <= MERGE_Y_TOLERANCE

        if len_ok and kind_ok and y_ok:
            for i in range(len(prev)):
                prev[i].extend(g[i])
        else:
            result.append(prev)
            prev = g

    if prev:
        result.append(prev)
    return result


# ==================================================
# 土工表の検出・測点との対応付け
# ==================================================

def _distance(a, b):
    a_x = a.get("anchor_x", a["x"])
    a_y = a.get("anchor_y", a["y"])
    b_x = b.get("anchor_x", b["x"])
    b_y = b.get("anchor_y", b["y"])
    return hypot(a_x - b_x, a_y - b_y)


def _split_by_x_gap(items, max_gap):
    chunks = []
    current = []
    last_x = None

    for item in sorted(items, key=lambda t: t["x"]):
        if last_x is not None and item["x"] - last_x > max_gap:
            chunks.append(current)
            current = []
        current.append(item)
        last_x = item["x"]

    if current:
        chunks.append(current)
    return chunks


def detect_earthwork_tables(all_texts):
    """左右2組の「種別・単位・数量」から土工表のヘッダーを検出する。"""
    header_words = set(EARTHWORK_HEADER_SEQUENCE)
    header_texts = [
        t for t in all_texts
        if t.get("layer", "").startswith(EARTHWORK_TEXT_LAYER_PREFIX)
        and t["text"] in header_words
    ]

    row_groups = []
    for item in sorted(
        header_texts,
        key=lambda t: (t["layer"], -t["y"], t["x"]),
    ):
        same_row = next(
            (
                group for group in row_groups
                if group["layer"] == item["layer"]
                and abs(group["y_ref"] - item["y"]) <= TABLE_HEADER_Y_TOLERANCE
            ),
            None,
        )
        if same_row is None:
            row_groups.append({
                "layer": item["layer"],
                "y_ref": item["y"],
                "items": [item],
            })
        else:
            same_row["items"].append(item)

    tables = []
    for row_group in row_groups:
        for chunk in _split_by_x_gap(
            row_group["items"],
            TABLE_HEADER_X_GAP,
        ):
            ordered_words = tuple(t["text"] for t in sorted(chunk, key=lambda t: t["x"]))
            if ordered_words != EARTHWORK_HEADER_SEQUENCE:
                continue

            tables.append({
                "x": sum(t["x"] for t in chunk) / len(chunk),
                "y": sum(t["y"] for t in chunk) / len(chunk),
                "layer": row_group["layer"],
                "headers": tuple(chunk),
            })

    tables.sort(key=lambda t: (-t["y"], t["x"]))
    return tables


def build_table_data_groups(all_texts):
    """第1段階では、表内データの構築だけ現行ロジックを再利用する。"""
    ba_records = build_ba_records(all_texts)
    final = group_by_y(group_by_x(ba_records))
    normalize_kinds(final)
    final = merge_x_groups(final)
    return [subgroup for x_group in final for subgroup in x_group]


def assign_data_groups_to_tables(tables, data_groups):
    """各データ群を最寄りの表へ割り当て、1表1群でなければ停止する。"""
    if not tables:
        raise ExtractionError("土工表の共通ヘッダーを検出できませんでした。")

    assignments = [[] for _ in tables]
    for data_group_index, data_group in enumerate(data_groups):
        if not data_group:
            continue
        table_index = min(
            range(len(tables)),
            key=lambda i: _distance(tables[i], data_group[0]),
        )
        assignments[table_index].append(data_group_index)

    invalid = [
        (index + 1, len(group_indexes))
        for index, group_indexes in enumerate(assignments)
        if len(group_indexes) != 1
    ]
    if invalid:
        count_summary = ", ".join(
            f"{count}群の表が{table_count}件"
            for count, table_count in sorted(
                Counter(count for _, count in invalid).items()
            )
        )
        raise ExtractionError(
            f"土工表{len(tables)}件にデータ群{len(data_groups)}件を"
            f"一意に対応付けできません（{count_summary}）。"
            "現段階では誤転記を避けるため、このDXFを出力しません。"
        )

    for table_index, group_indexes in enumerate(assignments):
        data_group_index = group_indexes[0]
        data_group = data_groups[data_group_index]
        table = tables[table_index]
        if _distance(table, data_group[0]) > MAX_DATA_GROUP_TO_TABLE_DISTANCE:
            raise ExtractionError(
                "土工表とデータ群の距離が安全上限を超えています。"
                "現段階では誤転記を避けるため、このDXFを出力しません。"
            )

        nearest_group_index = _nearest_item_index(
            table,
            [group[0] for group in data_groups],
            "データ群",
        )
        if nearest_group_index != data_group_index:
            raise ExtractionError(
                "土工表とデータ群が相互に最寄りではありません。"
                "現段階では誤転記を避けるため、このDXFを出力しません。"
            )

    return [
        data_groups[group_indexes[0]]
        for group_indexes in assignments
    ]


def _nearest_item_index(origin, items, target_name):
    ranked = sorted(
        (
            (_distance(origin, item), index)
            for index, item in enumerate(items)
        ),
        key=lambda pair: pair[0],
    )
    if (
        len(ranked) >= 2
        and ranked[1][0] - ranked[0][0] < MIN_NEAREST_DISTANCE_GAP
    ):
        raise ExtractionError(
            f"最寄りの{target_name}候補が近接しているため、一意に決定できません。"
        )
    return ranked[0][1]


def resolve_table_stations(tables, stations, center_markers):
    """
    表ごとの中心測点を解決する。

    CL TEXTがある場合は「表→CL→測点」、ない場合は測点を最寄り表へ
    逆割当し、候補が1件だけの表に限って採用する。
    """
    if not tables:
        raise ExtractionError("中心測点を対応付ける土工表がありません。")
    if not stations:
        raise ExtractionError("測点TEXTを検出できませんでした。")

    if center_markers:
        center_indexes = []
        for table in tables:
            center_index = _nearest_item_index(table, center_markers, "CL")
            center = center_markers[center_index]
            if (
                _distance(table, center) > MAX_TABLE_TO_CENTER_DISTANCE
                or abs(table["y"] - center["y"]) > MAX_TABLE_TO_CENTER_Y_DISTANCE
            ):
                raise ExtractionError(
                    "土工表に対応するCLを安全な距離内で特定できませんでした。"
                )
            center_indexes.append(center_index)

        if len(set(center_indexes)) != len(center_indexes):
            raise ExtractionError("複数の土工表が同じCLに対応し、中心断面を一意に特定できません。")

        station_indexes = []
        for center_index in center_indexes:
            center = center_markers[center_index]
            station_index = _nearest_item_index(center, stations, "測点")
            if _distance(center, stations[station_index]) > MAX_CENTER_TO_STATION_DISTANCE:
                raise ExtractionError("CLに対応する測点を安全な距離内で特定できませんでした。")
            station_indexes.append(station_index)

        if len(set(station_indexes)) != len(station_indexes):
            raise ExtractionError("複数の土工表が同じ測点に対応し、中心測点を一意に特定できません。")

        return [stations[index] for index in station_indexes]

    assignments = [[] for _ in tables]
    for station_index, station in enumerate(stations):
        table_index = _nearest_item_index(station, tables, "土工表")
        assignments[table_index].append(station_index)

    invalid = [
        (index + 1, len(station_indexes))
        for index, station_indexes in enumerate(assignments)
        if len(station_indexes) != 1
    ]
    if invalid:
        count_summary = ", ".join(
            f"{count}測点の表が{table_count}件"
            for count, table_count in sorted(
                Counter(count for _, count in invalid).items()
            )
        )
        raise ExtractionError(
            "CL TEXTがなく、各土工表に対応する測点を1件に絞れません"
            f"（{count_summary}）。"
        )

    for table_index, station_indexes in enumerate(assignments):
        station_index = station_indexes[0]
        table = tables[table_index]
        station = stations[station_index]
        if (
            _distance(table, station) > MAX_TABLE_TO_STATION_DISTANCE
            or abs(table["y"] - station.get("anchor_y", station["y"]))
            > MAX_TABLE_TO_STATION_Y_DISTANCE
        ):
            raise ExtractionError(
                "CL TEXTがない土工表と測点の距離が安全上限を超えています。"
            )

        nearest_station_index = _nearest_item_index(table, stations, "測点")
        if nearest_station_index != station_index:
            raise ExtractionError(
                "CL TEXTがない土工表と測点が相互に最寄りではありません。"
            )

    return [
        stations[station_indexes[0]]
        for station_indexes in assignments
    ]


def extract_output_records(all_texts):
    """DXFのTEXTから、Excel転記前のレコードを副作用なしで抽出する。"""
    tables = detect_earthwork_tables(all_texts)
    data_groups = build_table_data_groups(all_texts)
    table_data_groups = assign_data_groups_to_tables(tables, data_groups)

    stations = extract_no_texts(all_texts)
    center_markers = extract_center_markers(all_texts)
    table_stations = resolve_table_stations(tables, stations, center_markers)

    out_data = [
        {
            "測点": Decimal(station["text"]),
            "追加距離": station["plus"],
            "データ": data_group,
        }
        for station, data_group in zip(table_stations, table_data_groups)
    ]
    out_data.sort(key=sort_key)
    return out_data

# ==================================================
# メイン
# ==================================================

def write_output_workbook(template_path, output_path, out_data):
    wb = load_workbook(template_path)
    try:
        ws = wb[SHEET_NAME]

        for r, item in enumerate(out_data):
            row = ROW_START + r
            ws.cell(row=row, column=3).value = item["測点"]
            ws.cell(row=row, column=5).value = item["追加距離"]

            for data in item["データ"]:
                for col in range(1, ws.max_column + 1):
                    header = ws.cell(row=2, column=col)
                    if (
                        header.value is not None
                        and normalize_text(header.value) == normalize_text(data["種別"])
                    ):
                        cell = ws.cell(row=row, column=col)
                        cell.value = data["数量"]
                        cell.alignment = Alignment(horizontal="right")

        wb.save(output_path)
    finally:
        wb.close()


def proc(dxf_files):
    base = get_base_path()
    input_dir = os.path.join(base, "input")
    output_dir = os.path.join(base, "output")
    template_path = os.path.join(base, "template", TEMPLATE_NAME)

    os.makedirs(output_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%y%m%d_%H%M%S")

    input_subfolder  = os.path.join(input_dir,  f"used_{timestamp}")
    os.makedirs(input_subfolder,  exist_ok=True)

    output_subfolder = os.path.join(output_dir, f"exec_{timestamp}")
    os.makedirs(output_subfolder, exist_ok=True)

    failures = []
    for dxf_file in glob.glob(os.path.join(input_dir, "*.dxf")):
        print(dxf_file)
        doc = polyline_to_line_keep_others.main(ezdxf.readfile(dxf_file))
        all_texts = collect_all_texts(doc.modelspace())

        try:
            out_data = extract_output_records(all_texts)
        except ExtractionError as exc:
            failures.append(f"{os.path.basename(dxf_file)}: {exc}")
            continue

        # ---------- Excel ----------
        name = TEMPLATE_BASE + "_" + os.path.splitext(os.path.basename(dxf_file))[0]
        output_path = os.path.join(output_subfolder, name + ".xlsx")
        write_output_workbook(template_path, output_path, out_data)

        shutil.move(dxf_file, input_subfolder)
        dxf_files.append(os.path.basename(dxf_file))

    return dxf_files, failures


if __name__ == "__main__":
    print("oudan_red_yellow_checker [Version 1.0.0]")

    # Tkinterの初期化
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    if messagebox.askokcancel("処理確認", "DXFファイルの処理を開始しますか？"):
        dxf_files, failures = proc([])
        if failures:
            result_lines = []
            if dxf_files:
                result_lines.append("処理完了:\n" + "\n".join(dxf_files))
            result_lines.append("未処理:\n" + "\n".join(failures))
            messagebox.showwarning("処理結果", "\n\n".join(result_lines))
        else:
            messagebox.showinfo("処理完了", "\n".join(dxf_files))
