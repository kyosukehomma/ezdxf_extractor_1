import os
import sys
import re
import glob
import shutil
import unicodedata
from decimal import Decimal, InvalidOperation
from datetime import datetime
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
NO_PATTERN   = re.compile(r"(?:No\.|NO\.)\d+(?:\+\d+(?:\.\d+)?)?")

TOLERANCE_Y_ROW   =  0.5
TOLERANCE_X_GROUP =  0.5
TOLERANCE_Y_GROUP = 10.0
MERGE_Y_TOLERANCE =  1.0

SHEET_NAME = "入力"
ROW_START = 5
TEMPLATE_BASE = "01-01-03土工"
TEMPLATE_NAME = TEMPLATE_BASE + ".xlsx"

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
        texts.append({"text": s, "x": t.dxf.insert[0], "y": t.dxf.insert[1]})
    return texts

def extract_no_texts(all_texts):
    results = []
    for t in all_texts:
        norm = unicodedata.normalize("NFKC", t["text"])
        if any(unicodedata.east_asian_width(c) in ("W", "F") for c in norm):
            continue

        if NO_PATTERN.fullmatch(norm):
            m = re.match(
                r"(?:No\.|NO\.)(\d+)(?:\+(\d+(?:\.\d+)?))?$",
                norm
            )
            if m:
                results.append({
                "text": m.group(1),
                "plus": Decimal(m.group(2)) if m.group(2) else None,
                "x": t["x"],
                "y": t["y"]
            })
        else:
            match = re.search(r"\(([^)]+)\)", norm)
            if match and "+" in match.group(1):
                head, plus = match.group(1).split("+", 1)
                try:
                    plus = Decimal(plus)
                except InvalidOperation:
                    pass
                results.append({"text": head, "plus": plus, "x": t["x"], "y": t["y"]})

    for r in results:
        print(r)

    return results

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
# メイン
# ==================================================

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

    for dxf_file in glob.glob(os.path.join(input_dir, "*.dxf")):
        print(dxf_file)
        doc = polyline_to_line_keep_others.main(ezdxf.readfile(dxf_file))
        all_texts = collect_all_texts(doc.modelspace())

        ###
        target = "NO.0+7.935"
        for i, t in enumerate(all_texts):
            if t["text"] == target:
                print(f"index={i}, x={t['x']}, y={t['y']}")
                break
        ###

        no_texts  = extract_no_texts(all_texts)

        ba_records = build_ba_records(all_texts)
        x_groups   = group_by_x(ba_records)
        final      = group_by_y(x_groups)

        normalize_kinds(final)
        final = merge_x_groups(final)

        # ---------- No 照合 ----------
        out_data = []
        for xg in final:
            for sg in xg:
                base = sg[0]
                closest = min(
                    no_texts,
                    key=lambda n: ((n["x"] - base["x"])**2 + (n["y"] - base["y"])**2)**0.5,
                    default=None,
                )
                if closest:
                    print("closest=", closest)
                    out_data.append({
                        "測点": Decimal(closest["text"].replace("No.","").replace("NO.","")),
                        "追加距離": closest["plus"],
                        "データ": sg
                    })

        out_data.sort(key=sort_key)

        # ---------- Excel ----------
        wb = load_workbook(template_path)
        ws = wb[SHEET_NAME]

        name = TEMPLATE_BASE + "_" + os.path.splitext(os.path.basename(dxf_file))[0]
        output_path = os.path.join(output_subfolder, name + ".xlsx")

        for r, item in enumerate(out_data):
            row = ROW_START + r
            ws.cell(row=row, column=3).value = item["測点"]
            ws.cell(row=row, column=5).value = item["追加距離"]

            for d in item["データ"]:
                for col in range(1, ws.max_column + 1):
                    c_header = ws.cell(row=2, column=col)
                    if c_header.value is not None:
                        if normalize_text(c_header.value) == normalize_text(d["種別"]):
                            c = ws.cell(row=row, column=col)
                            c.value = d["数量"]
                            c.alignment = Alignment(horizontal="right")

        wb.save(output_path)
        wb.close()

        shutil.move(dxf_file, input_subfolder)

        dxf_files.append(os.path.basename(dxf_file))

    return dxf_files


if __name__ == "__main__":
    print("oudan_red_yellow_checker [Version 1.0.0]")

    # Tkinterの初期化
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    dxf_files = []
    if messagebox.askokcancel("処理確認", "DXFファイルの処理を開始しますか？"):
        dxf_files = proc(dxf_files)
        messagebox.showinfo("処理完@mermaid-chart了", "\n".join(dxf_files))