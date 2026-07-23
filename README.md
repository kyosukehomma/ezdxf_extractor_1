# DXF 横断図 赤黄チェッカー

`input` フォルダ内のDXF横断図から測点・種別・数量を抽出し、Excelテンプレートへ転記するデスクトップツールです。

## 処理内容

1. `input/*.dxf` を読み込む
2. DXFの `TEXT` エンティティから測点、種別、単位、数量を抽出する
3. `template/01-01-03土工.xlsx` の「入力」シートへ転記する
4. DXFごとのExcelを `output/exec_YYMMDD_HHMMSS/` に保存する
5. 処理済みDXFを `input/used_YYMMDD_HHMMSS/` へ移動する

処理前のDXFを残す必要がある場合は、あらかじめ別の場所へコピーしてください。

## ディレクトリ構成

```text
.
├── input/                         # DXF配置先（DXFはGit管理対象外）
├── output/                        # 実行時に自動作成
├── template/
│   └── 01-01-03土工.xlsx
├── tools/
├── oudan_red_yellow_checker.py
└── requirements.txt
```

## 実行環境

- Python 3
- Tkinter
- `ezdxf`
- `openpyxl`

Ubuntu／WSLでTkinterが未導入の場合:

```bash
sudo apt install python3-tk
```

Python依存パッケージの導入:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

## 実行

1. DXFファイルを `input/` に配置します。
2. 次のコマンドを実行します。

```bash
python oudan_red_yellow_checker.py
```

3. 確認ダイアログで処理を開始します。

## 現在の主な前提

- 抽出対象はモデル空間の `TEXT` エンティティです。`MTEXT` は対象外です。
- Excelテンプレートのシート名「入力」と2行目の種別見出しに依存します。
- テンプレートに存在しない種別は転記されません。
- Excel数式の再計算はExcelでブックを開いた際の再計算に依存します。
