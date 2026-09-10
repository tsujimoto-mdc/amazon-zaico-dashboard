# -*- coding: utf-8 -*-
"""
ZAICO自動転記ロジック（ダッシュボード共通コア）
====================================================================
元の「zaico_tool.py」から、業務ロジック部分だけを取り出したものです。
やっている事は元ツールとまったく同じで、

  ・Amazon FBA 納品プランExcel の「SKU」列
  ・ZAICO CSV の「型番」列

を突き合わせ、ZAICO CSV の保管場所別在庫数量を

  ・「本社在庫」   → 「本社ZAICO」列
  ・「★3F EC★」  → 「3FZAICO」列

として、納品プランExcelの「数量」列の右隣に2列追加して書き戻します。

違いは「フォルダを見に行く」代わりに、アップロードされたファイルの
バイト列（bytes）を直接受け取って、結果もバイト列で返す点だけです。
====================================================================
"""

from __future__ import annotations

import io
from copy import copy
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


# ====================================================================
# 業務ルール設定
#   運用ルールが変わった場合はここだけ直せば良いようにまとめてあります。
# ====================================================================

CSV_COL_SKU = "型番"              # ZAICO CSVで納品プランのSKUに対応する列
CSV_COL_LOCATION = "保管場所"      # ZAICO CSVの保管場所列（D列）
CSV_COL_QTY = "数量"              # ZAICO CSVの在庫数量列（F列）

LOCATION_HONSHA = "本社在庫"       # 本社ZAICOに集計する保管場所名
LOCATION_3F = "★3F EC★"          # 3FZAICOに集計する保管場所名

NEW_COL_HONSHA = "本社ZAICO"
NEW_COL_3F = "3FZAICO"

OUTPUT_SUFFIX = "_ZAICO数量記載"

COL_WIDTH_HONSHA = 9.625           # サンプルファイルに合わせた列幅
COL_WIDTH_3F = 8.25

FULLWIDTH_SPACE = "　"
CSV_ENCODINGS_TO_TRY = ("utf-8-sig", "utf-8", "cp932", "shift_jis")


# ====================================================================
# データ構造
# ====================================================================

@dataclass
class FileResult:
    input_name: str
    output_name: Optional[str] = None
    output_bytes: Optional[bytes] = None
    success: bool = False
    error: Optional[str] = None
    total_sku: int = 0
    missing_skus: list = field(default_factory=list)
    rows: list = field(default_factory=list)  # 画面プレビュー用（ダウンロードファイルには影響しません）


# ====================================================================
# 共通ユーティリティ
# ====================================================================

def normalize_text(value) -> str:
    """SKUや保管場所名を比較できるよう正規化する（前後の半角/全角スペース除去）"""
    if value is None:
        return ""
    s = str(value).strip()
    s = s.replace(FULLWIDTH_SPACE, "").strip()
    return s


# ====================================================================
# ZAICO CSV の読み込み
# ====================================================================

def read_zaico_csv(csv_bytes: bytes) -> pd.DataFrame:
    last_err = None
    for enc in CSV_ENCODINGS_TO_TRY:
        try:
            return pd.read_csv(io.BytesIO(csv_bytes), encoding=enc)
        except (UnicodeDecodeError, UnicodeError) as e:
            last_err = e
            continue
    raise RuntimeError("CSVの文字コードを判定できませんでした") from last_err


def build_zaico_lookup(df: pd.DataFrame) -> dict:
    """SKU(型番)ごとの {本社ZAICO: 数量, 3FZAICO: 数量} を作る"""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    for required_col in (CSV_COL_SKU, CSV_COL_LOCATION, CSV_COL_QTY):
        if required_col not in df.columns:
            raise RuntimeError(
                f"ZAICOのCSVに必要な列「{required_col}」が見つかりません。\n"
                f"CSVの列名: {list(df.columns)}"
            )

    work = pd.DataFrame({
        "sku": df[CSV_COL_SKU].map(normalize_text),
        "location": df[CSV_COL_LOCATION].astype(str).str.strip(),
        "qty": pd.to_numeric(df[CSV_COL_QTY], errors="coerce").fillna(0),
    })
    work = work[work["sku"] != ""]

    lookup: dict = {}
    for sku, sub in work.groupby("sku"):
        honsha_qty = int(sub.loc[sub["location"] == LOCATION_HONSHA, "qty"].sum())
        f3_qty = int(sub.loc[sub["location"] == LOCATION_3F, "qty"].sum())
        lookup[sku] = {NEW_COL_HONSHA: honsha_qty, NEW_COL_3F: f3_qty}

    return lookup


# ====================================================================
# 納品プランExcelの読み書き
# ====================================================================

def _find_header_row(ws, max_scan: int = 40) -> int:
    for r in range(1, max_scan + 1):
        a = normalize_text(ws.cell(row=r, column=1).value)
        b = normalize_text(ws.cell(row=r, column=2).value)
        if a == "SKU" and b == "商品名":
            return r
    raise RuntimeError("「SKU」「商品名」のヘッダー行が見つかりませんでした。")


def _find_qty_column(ws, header_row: int, max_scan_cols: int = 20) -> int:
    for c in range(1, max_scan_cols + 1):
        if normalize_text(ws.cell(row=header_row, column=c).value) == "数量":
            return c
    raise RuntimeError("「数量」列が見つかりませんでした。")


def _find_last_data_row(ws, header_row: int, sku_col: int, safety_limit: int = 20000) -> int:
    last = header_row
    r = header_row + 1
    while r - header_row <= safety_limit:
        v = normalize_text(ws.cell(row=r, column=sku_col).value)
        if v == "":
            break
        last = r
        r += 1
    return last


def process_one_workbook(excel_bytes: bytes, input_name: str, zaico_lookup: dict) -> FileResult:
    """1つの納品プランExcel（bytes）を処理して、結果Excel（bytes）を返す。"""
    result = FileResult(input_name=input_name)
    try:
        wb = load_workbook(io.BytesIO(excel_bytes))
        ws = wb.active

        header_row = _find_header_row(ws)
        sku_col = 1  # A列 = SKU
        qty_col = _find_qty_column(ws, header_row)
        last_row = _find_last_data_row(ws, header_row, sku_col)

        col_honsha = qty_col + 1
        col_3f = qty_col + 2

        # --- ヘッダー（「数量」の書式をコピー） ---
        qty_header_cell = ws.cell(row=header_row, column=qty_col)
        for offset, label in ((1, NEW_COL_HONSHA), (2, NEW_COL_3F)):
            cell = ws.cell(row=header_row, column=qty_col + offset)
            cell.value = label
            cell.font = copy(qty_header_cell.font)
            cell.border = copy(qty_header_cell.border)
            cell.fill = copy(qty_header_cell.fill)
            cell.alignment = copy(qty_header_cell.alignment)
            cell.number_format = qty_header_cell.number_format

        ws.column_dimensions[get_column_letter(col_honsha)].width = COL_WIDTH_HONSHA
        ws.column_dimensions[get_column_letter(col_3f)].width = COL_WIDTH_3F

        # --- データ行 ---
        missing_skus = []
        preview_rows = []
        total_sku = 0

        name_col = 2  # B列 = 商品名

        for r in range(header_row + 1, last_row + 1):
            sku = normalize_text(ws.cell(row=r, column=sku_col).value)
            if sku == "":
                continue
            total_sku += 1

            qty_cell = ws.cell(row=r, column=qty_col)
            values = zaico_lookup.get(sku)
            found = values is not None
            if values is None:
                values = {NEW_COL_HONSHA: 0, NEW_COL_3F: 0}
                missing_skus.append(sku)

            for offset, key in ((1, NEW_COL_HONSHA), (2, NEW_COL_3F)):
                cell = ws.cell(row=r, column=qty_col + offset)
                cell.value = values[key]
                cell.font = copy(qty_cell.font)
                cell.border = copy(qty_cell.border)
                cell.fill = copy(qty_cell.fill)
                cell.alignment = copy(qty_cell.alignment)
                cell.number_format = "General"

            plan_qty_raw = qty_cell.value
            try:
                plan_qty = int(plan_qty_raw)
            except (TypeError, ValueError):
                plan_qty = None

            honsha = values[NEW_COL_HONSHA]
            f3 = values[NEW_COL_3F]
            preview_rows.append({
                "SKU": sku,
                "商品名": ws.cell(row=r, column=name_col).value,
                "納品数量": plan_qty_raw,
                NEW_COL_HONSHA: honsha,
                NEW_COL_3F: f3,
                "在庫合計": honsha + f3,
                "ZAICO登録": "○" if found else "× 未登録",
                "在庫判定": _stock_judgement(plan_qty, honsha + f3, found),
            })

        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        stem = input_name.rsplit(".", 1)[0]
        result.success = True
        result.output_name = f"{stem}{OUTPUT_SUFFIX}.xlsx"
        result.output_bytes = buffer.read()
        result.total_sku = total_sku
        result.missing_skus = missing_skus
        result.rows = preview_rows
        return result

    except Exception as e:  # noqa: BLE001
        result.success = False
        result.error = str(e)
        return result


def _stock_judgement(plan_qty, stock_total: int, found: bool) -> str:
    """画面プレビュー用の在庫判定。ダウンロードするExcelの中身には影響しません。"""
    if not found:
        return "ZAICO未登録"
    if plan_qty is None:
        return "-"
    if stock_total >= plan_qty:
        return "OK"
    return f"不足 ({stock_total} < {plan_qty})"


# ====================================================================
# まとめ処理
# ====================================================================

def run_transcription(excel_files: list[tuple[str, bytes]], csv_bytes: bytes):
    """
    excel_files : [(ファイル名, bytes), ...]
    csv_bytes   : ZAICO CSV のバイト列
    戻り値      : (results: list[FileResult], error: Optional[str])
    """
    try:
        df = read_zaico_csv(csv_bytes)
        lookup = build_zaico_lookup(df)
    except Exception as e:  # noqa: BLE001
        return [], f"ZAICOのCSV読み込みでエラーが発生しました。\n{e}"

    if not excel_files:
        return [], "Amazon納品プランExcelがアップロードされていません。"

    results = [
        process_one_workbook(data, name, lookup)
        for name, data in excel_files
    ]
    return results, None
