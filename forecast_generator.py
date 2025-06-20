"""予算・FC・実績を横並びで比較するExcelを生成するスクリプト."""

from __future__ import annotations

import datetime
import re
import tkinter as tk
from tkinter import filedialog, simpledialog

import jpholiday
import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows


def find_date_column(df: pd.DataFrame) -> str:
    """Return the column name that represents a date.

    The function normalizes column names by stripping whitespace and converting
    them to lowercase. It then searches for a column likely representing a
    date, such as "日付", "宿泊日" or "date". If no suitable column is found,
    a ``KeyError`` is raised.
    """

    normalized = {c.strip().lower(): c for c in df.columns}
    candidates = ["日付", "宿泊日", "date"]

    for key, original in normalized.items():
        cleaned = re.sub(r"\s+", "", key)
        if cleaned in candidates:
            return original

    for key, original in normalized.items():
        if re.search(r"日\s*付", key) or "宿泊日" in key or "date" in key:
            return original

    raise KeyError("日付に該当する列が見つかりません")

# === GUIでキャパシティと期首月の取得 ===

try:
    root = tk.Tk()
    root.withdraw()
    capacity = simpledialog.askinteger(
        "キャパシティ入力",
        "客室キャパシティ（部屋数）を入力してください：",
    )
    start_month = simpledialog.askinteger(
        "期首月入力",
        "期首月（例：1〜12）を入力してください：",
    )
    file_path = filedialog.askopenfilename(title="日別予算Excelファイルを選択")
except tk.TclError:
    print("GUI を起動できないため CLI モードで実行します。")
    capacity = int(input("キャパシティ（部屋数）: "))
    start_month = int(input("期首月 (1-12): "))
    file_path = input("日別予算Excelファイルのパス: ")
# === Excel読み込み ===
xls = pd.read_excel(file_path, sheet_name=None)

# === 出力用Excel作成 ===
wb = Workbook()
wb.remove(wb.active)

# === 曜日装飾 ===
sat_fill = PatternFill(start_color="DDEEFF", end_color="DDEEFF", fill_type="solid")
sat_font = Font(color="003366")
sun_fill = PatternFill(start_color="FFE5E5", end_color="FFE5E5", fill_type="solid")
sun_font = Font(color="990000")
gray_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
budget_fill = PatternFill(start_color="E6F2FF", end_color="E6F2FF", fill_type="solid")
fc_fill = PatternFill(start_color="E6FFE6", end_color="E6FFE6", fill_type="solid")
thin = Side(style="thin", color="999999")
medium = Side(style="medium", color="999999")

# --- 各月シート作成 ---
summary_dict: dict[tuple[int, int], dict[str, int]] = {}
for sheet_name, df in xls.items():
    if not isinstance(df, pd.DataFrame) or df.empty:
        continue

    try:
        date_col = find_date_column(df)
    except KeyError:
        continue
    df["日付"] = pd.to_datetime(df[date_col], errors="coerce")
    df = df.dropna(subset=["日付"]).sort_values("日付")
    if df.empty:
        continue

    year = int(df["日付"].dt.year.iloc[0])
    month = int(df["日付"].dt.month.iloc[0])
    ws = wb.create_sheet(title=f"{year}年{month}月")

    # 日付ごとに横持ち化
    metrics = ["室数", "人数", "宿泊売上", "OCC", "ADR", "DOR", "RevPAR"]
    rows: list[dict] = []
    for _, r in df.iterrows():
        date = r["日付"].date()
        weekday_jp = ["月", "火", "水", "木", "金", "土", "日"][date.weekday()]
        if jpholiday.is_holiday(date):
            weekday_jp = f"{weekday_jp}・祝"

        base = {
            "日付": date.strftime("%Y/%m/%d"),
            "曜日": weekday_jp,
            "室数_予算": r.get("室数", ""),
            "人数_予算": r.get("人数", ""),
            "宿泊売上_予算": r.get("宿泊売上", ""),
        }
        for m in ["OCC", "ADR", "DOR", "RevPAR"]:
            base[f"{m}_予算"] = ""
        for kind in ["FC", "実績"]:
            for m in metrics:
                base[f"{m}_{kind}"] = ""
        rows.append(base)

    df_out = pd.DataFrame(rows)

    # 列順を定義
    metrics = ["室数", "人数", "宿泊売上", "OCC", "ADR", "DOR", "RevPAR"]
    cols = ["日付", "曜日"]
    cols += [f"{m}_予算" for m in metrics]
    cols += [f"{m}_FC" for m in metrics]
    cols += ["差_OCC_FC-予算", "差_ADR_FC-予算", "差_売上_FC-予算"]
    cols += [f"{m}_実績" for m in metrics]
    cols += ["差_OCC_実績-FC", "差_ADR_実績-FC", "差_売上_実績-FC"]
    for c in cols:
        if c not in df_out.columns:
            df_out[c] = ""
    df_out = df_out.reindex(columns=cols)

    # 出力
    for r in dataframe_to_rows(df_out, index=False, header=True):
        ws.append(r)

    header_map = {c.value: i for i, c in enumerate(ws[1], start=1)}
    ws.freeze_panes = "C2"
    data_end_row = ws.max_row
    for row in range(2, data_end_row + 1):
        # 予算列の指標をExcel数式で計算
        room_c = get_column_letter(header_map["室数_予算"])
        pax_c = get_column_letter(header_map["人数_予算"])
        sales_c = get_column_letter(header_map["宿泊売上_予算"])
        occ_c = ws.cell(row=row, column=header_map["OCC_予算"])
        adr_c = ws.cell(row=row, column=header_map["ADR_予算"])
        dor_c = ws.cell(row=row, column=header_map["DOR_予算"])
        rev_c = ws.cell(row=row, column=header_map["RevPAR_予算"])

        occ_c.value = f"={room_c}{row}/{capacity}"
        adr_c.value = f"={sales_c}{row}/{room_c}{row}"
        dor_c.value = f"={pax_c}{row}/{room_c}{row}"
        rev_c.value = f"={sales_c}{row}/{capacity}"

        occ_c.number_format = "0.0%"
        adr_c.number_format = "#,##0"
        dor_c.number_format = "0.00"
        rev_c.number_format = "#,##0"
        for col_name in ["室数_予算", "人数_予算", "宿泊売上_予算"]:
            ws.cell(row=row, column=header_map[col_name]).number_format = "#,##0"

        # FC列の数式
        fc_room_c = get_column_letter(header_map["室数_FC"])
        fc_pax_c = get_column_letter(header_map["人数_FC"])
        fc_sales_c = get_column_letter(header_map["宿泊売上_FC"])
        fc_occ = ws.cell(row=row, column=header_map["OCC_FC"])
        fc_adr = ws.cell(row=row, column=header_map["ADR_FC"])
        fc_dor = ws.cell(row=row, column=header_map["DOR_FC"])
        fc_rev = ws.cell(row=row, column=header_map["RevPAR_FC"])

        fc_occ.value = (
            f"=IF({fc_room_c}{row}=\"\", \"\", {fc_room_c}{row}/{capacity})"
        )
        fc_adr.value = (
            f"=IF(OR({fc_sales_c}{row}=\"\", {fc_room_c}{row}=\"\"), \"\", {fc_sales_c}{row}/{fc_room_c}{row})"
        )
        fc_dor.value = (
            f"=IF(OR({fc_pax_c}{row}=\"\", {fc_room_c}{row}=\"\"), \"\", {fc_pax_c}{row}/{fc_room_c}{row})"
        )
        fc_rev.value = (
            f"=IF({fc_sales_c}{row}=\"\", \"\", {fc_sales_c}{row}/{capacity})"
        )

        fc_occ.number_format = "0.0%"
        fc_adr.number_format = "#,##0"
        fc_dor.number_format = "0.00"
        fc_rev.number_format = "#,##0"

        # 実績列の数式
        act_room_c = get_column_letter(header_map["室数_実績"])
        act_pax_c = get_column_letter(header_map["人数_実績"])
        act_sales_c = get_column_letter(header_map["宿泊売上_実績"])
        act_occ = ws.cell(row=row, column=header_map["OCC_実績"])
        act_adr = ws.cell(row=row, column=header_map["ADR_実績"])
        act_dor = ws.cell(row=row, column=header_map["DOR_実績"])
        act_rev = ws.cell(row=row, column=header_map["RevPAR_実績"])

        act_occ.value = (
            f"=IF({act_room_c}{row}=\"\", \"\", {act_room_c}{row}/{capacity})"
        )
        act_adr.value = (
            f"=IF(OR({act_sales_c}{row}=\"\", {act_room_c}{row}=\"\"), \"\", {act_sales_c}{row}/{act_room_c}{row})"
        )
        act_dor.value = (
            f"=IF(OR({act_pax_c}{row}=\"\", {act_room_c}{row}=\"\"), \"\", {act_pax_c}{row}/{act_room_c}{row})"
        )
        act_rev.value = (
            f"=IF({act_sales_c}{row}=\"\", \"\", {act_sales_c}{row}/{capacity})"
        )

        act_occ.number_format = "0.0%"
        act_adr.number_format = "#,##0"
        act_dor.number_format = "0.00"
        act_rev.number_format = "#,##0"

        # 手入力セルの表示形式
        for col_name in ["室数_FC", "人数_FC", "宿泊売上_FC", "室数_実績", "人数_実績", "宿泊売上_実績"]:
            ws.cell(row=row, column=header_map[col_name]).number_format = "#,##0"

        # 差異列
        ws.cell(row=row, column=header_map["差_OCC_FC-予算"]).value = (
            f"=IF({get_column_letter(header_map['OCC_FC'])}{row}=\"\", \"\", {get_column_letter(header_map['OCC_FC'])}{row}-{get_column_letter(header_map['OCC_予算'])}{row})"
        )
        ws.cell(row=row, column=header_map["差_ADR_FC-予算"]).value = (
            f"=IF({get_column_letter(header_map['ADR_FC'])}{row}=\"\", \"\", {get_column_letter(header_map['ADR_FC'])}{row}-{get_column_letter(header_map['ADR_予算'])}{row})"
        )
        ws.cell(row=row, column=header_map["差_売上_FC-予算"]).value = (
            f"=IF({get_column_letter(header_map['宿泊売上_FC'])}{row}=\"\", \"\", {get_column_letter(header_map['宿泊売上_FC'])}{row}-{get_column_letter(header_map['宿泊売上_予算'])}{row})"
        )
        ws.cell(row=row, column=header_map["差_OCC_実績-FC"]).value = (
            f"=IF({get_column_letter(header_map['OCC_実績'])}{row}=\"\", \"\", {get_column_letter(header_map['OCC_実績'])}{row}-{get_column_letter(header_map['OCC_FC'])}{row})"
        )
        ws.cell(row=row, column=header_map["差_ADR_実績-FC"]).value = (
            f"=IF({get_column_letter(header_map['ADR_実績'])}{row}=\"\", \"\", {get_column_letter(header_map['ADR_実績'])}{row}-{get_column_letter(header_map['ADR_FC'])}{row})"
        )
        ws.cell(row=row, column=header_map["差_売上_実績-FC"]).value = (
            f"=IF({get_column_letter(header_map['宿泊売上_実績'])}{row}=\"\", \"\", {get_column_letter(header_map['宿泊売上_実績'])}{row}-{get_column_letter(header_map['宿泊売上_FC'])}{row})"
        )

        for diff_col in ["差_OCC_FC-予算", "差_OCC_実績-FC"]:
            ws.cell(row=row, column=header_map[diff_col]).number_format = "0.0%"
        for diff_col in ["差_ADR_FC-予算", "差_ADR_実績-FC"]:
            ws.cell(row=row, column=header_map[diff_col]).number_format = "#,##0"
        for diff_col in ["差_売上_FC-予算", "差_売上_実績-FC"]:
            ws.cell(row=row, column=header_map[diff_col]).number_format = "#,##0"
        for diff_col in [
            "差_OCC_FC-予算",
            "差_ADR_FC-予算",
            "差_売上_FC-予算",
            "差_OCC_実績-FC",
            "差_ADR_実績-FC",
            "差_売上_実績-FC",
        ]:
            ws.cell(row=row, column=header_map[diff_col]).fill = PatternFill(
                start_color="FFFAD0",
                end_color="FFFAD0",
                fill_type="solid",
            )

        # 曜日装飾
        w_cell = ws.cell(row=row, column=header_map["曜日"])
        date_cell = ws.cell(row=row, column=header_map["日付"])
        try:
            d_obj = datetime.datetime.strptime(str(date_cell.value), "%Y/%m/%d").date()
        except Exception:
            continue
        if jpholiday.is_holiday(d_obj) or w_cell.value in ["日", "祝"]:
            w_cell.fill = sun_fill
            w_cell.font = sun_font
        elif w_cell.value == "土":
            w_cell.fill = sat_fill
            w_cell.font = sat_font

    # 実績入力時にFC列をグレー化
    fc_col = get_column_letter(header_map["室数_FC"])
    actual_col = get_column_letter(header_map["室数_実績"])
    formula = f"LEN(${actual_col}2)>0"
    for offset in range(0, 7):
        col_letter = get_column_letter(header_map["室数_FC"] + offset)
        rule = FormulaRule(formula=[formula], fill=gray_fill)
        ws.conditional_formatting.add(
            f"{col_letter}2:{col_letter}{data_end_row}", rule
        )

    # 合計行
    total_row = data_end_row + 1
    ws.cell(row=total_row, column=1, value="合計")
    for kind in ["予算", "FC", "実績"]:
        r_col = header_map[f"室数_{kind}"]
        p_col = header_map[f"人数_{kind}"]
        s_col = header_map[f"宿泊売上_{kind}"]
        occ_col = header_map[f"OCC_{kind}"]
        adr_col = header_map[f"ADR_{kind}"]
        dor_col = header_map[f"DOR_{kind}"]
        rev_col = header_map[f"RevPAR_{kind}"]
        rl = get_column_letter(r_col)
        pl = get_column_letter(p_col)
        sl = get_column_letter(s_col)
        ws.cell(row=total_row, column=r_col).value = f"=SUM({rl}2:{rl}{data_end_row})"
        ws.cell(row=total_row, column=p_col).value = f"=SUM({pl}2:{pl}{data_end_row})"
        ws.cell(row=total_row, column=s_col).value = f"=SUM({sl}2:{sl}{data_end_row})"
        ws.cell(row=total_row, column=occ_col).value = (
            f"=IF(COUNT({rl}2:{rl}{data_end_row})=0, \"\", SUM({rl}2:{rl}{data_end_row})/{capacity}/COUNT({rl}2:{rl}{data_end_row}))"
        )
        ws.cell(row=total_row, column=adr_col).value = (
            f"=IF(COUNT({rl}2:{rl}{data_end_row})=0, \"\", SUM({sl}2:{sl}{data_end_row})/SUM({rl}2:{rl}{data_end_row}))"
        )
        ws.cell(row=total_row, column=dor_col).value = (
            f"=IF(COUNT({rl}2:{rl}{data_end_row})=0, \"\", SUM({pl}2:{pl}{data_end_row})/SUM({rl}2:{rl}{data_end_row}))"
        )
        ws.cell(row=total_row, column=rev_col).value = (
            f"=IF(COUNT({rl}2:{rl}{data_end_row})=0, \"\", SUM({sl}2:{sl}{data_end_row})/{capacity}/COUNT({rl}2:{rl}{data_end_row}))"
        )
        for col in [r_col, p_col, s_col, occ_col, adr_col, dor_col, rev_col]:
            ws.cell(row=total_row, column=col).number_format = ws.cell(row=2, column=col).number_format

    for diff_col, l, r in [
        ("差_OCC_FC-予算", "OCC_FC", "OCC_予算"),
        ("差_ADR_FC-予算", "ADR_FC", "ADR_予算"),
        ("差_売上_FC-予算", "宿泊売上_FC", "宿泊売上_予算"),
        ("差_OCC_実績-FC", "OCC_実績", "OCC_FC"),
        ("差_ADR_実績-FC", "ADR_実績", "ADR_FC"),
        ("差_売上_実績-FC", "宿泊売上_実績", "宿泊売上_FC"),
    ]:
        left_col = header_map[l]
        right_col = header_map[r]
        ltr = get_column_letter(left_col)
        rtr = get_column_letter(right_col)
        if diff_col in ["差_売上_FC-予算", "差_売上_実績-FC"]:
            range_left = f"{ltr}2:{ltr}{data_end_row}"
            range_right = f"{rtr}2:{rtr}{data_end_row}"
            ws.cell(row=total_row, column=header_map[diff_col]).value = (
                f"=IF(COUNT({range_left})=0, \"\", SUMIF({range_left}, \"<>\", {range_left})-SUMIF({range_left}, \"<>\", {range_right}))"
            )
        else:
            ws.cell(row=total_row, column=header_map[diff_col]).value = (
                f"=IF({ltr}{total_row}=\"\", \"\", {ltr}{total_row}-{rtr}{total_row})"
            )
        ws.cell(row=total_row, column=header_map[diff_col]).number_format = ws.cell(row=2, column=header_map[diff_col]).number_format
        ws.cell(row=total_row, column=header_map[diff_col]).fill = PatternFill(
            start_color="FFFAD0",
            end_color="FFFAD0",
            fill_type="solid",
        )
    # 修正月次フォーキャスト
    forecast_row = total_row + 1
    ws.cell(row=forecast_row, column=1, value="修正月次フォーキャスト")
    metrics = ["室数", "人数", "宿泊売上"]
    for m in metrics:
        fc_col = header_map[f"{m}_FC"]
        act_col = header_map[f"{m}_実績"]
        fc_letter = get_column_letter(fc_col)
        act_letter = get_column_letter(act_col)
        cell = ws.cell(row=forecast_row, column=fc_col)
        cell.value = (
            f"=SUM({act_letter}2:{act_letter}{data_end_row})+SUMIFS({fc_letter}2:{fc_letter}{data_end_row},{act_letter}2:{act_letter}{data_end_row},\"\")"
        )
        cell.number_format = ws.cell(row=2, column=fc_col).number_format

    days_count = data_end_row - 1
    ws.cell(row=forecast_row, column=header_map["OCC_FC"]).value = (
        f"=IF({get_column_letter(header_map['室数_FC'])}{forecast_row}=\"\", \"\", {get_column_letter(header_map['室数_FC'])}{forecast_row}/({capacity}*{days_count}))"
    )
    ws.cell(row=forecast_row, column=header_map["ADR_FC"]).value = (
        f"=IF(OR({get_column_letter(header_map['宿泊売上_FC'])}{forecast_row}=\"\", {get_column_letter(header_map['室数_FC'])}{forecast_row}=\"\"), \"\", {get_column_letter(header_map['宿泊売上_FC'])}{forecast_row}/{get_column_letter(header_map['室数_FC'])}{forecast_row})"
    )
    ws.cell(row=forecast_row, column=header_map["DOR_FC"]).value = (
        f"=IF(OR({get_column_letter(header_map['人数_FC'])}{forecast_row}=\"\", {get_column_letter(header_map['室数_FC'])}{forecast_row}=\"\"), \"\", {get_column_letter(header_map['人数_FC'])}{forecast_row}/{get_column_letter(header_map['室数_FC'])}{forecast_row})"
    )
    ws.cell(row=forecast_row, column=header_map["RevPAR_FC"]).value = (
        f"=IF({get_column_letter(header_map['宿泊売上_FC'])}{forecast_row}=\"\", \"\", {get_column_letter(header_map['宿泊売上_FC'])}{forecast_row}/({capacity}*{days_count}))"
    )
    for m in ["OCC", "ADR", "DOR", "RevPAR"]:
        ws.cell(row=forecast_row, column=header_map[f"{m}_FC"]).number_format = ws.cell(row=2, column=header_map[f"{m}_FC"]).number_format

    # 背景色設定
    all_metrics = ["室数", "人数", "宿泊売上", "OCC", "ADR", "DOR", "RevPAR"]
    budget_cols = [header_map[f"{m}_予算"] for m in all_metrics if f"{m}_予算" in header_map]
    fc_cols = [header_map[f"{m}_FC"] for m in all_metrics if f"{m}_FC" in header_map]
    for col in budget_cols:
        for r in range(2, forecast_row + 1):
            ws.cell(row=r, column=col).fill = budget_fill
    for col in fc_cols:
        for r in range(2, forecast_row + 1):
            ws.cell(row=r, column=col).fill = fc_fill

    diff_cols = [
        header_map["差_OCC_FC-予算"],
        header_map["差_ADR_FC-予算"],
        header_map["差_売上_FC-予算"],
        header_map["差_OCC_実績-FC"],
        header_map["差_ADR_実績-FC"],
        header_map["差_売上_実績-FC"],
    ]
    for col in diff_cols:
        for r in range(2, forecast_row + 1):
            ws.cell(row=r, column=col).fill = PatternFill(
                start_color="FFFAD0",
                end_color="FFFAD0",
                fill_type="solid",
            )
        col_letter = get_column_letter(col)
        neg_rule = FormulaRule(
            formula=[f"AND(ISNUMBER({col_letter}2),{col_letter}2<0)"],
            font=Font(color="FF0000"),
        )
        ws.conditional_formatting.add(
            f"{col_letter}2:{col_letter}{forecast_row}",
            neg_rule,
        )

    summary_dict[(year, month)] = {
        "sheet": ws.title,
        "total_row": total_row,
        "header_map": header_map,
        "data_end_row": data_end_row,
    }

    # 罫線設定
    block_ends = [header_map["RevPAR_予算"], header_map["RevPAR_FC"], header_map["RevPAR_実績"]]
    for r in ws.iter_rows(min_row=1, max_row=forecast_row, max_col=ws.max_column):
        for c in r:
            c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
    for end_col in block_ends:
        for row in range(1, forecast_row + 1):
            cell = ws.cell(row=row, column=end_col)
            cell.border = Border(
                top=cell.border.top,
                bottom=cell.border.bottom,
                left=cell.border.left,
                right=medium,
            )
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.border = Border(top=medium, bottom=medium, left=cell.border.left, right=cell.border.right)
    for row_idx in [total_row, forecast_row]:
        for cell in ws[row_idx]:
            cell.border = Border(top=medium, bottom=medium, left=cell.border.left, right=cell.border.right)


# === 年間集計シート ===
summary = wb.create_sheet(title="年間集計")
metrics = ["室数", "人数", "宿泊売上", "OCC", "ADR", "DOR", "RevPAR"]
header = ["月"]
for m in metrics:
    header += [f"{m}_予算", f"{m}_FC", f"{m}_実績", f"差_{m}_FC-予算", f"差_{m}_実績-FC"]
summary.append(header)
summary_header_map = {c.value: i for i, c in enumerate(summary[1], start=1)}
summary.freeze_panes = "B2"

room_refs = {"予算": [], "FC": [], "実績": []}
pax_refs = {"予算": [], "FC": [], "実績": []}
sales_refs = {"予算": [], "FC": [], "実績": []}
days = []
match = re.search(r"(20\d{2})", file_path)
start_year = int(match.group(1)) if match else datetime.date.today().year
year = start_year
month = start_month
for _ in range(12):
    label = f"{year}年{month}月"
    info = summary_dict.get((year, month))
    row = [label]
    if info:
        sheet = wb[info["sheet"]]
        tr = info["total_row"]
        hmap = info["header_map"]
        days.append(info["data_end_row"] - 1)
        for m in metrics:
            b_col = hmap.get(f"{m}_予算")
            f_col = hmap.get(f"{m}_FC")
            a_col = hmap.get(f"{m}_実績")
            if b_col:
                row.append(f"='{sheet.title}'!{get_column_letter(b_col)}{tr}")
            else:
                row.append(0)
            if f_col:
                row.append(f"='{sheet.title}'!{get_column_letter(f_col)}{tr}")
            else:
                row.append(0)
            if a_col:
                row.append(f"='{sheet.title}'!{get_column_letter(a_col)}{tr}")
            else:
                row.append(0)
            row += ["", ""]
        room_refs["予算"].append(f"='{sheet.title}'!{get_column_letter(hmap['室数_予算'])}{tr}")
        room_refs["FC"].append(f"='{sheet.title}'!{get_column_letter(hmap['室数_FC'])}{tr}")
        room_refs["実績"].append(f"='{sheet.title}'!{get_column_letter(hmap['室数_実績'])}{tr}")
        pax_refs["予算"].append(f"='{sheet.title}'!{get_column_letter(hmap['人数_予算'])}{tr}")
        pax_refs["FC"].append(f"='{sheet.title}'!{get_column_letter(hmap['人数_FC'])}{tr}")
        pax_refs["実績"].append(f"='{sheet.title}'!{get_column_letter(hmap['人数_実績'])}{tr}")
        sales_refs["予算"].append(f"='{sheet.title}'!{get_column_letter(hmap['宿泊売上_予算'])}{tr}")
        sales_refs["FC"].append(f"='{sheet.title}'!{get_column_letter(hmap['宿泊売上_FC'])}{tr}")
        sales_refs["実績"].append(f"='{sheet.title}'!{get_column_letter(hmap['宿泊売上_実績'])}{tr}")
    else:
        row += [0] * (len(metrics) * 5)
        days.append(0)
    summary.append(row)
    row_idx = summary.max_row
    for idx, m in enumerate(metrics):
        base = 2 + idx * 5
        f_col_idx = base + 1
        a_col_idx = base + 2
        d1 = base + 3
        d2 = base + 4
        f_cell = get_column_letter(f_col_idx)
        b_cell = get_column_letter(base)
        a_cell = get_column_letter(a_col_idx)
        if m in ["室数", "人数", "宿泊売上"]:
            summary.cell(row=row_idx, column=d1).value = (
                f"=IF(OR({f_cell}{row_idx}=\"\", {f_cell}{row_idx}=0), \"\", {f_cell}{row_idx}-{b_cell}{row_idx})"
            )
            summary.cell(row=row_idx, column=d2).value = (
                f"=IF(OR({a_cell}{row_idx}=\"\", {a_cell}{row_idx}=0), \"\", {a_cell}{row_idx}-{f_cell}{row_idx})"
            )
        else:
            summary.cell(row=row_idx, column=d1).value = (
                f"=IF({f_cell}{row_idx}=\"\", \"\", {f_cell}{row_idx}-{b_cell}{row_idx})"
            )
            summary.cell(row=row_idx, column=d2).value = (
                f"=IF({a_cell}{row_idx}=\"\", \"\", {a_cell}{row_idx}-{f_cell}{row_idx})"
            )
        fmt = {
            "室数": "#,##0",
            "人数": "#,##0",
            "宿泊売上": "#,##0",
            "OCC": "0.0%",
            "ADR": "#,##0",
            "DOR": "0.00",
            "RevPAR": "#,##0",
        }[m]
        for col in [base, f_col_idx, a_col_idx]:
            summary.cell(row=row_idx, column=col).number_format = fmt
        for col in [d1, d2]:
            summary.cell(row=row_idx, column=col).number_format = fmt
    if month == 12:
        month = 1
        year += 1
    else:
        month += 1

# 年間合計行
total_row = summary.max_row + 1
end_row = total_row - 1
summary.cell(row=total_row, column=1, value="年間合計")
for idx, m in enumerate(metrics):
    base = 2 + idx * 5
    fc = base + 1
    act = base + 2
    diff1 = base + 3
    diff2 = base + 4
    b_letter = get_column_letter(base)
    f_letter = get_column_letter(fc)
    a_letter = get_column_letter(act)
    summary.cell(row=total_row, column=base).value = f"=SUM({b_letter}2:{b_letter}{end_row})"
    summary.cell(row=total_row, column=fc).value = f"=SUM({f_letter}2:{f_letter}{end_row})"
    summary.cell(row=total_row, column=act).value = f"=SUM({a_letter}2:{a_letter}{end_row})"
    if m in ["室数", "人数", "宿泊売上"]:
        diff1_range = f"{get_column_letter(diff1)}2:{get_column_letter(diff1)}{end_row}"
        diff2_range = f"{get_column_letter(diff2)}2:{get_column_letter(diff2)}{end_row}"
        summary.cell(row=total_row, column=diff1).value = (
            f"=IF(COUNT({diff1_range})=0, \"\", SUMIF({diff1_range}, \"<>\", {diff1_range}))"
        )
        summary.cell(row=total_row, column=diff2).value = (
            f"=IF(COUNT({diff2_range})=0, \"\", SUMIF({diff2_range}, \"<>\", {diff2_range}))"
        )
    else:
        summary.cell(row=total_row, column=diff1).value = (
            f"=IF({f_letter}{total_row}=\"\", \"\", {f_letter}{total_row}-{b_letter}{total_row})"
        )
        summary.cell(row=total_row, column=diff2).value = (
            f"=IF({a_letter}{total_row}=\"\", \"\", {a_letter}{total_row}-{f_letter}{total_row})"
        )
    fmt = {
        "室数": "#,##0",
        "人数": "#,##0",
        "宿泊売上": "#,##0",
        "OCC": "0.0%",
        "ADR": "#,##0",
        "DOR": "0.00",
        "RevPAR": "#,##0",
    }[m]
    for c in [base, fc, act, diff1, diff2]:
        summary.cell(row=total_row, column=c).number_format = fmt
days_sum = sum(days) if days else 0
for kind, offset in [("予算", 0), ("FC", 1), ("実績", 2)]:
    room_col = 2 + metrics.index("室数") * 5 + offset
    pax_col = 2 + metrics.index("人数") * 5 + offset
    sales_col = 2 + metrics.index("宿泊売上") * 5 + offset
    occ_col = 2 + metrics.index("OCC") * 5 + offset
    adr_col = 2 + metrics.index("ADR") * 5 + offset
    dor_col = 2 + metrics.index("DOR") * 5 + offset
    rev_col = 2 + metrics.index("RevPAR") * 5 + offset
    r_letter = get_column_letter(room_col)
    p_letter = get_column_letter(pax_col)
    s_letter = get_column_letter(sales_col)
    summary.cell(row=total_row, column=occ_col).value = (
        f"=IFERROR({r_letter}{total_row}/{capacity}/{days_sum}, \"\")"
    )
    summary.cell(row=total_row, column=adr_col).value = (
        f"=IFERROR({s_letter}{total_row}/{r_letter}{total_row}, \"\")"
    )
    summary.cell(row=total_row, column=dor_col).value = (
        f"=IFERROR({p_letter}{total_row}/{r_letter}{total_row}, \"\")"
    )
    summary.cell(row=total_row, column=rev_col).value = (
        f"=IFERROR({s_letter}{total_row}/{capacity}/{days_sum}, \"\")"
    )
    for c in [occ_col, adr_col, dor_col, rev_col]:
        summary.cell(row=total_row, column=c).number_format = summary.cell(row=2, column=c).number_format

block_ends = [
    2 + metrics.index("RevPAR") * 5,  # 予算ブロック最後
    2 + metrics.index("RevPAR") * 5 + 1,  # FCブロック最後
    2 + metrics.index("RevPAR") * 5 + 2,  # 実績ブロック最後
]
for r in summary.iter_rows(min_row=1, max_row=total_row, max_col=summary.max_column):
    for c in r:
        c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
for end_col in block_ends:
    for row in range(1, total_row + 1):
        cell = summary.cell(row=row, column=end_col)
        cell.border = Border(
            top=cell.border.top,
            bottom=cell.border.bottom,
            left=cell.border.left,
            right=medium,
        )
for cell in summary[1]:
    cell.font = Font(bold=True)
    cell.border = Border(top=medium, bottom=medium, left=cell.border.left, right=cell.border.right)
for cell in summary[total_row]:
    cell.border = Border(top=medium, bottom=medium, left=cell.border.left, right=cell.border.right)

budget_cols = [2 + i * 5 for i in range(len(metrics))]
fc_cols = [c + 1 for c in budget_cols]
for col in budget_cols:
    for r in range(2, total_row + 1):
        summary.cell(row=r, column=col).fill = budget_fill
for col in fc_cols:
    for r in range(2, total_row + 1):
        summary.cell(row=r, column=col).fill = fc_fill

diff_cols = [
    summary_header_map[f"差_{m}_FC-予算"] for m in metrics
] + [
    summary_header_map[f"差_{m}_実績-FC"] for m in metrics
]
for col in diff_cols:
    for r in range(2, total_row + 1):
        summary.cell(row=r, column=col).fill = PatternFill(
            start_color="FFFAD0",
            end_color="FFFAD0",
            fill_type="solid",
        )
    col_letter = get_column_letter(col)
    neg_rule = FormulaRule(
        formula=[f"AND(ISNUMBER({col_letter}2),{col_letter}2<0)"],
        font=Font(color="FF0000"),
    )
    summary.conditional_formatting.add(
        f"{col_letter}2:{col_letter}{total_row}",
        neg_rule,
    )

# === 保存 ===
match = re.search(r"(20\d{2})", file_path)
year_str = match.group(1) if match else str(datetime.date.today().year)
out_path = f"予実管理表_{year_str}年度.xlsx"
wb.save(out_path)
print(f"✅ 出力完了: {out_path}")

