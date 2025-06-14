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
thin = Side(style="thin")
medium = Side(style="medium")

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
            weekday_jp = "祝"

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
        adr_c.value = f"={sales_c}{row}/{pax_c}{row}"
        dor_c.value = f"={pax_c}{row}/{room_c}{row}"
        rev_c.value = f"={sales_c}{row}/{capacity}"

        occ_c.number_format = "0.0%"
        adr_c.number_format = "#,##0"
        dor_c.number_format = "0.00"
        rev_c.number_format = "#,##0"

        # FC列の数式
        fc_room_c = get_column_letter(header_map["室数_FC"])
        fc_pax_c = get_column_letter(header_map["人数_FC"])
        fc_sales_c = get_column_letter(header_map["宿泊売上_FC"])
        fc_occ = ws.cell(row=row, column=header_map["OCC_FC"])
        fc_adr = ws.cell(row=row, column=header_map["ADR_FC"])
        fc_dor = ws.cell(row=row, column=header_map["DOR_FC"])
        fc_rev = ws.cell(row=row, column=header_map["RevPAR_FC"])

        fc_occ.value = f"=IFERROR({fc_room_c}{row}/{capacity}, \"\")"
        fc_adr.value = f"=IFERROR({fc_sales_c}{row}/{fc_pax_c}{row}, \"\")"
        fc_dor.value = f"=IFERROR({fc_pax_c}{row}/{fc_room_c}{row}, \"\")"
        fc_rev.value = f"=IFERROR({fc_sales_c}{row}/{capacity}, \"\")"

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

        act_occ.value = f"=IFERROR({act_room_c}{row}/{capacity}, \"\")"
        act_adr.value = f"=IFERROR({act_sales_c}{row}/{act_pax_c}{row}, \"\")"
        act_dor.value = f"=IFERROR({act_pax_c}{row}/{act_room_c}{row}, \"\")"
        act_rev.value = f"=IFERROR({act_sales_c}{row}/{capacity}, \"\")"

        act_occ.number_format = "0.0%"
        act_adr.number_format = "#,##0"
        act_dor.number_format = "0.00"
        act_rev.number_format = "#,##0"

        # 手入力セルの表示形式
        for col_name in ["室数_FC", "人数_FC", "宿泊売上_FC", "室数_実績", "人数_実績", "宿泊売上_実績"]:
            ws.cell(row=row, column=header_map[col_name]).number_format = "#,##0"

        # 差異列
        ws.cell(row=row, column=header_map["差_OCC_FC-予算"]).value = (
            f"=IFERROR({get_column_letter(header_map['OCC_FC'])}{row}-"
            f"{get_column_letter(header_map['OCC_予算'])}{row}, \"\")"
        )
        ws.cell(row=row, column=header_map["差_ADR_FC-予算"]).value = (
            f"=IFERROR({get_column_letter(header_map['ADR_FC'])}{row}-"
            f"{get_column_letter(header_map['ADR_予算'])}{row}, \"\")"
        )
        ws.cell(row=row, column=header_map["差_売上_FC-予算"]).value = (
            f"=IFERROR({get_column_letter(header_map['宿泊売上_FC'])}{row}-"
            f"{get_column_letter(header_map['宿泊売上_予算'])}{row}, \"\")"
        )
        ws.cell(row=row, column=header_map["差_OCC_実績-FC"]).value = (
            f"=IFERROR({get_column_letter(header_map['OCC_実績'])}{row}-"
            f"{get_column_letter(header_map['OCC_FC'])}{row}, \"\")"
        )
        ws.cell(row=row, column=header_map["差_ADR_実績-FC"]).value = (
            f"=IFERROR({get_column_letter(header_map['ADR_実績'])}{row}-"
            f"{get_column_letter(header_map['ADR_FC'])}{row}, \"\")"
        )
        ws.cell(row=row, column=header_map["差_売上_実績-FC"]).value = (
            f"=IFERROR({get_column_letter(header_map['宿泊売上_実績'])}{row}-"
            f"{get_column_letter(header_map['宿泊売上_FC'])}{row}, \"\")"
        )

        for diff_col in ["差_OCC_FC-予算", "差_OCC_実績-FC"]:
            ws.cell(row=row, column=header_map[diff_col]).number_format = "0.0%"
        for diff_col in ["差_ADR_FC-予算", "差_ADR_実績-FC"]:
            ws.cell(row=row, column=header_map[diff_col]).number_format = "#,##0"
        for diff_col in ["差_売上_FC-予算", "差_売上_実績-FC"]:
            ws.cell(row=row, column=header_map[diff_col]).number_format = "#,##0"

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
    for col in range(3, ws.max_column + 1):
        letter = get_column_letter(col)
        ws.cell(row=total_row, column=col).value = f"=SUM({letter}2:{letter}{data_end_row})"

    # 修正月次フォーキャスト
    forecast_row = total_row + 1
    ws.cell(row=forecast_row, column=1, value="修正月次フォーキャスト")
    metrics = ["室数", "人数", "宿泊売上", "OCC", "ADR", "DOR", "RevPAR"]
    for m in metrics:
        fc_col = header_map[f"{m}_FC"]
        act_col = header_map[f"{m}_実績"]
        fc_letter = get_column_letter(fc_col)
        act_letter = get_column_letter(act_col)
        cell = ws.cell(row=forecast_row, column=fc_col)
        cell.value = f"=SUM({act_letter}2:{act_letter}{data_end_row})+SUMIFS({fc_letter}2:{fc_letter}{data_end_row},{act_letter}2:{act_letter}{data_end_row},\"\")"
        cell.number_format = ws.cell(row=2, column=fc_col).number_format

    summary_dict[(year, month)] = {
        "sheet": ws.title,
        "total_row": total_row,
        "header_map": header_map,
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
metrics = ["宿泊売上", "室数", "人数", "OCC", "ADR", "DOR", "RevPAR"]
kinds = ["予算", "FC", "実績"]
header = ["月"] + [f"{k}_{m}" for k in kinds for m in metrics]
summary.append(header)
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
        total_row = info["total_row"]
        hmap = info["header_map"]
        for k in kinds:
            for m in metrics:
                if k == "予算":
                    col = hmap.get(f"{m}_予算")
                elif k == "FC":
                    col = hmap.get(f"{m}_FC")
                else:
                    col = hmap.get(f"{m}_実績")
                if col:
                    letter = get_column_letter(col)
                    row.append(f"='{sheet.title}'!{letter}{total_row}")
                else:
                    row.append(0)
    else:
        row += [0] * (len(metrics) * len(kinds))
    summary.append(row)
    if month == 12:
        month = 1
        year += 1
    else:
        month += 1

# 年間合計行
total_row = summary.max_row + 1
end_row = total_row - 1
summary.cell(row=total_row, column=1, value="年間合計")
for col in range(2, summary.max_column + 1):
    letter = get_column_letter(col)
    summary.cell(row=total_row, column=col).value = f"=SUM({letter}2:{letter}{end_row})"

# === 保存 ===
match = re.search(r"(20\d{2})", file_path)
year_str = match.group(1) if match else str(datetime.date.today().year)
out_path = f"予実管理表_{year_str}年度.xlsx"
wb.save(out_path)
print(f"✅ 出力完了: {out_path}")

