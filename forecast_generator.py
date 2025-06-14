import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.formatting.rule import FormulaRule
import tkinter as tk
from tkinter import filedialog, simpledialog
import calendar
import datetime
import jpholiday

# === GUIでキャパシティと期首月の取得 ===
root = tk.Tk()
root.withdraw()
capacity = simpledialog.askinteger("キャパシティ入力", "客室キャパシティ（部屋数）を入力してください：")
start_month = simpledialog.askinteger("期首月入力", "期首月（例：1〜12）を入力してください：")
file_path = filedialog.askopenfilename(title="日別予算Excelファイルを選択")

# === Excel読み込み ===
xls = pd.read_excel(file_path, sheet_name=None)

# === 出力用Excel作成 ===
wb = openpyxl.Workbook()
wb.remove(wb.active)

# === 曜日装飾 ===
sat_fill = PatternFill(start_color="DDEEFF", end_color="DDEEFF", fill_type="solid")
sat_font = Font(color="003366")
sun_fill = PatternFill(start_color="FFE5E5", end_color="FFE5E5", fill_type="solid")
sun_font = Font(color="990000")
gray_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")

# === 各月シート作成 ===
for sheet_name, df in xls.items():
    if not isinstance(df, pd.DataFrame) or df.empty:
        continue

    df['日付'] = pd.to_datetime(df['日付'], errors='coerce')
    df = df.dropna(subset=['日付'])

    # 日付ごとに横持ち1行
    rows = []
    for date in df['日付'].dt.date.unique():
        weekday = date.weekday()
        weekday_jp = ['月', '火', '水', '木', '金', '土', '日'][weekday]
        if jpholiday.is_holiday(date):
            weekday_jp = '祝'
        base = {'日付': date.strftime('%Y/%m/%d'), '曜日': weekday_jp}
        for kind in ['予算', 'FC', '実績']:
            subset = df[(df['日付'] == pd.Timestamp(date)) & (df['種別'] == kind)]
            if not subset.empty:
                row = subset.iloc[0]
                base[f'室数_{kind}'] = row['室数']
                base[f'人数_{kind}'] = row['人数']
                base[f'宿泊売上_{kind}'] = row['宿泊売上']
        rows.append(base)

    df_out = pd.DataFrame(rows)

    # 数式列を追加
    for kind in ['予算', 'FC', '実績']:
        df_out[f'OCC_{kind}'] = ''
        df_out[f'ADR_{kind}'] = ''
        df_out[f'DOR_{kind}'] = ''
        df_out[f'RevPAR_{kind}'] = ''

    # 差異列
    df_out['差_OCC_FC-予算'] = ''
    df_out['差_ADR_FC-予算'] = ''
    df_out['差_売上_FC-予算'] = ''
    df_out['差_OCC_実績-FC'] = ''
    df_out['差_ADR_実績-FC'] = ''
    df_out['差_売上_実績-FC'] = ''

    # Excelシートへ出力
    ws = wb.create_sheet(title=sheet_name)
    for r in dataframe_to_rows(df_out, index=False, header=True):
        ws.append(r)

    # 曜日装飾（色分け）
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=2):
        date_cell, weekday_cell = row
        try:
            date_obj = datetime.datetime.strptime(date_cell.value, "%Y/%m/%d").date()
        except Exception:
            continue

        if jpholiday.is_holiday(date_obj) or weekday_cell.value in ['日', '祝']:
            weekday_cell.fill = sun_fill
            weekday_cell.font = sun_font
        elif weekday_cell.value == '土':
            weekday_cell.fill = sat_fill
            weekday_cell.font = sat_font

    # 実績が入力されていたら、FC欄をグレー化（条件付き書式）
    headers = [cell.value for cell in ws[1]]
    for i, col in enumerate(headers):
        if col == '室数_FC':
            fc_col = openpyxl.utils.get_column_letter(i+1)
        if col == '室数_実績':
            actual_col = openpyxl.utils.get_column_letter(i+1)

    formula = f'LEN(${actual_col}2)>0'
    for i in range(0, 7):  # 室数〜RevPARまでの列
        col_letter = openpyxl.utils.get_column_letter(openpyxl.utils.column_index_from_string(fc_col) + i)
        rule = FormulaRule(formula=[formula], fill=gray_fill)
        ws.conditional_formatting.add(f"{col_letter}2:{col_letter}{ws.max_row}", rule)

# === 年間集計シート ===
summary = wb.create_sheet(title="年間集計")
summary.append(['月', '予算_売上', 'FC_売上', '実績_売上'])
month_order = [(start_month + i - 1) % 12 + 1 for i in range(12)]
for m in month_order:
    month_label = f"{m}月"
    summary.append([month_label, '', '', ''])

# === 保存 ===
year_str = ''.join(filter(str.isdigit, file_path))
out_path = f"予実管理表_{year_str}年度.xlsx"
wb.save(out_path)
print(f"✅ 出力完了: {out_path}")
