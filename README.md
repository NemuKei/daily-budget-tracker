# 📊 Daily Budget Tracker

ホテル向けの **日別予算・ローリングフォーキャスト・オンハンド・実績比較** を支援するExcel出力ツールです。  
予算と実績の差異を視覚化し、ローリングで着地予測や修正を行う業務を効率化します。

---

## 🚀 主な機能

- ✅ **横持ち形式**の日別予実フォーマットをExcelで自動生成  
- ✅ **予算 / FC（ローリングフォーキャスト） / OH（オンハンド） / 実績** を1行にまとめて比較  
- ✅ **OCC・ADR・RevPAR・DOR** をExcel関数で自動計算  
- ✅ **差異列（FC−予算、OH−FC、実績−FC、実績−予算）** を自動生成  
- ✅ 実績入力済みの日付に対して **FC・OHセルを自動グレー化**（条件付き書式）  
- ✅ 曜日列に応じた **色分け（例：土曜＝青、日祝＝赤）**  
- ✅ 月別シートに加えて、**年間集計シート（縦持ち形式）**を自動生成  
- ✅ **年間差異シート**も差異タイプ別にブロック表示＋年間合計を自動計算  
- ✅ GUIで **キャパシティ（部屋数）・期首月** を指定可能  
- ✅ `.exe` 化で非エンジニアでも利用可能  

---

## 📂 フォルダ構成（例）

```
daily-budget-tracker/
├── forecast_generator.py       # メインスクリプト（Excel出力）
├── README.md
├── requirements.txt            # 必要なPythonライブラリ
└── sample/
    └── 日別予算_2025.xlsx     # 入力用サンプルExcelファイル
```

---

## 🛠️ 使用方法

1. Python 3.x をインストール（推奨：3.10以上）  
2. 必要ライブラリをインストール：

```bash
pip install pandas openpyxl jpholiday
```

※Tkinter が入っていない場合、`python3-tk` の追加が必要です。

3. スクリプトを実行：

```bash
python forecast_generator.py
```

4. GUIで以下を入力：
- 宿泊キャパシティ（部屋数）
- 期首月（1〜12）
- 入力ファイル（例：`日別予算_2025.xlsx`）

5. 自動で `予実管理表_202X年度.xlsx` が出力されます。

---

## 🧮 入力ファイル（フォーマット要件）

- Excelファイル（例：`日別予算_2025.xlsx`）
- 各シートに対象月のデータを配置
- 以下の列を含むこと：

| 日付 | 種別（予算のみ） | 室数 | 人数 | 宿泊売上 |
|------|------------------|------|------|----------|

---

## 📅 出力形式（例）

| 日付       | 曜日 | 室数_予算 | ... | 室数_FC | ... | 室数_OH | ... | 室数_実績 | ... | 差_売上_実績-FC |
|------------|------|------------|-----|----------|-----|----------|-----|------------|-----|------------------|
| 2025/04/01 | 火   | 50         |     | 48       |     | 47       |     | 45         |     | -3,000           |

---

## 📊 年間シート構成

### 📘 年間集計シート
- **縦持ち形式で「月」列＋指標行（室数, 人数, 売上, OCC, ADR, etc.）**
- ブロック単位（予算 / FC / OH / 実績）で視認性向上
- 年間合計行付き（OCC等は再計算）

### 📕 年間差異シート
- 差異タイプ別にブロック表示（例：FC−予算、実績−予算、OH−FC、実績−FC）
- 年間合計行追加＋罫線付き
- 未入力月は除外、エラー値は空白化

---

## 📌 注意点

- 実績が未入力の日程は、月次フォーキャストに **OH値を加算**して修正着地を予測  
- 本ツールは宿泊部門用に最適化済み（朝食などは非対応）  
- 差異分析・視覚化に特化し、エクセルベースでの現場利用に最適

---
