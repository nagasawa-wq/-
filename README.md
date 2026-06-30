"""
売上明細（決済額明細）を Site ID 末尾2桁ごとに分けて
サマリー付きのExcelファイルを作成するスクリプト。

使い方:
    python split_sales_by_siteid.py 入力ファイル.xlsx [出力ファイル.xlsx] [--site-names site_names.json]

例:
    python split_sales_by_siteid.py 76208800_株式会社1_2_3_20260401-20260430.xlsx
    python split_sales_by_siteid.py 入力.xlsx 出力.xlsx --site-names sites.json

site_names.json の形式（任意）:
    {
        "76208801": "リスタート",
        "76208802": "High Cloud"
    }
"""

import sys
import json
import argparse
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# データ抽出
# ---------------------------------------------------------------------------
def parse_sales(filepath: str) -> pd.DataFrame:
    """《決済額明細》セクションから「売上確定」の行だけを抽出する。"""
    df_raw = pd.read_excel(filepath, header=None)

    rows = []
    in_sales = False
    for _, row in df_raw.iterrows():
        val = str(row[0]).strip() if pd.notna(row[0]) else ""

        if val == "《決済額明細》":
            in_sales = True
            continue

        # Site ID は 7 や 8 から始まる数字想定（必要に応じて調整）
        if in_sales and val[:1].isdigit():
            result = str(row[7]).strip() if pd.notna(row[7]) else ""
            if result == "売上確定":
                rows.append(
                    {
                        "Site ID": str(int(row[0])),
                        "Credit Card No.": str(row[2]).strip() if pd.notna(row[2]) else "",
                        "Transaction ID": str(row[4]).strip() if pd.notna(row[4]) else "",
                        "Day Time": str(row[6]).strip() if pd.notna(row[6]) else "",
                        "Result": result,
                        "Amount": int(row[8]) if pd.notna(row[8]) else 0,
                    }
                )

        # 次のセクションに入ったら終了
        if in_sales and (
            val.startswith("《ﾁｬｰｼﾞ")
            or val.startswith("《取消")
            or val.startswith("《認証")
        ):
            break

    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# スタイル定義
# ---------------------------------------------------------------------------
TITLE_COLOR = "2E5B8F"
HEADER_FILL = PatternFill("solid", start_color=TITLE_COLOR)
TOTAL_FILL = PatternFill("solid", start_color="D9E1F2")
ALT_FILL = PatternFill("solid", start_color="EEF2F9")
BORDER = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
NORMAL_FONT = Font(name="Arial", size=10)
TOTAL_FONT = Font(name="Arial", bold=True, size=10, color=TITLE_COLOR)


def make_header_row(ws, row_num, cols_data):
    for col, (val, width) in enumerate(cols_data, 1):
        cell = ws.cell(row=row_num, column=col, value=val)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(col)].width = width


# ---------------------------------------------------------------------------
# Excel 作成
# ---------------------------------------------------------------------------
def build_workbook(df: pd.DataFrame, period: str, file_label: str, site_names: dict) -> Workbook:
    df = df.copy()
    df["last2"] = df["Site ID"].str[-2:]
    df["SiteName"] = df["Site ID"].map(site_names).fillna("")
    digits = sorted(df["last2"].unique())

    wb = Workbook()

    # ---- サマリーシート ----
    ws = wb.active
    ws.title = "サマリー"
    ws.merge_cells("A1:G1")
    t = ws["A1"]
    t.value = f"Site ID 末尾2桁別　売上サマリー（{period}）　{file_label}"
    t.font = Font(name="Arial", bold=True, size=13, color=TITLE_COLOR)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    headers = [
        ("Site ID 末尾2桁", 16),
        ("サイト名", 22),
        ("件数", 10),
        ("売上合計 (円)", 18),
        ("構成比", 12),
        ("平均単価 (円)", 18),
        ("Site ID", 14),
    ]
    make_header_row(ws, 2, headers)
    ws.row_dimensions[2].height = 20

    total_row_num = len(digits) + 3
    for r_idx, digit in enumerate(digits, 3):
        sub = df[df["last2"] == digit]
        count = len(sub)
        amount = sub["Amount"].sum()
        avg = round(amount / count) if count else 0
        site_name = sub["SiteName"].iloc[0]
        site_id = sub["Site ID"].iloc[0]

        fill = ALT_FILL if r_idx % 2 == 0 else PatternFill()
        vals = [f"末尾 {digit}", site_name, count, amount, f"=D{r_idx}/D{total_row_num}", avg, site_id]
        aligns = ["center", "left", "center", "right", "center", "right", "center"]
        fmts = [None, None, None, "#,##0", "0.0%", "#,##0", None]

        for col, (val, align, fmt) in enumerate(zip(vals, aligns, fmts), 1):
            cell = ws.cell(row=r_idx, column=col, value=val)
            cell.font = NORMAL_FONT
            cell.border = BORDER
            cell.alignment = Alignment(horizontal=align, vertical="center")
            if fmt:
                cell.number_format = fmt
            if fill.fill_type:
                cell.fill = fill
        ws.row_dimensions[r_idx].height = 18

    ws.row_dimensions[total_row_num].height = 20
    total_vals = [
        "合計",
        "",
        f"=SUM(C3:C{total_row_num - 1})",
        f"=SUM(D3:D{total_row_num - 1})",
        "",
        f"=D{total_row_num}/C{total_row_num}",
        "",
    ]
    total_aligns = ["center", "left", "center", "right", "center", "right", "center"]
    total_fmts = [None, None, None, "#,##0", None, "#,##0", None]
    for col, (val, align, fmt) in enumerate(zip(total_vals, total_aligns, total_fmts), 1):
        cell = ws.cell(row=total_row_num, column=col, value=val)
        cell.font = TOTAL_FONT
        cell.fill = TOTAL_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal=align, vertical="center")
        if fmt:
            cell.number_format = fmt

    # ---- 末尾ごとの明細シート ----
    for digit in digits:
        sub = df[df["last2"] == digit].copy()
        site_name = sub["SiteName"].iloc[0]
        site_id = sub["Site ID"].iloc[0]

        sheet = wb.create_sheet(title=f"末尾{digit}")
        sheet.merge_cells("A1:F1")
        t = sheet["A1"]
        label = f"（{site_id} {site_name}）" if site_name else f"（{site_id}）"
        t.value = f"Site ID 末尾「{digit}」{label}の売上明細"
        t.font = Font(name="Arial", bold=True, size=12, color=TITLE_COLOR)
        t.alignment = Alignment(horizontal="center", vertical="center")
        sheet.row_dimensions[1].height = 28

        cols = [
            ("Site ID", 14),
            ("Credit Card No.", 22),
            ("Transaction ID", 20),
            ("Day Time", 20),
            ("Result", 12),
            ("Amount (円)", 14),
        ]
        make_header_row(sheet, 2, cols)
        sheet.row_dimensions[2].height = 20

        for r_idx, (_, row) in enumerate(sub.iterrows(), 3):
            fill = ALT_FILL if r_idx % 2 == 0 else PatternFill()
            vals = [
                row["Site ID"],
                row["Credit Card No."],
                row["Transaction ID"],
                row["Day Time"],
                row["Result"],
                row["Amount"],
            ]
            aligns = ["center", "left", "left", "center", "center", "right"]
            for col, (val, align) in enumerate(zip(vals, aligns), 1):
                cell = sheet.cell(row=r_idx, column=col, value=val)
                cell.font = NORMAL_FONT
                cell.border = BORDER
                cell.alignment = Alignment(horizontal=align, vertical="center")
                if col == 6:
                    cell.number_format = "#,##0"
                if fill.fill_type:
                    cell.fill = fill
            sheet.row_dimensions[r_idx].height = 16

        last_data_row = len(sub) + 2
        total_r = last_data_row + 1
        sheet.row_dimensions[total_r].height = 20
        for col in range(1, 7):
            cell = sheet.cell(row=total_r, column=col)
            cell.fill = TOTAL_FILL
            cell.border = BORDER
            cell.font = TOTAL_FONT
        sheet["A" + str(total_r)] = "合計"
        sheet["A" + str(total_r)].alignment = Alignment(horizontal="center", vertical="center")
        sheet["F" + str(total_r)] = f"=SUM(F3:F{last_data_row})"
        sheet["F" + str(total_r)].number_format = "#,##0"
        sheet["F" + str(total_r)].alignment = Alignment(horizontal="right", vertical="center")

    return wb


# ---------------------------------------------------------------------------
# メイン
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Site ID 末尾2桁別 売上集計Excelを作成")
    parser.add_argument("input", help="入力Excelファイルパス（決済額明細を含むもの）")
    parser.add_argument("output", nargs="?", default=None, help="出力Excelファイルパス（省略可）")
    parser.add_argument("--site-names", default=None, help="Site ID→サイト名 のJSONファイルパス（任意）")
    parser.add_argument("--period", default=None, help="集計期間ラベル（例: 2026年4月）省略時はファイル名から推測")
    parser.add_argument("--label", default=None, help="ファイルラベル（例: 株式会社1.2.3）省略時はファイル名を使用")
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"エラー: ファイルが見つかりません: {input_path}")
        sys.exit(1)

    site_names = {}
    if args.site_names:
        with open(args.site_names, encoding="utf-8") as f:
            site_names = json.load(f)

    period = args.period or "対象期間"
    label = args.label or input_path.stem

    df = parse_sales(str(input_path))
    if df.empty:
        print("警告: 売上確定データが見つかりませんでした。")
        sys.exit(1)

    wb = build_workbook(df, period, label, site_names)

    output_path = (
        Path(args.output)
        if args.output
        else Path.cwd() / f"売上_SiteID末尾2桁別_{input_path.stem}.xlsx"
    )
    wb.save(output_path)

    print(f"完了: {output_path}")
    print(f"総件数: {len(df)}件 / 総売上: {df['Amount'].sum():,}円")


if __name__ == "__main__":
    main()
