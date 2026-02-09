"""
訪問看護・介護スケジュール自動生成システム

入力CSVからケア予定と利用者マスタを結合し、
保険種別に応じた色分け付きExcelを出力する。
"""

import os
import shutil
import sys
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# --- 定数 ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "01_input")
OUTPUT_DIR = os.path.join(BASE_DIR, "02_output")
ARCHIVE_DIR = os.path.join(BASE_DIR, "99_archive")

REQUIRED_FILES = ["staff_shift.csv", "care_plan.csv", "medical_master.csv"]

# 色定義 (保険種別ごとの背景色)
FILL_DISABILITY = PatternFill(start_color="D8BFD8", end_color="D8BFD8", fill_type="solid")  # 紫: 障がい
FILL_MEDICAL = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")      # 青: 医療
FILL_DEFAULT = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")      # 緑: デフォルト


def ensure_directories():
    """必要なフォルダが存在しない場合は作成する。"""
    for d in [INPUT_DIR, OUTPUT_DIR, ARCHIVE_DIR]:
        os.makedirs(d, exist_ok=True)


def check_input_files():
    """入力フォルダ内に必要な3つのCSVが揃っているか確認する。"""
    missing = [f for f in REQUIRED_FILES if not os.path.isfile(os.path.join(INPUT_DIR, f))]
    if missing:
        print(f"エラー: 以下の入力ファイルが見つかりません: {', '.join(missing)}")
        print(f"  入力フォルダ: {INPUT_DIR}")
        sys.exit(1)


def load_csv(filename, description):
    """CSVファイルを読み込み、DataFrameとして返す。"""
    filepath = os.path.join(INPUT_DIR, filename)
    try:
        df = pd.read_csv(filepath, encoding="utf-8-sig")
        print(f"  {description} ({filename}): {len(df)} 件読み込み")
        return df
    except UnicodeDecodeError:
        try:
            df = pd.read_csv(filepath, encoding="cp932")
            print(f"  {description} ({filename}): {len(df)} 件読み込み (cp932)")
            return df
        except Exception as e:
            print(f"エラー: {description} ({filename}) の読み込みに失敗しました。")
            print(f"  原因: {e}")
            sys.exit(1)
    except Exception as e:
        print(f"エラー: {description} ({filename}) の読み込みに失敗しました。")
        print(f"  原因: {e}")
        sys.exit(1)


def normalize_boolean(value):
    """TRUE/FALSE 文字列やブール値を Python bool に変換する。"""
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().upper() == "TRUE"
    return False


def merge_data(care_plan, medical_master):
    """ケア予定にマスタ情報を結合する。"""
    merged = care_plan.merge(medical_master, on="利用者名", how="left")
    print(f"  結合後レコード数: {len(merged)} 件")
    return merged


def determine_fill(row):
    """行データから適用すべき背景色を返す (優先順位付き)。"""
    if normalize_boolean(row.get("判定_障がい", False)):
        return FILL_DISABILITY
    if normalize_boolean(row.get("判定_医療", False)):
        return FILL_MEDICAL
    return FILL_DEFAULT


def export_excel(merged_df, staff_shift_df, timestamp_str):
    """結合データを色分け付きExcelとして出力する。"""
    filename = f"週間スケジュール_{timestamp_str}.xlsx"
    filepath = os.path.join(OUTPUT_DIR, filename)

    try:
        with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
            merged_df.to_excel(writer, sheet_name="週間スケジュール", index=False)
            staff_shift_df.to_excel(writer, sheet_name="スタッフ出勤情報", index=False)
    except PermissionError:
        print(f"エラー: ファイルへの書き込み権限がありません: {filepath}")
        sys.exit(1)
    except Exception as e:
        print(f"エラー: Excel出力に失敗しました。")
        print(f"  原因: {e}")
        sys.exit(1)

    # 色分け処理
    try:
        wb = load_workbook(filepath)
        ws = wb["週間スケジュール"]

        # ヘッダー行からカラム位置を取得
        headers = {cell.value: cell.column for cell in ws[1]}

        for row_idx in range(2, ws.max_row + 1):
            row_data = {}
            for col_name, col_num in headers.items():
                row_data[col_name] = ws.cell(row=row_idx, column=col_num).value

            fill = determine_fill(row_data)
            for col_num in range(1, ws.max_column + 1):
                ws.cell(row=row_idx, column=col_num).fill = fill

        # 列幅の自動調整
        for sheet in wb.worksheets:
            for column_cells in sheet.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter
                for cell in column_cells:
                    if cell.value is not None:
                        # 日本語は幅が広いため係数を掛ける
                        cell_len = sum(2 if ord(c) > 127 else 1 for c in str(cell.value))
                        max_length = max(max_length, cell_len)
                sheet.column_dimensions[column_letter].width = min(max_length + 2, 50)

        wb.save(filepath)
        print(f"  出力ファイル: {filepath}")
    except Exception as e:
        print(f"エラー: Excel色分け処理に失敗しました。")
        print(f"  原因: {e}")
        sys.exit(1)

    return filepath


def archive_inputs(timestamp_str):
    """処理済みの入力ファイルをアーカイブフォルダに移動する。"""
    for filename in REQUIRED_FILES:
        src = os.path.join(INPUT_DIR, filename)
        dst = os.path.join(ARCHIVE_DIR, f"{timestamp_str}_{filename}")
        try:
            shutil.move(src, dst)
            print(f"  アーカイブ: {filename} -> {os.path.basename(dst)}")
        except Exception as e:
            print(f"警告: {filename} のアーカイブに失敗しました。")
            print(f"  原因: {e}")


def main():
    """メイン処理フロー。"""
    print("=" * 60)
    print("訪問看護・介護スケジュール自動生成システム")
    print("=" * 60)

    now = datetime.now()
    timestamp_str = now.strftime("%Y%m%d_%H%M%S")

    # 1. 初期化
    print("\n[1/6] フォルダ初期化...")
    ensure_directories()
    print("  完了")

    # 2. ファイル確認
    print("\n[2/6] 入力ファイル確認...")
    check_input_files()
    print("  必要なファイルがすべて揃っています")

    # 3. データ読み込み
    print("\n[3/6] データ読み込み...")
    staff_shift = load_csv("staff_shift.csv", "スタッフ出勤情報")
    care_plan = load_csv("care_plan.csv", "ケア予定")
    medical_master = load_csv("medical_master.csv", "利用者マスタ")

    # 4. データ結合
    print("\n[4/6] データ結合...")
    merged = merge_data(care_plan, medical_master)

    # 5. Excel生成 & 6. 色分け処理
    print("\n[5/6] Excel出力 & 色分け処理...")
    output_path = export_excel(merged, staff_shift, timestamp_str)

    # 7. アーカイブ処理
    print("\n[6/6] アーカイブ処理...")
    archive_inputs(timestamp_str)

    print("\n" + "=" * 60)
    print("処理が正常に完了しました。")
    print(f"  出力先: {output_path}")
    print("=" * 60)


if __name__ == "__main__":
    main()
