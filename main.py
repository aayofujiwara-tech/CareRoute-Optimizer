"""
訪問看護・介護スケジュール自動生成システム - CareRoute Optimizer

横軸に時間(30分刻み)、縦軸にスタッフを配置した
ガントチャート風マトリクス形式のExcelを出力する。
"""

import os
import shutil
import sys
from datetime import datetime, timedelta

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --- 定数 ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "01_input")
OUTPUT_DIR = os.path.join(BASE_DIR, "02_output")
ARCHIVE_DIR = os.path.join(BASE_DIR, "99_archive")

REQUIRED_FILES = ["staff_shift.csv", "care_plan.csv", "medical_master.csv"]

REQUIRED_COLUMNS = {
    "staff_shift.csv": ["日付", "スタッフ名", "開始時間", "終了時間", "NG時間帯"],
    "care_plan.csv": ["利用者名", "曜日", "頻度", "固定時間指定", "所要時間", "サービス種類"],
    "medical_master.csv": ["利用者名", "住所", "判定_医療", "判定_介護", "判定_障がい"],
}

# 時間軸: 09:00 - 18:00 (30分刻み, 19スロット)
TIME_SLOTS = []
_t = datetime(2000, 1, 1, 9, 0)
while _t <= datetime(2000, 1, 1, 18, 0):
    TIME_SLOTS.append(_t.strftime("%H:%M"))
    _t += timedelta(minutes=30)

# 曜日変換 (Python weekday -> 日本語)
WEEKDAY_JP = {0: "月", 1: "火", 2: "水", 3: "木", 4: "金", 5: "土", 6: "日"}

# スタイル定義
FILL_HEADER = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
FILL_DISABILITY = PatternFill(start_color="D8BFD8", end_color="D8BFD8", fill_type="solid")
FILL_MEDICAL = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
FILL_CARE = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
FONT_HEADER = Font(bold=True, size=10)
FONT_CELL = Font(size=9)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


# ====================================================================
# ユーティリティ関数
# ====================================================================

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
    except UnicodeDecodeError:
        try:
            df = pd.read_csv(filepath, encoding="cp932")
            print(f"  {description} ({filename}): {len(df)} 件読み込み (cp932)")
        except Exception as e:
            print(f"エラー: {description} ({filename}) の読み込みに失敗しました。")
            print(f"  原因: {e}")
            sys.exit(1)
    except Exception as e:
        print(f"エラー: {description} ({filename}) の読み込みに失敗しました。")
        print(f"  原因: {e}")
        sys.exit(1)
    else:
        print(f"  {description} ({filename}): {len(df)} 件読み込み")

    # 必要な列の存在チェック
    required = REQUIRED_COLUMNS[filename]
    missing_cols = [c for c in required if c not in df.columns]
    if missing_cols:
        print(f"エラー: {filename} に必要な列が不足しています: {', '.join(missing_cols)}")
        print(f"  存在する列: {', '.join(df.columns)}")
        sys.exit(1)

    return df


def normalize_boolean(value):
    """TRUE/FALSE 文字列やブール値を Python bool に変換する。"""
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().upper() == "TRUE"
    return False


def get_fill_for_user(row):
    """利用者の保険判定に基づく背景色を返す (優先順位付き)。"""
    if normalize_boolean(row.get("判定_障がい", False)):
        return FILL_DISABILITY
    if normalize_boolean(row.get("判定_医療", False)):
        return FILL_MEDICAL
    return FILL_CARE


def parse_time(time_str):
    """時間文字列 (HH:MM) を datetime.time に変換する。"""
    if pd.isna(time_str) or str(time_str).strip() == "":
        return None
    s = str(time_str).strip()
    try:
        return datetime.strptime(s, "%H:%M").time()
    except ValueError:
        return None


def time_to_slot_index(time_str):
    """時間文字列 (HH:MM) を TIME_SLOTS 内のインデックスに変換する。"""
    t = parse_time(time_str)
    if t is None:
        return None
    label = t.strftime("%H:%M")
    if label in TIME_SLOTS:
        return TIME_SLOTS.index(label)
    return None


def is_in_ng_range(slot_label, ng_str):
    """指定スロットがNG時間帯に含まれるか判定する。"""
    if pd.isna(ng_str) or str(ng_str).strip() == "":
        return False
    slot_time = parse_time(slot_label)
    if slot_time is None:
        return False
    for ng_part in str(ng_str).split(","):
        ng_part = ng_part.strip()
        if "-" not in ng_part:
            continue
        parts = ng_part.split("-")
        ng_start = parse_time(parts[0].strip())
        ng_end = parse_time(parts[1].strip())
        if ng_start and ng_end and ng_start <= slot_time < ng_end:
            return True
    return False


def is_within_shift(slot_label, shift_start_str, shift_end_str):
    """指定スロットがスタッフの勤務時間内か判定する。"""
    slot_time = parse_time(slot_label)
    start = parse_time(shift_start_str)
    end = parse_time(shift_end_str)
    if slot_time is None or start is None or end is None:
        return False
    return start <= slot_time < end


# ====================================================================
# Excel生成
# ====================================================================

def build_matrix_excel(staff_shift_df, merged_df, timestamp_str):
    """ガントチャート風マトリクスExcelを生成する。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "週間スケジュール"

    # --- ヘッダー行 ---
    headers = ["日付", "スタッフ名"] + TIME_SLOTS
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.border = THIN_BORDER
        cell.alignment = ALIGN_CENTER

    # --- 日付ごとにスタッフ行を作成 ---
    # 日付ごとのグループ化
    dates = staff_shift_df["日付"].unique()

    # care_plan を曜日別に整理 (固定時間指定ありのもの)
    tasks_by_weekday = {}
    for _, task_row in merged_df.iterrows():
        weekday = str(task_row["曜日"]).strip()
        if weekday not in tasks_by_weekday:
            tasks_by_weekday[weekday] = []
        tasks_by_weekday[weekday].append(task_row)

    current_row = 2

    for date_str in sorted(dates):
        # 日付から曜日を取得
        try:
            date_obj = pd.to_datetime(date_str)
            weekday_jp = WEEKDAY_JP[date_obj.weekday()]
        except Exception:
            weekday_jp = None

        # この日に出勤しているスタッフ一覧
        day_staff = staff_shift_df[staff_shift_df["日付"] == date_str]

        for _, staff_row in day_staff.iterrows():
            staff_name = staff_row["スタッフ名"]
            shift_start = str(staff_row["開始時間"]).strip()
            shift_end = str(staff_row["終了時間"]).strip()
            ng_range = staff_row["NG時間帯"]

            # A列: 日付, B列: スタッフ名
            date_cell = ws.cell(row=current_row, column=1, value=date_str)
            date_cell.border = THIN_BORDER
            date_cell.font = FONT_CELL
            date_cell.alignment = ALIGN_CENTER

            name_cell = ws.cell(row=current_row, column=2, value=staff_name)
            name_cell.border = THIN_BORDER
            name_cell.font = FONT_CELL
            name_cell.alignment = ALIGN_CENTER

            # 各タイムスロットのセルを初期化 (罫線のみ)
            for slot_idx, slot_label in enumerate(TIME_SLOTS):
                col = slot_idx + 3  # C列から
                cell = ws.cell(row=current_row, column=col)
                cell.border = THIN_BORDER
                cell.font = FONT_CELL
                cell.alignment = ALIGN_CENTER

            # タスクの配置
            if weekday_jp and weekday_jp in tasks_by_weekday:
                for task_row in tasks_by_weekday[weekday_jp]:
                    fixed_time = str(task_row["固定時間指定"]).strip()
                    start_idx = time_to_slot_index(fixed_time)
                    if start_idx is None:
                        continue

                    # 所要時間からスロット数を計算
                    try:
                        duration_min = int(task_row["所要時間"])
                    except (ValueError, TypeError):
                        duration_min = 30
                    slot_count = max(1, duration_min // 30)

                    # シフト時間内か・NG時間帯でないかチェック
                    slot_label = TIME_SLOTS[start_idx]
                    if not is_within_shift(slot_label, shift_start, shift_end):
                        continue
                    if is_in_ng_range(slot_label, ng_range):
                        continue

                    # セルに書き込む表示テキスト
                    user_name = str(task_row["利用者名"])
                    service = str(task_row["サービス種類"])
                    display_text = f"{user_name}\n({service})"
                    fill = get_fill_for_user(task_row)

                    # 所要時間分のスロットを塗る
                    for offset in range(slot_count):
                        idx = start_idx + offset
                        if idx >= len(TIME_SLOTS):
                            break
                        target_slot = TIME_SLOTS[idx]
                        if not is_within_shift(target_slot, shift_start, shift_end):
                            break
                        if is_in_ng_range(target_slot, ng_range):
                            break

                        col = idx + 3
                        cell = ws.cell(row=current_row, column=col)
                        # 既にタスクが入っている場合は追記
                        if cell.value:
                            cell.value = f"{cell.value}\n---\n{display_text}"
                        else:
                            cell.value = display_text
                        cell.fill = fill

            current_row += 1

    # --- 列幅調整 ---
    ws.column_dimensions["A"].width = 14  # 日付
    ws.column_dimensions["B"].width = 14  # スタッフ名
    for slot_idx in range(len(TIME_SLOTS)):
        col_letter = get_column_letter(slot_idx + 3)
        ws.column_dimensions[col_letter].width = 18  # タイムスロット

    # 行の高さ調整
    for row_idx in range(2, current_row):
        ws.row_dimensions[row_idx].height = 50

    # --- 保存 ---
    filename = f"週間スケジュール_マトリクス_{timestamp_str}.xlsx"
    filepath = os.path.join(OUTPUT_DIR, filename)
    try:
        wb.save(filepath)
        print(f"  出力ファイル: {filepath}")
    except PermissionError:
        print(f"エラー: ファイルへの書き込み権限がありません: {filepath}")
        sys.exit(1)
    except Exception as e:
        print(f"エラー: Excel出力に失敗しました。")
        print(f"  原因: {e}")
        sys.exit(1)

    return filepath


# ====================================================================
# アーカイブ処理
# ====================================================================

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


# ====================================================================
# メイン処理
# ====================================================================

def main():
    """メイン処理フロー。"""
    print("=" * 60)
    print("CareRoute Optimizer - 訪問看護・介護スケジュール自動生成")
    print("=" * 60)

    now = datetime.now()
    timestamp_str = now.strftime("%Y%m%d_%H%M%S")
    date_str = now.strftime("%Y%m%d")

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
    merged = care_plan.merge(medical_master, on="利用者名", how="left")
    print(f"  結合後レコード数: {len(merged)} 件")

    # 5. Excel生成 (マトリクス形式) & 6. 色分け処理
    print("\n[5/6] マトリクスExcel生成 & 色分け...")
    output_path = build_matrix_excel(staff_shift, merged, date_str)

    # 7. アーカイブ処理
    print("\n[6/6] アーカイブ処理...")
    archive_inputs(timestamp_str)

    print("\n" + "=" * 60)
    print("処理が正常に完了しました。")
    print(f"  出力先: {output_path}")
    print("=" * 60)


if __name__ == "__main__":
    main()
