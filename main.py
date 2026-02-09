"""
訪問看護・介護スケジュール自動生成システム - CareRoute Optimizer

Excel形式の入力ファイルをフォルダ監視型で読み込み、
横軸に時間(30分刻み)・縦軸にスタッフを配置した
ガントチャート風マトリクス形式のExcelを出力する。
"""

import glob
import os
import shutil
import sys
from datetime import datetime, timedelta

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ====================================================================
# 定数
# ====================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "01_input")
SHIFT_DIR = os.path.join(INPUT_DIR, "01_shift")
CARE_DIR = os.path.join(INPUT_DIR, "02_care")
MASTER_DIR = os.path.join(INPUT_DIR, "03_master")
OUTPUT_DIR = os.path.join(BASE_DIR, "02_output")
ARCHIVE_DIR = os.path.join(BASE_DIR, "99_archive")

ALL_DIRS = [SHIFT_DIR, CARE_DIR, MASTER_DIR, OUTPUT_DIR, ARCHIVE_DIR]

# 各フォルダの表示名と必須カラム
FOLDER_CONFIG = {
    "01_shift": {
        "path": SHIFT_DIR,
        "label": "シフト情報",
        "required_columns": ["日付", "スタッフ名", "開始時間", "終了時間", "NG時間帯"],
    },
    "02_care": {
        "path": CARE_DIR,
        "label": "ケアプラン",
        "required_columns": [
            "利用者名", "曜日", "頻度", "固定時間指定", "所要時間", "サービス種類",
        ],
    },
    "03_master": {
        "path": MASTER_DIR,
        "label": "利用者マスタ",
        "required_columns": ["利用者名", "住所", "判定_医療", "判定_介護", "判定_障がい"],
    },
}

# 時間軸: 09:00 - 18:00 (30分刻み)
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
# フォルダ初期化・ファイル探索
# ====================================================================

def ensure_directories():
    """必要なフォルダが存在しない場合は作成する。"""
    for d in ALL_DIRS:
        os.makedirs(d, exist_ok=True)


def find_excel_in_folder(folder_key):
    """指定フォルダ内の .xlsx ファイルを探索し、パスを1つ返す。

    - ファイルが0個 → エラー終了
    - ファイルが1個 → そのまま返す
    - ファイルが複数 → 更新日時が最新のものを警告付きで返す
    """
    config = FOLDER_CONFIG[folder_key]
    folder_path = config["path"]
    label = config["label"]
    folder_name = folder_key

    xlsx_files = sorted(glob.glob(os.path.join(folder_path, "*.xlsx")))
    # 一時ファイル (~$...) を除外
    xlsx_files = [f for f in xlsx_files if not os.path.basename(f).startswith("~$")]

    if not xlsx_files:
        print(f"[エラー] {folder_name}フォルダにExcelファイルが見つかりません")
        print(f"  対象フォルダ: {folder_path}")
        sys.exit(1)

    if len(xlsx_files) == 1:
        chosen = xlsx_files[0]
        print(f"  {label}: {os.path.basename(chosen)}")
        return chosen

    # 複数ファイルがある場合 → 更新日時が最新のものを選択
    xlsx_files.sort(key=lambda f: os.path.getmtime(f), reverse=True)
    chosen = xlsx_files[0]
    print(f"  [警告] {folder_name}フォルダに複数のExcelファイルがあります ({len(xlsx_files)}件)")
    for f in xlsx_files:
        mtime = datetime.fromtimestamp(os.path.getmtime(f)).strftime("%Y-%m-%d %H:%M:%S")
        marker = " ← 採用" if f == chosen else ""
        print(f"    - {os.path.basename(f)} (更新: {mtime}){marker}")
    return chosen


def load_excel(filepath, folder_key):
    """Excelファイルの1シート目を読み込み、カラム検証して DataFrame を返す。"""
    config = FOLDER_CONFIG[folder_key]
    label = config["label"]
    filename = os.path.basename(filepath)

    try:
        df = pd.read_excel(filepath, sheet_name=0, engine="openpyxl")
    except Exception as e:
        print(f"[エラー] {label} ({filename}) の読み込みに失敗しました。")
        print(f"  ファイルが破損しているか、Excel形式ではない可能性があります。")
        print(f"  原因: {e}")
        sys.exit(1)

    # 必須カラムの存在チェック
    required = config["required_columns"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        print(f"[エラー] {label} ({filename}) に必要な列が不足しています。")
        print(f"  不足列: {', '.join(missing)}")
        print(f"  存在する列: {', '.join(df.columns.tolist())}")
        sys.exit(1)

    print(f"  {label} ({filename}): {len(df)} 件読み込み")
    return df


# ====================================================================
# データ変換ユーティリティ
# ====================================================================

def normalize_boolean(value):
    """TRUE/FALSE 文字列・ブール値・1/0 を Python bool に変換する。"""
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value == 1
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
    # datetime オブジェクトがそのまま入っている場合に対応
    if hasattr(time_str, "hour"):
        return time_str if hasattr(time_str, "second") else None
    try:
        return datetime.strptime(s, "%H:%M").time()
    except ValueError:
        # "09:00:00" のような秒付きフォーマットにも対応
        try:
            return datetime.strptime(s, "%H:%M:%S").time()
        except ValueError:
            return None


def time_to_slot_index(time_str):
    """時間文字列を TIME_SLOTS 内のインデックスに変換する。"""
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
# Excel生成 (マトリクスレイアウト)
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

    # --- ケアプランを曜日別に整理 ---
    tasks_by_weekday = {}
    for _, task_row in merged_df.iterrows():
        weekday = str(task_row["曜日"]).strip()
        if weekday not in tasks_by_weekday:
            tasks_by_weekday[weekday] = []
        tasks_by_weekday[weekday].append(task_row)

    # --- 日付ごと・スタッフごとに行を生成 ---
    dates = staff_shift_df["日付"].unique()
    current_row = 2

    for date_val in sorted(dates):
        # 日付から曜日を取得
        try:
            date_obj = pd.to_datetime(date_val)
            weekday_jp = WEEKDAY_JP[date_obj.weekday()]
            date_display = date_obj.strftime("%Y-%m-%d")
        except Exception:
            weekday_jp = None
            date_display = str(date_val)

        day_staff = staff_shift_df[staff_shift_df["日付"] == date_val]

        for _, staff_row in day_staff.iterrows():
            staff_name = staff_row["スタッフ名"]
            shift_start = str(staff_row["開始時間"]).strip()
            shift_end = str(staff_row["終了時間"]).strip()
            ng_range = staff_row["NG時間帯"]

            # A列: 日付, B列: スタッフ名
            date_cell = ws.cell(row=current_row, column=1, value=date_display)
            date_cell.border = THIN_BORDER
            date_cell.font = FONT_CELL
            date_cell.alignment = ALIGN_CENTER

            name_cell = ws.cell(row=current_row, column=2, value=staff_name)
            name_cell.border = THIN_BORDER
            name_cell.font = FONT_CELL
            name_cell.alignment = ALIGN_CENTER

            # 各タイムスロットのセルを初期化 (罫線のみ)
            for slot_idx in range(len(TIME_SLOTS)):
                col = slot_idx + 3
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

                    # 開始時刻がシフト時間内か・NG時間帯でないか
                    slot_label = TIME_SLOTS[start_idx]
                    if not is_within_shift(slot_label, shift_start, shift_end):
                        continue
                    if is_in_ng_range(slot_label, ng_range):
                        continue

                    user_name = str(task_row["利用者名"])
                    service = str(task_row["サービス種類"])
                    display_text = f"{user_name}\n({service})"
                    fill = get_fill_for_user(task_row)

                    # 所要時間分のスロットにタスクを書き込む
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
                        if cell.value:
                            cell.value = f"{cell.value}\n---\n{display_text}"
                        else:
                            cell.value = display_text
                        cell.fill = fill

            current_row += 1

    # --- 列幅・行高さ調整 ---
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 14
    for slot_idx in range(len(TIME_SLOTS)):
        col_letter = get_column_letter(slot_idx + 3)
        ws.column_dimensions[col_letter].width = 18
    for row_idx in range(2, current_row):
        ws.row_dimensions[row_idx].height = 50

    # --- 保存 ---
    filename = f"週間スケジュール_{timestamp_str}.xlsx"
    filepath = os.path.join(OUTPUT_DIR, filename)
    try:
        wb.save(filepath)
        print(f"  出力ファイル: {filepath}")
    except PermissionError:
        print(f"[エラー] ファイルへの書き込み権限がありません: {filepath}")
        sys.exit(1)
    except Exception as e:
        print(f"[エラー] Excel出力に失敗しました。")
        print(f"  原因: {e}")
        sys.exit(1)

    return filepath


# ====================================================================
# アーカイブ処理
# ====================================================================

def archive_files(file_paths, timestamp_str):
    """処理に使用した入力ファイルをアーカイブフォルダへ移動する。"""
    for src in file_paths:
        original_name = os.path.basename(src)
        dst = os.path.join(ARCHIVE_DIR, f"{timestamp_str}_{original_name}")
        try:
            shutil.move(src, dst)
            print(f"  アーカイブ: {original_name} -> {os.path.basename(dst)}")
        except Exception as e:
            print(f"  [警告] {original_name} のアーカイブに失敗しました。")
            print(f"    原因: {e}")


# ====================================================================
# メイン処理
# ====================================================================

def main():
    print("=" * 60)
    print("CareRoute Optimizer - 訪問看護・介護スケジュール自動生成")
    print("=" * 60)

    now = datetime.now()
    timestamp_str = now.strftime("%Y%m%d_%H%M%S")

    # 1. フォルダ初期化
    print("\n[1/6] フォルダ初期化...")
    ensure_directories()
    print("  完了")

    # 2. 入力ファイル探索
    print("\n[2/6] 入力ファイル探索...")
    shift_file = find_excel_in_folder("01_shift")
    care_file = find_excel_in_folder("02_care")
    master_file = find_excel_in_folder("03_master")
    used_files = [shift_file, care_file, master_file]

    # 3. データ読み込み
    print("\n[3/6] データ読み込み...")
    staff_shift = load_excel(shift_file, "01_shift")
    care_plan = load_excel(care_file, "02_care")
    medical_master = load_excel(master_file, "03_master")

    # 4. データ結合
    print("\n[4/6] データ結合...")
    merged = care_plan.merge(medical_master, on="利用者名", how="left")
    print(f"  ケアプラン + 利用者マスタ → {len(merged)} 件")

    # 5. マトリクスExcel生成 & 色分け
    print("\n[5/6] マトリクスExcel生成 & 色分け...")
    output_path = build_matrix_excel(staff_shift, merged, timestamp_str)

    # 6. アーカイブ
    print("\n[6/6] アーカイブ処理...")
    archive_files(used_files, timestamp_str)

    print("\n" + "=" * 60)
    print("処理が正常に完了しました。")
    print(f"  出力先: {output_path}")
    print("=" * 60)


if __name__ == "__main__":
    main()
