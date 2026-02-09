"""
訪問看護・介護スケジュール自動生成システム - CareRoute Optimizer
(Multi-Building Edition)

Excel形式の入力ファイルをフォルダ監視型で読み込み、
スコアリング方式（負荷平準化 + 建物間動線最適化）で
1タスク=1スタッフの割り当てを行い、
ガントチャート風マトリクス形式のExcelを出力する。

建物間ルール:
  - C棟は隔離棟: 他の棟との行き来不可
  - 同一建物: -30pt (ボーナス)
  - 同一建物かつ同一階: さらに -20pt
  - 同一建物で異なる階: 階差 × 5pt (ペナルティ)
  - 異なる建物間 (A↔B等): +20pt (ペナルティ)
"""

import glob
import os
import random
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
    xlsx_files = [f for f in xlsx_files if not os.path.basename(f).startswith("~$")]

    if not xlsx_files:
        print(f"[エラー] {folder_name}フォルダにExcelファイルが見つかりません")
        print(f"  対象フォルダ: {folder_path}")
        sys.exit(1)

    if len(xlsx_files) == 1:
        chosen = xlsx_files[0]
        print(f"  {label}: {os.path.basename(chosen)}")
        return chosen

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


def safe_str(value, default="不明"):
    """値を安全に文字列化する。NaN や None は default を返す。"""
    if pd.isna(value):
        return default
    return str(value).strip()


def safe_int(value, default=0):
    """値を安全に整数化する。変換できなければ default を返す。"""
    if pd.isna(value):
        return default
    try:
        return int(value)
    except (ValueError, TypeError):
        return default


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
    if hasattr(time_str, "hour"):
        return time_str if hasattr(time_str, "second") else None
    s = str(time_str).strip()
    try:
        return datetime.strptime(s, "%H:%M").time()
    except ValueError:
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
# Greedy 割り当てロジック
# ====================================================================

def is_movement_allowed(prev_building, curr_building):
    """建物間の移動が許可されるか判定する。

    C棟は隔離棟のため、他の棟との行き来は不可。
    C棟 → C棟 のみ許可。他の棟同士 (A↔B等) は許可。
    """
    if not prev_building or not curr_building:
        return True
    if prev_building == curr_building:
        return True
    if prev_building == "C棟" or curr_building == "C棟":
        return False
    return True


def calculate_movement_score(task_floor, task_building, last_floor, last_building):
    """建物・階を考慮した動線スコアを計算する。

    スコアが低いほど動線が良い (ボーナスはマイナス、ペナルティはプラス)。
    - 初回訪問 (前回情報なし): 0pt
    - 同一建物: -30pt
      - さらに同一階: -20pt (合計 -50pt)
      - 異なる階: +階差×5pt
    - 異なる建物: +20pt
    """
    if last_building is None and last_floor is None:
        return 0

    if last_building and task_building and last_building == task_building:
        # 同一建物ボーナス
        score = -30
        if last_floor is not None and task_floor is not None:
            if last_floor == task_floor:
                score -= 20  # 同一階追加ボーナス
            else:
                score += abs(task_floor - last_floor) * 5  # 階差ペナルティ
        return score
    else:
        # 異なる建物ペナルティ
        return 20


def assign_tasks_for_date(day_staff_df, day_tasks, date_display, carry_over):
    """1日分のタスクをスコアリング方式で割り当てる。

    Score = (累積ランク合計 * 10) - 動線ボーナス
    スコアが低いスタッフが優先的に選ばれる。
    累積ランクは日をまたいで引き継ぎ、負荷を全日程で平準化する。

    Args:
        carry_over: dict[staff_name] -> {"cumulative_rank": int, ...}
            前日までの累積状態。この日の出勤者のみ使用する。

    Returns:
        assignments: dict[staff_name] -> list of (slot_indices, task_row)
        warnings: list of str
        day_stats: dict[staff_name] -> {"count": int, "total_rank": int}
        carry_over: 更新済みの累積状態
    """
    warnings = []

    # スタッフごとの情報を構造化
    staff_info = {}
    for _, row in day_staff_df.iterrows():
        name = row["スタッフ名"]
        staff_info[name] = {
            "shift_start": str(row["開始時間"]).strip(),
            "shift_end": str(row["終了時間"]).strip(),
            "ng_range": row["NG時間帯"],
            "occupied_slots": set(),
        }

    # 累積状態を引き継ぐ (初出のスタッフは0初期化)
    for name in staff_info:
        if name not in carry_over:
            carry_over[name] = {"cumulative_rank": 0}

    # 当日の動線追跡 (日ごとにリセット)
    staff_last_floor = {name: None for name in staff_info}
    staff_last_building = {name: None for name in staff_info}

    assignments = {name: [] for name in staff_info}
    day_stats = {name: {"count": 0, "total_rank": 0} for name in staff_info}

    # タスクを順に処理
    for _, task_row in day_tasks.iterrows():
        fixed_time = str(task_row["固定時間指定"]).strip()
        start_idx = time_to_slot_index(fixed_time)
        if start_idx is None:
            continue

        try:
            duration_min = int(task_row["所要時間"])
        except (ValueError, TypeError):
            duration_min = 30
        slot_count = max(1, duration_min // 30)

        needed_slots = list(range(start_idx, min(start_idx + slot_count, len(TIME_SLOTS))))

        user_name = safe_str(task_row["利用者名"])
        task_floor = safe_int(task_row.get("階数", None), default=0)
        task_building = safe_str(task_row.get("建物名", None), default="")
        task_rank = safe_int(task_row.get("ランク", None), default=1)

        # 候補スタッフをリストアップ
        candidates = []
        for staff_name, info in staff_info.items():
            all_in_shift = all(
                is_within_shift(TIME_SLOTS[idx], info["shift_start"], info["shift_end"])
                for idx in needed_slots
            )
            if not all_in_shift:
                continue

            any_ng = any(
                is_in_ng_range(TIME_SLOTS[idx], info["ng_range"])
                for idx in needed_slots
            )
            if any_ng:
                continue

            if info["occupied_slots"] & set(needed_slots):
                continue

            # C棟隔離チェック: 前回の建物からの移動が許可されるか
            if not is_movement_allowed(staff_last_building[staff_name], task_building):
                continue

            candidates.append(staff_name)

        if not candidates:
            slot_label = TIME_SLOTS[start_idx]
            bldg_info = f" ({task_building})" if task_building else ""
            msg = f"  [警告] {date_display} {slot_label}の{user_name}様{bldg_info}を担当できるスタッフがいません"
            warnings.append(msg)
            continue

        # スコアリングで最適スタッフを選択
        # Score = 累積ランク*10 + 動線スコア + ランダムジッター(0-5)
        def calc_score(name):
            movement = calculate_movement_score(
                task_floor, task_building,
                staff_last_floor[name], staff_last_building[name],
            )
            jitter = random.uniform(0, 5)
            return (carry_over[name]["cumulative_rank"] * 10) + movement + jitter

        best = min(candidates, key=calc_score)

        # 割り当て実行
        staff_info[best]["occupied_slots"].update(needed_slots)
        staff_last_floor[best] = task_floor
        staff_last_building[best] = task_building
        carry_over[best]["cumulative_rank"] += task_rank
        day_stats[best]["count"] += 1
        day_stats[best]["total_rank"] += task_rank
        assignments[best].append((needed_slots, task_row))

    return assignments, warnings, day_stats, carry_over


# ====================================================================
# Excel生成 (マトリクスレイアウト)
# ====================================================================

def build_matrix_excel(staff_shift_df, merged_df, timestamp_str):
    """割り当て結果をガントチャート風マトリクスExcelとして出力する。"""
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

    # --- 日付ごとに割り当て → 書き込み ---
    dates = staff_shift_df["日付"].unique()
    current_row = 2
    all_warnings = []
    # 全日程を通じたスタッフ別統計・累積状態
    global_stats = {}
    carry_over = {}  # 日をまたいで累積ランクを引き継ぐ

    # ケアプランを曜日別に整理
    tasks_by_weekday = {}
    for _, task_row in merged_df.iterrows():
        weekday = str(task_row["曜日"]).strip()
        if weekday not in tasks_by_weekday:
            tasks_by_weekday[weekday] = []
        tasks_by_weekday[weekday].append(task_row)

    for date_val in sorted(dates):
        try:
            date_obj = pd.to_datetime(date_val)
            weekday_jp = WEEKDAY_JP[date_obj.weekday()]
            date_display = date_obj.strftime("%Y-%m-%d")
        except Exception:
            weekday_jp = None
            date_display = str(date_val)

        day_staff_df = staff_shift_df[staff_shift_df["日付"] == date_val]

        # この曜日のタスクを取得してソート済みDataFrameにする
        day_task_list = tasks_by_weekday.get(weekday_jp, [])
        if day_task_list:
            day_tasks_df = pd.DataFrame(day_task_list)
        else:
            day_tasks_df = pd.DataFrame()

        # スコアリング割り当て実行
        if not day_tasks_df.empty:
            assignments, warnings, day_stats, carry_over = assign_tasks_for_date(
                day_staff_df, day_tasks_df, date_display, carry_over
            )
            all_warnings.extend(warnings)
            for name, stats in day_stats.items():
                if name not in global_stats:
                    global_stats[name] = {"count": 0, "total_rank": 0}
                global_stats[name]["count"] += stats["count"]
                global_stats[name]["total_rank"] += stats["total_rank"]
        else:
            assignments = {row["スタッフ名"]: [] for _, row in day_staff_df.iterrows()}

        # 各スタッフの行を書き込む
        for _, staff_row in day_staff_df.iterrows():
            staff_name = staff_row["スタッフ名"]

            # A列: 日付, B列: スタッフ名
            date_cell = ws.cell(row=current_row, column=1, value=date_display)
            date_cell.border = THIN_BORDER
            date_cell.font = FONT_CELL
            date_cell.alignment = ALIGN_CENTER

            name_cell = ws.cell(row=current_row, column=2, value=staff_name)
            name_cell.border = THIN_BORDER
            name_cell.font = FONT_CELL
            name_cell.alignment = ALIGN_CENTER

            # 全スロットに罫線を設定 + NG時間帯に「休憩」表示
            ng_range = staff_row["NG時間帯"]
            for slot_idx in range(len(TIME_SLOTS)):
                col = slot_idx + 3
                cell = ws.cell(row=current_row, column=col)
                cell.border = THIN_BORDER
                cell.font = FONT_CELL
                cell.alignment = ALIGN_CENTER
                if is_in_ng_range(TIME_SLOTS[slot_idx], ng_range):
                    cell.value = "休憩"
                    cell.fill = FILL_HEADER

            # 割り当て済みタスクを書き込む
            for slot_indices, task_row in assignments.get(staff_name, []):
                user_name = safe_str(task_row["利用者名"])
                room = safe_str(task_row.get("部屋番号", None), default="")
                building = safe_str(task_row.get("建物名", None), default="")
                service = safe_str(task_row.get("サービス種類", None), default="")
                # 表示: 利用者名 / 建物名 部屋番号 (サービス種類)
                location_parts = []
                if building:
                    location_parts.append(building)
                if room:
                    location_parts.append(room)
                location = " ".join(location_parts)
                if location and service:
                    display_text = f"{user_name}\n{location} ({service})"
                elif location:
                    display_text = f"{user_name}\n{location}"
                else:
                    display_text = user_name
                fill = get_fill_for_user(task_row)

                for idx in slot_indices:
                    col = idx + 3
                    cell = ws.cell(row=current_row, column=col)
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

    return filepath, all_warnings, global_stats


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
    print("\n[1/7] フォルダ初期化...")
    ensure_directories()
    print("  完了")

    # 2. 入力ファイル探索
    print("\n[2/7] 入力ファイル探索...")
    shift_file = find_excel_in_folder("01_shift")
    care_file = find_excel_in_folder("02_care")
    master_file = find_excel_in_folder("03_master")
    used_files = [shift_file, care_file, master_file]

    # 3. データ読み込み
    print("\n[3/7] データ読み込み...")
    staff_shift = load_excel(shift_file, "01_shift")
    care_plan = load_excel(care_file, "02_care")
    medical_master = load_excel(master_file, "03_master")

    # 4. データ結合
    print("\n[4/7] データ結合...")
    merged = care_plan.merge(medical_master, on="利用者名", how="left")

    # オプション列が無い場合に備えてデフォルト値を補填
    if "階数" not in merged.columns:
        merged["階数"] = 0
        print("  [情報] マスタに「階数」列がないため、全件 0 で補填しました")
    if "部屋番号" not in merged.columns:
        merged["部屋番号"] = ""
        print("  [情報] マスタに「部屋番号」列がないため、空文字で補填しました")
    if "建物名" not in merged.columns:
        merged["建物名"] = ""
        print("  [情報] マスタに「建物名」列がないため、空文字で補填しました")
    if "ランク" not in merged.columns:
        merged["ランク"] = 1
        print("  [情報] ケアプランに「ランク」列がないため、全件 1 で補填しました")

    # ソート: 固定時間指定 (昇順) > 建物名 (昇順) > 階数 (昇順)
    # 時間順で処理することで、スタッフが同一建物・同一階に留まりやすくなる
    merged["_sort_time"] = merged["固定時間指定"].apply(
        lambda v: safe_str(v, "99:99")
    )
    merged["_sort_building"] = merged["建物名"].apply(lambda v: safe_str(v, ""))
    merged["_sort_floor"] = merged["階数"].apply(lambda v: safe_int(v, 0))
    merged = merged.sort_values(
        ["_sort_time", "_sort_building", "_sort_floor"]
    ).drop(columns=["_sort_time", "_sort_building", "_sort_floor"])

    print(f"  ケアプラン + 利用者マスタ → {len(merged)} 件 (時間→建物→階数でソート済)")

    # 5. 割り当て & Excel生成
    print("\n[5/7] スコアリング割り当て & マトリクスExcel生成...")
    output_path, warnings, global_stats = build_matrix_excel(
        staff_shift, merged, timestamp_str
    )

    # 6. 割り当て結果サマリ
    print("\n[6/7] 割り当て結果...")
    if warnings:
        print(f"  割り当て不能タスク: {len(warnings)} 件")
        for w in warnings[:5]:
            print(w)
        if len(warnings) > 5:
            print(f"  ...他 {len(warnings) - 5} 件")
    else:
        print("  全タスクを正常に割り当てました")

    # 負荷分散サマリ
    if global_stats:
        print("\n  --- スタッフ別 負荷サマリ ---")
        print(f"  {'スタッフ名':　<10s}  件数  トータルランク")
        print(f"  {'-' * 36}")
        for name in sorted(global_stats.keys()):
            s = global_stats[name]
            print(f"  {name:　<10s}  {s['count']:>4d}  {s['total_rank']:>14d}")
        total_count = sum(s["count"] for s in global_stats.values())
        total_rank = sum(s["total_rank"] for s in global_stats.values())
        print(f"  {'-' * 36}")
        print(f"  {'合計':　<10s}  {total_count:>4d}  {total_rank:>14d}")

    # 7. アーカイブ
    print("\n[7/7] アーカイブ処理...")
    archive_files(used_files, timestamp_str)

    print("\n" + "=" * 60)
    print("処理が正常に完了しました。")
    print(f"  出力先: {output_path}")
    print("=" * 60)


if __name__ == "__main__":
    main()
