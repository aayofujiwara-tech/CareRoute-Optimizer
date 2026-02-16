"""
訪問看護・介護スケジュール自動生成システム - CareRoute Optimizer
(Multi-Building Edition)

Excel形式の入力ファイルをフォルダ監視型で読み込み、
2段階方式（固定タスク → フリータスク自動配置）の
スコアリング割り当て（負荷平準化 + 建物間動線最適化）を行い、
ガントチャート風マトリクス形式のExcelを出力する。

建物間ルール:
  - C棟は隔離棟: 他の棟との行き来不可
  - 同一建物: -30pt (ボーナス)
  - 同一建物かつ同一階: さらに -20pt
  - 同一建物で異なる階: 階差 × 5pt (ペナルティ)
  - 異なる建物間 (A↔B等): +30pt (基礎) + 当日移動回数×40pt (累積)
"""

import glob
import os
import random
import re
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
INTEGRATED_DIR = os.path.join(INPUT_DIR, "00_integrated")
SHIFT_DIR = os.path.join(INPUT_DIR, "01_shift")
CARE_DIR = os.path.join(INPUT_DIR, "02_care")
MASTER_DIR = os.path.join(INPUT_DIR, "03_master")
ROOM_DIR = os.path.join(INPUT_DIR, "04_room")
OUTPUT_DIR = os.path.join(BASE_DIR, "02_output")
ARCHIVE_DIR = os.path.join(BASE_DIR, "99_archive")

ALL_DIRS = [INTEGRATED_DIR, SHIFT_DIR, CARE_DIR, MASTER_DIR, ROOM_DIR, OUTPUT_DIR, ARCHIVE_DIR]

# 曜日名の正規化マッピング
WEEKDAY_NAMES = ["月", "火", "水", "木", "金", "土", "日"]

# デフォルトシフト設定
DEFAULT_SHIFT_START = "09:00"
DEFAULT_SHIFT_END = "18:00"
DEFAULT_SHIFT_NG = "12:00-13:00"

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

# 時間軸: 09:00 - 17:30 (30分刻み, 実例準拠)
TIME_SLOTS = []
_t = datetime(2000, 1, 1, 9, 0)
while _t <= datetime(2000, 1, 1, 17, 30):
    TIME_SLOTS.append(_t.strftime("%H:%M"))
    _t += timedelta(minutes=30)

# 曜日変換 (Python weekday -> 日本語)
WEEKDAY_JP = {0: "月", 1: "火", 2: "水", 3: "木", 4: "金", 5: "土", 6: "日"}

# スタイル定義
FILL_HEADER = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
FILL_MENTAL = PatternFill(start_color="6AA84F", end_color="6AA84F", fill_type="solid")      # 精神 = 緑
FILL_MEDICAL = PatternFill(start_color="6D9EEB", end_color="6D9EEB", fill_type="solid")     # 医療 = 青
FILL_CARE = PatternFill(start_color="E69138", end_color="E69138", fill_type="solid")        # 介護 = 橙
FILL_INFORMAL = PatternFill(start_color="FFE599", end_color="FFE599", fill_type="solid")    # インフォーマル = 黄
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
        return FILL_MENTAL    # 精神 = 緑
    if normalize_boolean(row.get("判定_医療", False)):
        return FILL_MEDICAL   # 医療 = 青
    return FILL_CARE          # 介護 = 橙


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
# 統合入力パーサー
# ====================================================================

def parse_visit_schedule(visit_str):
    """訪問日の自然言語記述をパースし、構造化データのリストを返す。

    入力例:
      "週3回(月水金)\n1回30分"
      "週7回\n1日3回\n1回30分"
      "週3回(月水金)\n月金は60分\n水は30分"
      "週1回(木)\n1回30分\n9:30〜"
      "週3回(月水金)\n1回30分\n午後"
      "月火木金日\n9:30〜\n13:00〜\n16:00〜\n\n水土\n9:00〜\n16:00〜\n\n1日3回\n1回60分"

    返り値: list of dict
      [{"weekday": "月", "duration": 30, "fixed_time": None, "afternoon": False}, ...]
    """
    if pd.isna(visit_str) or str(visit_str).strip() == "":
        return []

    text = str(visit_str).strip()
    lines = [l.strip() for l in text.replace("\\n", "\n").split("\n") if l.strip()]

    # --- 基本所要時間 (全パターン共通) ---
    base_duration = 30  # デフォルト
    for line in lines:
        m = re.search(r"1回(\d+)分", line)
        if m:
            base_duration = int(m.group(1))

    # --- 1日N回 (全パターン共通) ---
    times_per_day = 1
    for line in lines:
        m = re.search(r"1日(\d+)回", line)
        if m:
            times_per_day = int(m.group(1))

    # --- 曜日グループパターン検出 ---
    # 「月火木金日」のように曜日文字のみの行があれば、グループ別固定時間モード
    # 例: 月火木金日\n9:30〜\n13:00〜\n16:00〜\n水土\n9:00〜\n16:00〜
    weekday_group_re = re.compile(r"^[月火水木金土日]+$")
    has_weekday_groups = any(weekday_group_re.match(line) for line in lines)

    if has_weekday_groups:
        groups = []
        current_weekdays = None
        current_times = []

        for line in lines:
            if weekday_group_re.match(line):
                # 前のグループを保存
                if current_weekdays is not None:
                    groups.append((current_weekdays, current_times))
                current_weekdays = list(line)
                current_times = []
            elif current_weekdays is not None:
                # 時刻パターン (9:30〜 等)
                m = re.match(r"(\d{1,2}:\d{2})[〜~]?$", line)
                if m:
                    t = m.group(1)
                    parts = t.split(":")
                    current_times.append(f"{int(parts[0]):02d}:{parts[1]}")

        # 最後のグループを保存
        if current_weekdays is not None:
            groups.append((current_weekdays, current_times))

        results = []
        for weekdays, times in groups:
            for wd in weekdays:
                if times:
                    for t in times:
                        results.append({
                            "weekday": wd,
                            "duration": base_duration,
                            "fixed_time": t,
                            "afternoon": False,
                        })
                else:
                    # 固定時間なしのグループ → times_per_day で展開
                    for _ in range(times_per_day):
                        results.append({
                            "weekday": wd,
                            "duration": base_duration,
                            "fixed_time": None,
                            "afternoon": False,
                        })
        return results

    # --- 以下: 標準パターン (週N回 形式) ---

    # --- 週N回 / 曜日の抽出 ---
    freq = 0
    weekdays = []
    for line in lines:
        # 週N回(月水金) パターン ("回" 省略可)
        m = re.match(r"週(\d+)回?[（(]([月火水木金土日]+)[)）]", line)
        if m:
            freq = int(m.group(1))
            weekdays = list(m.group(2))
            continue
        # 週N回 (曜日指定なし, "回" 省略可)
        m = re.match(r"週(\d+)回?(?!\d)", line)
        if m:
            freq = int(m.group(1))
            continue

    # 曜日指定がない場合、頻度に基づいてデフォルト割り当て
    if not weekdays:
        if freq >= 7:
            weekdays = list(WEEKDAY_NAMES)  # 全曜日
        elif freq == 6:
            weekdays = WEEKDAY_NAMES[:6]    # 月〜土
        elif freq == 5:
            weekdays = WEEKDAY_NAMES[:5]    # 月〜金
        elif freq == 4:
            weekdays = ["月", "火", "木", "金"]
        elif freq == 3:
            weekdays = ["月", "水", "金"]
        elif freq == 2:
            weekdays = ["火", "金"]
        elif freq == 1:
            weekdays = ["水"]
        else:
            weekdays = WEEKDAY_NAMES[:5]  # フォールバック: 平日

    # --- 曜日別所要時間 (例: "月金は60分", "水は30分") ---
    weekday_durations = {}
    for line in lines:
        m = re.match(r"([月火水木金土日]+)は(\d+)分", line)
        if m:
            for wd in list(m.group(1)):
                weekday_durations[wd] = int(m.group(2))

    # --- 固定時間 (例: "9:30〜", "13:00〜") ---
    fixed_time = None
    for line in lines:
        m = re.search(r"(\d{1,2}:\d{2})[〜~]?$", line)
        if m and not re.match(r"1回\d+分", line) and "迎え" not in line and "送り" not in line:
            fixed_time = m.group(1)
            # "9:30" → "09:30" に正規化
            parts = fixed_time.split(":")
            fixed_time = f"{int(parts[0]):02d}:{parts[1]}"

    # --- 午後指定 ---
    is_afternoon = False
    for line in lines:
        if line == "午後" or "午後" in line:
            # "午後" 単独の場合、13:00〜 として扱う
            is_afternoon = True
            if fixed_time is None:
                fixed_time = "13:00"

    # --- 結果の組み立て ---
    results = []
    for wd in weekdays:
        duration = weekday_durations.get(wd, base_duration)
        for _ in range(times_per_day):
            results.append({
                "weekday": wd,
                "duration": duration,
                "fixed_time": fixed_time,
                "afternoon": is_afternoon,
            })

    return results


def parse_day_service(ds_str):
    """デイサービス情報をパースし、曜日ごとの不在時間帯を返す。

    入力例:
      "週2回(水金)\n迎え9:40\n送り14:45"
      "週3回(月水金)\n迎え8:50\n送り12:00"
      "週1(月金)\n迎え10:30"           ← "回"省略 + "送り"なし
      "週2(水土)\n迎え9:30\n送り15:00"  ← "回"省略

    返り値: dict
      {"水": ("09:40", "14:45"), "金": ("09:40", "14:45")}
    """
    if pd.isna(ds_str) or str(ds_str).strip() == "":
        return {}

    text = str(ds_str).strip()
    lines = [l.strip() for l in text.replace("\\n", "\n").split("\n") if l.strip()]

    # 曜日抽出 ("回" は省略可)
    weekdays = []
    for line in lines:
        m = re.match(r"(?:デイ)?週\d+回?[（(]([月火水木金土日]+)[)）]", line)
        if m:
            weekdays = list(m.group(1))
            break

    # 迎え・送り時刻
    pickup_time = None
    dropoff_time = None
    for line in lines:
        m = re.search(r"迎え\s*(\d{1,2}:\d{2})", line)
        if m:
            t = m.group(1)
            parts = t.split(":")
            pickup_time = f"{int(parts[0]):02d}:{parts[1]}"
        # 送り時刻: "送り16:00" or "送り16:00-16:40" (範囲の場合は遅い方を採用)
        m = re.search(r"送り\s*(\d{1,2}:\d{2})(?:\s*[-〜~]\s*(\d{1,2}:\d{2}))?", line)
        if m:
            t = m.group(2) if m.group(2) else m.group(1)  # 範囲なら遅い方
            parts = t.split(":")
            dropoff_time = f"{int(parts[0]):02d}:{parts[1]}"

    if not weekdays or not pickup_time:
        return {}

    # 「送り」がない場合は業務終了時間 (18:00) まで不在として扱う
    if not dropoff_time:
        dropoff_time = "18:00"

    return {wd: (pickup_time, dropoff_time) for wd in weekdays}


def load_room_map(room_filepath=None):
    """部屋番号表Excelを読み込み、利用者名→(建物名, 部屋番号, 階数) のマッピングを返す。

    部屋番号表のフォーマット:
      左3列: パシフィック塚本 (部屋番号=3桁数字, 1行目=カナ, 2行目=漢字)
      右3列: ルネッサンス塚本 (部屋番号=階数+英字, 1行目=カナ, 2行目=漢字)

    ファイルがない場合は空dictを返す。
    """
    if room_filepath is None:
        room_files = sorted(glob.glob(os.path.join(ROOM_DIR, "*.xlsx")))
        room_files = [f for f in room_files if not os.path.basename(f).startswith("~$")]
        if not room_files:
            return {}
        room_filepath = room_files[-1]

    try:
        df = pd.read_excel(room_filepath, sheet_name=0, header=None, engine="openpyxl")
    except Exception as e:
        print(f"  [警告] 部屋番号表の読み込み失敗: {e}")
        return {}

    room_map = {}
    # 左ブロック: 建物名は行0・列0
    left_building = str(df.iloc[0, 0]).strip() if pd.notna(df.iloc[0, 0]) else "建物A"
    # 右ブロック: 建物名は行0・列3
    right_building = str(df.iloc[0, 3]).strip() if pd.notna(df.iloc[0, 3]) else "建物B"

    for i in range(1, len(df), 2):
        # 左ブロック
        room_val = df.iloc[i, 0]
        kanji = df.iloc[i + 1, 1] if i + 1 < len(df) and pd.notna(df.iloc[i + 1, 1]) else None
        if pd.notna(room_val) and kanji:
            room_s = str(int(room_val)) if isinstance(room_val, (int, float)) else str(room_val)
            floor = int(room_s[0]) if room_s and room_s[0].isdigit() else 0
            room_map[str(kanji).strip()] = (left_building, room_s, floor)

        # 右ブロック
        room_val = df.iloc[i, 3]
        kanji = df.iloc[i + 1, 4] if i + 1 < len(df) and pd.notna(df.iloc[i + 1, 4]) else None
        if pd.notna(room_val) and kanji:
            room_s = str(room_val).strip()
            floor = int(room_s[0]) if room_s and room_s[0].isdigit() else 0
            room_map[str(kanji).strip()] = (right_building, room_s, floor)

    return room_map


# 出勤マーク (これらのマークがある日は出勤)
EEKANGO_WORK_MARKS = {"B", "D"}
# 休みマーク (これらのマークがある日は休み)
EEKANGO_OFF_MARKS = {"休", "誕", "有"}
# 除外名 (スタッフとして扱わない: 集計行・代行等)
EEKANGO_EXCLUDE_NAMES = {"ヘルパー代行", "常勤合計", "非常勤合計", "勤務体制"}


def load_eekango_shift(filepath, target_dates):
    """ええかんごシフト表Excelを読み込み、標準シフトDataFrameに変換する。

    フォーマット:
      - 月次シート名: R{令和年}.{月} (例: R8.2 = 令和8年2月 = 2026年2月)
      - Row 2: 年月ヘッダー (B=年, E=月)
      - Row 3: 日付ヘッダー (D列〜: datetime)
      - Row 4: 曜日ヘッダー
      - Row 5+: スタッフ行 (B列=名前, D列〜=シフトマーク)
      - マーク: B/D=出勤, 休/誕=休み, 空=判定ロジックで処理

    Args:
        filepath: ええかんごシフト.xlsx のパス
        target_dates: 対象日付リスト (datetime)

    Returns:
        DataFrame (日付, スタッフ名, 開始時間, 終了時間, NG時間帯) or None
    """
    import openpyxl as oxl

    try:
        wb = oxl.load_workbook(filepath, data_only=True)
    except Exception as e:
        print(f"  [警告] ええかんごシフト読み込み失敗: {e}")
        return None

    # 対象月のシートを特定 (target_datesの最初の日付から月を取得)
    target_month = target_dates[0]
    reiwa_year = target_month.year - 2018  # 2019=R1, 2026=R8
    sheet_name = f"R{reiwa_year}.{target_month.month}"

    if sheet_name not in wb.sheetnames:
        print(f"  [警告] シート「{sheet_name}」が見つかりません (シート一覧: {wb.sheetnames})")
        return None

    ws = wb[sheet_name]
    print(f"  ええかんごシフト: シート「{sheet_name}」を使用")

    # Row 3 から日付→列番号のマッピングを作成
    date_col_map = {}  # {datetime.date: col_idx}
    for col_idx in range(4, ws.max_column + 1):
        cell_val = ws.cell(row=3, column=col_idx).value
        if cell_val is not None:
            try:
                if isinstance(cell_val, datetime):
                    d = cell_val.date()
                else:
                    d = pd.to_datetime(cell_val).date()
                date_col_map[d] = col_idx
            except Exception:
                pass

    # 対象日付のうち、シートに存在する日付を抽出
    target_date_set = {}
    for td in target_dates:
        d = td.date() if hasattr(td, "date") else td
        if d in date_col_map:
            target_date_set[d] = date_col_map[d]

    if not target_date_set:
        print(f"  [警告] 対象日付がシート内に見つかりません")
        return None

    # スタッフ行を読み込み (Row 5〜)
    staff_rows = []
    for row_idx in range(5, ws.max_row + 1):
        name_cell = ws.cell(row=row_idx, column=2).value
        if name_cell is None:
            continue
        name = str(name_cell).strip().replace("\u3000", " ")
        if not name or name in EEKANGO_EXCLUDE_NAMES:
            continue
        staff_rows.append((row_idx, name))

    # シフトDataFrameの生成
    # 出勤判定: B/Dマークがある日のみ出勤。空セルや休マークは非出勤。
    rows = []
    active_staff = set()
    for target_date, col_idx in sorted(target_date_set.items()):
        date_str = target_date.strftime("%Y-%m-%d") if hasattr(target_date, "strftime") else str(target_date)

        for row_idx, name in staff_rows:
            mark = ws.cell(row=row_idx, column=col_idx).value
            mark_str = str(mark).strip() if mark is not None else ""

            # 出勤判定: B/D マークのみ出勤
            if mark_str not in EEKANGO_WORK_MARKS:
                continue

            active_staff.add(name)
            rows.append({
                "日付": date_str,
                "スタッフ名": name,
                "開始時間": DEFAULT_SHIFT_START,
                "終了時間": DEFAULT_SHIFT_END,
                "NG時間帯": DEFAULT_SHIFT_NG,
            })

    if not rows:
        print(f"  [警告] 対象期間に出勤スタッフが見つかりません")
        return None

    print(f"  スタッフ: {len(active_staff)}名 ({', '.join(sorted(active_staff))})")
    return pd.DataFrame(rows)


def map_insurance_type(insurance_str):
    """保険種別文字列を判定フラグに変換する。

    入力: "精神", "医療", "介護" 等
    返り値: dict {"判定_障がい": bool, "判定_医療": bool, "判定_介護": bool}
    """
    s = str(insurance_str).strip() if not pd.isna(insurance_str) else ""
    return {
        "判定_障がい": s in ("精神", "障がい", "障害"),
        "判定_医療": s == "医療",
        "判定_介護": s == "介護",
    }


def expand_integrated_to_care_plan(integrated_df, target_dates, room_map=None):
    """統合入力データを、従来のケアプラン形式 (曜日ベース) に展開する。

    Args:
        integrated_df: 統合入力DataFrame (状況, 利用者名, 訪問日, デイサービス, 保険)
        target_dates: 対象日付のリスト (datetime)
        room_map: 部屋番号表マッピング {利用者名: (建物名, 部屋番号, 階数)}

    Returns:
        care_df: ケアプラン相当DataFrame
        master_df: 利用者マスタ相当DataFrame
        user_ds_constraints: 利用者ごとのデイサービス制約 {利用者名: {曜日: (from, to)}}
    """
    if room_map is None:
        room_map = {}
    # 対象曜日を取得
    target_weekdays = set()
    for d in target_dates:
        target_weekdays.add(WEEKDAY_JP[d.weekday()])

    care_rows = []
    master_rows = []
    user_ds_constraints = {}
    seen_users = set()

    # 「訪問中」のみ処理
    active_df = integrated_df[
        integrated_df["状況"].astype(str).str.strip() == "訪問中"
    ] if "状況" in integrated_df.columns else integrated_df

    for _, row in active_df.iterrows():
        user_name = str(row["利用者名"]).strip()
        visit_str = row.get("訪問日", "")
        ds_str = row.get("デイサービス", "")
        insurance_str = row.get("保険", "")

        # マスタ情報 (1利用者1行)
        if user_name not in seen_users:
            seen_users.add(user_name)
            ins_flags = map_insurance_type(insurance_str)
            master_row = {
                "利用者名": user_name,
                "住所": "",
                "判定_障がい": ins_flags["判定_障がい"],
                "判定_医療": ins_flags["判定_医療"],
                "判定_介護": ins_flags["判定_介護"],
            }
            # 部屋番号表から建物・階数・部屋番号を付与
            name_base = user_name.replace("様", "")
            if name_base in room_map:
                bldg, room_no, floor = room_map[name_base]
                master_row["建物名"] = bldg
                master_row["部屋番号"] = room_no
                master_row["階数"] = floor
            else:
                # 元データに建物情報があれば引き継ぐ
                for col in ["建物名", "階数", "部屋番号"]:
                    if col in row.index:
                        master_row[col] = row[col]
            master_rows.append(master_row)

        # デイサービス制約
        ds_constraints = parse_day_service(ds_str)
        if ds_constraints:
            user_ds_constraints[user_name] = ds_constraints

        # 訪問スケジュールをパースしてケアプラン行に展開
        schedule_items = parse_visit_schedule(visit_str)
        for item in schedule_items:
            wd = item["weekday"]
            # 対象曜日のみ展開
            if wd not in target_weekdays:
                continue

            # デイサービスの日の処理
            user_ng_time = ""
            if wd in ds_constraints:
                ds_from, ds_to = ds_constraints[wd]
                if item["fixed_time"]:
                    # 固定時間指定がデイサービス不在時間内ならスキップ
                    ft = parse_time(item["fixed_time"])
                    ds_start = parse_time(ds_from)
                    ds_end = parse_time(ds_to)
                    if ft and ds_start and ds_end and ds_start <= ft < ds_end:
                        continue
                else:
                    # フリータスク → 不在時間帯をNG制約として付与
                    user_ng_time = f"{ds_from}-{ds_to}"

            care_rows.append({
                "利用者名": user_name,
                "曜日": wd,
                "頻度": "週1回",  # 展開済みなので各行は週1回
                "固定時間指定": item["fixed_time"] if item["fixed_time"] else "",
                "所要時間": item["duration"],
                "サービス種類": "訪問看護",
                "利用者NG時間帯": user_ng_time,
            })

    care_df = pd.DataFrame(care_rows) if care_rows else pd.DataFrame(
        columns=["利用者名", "曜日", "頻度", "固定時間指定", "所要時間", "サービス種類"]
    )
    master_df = pd.DataFrame(master_rows) if master_rows else pd.DataFrame(
        columns=["利用者名", "住所", "判定_障がい", "判定_医療", "判定_介護"]
    )

    return care_df, master_df, user_ds_constraints


def generate_default_shift(staff_names, target_dates):
    """デフォルトのシフト情報を生成する。

    全スタッフに対し、対象日付すべてで 09:00-18:00 / NG 12:00-13:00 を設定。

    Args:
        staff_names: スタッフ名リスト
        target_dates: 対象日付リスト (datetime)

    Returns:
        DataFrame (日付, スタッフ名, 開始時間, 終了時間, NG時間帯)
    """
    rows = []
    for d in target_dates:
        date_str = d.strftime("%Y-%m-%d")
        for name in staff_names:
            rows.append({
                "日付": date_str,
                "スタッフ名": name,
                "開始時間": DEFAULT_SHIFT_START,
                "終了時間": DEFAULT_SHIFT_END,
                "NG時間帯": DEFAULT_SHIFT_NG,
            })
    return pd.DataFrame(rows)


def detect_input_mode():
    """入力モードを判定する。

    統合入力フォルダ (00_integrated) にExcelがあれば "integrated" モード、
    従来の3フォルダにファイルがあれば "classic" モード。

    Returns:
        ("integrated", filepath) or ("classic", None)
    """
    integrated_files = sorted(glob.glob(os.path.join(INTEGRATED_DIR, "*.xlsx")))
    integrated_files = [f for f in integrated_files if not os.path.basename(f).startswith("~$")]

    if integrated_files:
        # 最新ファイルを採用
        if len(integrated_files) > 1:
            integrated_files.sort(key=lambda f: os.path.getmtime(f), reverse=True)
            chosen = integrated_files[0]
            print(f"  [情報] 統合入力フォルダに複数ファイルあり ({len(integrated_files)}件) → 最新を採用")
            print(f"    採用: {os.path.basename(chosen)}")
        else:
            chosen = integrated_files[0]
        return "integrated", chosen

    return "classic", None


# ====================================================================
# Greedy 割り当てロジック
# ====================================================================

def is_c_building(building_name):
    """建物名がC棟（隔離棟）か判定する。

    「C棟」「C棟（離れ）」「C棟(別館)」等の表記揺れに対応するため、
    部分一致で判定する。
    """
    if not building_name:
        return False
    return "C棟" in str(building_name)


def is_isolation_ok(staff_name, task_building, visited_c, visited_non_c):
    """C棟隔離ルールに基づき、スタッフがこの建物に行けるか判定する。

    ルール:
      - C棟に一度でも行ったスタッフ → その日はC棟以外に行けない
      - C棟以外に一度でも行ったスタッフ → その日はC棟に行けない
      - まだどこにも行っていない (朝イチ) → どこでもOK
    """
    task_is_c = is_c_building(task_building)

    if staff_name in visited_c and not task_is_c:
        # C棟に行った人が非C棟タスクに行こうとしている → NG
        return False
    if staff_name in visited_non_c and task_is_c:
        # 非C棟に行った人がC棟タスクに行こうとしている → NG
        return False
    return True


def calculate_movement_score(task_floor, task_building, last_floor, last_building):
    """建物・階を考慮した動線スコアを計算する。

    スコアが低いほど動線が良い (ボーナスはマイナス、ペナルティはプラス)。
    - 初回訪問 (前回情報なし): 0pt
    - 同一建物: -30pt
      - さらに同一階: -20pt (合計 -50pt)
      - 異なる階: +階差×5pt
    - 異なる建物: +30pt
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
        # 異なる建物ペナルティ (基礎分。累積移動ペナルティは呼び出し側で加算)
        return 30


def _is_slot_range_available(needed_slots, staff_name, staff_info):
    """指定スロット範囲がスタッフにとって利用可能か判定する。

    勤務時間内 / NG時間帯でない / 空きスロット の全条件を満たす場合 True。
    """
    info = staff_info[staff_name]
    for idx in needed_slots:
        slot_label = TIME_SLOTS[idx]
        if not is_within_shift(slot_label, info["shift_start"], info["shift_end"]):
            return False
        if is_in_ng_range(slot_label, info["ng_range"]):
            return False
    if info["occupied_slots"] & set(needed_slots):
        return False
    return True


def _do_assign(best, needed_slots, task_row, staff_info, staff_last_floor,
               staff_last_building, visited_c, visited_non_c,
               carry_over, day_stats, assignments):
    """タスクをスタッフに割り当て、各種状態を更新する共通処理。"""
    task_floor = safe_int(task_row.get("階数", None), default=0)
    task_building = safe_str(task_row.get("建物名", None), default="")
    task_rank = safe_int(task_row.get("ランク", None), default=1)

    # 建物移動カウント: 前回の建物から変わった場合にカウントアップ
    last_bldg = staff_last_building[best]
    if last_bldg is not None and task_building and last_bldg != task_building:
        day_stats[best]["move_count"] = day_stats[best].get("move_count", 0) + 1

    staff_info[best]["occupied_slots"].update(needed_slots)
    staff_last_floor[best] = task_floor
    staff_last_building[best] = task_building

    if task_building:
        if is_c_building(task_building):
            visited_c.add(best)
        else:
            visited_non_c.add(best)

    carry_over[best]["cumulative_rank"] += task_rank
    day_stats[best]["count"] += 1
    day_stats[best]["total_rank"] += task_rank
    assignments[best].append((needed_slots, task_row))


def assign_tasks_for_date(day_staff_df, day_tasks_fixed, day_tasks_free,
                          date_display, carry_over):
    """1日分のタスクを2段階方式で割り当てる。

    フェーズ1 (固定タスク): 固定時間指定ありのタスクを時間順に処理。
    フェーズ2 (フリータスク): 固定時間指定なしのタスクを、
        全スタッフ×全空きスロットから最適な場所を探索して自動配置。

    累積ランクは日をまたいで引き継ぎ (carry_over)、
    C棟隔離履歴は日ごとにリセットされる。
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

    # C棟隔離の履歴管理 (日ごと)
    visited_c = set()
    visited_non_c = set()

    assignments = {name: [] for name in staff_info}
    day_stats = {name: {"count": 0, "total_rank": 0} for name in staff_info}

    # ================================================================
    # フェーズ1: 固定タスクの割り当て (従来ロジック)
    # ================================================================
    for _, task_row in day_tasks_fixed.iterrows():
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

        candidates = []
        for staff_name in staff_info:
            if not _is_slot_range_available(needed_slots, staff_name, staff_info):
                continue
            if not is_isolation_ok(staff_name, task_building, visited_c, visited_non_c):
                continue
            candidates.append(staff_name)

        if not candidates:
            slot_label = TIME_SLOTS[start_idx]
            bldg_info = f" ({task_building})" if task_building else ""
            msg = f"  [警告] {date_display} {slot_label}の{user_name}様{bldg_info}を担当できるスタッフがいません"
            warnings.append(msg)
            continue

        def calc_score(name):
            movement = calculate_movement_score(
                task_floor, task_building,
                staff_last_floor[name], staff_last_building[name],
            )
            move_penalty = day_stats[name].get("move_count", 0) * 40
            jitter = random.uniform(0, 5)
            return (carry_over[name]["cumulative_rank"] * 10) + movement + move_penalty + jitter

        best = min(candidates, key=calc_score)
        _do_assign(best, needed_slots, task_row, staff_info,
                   staff_last_floor, staff_last_building,
                   visited_c, visited_non_c,
                   carry_over, day_stats, assignments)

    # ================================================================
    # フェーズ2: フリータスクの自動配置
    # 全スタッフ × 全空きスロットから最適な (スタッフ, 開始時間) を探索
    # ================================================================
    for _, task_row in day_tasks_free.iterrows():
        try:
            duration_min = int(task_row["所要時間"])
        except (ValueError, TypeError):
            duration_min = 30
        slot_count = max(1, duration_min // 30)

        user_name = safe_str(task_row["利用者名"])
        task_floor = safe_int(task_row.get("階数", None), default=0)
        task_building = safe_str(task_row.get("建物名", None), default="")
        task_rank = safe_int(task_row.get("ランク", None), default=1)
        user_ng = safe_str(task_row.get("利用者NG時間帯", None), default="")

        # 全候補 (スタッフ, 開始スロット) を列挙してスコアリング
        best_candidate = None
        best_score = float("inf")

        for staff_name in staff_info:
            if not is_isolation_ok(staff_name, task_building, visited_c, visited_non_c):
                continue

            # このスタッフで配置可能な全開始スロットを探索
            max_start = len(TIME_SLOTS) - slot_count + 1
            for start_idx in range(max_start):
                needed_slots = list(range(start_idx, start_idx + slot_count))
                if not _is_slot_range_available(needed_slots, staff_name, staff_info):
                    continue
                # 利用者NG時間帯チェック (デイサービス不在等)
                if user_ng:
                    ng_conflict = False
                    for idx in needed_slots:
                        if is_in_ng_range(TIME_SLOTS[idx], user_ng):
                            ng_conflict = True
                            break
                    if ng_conflict:
                        continue

                # スコア計算: 負荷 + 動線 + 累積移動ペナルティ + ジッター
                movement = calculate_movement_score(
                    task_floor, task_building,
                    staff_last_floor[staff_name],
                    staff_last_building[staff_name],
                )
                move_penalty = day_stats[staff_name].get("move_count", 0) * 40
                jitter = random.uniform(0, 5)
                score = (carry_over[staff_name]["cumulative_rank"] * 10) + movement + move_penalty + jitter

                if score < best_score:
                    best_score = score
                    best_candidate = (staff_name, needed_slots)

        if best_candidate is None:
            bldg_info = f" ({task_building})" if task_building else ""
            msg = f"  [警告] {date_display} フリータスク {user_name}様{bldg_info}を配置できる空きがありません"
            warnings.append(msg)
            continue

        best_staff, best_slots = best_candidate
        _do_assign(best_staff, best_slots, task_row, staff_info,
                   staff_last_floor, staff_last_building,
                   visited_c, visited_non_c,
                   carry_over, day_stats, assignments)

    # デバッグ: C棟隔離状態をログ出力
    if visited_c:
        c_staff = ", ".join(sorted(visited_c))
        print(f"    [隔離ログ] {date_display}: C棟担当 → {c_staff}")
        violation = visited_c & visited_non_c
        if violation:
            v_staff = ", ".join(sorted(violation))
            print(f"    [隔離違反!] {date_display}: {v_staff} がC棟と他棟の両方に割り当てられています")

    return assignments, warnings, day_stats, carry_over


# ====================================================================
# Excel生成 (マトリクスレイアウト)
# ====================================================================

def build_matrix_excel(staff_shift_df, merged_df, timestamp_str, master_df=None):
    """割り当て結果をガントチャート風マトリクスExcelとして出力する。

    実例 (週間スケジュール実例.xlsx) 準拠レイアウト:
      行1: ヘッダー (B1=訪問予定, D1=開始日, E1=～, F1=終了日)
      行18: 時間ヘッダー (E-V: 09:00〜17:30)
      行19〜: 曜日ブロック
        B列: 曜日 (ブロック内セル結合)
        C列: 日付 (ブロック内セル結合)
        D列: スタッフ名
        E-V列: タスクセル (同一タスクはセル結合)
      右端: 利用者マスタ一覧 (X列〜)
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "週間スケジュール"

    # 定数: 列オフセット
    COL_WEEKDAY = 2    # B列: 曜日
    COL_DATE = 3       # C列: 日付
    COL_STAFF = 4      # D列: スタッフ名
    COL_TIME_START = 5 # E列: 最初の時間スロット (09:00)
    ROWS_PER_BLOCK = 7 # 各曜日ブロックの行数 (スタッフ行 + 余白)
    HEADER_ROW = 17    # 時間ヘッダー行 (実例では18だが0-indexed調整で17)
    DATA_START_ROW = 18  # データ開始行

    # --- 行1: タイトルヘッダー ---
    dates = sorted(staff_shift_df["日付"].unique())
    try:
        first_date = pd.to_datetime(dates[0])
        last_date = pd.to_datetime(dates[-1])
    except Exception:
        first_date = datetime.now()
        last_date = first_date + timedelta(days=6)

    ws.cell(row=1, column=COL_WEEKDAY, value="訪問予定").font = Font(bold=True, size=12)
    ws.cell(row=1, column=COL_STAFF, value=first_date.strftime("%Y/%m/%d")).font = FONT_HEADER
    ws.cell(row=1, column=COL_TIME_START, value="～").font = FONT_HEADER
    ws.cell(row=1, column=COL_TIME_START + 1, value=last_date.strftime("%Y/%m/%d")).font = FONT_HEADER

    # --- 時間ヘッダー行 ---
    for slot_idx, slot_label in enumerate(TIME_SLOTS):
        col = COL_TIME_START + slot_idx
        cell = ws.cell(row=HEADER_ROW, column=col, value=slot_label)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.border = THIN_BORDER
        cell.alignment = ALIGN_CENTER

    # B-D列のヘッダーも設定
    for col_idx, label in [(COL_WEEKDAY, "曜日"), (COL_DATE, "日付"), (COL_STAFF, "スタッフ")]:
        cell = ws.cell(row=HEADER_ROW, column=col_idx, value=label)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.border = THIN_BORDER
        cell.alignment = ALIGN_CENTER

    # --- 割り当て計算 ---
    all_warnings = []
    global_stats = {}
    carry_over = {}
    all_assignments = {}  # {date_val: {staff_name: [(slots, task_row), ...]}}

    # ケアプランを曜日別に整理 (固定/フリー分離)
    fixed_by_weekday = {}
    free_by_weekday = {}
    for _, task_row in merged_df.iterrows():
        weekday = str(task_row["曜日"]).strip()
        fixed_time = task_row.get("固定時間指定", None)
        is_free = pd.isna(fixed_time) or str(fixed_time).strip() == ""
        if is_free:
            free_by_weekday.setdefault(weekday, []).append(task_row)
        else:
            fixed_by_weekday.setdefault(weekday, []).append(task_row)

    for date_val in dates:
        try:
            date_obj = pd.to_datetime(date_val)
            weekday_jp = WEEKDAY_JP[date_obj.weekday()]
            date_display = date_obj.strftime("%Y-%m-%d")
        except Exception:
            weekday_jp = None
            date_display = str(date_val)

        day_staff_df = staff_shift_df[staff_shift_df["日付"] == date_val]
        fixed_list = fixed_by_weekday.get(weekday_jp, [])
        free_list = free_by_weekday.get(weekday_jp, [])
        day_fixed_df = pd.DataFrame(fixed_list) if fixed_list else pd.DataFrame()
        day_free_df = pd.DataFrame(free_list) if free_list else pd.DataFrame()

        has_tasks = not day_fixed_df.empty or not day_free_df.empty
        if has_tasks:
            assignments, warnings, day_stats, carry_over = assign_tasks_for_date(
                day_staff_df, day_fixed_df, day_free_df, date_display, carry_over
            )
            all_warnings.extend(warnings)
            for name, stats in day_stats.items():
                if name not in global_stats:
                    global_stats[name] = {"count": 0, "total_rank": 0}
                global_stats[name]["count"] += stats["count"]
                global_stats[name]["total_rank"] += stats["total_rank"]
        else:
            assignments = {row["スタッフ名"]: [] for _, row in day_staff_df.iterrows()}

        all_assignments[date_val] = assignments

    # --- 曜日ブロック書き込み ---
    current_row = DATA_START_ROW
    for date_val in dates:
        try:
            date_obj = pd.to_datetime(date_val)
            weekday_jp = WEEKDAY_JP[date_obj.weekday()]
            date_display = date_obj.strftime("%m/%d")
        except Exception:
            weekday_jp = "?"
            date_display = str(date_val)

        day_staff_df = staff_shift_df[staff_shift_df["日付"] == date_val]
        staff_names = day_staff_df["スタッフ名"].tolist()
        assignments = all_assignments.get(date_val, {})

        block_start_row = current_row
        # スタッフ行数 = 実スタッフ数 (最低1行)
        num_staff_rows = max(len(staff_names), 1)
        block_end_row = block_start_row + num_staff_rows - 1

        # B列: 曜日 (セル結合)
        wd_cell = ws.cell(row=block_start_row, column=COL_WEEKDAY, value=weekday_jp)
        wd_cell.font = Font(bold=True, size=11)
        wd_cell.border = THIN_BORDER
        wd_cell.alignment = ALIGN_CENTER
        if num_staff_rows > 1:
            ws.merge_cells(
                start_row=block_start_row, start_column=COL_WEEKDAY,
                end_row=block_end_row, end_column=COL_WEEKDAY,
            )
            # 結合後も罫線を設定
            for r in range(block_start_row, block_end_row + 1):
                ws.cell(row=r, column=COL_WEEKDAY).border = THIN_BORDER

        # C列: 日付 (セル結合)
        dt_cell = ws.cell(row=block_start_row, column=COL_DATE, value=date_display)
        dt_cell.font = FONT_CELL
        dt_cell.border = THIN_BORDER
        dt_cell.alignment = ALIGN_CENTER
        if num_staff_rows > 1:
            ws.merge_cells(
                start_row=block_start_row, start_column=COL_DATE,
                end_row=block_end_row, end_column=COL_DATE,
            )
            for r in range(block_start_row, block_end_row + 1):
                ws.cell(row=r, column=COL_DATE).border = THIN_BORDER

        # 各スタッフ行を書き込む
        for staff_idx, (_, staff_row) in enumerate(day_staff_df.iterrows()):
            row = block_start_row + staff_idx
            staff_name = staff_row["スタッフ名"]

            # D列: スタッフ名
            name_cell = ws.cell(row=row, column=COL_STAFF, value=staff_name)
            name_cell.border = THIN_BORDER
            name_cell.font = FONT_CELL
            name_cell.alignment = ALIGN_CENTER

            # E-V列: 全スロットに罫線設定 + NG時間帯
            ng_range = staff_row["NG時間帯"]
            for slot_idx in range(len(TIME_SLOTS)):
                col = COL_TIME_START + slot_idx
                cell = ws.cell(row=row, column=col)
                cell.border = THIN_BORDER
                cell.font = FONT_CELL
                cell.alignment = ALIGN_CENTER
                if is_in_ng_range(TIME_SLOTS[slot_idx], ng_range):
                    cell.value = "休憩"
                    cell.fill = FILL_HEADER

            # タスクを書き込む (セル結合あり)
            for slot_indices, task_row in assignments.get(staff_name, []):
                user_name = safe_str(task_row["利用者名"])
                service = safe_str(task_row.get("サービス種類", None), default="")
                fill = get_fill_for_user(task_row)

                start_time = TIME_SLOTS[slot_indices[0]] if slot_indices else ""
                display_text = f"{start_time}\n{user_name}"
                if service and service != "訪問看護":
                    display_text = f"{start_time}\n{user_name}（{service}）"

                # 最初のセルに値を書き込み
                first_col = COL_TIME_START + slot_indices[0]
                cell = ws.cell(row=row, column=first_col)
                cell.value = display_text
                cell.fill = fill
                cell.font = FONT_CELL
                cell.alignment = ALIGN_CENTER

                # 複数スロットにまたがる場合はセル結合
                if len(slot_indices) > 1:
                    last_col = COL_TIME_START + slot_indices[-1]
                    ws.merge_cells(
                        start_row=row, start_column=first_col,
                        end_row=row, end_column=last_col,
                    )
                    # 結合セル全体に背景色・罫線を設定
                    for idx in slot_indices:
                        c = COL_TIME_START + idx
                        ws.cell(row=row, column=c).fill = fill
                        ws.cell(row=row, column=c).border = THIN_BORDER

        current_row = block_end_row + 1

    # --- 右側: 利用者マスタ一覧 ---
    MASTER_COL_START = COL_TIME_START + len(TIME_SLOTS) + 1  # 1列空けて配置
    master_headers = ["状況", "利用者名", "訪問日", "DS", "保険"]
    if master_df is not None and not master_df.empty:
        # ヘッダー
        for i, h in enumerate(master_headers):
            col = MASTER_COL_START + i
            cell = ws.cell(row=HEADER_ROW, column=col, value=h)
            cell.fill = FILL_HEADER
            cell.font = FONT_HEADER
            cell.border = THIN_BORDER
            cell.alignment = ALIGN_CENTER

        # データ行
        master_row = DATA_START_ROW
        for _, mrow in master_df.iterrows():
            user_name = safe_str(mrow.get("利用者名", ""), "")
            status = safe_str(mrow.get("状況", ""), "")
            visit = safe_str(mrow.get("訪問日", ""), "")
            ds = safe_str(mrow.get("デイサービス", ""), "")
            insurance = safe_str(mrow.get("保険", ""), "")

            values = [status, user_name, visit, ds, insurance]
            for i, val in enumerate(values):
                col = MASTER_COL_START + i
                cell = ws.cell(row=master_row, column=col, value=val)
                cell.font = FONT_CELL
                cell.border = THIN_BORDER
                cell.alignment = Alignment(
                    horizontal="left", vertical="center", wrap_text=True
                )
            master_row += 1

    # --- 列幅・行高さ調整 ---
    ws.column_dimensions[get_column_letter(COL_WEEKDAY)].width = 5   # B: 曜日
    ws.column_dimensions[get_column_letter(COL_DATE)].width = 8      # C: 日付
    ws.column_dimensions[get_column_letter(COL_STAFF)].width = 10    # D: スタッフ
    for slot_idx in range(len(TIME_SLOTS)):
        col_letter = get_column_letter(COL_TIME_START + slot_idx)
        ws.column_dimensions[col_letter].width = 14
    # マスタ列の幅
    if master_df is not None:
        master_widths = [6, 14, 20, 20, 6]
        for i, w in enumerate(master_widths):
            col_letter = get_column_letter(MASTER_COL_START + i)
            ws.column_dimensions[col_letter].width = w
    # 行高さ
    for row_idx in range(DATA_START_ROW, current_row):
        ws.row_dimensions[row_idx].height = 40

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

def _prepare_merged_df(care_plan, medical_master):
    """ケアプランとマスタを結合し、欠損列を補填してソート済みDataFrameを返す。"""
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

    # 固定タスク / フリータスクの件数を集計
    def _is_free(v):
        return pd.isna(v) or str(v).strip() == ""
    n_fixed = sum(1 for v in merged["固定時間指定"] if not _is_free(v))
    n_free = len(merged) - n_fixed

    # ソート: 固定時間指定 (昇順, 空欄は末尾) > 建物名 (昇順) > 階数 (昇順)
    merged["_sort_time"] = merged["固定時間指定"].apply(
        lambda v: safe_str(v, "99:99")
    )
    merged["_sort_building"] = merged["建物名"].apply(lambda v: safe_str(v, ""))
    merged["_sort_floor"] = merged["階数"].apply(lambda v: safe_int(v, 0))
    merged = merged.sort_values(
        ["_sort_time", "_sort_building", "_sort_floor"]
    ).drop(columns=["_sort_time", "_sort_building", "_sort_floor"])

    return merged, n_fixed, n_free


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

    # 2. 入力モード判定
    print("\n[2/7] 入力ファイル探索...")
    input_mode, integrated_file = detect_input_mode()

    if input_mode == "integrated":
        # ============================================================
        # 統合入力モード
        # ============================================================
        print(f"  入力モード: 統合 (1ファイル)")
        print(f"  統合ファイル: {os.path.basename(integrated_file)}")
        used_files = [integrated_file]

        # 3. データ読み込み (ヘッダー行を自動検出)
        print("\n[3/7] 統合データ読み込み...")
        try:
            integrated_df = pd.read_excel(integrated_file, sheet_name=0, engine="openpyxl")
            # 「利用者名」列がなければヘッダー行がずれている → 先頭行をスキャン
            if "利用者名" not in integrated_df.columns:
                df_raw = pd.read_excel(integrated_file, sheet_name=0, header=None, engine="openpyxl")
                header_row = None
                for i in range(min(10, len(df_raw))):
                    row_vals = [str(v).strip() for v in df_raw.iloc[i] if pd.notna(v)]
                    if "利用者名" in row_vals:
                        header_row = i
                        break
                if header_row is not None:
                    integrated_df = pd.read_excel(
                        integrated_file, sheet_name=0, header=header_row, engine="openpyxl"
                    )
                    # Unnamed列を除去
                    integrated_df = integrated_df.loc[
                        :, ~integrated_df.columns.astype(str).str.startswith("Unnamed")
                    ]
                    print(f"  [情報] ヘッダー行を自動検出: 行{header_row + 1}")
        except Exception as e:
            print(f"[エラー] 統合ファイルの読み込みに失敗しました: {e}")
            sys.exit(1)
        # Unnamed列を除去 (Excelの空列対策)
        integrated_df = integrated_df.loc[
            :, ~integrated_df.columns.astype(str).str.startswith("Unnamed")
        ]
        # NaN行を除去
        if "利用者名" in integrated_df.columns:
            integrated_df = integrated_df.dropna(subset=["利用者名"])
        print(f"  統合入力: {len(integrated_df)} 件")

        # 「訪問中」のみ抽出
        if "状況" in integrated_df.columns:
            active = integrated_df[integrated_df["状況"].astype(str).str.strip() == "訪問中"]
            skipped = len(integrated_df) - len(active)
            if skipped > 0:
                print(f"  [情報] 「訪問中」以外の {skipped} 件をスキップ")
        else:
            active = integrated_df

        # 対象日付の決定
        # 今週の月曜日を起点とした7日間をデフォルトとする
        today = datetime.now()
        monday = today - timedelta(days=today.weekday())
        target_dates = [monday + timedelta(days=i) for i in range(7)]

        # シフトファイル探索 (優先順位: 01_shift標準形式 > ええかんごシフト > デフォルト)
        shift_files = sorted(glob.glob(os.path.join(SHIFT_DIR, "*.xlsx")))
        shift_files = [f for f in shift_files if not os.path.basename(f).startswith("~$")]

        # ええかんごシフト探索 (01_shift/ と 00_analysis/ の両方をチェック)
        eekango_shift_file = None
        for search_dir in [SHIFT_DIR, os.path.join(BASE_DIR, "00_analysis")]:
            candidates = sorted(glob.glob(os.path.join(search_dir, "*シフト*.xlsx")))
            candidates = [f for f in candidates if not os.path.basename(f).startswith("~$")]
            # 標準形式のシフトファイルと区別するため、ええかんご形式かどうか確認
            for c in candidates:
                if "ええかんご" in os.path.basename(c) or "シフト" in os.path.basename(c):
                    eekango_shift_file = c
                    break
            if eekango_shift_file:
                break

        # 標準形式のシフトファイルがある場合はそちらを優先
        standard_shift_files = [f for f in shift_files
                                if f != eekango_shift_file]
        if standard_shift_files:
            shift_file = standard_shift_files[-1]
            print(f"  シフトファイル: {os.path.basename(shift_file)}")
            staff_shift = load_excel(shift_file, "01_shift")
            used_files.append(shift_file)
            target_dates = [pd.to_datetime(d) for d in staff_shift["日付"].unique()]
        elif eekango_shift_file:
            # ええかんごシフト形式で読み込み
            print(f"  シフトファイル: {os.path.basename(eekango_shift_file)} (ええかんご形式)")
            staff_shift = load_eekango_shift(eekango_shift_file, target_dates)
            if staff_shift is not None:
                # 00_analysis/ 内のファイルはアーカイブ対象外
                if not eekango_shift_file.startswith(os.path.join(BASE_DIR, "00_analysis")):
                    used_files.append(eekango_shift_file)
            else:
                # 読み込み失敗 → デフォルトにフォールバック
                print("  [情報] ええかんごシフト読み込み失敗 → デフォルトシフトを使用")
                default_staff = ["スタッフA", "スタッフB", "スタッフC", "スタッフD", "スタッフE"]
                staff_shift = generate_default_shift(default_staff, target_dates)
        else:
            # シフトファイルなし → デフォルトシフトを生成
            print("  [情報] シフトファイルなし → デフォルトシフトを自動生成")
            default_staff = ["スタッフA", "スタッフB", "スタッフC", "スタッフD", "スタッフE"]
            staff_shift = generate_default_shift(default_staff, target_dates)
            print(f"  デフォルトスタッフ: {', '.join(default_staff)}")

        print(f"  対象期間: {target_dates[0].strftime('%Y-%m-%d')} ～ {target_dates[-1].strftime('%Y-%m-%d')}")

        # 部屋番号表の読み込み
        room_map = load_room_map()
        if room_map:
            print(f"  部屋番号表: {len(room_map)} 名分の部屋情報を取得")
        else:
            print("  [情報] 部屋番号表なし → 建物・階数情報なしで動作")

        # 4. 統合データ → ケアプラン + マスタに展開
        print("\n[4/7] データ変換 (統合 → ケアプラン + マスタ)...")
        care_plan, medical_master, user_ds_constraints = expand_integrated_to_care_plan(
            integrated_df, target_dates, room_map=room_map
        )
        print(f"  ケアプラン: {len(care_plan)} 件に展開")
        print(f"  利用者マスタ: {len(medical_master)} 名")
        if user_ds_constraints:
            print(f"  デイサービス制約: {len(user_ds_constraints)} 名")
            for uname, constraints in user_ds_constraints.items():
                days_str = ",".join(sorted(constraints.keys()))
                print(f"    {uname}: {days_str}")

        merged, n_fixed, n_free = _prepare_merged_df(care_plan, medical_master)
        print(f"  結合結果: {len(merged)} 件 (固定: {n_fixed}件, フリー: {n_free}件)")

        # 右側マスタ表示用に元データを保持
        right_master_df = integrated_df

    else:
        # ============================================================
        # 従来モード (3フォルダ)
        # ============================================================
        print(f"  入力モード: クラシック (3フォルダ)")
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
        merged, n_fixed, n_free = _prepare_merged_df(care_plan, medical_master)
        print(f"  ケアプラン + 利用者マスタ → {len(merged)} 件 (固定: {n_fixed}件, フリー: {n_free}件)")
        right_master_df = None

    # 5. 割り当て & Excel生成
    print("\n[5/7] スコアリング割り当て & マトリクスExcel生成...")
    output_path, warnings, global_stats = build_matrix_excel(
        staff_shift, merged, timestamp_str, master_df=right_master_df
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
