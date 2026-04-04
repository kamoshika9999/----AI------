"""
planning_core — 工場向け「配台計画」シミュレーションの中核（単一モジュール）

【このファイルの役割】
    VBA / マクロ実行ブックから環境変数で渡されたパスを読み、加工計画データ・勤怠・
    スキル need を統合して、設備・担当者への割付結果（Excel）を生成する。

【外部から直接使う入口（他 .py から import される想定）】
    - ``TASKS_INPUT_WORKBOOK`` … マクロブックのパス（環境変数 TASK_INPUT_WORKBOOK）
    - ``run_stage1_extract()`` … 段階1。`task_extract_stage1.py` から呼ぶ
    - ``generate_plan()`` … 段階2。`plan_simulation_stage2.py` から呼ぶ

【処理の流れ（ざっくり）】
    1. 段階1: 「加工計画DATA」→ 中間 `output/plan_input_tasks.xlsx`、
       マクロブック内「設定_配台不要工程」の行同期・D列→E列(Gemini)・保存
    2. 段階2: master の skills / 勤怠 / 配台計画シートを読み、日付ループで割付、
       `output/年/月/日/` に結果ブック・個人スケジュールを出力

【ソース上の構成（=#= 見出しでスクロール検索可能）】
    - 先頭 … ログ・パス・レガシーファイル掃除
    - 「【設定】APIキー / 基本ルール / ファイル名」… 列名定数・パス
    - 配台計画シート列・参照列ヘルパ
    - 結果ガント・タスク一覧の整形
    - 実績 DATA・特別指定備考・Gemini 連携
    - 配台不要（分割自動・設定シート・COM 保存フォールバック）
    - ``run_stage1_extract`` … 段階1本体
    - 「1. コア計算ロジック」… 時刻・休憩を踏んだ実働分計算
    - 「2. マスタデータ・出勤簿 と AI解析」… skills/need/勤怠
    - ``generate_plan`` … 段階2本体（メインループ）

【命名】
    - 先頭 ``_`` … モジュール内専用ヘルパ（外部から呼ばない想定）
    - ``PLAN_*`` / ``TASK_*`` … Excel 見出しと一致させる定数

【依存】
    pandas, openpyxl, google.genai（GEMINI_API_KEY）, Windows では pywin32（COM はテスト用のみ。本番のシート保存は openpyxl または VBA）
"""

import pandas as pd
from datetime import datetime, timedelta, time, date
from collections import defaultdict
import itertools
import csv
import json
import re
import traceback
import base64
import hashlib
import unicodedata
import time as time_module
from google import genai
import logging
import calendar
import math
import os
import sys
import ctypes
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.styles.borders import Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# =========================================================
# 【重要】実行ファイルと同じフォルダをカレントディレクトリに設定
# （UNCパスやVBAからの呼び出し時のパスエラーを回避します）
# =========================================================
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# cmd で chcp 65001 時にログの日本語をコンソールへ出しやすくする（段階2でリダイレクト無し実行時向け）
if os.name == "nt" and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# ---------------------------------------------------------
# コマンドプロンプト画面を「最前面（トップレベル）」にする処理
# ---------------------------------------------------------
if os.name == 'nt':
    hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    if hwnd:
        # HWND_TOPMOST = -1, SWP_NOMOVE = 0x0002, SWP_NOSIZE = 0x0001
        ctypes.windll.user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 3)

# ---------------------------------------------------------
# ログを「画面（コンソール）」と「一時ファイル」の両方に出力する
# ---------------------------------------------------------
logger = logging.getLogger()
logger.setLevel(logging.INFO)
# 既存のハンドラをクリア
if logger.hasHandlers():
    logger.handlers.clear()

formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

# 1. 画面(コンソール)用ハンドラ
stream_handler = logging.StreamHandler(sys.stdout)
stream_handler.setFormatter(formatter)
logger.addHandler(stream_handler)

# 2. 成果物用 output / 実行ログ用 log（いずれも常に最新1ファイルのみ上書き）
output_dir = os.path.join(os.getcwd(), 'output')
os.makedirs(output_dir, exist_ok=True)
log_dir = os.path.join(os.getcwd(), 'log')
os.makedirs(log_dir, exist_ok=True)
# 旧仕様: ログが output 直下にあった名残を削除（常に log 側の1ファイルのみが最新）
for _legacy_name in (
    "execution_log.txt",
    "ai_task_special_remark_last.txt",
    "ai_task_special_last_prompt.txt",
    "planning_conflict_highlight.tsv",
    "cmd_stage2.log",
):
    try:
        os.remove(os.path.join(output_dir, _legacy_name))
    except OSError:
        pass


def get_dated_output_dir(base_dt=None):
    """
    output 配下を 年/月/日 の3階層に整理して返す。
    例: output/2026/03/27
    """
    dt = base_dt if isinstance(base_dt, datetime) else datetime.now()
    y = f"{dt.year:04d}"
    m = f"{dt.month:02d}"
    d = f"{dt.day:02d}"
    path = os.path.join(output_dir, y, m, d)
    os.makedirs(path, exist_ok=True)
    return path

# 3. ファイル用ハンドラ（VBAで後から読み取るため UTF-8 で保存）
log_file_path = os.path.join(log_dir, 'execution_log.txt')
# BOM 付き UTF-8（Excel / VBA の ADODB.Stream が文字化けしにくい）
file_handler = logging.FileHandler(log_file_path, encoding='utf-8-sig')
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)
ai_cache_path = os.path.join(output_dir, 'ai_remarks_cache.json')
# 「設定_配台不要工程」シート作成・保存の成否デバッグ（execution_log と併用）
exclude_rules_sheet_debug_log_path = os.path.join(log_dir, "exclude_rules_sheet_debug.txt")
# 保存失敗時に E 列（ロジック式）だけを退避し、次回 run_exclude_rules_sheet_maintenance で自動適用する
EXCLUDE_RULES_E_SIDECAR_FILENAME = "exclude_rules_e_column_pending.json"
# openpyxl 保存失敗時に VBA が E 列へ書き込むための UTF-8 TSV（Base64）。
EXCLUDE_RULES_E_VBA_TSV_FILENAME = "exclude_rules_e_column_vba.tsv"
# openpyxl 保存失敗時に VBA が A〜E を一括反映する UTF-8 TSV（行ごとに 5 セル分 Base64）。
EXCLUDE_RULES_MATRIX_VBA_FILENAME = "exclude_rules_matrix_vba.tsv"
# VBA がメイン P 列へ書き込むための UTF-8 テキスト（Excel 開いたまま save できない問題の回避）
GEMINI_USAGE_SUMMARY_FOR_MAIN_FILE = "gemini_usage_summary_for_main.txt"
# 全実行を通した Gemini 利用・推定料金の累計（API 応答ごとに更新。コンソールには単発料金を出さない）
GEMINI_USAGE_CUMULATIVE_JSON_FILE = "gemini_usage_cumulative.json"
# 期間別バケットをフラット化した CSV（Excel の折れ線・棒グラフ用）
GEMINI_USAGE_BUCKETS_CSV_FILE = "gemini_usage_buckets_for_chart.csv"
# テスト: EXCLUDE_RULES_TEST_E1234=1 で EXCLUDE_RULES_SHEET_NAME（「設定_配台不要工程」）の E 列に "1234" を書く（COM 保存の確認用）。
# TASK_INPUT_WORKBOOK は「加工計画DATA」シート付きブック（例: 生産管理_AI配台テスト.xlsm）を指定すること。
# 行は EXCLUDE_RULES_TEST_E1234_ROW（既定 9、2 未満は 9 に丸める）。

# =========================================================
# 【設定】APIキー / 基本ルール / ファイル名
# =========================================================
# 環境変数 GEMINI_API_KEY を推奨（平文でのコード埋め込みは避ける）
API_KEY = os.environ.get("GEMINI_API_KEY", "").strip() or None
if not API_KEY:
    logging.warning("環境変数 GEMINI_API_KEY が未設定です。備考のAI解析はスキップされます。")

GEMINI_MODEL_FLASH = "gemini-2.5-flash"
# 推定料金: USD / 1M tokens（入力, 出力）。公式の最新単価に合わせて更新すること。
# 環境変数 GEMINI_PRICE_USD_IN_PER_M / GEMINI_PRICE_USD_OUT_PER_M で上書き可（Flash 向け）。
_GEMINI_FLASH_IN_PER_M = float(
    os.environ.get("GEMINI_PRICE_USD_IN_PER_M", "0.075") or 0.075
)
_GEMINI_FLASH_OUT_PER_M = float(
    os.environ.get("GEMINI_PRICE_USD_OUT_PER_M", "0.30") or 0.30
)
GEMINI_JPY_PER_USD = float(os.environ.get("GEMINI_JPY_PER_USD", "150") or 150)

# ---------------------------------------------------------------------------
# 以降の定数ブロックは「Excel 列見出し」と 1:1 で対応させる。
# 列名を変える場合は VBA・マクロ側シートと同時に直すこと。
# ---------------------------------------------------------------------------

MASTER_FILE = "master.xlsm" # skillsとattendance(およびtasks)を統合したファイル
# メンバー別勤怠シート: master.xlsm では「休暇区分」と「備考」が別列。
# 勤怠AIの入力は備考のみ。ただし reason（表示・中抜け補正・個人シートの休憩/休暇文言）は、備考が空のとき休暇区分を引き継ぐ。
# master カレンダー／出勤簿.txt 準拠: 前休=午前年休・12:45～17:00（午後休憩14:45～15:00）／後休=8:45～12:00・午後年休／国=他拠点勤務。
# 備考列に文言が入っている行は出勤状態の変更の可能性が高いため、内容に関わらず勤怠AI判定の対象とする（空欄のみスキップ）。
ATT_COL_LEAVE_TYPE = "休暇区分"
ATT_COL_REMARK = "備考"
# need シート: 「基本必要人数」行（A列に「必要人数」を含む）＋ その直下の「配台時追加人数」等（余剰時に増やせる人数・工程×機械列）
# ＋ 行「特別指定1」～「特別指定99」（必要人数の上書き・1～99）
NEED_COL_CONDITION = "依頼NO条件"
NEED_COL_NOTE = "備考"
# need「配台時追加人数」を満枠使っても、単位あたり加工時間が短くなるのは最大でこの割合（例: 0.05 ≒ 5%）
SURPLUS_TEAM_MAX_SPEEDUP_RATIO = 0.05
# タスクは tasks.xlsx を使わず、VBA から渡す TASK_INPUT_WORKBOOK の「加工計画DATA」のみ
TASKS_INPUT_WORKBOOK = os.environ.get("TASK_INPUT_WORKBOOK", "").strip()
TASKS_SHEET_NAME = "加工計画DATA"

# 計画結果ブックの既定フォント（VBA の BIZ_UDP と揃える）。列幅は openpyxl では触らず、取り込みマクロ側の AutoFit に任せる。
OUTPUT_BOOK_FONT_NAME = "BIZ UDPゴシック"
OUTPUT_BOOK_FONT_SIZE = 11
RESULT_SHEET_GANTT_NAME = "結果_設備ガント"

# タスク列名（マクロ実行ブック「加工計画DATA」）
TASK_COL_TASK_ID = "依頼NO"
TASK_COL_MACHINE = "工程名"
TASK_COL_MACHINE_NAME = "機械名"
TASK_COL_QTY = "換算数量"
TASK_COL_ORDER_QTY = "受注数"
TASK_COL_SPEED = "加工速度"
TASK_COL_PRODUCT = "製品名"
TASK_COL_ANSWER_DUE = "回答納期"
TASK_COL_SPECIFIED_DUE = "指定納期"
TASK_COL_RAW_INPUT_DATE = "原反投入日"
# 投入可能日の目安は「回答納期」、未入力時は「指定納期」（前日基準・当日/遅れは最優先）。「加工開始日」列は参照しない。
# 完了判定・進捗（加工計画DATA）
TASK_COL_COMPLETION_FLAG = "加工完了区分"
TASK_COL_ACTUAL_DONE = "実加工数"   # 旧互換（直接の加工済数量）
TASK_COL_ACTUAL_OUTPUT = "実出来高"  # 完成品数量（換算に使う）
TASK_COL_DATA_EXTRACTION_DT = "データ抽出日"
AI_CACHE_TTL_SECONDS = 6 * 60 * 60  # 6時間
# output/ai_remarks_cache.json 内のキー接頭辞（設定_配台不要工程・配台不能ロジック D→E）
AI_CACHE_KEY_PREFIX_EXCLUDE_RULE_DE = "exclude_rule_de_v1"

# マクロブック「加工実績DATA」（Power Query 等で取り込み想定）
ACTUALS_SHEET_NAME = "加工実績DATA"
ACT_COL_TASK_ID = "依頼NO"
ACT_COL_PROCESS = "工程名"
ACT_COL_OPERATOR = "担当者"
ACT_COL_START_DT = "開始日時"
ACT_COL_END_DT = "終了日時"
ACT_COL_START_ALT = "実績開始"
ACT_COL_END_ALT = "実績終了"
ACT_COL_DAY = "日付"
ACT_COL_TIME_START = "開始時刻"
ACT_COL_TIME_END = "終了時刻"
ACTUAL_HEADER_CANONICAL = (
    ACT_COL_TASK_ID,
    ACT_COL_PROCESS,
    ACT_COL_OPERATOR,
    ACT_COL_START_DT,
    ACT_COL_END_DT,
    ACT_COL_START_ALT,
    ACT_COL_END_ALT,
    ACT_COL_DAY,
    ACT_COL_TIME_START,
    ACT_COL_TIME_END,
)

# --- 2段階処理: 段階1抽出 → ブック「配台計画_タスク入力」編集 → 段階2計画 ---
STAGE1_OUTPUT_FILENAME = "plan_input_tasks.xlsx"
PLAN_INPUT_SHEET_NAME = os.environ.get("TASK_PLAN_SHEET", "").strip() or "配台計画_タスク入力"
PLAN_COL_REQUIRED_OP = "必要人数"
# 上書き列「必要OP人数」の左隣の参照列見出し（「（元）必要OP人数」から改称）
PLAN_REF_REQUIRED_OP = "（元）必要人数"
PLAN_COL_SPEED_OVERRIDE = "加工速度_上書き"
PLAN_COL_TASK_EFFICIENCY = "タスク加工効率"
PLAN_COL_PRIORITY = "優先度"
PLAN_COL_SPECIFIED_DUE_OVERRIDE = "指定納期_上書き"
PLAN_COL_START_DATE_OVERRIDE = "加工開始日_指定"
PLAN_COL_START_TIME_OVERRIDE = "加工開始時刻_指定"
PLAN_COL_PREFERRED_OP = "担当OP_指定"
PLAN_COL_SPECIAL_REMARK = "特別指定_備考"
# 参照列「（元）配台不要」は置かない（元データに相当するマスタ列が無いため）。
# セル値の例（配台から外す）: Excel の TRUE / 数値 1 / 文字列「はい」「yes」「true」「○」「〇」「●」等。
# 空・FALSE・0・「いいえ」等は配台対象。詳細は _plan_row_exclude_from_assignment。
PLAN_COL_EXCLUDE_FROM_ASSIGNMENT = "配台不要"
PLAN_COL_AI_PARSE = "AI特別指定_解析"
PLAN_COL_PROCESS_FACTOR = "加工工程の決定プロセスの因子"
DEBUG_TASK_ID = os.environ.get("DEBUG_TASK_ID", "Y3-26").strip()
# 例: set TRACE_TEAM_ASSIGN_TASK_ID=W3-14 … 配台ループで「人数別の最良候補」と採用理由を INFO ログに出す
TRACE_TEAM_ASSIGN_TASK_ID = os.environ.get("TRACE_TEAM_ASSIGN_TASK_ID", "").strip()
# True（既定）: チームは「人数多い→開始早い→本日単位多い→優先度合計小さい」。False: 従来どおり開始時刻最優先
TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF = os.environ.get(
    "TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF", "1"
).strip().lower() not in ("0", "false", "no", "off", "いいえ")

# マクロブック「設定_配台不要工程」: A〜E は通常 Excel COM で保存。COM 失敗時のみ E を TSV→VBA 反映。
EXCLUDE_RULES_SHEET_NAME = "設定_配台不要工程"
EXCLUDE_RULE_COL_PROCESS = "工程名"
EXCLUDE_RULE_COL_MACHINE = "機械名"
EXCLUDE_RULE_COL_FLAG = "配台不要"
EXCLUDE_RULE_COL_LOGIC_JA = "配台不能ロジック"
EXCLUDE_RULE_COL_LOGIC_JSON = "ロジック式"
# 元ブックがロックされ別名保存した場合、同一プロセス内のルール読込はこのパスを優先
_exclude_rules_effective_read_path: str | None = None
# 直後の apply_exclude_rules（同一プロセス）用: VBA 反映前でも E 列付きルールを使う
_exclude_rules_rules_snapshot: list | None = None
_exclude_rules_snapshot_wb: str | None = None
# ルール JSON の conditions で参照可能な列（AI プロンプトと評価器を一致させる）
EXCLUDE_RULE_ALLOWED_COLUMNS = frozenset(
    {
        TASK_COL_TASK_ID,
        TASK_COL_MACHINE,
        TASK_COL_MACHINE_NAME,
        TASK_COL_QTY,
        TASK_COL_ORDER_QTY,
        TASK_COL_SPEED,
        TASK_COL_PRODUCT,
        TASK_COL_ANSWER_DUE,
        TASK_COL_SPECIFIED_DUE,
        TASK_COL_RAW_INPUT_DATE,
        TASK_COL_COMPLETION_FLAG,
        TASK_COL_ACTUAL_DONE,
        TASK_COL_ACTUAL_OUTPUT,
        TASK_COL_DATA_EXTRACTION_DT,
        PLAN_COL_REQUIRED_OP,
        PLAN_COL_SPEED_OVERRIDE,
        PLAN_COL_TASK_EFFICIENCY,
        PLAN_COL_PRIORITY,
        PLAN_COL_SPECIFIED_DUE_OVERRIDE,
        PLAN_COL_START_DATE_OVERRIDE,
        PLAN_COL_START_TIME_OVERRIDE,
        PLAN_COL_PREFERRED_OP,
        PLAN_COL_SPECIAL_REMARK,
        PLAN_COL_PROCESS_FACTOR,
    }
)

# 計画結果ブック「結果_タスク一覧」の列順・表示（マクロ実行ブックの同名シートで上書き可）
RESULT_TASK_SHEET_NAME = "結果_タスク一覧"
COLUMN_CONFIG_SHEET_NAME = "列設定_結果_タスク一覧"
COLUMN_CONFIG_HEADER_COL = "列名"
COLUMN_CONFIG_VISIBLE_COL = "表示"

SOURCE_BASE_COLUMNS = [
    TASK_COL_TASK_ID, TASK_COL_MACHINE, TASK_COL_MACHINE_NAME, TASK_COL_QTY, TASK_COL_ORDER_QTY, TASK_COL_SPEED, TASK_COL_PRODUCT,
    TASK_COL_ANSWER_DUE, TASK_COL_SPECIFIED_DUE, TASK_COL_RAW_INPUT_DATE,
    TASK_COL_COMPLETION_FLAG, TASK_COL_ACTUAL_DONE, TASK_COL_ACTUAL_OUTPUT,
]
PLAN_OVERRIDE_COLUMNS = [
    PLAN_COL_EXCLUDE_FROM_ASSIGNMENT,
    PLAN_COL_REQUIRED_OP, PLAN_COL_SPEED_OVERRIDE, PLAN_COL_TASK_EFFICIENCY,
    PLAN_COL_PRIORITY, PLAN_COL_SPECIFIED_DUE_OVERRIDE, PLAN_COL_START_DATE_OVERRIDE, PLAN_COL_START_TIME_OVERRIDE,
    PLAN_COL_PREFERRED_OP,
    PLAN_COL_SPECIAL_REMARK,
    PLAN_COL_AI_PARSE,
]
# 矛盾検出でリセット対象にする列（見出し行の文言と一致すること）
PLAN_CONFLICT_STYLABLE_COLS = tuple(PLAN_OVERRIDE_COLUMNS)
# 段階1再抽出時、既存「配台計画_タスク入力」から継承する列（AIの解析結果列は毎回空に戻す）
PLAN_STAGE1_MERGE_COLUMNS = tuple(c for c in PLAN_OVERRIDE_COLUMNS if c != PLAN_COL_AI_PARSE)
# openpyxl 保存がブックロックで失敗したとき、VBA が開いているブックへ書式適用するための指示ファイル
PLANNING_CONFLICT_SIDECAR = "planning_conflict_highlight.tsv"


def plan_reference_column_name(override_col: str) -> str:
    """上書き列の左隣に置く参照列の見出し（セル値は括弧付きで元データを表示）。"""
    if override_col == PLAN_COL_REQUIRED_OP:
        return PLAN_REF_REQUIRED_OP
    return f"（元）{override_col}"


def plan_input_sheet_column_order():
    """
    配台計画_タスク入力の列順。
    「配台不要」は先頭（参照列なし）。その他の上書き列は各「（元）…」の右に本体列。
    """
    cols = [PLAN_COL_EXCLUDE_FROM_ASSIGNMENT]
    cols.extend(SOURCE_BASE_COLUMNS)
    cols.append(PLAN_COL_PROCESS_FACTOR)
    for c in PLAN_OVERRIDE_COLUMNS:
        if c == PLAN_COL_EXCLUDE_FROM_ASSIGNMENT:
            continue
        if c == PLAN_COL_AI_PARSE:
            cols.append(c)
        else:
            cols.append(plan_reference_column_name(c))
            cols.append(c)
    return cols


def _format_paren_ref_scalar(val):
    """参照表示用: 空は（―）、日付・その他は（値）。"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "（―）"
    if isinstance(val, datetime):
        d = val.date() if hasattr(val, "date") else val
        if isinstance(d, date):
            return f"（{d.year}/{d.month}/{d.day}）"
    if isinstance(val, date):
        return f"（{val.year}/{val.month}/{val.day}）"
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none"):
        return "（―）"
    return f"（{s}）"


def _reference_text_for_override_row(row, override_col: str, req_map: dict, need_rules: list) -> str:
    """1行分の上書き列に対応する参照文言（括弧付き）。"""
    tid = str(row.get(TASK_COL_TASK_ID, "") or "").strip()
    mach = str(row.get(TASK_COL_MACHINE, "") or "").strip()
    mname = str(row.get(TASK_COL_MACHINE_NAME, "") or "").strip()

    if override_col == PLAN_COL_REQUIRED_OP:
        try:
            n = resolve_need_required_op(mach, mname, tid, req_map, need_rules)
            return f"（{n}）"
        except Exception:
            return "（―）"
    if override_col == PLAN_COL_SPEED_OVERRIDE:
        v = row.get(TASK_COL_SPEED)
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return "（―）"
        try:
            x = float(v)
            if abs(x - round(x)) < 1e-9:
                return f"（{int(round(x))}）"
            return f"（{x}）"
        except (TypeError, ValueError):
            return _format_paren_ref_scalar(v)
    if override_col in (
        PLAN_COL_TASK_EFFICIENCY,
        PLAN_COL_PRIORITY,
        PLAN_COL_START_TIME_OVERRIDE,
        PLAN_COL_PREFERRED_OP,
        PLAN_COL_SPECIAL_REMARK,
    ):
        return "（―）"
    if override_col == PLAN_COL_SPECIFIED_DUE_OVERRIDE:
        sd = parse_optional_date(row.get(TASK_COL_SPECIFIED_DUE))
        if sd is not None:
            return _format_paren_ref_scalar(sd)
        ad = parse_optional_date(row.get(TASK_COL_ANSWER_DUE))
        if ad is not None:
            return _format_paren_ref_scalar(ad)
        return "（―）"
    if override_col == PLAN_COL_START_DATE_OVERRIDE:
        return _format_paren_ref_scalar(parse_optional_date(row.get(TASK_COL_RAW_INPUT_DATE)))
    return "（―）"


def _refresh_plan_reference_columns(df, req_map: dict, need_rules: list):
    """加工計画DATA／need に基づき「（元）…」列を再計算（マージ後に必ず呼ぶ）。"""
    if df is None or df.empty:
        return df
    need_rules = need_rules or []
    req_map = req_map or {}
    for i in df.index:
        row = df.loc[i]
        for oc in PLAN_OVERRIDE_COLUMNS:
            if oc == PLAN_COL_AI_PARSE:
                continue
            if oc == PLAN_COL_EXCLUDE_FROM_ASSIGNMENT:
                continue
            ref_col = plan_reference_column_name(oc)
            if ref_col not in df.columns:
                continue
            df.at[i, ref_col] = _reference_text_for_override_row(row, oc, req_map, need_rules)
    return df


def _apply_plan_input_visual_format(path: str, sheet_name: str = "タスク一覧"):
    """上書き入力列に薄い黄色を付与（参照列は未着色。AI解析列は除外）。"""
    # 見出し文字の表記ゆれで列名検索に失敗しがちなため、段階1の列順（plan_input_sheet_column_order）の
    # 1-based 列番号で塗る（to_excel の列順と一致させる）。
    fill_yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    order = plan_input_sheet_column_order()
    col_1based = {name: i + 1 for i, name in enumerate(order)}
    wb = load_workbook(path)
    try:
        if sheet_name not in wb.sheetnames:
            return
        ws = wb[sheet_name]
        last_row = ws.max_row or 1
        if last_row < 2:
            return
        for oc in PLAN_OVERRIDE_COLUMNS:
            if oc == PLAN_COL_AI_PARSE:
                continue
            ci = col_1based.get(oc)
            if not ci:
                continue
            for r in range(2, last_row + 1):
                ws.cell(row=r, column=ci).fill = fill_yellow
        wb.save(path)
    finally:
        wb.close()


def _planning_conflict_sidecar_path():
    return os.path.join(log_dir, PLANNING_CONFLICT_SIDECAR)


def _remove_planning_conflict_sidecar_safe():
    try:
        os.remove(_planning_conflict_sidecar_path())
    except OSError:
        pass


def write_planning_conflict_highlight_sidecar(sheet_name, num_data_rows, conflicts_by_row):
    """
    Excel がブックを開いたままのとき保存できない場合に、VBA 用の TSV を log に書く。
    形式: V1 / シート名 / データ行数 / クリア列をタブ結合 / 以降 行番号\\t列名
    """
    path = _planning_conflict_sidecar_path()
    clear_cols = "\t".join(PLAN_CONFLICT_STYLABLE_COLS)
    lines = ["V1", sheet_name, str(int(num_data_rows)), clear_cols]
    for r in sorted(conflicts_by_row.keys()):
        for name in sorted(conflicts_by_row[r]):
            lines.append(f"{int(r)}\t{name}")
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines) + "\n")

# 段階1出力・ブック内の日付列を Excel 上「日付のみ」(時刻なし表示) に整える
STAGE1_SHEET_DATEONLY_HEADERS = frozenset(
    {
        TASK_COL_ANSWER_DUE,
        TASK_COL_SPECIFIED_DUE,
        TASK_COL_RAW_INPUT_DATE,
        PLAN_COL_SPECIFIED_DUE_OVERRIDE,
        PLAN_COL_START_DATE_OVERRIDE,
    }
)


def _output_book_font(bold=False):
    return Font(name=OUTPUT_BOOK_FONT_NAME, size=OUTPUT_BOOK_FONT_SIZE, bold=bold)


def _apply_output_font_to_result_sheet(ws):
    """結果_* のうちガント以外向け: 既定フォント・1行目太字のみ（列幅は VBA AutoFit）。"""
    base = _output_book_font(bold=False)
    hdr = _output_book_font(bold=True)
    mr, mc = ws.max_row or 1, ws.max_column or 1
    for row in ws.iter_rows(min_row=1, max_row=mr, min_col=1, max_col=mc):
        for cell in row:
            cell.font = base
    for cell in ws[1]:
        cell.font = hdr


def _apply_excel_date_columns_date_only_display(path, sheet_name, header_names=None):
    """openpyxl: 指定ヘッダー列を yyyy/mm/dd の日付表示にする（時刻を表示しない）。"""
    from openpyxl import load_workbook

    headers = header_names or STAGE1_SHEET_DATEONLY_HEADERS
    wb = load_workbook(path)
    try:
        ws = wb[sheet_name] if isinstance(sheet_name, str) else wb.worksheets[int(sheet_name)]
        cmap = {}
        for cell in ws[1]:
            if cell.value is not None:
                cmap[str(cell.value).strip()] = cell.column
        fmt = "yyyy/mm/dd"
        for h in headers:
            col = cmap.get(h)
            if not col:
                continue
            for r in range(2, ws.max_row + 1):
                c = ws.cell(row=r, column=col)
                v = c.value
                if v is None:
                    continue
                if isinstance(v, datetime):
                    c.value = v.date()
                elif isinstance(v, date):
                    pass
                else:
                    try:
                        d0 = pd.to_datetime(v, errors="coerce")
                        if pd.isna(d0):
                            continue
                        c.value = d0.date()
                    except Exception:
                        continue
                c.number_format = fmt
        wb.save(path)
    finally:
        wb.close()


def _extract_data_extraction_datetime():
    """
    `加工計画DATA` シートの `データ抽出日` から datetime を取得する。
    """
    try:
        if not TASKS_INPUT_WORKBOOK or not os.path.exists(TASKS_INPUT_WORKBOOK):
            return None
        df = pd.read_excel(TASKS_INPUT_WORKBOOK, sheet_name=TASKS_SHEET_NAME)
        df.columns = df.columns.str.strip()
        if TASK_COL_DATA_EXTRACTION_DT not in df.columns:
            return None
        s = df[TASK_COL_DATA_EXTRACTION_DT]
        first = None
        for v in s:
            if v is None or (isinstance(v, float) and pd.isna(v)):
                continue
            first = v
            break
        if first is None:
            return None
        dt = pd.to_datetime(first, errors="coerce")
        if pd.isna(dt):
            return None
        if isinstance(dt, pd.Timestamp):
            return dt.to_pydatetime()
        return dt if isinstance(dt, datetime) else None
    except Exception:
        return None


def _extract_data_extraction_datetime_str():
    """
    `加工計画DATA` シートの `データ抽出日` から「データ吸出し日時」を取得して文字列化する。
    """
    try:
        dt = _extract_data_extraction_datetime()
        if dt is None:
            return "—"
        return dt.strftime("%Y/%m/%d %H:%M:%S")
    except Exception:
        return "—"


def _weekday_jp(d):
    return "月火水木金土日"[d.weekday()]


# ガントの作業バー：いずれも明るい地色＋黒文字が読めるトーン（モノクロ印刷でも濃淡で識別しやすい）
_GANTT_BAR_FILLS_PRINT_SAFE = (
    "E8E8E8",
    "D8E4EF",
    "E6E2DB",
    "DEEADF",
    "E8E0E8",
    "EAE8D8",
    "DDE6EA",
    "E5DCE5",
)

# 実績バー用（計画と並べてもモノクロで区別しやすいトーン）
_GANTT_BAR_FILLS_ACTUAL = (
    "D4E4D4",
    "C9DDE8",
    "DED8CC",
    "D2E5CD",
    "DAD2D9",
    "E0DCCF",
    "CDE2E8",
    "DCD2DC",
)


def _gantt_bar_fill_for_task_id(task_id):
    """依頼NOごとに上記パレットから1色（RRGGBB）。濃色＋白文字の組み合わせは使わない。"""
    h = hashlib.md5(str(task_id).encode("utf-8")).hexdigest()
    i = int(h[:8], 16) % len(_GANTT_BAR_FILLS_PRINT_SAFE)
    return _GANTT_BAR_FILLS_PRINT_SAFE[i]


def _gantt_bar_fill_actual_for_task_id(task_id):
    h = hashlib.md5(str(task_id).encode("utf-8")).hexdigest()
    i = int(h[:8], 16) % len(_GANTT_BAR_FILLS_ACTUAL)
    return _GANTT_BAR_FILLS_ACTUAL[i]


def _gantt_slot_state_tuple(evlist, slot_mid, task_fill_fn=None):
    """スロット中央時刻における1マス分の状態。('idle',) | ('break',) | ('task', tid, fill_hex, pct)"""
    fill_fn = task_fill_fn or _gantt_bar_fill_for_task_id
    active = None
    for e in evlist:
        if e["start_dt"] <= slot_mid < e["end_dt"]:
            active = e
            break
    if active is None:
        return ("idle",)
    if any(b_s <= slot_mid < b_e for b_s, b_e in active.get("breaks") or ()):
        return ("break",)
    tid = str(active["task_id"])
    gh = fill_fn(active["task_id"])
    pct = None
    try:
        # 「マクロ実行時点」の完了率を優先（pct_macro を timeline_event に持たせる）
        if active.get("pct_macro") is not None:
            pct = int(round(parse_float_safe(active.get("pct_macro"), 0.0)))
            pct = max(0, min(100, pct))
        else:
            # フェイルセーフ（従来の擬似進捗計算）
            tot = parse_float_safe(active.get("total_units"), 0.0)
            done = parse_float_safe(active.get("already_done_units"), 0.0) + parse_float_safe(
                active.get("units_done"), 0.0
            )
            if tot > 0:
                pct = max(0, min(100, int(round((done / tot) * 100))))
    except Exception:
        pct = None
    return ("task", tid, gh, pct)


def _gantt_merge_key(st):
    if st[0] == "idle":
        return ("idle",)
    if st[0] == "break":
        return ("break",)
    return ("task", st[1])


def _paint_gantt_timeline_row_merged(
    ws,
    row,
    n_fixed,
    slots,
    slot_mins,
    evlist,
    idle_fill,
    break_fill,
    gantt_label_font,
    grid_border,
    task_fill_fn=None,
    label_font=None,
):
    """
    時間軸を塗り分けたうえで、同一状態が連続するセルを横結合し帯状のバーにする。
    （細マス単体の塗りではなく15分刻み＋同一状態のセル結合で、帯状のバーとして表現する）
    """
    bar_label_font = label_font or gantt_label_font
    n_slots = len(slots)
    states = []
    for slot_start in slots:
        mid = slot_start + timedelta(minutes=slot_mins / 2)
        states.append(_gantt_slot_state_tuple(evlist, mid, task_fill_fn))
    tcol0 = n_fixed + 1
    i = 0
    while i < n_slots:
        mk = _gantt_merge_key(states[i])
        j = i + 1
        while j < n_slots and _gantt_merge_key(states[j]) == mk:
            j += 1
        col_s = tcol0 + i
        col_e = tcol0 + j - 1
        st0 = states[i]
        if col_s < col_e:
            ws.merge_cells(start_row=row, start_column=col_s, end_row=row, end_column=col_e)
        c = ws.cell(row=row, column=col_s)
        c.border = grid_border
        # 左寄せ・折り返しなし・縮小なし（長いラベルはセル外へはみ出し表示しやすくする）
        c.alignment = Alignment(
            horizontal="left",
            vertical="center",
            wrap_text=False,
            shrink_to_fit=False,
            indent=1,
        )
        if st0[0] == "idle":
            c.fill = idle_fill
        elif st0[0] == "break":
            c.fill = break_fill
        else:
            _, tid, gh, pct = st0
            c.fill = PatternFill(fill_type="solid", start_color=gh, end_color=gh)
            c.value = f"{tid[:9]} {pct}%" if pct is not None else tid[:9]
            c.font = bar_label_font
        i = j


def _write_results_equipment_gantt_sheet(
    writer,
    timeline_events,
    equipment_list,
    sorted_dates,
    attendance_data,
    data_extract_dt_str,
    base_now_dt=None,
    actual_timeline_events=None,
):
    """
    結果_設備毎の時間割と同一データ源（timeline_events）に基づき、
    設備×横軸時間のガンチャート風シートを追加する。
    横軸は15分刻み。連続する同一タスク／休憩／空きはセル結合して帯状に表示する。
    actual_timeline_events があれば設備ごとに「実績」行を計画行の下へ追加する。
    """
    wb = writer.book
    try:
        insert_at = wb.sheetnames.index("結果_設備毎の時間割") + 1
    except ValueError:
        insert_at = len(wb.sheetnames)
    ws = wb.create_sheet("結果_設備ガント", insert_at)
    try:
        ws.sheet_properties.tabColor = "7F7F7F"
    except Exception:
        pass

    events_by_date = defaultdict(list)
    for e in timeline_events:
        events_by_date[e["date"]].append(e)

    by_dm = defaultdict(lambda: defaultdict(list))
    for e in timeline_events:
        by_dm[e["date"]][e["machine"]].append(e)
    for d0 in by_dm:
        for mk in by_dm[d0]:
            by_dm[d0][mk].sort(key=lambda x: x["start_dt"])

    by_dm_actual = defaultdict(lambda: defaultdict(list))
    show_actual_rows = bool(actual_timeline_events)
    actual_events_by_date = defaultdict(list)
    if show_actual_rows:
        for e in actual_timeline_events:
            actual_events_by_date[e["date"]].append(e)
            by_dm_actual[e["date"]][e["machine"]].append(e)
        for d0 in by_dm_actual:
            for mk in by_dm_actual[d0]:
                by_dm_actual[d0][mk].sort(key=lambda x: x["start_dt"])

    slot_mins = 15
    fn = OUTPUT_BOOK_FONT_NAME
    hdr_font = Font(name=fn, bold=True, color="000000", size=10)
    hdr_fill = PatternFill(fill_type="solid", start_color="D9D9D9", end_color="D9D9D9")
    hdr_time_font = Font(name=fn, bold=True, color="000000", size=9)
    title_font = Font(name=fn, bold=True, size=20, color="1A1A1A")
    title_fill = PatternFill(fill_type="solid", start_color="DDDDDD", end_color="DDDDDD")
    meta_font = Font(name=fn, size=9, color="333333")
    meta_fill = PatternFill(fill_type="solid", start_color="F3F3F3", end_color="F3F3F3")
    day_banner_font = Font(name=fn, bold=True, size=11, color="1A1A1A")
    day_banner_fill = PatternFill(fill_type="solid", start_color="D0D0D0", end_color="D0D0D0")
    accent_left = Side(style="thick", color="2B2B2B")
    banner_sep = Side(style="thin", color="7A7A7A")
    thin = Side(style="thin", color="666666")
    grid_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    idle_fill = PatternFill(fill_type="solid", start_color="FFFFFF", end_color="FFFFFF")
    break_fill = PatternFill(fill_type="solid", start_color="B8B8B8", end_color="B8B8B8")
    gantt_label_font = Font(name=fn, size=8, bold=True, color="000000")
    gantt_label_font_actual = Font(name=fn, size=8, bold=True, color="000000", italic=True)

    # 横軸(10分刻み)は日付で共通のため、slot_times を先に確定
    base_dt = base_now_dt if isinstance(base_now_dt, datetime) else datetime.now()
    dummy_d = sorted_dates[0] if sorted_dates else base_dt.date()
    d_start0 = datetime.combine(dummy_d, DEFAULT_START_TIME)
    d_end0 = datetime.combine(dummy_d, DEFAULT_END_TIME)
    slot_times = []
    t0 = d_start0
    while t0 < d_end0:
        slot_times.append(t0.time())
        t0 += timedelta(minutes=slot_mins)

    n_slots = len(slot_times)
    n_fixed = 3  # 設備 / タスク概要 / 主担当（日付は日ブロック見出しのみ）
    last_col = n_fixed + n_slots

    # タイトル＆日時（ページ上部）
    create_ts = base_dt.strftime("%Y/%m/%d %H:%M:%S")
    master_path = os.path.join(os.getcwd(), MASTER_FILE) if MASTER_FILE else ""

    def _fmt_mtime(p):
        try:
            if p and os.path.exists(p):
                return datetime.fromtimestamp(os.path.getmtime(p)).strftime("%Y/%m/%d %H:%M:%S")
        except Exception:
            pass
        return "—"

    master_mtime = _fmt_mtime(master_path)

    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    tcell = ws.cell(row=row, column=1, value="湖南工場 加工計画")
    tcell.font = title_font
    tcell.fill = title_fill
    # 結合セルでも左端から表示（縮小・折り返しなし）
    tcell.alignment = Alignment(
        horizontal="left",
        vertical="center",
        wrap_text=False,
        shrink_to_fit=False,
        indent=1,
    )
    tcell.border = Border(left=accent_left, bottom=banner_sep)
    ws.row_dimensions[row].height = 34
    row += 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    meta_line = (
        f"作成　{create_ts}"
        f"　・　データ吸出し　{data_extract_dt_str or '—'}"
        f"　・　マスタ（master.xlsm）　{master_mtime}"
    )
    mtop = ws.cell(row=row, column=1, value=meta_line)
    mtop.font = meta_font
    mtop.fill = meta_fill
    mtop.alignment = Alignment(
        horizontal="left",
        vertical="center",
        indent=1,
        wrap_text=False,
        shrink_to_fit=False,
    )
    mtop.border = Border(left=accent_left, bottom=banner_sep)
    ws.row_dimensions[row].height = 22
    row += 1

    first_freeze_set = False

    for d in sorted_dates:
        evs = events_by_date.get(d, [])
        a_evs_day = actual_events_by_date.get(d, []) if show_actual_rows else []
        is_anyone_working = any(
            attendance_data[d][mm]["is_working"] for mm in attendance_data[d] if mm in attendance_data[d]
        )
        if not evs and not a_evs_day and not is_anyone_working:
            continue

        slots = [datetime.combine(d, tm) for tm in slot_times]

        # 日付見出し（左の固定列幅に合わせて A〜C に表示）
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_fixed)
        ban = ws.cell(row=row, column=1, value=f"▶ {d.strftime('%Y/%m/%d')}　{_weekday_jp(d)}")
        ban.font = day_banner_font
        ban.fill = day_banner_fill
        ban.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ban.border = Border(left=accent_left, bottom=thin)
        ws.row_dimensions[row].height = 22
        row += 1

        fixed_hdr = ["設備", "タスク概要", "主担当"]
        for ci, h in enumerate(fixed_hdr, 1):
            c = ws.cell(row=row, column=ci, value=h)
            c.font = hdr_font
            c.fill = hdr_fill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
        for si, st in enumerate(slots):
            c = ws.cell(row=row, column=n_fixed + 1 + si, value=st.strftime("%H:%M"))
            c.font = hdr_time_font
            c.fill = hdr_fill
            c.alignment = Alignment(horizontal="center", vertical="bottom", textRotation=90)
        ws.row_dimensions[row].height = 38
        row += 1

        data_row0 = row
        if not first_freeze_set:
            ws.freeze_panes = f"{get_column_letter(n_fixed + 1)}{data_row0}"
            first_freeze_set = True

        zebra = False
        for eq in equipment_list:
            zebra = not zebra
            evlist = by_dm[d].get(eq, [])
            if evlist:
                parts = []
                ops = []
                for e in evlist:
                    cum = e["already_done_units"] + e["units_done"]
                    tot = e["total_units"]
                    parts.append(f"[{e['task_id']}] {cum}/{tot}R")
                    ops.append(str(e["op"]))
                task_sum = " ｜ ".join(parts)[:120]
                op_disp = ops[0] if len(set(ops)) == 1 else ",".join(dict.fromkeys(ops))
            else:
                task_sum = "—"
                op_disp = "—"

            lab_fill = PatternFill(fill_type="solid", start_color="FAFAFA", end_color="FAFAFA")
            if zebra:
                lab_fill = PatternFill(fill_type="solid", start_color="F0F4FA", end_color="F0F4FA")

            c1 = ws.cell(row=row, column=1, value=eq)
            c2 = ws.cell(row=row, column=2, value=task_sum)
            c3 = ws.cell(row=row, column=3, value=op_disp)
            for c in (c1, c2, c3):
                c.font = Font(name=fn, size=10, color="000000")
                c.fill = lab_fill
                c.border = grid_border
                c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
            c1.font = Font(name=fn, size=10, bold=True, color="000000")

            _paint_gantt_timeline_row_merged(
                ws,
                row,
                n_fixed,
                slots,
                slot_mins,
                evlist,
                idle_fill,
                break_fill,
                gantt_label_font,
                grid_border,
            )

            ws.row_dimensions[row].height = 22
            row += 1

            if show_actual_rows:
                evlist_a = by_dm_actual[d].get(eq, [])
                if evlist_a:
                    parts_a = [f"[{e_a['task_id']}]" for e_a in evlist_a]
                    task_sum_a = " ".join(dict.fromkeys(parts_a))[:120]
                    ops_a = [
                        str(e_a.get("op") or "").strip()
                        for e_a in evlist_a
                        if str(e_a.get("op") or "").strip()
                    ]
                    op_disp_a = (
                        ops_a[0]
                        if len(set(ops_a)) == 1
                        else ",".join(dict.fromkeys(ops_a))
                    )
                else:
                    task_sum_a = "—"
                    op_disp_a = "—"

                lab_fill_a = PatternFill(
                    fill_type="solid", start_color="EEF6EE", end_color="EEF6EE"
                )
                if zebra:
                    lab_fill_a = PatternFill(
                        fill_type="solid", start_color="E0EEE0", end_color="E0EEE0"
                    )

                ca1 = ws.cell(row=row, column=1, value=f"{eq}（実績）")
                ca2 = ws.cell(row=row, column=2, value=task_sum_a)
                ca3 = ws.cell(row=row, column=3, value=op_disp_a)
                for c in (ca1, ca2, ca3):
                    c.font = Font(name=fn, size=10, color="000000")
                    c.fill = lab_fill_a
                    c.border = grid_border
                    c.alignment = Alignment(
                        horizontal="left", vertical="center", wrap_text=False
                    )
                ca1.font = Font(name=fn, size=10, bold=True, color="000000", italic=True)

                _paint_gantt_timeline_row_merged(
                    ws,
                    row,
                    n_fixed,
                    slots,
                    slot_mins,
                    evlist_a,
                    idle_fill,
                    break_fill,
                    gantt_label_font_actual,
                    grid_border,
                    task_fill_fn=_gantt_bar_fill_actual_for_task_id,
                    label_font=gantt_label_font_actual,
                )

                ws.row_dimensions[row].height = 22
                row += 1
        # 凡例は高さ確保のため省略（モノクロ印刷は色の濃淡/セルの枠で識別）
    # 列幅は VBA 取り込み時（結果_設備ガント_列幅を設定）で設定

    try:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToHeight = False
        ws.page_setup.fitToWidth = 1
        # A3（openpyxl 上で paperSize=8 が A3 相当）
        ws.page_setup.paperSize = 8
        # 余白を狭めて横1ページに収まりやすくする（単位: インチ）
        ws.page_margins.left = 0.2
        ws.page_margins.right = 0.2
        ws.page_margins.top = 0.2
        ws.page_margins.bottom = 0.2
        # タイトル・表をページ左基準に（レポート風）
        ws.print_options.horizontalCentered = False
        ws.print_options.verticalCentered = False
        ws.print_options.gridLines = False
    except Exception:
        pass


def row_has_completion_keyword(row):
    """加工完了区分に「完了」の文字が含まれる場合はタスク完了とみなす。"""
    v = row.get(TASK_COL_COMPLETION_FLAG)
    if v is None or pd.isna(v):
        return False
    return "完了" in str(v)


def _plan_row_exclude_from_assignment(row) -> bool:
    """
    「配台不要」列がオンなら、その行は配台キューへ入れず、特別指定_備考の AI 解析行からも除く。

    配台から外す（真）: 論理値 True、数値 1、文字列（NFKC 後・小文字）
      true / 1 / yes / on / y / t / はい / ○ / 〇 / ●
    配台対象（偽）: 空、None、False、0、no / off / false / いいえ / 否 等
    上記以外の文字列は偽（配台する）。チェックボックス連動セルは通常 TRUE/FALSE または 1/0。
    """
    v = row.get(PLAN_COL_EXCLUDE_FROM_ASSIGNMENT)
    if v is True:
        return True
    if v is False:
        return False
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return False
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            iv = int(v)
            if iv == 1:
                return True
            if iv == 0:
                return False
        except (TypeError, ValueError):
            pass
    s = unicodedata.normalize("NFKC", str(v).strip()).lower()
    if not s or s in ("nan", "none", "false", "0", "no", "off", "いいえ", "否"):
        return False
    if s in ("true", "1", "yes", "on", "はい", "y", "t", "○", "〇", "●"):
        return True
    return False


def _coerce_plan_exclude_column_value_for_storage(v):
    """
    「配台不要」列へ書き込む値を、StringDtype 列でも代入エラーにならない形にそろえる。
    Excel 取り込みの True / 1 / False / 0 と文字列を保持し、_plan_row_exclude_from_assignment と整合する。
    """
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    if v is True:
        return "yes"
    if v is False:
        return ""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            iv = int(v)
            if iv == 1:
                return "yes"
            if iv == 0:
                return ""
        except (TypeError, ValueError):
            pass
    return str(v).strip()


def parse_float_safe(val, default=0.0):
    try:
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return default
        return float(val)
    except (TypeError, ValueError):
        return default


def calc_done_qty_equivalent_from_row(row):
    """
    加工済数量（工程投入量換算）を返す。

    基本式:
      実出来高 ÷ (受注数 ÷ 換算数量)
    = 実出来高 * 換算数量 / 受注数

    受注数が無い/不正な場合は、旧列「実加工数」を互換フォールバックとして使う。
    """
    qty_total = parse_float_safe(row.get(TASK_COL_QTY), 0.0)
    order_qty = parse_float_safe(row.get(TASK_COL_ORDER_QTY), 0.0)
    actual_output = parse_float_safe(row.get(TASK_COL_ACTUAL_OUTPUT), 0.0)
    legacy_done = parse_float_safe(row.get(TASK_COL_ACTUAL_DONE), 0.0)

    if qty_total <= 0:
        return max(0.0, legacy_done)

    if order_qty > 0 and actual_output >= 0:
        done_qty = actual_output * qty_total / order_qty
        return max(0.0, done_qty)

    return max(0.0, legacy_done)


def parse_optional_int(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", ""):
        return None
    try:
        return int(round(float(s)))
    except (TypeError, ValueError):
        return None


def parse_optional_date(val):
    if val is None or pd.isna(val):
        return None
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return None
    try:
        return pd.to_datetime(val).date()
    except Exception:
        return None


def load_ai_cache():
    try:
        if os.path.exists(ai_cache_path):
            with open(ai_cache_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    # 期限切れエントリを除去（6時間）
                    now_ts = time_module.time()
                    cleaned = {}
                    expired_count = 0
                    for k, v in data.items():
                        # 新形式: {"ts": epoch_seconds, "data": {...}}
                        if isinstance(v, dict) and "ts" in v and "data" in v:
                            ts = parse_float_safe(v.get("ts"), 0.0)
                            if ts > 0 and (now_ts - ts) <= AI_CACHE_TTL_SECONDS:
                                cleaned[k] = v
                            else:
                                expired_count += 1
                        # 旧形式: 値が直接AI結果dict（互換で読み取り、即時に新形式へ再保存される）
                        else:
                            cleaned[k] = {"ts": now_ts, "data": v}
                    if expired_count > 0:
                        logging.info(f"AIキャッシュ期限切れを削除: {expired_count}件")
                    return cleaned
    except Exception as e:
        logging.warning(f"AIキャッシュ読み込み失敗: {e}")
    return {}

def save_ai_cache(cache_obj):
    try:
        with open(ai_cache_path, "w", encoding="utf-8") as f:
            json.dump(cache_obj, f, ensure_ascii=False)
    except Exception as e:
        logging.warning(f"AIキャッシュ保存失敗: {e}")

def get_cached_ai_result(cache_obj, cache_key, content_key=None):
    """
    content_key: オプション。保存時と同一の文字列でないヒットは無効化する（特別指定・照合用の二次チェック）。
    旧エントリに content_key が無い場合は SHA256 キー一致のみで従来どおりヒットとみなす。
    """
    entry = cache_obj.get(cache_key)
    if not isinstance(entry, dict):
        return None
    ts = parse_float_safe(entry.get("ts"), 0.0)
    if ts <= 0:
        return None
    if (time_module.time() - ts) > AI_CACHE_TTL_SECONDS:
        return None
    if content_key is not None:
        stored_ck = entry.get("content_key")
        if stored_ck is not None and stored_ck != content_key:
            logging.info(
                "AIキャッシュ: キーは一致しますが content_key が現行入力と異なるため無効化します。"
            )
            return None
    data = entry.get("data")
    if isinstance(data, dict):
        return data
    return None

def put_cached_ai_result(cache_obj, cache_key, parsed_obj, content_key=None):
    payload = {"ts": time_module.time(), "data": parsed_obj}
    if content_key is not None:
        payload["content_key"] = content_key
    cache_obj[cache_key] = payload

def extract_retry_seconds(err_text):
    # 例: "Please retry in 57.089735313s."
    m = re.search(r"retry in ([0-9]+(?:\.[0-9]+)?)s", err_text, re.IGNORECASE)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            pass
    # 例: "'retryDelay': '57s'"
    m = re.search(r"retryDelay'\s*:\s*'([0-9]+)s'", err_text)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            pass
    return None


def infer_unit_m_from_product_name(product_name, fallback_unit):
    """
    製品名文字列から加工単位(m)を推定する暫定ルール。
    例: 15020-JX5R- 770X300F-A   R -> 300
    ※ バリエーションが多い前提のため、ここを都度調整できるよう関数化している。
    """
    if product_name is None or pd.isna(product_name):
        return fallback_unit
    s = str(product_name)
    # "770X300..." のようなパターンから X の後の数値を拾う（最後に見つかったXを優先）
    matches = re.findall(r"[xX]\s*(\d{2,6})", s)
    if matches:
        try:
            v = int(matches[-1])
            if v > 0:
                return v
        except ValueError:
            pass
    return fallback_unit

def load_tasks_df():
    """
    タスク入力を取得する（tasks.xlsx は使用しない）。
    必須: 環境変数 TASK_INPUT_WORKBOOK にマクロ実行ブックのフルパス（VBA が設定）
         シート「加工計画DATA」を読み込む（投入目安は「回答納期」、未入力時は「指定納期」）。
    """
    if not TASKS_INPUT_WORKBOOK:
        raise FileNotFoundError(
            "TASK_INPUT_WORKBOOK が未設定です。VBA の RunPython でマクロ実行ブックのパスを渡してください。"
        )
    if not os.path.exists(TASKS_INPUT_WORKBOOK):
        raise FileNotFoundError(f"TASK_INPUT_WORKBOOK が存在しません: {TASKS_INPUT_WORKBOOK}")
    df = pd.read_excel(TASKS_INPUT_WORKBOOK, sheet_name=TASKS_SHEET_NAME)
    df.columns = df.columns.str.strip()
    logging.info(f"タスク入力: '{TASKS_INPUT_WORKBOOK}' の '{TASKS_SHEET_NAME}' を読み込みました。")
    return df


def _nfkc_column_aliases(canonical_name):
    """見出しの表記ゆれ（全角記号・互換文字）を吸収するための比較キー。"""
    return unicodedata.normalize("NFKC", str(canonical_name).strip())


def _rename_legacy_plan_input_ref_columns(df):
    """旧参照列見出し「（元）必要OP人数」を現行「（元）必要人数」へ寄せる。"""
    if df is None or df.empty:
        return df
    legacy = unicodedata.normalize("NFKC", "（元）必要OP人数")
    renames = {}
    for c in df.columns:
        if _nfkc_column_aliases(c) == legacy:
            renames[c] = PLAN_REF_REQUIRED_OP
    if renames:
        df = df.rename(columns=renames)
    return df


def _align_dataframe_headers_to_canonical(df, canonical_names):
    """列名を NFKC 一致で canonical に寄せる（Excel 側が全角 '_' 等でも読めるように）。"""
    key_to_canonical = {_nfkc_column_aliases(c): c for c in canonical_names}
    rename_map = {}
    for col in df.columns:
        k = _nfkc_column_aliases(col)
        if k in key_to_canonical:
            target = key_to_canonical[k]
            if col != target:
                rename_map[col] = target
    if rename_map:
        df = df.rename(columns=rename_map)
    return df


def _normalize_equipment_match_key(val):
    """
    工程名（設備名）の照合用キー。
    NFKC・前後空白・連続空白・NBSP/全角スペース・ゼロ幅文字を正規化する。
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    t = unicodedata.normalize("NFKC", str(val))
    t = t.replace("\u00a0", " ").replace("\u3000", " ")
    t = re.sub(r"[\u200b\u200c\u200d\ufeff]", "", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def _equipment_lookup_normalized_to_canonical(equipment_list):
    """正規化キー → master スキルシート上の列名（canonical 表記）。"""
    lookup = {}
    for eq in equipment_list:
        k = _normalize_equipment_match_key(eq)
        if k and k not in lookup:
            lookup[k] = eq
    return lookup


def load_planning_tasks_df():
    """
    2段階目用: マクロブック上の「配台計画_タスク入力」シートを読み込む。

    「担当OP_指定」列または特別指定備考の AI 出力 preferred_operator で主担当 OP を指名できる（skills のメンバー名とあいまい一致）。
    メイン「再優先特別記載」の task_preferred_operators は generate_plan 側で最優先マージされる。
    「配台不要」がオン（TRUE/1/はい 等）の行は配台対象外。
    読み込み後、同一依頼NO・重複機械名があるグループの工程「分割」行へ空なら「配台不要」=yes（段階1と同じ）。
    「設定_配台不要工程」で工程+機械の組を同期し、C/D/E に基づき配台不要を反映する（シート作成は VBA）。
    """
    if not TASKS_INPUT_WORKBOOK:
        raise FileNotFoundError(
            "TASK_INPUT_WORKBOOK が未設定です。VBA の RunPython でマクロ実行ブックのパスを渡してください。"
        )
    if not os.path.exists(TASKS_INPUT_WORKBOOK):
        raise FileNotFoundError(f"TASK_INPUT_WORKBOOK が存在しません: {TASKS_INPUT_WORKBOOK}")
    df = pd.read_excel(TASKS_INPUT_WORKBOOK, sheet_name=PLAN_INPUT_SHEET_NAME)
    df.columns = df.columns.str.strip()
    df = _rename_legacy_plan_input_ref_columns(df)
    df = _align_dataframe_headers_to_canonical(
        df, plan_input_sheet_column_order()
    )
    for c in plan_input_sheet_column_order():
        if c not in df.columns:
            df[c] = ""
    try:
        _pairs_lr = []
        _seen_lr = set()
        for _, _row_lr in df.iterrows():
            _p = str(_row_lr.get(TASK_COL_MACHINE, "") or "").strip()
            _m = str(_row_lr.get(TASK_COL_MACHINE_NAME, "") or "").strip()
            if not _p:
                continue
            _k = (
                _normalize_process_name_for_rule_match(_p),
                _normalize_equipment_match_key(_m),
            )
            if _k in _seen_lr:
                continue
            _seen_lr.add(_k)
            _pairs_lr.append((_p, _m))
        run_exclude_rules_sheet_maintenance(TASKS_INPUT_WORKBOOK, _pairs_lr, "配台シート読込")
    except Exception:
        logging.exception("配台シート読込: 設定_配台不要工程の保守で例外（続行）")
    try:
        _apply_auto_exclude_bunkatsu_duplicate_machine(df, log_prefix="配台シート読込")
    except Exception as ex:
        logging.warning("配台シート読込: 分割行の配台不要自動設定で例外（続行）: %s", ex)
    try:
        df = apply_exclude_rules_config_to_plan_df(df, TASKS_INPUT_WORKBOOK, "配台シート読込")
    except Exception as ex:
        logging.warning("配台シート読込: 設定シートによる配台不要適用で例外（続行）: %s", ex)
    logging.info(
        f"計画タスク入力: '{TASKS_INPUT_WORKBOOK}' の '{PLAN_INPUT_SHEET_NAME}' を読み込みました。"
    )
    return df


def _main_sheet_cell_is_global_comment_label(val) -> bool:
    """メインシート上「グローバルコメント」見出しセルか（表記ゆれ許容）。"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return False
    s = unicodedata.normalize("NFKC", str(val).strip())
    if not s:
        return False
    if _nfkc_column_aliases(s) == _nfkc_column_aliases("グローバルコメント"):
        return True
    if "グローバル" in s and "コメント" in s:
        return True
    return False


def load_main_sheet_global_priority_override_text() -> str:
    """
    TASK_INPUT_WORKBOOK のメインシートで「グローバルコメント」と書かれたセルの **直下** を読む。
    シート名: 「メイン」「Main」、または名前に「メイン」を含む（VBA GetMainWorksheet と同趣旨）。

    同じ欄に「4/1は工場臨時休業」などと書いた場合、parse_factory_closure_dates_from_global_comment により
    その日は全員非稼働（配台で加工しない）として扱われる。
    """
    wb_path = TASKS_INPUT_WORKBOOK.strip() if TASKS_INPUT_WORKBOOK else ""
    if not wb_path or not os.path.exists(wb_path):
        return ""
    try:
        wb = load_workbook(wb_path, data_only=True, read_only=False)
    except Exception as e:
        logging.warning("メイン再優先特記: ブックを開けませんでした: %s", e)
        return ""
    try:
        ws = None
        for name in ("メイン", "Main"):
            if name in wb.sheetnames:
                ws = wb[name]
                break
        if ws is None:
            for sn in wb.sheetnames:
                if "メイン" in sn:
                    ws = wb[sn]
                    break
        if ws is None:
            return ""
        max_r = min(ws.max_row or 0, 400)
        max_c = min(ws.max_column or 0, 40)
        if max_r < 1 or max_c < 1:
            return ""
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                cell = ws.cell(row=r, column=c)
                if not _main_sheet_cell_is_global_comment_label(cell.value):
                    continue
                below = ws.cell(row=r + 1, column=c).value
                if below is None or (isinstance(below, float) and pd.isna(below)):
                    return ""
                return str(below).strip()
        return ""
    finally:
        pass


def _global_comment_chunk_implies_factory_closure(chunk: str) -> bool:
    """
    メイン「グローバルコメント」の断片が、工場単位の休業・非稼働を意味するか（個人休みだけを誤検出しない）。
    """
    c = unicodedata.normalize("NFKC", str(chunk or ""))
    if not c.strip():
        return False
    if re.search(r"臨時\s*休業", c):
        return True
    if "休場" in c:
        return True
    if re.search(r"工場", c) and re.search(r"休|休業|休み|停止|お休み", c):
        return True
    if re.search(r"(?:全社|全館|全工場).{0,15}(?:休|休業|停止)", c):
        return True
    if re.search(r"(?:稼働|生産|ライン).{0,12}(?:停止|なし|無し)", c):
        return True
    if re.search(r"加工.{0,15}(?:しない|無し|なし|お休み)", c):
        return True
    if "休業" in c and re.search(
        r"(?:工場|全社|本社|当日|弊社|当社|全員|社全体)", c
    ):
        return True
    return False


def _extract_calendar_dates_from_text(s: str, default_year: int) -> list[date]:
    """グローバルコメント内の日付表記を date に変換（基準年は計画の基準年）。"""
    t = unicodedata.normalize("NFKC", str(s or ""))
    found: list[date] = []
    seen: set[date] = set()

    def add(y: int, mo: int, d: int) -> None:
        try:
            dd = date(y, mo, d)
        except ValueError:
            return
        if dd not in seen:
            seen.add(dd)
            found.append(dd)

    for m in re.finditer(
        r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?",
        t,
    ):
        add(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    for m in re.finditer(
        r"(\d{4})\s*[/\-\.／]\s*(\d{1,2})\s*[/\-\.／]\s*(\d{1,2})",
        t,
    ):
        add(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    for m in re.finditer(r"(\d{1,2})\s*月\s*(\d{1,2})\s*日", t):
        add(int(default_year), int(m.group(1)), int(m.group(2)))
    for m in re.finditer(
        r"(?<!\d)(\d{1,2})\s*[/／]\s*(\d{1,2})(?!\d)",
        t,
    ):
        add(int(default_year), int(m.group(1)), int(m.group(2)))
    return found


def parse_factory_closure_dates_from_global_comment(
    text: str, default_year: int
) -> set[date]:
    """
    メインシート「グローバルコメント」に、工場臨時休業などと日付が書かれている場合に
    その日を工場休み（全員非稼働・配台で加工しない）として扱う日付集合を返す。
    """
    blob = unicodedata.normalize("NFKC", str(text or "").strip())
    if not blob:
        return set()
    chunks = [c.strip() for c in re.split(r"[\n\r。;；]+", blob) if c.strip()]
    if not chunks:
        chunks = [blob]
    out: set[date] = set()
    y0 = int(default_year)
    for ch in chunks:
        if not _global_comment_chunk_implies_factory_closure(ch):
            continue
        for d in _extract_calendar_dates_from_text(ch, y0):
            out.add(d)
    if not out and _global_comment_chunk_implies_factory_closure(blob):
        for d in _extract_calendar_dates_from_text(blob, y0):
            out.add(d)
    return out


def apply_factory_closure_dates_to_attendance(
    attendance_data: dict, members: list, closure_dates: set[date]
) -> None:
    """工場休業日: 勤怠上は全員 is_working=False とし、その日は設備割付を行わない。"""
    if not closure_dates or not attendance_data:
        return
    tag = "工場休業（メイン・グローバルコメント）"
    for d in sorted(closure_dates):
        if d not in attendance_data:
            logging.warning(
                "グローバルコメントの工場休業日 %s はマスタ勤怠に行がありません。"
                " その日は計画ループに含まれない場合、配台上の効果が限定的です。",
                d,
            )
            continue
        day = attendance_data[d]
        for m in members:
            if m not in day:
                continue
            ent = day[m]
            ent["is_working"] = False
            prev = str(ent.get("reason") or "").strip()
            ent["reason"] = f"{tag} {prev}".strip() if prev else tag


def _apply_global_priority_abolish_heuristic(blob: str, coerced: dict) -> dict:
    """
    「制限撤廃」「あらゆる条件」等: 設備専有・時刻ガードまで含め配台制約を緩める（abolish_all_scheduling_limits）。
    """
    b = unicodedata.normalize("NFKC", str(blob or ""))
    strong = (
        "制限撤廃",
        "制限を撤廃",
        "すべての制限",
        "全ての制限",
        "あらゆる制限",
        "あらゆる条件",
        "すべての条件",
        "全ての条件",
        "撤廃して",
        "撤廃し",
    )
    if any(k in b for k in strong):
        out = dict(coerced)
        out["abolish_all_scheduling_limits"] = True
        out["ignore_skill_requirements"] = True
        out["ignore_need_minimum"] = True
        logging.warning(
            "メイン再優先特記: 制限撤廃キーワードを検出。設備専有・時刻ガードを含め配台上の制約を緩めます。"
        )
        return out
    return coerced


def _finalize_global_priority_override(blob: str, coerced: dict) -> dict:
    """ソロ補正の後、abolish が true ならスキル・人数も強制オン。"""
    coerced = _apply_global_priority_solo_heuristic(blob, coerced)
    coerced = _apply_global_priority_abolish_heuristic(blob, coerced)
    if coerced.get("abolish_all_scheduling_limits"):
        out = dict(coerced)
        out["ignore_skill_requirements"] = True
        out["ignore_need_minimum"] = True
        return out
    return coerced


def _apply_global_priority_solo_heuristic(blob: str, coerced: dict) -> dict:
    """
    「一人で担当」「単独」等で人数だけ緩めても、指名メンバーがスキル非該当だと配台されない。
    その場合はスキル無視を同時に立てる。
    """
    if not coerced.get("ignore_need_minimum") or coerced.get("ignore_skill_requirements"):
        return coerced
    b = unicodedata.normalize("NFKC", str(blob or ""))
    solo_kw = ("一人", "ひとり", "単独", "１人", "1人", "独自", "単身")
    if any(k in b for k in solo_kw):
        out = dict(coerced)
        out["ignore_skill_requirements"] = True
        logging.info(
            "メイン再優先特記: 単独系キーワードのため ignore_skill_requirements を補助的に true にしました。"
        )
        return out
    return coerced


def _coerce_task_preferred_operators_dict(raw_val) -> dict:
    """AI の task_preferred_operators を {依頼NO: 氏名} に正規化。"""
    out = {}
    if not isinstance(raw_val, dict):
        return out
    for k, v in raw_val.items():
        ks = str(k).strip()
        if not ks:
            continue
        if v is None or (isinstance(v, float) and pd.isna(v)):
            continue
        vs = str(v).strip()
        if vs and vs.lower() not in ("nan", "none", "null"):
            out[ks] = vs
    return out


def _coerce_global_priority_override_dict(raw) -> dict:
    """Gemini 戻りを配台用フラグに正規化。"""

    def as_bool(v):
        if v is True:
            return True
        if v is False:
            return False
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return False
        s = unicodedata.normalize("NFKC", str(v).strip()).lower()
        return s in ("true", "1", "yes", "はい", "on")

    base = {
        "ignore_skill_requirements": False,
        "ignore_need_minimum": False,
        "abolish_all_scheduling_limits": False,
        "task_preferred_operators": {},
        "interpretation_ja": "",
    }
    if not isinstance(raw, dict):
        return base
    base["ignore_skill_requirements"] = as_bool(raw.get("ignore_skill_requirements"))
    base["ignore_need_minimum"] = as_bool(raw.get("ignore_need_minimum"))
    base["abolish_all_scheduling_limits"] = as_bool(
        raw.get("abolish_all_scheduling_limits")
    )
    base["task_preferred_operators"] = _coerce_task_preferred_operators_dict(
        raw.get("task_preferred_operators")
    )
    ij = raw.get("interpretation_ja")
    if ij is not None and not (isinstance(ij, float) and pd.isna(ij)):
        base["interpretation_ja"] = str(ij).strip()[:800]
    return base


def _parse_global_priority_override_gemini_response(res):
    """Gemini 応答から JSON オブジェクト1つを取り出す（```json フェンス付きでも可）。"""
    raw = (_gemini_result_text(res) or "").strip()
    if not raw:
        return None
    candidate = None
    fence = re.search(
        r"```(?:json)?\s*(\{.*\})\s*```",
        raw,
        re.DOTALL | re.IGNORECASE,
    )
    if fence:
        candidate = fence.group(1).strip()
    elif raw.startswith("{"):
        candidate = raw
    else:
        loose = re.search(r"\{.*\}", raw, re.DOTALL)
        candidate = loose.group(0).strip() if loose else None
    if not candidate:
        return None
    try:
        parsed = json.loads(candidate)
    except json.JSONDecodeError:
        return None
    return parsed if isinstance(parsed, dict) else None


def analyze_global_priority_override_comment(
    text: str, members: list, reference_year: int, ai_sheet_sink: dict | None = None
) -> dict:
    """
    メインシート「再優先特別記載」欄の自由記述を AI で解釈し、配台で最優先適用するフラグを返す。
    - ignore_skill_requirements: True のとき skills の OP/AS 判定を無視し、当日出勤メンバー全員を OP 扱いで候補に含める。
    - ignore_need_minimum: True のとき need シート・行の必要人数・上書きより優先しチーム人数を 1 にする。
    - abolish_all_scheduling_limits: True のとき同一設備の直列ロック・原反同日10:30・指定開始時刻・マクロ実行時刻下限を適用しない
      （メンバーの勤務終了・休憩・avail は維持）。
    - task_preferred_operators: 依頼NO → 主担当OP名。メイン原文に明示されたときのみ。配台ではセル・特別指定備考AIの指名より優先。

    空文言・API 未設定・失敗時はいずれも False（task_preferred_operators は {}）。
    """
    empty = _coerce_global_priority_override_dict({})
    if not text or not str(text).strip():
        if ai_sheet_sink is not None:
            ai_sheet_sink["メイン再優先特記_AI_API"] = "スキップ（メイン原文なし）"
        return empty
    blob = str(text).strip()
    ref_y = int(reference_year) if reference_year is not None else date.today().year
    mem_sig = ",".join(sorted(str(m).strip() for m in (members or []) if m))
    cache_fingerprint = f"{GLOBAL_PRIORITY_OVERRIDE_CACHE_PREFIX}{ref_y}\n{blob}\n{mem_sig}"
    cache_key = hashlib.sha256(cache_fingerprint.encode("utf-8")).hexdigest()
    ai_cache = load_ai_cache()
    cached = get_cached_ai_result(ai_cache, cache_key, content_key=cache_fingerprint)
    if cached is not None:
        logging.info("メイン再優先特記: キャッシュヒット（Gemini は呼びません）。")
        if ai_sheet_sink is not None:
            ai_sheet_sink["メイン再優先特記_AI_API"] = "なし（キャッシュ使用）"
        return _finalize_global_priority_override(
            blob, _coerce_global_priority_override_dict(cached)
        )

    if not API_KEY:
        logging.info("GEMINI_API_KEY 未設定のためメイン再優先特記の AI 解析をスキップしました。")
        if ai_sheet_sink is not None:
            ai_sheet_sink["メイン再優先特記_AI_API"] = "なし（APIキー未設定）"
        return _finalize_global_priority_override(blob, _coerce_global_priority_override_dict({}))

    member_sample = ", ".join(str(m) for m in (members or [])[:80])
    if len(members or []) > 80:
        member_sample += " …"

    prompt = f"""あなたは工場の配台計画システム用アシスタントです。
Excel メインシートの「再優先特別記載」欄にユーザーが入力した **全文** を読み、次のキーを JSON で返してください。

【最優先ルール】
この欄はマスタ・スキル・need・タスク行・特別指定_備考の AI 指名より優先される「例外指示」です。

1) ignore_skill_requirements: スキルが無い人でも割り当てよい／誰でも担当、などのとき true
2) ignore_need_minimum: 必要人数を減らす／1人でよい、などのとき true
3) **abolish_all_scheduling_limits**: 原文に **制限を撤廃／あらゆる条件より最優先／すべての制限を無視／設備や時刻の制約も含め配台だけ最優先** など、
   「配台上のルールを可能な限り外してよい」という意味があるとき **true**。
   true のときは (1)(2) も必ず true にすること（スキル・人数は実質無効化）。
4) **特定の氏名＋一人／単独で担当** → (1)(2) は両方 true（片方だけだと配台されないことがある）。
5) **task_preferred_operators**: オブジェクト。**原文に「依頼NO（例 W3-14、Y3-26）」と、その依頼の主担当として指名する人物が対応して書かれている場合のみ** 出力する。
   - キー: 原文の依頼NO文字列と **完全一致**（ハイフン・英大文字小文字も原文どおり）
   - 値: その人物の識別名（**姓だけでよい**。登録メンバー名と照合しやすい表記。敬称「さん」等は付けない）
   - 例: 「W3-14の検査は菅沼さんが一人で担当」→ {{"W3-14":"菅沼"}}
   - 依頼NOと担当者の対応が原文に無いときは **{{}} 空オブジェクト** とする（キーを推測で増やさない）。

推測で true にしないこと。明確な根拠があるときだけ true。

【返答形式】
先頭が {{ で終わりが }} の **JSON オブジェクト1つのみ**（説明文・マークダウン禁止）。

キー:
- "ignore_skill_requirements": true または false
- "ignore_need_minimum": true または false
- "abolish_all_scheduling_limits": true または false
- "task_preferred_operators": オブジェクト（上記5のとおり。該当なしは {{}}）
- "interpretation_ja": 原文の意図を1文で（120文字以内）

【基準年】 日付言及があれば西暦 {ref_y} 年として解釈してよい。

【登録メンバー名の参考】（照合用。JSON キーには含めない）
{member_sample}

【再優先特別記載・原文】
{blob}
"""
    try:
        ppath = os.path.join(log_dir, "ai_global_priority_override_last_prompt.txt")
        with open(ppath, "w", encoding="utf-8", newline="\n") as pf:
            pf.write(prompt)
        logging.info("メイン再優先特記: プロンプト全文 → %s", ppath)
    except OSError as ex:
        logging.warning("メイン再優先特記: プロンプト保存失敗: %s", ex)

    client = genai.Client(api_key=API_KEY)
    try:
        res = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
        record_gemini_response_usage(res, GEMINI_MODEL_FLASH)
        parsed = _parse_global_priority_override_gemini_response(res)
        if parsed is None:
            logging.warning(
                "メイン再優先特記: AI 応答から JSON を解釈できませんでした。キャッシュせず、次回再試行されます。"
            )
            try:
                rpath = os.path.join(log_dir, "ai_global_priority_override_last_response.txt")
                with open(rpath, "w", encoding="utf-8", newline="\n") as rf:
                    rf.write(_gemini_result_text(res) or "")
            except OSError:
                pass
            if ai_sheet_sink is not None:
                ai_sheet_sink["メイン再優先特記_AI_API"] = "あり（JSON解釈失敗・既定値）"
            return _finalize_global_priority_override(
                blob, _coerce_global_priority_override_dict({})
            )
        coerced = _coerce_global_priority_override_dict(parsed)
        coerced = _finalize_global_priority_override(blob, coerced)
        try:
            rpath = os.path.join(log_dir, "ai_global_priority_override_last_response.txt")
            with open(rpath, "w", encoding="utf-8", newline="\n") as rf:
                rf.write(_gemini_result_text(res) or "")
        except OSError:
            pass
        put_cached_ai_result(ai_cache, cache_key, coerced, content_key=cache_fingerprint)
        save_ai_cache(ai_cache)
        _tpo = coerced.get("task_preferred_operators") or {}
        logging.info(
            "メイン再優先特記: AI 解釈 skill=%s need1=%s abolish=%s task_pref=%s件 — %s",
            coerced["ignore_skill_requirements"],
            coerced["ignore_need_minimum"],
            coerced.get("abolish_all_scheduling_limits"),
            len(_tpo),
            coerced.get("interpretation_ja", "")[:100],
        )
        if ai_sheet_sink is not None:
            ai_sheet_sink["メイン再優先特記_AI_API"] = "あり"
        return coerced
    except Exception as e:
        logging.warning("メイン再優先特記: Gemini 呼び出し失敗: %s", e)
        if ai_sheet_sink is not None:
            ai_sheet_sink["メイン再優先特記_AI_API"] = f"失敗: {e}"[:500]
        return _finalize_global_priority_override(
            blob, _coerce_global_priority_override_dict({})
        )


def default_result_task_sheet_column_order(max_history_len: int) -> list:
    """結果_タスク一覧の既定列順（履歴列数は実行時に決まる）。"""
    hist = [f"履歴{i+1}" for i in range(max_history_len)]
    return [
        "ステータス",
        "タスクID",
        "工程名",
        "機械名",
        "優先度",
        *hist,
        "必要OP(上書)",
        "タスク効率",
        "加工途中",
        "特別指定あり",
        "担当OP指名",
        "回答納期",
        "指定納期",
        "計画基準納期",
        "納期緊急",
        "加工開始日",
        "総加工量",
        "残加工量",
        "完了率(実行時点)",
        "特別指定_AI",
    ]


def _task_date_key_for_result_sheet_sort(val):
    """結果_タスク一覧の並べ替え用。欠損・解釈不能は最後（date.max）。"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return date.max
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    try:
        ts = pd.Timestamp(val)
        if pd.isna(ts):
            return date.max
        return ts.date()
    except Exception:
        return date.max


def _result_task_sheet_sort_key(t: dict):
    """
    結果_タスク一覧の表示順（配台キュー順とは独立）。
    ①加工開始日が早い ②回答納期が早い ③指定納期が早い。
    同一キー内は依頼NO（タスクID）文字列でまとめ、さらに工程名で安定化。
    """
    return (
        _task_date_key_for_result_sheet_sort(t.get("start_date_req")),
        _task_date_key_for_result_sheet_sort(t.get("answer_due_date")),
        _task_date_key_for_result_sheet_sort(t.get("specified_due_date")),
        str(t.get("task_id", "")).strip(),
        str(t.get("machine", "")).strip(),
    )


def _is_result_task_history_expand_token(cell_val) -> bool:
    """列設定シートで「履歴」1行を置くと履歴1～n をその位置に展開する。"""
    if cell_val is None or (isinstance(cell_val, float) and pd.isna(cell_val)):
        return False
    s = unicodedata.normalize("NFKC", str(cell_val).strip())
    return s in ("履歴", "履歴*")


def _result_task_column_alias_map(df_columns) -> dict:
    """見出しの NFKC 正規化キー → DataFrame 上の実列名。"""
    m = {}
    for c in df_columns:
        m[_nfkc_column_aliases(str(c).strip())] = c
    return m


def _resolve_result_task_column_label(label, col_by_norm: dict):
    if label is None or (isinstance(label, float) and pd.isna(label)):
        return None
    s = unicodedata.normalize("NFKC", str(label).strip())
    if not s or s.lower() in ("nan", "none"):
        return None
    return col_by_norm.get(_nfkc_column_aliases(s))


def _parse_column_visible_cell(val) -> bool:
    """表示列: 空・未記入は True（表示）。FALSE/0/いいえ 等で非表示。"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return True
    if isinstance(val, bool):
        return val
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        if val == 0:
            return False
        if val == 1:
            return True
    s = unicodedata.normalize("NFKC", str(val).strip()).lower()
    if s in ("", "true", "1", "はい", "yes", "on", "表示", "○"):
        return True
    if s in ("false", "flase", "0", "いいえ", "no", "off", "非表示", "隠す", "×"):
        return False
    return True


def load_result_task_column_rows_from_input_workbook(max_history_len: int) -> list | None:
    """
    TASK_INPUT_WORKBOOK の「列設定_結果_タスク一覧」シートから (列ラベル, 表示) を上から読む。
    見出し「列名」と「表示」（無い場合は表示はすべて True）。
    「履歴」「履歴*」の1行は履歴1～履歴n に展開し、同一行の表示フラグを共有する。
    """
    wb = TASKS_INPUT_WORKBOOK
    if not wb or not os.path.exists(wb):
        return None
    try:
        df_cfg = pd.read_excel(wb, sheet_name=COLUMN_CONFIG_SHEET_NAME, header=0)
    except ValueError:
        return None
    except Exception as e:
        logging.warning(
            "シート「%s」: 読み込みに失敗したため既定の列順を使います (%s)",
            COLUMN_CONFIG_SHEET_NAME,
            e,
        )
        return None
    if df_cfg is None or df_cfg.empty:
        return None

    name_col = None
    for c in df_cfg.columns:
        if _nfkc_column_aliases(str(c).strip()) == _nfkc_column_aliases(COLUMN_CONFIG_HEADER_COL):
            name_col = c
            break
    if name_col is None:
        name_col = df_cfg.columns[0]

    vis_col = None
    for c in df_cfg.columns:
        if _nfkc_column_aliases(str(c).strip()) == _nfkc_column_aliases(COLUMN_CONFIG_VISIBLE_COL):
            vis_col = c
            break

    out = []
    for i in range(len(df_cfg)):
        raw = df_cfg[name_col].iloc[i]
        vis = _parse_column_visible_cell(df_cfg[vis_col].iloc[i] if vis_col is not None else None)
        if _is_result_task_history_expand_token(raw):
            for j in range(max_history_len):
                out.append((f"履歴{j+1}", vis))
            continue
        if raw is None or (isinstance(raw, float) and pd.isna(raw)):
            continue
        s = unicodedata.normalize("NFKC", str(raw).strip())
        if not s or s.lower() in ("nan", "none"):
            continue
        out.append((s, vis))
    return out or None


def apply_result_task_sheet_column_order(df: pd.DataFrame, max_history_len: int):
    """
    列設定シートがあればその順・表示を優先し、無い列は既定順で後ろに追記（表示は True）。
    戻り値: (並べ替え後 DataFrame, 実際の列名リスト, 設定ソース説明文字列, 列名→表示bool)
    """
    default_order = default_result_task_sheet_column_order(max_history_len)
    user_rows = load_result_task_column_rows_from_input_workbook(max_history_len)
    if user_rows:
        primary = user_rows
        source = f"マクロブック「{COLUMN_CONFIG_SHEET_NAME}」"
    else:
        primary = [(n, True) for n in default_order]
        source = "既定"

    actual = list(df.columns)
    actual_set = set(actual)
    col_by_norm = _result_task_column_alias_map(actual)
    vis_map = {c: True for c in actual}

    seen = set()
    ordered = []
    unknown = []

    for item, vis in primary:
        resolved = _resolve_result_task_column_label(item, col_by_norm)
        if resolved and resolved not in seen:
            ordered.append(resolved)
            seen.add(resolved)
            vis_map[resolved] = vis
        elif not resolved:
            if item is not None and not (isinstance(item, float) and pd.isna(item)):
                lab = str(item).strip()
                if lab and lab not in unknown:
                    unknown.append(lab)

    for name in default_order:
        if name in actual_set and name not in seen:
            ordered.append(name)
            seen.add(name)
    for name in actual:
        if name not in seen:
            ordered.append(name)
            seen.add(name)

    if unknown:
        logging.warning(
            "列設定: 結果に無い列名を無視しました（最大20件）: %s",
            ", ".join(unknown[:20]) + (" …" if len(unknown) > 20 else ""),
        )
    logging.info("結果_タスク一覧の列順ソース: %s（%s 列）", source, len(ordered))
    if not user_rows:
        logging.info(
            "列順・表示のカスタマイズ: マクロ実行ブックにシート「%s」を追加。"
            " 見出し「%s」「%s」… 表示が FALSE の列は結果シートで非表示。"
            " 1行「履歴」で履歴1～n を挿入。VBA の「列設定_結果_タスク一覧_チェックボックスを配置」でチェックボックスを表示列に連動可能。",
            COLUMN_CONFIG_SHEET_NAME,
            COLUMN_CONFIG_HEADER_COL,
            COLUMN_CONFIG_VISIBLE_COL,
        )
    return df[ordered], ordered, source, vis_map


def _apply_result_task_sheet_column_visibility(worksheet, column_names: list, vis_map: dict):
    """結果_タスク一覧で、vis_map が False の列を非表示にする。"""
    for idx, col_name in enumerate(column_names, 1):
        if not vis_map.get(col_name, True):
            worksheet.column_dimensions[get_column_letter(idx)].hidden = True


_RESULT_TASK_HISTORY_RICH_HEAD_RE = re.compile(r"^・(【[^】]*】)(.*)$", re.DOTALL)


def _apply_result_task_history_rich_text(worksheet, column_names: list):
    """
    履歴列: 「・【日付】：…」の日付括弧部分を青色リッチテキストにする。
    openpyxl 3.1 未満ではスキップ（文字列の【】のみ）。
    """
    try:
        from openpyxl.cell.rich_text import CellRichText, TextBlock
        from openpyxl.cell.text import InlineFont
        from openpyxl.styles.colors import Color
    except ImportError:
        return

    hist_cols = [
        i + 1 for i, c in enumerate(column_names) if str(c).startswith("履歴")
    ]
    if not hist_cols:
        return

    plain_if = InlineFont(rFont=OUTPUT_BOOK_FONT_NAME, sz=OUTPUT_BOOK_FONT_SIZE)
    blue_if = InlineFont(
        rFont=OUTPUT_BOOK_FONT_NAME,
        sz=OUTPUT_BOOK_FONT_SIZE,
        color=Color(rgb="FF0070C0"),
    )
    top = Alignment(wrap_text=False, vertical="top")

    for r in range(2, worksheet.max_row + 1):
        for ci in hist_cols:
            cell = worksheet.cell(row=r, column=ci)
            v = cell.value
            if not isinstance(v, str) or not v.startswith("・【"):
                continue
            m = _RESULT_TASK_HISTORY_RICH_HEAD_RE.match(v)
            if not m:
                continue
            bracketed, rest = m.group(1), m.group(2)
            cell.value = CellRichText(
                TextBlock(plain_if, "・"),
                TextBlock(blue_if, bracketed),
                TextBlock(plain_if, rest),
            )
            cell.alignment = top


def _add_column_config_sheet_helpers(ws_cfg, num_data_rows: int):
    """表示列に TRUE/FALSE リスト（チェックの代わりにプルダウン）を付与。"""
    last_r = max(num_data_rows + 1, 2)
    cap = max(last_r + 50, 500)
    dv = DataValidation(type="list", formula1='"TRUE,FALSE"', allow_blank=True)
    ws_cfg.add_data_validation(dv)
    dv.add(f"B2:B{cap}")


def _coerce_actual_sheet_datetime(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, datetime):
        return val
    if isinstance(val, date) and not isinstance(val, datetime):
        return datetime.combine(val, time(0, 0))
    try:
        ts = pd.to_datetime(val, errors="coerce")
        if pd.isna(ts) or ts is pd.NaT:
            return None
        if isinstance(ts, pd.Timestamp):
            return ts.to_pydatetime()
        return ts if isinstance(ts, datetime) else None
    except Exception:
        return None


def _actual_row_time_bounds(row):
    """加工実績DATA の1行から (開始, 終了) を得る。解けなければ (None, None)。"""
    s_dt = _coerce_actual_sheet_datetime(row.get(ACT_COL_START_DT))
    e_dt = _coerce_actual_sheet_datetime(row.get(ACT_COL_END_DT))
    if s_dt and e_dt and s_dt < e_dt:
        return s_dt, e_dt
    s_dt = _coerce_actual_sheet_datetime(row.get(ACT_COL_START_ALT))
    e_dt = _coerce_actual_sheet_datetime(row.get(ACT_COL_END_ALT))
    if s_dt and e_dt and s_dt < e_dt:
        return s_dt, e_dt

    d_date = parse_optional_date(row.get(ACT_COL_DAY))
    if not d_date:
        cd = _coerce_actual_sheet_datetime(row.get(ACT_COL_DAY))
        if isinstance(cd, datetime):
            d_date = cd.date()
        elif isinstance(cd, date):
            d_date = cd
    if not d_date:
        return None, None

    ts_s = row.get(ACT_COL_TIME_START)
    ts_e = row.get(ACT_COL_TIME_END)
    if ts_s is None or pd.isna(ts_s) or ts_e is None or pd.isna(ts_e):
        return None, None

    if isinstance(ts_s, time):
        t0 = ts_s
    elif isinstance(ts_s, datetime):
        t0 = ts_s.time()
    else:
        t0 = parse_time_str(ts_s, None)

    if isinstance(ts_e, time):
        t1 = ts_e
    elif isinstance(ts_e, datetime):
        t1 = ts_e.time()
    else:
        t1 = parse_time_str(ts_e, None)

    if t0 is None or t1 is None or t0 >= t1:
        return None, None
    return datetime.combine(d_date, t0), datetime.combine(d_date, t1)


def load_machining_actuals_df():
    """
    マクロブックの「加工実績DATA」を読む（無ければ空 DataFrame）。
    Power Query 等で用意したシートを想定。
    """
    if not TASKS_INPUT_WORKBOOK or not os.path.exists(TASKS_INPUT_WORKBOOK):
        return pd.DataFrame()
    try:
        df = pd.read_excel(TASKS_INPUT_WORKBOOK, sheet_name=ACTUALS_SHEET_NAME)
    except ValueError:
        logging.info(
            f"シート「{ACTUALS_SHEET_NAME}」が無いため、ガントの実績行は出力しません。"
        )
        return pd.DataFrame()
    df.columns = df.columns.str.strip()
    df = _align_dataframe_headers_to_canonical(df, ACTUAL_HEADER_CANONICAL)
    logging.info(
        f"加工実績: '{TASKS_INPUT_WORKBOOK}' の '{ACTUALS_SHEET_NAME}' を {len(df)} 行読み込み。"
    )
    return df


def build_actual_timeline_events(df, equipment_list, sorted_dates):
    """
    実績シートの各行をガント用イベントへ変換。
    計画表示日（sorted_dates）かつ設備マスタに一致する「工程名」だけ対象。
    工程名は NFKC・空白正規化後にマスタ列名へマッピングする。
    時刻は DEFAULT_START_TIME / DEFAULT_END_TIME の枠内にクリップ。
    """
    if df is None or len(df) == 0:
        return []
    equip_lookup = _equipment_lookup_normalized_to_canonical(equipment_list)
    date_ok = set(sorted_dates)
    events = []
    bad_eq = 0
    bad_time = 0
    no_plan_overlap = 0
    mismatch_norm_samples = []

    for _, row in df.iterrows():
        tid = row.get(ACT_COL_TASK_ID)
        if tid is None or pd.isna(tid):
            continue
        tid_s = str(tid).strip()
        if not tid_s:
            continue
        proc = row.get(ACT_COL_PROCESS)
        if proc is None or pd.isna(proc):
            continue
        proc_key = _normalize_equipment_match_key(proc)
        mach = equip_lookup.get(proc_key)
        if not mach:
            bad_eq += 1
            if len(mismatch_norm_samples) < 12 and proc_key:
                if proc_key not in mismatch_norm_samples:
                    mismatch_norm_samples.append(proc_key)
            continue
        start_dt, end_dt = _actual_row_time_bounds(row)
        if not start_dt or not end_dt or start_dt >= end_dt:
            bad_time += 1
            continue
        op_val = row.get(ACT_COL_OPERATOR)
        op_s = ""
        if op_val is not None and not pd.isna(op_val):
            op_s = str(op_val).strip()

        before = len(events)
        for d in sorted_dates:
            if d not in date_ok:
                continue
            day_start = datetime.combine(d, DEFAULT_START_TIME)
            day_end = datetime.combine(d, DEFAULT_END_TIME)
            if end_dt <= day_start or start_dt >= day_end:
                continue
            s_clip = max(start_dt, day_start)
            e_clip = min(end_dt, day_end)
            if s_clip >= e_clip:
                continue
            events.append(
                {
                    "date": d,
                    "task_id": tid_s,
                    "machine": mach,
                    "op": op_s,
                    "sub": "",
                    "start_dt": s_clip,
                    "end_dt": e_clip,
                    "breaks": [],
                    "units_done": 0,
                    "already_done_units": 0,
                    "total_units": 0,
                    "eff_time_per_unit": 0.0,
                    "unit_m": 0.0,
                }
            )
        if len(events) == before:
            no_plan_overlap += 1

    if bad_eq:
        logging.warning(
            f"加工実績DATA: 工程名がマスタ設備と一致しない行を {bad_eq} 件スキップしました（空白等は正規化済み）。"
        )
        if mismatch_norm_samples:
            logging.info(
                "  不一致となった工程名の正規化後サンプル: "
                + " | ".join(mismatch_norm_samples[:12])
            )
    if bad_time:
        logging.info(
            f"加工実績DATA: 開始/終了日時が解釈できない行を {bad_time} 件スキップしました。"
        )
    if no_plan_overlap and sorted_dates:
        logging.info(
            f"加工実績DATA: 設備・日時は有効だが、計画対象日（当日以降の勤怠日×{DEFAULT_START_TIME}～{DEFAULT_END_TIME}）と重ならない行が {no_plan_overlap} 件ありました。"
        )
    if not events and len(df) > 0:
        logging.info(
            "加工実績DATA: ガント用セグメントが0件です。過去日の実績のみの場合、計画の表示日（sorted_dates）に含まれないため描画されません。"
        )
    logging.info(f"加工実績DATA からガント用セグメント {len(events)} 件を生成しました。")
    return events


TASK_SPECIAL_AI_LAST_RESPONSE_FILE = "ai_task_special_remark_last.txt"
# 勤怠備考キャッシュとキー空間を分離（同一SHA衝突を避ける）。指紋に基準年を含め日付解釈のズレを防ぐ。
TASK_SPECIAL_CACHE_KEY_PREFIX = "TASK_SPECIAL_v3|"
# メインシート「グローバルコメント」下の自由記述 → Gemini 解釈（配台の最優先オーバーライド）
GLOBAL_PRIORITY_OVERRIDE_CACHE_PREFIX = "GLOBAL_PRIO_v4|"


def _normalize_special_task_id_for_ai(val):
    """
    依頼NOをキャッシュキー・プロンプト行で一貫させる。
    Excel の数値セルは float になりがちなので 12345.0 → \"12345\" に揃える。
    文字列は NFKC（全角英数字など）で表記ゆれを吸収（同一実体の再API呼び出しを減らす）。
    """
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except TypeError:
        pass
    if isinstance(val, bool):
        return None
    if isinstance(val, int):
        return str(val)
    if isinstance(val, float):
        if math.isnan(val):
            return None
        if val.is_integer():
            return str(int(val))
        s = str(val).strip()
        if not s or s.lower() in ("nan", "none", "null"):
            return None
        return unicodedata.normalize("NFKC", s).strip() or None
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return None
    return unicodedata.normalize("NFKC", s).strip() or None


def _cell_text_task_special_remark(val):
    """
    特別指定_備考をプロンプト用に取り出す。仕様どおり **strip のみ**
    （先頭末尾の空白・Excel の偽空白を除き、文中の改行・スペースは保持。数値セルは表記を固定）。
    """
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except TypeError:
        pass
    if isinstance(val, bool):
        s = str(val)
    elif isinstance(val, float):
        if math.isnan(val):
            return ""
        # 備考列に数値だけ入っている場合の表記ゆれを減らす
        if val.is_integer():
            s = str(int(val))
        else:
            s = str(val)
    elif isinstance(val, int):
        s = str(val)
    else:
        s = str(val)
    s = s.strip()
    if s.lower() in ("nan", "none", "null"):
        return ""
    return s


def _task_special_prompt_lines(tasks_df):
    """プロンプトに載せる行リスト（ソート前）。正規化は上記ヘルパーに統一。"""
    lines = []
    for _, row in tasks_df.iterrows():
        if _plan_row_exclude_from_assignment(row):
            continue
        tid = _normalize_special_task_id_for_ai(row.get(TASK_COL_TASK_ID))
        rem = _cell_text_task_special_remark(row.get(PLAN_COL_SPECIAL_REMARK))
        if not tid or not rem:
            continue
        proc = str(row.get(TASK_COL_MACHINE, "") or "").strip()
        macn = str(row.get(TASK_COL_MACHINE_NAME, "") or "").strip()
        proc_disp = proc if proc else "（空）"
        macn_disp = macn if macn else "（空）"
        lines.append(
            f"- 依頼NO {tid} | 工程名「{proc_disp}」 | 機械名「{macn_disp}」 | 備考本文: {rem}"
        )
    return lines


def _normalize_task_special_scope_str(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return unicodedata.normalize("NFKC", str(s).strip())


def _task_special_scope_matches_row_field(row_val, restrict_val) -> bool:
    """
    restrict が無い・空なら制限なし（True）。
    非空なら Excel 側の値とあいまい一致（部分一致可）。
    """
    if restrict_val is None:
        return True
    r = _normalize_task_special_scope_str(restrict_val)
    if not r:
        return True
    v = _normalize_task_special_scope_str(row_val)
    if not v:
        return False
    if v == r:
        return True
    if r in v or v in r:
        return True
    return False


def _ai_remark_entry_applies_to_row(entry: dict, row) -> bool:
    """restrict_to_* が無いときは同一依頼NOの全行に適用。"""
    if not isinstance(entry, dict):
        return False
    rp = row.get(TASK_COL_MACHINE, "")
    rm = row.get(TASK_COL_MACHINE_NAME, "")
    if not _task_special_scope_matches_row_field(rp, entry.get("restrict_to_process_name")):
        return False
    if not _task_special_scope_matches_row_field(rm, entry.get("restrict_to_machine_name")):
        return False
    return True


def _row_matches_remark_source_row(entry: dict, row) -> bool:
    """
    JSON の process_name / machine_name が、当該 Excel 行の工程名・機械名と一致するか。
    （プロンプトで渡した「備考があった行」と対応づける。片方だけ一致でも可）
    """
    if not isinstance(entry, dict):
        return False
    rp = _normalize_task_special_scope_str(row.get(TASK_COL_MACHINE))
    rm = _normalize_task_special_scope_str(row.get(TASK_COL_MACHINE_NAME))
    ep = _normalize_task_special_scope_str(entry.get("process_name"))
    em = _normalize_task_special_scope_str(entry.get("machine_name"))
    proc_ok = (not ep) or (not rp) or ep == rp or ep in rp or rp in ep
    mac_ok = (not em) or (not rm) or em == rm or em in rm or rm in em
    return proc_ok and mac_ok


def _entry_is_global_task_special_scope(entry: dict) -> bool:
    """restrict_to_* が無い・空＝同一依頼NOの全工程行に効かせる指定。"""
    if not isinstance(entry, dict):
        return False
    a = _normalize_task_special_scope_str(entry.get("restrict_to_process_name"))
    b = _normalize_task_special_scope_str(entry.get("restrict_to_machine_name"))
    return not a and not b


def _select_ai_task_special_entry_for_tid_value(val, row):
    """1依頼NOに対する値が dict または dict の配列のどちらでも行に合う要素を返す。"""
    if val is None:
        return None
    if isinstance(val, list):
        for item in val:
            if (
                isinstance(item, dict)
                and _ai_remark_entry_applies_to_row(item, row)
                and _row_matches_remark_source_row(item, row)
            ):
                return item
        for item in val:
            if (
                isinstance(item, dict)
                and _ai_remark_entry_applies_to_row(item, row)
                and _entry_is_global_task_special_scope(item)
            ):
                return item
        for item in val:
            if isinstance(item, dict) and _ai_remark_entry_applies_to_row(item, row):
                return item
        return None
    if isinstance(val, dict):
        if _ai_remark_entry_applies_to_row(val, row):
            return val
        return None
    return None


def _ai_task_special_entry_for_row(ai_by_tid, row):
    """
    analyze_task_special_remarks の戻りから当該行のエントリを取る。
    プロンプトキーは正規化済み依頼NOなので、Excel が 12345.0 でもヒットする。
    restrict_to_process_name / restrict_to_machine_name が無い・空のときは
    同一依頼NOの工程・機械が異なる全行に同じ指示を適用する。
    """
    if not isinstance(ai_by_tid, dict) or not ai_by_tid:
        return None
    tid_norm = _normalize_special_task_id_for_ai(row.get(TASK_COL_TASK_ID))
    tid_raw = str(row.get(TASK_COL_TASK_ID, "") or "").strip()

    def try_val(v):
        return _select_ai_task_special_entry_for_tid_value(v, row)

    if tid_norm and tid_norm in ai_by_tid:
        hit = try_val(ai_by_tid[tid_norm])
        if hit is not None:
            return hit
    if tid_raw:
        for key in (tid_raw, str(tid_raw)):
            if key in ai_by_tid:
                hit = try_val(ai_by_tid[key])
                if hit is not None:
                    return hit
    if tid_norm:
        for k, v in ai_by_tid.items():
            if str(k).strip() == tid_norm:
                hit = try_val(v)
                if hit is not None:
                    return hit
    if tid_raw:
        for k, v in ai_by_tid.items():
            if str(k).strip() == tid_raw:
                hit = try_val(v)
                if hit is not None:
                    return hit
    return None


def _gemini_result_text(res):
    try:
        return (res.text or "").strip()
    except Exception:
        return ""


# 1 回の Python 実行（段階1 または 段階2）単位でリセットする
_gemini_usage_session: dict[str, dict[str, int]] = {}


def reset_gemini_usage_tracker() -> None:
    global _gemini_usage_session
    _gemini_usage_session = {}


def _gemini_cumulative_json_path() -> str:
    return os.path.join(log_dir, GEMINI_USAGE_CUMULATIVE_JSON_FILE)


def _load_gemini_cumulative_payload() -> dict:
    """log 内の累計 JSON を読む。無い・壊れていれば初期形を返す。"""
    path = _gemini_cumulative_json_path()
    default: dict = {
        "version": 1,
        "updated_at": None,
        "calls_total": 0,
        "prompt_total": 0,
        "candidates_total": 0,
        "thoughts_total": 0,
        "total_tokens_reported": 0,
        "estimated_cost_usd_total": 0.0,
        "by_model": {},
    }
    if not os.path.isfile(path):
        _gemini_buckets_ensure_structure(default)
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict) or int(data.get("version") or 0) != 1:
            _gemini_buckets_ensure_structure(default)
            return default
        data.setdefault("by_model", {})
        _gemini_buckets_ensure_structure(data)
        for k in (
            "calls_total",
            "prompt_total",
            "candidates_total",
            "thoughts_total",
            "total_tokens_reported",
        ):
            data[k] = int(data.get(k) or 0)
        data["estimated_cost_usd_total"] = float(data.get("estimated_cost_usd_total") or 0.0)
        return data
    except Exception:
        _gemini_buckets_ensure_structure(default)
        return default


def _save_gemini_cumulative_payload(data: dict) -> None:
    path = _gemini_cumulative_json_path()
    try:
        os.makedirs(log_dir, exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8", newline="\n") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    except OSError as ex:
        logging.debug("Gemini 累計 JSON の保存に失敗: %s", ex)


def _gemini_buckets_ensure_structure(data: dict) -> None:
    """累計 JSON に期間別バケット用の辞書を用意する（既存 v1 ファイルもマージ）。"""
    b = data.setdefault("buckets", {})
    if not isinstance(b, dict):
        b = {}
        data["buckets"] = b
    for sub in ("by_year", "by_month", "by_week", "by_day", "by_hour"):
        x = b.setdefault(sub, {})
        if not isinstance(x, dict):
            b[sub] = {}
    b.setdefault(
        "timezone_note",
        "period_key は PC ローカル時刻（datetime.now）で付与。他 PC との集計は混ぜないでください。",
    )


def _gemini_time_bucket_keys(dt: datetime) -> tuple[str, str, str, str, str]:
    """年・月・ISO週・日・時 のキー（文字列ソートで時系列比較しやすい形）。"""
    iy, iw, _ = dt.isocalendar()
    y = dt.strftime("%Y")
    ym = dt.strftime("%Y-%m")
    wk = f"{iy}-W{iw:02d}"
    d = dt.strftime("%Y-%m-%d")
    h = dt.strftime("%Y-%m-%dT%H")
    return y, ym, wk, d, h


def _gemini_bucket_add_one_call(
    buckets_root: dict,
    pt: int,
    ct: int,
    th: int,
    tt: int,
    inc_usd: float | None,
    *,
    when: datetime | None = None,
) -> None:
    """1 回の API 呼出しを年・月・週・日・時の各バケットに加算する。"""
    dt = when or datetime.now()
    y, ym, wk, d, h = _gemini_time_bucket_keys(dt)
    pairs = (
        ("by_year", y),
        ("by_month", ym),
        ("by_week", wk),
        ("by_day", d),
        ("by_hour", h),
    )
    for sub, pk in pairs:
        subd = buckets_root.setdefault(sub, {})
        ent = subd.setdefault(
            pk,
            {
                "calls": 0,
                "prompt": 0,
                "candidates": 0,
                "thoughts": 0,
                "total_tokens": 0,
                "estimated_cost_usd": 0.0,
            },
        )
        ent["calls"] = int(ent.get("calls") or 0) + 1
        ent["prompt"] = int(ent.get("prompt") or 0) + pt
        ent["candidates"] = int(ent.get("candidates") or 0) + ct
        ent["thoughts"] = int(ent.get("thoughts") or 0) + th
        ent["total_tokens"] = int(ent.get("total_tokens") or 0) + tt
        if inc_usd is not None:
            ent["estimated_cost_usd"] = float(
                ent.get("estimated_cost_usd") or 0.0
            ) + float(inc_usd)


def _append_gemini_cumulative_one_call(
    model_id: str, pt: int, ct: int, th: int, tt: int
) -> None:
    """1 回の API 応答分を累計 JSON に加算する（ログに単発料金は出さない）。"""
    mid = str(model_id).strip()
    data = _load_gemini_cumulative_payload()
    data["calls_total"] = int(data["calls_total"]) + 1
    data["prompt_total"] = int(data["prompt_total"]) + pt
    data["candidates_total"] = int(data["candidates_total"]) + ct
    data["thoughts_total"] = int(data["thoughts_total"]) + th
    data["total_tokens_reported"] = int(data["total_tokens_reported"]) + tt
    bm: dict = data["by_model"]
    if mid not in bm or not isinstance(bm[mid], dict):
        bm[mid] = {
            "calls": 0,
            "prompt": 0,
            "candidates": 0,
            "thoughts": 0,
            "total": 0,
            "estimated_cost_usd": 0.0,
        }
    m = bm[mid]
    m["calls"] = int(m.get("calls") or 0) + 1
    m["prompt"] = int(m.get("prompt") or 0) + pt
    m["candidates"] = int(m.get("candidates") or 0) + ct
    m["thoughts"] = int(m.get("thoughts") or 0) + th
    m["total"] = int(m.get("total") or 0) + tt
    inc_usd = _gemini_estimate_cost_usd(mid, pt, ct, th)
    if inc_usd is not None:
        m["estimated_cost_usd"] = float(m.get("estimated_cost_usd") or 0.0) + float(inc_usd)
        data["estimated_cost_usd_total"] = float(
            data.get("estimated_cost_usd_total") or 0.0
        ) + float(inc_usd)
    _gemini_buckets_ensure_structure(data)
    _gemini_bucket_add_one_call(
        data["buckets"], pt, ct, th, tt, inc_usd, when=datetime.now()
    )
    data["updated_at"] = datetime.now().isoformat(timespec="seconds")
    _save_gemini_cumulative_payload(data)


def record_gemini_response_usage(res, model_id: str) -> None:
    """generate_content の応答から usage_metadata を集計する（セッション＋累計 JSON）。"""
    global _gemini_usage_session
    if res is None or not str(model_id or "").strip():
        return
    um = getattr(res, "usage_metadata", None)
    if um is None:
        return

    def _iv(name: str) -> int:
        v = getattr(um, name, None)
        try:
            return int(v) if v is not None else 0
        except (TypeError, ValueError):
            return 0

    pt = _iv("prompt_token_count")
    ct = _iv("candidates_token_count")
    tt = _iv("total_token_count")
    th = _iv("thoughts_token_count")
    if tt <= 0 and (pt > 0 or ct > 0 or th > 0):
        tt = pt + ct + th
    mid = str(model_id).strip()
    b = _gemini_usage_session.setdefault(
        mid,
        {"prompt": 0, "candidates": 0, "total": 0, "thoughts": 0, "calls": 0},
    )
    b["prompt"] += pt
    b["candidates"] += ct
    b["total"] += tt
    b["thoughts"] += th
    b["calls"] += 1
    try:
        _append_gemini_cumulative_one_call(mid, pt, ct, th, tt)
    except Exception as ex:
        logging.debug("Gemini 累計の更新で例外（続行）: %s", ex)


def _gemini_estimate_cost_usd(
    model_id: str, prompt_tok: int, cand_tok: int, thoughts_tok: int
) -> float | None:
    m = str(model_id).strip().lower()
    rin, rout = None, None
    if "flash" in m:
        rin, rout = _GEMINI_FLASH_IN_PER_M, _GEMINI_FLASH_OUT_PER_M
    elif "pro" in m:
        # 目安（未使用モデル向けフォールバック）
        rin, rout = 1.25, 5.0
    if rin is None:
        return None
    out_equiv = cand_tok + thoughts_tok
    return (prompt_tok / 1_000_000.0) * rin + (out_equiv / 1_000_000.0) * rout


def _gemini_bucket_entry_line(period_key: str, ent: dict) -> str:
    calls = int(ent.get("calls") or 0)
    pt = int(ent.get("prompt") or 0)
    cc = int(ent.get("candidates") or 0)
    th = int(ent.get("thoughts") or 0)
    usd = float(ent.get("estimated_cost_usd") or 0.0)
    parts = [
        f"  {period_key}",
        f"呼出し {calls:,} 回",
        f"入力 {pt:,}",
        f"出力 {cc:,}",
    ]
    if th:
        parts.append(f"思考 {th:,}")
    if usd > 0:
        parts.append(f"推定USD ${usd:.6f}")
        parts.append(f"推定JPY ¥{usd * GEMINI_JPY_PER_USD:.2f}")
    return "  ".join(parts)


def _gemini_time_bucket_summary_lines(cum: dict) -> list[str]:
    b = cum.get("buckets")
    if not isinstance(b, dict):
        return []
    lines: list[str] = []
    note = b.get("timezone_note")
    lines.append("【期間別集計】（トレンドは log の CSV を Excel でグラフ化）")
    if note:
        lines.append(f"  （{note}）")

    def emit_block(title: str, sub: str, limit: int | None) -> None:
        subd = b.get(sub)
        if not isinstance(subd, dict) or not subd:
            return
        keys = sorted(subd.keys(), reverse=True)
        if limit is not None:
            keys = keys[:limit]
        lines.append(f"  [{title}]")
        for pk in keys:
            ent = subd.get(pk)
            if isinstance(ent, dict):
                lines.append(_gemini_bucket_entry_line(pk, ent))

    emit_block("年", "by_year", None)
    emit_block("月（新しい順・最大12）", "by_month", 12)
    emit_block("週 ISO（新しい順・最大8）", "by_week", 8)
    emit_block("日（新しい順・最大14）", "by_day", 14)
    emit_block("時間（新しい順・最大48）", "by_hour", 48)
    lines.append(
        f"  グラフ用: log\\{GEMINI_USAGE_BUCKETS_CSV_FILE}"
    )
    return lines


def _export_gemini_buckets_csv_for_charts(cum: dict) -> None:
    """Excel 折れ線・棒グラフ向けに長形式 CSV を log に書き出す。"""
    b = cum.get("buckets")
    if not isinstance(b, dict):
        return
    mapping = (
        ("year", "by_year"),
        ("month", "by_month"),
        ("week_iso", "by_week"),
        ("day", "by_day"),
        ("hour", "by_hour"),
    )
    rows_out: list[dict[str, object]] = []
    for gran_label, sub in mapping:
        subd = b.get(sub)
        if not isinstance(subd, dict):
            continue
        for pk in sorted(subd.keys()):
            ent = subd.get(pk)
            if not isinstance(ent, dict):
                continue
            calls = int(ent.get("calls") or 0)
            pt = int(ent.get("prompt") or 0)
            cc = int(ent.get("candidates") or 0)
            th = int(ent.get("thoughts") or 0)
            tt = int(ent.get("total_tokens") or 0)
            usd = float(ent.get("estimated_cost_usd") or 0.0)
            rows_out.append(
                {
                    "granularity": gran_label,
                    "period_key": pk,
                    "calls": calls,
                    "prompt_tokens": pt,
                    "candidates_tokens": cc,
                    "thoughts_tokens": th,
                    "total_tokens": tt,
                    "estimated_cost_usd": round(usd, 8),
                    "estimated_cost_jpy": round(usd * GEMINI_JPY_PER_USD, 4),
                }
            )
    if not rows_out:
        return
    path = os.path.join(log_dir, GEMINI_USAGE_BUCKETS_CSV_FILE)
    fieldnames = [
        "granularity",
        "period_key",
        "calls",
        "prompt_tokens",
        "candidates_tokens",
        "thoughts_tokens",
        "total_tokens",
        "estimated_cost_usd",
        "estimated_cost_jpy",
    ]
    try:
        os.makedirs(log_dir, exist_ok=True)
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            w.writeheader()
            w.writerows(rows_out)
    except OSError as ex:
        logging.debug("Gemini バケット CSV の保存に失敗: %s", ex)


def build_gemini_usage_summary_text() -> str:
    """メイン表示・結果ログ用の複数行テキスト（この実行分＋累計 JSON）。"""
    cum = _load_gemini_cumulative_payload()
    ct_tot = int(cum.get("calls_total") or 0)
    if not _gemini_usage_session and ct_tot <= 0:
        return ""

    lines: list[str] = []
    ts = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    if _gemini_usage_session:
        lines.append(f"集計時刻: {ts}（この実行での Gemini API 合計）")
        tot_calls = sum(b["calls"] for b in _gemini_usage_session.values())
        tot_p = sum(b["prompt"] for b in _gemini_usage_session.values())
        tot_c = sum(b["candidates"] for b in _gemini_usage_session.values())
        tot_th = sum(b["thoughts"] for b in _gemini_usage_session.values())
        tot_t = sum(b["total"] for b in _gemini_usage_session.values())
        lines.append(
            f"【合計】呼出し {tot_calls} 回 / 入力 {tot_p:,} / 出力 {tot_c:,}"
            + (f" / 思考 {tot_th:,}" if tot_th else "")
            + f" / total 報告 {tot_t:,}"
        )
        grand_usd = 0.0
        any_price = False
        for mid in sorted(_gemini_usage_session.keys()):
            b = _gemini_usage_session[mid]
            lines.append(f"--- モデル: {mid} ---")
            lines.append(f"  呼出し: {b['calls']} 回")
            lines.append(f"  入力トークン: {b['prompt']:,}")
            lines.append(f"  出力トークン(candidates): {b['candidates']:,}")
            if b.get("thoughts", 0):
                lines.append(f"  思考トークン(thoughts): {b['thoughts']:,}")
            lines.append(f"  total_token_count 合計: {b['total']:,}")
            est = _gemini_estimate_cost_usd(
                mid, b["prompt"], b["candidates"], b.get("thoughts", 0)
            )
            if est is not None:
                any_price = True
                grand_usd += est
                lines.append(f"  推定料金(USD): ${est:.6f}")
                lines.append(
                    f"  推定料金(JPY・{GEMINI_JPY_PER_USD:.0f}円/USD): ¥{est * GEMINI_JPY_PER_USD:.2f}"
                )
            else:
                lines.append("  推定料金: （単価未登録モデル）")
        if any_price:
            lines.append(f"【推定料金 合計(USD)】${grand_usd:.6f}")
            lines.append(
                f"【推定料金 合計(JPY)】¥{grand_usd * GEMINI_JPY_PER_USD:.2f}"
            )
    else:
        lines.append(f"集計時刻: {ts}")
        lines.append("（この実行での Gemini API 呼出しはありません）")
    lines.append("※ トークンは API の usage_metadata に基づきます。")
    lines.append(
        "※ USD 単価はコード／環境変数の目安です。実課金は Google の請求を参照してください。"
    )
    lines.append(
        "※ 各 API 呼出しごとの推定料金はコンソールに出さず、下記累計 JSON にのみ積み上げます。"
    )

    if ct_tot > 0:
        lines.append("")
        lines.append(
            f"【累計】{GEMINI_USAGE_CUMULATIVE_JSON_FILE}（全実行・推定値・ファイル: log フォルダ）"
        )
        lines.append(f"  最終更新: {cum.get('updated_at') or '—'}")
        pt0 = int(cum.get("prompt_total") or 0)
        cc0 = int(cum.get("candidates_total") or 0)
        th0 = int(cum.get("thoughts_total") or 0)
        tt0 = int(cum.get("total_tokens_reported") or 0)
        lines.append(
            f"  呼出し {ct_tot:,} 回 / 入力 {pt0:,} / 出力 {cc0:,}"
            + (f" / 思考 {th0:,}" if th0 else "")
            + f" / total 報告 {tt0:,}"
        )
        usd_all = float(cum.get("estimated_cost_usd_total") or 0.0)
        if usd_all > 0:
            lines.append(f"  推定料金累計(USD): ${usd_all:.6f}")
            lines.append(
                f"  推定料金累計(JPY・{GEMINI_JPY_PER_USD:.0f}円/USD): ¥{usd_all * GEMINI_JPY_PER_USD:.2f}"
            )
        bm = cum.get("by_model") or {}
        if isinstance(bm, dict) and bm:
            for mid in sorted(bm.keys()):
                m = bm[mid]
                if not isinstance(m, dict):
                    continue
                lines.append(f"  --- 累計 モデル: {mid} ---")
                lines.append(f"    呼出し: {int(m.get('calls') or 0):,} 回")
                lines.append(f"    入力: {int(m.get('prompt') or 0):,} / 出力: {int(m.get('candidates') or 0):,}")
                if int(m.get("thoughts") or 0):
                    lines.append(f"    思考: {int(m.get('thoughts') or 0):,}")
                mud = float(m.get("estimated_cost_usd") or 0.0)
                if mud > 0:
                    lines.append(f"    推定料金累計(USD): ${mud:.6f}")
                    lines.append(
                        f"    推定料金累計(JPY): ¥{mud * GEMINI_JPY_PER_USD:.2f}"
                    )
        lines.extend(_gemini_time_bucket_summary_lines(cum))
    return "\n".join(lines)


def _write_main_sheet_gemini_usage_via_openpyxl(
    macro_wb_path: str, text: str, log_prefix: str
) -> bool:
    """メインシート P 列 16 行目以降に Gemini サマリを openpyxl で書き save する（ブックが閉じているとき）。"""
    keep_vba = str(macro_wb_path).lower().endswith(".xlsm")
    start_r, col_p, clear_n = 16, 16, 120
    wb = None
    try:
        wb = load_workbook(
            macro_wb_path, keep_vba=keep_vba, read_only=False, data_only=False
        )
    except Exception as ex:
        logging.info(
            "%s: メイン P 列への openpyxl 書込のためブックを開けません（Excel で開きっぱなし等）: %s",
            log_prefix,
            ex,
        )
        return False
    try:
        ws_main = None
        for name in ("メイン", "Main"):
            if name in wb.sheetnames:
                ws_main = wb[name]
                break
        if ws_main is None:
            for sn in wb.sheetnames:
                if "メイン" in str(sn):
                    ws_main = wb[sn]
                    break
        if ws_main is None:
            logging.info(
                "%s: メインシートが無いため openpyxl での AI サマリをスキップしました。",
                log_prefix,
            )
            return False
        last_clear = start_r + clear_n - 1
        for i in range(clear_n):
            ws_main.cell(row=start_r + i, column=col_p, value=None)
        lines = text.split("\n") if (text or "").strip() else []
        for i, line in enumerate(lines):
            if i >= clear_n:
                break
            cell = ws_main.cell(row=start_r + i, column=col_p, value=line)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        wb.save(macro_wb_path)
        logging.info(
            "%s: メインシート P%d 以降に AI 利用サマリを openpyxl で保存しました。",
            log_prefix,
            start_r,
        )
        return True
    except Exception as ex:
        logging.warning(
            "%s: メイン AI サマリの openpyxl 保存に失敗: %s", log_prefix, ex
        )
        return False
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass


def write_main_sheet_gemini_usage_summary(wb_path: str, log_prefix: str) -> None:
    """Gemini 利用サマリを log に書き、可能なら openpyxl でメイン P 列へ保存。開いたまま保存できないときは VBA 用テキストのみ。"""
    text = build_gemini_usage_summary_text()
    path = os.path.join(log_dir, GEMINI_USAGE_SUMMARY_FOR_MAIN_FILE)
    disk_ok = False
    if wb_path and os.path.isfile(wb_path):
        try:
            disk_ok = _write_main_sheet_gemini_usage_via_openpyxl(
                wb_path, text, log_prefix
            )
        except Exception as ex:
            logging.warning("%s: AI サマリの openpyxl 書き込みで例外: %s", log_prefix, ex)
    try:
        os.makedirs(log_dir, exist_ok=True)
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            f.write(text)
    except OSError:
        pass
    try:
        cum2 = _load_gemini_cumulative_payload()
        if int(cum2.get("calls_total") or 0) > 0:
            _export_gemini_buckets_csv_for_charts(cum2)
    except Exception as ex:
        logging.debug("Gemini バケット CSV 出力で例外（続行）: %s", ex)
    if disk_ok:
        return
    if text.strip():
        logging.info(
            "%s: メイン P 列は openpyxl で保存できませんでした（ブックが Excel で開いている可能性）。"
            " %s に出力済み → マクロ「メインシート_Gemini利用サマリをP列に反映」で反映してください。",
            log_prefix,
            path,
        )
    else:
        logging.info(
            "%s: Gemini 未使用: サマリを空で %s に出力。",
            log_prefix,
            path,
        )


def _try_write_main_sheet_gemini_usage_summary(phase: str) -> None:
    try:
        write_main_sheet_gemini_usage_summary(TASKS_INPUT_WORKBOOK, phase)
    except Exception as ex:
        logging.warning(
            "%s: メインシートへの AI 利用サマリ書き込みで例外（続行）: %s", phase, ex
        )


def _log_task_special_ai_response(raw_text, parsed, extracted_json_str, prompt_text=None):
    """特別指定_備考向け Gemini のプロンプト・生テキスト・抽出JSON・パース結果を1ファイルに残す。"""
    path = os.path.join(log_dir, TASK_SPECIAL_AI_LAST_RESPONSE_FILE)
    try:
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            if prompt_text is not None and str(prompt_text).strip():
                f.write("=== Gemini へ送信したプロンプト（全文） ===\n")
                f.write(str(prompt_text).strip())
                f.write("\n\n")
            f.write("=== Gemini 返却テキスト（モデル出力そのまま） ===\n")
            f.write(raw_text or "")
            f.write(
                "\n\n=== AI が返したテキストからクライアントが切り出した JSON 文字列 ===\n"
                "（※ユーザー特別指定の解析に正規表現は使っていません。モデル応答のパース用です）\n"
            )
            f.write(extracted_json_str if extracted_json_str else "(抽出なし)")
            f.write("\n\n=== json.loads 後（依頼NOキー） ===\n")
            if isinstance(parsed, dict):
                f.write(json.dumps(parsed, ensure_ascii=False, indent=2))
            else:
                f.write("(パースできず)")
        logging.info(
            "タスク特別指定: プロンプト＋AI応答の詳細 → %s",
            path,
        )
    except OSError as ex:
        logging.warning("タスク特別指定: AI応答ファイル保存に失敗: %s", ex)
    if isinstance(parsed, dict) and parsed:
        logging.info(
            "タスク特別指定: 解析された依頼NO: %s",
            ", ".join(sorted(parsed.keys(), key=lambda x: str(x))),
        )
        for tid_k in sorted(parsed.keys(), key=lambda x: str(x)):
            logging.info(
                "  依頼NO [%s] AI解析フィールド: %s",
                tid_k,
                json.dumps(parsed[tid_k], ensure_ascii=False),
            )


def _parse_and_log_task_special_gemini_response(res, prompt_text=None):
    """
    API レスポンスを JSON 化しログ／ファイルへ記録。失敗時は None。
    ユーザーの特別指定文言には触れず、モデル出力から JSON ブロックを取り出す処理のみ。
    """
    raw = _gemini_result_text(res)
    if raw:
        stripped = raw.strip()
        if stripped.startswith("{"):
            try:
                trial = json.loads(stripped)
                if isinstance(trial, dict):
                    _log_task_special_ai_response(raw, trial, stripped, prompt_text)
                    return trial
            except json.JSONDecodeError:
                pass
    match = re.search(r"\{.*\}", raw, re.DOTALL)
    if not match:
        _log_task_special_ai_response(raw, {}, None, prompt_text)
        logging.warning(
            "タスク特別指定: AI応答から JSON を抽出できませんでした。生テキスト先頭 3000 文字:\n%s",
            (raw[:3000] if raw else "(空)"),
        )
        return None
    extracted = match.group(0)
    try:
        parsed = json.loads(extracted)
    except json.JSONDecodeError as je:
        _log_task_special_ai_response(raw, None, extracted, prompt_text)
        logging.warning("タスク特別指定: JSON パース失敗: %s", je)
        return None
    if not isinstance(parsed, dict):
        _log_task_special_ai_response(raw, None, extracted, prompt_text)
        logging.warning("タスク特別指定: トップレベルが JSON オブジェクトではありません。")
        return None
    _log_task_special_ai_response(raw, parsed, extracted, prompt_text)
    return parsed


def analyze_task_special_remarks(tasks_df, reference_year=None, ai_sheet_sink: dict | None = None):
    """
    「配台計画_タスク入力」の「特別指定_備考」を AI で構造化（セルに値がある項目は後段でセルを優先）。
    「配台不要」がオンな行はプロンプトに載せない（API 節約・当該行は配台しないため）。
    担当OP指名はプロンプトの返却契約でモデルに preferred_operator を出力させる（備考を正規表現で切り出す処理は行わない）。
    output/ai_remarks_cache.json に TTL AI_CACHE_TTL_SECONDS でキャッシュ（同一入力・同一基準年なら API を呼ばない）。
    依頼NOは数値表記・全角などを正規化してキーを安定化し、基準年は指紋に含めて日付解釈の変化とキャッシュの食い違いを防ぐ。

    戻り値の例: 依頼NO -> オブジェクト、または同一依頼NOに備考行が複数ある場合はオブジェクトの配列。
      process_name, machine_name … 当該備考セルがある行の工程名・機械名（プロンプトの行と一致）
      restrict_to_process_name, restrict_to_machine_name … 省略または空なら同一依頼NOの全工程・全機械行に適用。
      その他 required_op, speed_override, task_efficiency, priority, start_date, start_time,
      target_completion_date, ship_by_date, preferred_operator など。
    """
    lines = _task_special_prompt_lines(tasks_df)
    if not lines:
        n_rows = len(tasks_df)
        n_rem_only = 0
        n_tid_raw = 0
        for _, row in tasks_df.iterrows():
            tid = _normalize_special_task_id_for_ai(row.get(TASK_COL_TASK_ID))
            rem = _cell_text_task_special_remark(row.get(PLAN_COL_SPECIAL_REMARK))
            if tid:
                n_tid_raw += 1
            if rem:
                n_rem_only += 1
        miss_col = PLAN_COL_SPECIAL_REMARK not in tasks_df.columns
        logging.warning(
            "タスク特別指定: AI 解析対象がありません（「%s」列は%s）。"
            "总行数=%s、依頼NOのある行=%s、備考が入っている行=%s。"
            "段階2実行前にブックを保存し、本当に「%s」列へ入力しているか確認してください。",
            PLAN_COL_SPECIAL_REMARK,
            "見つかりません" if miss_col else "空の可能性があります",
            n_rows,
            n_tid_raw,
            n_rem_only,
            PLAN_COL_SPECIAL_REMARK,
        )
        if ai_sheet_sink is not None:
            ai_sheet_sink["特別指定備考_AI_API"] = "スキップ（対象行なし）"
        return {}

    blob = "\n".join(sorted(lines))
    ref_y = int(reference_year) if reference_year is not None else date.today().year
    cache_fingerprint = f"{ref_y}\n{blob}"
    cache_key_input = f"{TASK_SPECIAL_CACHE_KEY_PREFIX}{cache_fingerprint}"
    cache_key = hashlib.sha256(cache_key_input.encode("utf-8")).hexdigest()
    ai_cache = load_ai_cache()
    cached_parsed = get_cached_ai_result(
        ai_cache, cache_key, content_key=cache_fingerprint
    )
    if cached_parsed is not None:
        logging.info(
            "タスク特別指定: キャッシュヒット（%s 件・基準年=%s）。Gemini は呼びません。",
            len(lines),
            ref_y,
        )
        if ai_sheet_sink is not None:
            ai_sheet_sink["特別指定備考_AI_API"] = "なし（キャッシュ使用）"
        return cached_parsed

    logging.info(
        "タスク特別指定: キャッシュなし。Gemini で %s 件の備考を解析します（基準年=%s）。",
        len(lines),
        ref_y,
    )

    if not API_KEY:
        logging.info("GEMINI_API_KEY 未設定のためタスク特別指定のAI解析をスキップしました。")
        if ai_sheet_sink is not None:
            ai_sheet_sink["特別指定備考_AI_API"] = "なし（APIキー未設定）"
        return {}

    prompt = f"""
あなたは工場の配台計画向けに、Excel「特別指定_備考」欄への自由記述を読み、配台ロジックが使えるフィールドだけに落とし込むアシスタントです。

【最重要】
1) 【特別指定原文】の各行は、ユーザーがセルに入力した文字列を **改変・要約・断ち切りはしておらず**（先頭末尾の空白のみ除去）、そのまま渡しています。**原文の事実や意図を別の文言に置き換えないでください。**
2) あなたの応答は **1個の JSON オブジェクトのみ**（先頭が {{ 、末尾が }} ）。説明文・マークダウン・コードフェンスは禁止。

【返却JSONの契約（この節どおりに出力すること）】
■ トップレベル
- キー: 【特別指定原文】に現れる **依頼NO文字列と完全一致**（表記・ハイフン・英大文字小文字を原文どおり）。
- 値: 次のいずれか。
  (A) **JSONオブジェクト1つ** … 当該依頼NOの備考がプロンプト上 **1行だけ** のとき。
  (B) **JSON配列**（要素はオブジェクト）… 同一依頼NOで工程名・機械名が異なる備考行が **複数** あるとき。要素の順はプロンプトの行順と対応させる。

■ process_name（文字列）・machine_name（文字列）— **必須**
- 当該備考に対応するプロンプト行の **工程名「…」**・**機械名「…」** の値と **一致** させる（「（空）」のときは空文字列 ""）。
- ログ・トレース用。省略不可。

■ restrict_to_process_name（文字列）・restrict_to_machine_name（文字列）— **任意**
- **原文が「特定の工程だけ」「この機械だけ」など、適用範囲を絞っているときだけ** 出力する。
- **原文に工程名・機械名の限定が無い**（依頼全体・全行程への指示）ときは **両方とも省略** するか **空文字列 ""** とする。
- その場合、配台ロジックは **同一依頼NOの別行（例: エンボス行と分割行）にも同じ指示を適用** する。
- 絞る場合は、原文で示された識別名を入れる（Excel の工程名・機械名と照合しやすい表記）。

■ preferred_operator（文字列）— 条件付き**必須**
- **必要条件**: 当該依頼の原文を読み、「**誰がこの加工・作業の主担当（OP）として割り当てたいか**」が **意味として** 読み取れるとき。
  例: 特定の人にやってもらう／その人に任せる／担当はあの人／OPは〜／〜さん（氏名）に依頼、など。**表現の型に依存せず**、文の意味で判断する。
- **満たしたときの出力義務**: 上記の意味が成立すると判断したオブジェクトでは、**必ず** キー `preferred_operator` を含め、値は **空でない文字列** とする。併せて **process_name / machine_name は必須**（例: `{{"process_name":"…","machine_name":"…","preferred_operator":"…"}}`）。
- **値の形式**: 原文で示された **担当者の識別名を1名分**（姓・名・ニックネーム等、原文に現れた表記を維持）。末尾の敬称（さん・君・氏）のみ除去。例:「森岡さんにやってもらいます」→ `"森岡"`。
- **出力してはいけないとき**: 原文に担当者の指意が **一切ない** と判断した依頼NOでは `preferred_operator` キー自体を **省略** する（空文字列も付けない）。

■ その他フィールド（required_op, speed_override, task_efficiency, priority, start_date, start_time, target_completion_date, ship_by_date）
- 原文から **明確に** 読み取れる場合のみ出力。読み取れない数値・日付は **省略**（推測で埋めない）。

【同一依頼NO・複数工程の例】
依頼NO Y4-2 に「エンボス」と「分割」の行があり、備考が「4/5までに終わらせる」のみで工程の限定が無い場合:
- process_name / machine_name は **備考が書かれた行** の値を入れる。
- restrict_to_* は **出さないか空** にし、**エンボス行・分割行の両方** に同じ優先度・日付等が効くようにする。

【基準年（年なし日付用）】
「4/5」「4/5に出荷」のように **年が無い** 日付は原則 **西暦 {ref_y} 年** とし、YYYY-MM-DD で出力。

【フィールド一覧（型の参考）】
- process_name, machine_name: 文字列（必須。プロンプト行と一致）
- restrict_to_process_name, restrict_to_machine_name: 文字列（任意。限定なら）
- preferred_operator: 文字列（上記契約に従う）
- required_op: 正の整数
- speed_override: 正の数（m/分）
- task_efficiency: 0〜1
- priority: 整数（小さいほど先に割付）
- start_date: YYYY-MM-DD / start_time: HH:MM
- target_completion_date, ship_by_date: YYYY-MM-DD

【解釈の指針】
- 「間に合うように」「繰り上げる」→ priority を上げる（数値を下げる）。日付が文中にあれば target_completion_date または ship_by_date に入れる。
- 担当者指名は **意味理解** で preferred_operator を決める（特定のキーワード列挙に頼らない）。
- 数値・日付は推測で補わない。
- **備考が特定の工程・機械にだけ言及していない限り**、restrict_to_* は空にし、同一依頼NOの他行にも適用される形にする。

【出力直前の自己検証（必ず実行してから JSON を閉じる）】
- 【特別指定原文】の **各行** について、対応するオブジェクトに **process_name** と **machine_name** があるか。
- 同一依頼NOが複数行あるときは **配列** で各行に1オブジェクト、または適切にマージした単一オブジェクト＋restrict の運用を一貫させる。
- 「主担当OPの指意」がある行では **非空の preferred_operator** を付ける。

【出力形式の例】（依頼NO・値は実データに合わせ替えること）
{{
  "W3-14": {{
    "process_name": "検査",
    "machine_name": "ラインA",
    "preferred_operator": "森岡"
  }},
  "Y3-26": {{
    "process_name": "コーティング",
    "machine_name": "",
    "priority": 1,
    "ship_by_date": "{ref_y}-04-05",
    "target_completion_date": "{ref_y}-04-05"
  }},
  "Y4-2": {{
    "process_name": "エンボス",
    "machine_name": "E1",
    "priority": 2,
    "restrict_to_process_name": "",
    "restrict_to_machine_name": ""
  }}
}}

【特別指定原文】（Excel からそのまま。1行＝依頼NOと備考のペア）
{blob}
"""
    try:
        ppath = os.path.join(log_dir, "ai_task_special_last_prompt.txt")
        with open(ppath, "w", encoding="utf-8", newline="\n") as pf:
            pf.write(prompt)
        logging.info("タスク特別指定: 今回 Gemini に渡したプロンプト全文 → %s", ppath)
    except OSError as ex:
        logging.warning("タスク特別指定: プロンプト保存失敗: %s", ex)

    client = genai.Client(api_key=API_KEY)
    try:
        res = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
        record_gemini_response_usage(res, GEMINI_MODEL_FLASH)
        parsed = _parse_and_log_task_special_gemini_response(res, prompt_text=prompt)
        if parsed is not None:
            put_cached_ai_result(
                ai_cache, cache_key, parsed, content_key=cache_fingerprint
            )
            save_ai_cache(ai_cache)
            logging.info("タスク特別指定: AI解析が完了しました。")
            if ai_sheet_sink is not None:
                ai_sheet_sink["特別指定備考_AI_API"] = "あり"
            return parsed
        if ai_sheet_sink is not None:
            ai_sheet_sink["特別指定備考_AI_API"] = "あり（JSON解釈失敗）"
        return {}
    except Exception as e:
        err_text = str(e)
        is_quota = ("429" in err_text) or ("RESOURCE_EXHAUSTED" in err_text)
        is_unavailable = ("503" in err_text) or ("UNAVAILABLE" in err_text)
        retry_sec = extract_retry_seconds(err_text) if is_quota else None
        if is_quota and retry_sec is not None:
            wait_sec = min(max(retry_sec, 1.0), 90.0)
            logging.warning(f"タスク特別指定 AI 429。{wait_sec:.1f}秒待機して再試行します。")
            time_module.sleep(wait_sec)
            try:
                res = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
                record_gemini_response_usage(res, GEMINI_MODEL_FLASH)
                parsed = _parse_and_log_task_special_gemini_response(res, prompt_text=prompt)
                if parsed is not None:
                    put_cached_ai_result(
                        ai_cache, cache_key, parsed, content_key=cache_fingerprint
                    )
                    save_ai_cache(ai_cache)
                    if ai_sheet_sink is not None:
                        ai_sheet_sink["特別指定備考_AI_API"] = "あり（429再試行後）"
                    return parsed
            except Exception as e2:
                logging.warning(f"タスク特別指定 AI 再試行失敗: {e2}")
        elif is_unavailable:
            wait_sec = 8.0
            logging.warning(
                f"タスク特別指定 AI 503/UNAVAILABLE。{wait_sec:.1f}秒待機して再試行します。"
            )
            time_module.sleep(wait_sec)
            try:
                res = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
                record_gemini_response_usage(res, GEMINI_MODEL_FLASH)
                parsed = _parse_and_log_task_special_gemini_response(res, prompt_text=prompt)
                if parsed is not None:
                    put_cached_ai_result(
                        ai_cache, cache_key, parsed, content_key=cache_fingerprint
                    )
                    save_ai_cache(ai_cache)
                    logging.info("タスク特別指定: AI再試行で解析が完了しました。")
                    if ai_sheet_sink is not None:
                        ai_sheet_sink["特別指定備考_AI_API"] = "あり（503再試行後）"
                    return parsed
                logging.warning("タスク特別指定 AI 503再試行: JSON 抽出に失敗しました。")
            except Exception as e2:
                logging.warning(f"タスク特別指定 AI 503再試行失敗: {e2}")
        else:
            logging.warning(f"タスク特別指定 AI エラー: {e}")
        logging.warning(
            "タスク特別指定: AI解析結果を取得できなかったため、特別指定_備考の開始日/優先指示は反映されません。"
            " 必要なら『加工開始日_指定』または『指定納期_上書き』を入力してください。"
        )
        if ai_sheet_sink is not None:
            ai_sheet_sink["特別指定備考_AI_API"] = f"失敗: {e}"[:500]
        return {}


def _merge_preferred_operator_cell_and_ai(row, ai_for_tid):
    """Excel「担当OP_指定」を優先し、空なら AI の preferred_operator。"""
    ai = ai_for_tid if isinstance(ai_for_tid, dict) else {}
    v = row.get(PLAN_COL_PREFERRED_OP)
    if v is not None and not (isinstance(v, float) and pd.isna(v)):
        s = str(v).strip()
        if s and s.lower() not in ("nan", "none", "null"):
            return s
    a = ai.get("preferred_operator")
    if a is not None:
        s = str(a).strip()
        if s and s.lower() not in ("nan", "none", "null"):
            return s
    return ""


def _global_override_preferred_operator_for_task(tpref, task_id) -> str | None:
    """
    メイン「再優先特別記載」の task_preferred_operators。
    キーは依頼NO（大文字・小文字の差は無視）。
    """
    if not isinstance(tpref, dict) or not tpref:
        return None
    tid = str(task_id).strip()
    if not tid:
        return None
    tlo = tid.lower()
    for k, v in tpref.items():
        if str(k).strip().lower() != tlo:
            continue
        s = str(v).strip()
        if s and s.lower() not in ("nan", "none", "null"):
            return s
        return None
    return None


def _merge_task_row_with_ai(row, ai_for_tid):
    """セル優先で上書き値を決定。ai_for_tid は analyze_task_special_remarks の1エントリ。"""
    ai = ai_for_tid if isinstance(ai_for_tid, dict) else {}

    def first_int(cell, ai_key):
        v = parse_optional_int(row.get(cell))
        if v is not None:
            return v
        return parse_optional_int(ai.get(ai_key))

    def first_float_pos(cell, ai_key):
        v = parse_float_safe(row.get(cell), None)
        if v is not None and (not isinstance(v, float) or not pd.isna(v)) and float(v) > 0:
            return float(v)
        a = ai.get(ai_key)
        try:
            if a is not None and float(a) > 0:
                return float(a)
        except (TypeError, ValueError):
            pass
        return None

    req_op = first_int(PLAN_COL_REQUIRED_OP, "required_op")
    if req_op is not None and req_op < 1:
        req_op = None

    te = first_float_pos(PLAN_COL_TASK_EFFICIENCY, "task_efficiency")
    if te is None or te <= 0:
        te = 1.0

    pri = first_int(PLAN_COL_PRIORITY, "priority")
    if pri is None:
        pri = 999

    st_date = parse_optional_date(row.get(PLAN_COL_START_DATE_OVERRIDE))
    if st_date is None and ai.get("start_date"):
        st_date = parse_optional_date(ai.get("start_date"))

    st_time = parse_time_str(row.get(PLAN_COL_START_TIME_OVERRIDE), None)
    if st_time is None and ai.get("start_time"):
        st_time = parse_time_str(str(ai.get("start_time")), None)

    speed_ov = first_float_pos(PLAN_COL_SPEED_OVERRIDE, "speed_override")

    return req_op, speed_ov, te, pri, st_date, st_time, ai


def _plan_row_cell_nonempty(row, col_name):
    v = row.get(col_name)
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return False
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none"):
        return False
    return True


def _ai_int_for_conflict(ai, key):
    if not ai or ai.get(key) is None:
        return None
    return parse_optional_int(ai.get(key))


def _ai_float_for_conflict(ai, key):
    if not ai or ai.get(key) is None:
        return None
    try:
        f = float(ai.get(key))
        return f if f > 0 else None
    except (TypeError, ValueError):
        return None


def detect_planning_remark_ai_conflicts(row, ai_for_tid):
    """
    特別指定_備考に依る AI 解析結果と、明示セルの両方に値があり食い違う列を返す。
    備考・AIいずれか欠ける場合は空集合。
    """
    remark = str(row.get(PLAN_COL_SPECIAL_REMARK, "") or "").strip()
    if not remark or remark.lower() in ("nan", "none"):
        return set()
    ai = ai_for_tid if isinstance(ai_for_tid, dict) else {}
    if not ai:
        return set()
    out = set()

    if _plan_row_cell_nonempty(row, PLAN_COL_REQUIRED_OP):
        cv = parse_optional_int(row.get(PLAN_COL_REQUIRED_OP))
        av = _ai_int_for_conflict(ai, "required_op")
        if cv is not None and av is not None and cv != av:
            out.add(PLAN_COL_REQUIRED_OP)

    if _plan_row_cell_nonempty(row, PLAN_COL_SPEED_OVERRIDE):
        cv = parse_float_safe(row.get(PLAN_COL_SPEED_OVERRIDE), None)
        if cv is not None and cv > 0:
            av = _ai_float_for_conflict(ai, "speed_override")
            if av is not None and abs(cv - av) > 1e-5:
                out.add(PLAN_COL_SPEED_OVERRIDE)

    if _plan_row_cell_nonempty(row, PLAN_COL_TASK_EFFICIENCY):
        cv = parse_float_safe(row.get(PLAN_COL_TASK_EFFICIENCY), None)
        if cv is not None and cv > 0:
            av = _ai_float_for_conflict(ai, "task_efficiency")
            if av is not None and abs(cv - av) > 1e-5:
                out.add(PLAN_COL_TASK_EFFICIENCY)

    if _plan_row_cell_nonempty(row, PLAN_COL_PRIORITY):
        cv = parse_optional_int(row.get(PLAN_COL_PRIORITY))
        av = _ai_int_for_conflict(ai, "priority")
        if cv is not None and av is not None and cv != av:
            out.add(PLAN_COL_PRIORITY)

    if _plan_row_cell_nonempty(row, PLAN_COL_START_DATE_OVERRIDE):
        cv = parse_optional_date(row.get(PLAN_COL_START_DATE_OVERRIDE))
        av = parse_optional_date(ai.get("start_date")) if ai.get("start_date") else None
        if cv is not None and av is not None and cv != av:
            out.add(PLAN_COL_START_DATE_OVERRIDE)

    if _plan_row_cell_nonempty(row, PLAN_COL_START_TIME_OVERRIDE):
        cv = parse_time_str(row.get(PLAN_COL_START_TIME_OVERRIDE), None)
        av = parse_time_str(str(ai.get("start_time")), None) if ai.get("start_time") else None
        if cv is not None and av is not None and cv != av:
            out.add(PLAN_COL_START_TIME_OVERRIDE)

    if _plan_row_cell_nonempty(row, PLAN_COL_PREFERRED_OP):
        cv = _normalize_person_name_for_match(row.get(PLAN_COL_PREFERRED_OP))
        av = _normalize_person_name_for_match(ai.get("preferred_operator"))
        if cv and av and cv != av:
            out.add(PLAN_COL_PREFERRED_OP)

    if out:
        out.add(PLAN_COL_SPECIAL_REMARK)
    return out


def collect_planning_conflicts_by_excel_row(tasks_df, ai_by_tid):
    """Excel 行番号(1始まり・ヘッダー=1行目) -> 矛盾があった列名の集合"""
    res = {}
    for i, (_, row) in enumerate(tasks_df.iterrows()):
        if _plan_row_exclude_from_assignment(row):
            continue
        ai_one = _ai_task_special_entry_for_row(ai_by_tid, row)
        cset = detect_planning_remark_ai_conflicts(row, ai_one)
        if cset:
            res[i + 2] = cset
    return res


def apply_planning_sheet_conflict_styles(wb_path, sheet_name, num_data_rows, conflicts_by_row):
    """
    配台計画_タスク入力シートのデータ行を、矛盾列のみ赤地・白太字にする。
    事前パスでは上書き入力列を段階1と同じ薄黄色に戻し、フォントは変更しない（体裁維持）。
    AI解析列は着色しない（段階1の仕様に合わせる）。
    .xlsm は keep_vba=True で保存する。
    """
    if not wb_path or not os.path.exists(wb_path):
        return
    keep_vba = str(wb_path).lower().endswith(".xlsm")
    wb = load_workbook(wb_path, keep_vba=keep_vba)
    try:
        if sheet_name not in wb.sheetnames:
            logging.warning(f"矛盾書式: シート '{sheet_name}' が見つかりません。")
            return
        ws = wb[sheet_name]
        header_map = {}
        for col_idx in range(1, ws.max_column + 1):
            v = ws.cell(1, col_idx).value
            if v is not None:
                header_map[str(v).strip()] = col_idx

        last_row = max(2, 1 + int(num_data_rows))
        clear_fill = PatternFill(fill_type=None)
        yellow_input_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        conflict_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        conflict_font = Font(color="FFFFFF", bold=True)

        for r in range(2, last_row + 1):
            for name in PLAN_CONFLICT_STYLABLE_COLS:
                ci = header_map.get(name)
                if not ci:
                    continue
                cell = ws.cell(row=r, column=ci)
                if name == PLAN_COL_AI_PARSE:
                    cell.fill = clear_fill
                else:
                    cell.fill = yellow_input_fill
                # フォントは上書きしない（ブック既定・ユーザー設定を維持）

        for r, colnames in conflicts_by_row.items():
            if r < 2:
                continue
            for name in colnames:
                ci = header_map.get(name)
                if not ci:
                    continue
                cell = ws.cell(row=r, column=ci)
                cell.fill = conflict_fill
                cell.font = conflict_font

        try:
            wb.save(wb_path)
        except OSError as e:
            write_planning_conflict_highlight_sidecar(sheet_name, num_data_rows, conflicts_by_row)
            logging.warning(
                "配台シートへの矛盾ハイライトをファイル保存できませんでした（Excel でブックを開いたまま等）。"
                " '%s' に指示を書き出しました。マクロがシート上に直接適用します。 (%s)",
                _planning_conflict_sidecar_path(),
                e,
            )
        else:
            _remove_planning_conflict_sidecar_safe()
            if conflicts_by_row:
                logging.info(
                    f"特別指定_備考と列の矛盾: {len(conflicts_by_row)} 行を '{sheet_name}' でハイライトしました。"
                )
    finally:
        wb.close()


def _ai_planning_target_due_date(ai_dict):
    """AI JSON の完了・出荷目標日から、配台の目標日1つを決める（複数あれば最も早い日＝厳しい方）。"""
    if not isinstance(ai_dict, dict):
        return None
    dates = []
    for k in ("target_completion_date", "ship_by_date", "latest_ship_date", "due_date"):
        d = parse_optional_date(ai_dict.get(k))
        if d is not None:
            dates.append(d)
    if not dates:
        return None
    return min(dates)


# ---------------------------------------------------------------------------
# 配台用タスクキュー
#   配台計画 DataFrame 1行 → 割付アルゴリズム用 dict への変換（優先度・納期・AI 上書きを集約）
# ---------------------------------------------------------------------------
def build_task_queue_from_planning_df(
    tasks_df, run_date, req_map, ai_by_tid=None, global_priority_override=None
):
    """
    ``generate_plan`` 内で呼ばれる。完了済み・配台不要行を除き、残りを task_queue に積む。
    ai_by_tid が None のときだけ内部で analyze_task_special_remarks を実行する。
    """
    if ai_by_tid is None:
        ai_by_tid = analyze_task_special_remarks(tasks_df, reference_year=run_date.year)
    task_queue = []
    n_exclude_plan = 0

    for _, row in tasks_df.iterrows():
        if row_has_completion_keyword(row):
            continue
        if _plan_row_exclude_from_assignment(row):
            n_exclude_plan += 1
            continue

        task_id = str(row.get(TASK_COL_TASK_ID, "")).strip()
        machine = str(row.get(TASK_COL_MACHINE, "")).strip()
        machine_name = str(row.get(TASK_COL_MACHINE_NAME, "")).strip()
        qty_total = parse_float_safe(row.get(TASK_COL_QTY), 0.0)
        done_qty = calc_done_qty_equivalent_from_row(row)
        speed_raw = row.get(TASK_COL_SPEED, 1)
        product_name = row.get(TASK_COL_PRODUCT, None)
        answer_due = parse_optional_date(row.get(TASK_COL_ANSWER_DUE))
        specified_due = parse_optional_date(row.get(TASK_COL_SPECIFIED_DUE))
        specified_due_ov = parse_optional_date(row.get(PLAN_COL_SPECIFIED_DUE_OVERRIDE))
        # 納期/開始日の採用元（優先順位）:
        # 1) 指定納期_上書き（セル）
        # 2) 特別指定_備考のAI: target_completion_date / ship_by_date 等
        # 3) 特別指定_備考のAI: start_date（従来互換・目標日が無い場合のみ締めに使う）
        # 4) 回答納期
        # 5) 指定納期
        due_basis = None
        due_source = "none"
        due_source_rank = 9
        raw_input_date = parse_optional_date(row.get(TASK_COL_RAW_INPUT_DATE))

        qty = max(0.0, qty_total - done_qty)
        speed = parse_float_safe(speed_raw, 1.0)
        if speed <= 0:
            speed = 1.0

        if qty <= 0 or not machine or not task_id:
            continue

        remark_raw = str(row.get(PLAN_COL_SPECIAL_REMARK, "") or "").strip()
        has_special_remark = bool(remark_raw) and remark_raw.lower() not in ("nan", "none")
        in_progress = done_qty > 0.0
        has_done_deadline_override = False

        ai_one = _ai_task_special_entry_for_row(ai_by_tid, row)
        req_op, speed_ov, task_eff_factor, priority, start_date_ov, start_time_ov, ai_used = _merge_task_row_with_ai(
            row, ai_one
        )
        preferred_operator_raw = _merge_preferred_operator_cell_and_ai(row, ai_one)
        gpo = global_priority_override or {}
        gop_name = _global_override_preferred_operator_for_task(
            gpo.get("task_preferred_operators"), task_id
        )
        if gop_name is not None:
            preferred_operator_raw = gop_name
            logging.info(
                "メイン再優先特記: 依頼NO=%s の担当OPをグローバル指名で上書き %r（セル・特別指定備考AIより優先）",
                task_id,
                gop_name,
            )

        ai_target_due = _ai_planning_target_due_date(ai_used)
        ai_start_date = None
        if isinstance(ai_used, dict) and ai_used.get("start_date") is not None:
            ai_start_date = parse_optional_date(ai_used.get("start_date"))

        if specified_due_ov is not None:
            due_basis = specified_due_ov
            due_source = "specified_due_override"
            due_source_rank = 0
        elif ai_target_due is not None:
            due_basis = ai_target_due
            due_source = "ai_target_due"
            due_source_rank = 1
            has_done_deadline_override = True
        elif ai_start_date is not None:
            has_done_deadline_override = True
            due_basis = ai_start_date
            due_source = "ai_start_date"
            due_source_rank = 2
        elif answer_due is not None:
            due_basis = answer_due
            due_source = "answer_due"
            due_source_rank = 3
        elif specified_due is not None:
            due_basis = specified_due
            due_source = "specified_due"
            due_source_rank = 4

        if speed_ov is not None:
            speed = speed_ov
        if speed <= 0:
            speed = 1.0

        unit = infer_unit_m_from_product_name(product_name, fallback_unit=qty_total if qty_total > 0 else qty)
        try:
            unit = float(unit)
        except Exception:
            unit = qty
        if unit <= 0:
            unit = qty

        # 納期は優先順位・緊急度には使うが、開始日の下限には使わない（余力があれば前倒し開始するため）。
        if due_basis is None:
            due_urgent = False
        else:
            due_urgent = due_basis <= run_date

        # 開始日ルール:
        # 1) デフォルトは「原反投入日の1日後」（原反投入日が無い場合は run_date）
        # 2) 特別指定（セル/AI）の開始日がある場合はそれを優先
        if raw_input_date:
            effective_start_date = max(run_date, raw_input_date + timedelta(days=1))
        else:
            effective_start_date = run_date
        if start_date_ov is not None:
            effective_start_date = start_date_ov
            if raw_input_date and start_date_ov <= raw_input_date:
                logging.info(
                    "開始日上書きを優先: 依頼NO=%s 指定開始日=%s 原反投入日=%s（当日開始を許容）",
                    task_id,
                    start_date_ov,
                    raw_input_date,
                )

        same_day_raw_start_limit = (
            time(10, 30)
            if (raw_input_date and start_date_ov is None and effective_start_date == raw_input_date)
            else None
        )

        calc_time_val = qty * speed
        ai_note = ""
        if ai_used:
            try:
                ai_note = json.dumps(ai_used, ensure_ascii=False)[:500]
            except Exception:
                ai_note = str(ai_used)[:500]

        task_queue.append(
            {
                "task_id": task_id,
                "machine": machine,
                "machine_name": machine_name,
                "start_date_req": effective_start_date,
                "answer_due_date": answer_due,
                "specified_due_date": specified_due,
                "specified_due_override": specified_due_ov,
                "due_basis_date": due_basis,
                "due_source": due_source,
                "due_source_rank": due_source_rank,
                "due_urgent": due_urgent,
                "raw_input_date": raw_input_date,
                "same_day_raw_start_limit": same_day_raw_start_limit,
                "total_qty_m": int(qty_total),
                "unit_m": int(unit),
                "remaining_units": qty / unit if unit else 0,
                "base_time_per_unit": (qty / speed) / (qty / unit) if unit and speed and qty else 0,
                "assigned_history": [],
                "calc_time_value": calc_time_val,
                "required_op": req_op,
                "task_eff_factor": task_eff_factor,
                "priority": priority,
                "earliest_start_time": start_time_ov,
                "preferred_operator_raw": preferred_operator_raw,
                "task_special_ai_note": ai_note,
                "in_progress": in_progress,
                "has_special_remark": has_special_remark,
                "has_done_deadline_override": has_done_deadline_override,
                "done_qty_reported": done_qty,
            }
        )

    logging.info(
        "task_queue 構築完了: total=%s（配台不要によりスキップ %s 行）",
        len(task_queue),
        n_exclude_plan,
    )
    return task_queue


def _task_id_priority_key(task_id):
    """
    依頼NOの同条件タイブレーク用キー。
    例: Y3-24, Y3-34 のような場合はハイフン後半の数値が小さい方を優先。
    """
    s = str(task_id or "").strip()
    if not s:
        return ("", 10**9, "")
    parts = s.rsplit("-", 1)
    if len(parts) == 2:
        head = parts[0].strip()
        tail = parts[1].strip()
        if re.match(r"^\d+$", tail):
            return (head, int(tail), s)
    return (s, 10**9, s)


def _merge_plan_sheet_user_overrides(out_df):
    """
    ブック内の「配台計画_タスク入力」にユーザーが入力した上書き列を、
    段階1の抽出結果へ (依頼NO, 工程名) 単位で引き継ぐ。
    空のセルはマージしない（新規抽出側の空のまま）。
    """
    if out_df is None or out_df.empty:
        return out_df
    if not TASKS_INPUT_WORKBOOK or not os.path.exists(TASKS_INPUT_WORKBOOK):
        return out_df
    try:
        df_old = pd.read_excel(TASKS_INPUT_WORKBOOK, sheet_name=PLAN_INPUT_SHEET_NAME)
    except Exception as e:
        logging.info("段階1: 既存の配台シートを読めないため上書き継承をスキップ (%s)", e)
        return out_df
    df_old.columns = df_old.columns.str.strip()
    df_old = _rename_legacy_plan_input_ref_columns(df_old)
    df_old = _align_dataframe_headers_to_canonical(
        df_old,
        plan_input_sheet_column_order(),
    )
    if TASK_COL_TASK_ID not in df_old.columns or TASK_COL_MACHINE not in df_old.columns:
        return out_df

    lookup = {}
    for _, r in df_old.iterrows():
        tid = str(r.get(TASK_COL_TASK_ID, "") or "").strip()
        mach = str(r.get(TASK_COL_MACHINE, "") or "").strip()
        if not tid or not mach:
            continue
        key = (tid, mach)
        bucket = lookup.setdefault(key, {})
        for c in PLAN_STAGE1_MERGE_COLUMNS:
            if c not in df_old.columns or c not in out_df.columns:
                continue
            v = r.get(c)
            if v is None or (isinstance(v, float) and pd.isna(v)):
                continue
            if isinstance(v, str):
                s = v.strip()
                if not s or s.lower() in ("nan", "none"):
                    continue
            bucket[c] = v

    if not lookup:
        return out_df

    merged_rows = 0
    for i, row in out_df.iterrows():
        tid = str(row.get(TASK_COL_TASK_ID, "") or "").strip()
        mach = str(row.get(TASK_COL_MACHINE, "") or "").strip()
        bucket = lookup.get((tid, mach))
        if not bucket:
            continue
        merged_rows += 1
        for c, v in bucket.items():
            if c == PLAN_COL_EXCLUDE_FROM_ASSIGNMENT:
                v = _coerce_plan_exclude_column_value_for_storage(v)
            out_df.at[i, c] = v

    if merged_rows:
        logging.info(
            "段階1: 既存シートから上書き列を %s 行へ引き継ぎました（キー: 依頼NO+工程名）。",
            merged_rows,
        )
    return out_df


# ---------------------------------------------------------------------------
# 配台不要（2系統）
#   (A) DataFrame 上のルール … 同一依頼NO×同一機械で「分割」行に yes（手入力は上書きしない）
#   (B) マクロブック「設定_配台不要工程」… 工程+機械ごとの C/D/E 列、Gemini で D→E、
#       保存ロック時は Excel COM で A:E のみ同期するフォールバックあり
#   いずれも apply_exclude_rules_config_to_plan_df で計画 DataFrame に反映される。
# ---------------------------------------------------------------------------

def _auto_exclude_cell_empty_for_autofill(v) -> bool:
    """配台不要セルが未入力のときだけ自動で yes を書き込む。"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return True
    if isinstance(v, str):
        s = str(v).strip()
        return not s or s.lower() in ("nan", "none")
    return False


def _normalize_task_id_for_dup_grouping(raw) -> str:
    """同一依頼NOのグルーピング用（表記ゆれ・英字の大小を寄せる）。"""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return ""
    if isinstance(raw, float) and raw == int(raw):
        s = str(int(raw))
    else:
        s = unicodedata.normalize("NFKC", str(raw).strip())
    s = s.strip()
    if not s or s.lower() == "nan":
        return ""
    return s.upper()


def _process_name_is_bunkatsu_for_auto_exclude(raw) -> bool:
    """工程名が「分割」（空白除去・NFKC 後）。"""
    t = unicodedata.normalize("NFKC", str(raw or "").strip())
    t = re.sub(r"[\s　]+", "", t)
    return t == "分割"


def _apply_auto_exclude_bunkatsu_duplicate_machine(
    df: pd.DataFrame, log_prefix: str = "段階1"
) -> pd.DataFrame:
    """
    同一依頼NOが2行以上あり、かつ空でない同一機械名が2行以上あるグループでは、
    工程名が「分割」の行の「配台不要」に yes を入れる（セルが空のときのみ）。
    機械名は _normalize_equipment_match_key で重複判定。
    """
    if df is None or df.empty:
        return df
    need_cols = (TASK_COL_TASK_ID, TASK_COL_MACHINE, TASK_COL_MACHINE_NAME)
    for c in need_cols:
        if c not in df.columns:
            return df
    if PLAN_COL_EXCLUDE_FROM_ASSIGNMENT not in df.columns:
        df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = ""
    # read_excel 等で StringDtype になると数値・真偽の .at 代入で TypeError になるため object に寄せる
    df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT].astype(object)

    by_tid = defaultdict(list)
    for i in df.index:
        tid = _normalize_task_id_for_dup_grouping(df.at[i, TASK_COL_TASK_ID])
        if not tid:
            continue
        by_tid[tid].append(i)

    n_set = 0
    for _tid_key, idx_list in by_tid.items():
        if len(idx_list) < 2:
            continue
        counts = defaultdict(int)
        for i in idx_list:
            mn_key = _normalize_equipment_match_key(df.at[i, TASK_COL_MACHINE_NAME])
            if not mn_key:
                continue
            counts[mn_key] += 1
        if not any(c >= 2 for c in counts.values()):
            continue
        for i in idx_list:
            if not _process_name_is_bunkatsu_for_auto_exclude(df.at[i, TASK_COL_MACHINE]):
                continue
            if not _auto_exclude_cell_empty_for_autofill(
                df.at[i, PLAN_COL_EXCLUDE_FROM_ASSIGNMENT]
            ):
                continue
            # 列が StringDtype のとき int 代入で TypeError になるため文字列にする（_plan_row_exclude_from_assignment は yes を真とみなす）
            df.at[i, PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = "yes"
            n_set += 1

    if n_set:
        logging.info(
            "%s: 同一依頼NOかつ同一機械名が複数行あるグループで、工程名「分割」の行 %s 件に「配台不要」=yes を自動設定しました。",
            log_prefix,
            n_set,
        )
    return df


def _normalize_process_name_for_rule_match(raw) -> str:
    """工程名のルール照合（NFKC・空白除去）。"""
    t = unicodedata.normalize("NFKC", str(raw or "").strip())
    t = re.sub(r"[\s　]+", "", t)
    return t


def _exclude_rules_sheet_header_map(ws) -> dict:
    """1行目見出し → 列番号(1始まり)。
    openpyxl は新規シート直後に max_column が 0 のままのことがあり、見出しが読めず保存前に return してしまう。
    そのため最低 A～E 列は必ず走査する。
    """
    h = {}
    last_col = max(5, int(ws.max_column or 0))
    for col in range(1, last_col + 1):
        v = ws.cell(1, col).value
        if v is not None:
            h[str(v).strip()] = col
    return h


def _ensure_exclude_rules_sheet_headers_and_columns(ws, log_prefix: str) -> tuple[int, int, int, int, int]:
    """
    1行目に標準見出し（工程名・機械名・配台不要・配台不能ロジック・ロジック式）があることを保証する。
    手動で空シートだけ追加した場合は A1:E1 が空のため、ここで書き込んで列番号を返す。
    """
    headers = (
        EXCLUDE_RULE_COL_PROCESS,
        EXCLUDE_RULE_COL_MACHINE,
        EXCLUDE_RULE_COL_FLAG,
        EXCLUDE_RULE_COL_LOGIC_JA,
        EXCLUDE_RULE_COL_LOGIC_JSON,
    )
    hm = _exclude_rules_sheet_header_map(ws)
    if all(hm.get(x) for x in headers):
        return tuple(hm[x] for x in headers)
    for i, name in enumerate(headers, start=1):
        ws.cell(row=1, column=i, value=name)
    logging.info(
        "%s: 「%s」の見出しが無い／列名が一致しないため、標準の1行目（A1:E1）を設定しました。",
        log_prefix,
        EXCLUDE_RULES_SHEET_NAME,
    )
    return (1, 2, 3, 4, 5)


def _compact_exclude_rules_data_rows(
    ws,
    c_proc: int,
    c_mach: int,
    c_flag: int,
    c_d: int,
    c_e: int,
    log_prefix: str,
) -> tuple[int, int]:
    """
    2 行目以降から「空行」を除いて上に詰める（元の並びは維持、ソートしない）。
    空行: 工程名が空、または A～E 相当の5セルがすべて空白相当。
    Returns (残したデータ行数, 削除した行数).
    """
    max_r = int(ws.max_row or 1)
    if max_r < 2:
        return 0, 0

    old_body = max_r - 1
    cols = (c_proc, c_mach, c_flag, c_d, c_e)
    rows: list[tuple[str, str, object, object, object]] = []
    for r in range(2, max_r + 1):
        pv = ws.cell(row=r, column=c_proc).value
        mv = ws.cell(row=r, column=c_mach).value
        cv = ws.cell(row=r, column=c_flag).value
        dv = ws.cell(row=r, column=c_d).value
        ev = ws.cell(row=r, column=c_e).value
        all_blank = all(
            _cell_is_blank_for_rule(ws.cell(row=r, column=c).value) for c in cols
        )
        p = str(pv).strip() if pv is not None and not (isinstance(pv, float) and pd.isna(pv)) else ""
        m = str(mv).strip() if mv is not None and not (isinstance(mv, float) and pd.isna(mv)) else ""
        if all_blank or not p:
            continue
        rows.append((p, m, cv, dv, ev))

    n_skip = old_body - len(rows)

    if not rows:
        ws.delete_rows(2, old_body)
        if old_body > 0:
            logging.info(
                "%s: 「%s」は有効なデータ行が無かったため、データ行 %s 行を削除しました。",
                log_prefix,
                EXCLUDE_RULES_SHEET_NAME,
                old_body,
            )
        return 0, n_skip

    ws.delete_rows(2, old_body)
    for i, (p, m, cv, dv, ev) in enumerate(rows, start=2):
        ws.cell(row=i, column=c_proc, value=p)
        ws.cell(row=i, column=c_mach, value=m)
        ws.cell(row=i, column=c_flag, value=cv)
        ws.cell(row=i, column=c_d, value=dv)
        ws.cell(row=i, column=c_e, value=ev)

    if n_skip:
        logging.info(
            "%s: 「%s」から空行を %s 件削除し、%s 行に詰めました（並び順は維持）。",
            log_prefix,
            EXCLUDE_RULES_SHEET_NAME,
            n_skip,
            len(rows),
        )
    return len(rows), n_skip


def _cell_is_blank_for_rule(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    s = str(v).strip()
    return not s or s.lower() in ("nan", "none", "null")


def _exclude_rule_c_column_is_yes(v) -> bool:
    """C列「配台不要」がオン（この工程+機械パターンは常に配台不要）。"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            return int(v) == 1
        except (TypeError, ValueError):
            pass
    s = unicodedata.normalize("NFKC", str(v).strip()).lower()
    return s in ("yes", "true", "1", "y", "はい", "○", "〇", "●")


def _task_row_matches_exclude_rule_target(
    task_proc: str, task_mach: str, rule_proc: str, rule_mach: str
) -> bool:
    if _normalize_process_name_for_rule_match(task_proc) != _normalize_process_name_for_rule_match(
        rule_proc
    ):
        return False
    rm = str(rule_mach or "").strip()
    if not rm:
        return True
    return _normalize_equipment_match_key(task_mach) == _normalize_equipment_match_key(rm)


def _collect_process_machine_pairs_for_exclude_rules(df_src: pd.DataFrame) -> list[tuple[str, str]]:
    """加工計画DATA から、段階1と同じ抽出条件で (工程名, 機械名) の一覧（重複除く・順序維持）。"""
    out: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for _, row in df_src.iterrows():
        if row_has_completion_keyword(row):
            continue
        task_id = str(row.get(TASK_COL_TASK_ID, "") or "").strip()
        machine = str(row.get(TASK_COL_MACHINE, "") or "").strip()
        machine_name = str(row.get(TASK_COL_MACHINE_NAME, "") or "").strip()
        qty_total = parse_float_safe(row.get(TASK_COL_QTY), 0.0)
        done_qty = calc_done_qty_equivalent_from_row(row)
        qty = max(0.0, qty_total - done_qty)
        if qty <= 0 or not machine or not task_id:
            continue
        key = (
            _normalize_process_name_for_rule_match(machine),
            _normalize_equipment_match_key(machine_name),
        )
        if key in seen:
            continue
        seen.add(key)
        out.append((machine, machine_name))
    return out


def _parse_exclude_rule_json_cell(raw) -> dict | None:
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None
    s = str(raw).strip()
    if not s:
        return None
    fence = re.search(
        r"```(?:json)?\s*(\{.*\})\s*```",
        s,
        re.DOTALL | re.IGNORECASE,
    )
    if fence:
        s = fence.group(1).strip()
    try:
        obj = json.loads(s)
    except json.JSONDecodeError:
        return None
    return obj if isinstance(obj, dict) else None


def _validate_exclude_rule_parsed_dict(o: object) -> dict | None:
    """Gemini／E列から得た dict が配台不要ルールとして有効か。"""
    if not isinstance(o, dict):
        return None
    if int(o.get("version") or 0) != 1:
        return None
    mode = str(o.get("mode") or "").strip().lower()
    if mode not in ("always_exclude", "conditions"):
        return None
    return o


def _exclude_rule_de_cache_key(stripped_blob: str) -> str:
    """「配台不能ロジック」文言（正規化済み）に対する ai_remarks_cache 用キー。"""
    h = hashlib.sha256(stripped_blob.encode("utf-8")).hexdigest()
    return f"{AI_CACHE_KEY_PREFIX_EXCLUDE_RULE_DE}:{h}"


def _cache_get_exclude_rule_de_parsed(cache_obj: dict, blob: str) -> dict | None:
    s = str(blob or "").strip()
    if not s:
        return None
    data = get_cached_ai_result(
        cache_obj, _exclude_rule_de_cache_key(s), content_key=s
    )
    if not isinstance(data, dict):
        return None
    return _validate_exclude_rule_parsed_dict(data)


def _cache_put_exclude_rule_de_parsed(
    cache_obj: dict, blob: str, parsed: dict | None
) -> None:
    if parsed is None:
        return
    s = str(blob or "").strip()
    if not s:
        return
    put_cached_ai_result(
        cache_obj, _exclude_rule_de_cache_key(s), parsed, content_key=s
    )


def _exclude_rule_logic_gemini_schema_instructions() -> str:
    allowed = ", ".join(sorted(EXCLUDE_RULE_ALLOWED_COLUMNS))
    return (
        "【スキーマ version は必ず 1】\n"
        "1) 常に配台不要（説明が条件なしで外す意味）のとき:\n"
        '{"version":1,"mode":"always_exclude"}\n\n'
        "2) 列の条件で配台不要とするとき:\n"
        '{"version":1,"mode":"conditions","require_all": true または false,"conditions":[ ... ]}\n\n'
        "conditions の各要素:\n"
        "- {\"column\":\"列名\",\"op\":\"empty\"} … セルが空\n"
        "- {\"column\":\"列名\",\"op\":\"not_empty\"}\n"
        "- {\"column\":\"列名\",\"op\":\"eq\",\"value\":\"文字列\"} / ne / contains / not_contains / regex（正規表現）\n"
        "- {\"column\":\"列名\",\"op\":\"gt\"|\"gte\"|\"lt\"|\"lte\",\"value\":数値} … 数値比較（列は数として解釈）\n\n"
        f"【使用可能な列名のみ】（これ以外は使わない）:\n{allowed}\n"
    )


def _parse_exclude_rule_json_array_response(text: str) -> list | None:
    """モデル応答から JSON 配列を取り出す（```json フェンス付き可）。"""
    s = (text or "").strip()
    if not s:
        return None
    fence = re.search(
        r"```(?:json)?\s*(\[.*\])\s*```",
        s,
        re.DOTALL | re.IGNORECASE,
    )
    if fence:
        s = fence.group(1).strip()
    try:
        obj = json.loads(s)
    except json.JSONDecodeError:
        return None
    return obj if isinstance(obj, list) else None


def _row_scalar_for_exclude_rule(row, col_name: str):
    try:
        return row.get(col_name)
    except Exception:
        return None


def _evaluate_exclude_rule_one_condition(cond: dict, row) -> bool:
    if not isinstance(cond, dict):
        return False
    col = cond.get("column")
    if col not in EXCLUDE_RULE_ALLOWED_COLUMNS:
        logging.warning("配台不要ルール: 未対応の列名をスキップしました: %s", col)
        return False
    op = str(cond.get("op") or "").strip().lower()
    val = _row_scalar_for_exclude_rule(row, col)
    val_s = "" if val is None or (isinstance(val, float) and pd.isna(val)) else str(val).strip()
    val_s_lower = val_s.lower()

    if op == "empty":
        return val_s == ""
    if op == "not_empty":
        return val_s != ""

    if op in ("contains", "not_contains", "regex", "eq", "ne"):
        rhs = cond.get("value", "")
        pat = "" if rhs is None else str(rhs)
        if op == "contains":
            return pat in val_s
        if op == "not_contains":
            return pat not in val_s
        if op == "regex":
            try:
                return re.search(pat, val_s) is not None
            except re.error:
                return False
        if op == "eq":
            return val_s == pat
        if op == "ne":
            return val_s != pat

    def _num(x):
        try:
            return float(x)
        except (TypeError, ValueError):
            return None

    nv = _num(val)
    cv = _num(cond.get("value"))
    if nv is None or cv is None:
        return False
    if op == "gt":
        return nv > cv
    if op == "gte":
        return nv >= cv
    if op == "lt":
        return nv < cv
    if op == "lte":
        return nv <= cv
    return False


def evaluate_exclude_rule_json_for_row(rule: dict, row) -> bool:
    """
    E列の JSON（version=1）を評価し、当該タスク行を配台不要とすべきなら True。
    mode: always_exclude | conditions
    """
    if not isinstance(rule, dict) or int(rule.get("version") or 0) != 1:
        return False
    mode = str(rule.get("mode") or "").strip().lower()
    if mode == "always_exclude":
        return True
    if mode != "conditions":
        return False
    conds = rule.get("conditions")
    if not isinstance(conds, list) or not conds:
        return False
    require_all = bool(rule.get("require_all", True))
    checks = []
    for c in conds:
        if isinstance(c, dict) and c.get("column") in EXCLUDE_RULE_ALLOWED_COLUMNS:
            checks.append(_evaluate_exclude_rule_one_condition(c, row))
    if not checks:
        return False
    return all(checks) if require_all else any(checks)


def _ai_compile_exclude_rule_logic_to_json(natural_language: str) -> dict | None:
    """
    D列の自然言語を Gemini で JSON ルールに変換。失敗時 None。
    output/ai_remarks_cache.json に TTL 付きでキャッシュ（同一文言なら API を呼ばない）。
    """
    blob = str(natural_language or "").strip()
    if not blob:
        return None
    ai_cache = load_ai_cache()
    hit = _cache_get_exclude_rule_de_parsed(ai_cache, blob)
    if hit is not None:
        logging.info("配台不要ルール: AIキャッシュヒット（配台不能ロジック→JSON）")
        return hit
    if not API_KEY:
        return None
    schema = _exclude_rule_logic_gemini_schema_instructions()
    prompt = (
        "あなたは工場の配台システム用です。次の「配台不能の説明」を、タスク1行を判定する機械可読ルールに変換してください。\n\n"
        "【出力】先頭が { で終わりが } の JSON オブジェクト1つのみ（説明・マークダウン禁止）。\n\n"
        f"{schema}\n"
        f"【説明文】\n{blob}\n"
    )
    try:
        ppath = os.path.join(log_dir, "ai_exclude_rule_logic_last_prompt.txt")
        with open(ppath, "w", encoding="utf-8", newline="\n") as pf:
            pf.write(prompt)
        logging.info("配台不要ルール: プロンプト → %s", ppath)
    except OSError as ex:
        logging.warning("配台不要ルール: プロンプト保存失敗: %s", ex)
    try:
        client = genai.Client(api_key=API_KEY)
        res = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
        record_gemini_response_usage(res, GEMINI_MODEL_FLASH)
        raw = (_gemini_result_text(res) or "").strip()
        rpath = os.path.join(log_dir, "ai_exclude_rule_logic_last_response.txt")
        try:
            with open(rpath, "w", encoding="utf-8", newline="\n") as rf:
                rf.write(raw)
        except OSError:
            pass
        parsed = _validate_exclude_rule_parsed_dict(_parse_exclude_rule_json_cell(raw))
        if parsed:
            _cache_put_exclude_rule_de_parsed(ai_cache, blob, parsed)
            save_ai_cache(ai_cache)
        return parsed
    except Exception as e:
        logging.warning("配台不要ルール: Gemini 変換失敗: %s", e)
        return None


def _ai_compile_exclude_rule_logics_batch(blobs: list[str]) -> list[dict | None]:
    """
    複数の D 列文言を 1 回の Gemini 呼び出しで JSON 化。失敗・要素数不一致時は 1 件ずつにフォールバック。
    ai_remarks_cache.json にヒットした文言は API を呼ばない。
    """
    n = len(blobs)
    if n == 0:
        return []
    ai_cache = load_ai_cache()
    out: list[dict | None] = [None] * n
    pend_i: list[int] = []
    pend_b: list[str] = []
    for i, b in enumerate(blobs):
        s = str(b).strip()
        hit = _cache_get_exclude_rule_de_parsed(ai_cache, s) if s else None
        if hit is not None:
            out[i] = hit
        else:
            pend_i.append(i)
            pend_b.append(s)
    if not pend_b:
        logging.info(
            "配台不要ルール: AIキャッシュのみで D→E バッチ %s 件を完結（API 呼び出しなし）。",
            n,
        )
        return out
    if not API_KEY:
        return out
    m = len(pend_b)
    if m == 1:
        out[pend_i[0]] = _ai_compile_exclude_rule_logic_to_json(pend_b[0])
        return out

    schema = _exclude_rule_logic_gemini_schema_instructions()
    numbered = "\n".join(f"[{i + 1}] {str(b).strip()}" for i, b in enumerate(pend_b))
    prompt = (
        "あなたは工場の配台システム用です。以下の N 個の「配台不能の説明」を、与えた順序でそれぞれ JSON ルールに変換してください。\n\n"
        f"【出力】JSON 配列のみ。先頭が [ で終わりが ] 。要素数は必ず {m}（Markdown・説明禁止）。\n"
        f"配列の先頭要素が [1]、2 番目が [2] … に対応します。\n\n"
        f"{schema}\n"
        f"【説明文】\n{numbered}\n"
    )
    try:
        ppath = os.path.join(log_dir, "ai_exclude_rule_logic_batch_last_prompt.txt")
        with open(ppath, "w", encoding="utf-8", newline="\n") as pf:
            pf.write(prompt)
        logging.info("配台不要ルール(バッチ): プロンプト → %s", ppath)
    except OSError as ex:
        logging.warning("配台不要ルール(バッチ): プロンプト保存失敗: %s", ex)
    try:
        client = genai.Client(api_key=API_KEY)
        res = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
        record_gemini_response_usage(res, GEMINI_MODEL_FLASH)
        raw = (_gemini_result_text(res) or "").strip()
        rpath = os.path.join(log_dir, "ai_exclude_rule_logic_batch_last_response.txt")
        try:
            with open(rpath, "w", encoding="utf-8", newline="\n") as rf:
                rf.write(raw)
        except OSError:
            pass
        arr = _parse_exclude_rule_json_array_response(raw)
        if not isinstance(arr, list) or len(arr) != m:
            logging.warning(
                "配台不要ルール: バッチ応答が不正（要素数 %s、期待 %s）。1 件ずつ再試行します。",
                len(arr) if isinstance(arr, list) else None,
                m,
            )
            for j, idx in enumerate(pend_i):
                out[idx] = _ai_compile_exclude_rule_logic_to_json(pend_b[j])
            return out
        cache_dirty = False
        for j, item in enumerate(arr):
            parsed = _validate_exclude_rule_parsed_dict(item)
            out[pend_i[j]] = parsed
            if parsed:
                _cache_put_exclude_rule_de_parsed(ai_cache, pend_b[j], parsed)
                cache_dirty = True
        if cache_dirty:
            save_ai_cache(ai_cache)
        return out
    except Exception as e:
        logging.warning("配台不要ルール: バッチ Gemini 失敗、単発にフォールバック: %s", e)
        for j, idx in enumerate(pend_i):
            out[idx] = _ai_compile_exclude_rule_logic_to_json(pend_b[j])
        return out


def _log_exclude_rules_sheet_debug(
    event: str,
    log_prefix: str,
    summary: str,
    details: str = "",
    exc: BaseException | None = None,
) -> None:
    """
    「設定_配台不要工程」の保守処理のイベントログ。

    設定シート処理の成否を log/exclude_rules_sheet_debug.txt に追記し、execution_log にもタグ付きで出力する。
    event 例: START, OPEN_OK, OPEN_RETRY, OPEN_FAIL, HEADER_FIX, SYNC_ROWS, OPENPYXL_SAVE_OK, OPENPYXL_SAVE_FAIL,
    OPENPYXL_RETRY_WAIT, OPENPYXL_VBA_FALLBACK, MATRIX_TSV_WRITTEN, COM_ATTACH_OPEN_FAIL,
    E_SIDECAR_WRITTEN, E_SIDECAR_APPLIED, FALLBACK_FAIL,
    SKIP_NO_PATH, SKIP_NO_FILE, SKIP_NO_SHEET, DATA_COMPACT
    """
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [
        f"--- {ts} ---",
        f"event={event}",
        f"phase={log_prefix}",
        f"summary={summary}",
    ]
    if details:
        lines.append(f"details={details}")
    if exc is not None:
        lines.append(f"exception={type(exc).__name__}: {exc}")
        lines.append(traceback.format_exc().rstrip())
    block = "\n".join(lines) + "\n\n"
    try:
        os.makedirs(log_dir, exist_ok=True)
        with open(exclude_rules_sheet_debug_log_path, "a", encoding="utf-8", newline="\n") as df:
            df.write(block)
    except OSError as wex:
        logging.warning("exclude_rules_sheet_debug.txt へ書けません: %s", wex)

    tag = "[設定_配台不要工程]"
    msg = f"{tag} {event} | {log_prefix} | {summary}"
    if details:
        msg += f" | {details}"
    if event in (
        "OPEN_FAIL",
        "SAVE_FAIL",
        "COM_MERGE_FAIL",
        "FALLBACK_FAIL",
        "SKIP_NO_PATH",
        "SKIP_NO_FILE",
        "SKIP_NO_SHEET",
        "FATAL",
    ):
        logging.error(msg)
    elif event in (
        "OPEN_RETRY",
        "SAVE_FAIL_HINT",
        "SAVE_RETRY",
        "COM_SYNC_UNAVAILABLE",
        "COM_ATTACH_OPEN_FAIL",
        "E_SIDECAR_WRITTEN",
        "OPENPYXL_SAVE_FAIL",
        "OPENPYXL_VBA_FALLBACK",
    ):
        logging.warning(msg)
    elif event in (
        "COM_MERGE_SKIP",
        "MATRIX_TSV_WRITTEN",
        "OPENPYXL_SAVE_OK",
        "OPENPYXL_RETRY_WAIT",
    ):
        logging.info(msg)
    else:
        logging.info(msg)


def _com_workbook_matches_path(wb_com, disk_path: str) -> bool:
    """Excel COM の Workbook が disk_path と同一ファイルか（表記ゆれ・クラウド同期パスを多少吸収）。"""
    try:
        fn = str(wb_com.FullName)
    except Exception:
        return False

    def _norm(p: str) -> str:
        p = os.path.normpath(str(p).strip().replace("/", "\\"))
        return os.path.normcase(os.path.abspath(p))

    try:
        if _norm(disk_path) == _norm(fn):
            return True
    except Exception:
        pass
    try:
        return os.path.samefile(disk_path, fn)
    except Exception:
        pass
    # 8.3 短い名前 / ロングパス差
    try:
        import win32api  # type: ignore

        a = _norm(win32api.GetLongPathName(disk_path))
        b = _norm(win32api.GetLongPathName(fn))
        if a == b:
            return True
    except Exception:
        pass
    # 同一フォルダ + 同一ファイル名（ドライブレター大小など）
    try:
        if os.path.basename(_norm(disk_path)).lower() == os.path.basename(_norm(fn)).lower():
            if _norm(os.path.dirname(disk_path)) == _norm(os.path.dirname(fn)):
                return True
    except Exception:
        pass
    return False


def _try_attach_open_workbook_via_getobject(abs_path: str):
    """
    既に Excel で開かれているブックを GetObject(フルパス) で取得する。
    Excel が起動していないと Excel を起動しうるため、呼び出し側は「Application 取得済み」のときだけ使う。
    戻り値: (Application, Workbook) または (None, None)
    """
    try:
        from win32com.client import GetObject  # type: ignore[import-not-found]
    except ImportError:
        return None, None
    try:
        wb = GetObject(abs_path)
    except Exception:
        return None, None
    try:
        xl = wb.Application
        return xl, wb
    except Exception:
        return None, None


# 設定シートの列範囲（A〜E）。VBA 行列 TSV 出力でも使用。
EXCLUDE_RULES_SHEET_COM_SYNC_MAX_COL = 5
EXCLUDE_RULES_MATRIX_CLIP_MAX_COL = 5


def _find_com_workbook_by_path(xl_app, disk_path: str):
    """起動中 Excel の Workbooks から disk_path に一致するブックを返す。無ければ None。"""
    try:
        n = int(xl_app.Workbooks.Count)
    except Exception:
        n = 0
    for i in range(1, n + 1):
        try:
            cand = xl_app.Workbooks(i)
        except Exception:
            continue
        if _com_workbook_matches_path(cand, disk_path):
            return cand
    return None


def _com_release_workbook_after_mutation(xl_app, wb, info: dict, mutation_ok: bool) -> None:
    """専用起動した Excel は終了する。実行中 Excel でだけ Open したブックは失敗時のみ閉じる。"""
    mode = info.get("mode", "keep")
    opened_here = bool(info.get("opened_wb_here"))
    if mode == "quit_excel":
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            xl_app.Quit()
        except Exception:
            pass
        return
    if opened_here and not mutation_ok:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass


def _com_attach_open_macro_workbook(
    macro_wb_path: str, log_prefix: str
) -> tuple[object, object, dict] | None:
    """
    マクロブックを Excel COM で取得する。
    戻り値: (Application, Workbook, release_info) / 失敗時 None。
    release_info: mode が keep または quit_excel、opened_wb_here が bool。
    """
    try:
        from win32com.client import Dispatch, GetActiveObject, GetObject  # type: ignore[import-not-found]
    except ImportError:
        return None

    abs_path = os.path.abspath(macro_wb_path)

    def _attach_on_running_xl(xl_app) -> tuple[object, object, dict] | None:
        macro_wb = _find_com_workbook_by_path(xl_app, macro_wb_path)
        if macro_wb is not None:
            return xl_app, macro_wb, {"mode": "keep", "opened_wb_here": False}
        try:
            macro_wb = xl_app.Workbooks.Open(abs_path, 0, False)
            return xl_app, macro_wb, {"mode": "keep", "opened_wb_here": True}
        except Exception:
            pass
        xl_go, wb_go = _try_attach_open_workbook_via_getobject(abs_path)
        if wb_go is not None and xl_go is not None:
            return xl_go, wb_go, {"mode": "keep", "opened_wb_here": False}
        return None

    xl_run = None
    try:
        xl_run = GetActiveObject("Excel.Application")
    except Exception:
        pass

    if xl_run is not None:
        got = _attach_on_running_xl(xl_run)
        if got is not None:
            return got

    try:
        xl2 = GetObject(Class="Excel.Application")
        got = _attach_on_running_xl(xl2)
        if got is not None:
            return got
    except Exception:
        pass

    try:
        xl_own = Dispatch("Excel.Application")
        xl_own.Visible = False
        xl_own.DisplayAlerts = False
        wb_own = xl_own.Workbooks.Open(abs_path, 0, False)
        return xl_own, wb_own, {"mode": "quit_excel", "opened_wb_here": True}
    except Exception as ex:
        _log_exclude_rules_sheet_debug(
            "COM_ATTACH_OPEN_FAIL",
            log_prefix,
            "Excel COM でブックを開けませんでした。",
            details=f"path={abs_path}",
            exc=ex,
        )
        return None


def _persist_exclude_rules_workbook(_wb, wb_path: str, ws, log_prefix: str) -> bool:
    """
    設定シートのディスク反映。ブックが他プロセスで開かれていなければ openpyxl で save。
    ロック等で保存できないときは log に行列 TSV を出し、VBA「設定_配台不要工程_AからE_TSVから反映」で反映する。

    _wb … 編集済み openpyxl ブック（save に使用）。
    """
    global _exclude_rules_effective_read_path

    def _openpyxl_persist_ok(which: str) -> bool:
        try:
            _wb.save(wb_path)
        except Exception as ex:
            _log_exclude_rules_sheet_debug(
                "OPENPYXL_SAVE_FAIL",
                log_prefix,
                f"openpyxl での .xlsm 保存に失敗しました {which}（Excel で開きっぱなし・ロックの可能性）。",
                details=f"path={wb_path}",
                exc=ex,
            )
            return False
        _exclude_rules_effective_read_path = wb_path
        _clear_exclude_rules_e_apply_files()
        _log_exclude_rules_sheet_debug(
            "OPENPYXL_SAVE_OK",
            log_prefix,
            "openpyxl で設定シートを含むブックを保存しました（A〜E）。",
            details=f"path={wb_path} {which}",
        )
        logging.info(
            "%s: 設定シートを openpyxl でマクロブックに保存しました。%s",
            log_prefix,
            which,
        )
        return True

    logging.info(
        "%s: 設定_配台不要工程は openpyxl で保存します（不可のときは VBA 用行列 TSV）。",
        log_prefix,
    )
    labels = ("(1/4)", "(2/4)", "(3/4)", "(4/4)")
    for i, label in enumerate(labels):
        if i:
            _log_exclude_rules_sheet_debug(
                "OPENPYXL_RETRY_WAIT",
                log_prefix,
                f"openpyxl 再保存まで 2 秒待ちます {label}。",
                details=f"path={wb_path}",
            )
            time_module.sleep(2.0)
        if _openpyxl_persist_ok(label):
            return True

    if _write_exclude_rules_matrix_vba_tsv(wb_path, ws, log_prefix):
        logging.warning(
            "%s: 設定シートを log\\%s に出力しました。"
            " Excel でマクロ「設定_配台不要工程_AからE_TSVから反映」を実行してください。",
            log_prefix,
            EXCLUDE_RULES_MATRIX_VBA_FILENAME,
        )

    _log_exclude_rules_sheet_debug(
        "OPENPYXL_VBA_FALLBACK",
        log_prefix,
        "openpyxl 保存に失敗したため VBA 用行列 TSV を出力しました（ブックは Excel 上で手動反映が必要な場合があります）。",
        details=f"path={wb_path}",
    )
    return False


def _exclude_rules_e_sidecar_path() -> str:
    return os.path.join(log_dir, EXCLUDE_RULES_E_SIDECAR_FILENAME)


def _exclude_rules_e_vba_tsv_path() -> str:
    return os.path.join(log_dir, EXCLUDE_RULES_E_VBA_TSV_FILENAME)


def _exclude_rules_matrix_vba_path() -> str:
    return os.path.join(log_dir, EXCLUDE_RULES_MATRIX_VBA_FILENAME)


def _serialize_cell_for_matrix_tsv(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, bool):
        return "TRUE" if val else "FALSE"
    return str(val)


def _write_exclude_rules_matrix_vba_tsv(
    wb_path: str, ws, log_prefix: str
) -> bool:
    """VBA 用: 設定シート 1 行目〜 max_row の A〜E を Base64(UTF-8) 付き TSV で出力する。"""
    max_r = max(1, int(ws.max_row or 1))
    lines = [
        "v1",
        "workbook\t" + os.path.abspath(wb_path),
        "sheet\t" + EXCLUDE_RULES_SHEET_NAME,
        "ncols\t5",
        "---",
    ]
    for r in range(1, max_r + 1):
        parts: list[str] = [str(r)]
        for c in range(1, 6):
            s = _serialize_cell_for_matrix_tsv(ws.cell(row=r, column=c).value)
            parts.append(base64.b64encode(s.encode("utf-8")).decode("ascii"))
        lines.append("\t".join(parts))
    path = _exclude_rules_matrix_vba_path()
    try:
        os.makedirs(log_dir, exist_ok=True)
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            f.write("\n".join(lines) + "\n")
        _log_exclude_rules_sheet_debug(
            "MATRIX_TSV_WRITTEN",
            log_prefix,
            "設定シート A〜E を VBA 反映用 TSV に書き出しました（openpyxl 保存不可時）。",
            details=f"path={path} rows={max_r}",
        )
        return True
    except OSError as ex:
        logging.warning("%s: 行列 VBA 用 TSV を書けません: %s", log_prefix, ex)
        return False


def _build_exclude_rules_list_from_openpyxl_ws(
    ws, c_proc: int, c_mach: int, c_flag: int, c_e: int
) -> list[dict]:
    """openpyxl 上の設定シートから _load_exclude_rules_from_workbook と同形のリストを構築。"""
    rules: list[dict] = []
    max_r = int(ws.max_row or 1)
    for r in range(2, max_r + 1):
        pv = ws.cell(row=r, column=c_proc).value
        proc = (
            ""
            if pv is None or (isinstance(pv, float) and pd.isna(pv))
            else str(pv).strip()
        )
        if not proc:
            continue
        mv = ws.cell(row=r, column=c_mach).value
        mach = (
            ""
            if mv is None or (isinstance(mv, float) and pd.isna(mv))
            else str(mv).strip()
        )
        cv = ws.cell(row=r, column=c_flag).value
        ev = ws.cell(row=r, column=c_e).value
        parsed = _parse_exclude_rule_json_cell(ev)
        rules.append(
            {"proc": proc, "mach": mach, "c_val": cv, "parsed": parsed}
        )
    return rules


def _set_exclude_rules_snapshot_from_ws(
    wb_path: str, ws, c_proc: int, c_mach: int, c_flag: int, c_e: int
) -> None:
    global _exclude_rules_rules_snapshot, _exclude_rules_snapshot_wb
    _exclude_rules_rules_snapshot = _build_exclude_rules_list_from_openpyxl_ws(
        ws, c_proc, c_mach, c_flag, c_e
    )
    _exclude_rules_snapshot_wb = os.path.normcase(os.path.abspath(wb_path))


def _clear_exclude_rules_e_apply_files() -> None:
    for rel in (
        EXCLUDE_RULES_E_SIDECAR_FILENAME,
        EXCLUDE_RULES_E_VBA_TSV_FILENAME,
        EXCLUDE_RULES_MATRIX_VBA_FILENAME,
    ):
        p = os.path.join(log_dir, rel)
        if os.path.isfile(p):
            try:
                os.remove(p)
            except OSError:
                pass


def _write_exclude_rules_e_vba_tsv_from_cells(
    wb_path: str, c_e: int, cells: dict[str, str], log_prefix: str
) -> None:
    """VBA 用: 行番号と Base64(UTF-8) セル文字列の TSV。"""
    lines = [
        "v1",
        "workbook\t" + os.path.abspath(wb_path),
        "sheet\t" + EXCLUDE_RULES_SHEET_NAME,
        "column_e\t" + str(int(c_e)),
        "---",
    ]
    for rk in sorted(cells.keys(), key=lambda x: int(x)):
        s = cells[rk]
        b64 = base64.b64encode(s.encode("utf-8")).decode("ascii")
        lines.append(rk + "\t" + b64)
    path_tsv = _exclude_rules_e_vba_tsv_path()
    try:
        os.makedirs(log_dir, exist_ok=True)
        with open(path_tsv, "w", encoding="utf-8", newline="\n") as f:
            f.write("\n".join(lines) + "\n")
        _log_exclude_rules_sheet_debug(
            "E_VBA_TSV_WRITTEN",
            log_prefix,
            "E 列を VBA 反映用 TSV に書き出しました（COM 保存失敗時のフォールバック用）。",
            details=f"path={path_tsv} cells={len(cells)}",
        )
    except OSError as ex:
        logging.warning("%s: E 列 VBA 用 TSV を書けません: %s", log_prefix, ex)


def _write_exclude_rules_e_apply_artifacts(
    wb_path: str, ws, c_e: int, log_prefix: str
) -> None:
    """
    E 列（非空）を JSON サイドカードと VBA 用 TSV に書く。空なら両ファイルを削除。
    Python 次回起動時の E 復元用 JSON と、マクロからの E 書込み用 TSV。
    """
    cells: dict[str, str] = {}
    max_r = int(ws.max_row or 1)
    for r in range(2, max_r + 1):
        ev = ws.cell(row=r, column=c_e).value
        if _cell_is_blank_for_rule(ev):
            continue
        s = str(ev).strip() if ev is not None else ""
        if not s:
            continue
        cells[str(r)] = s
    if not cells:
        _clear_exclude_rules_e_apply_files()
        return
    payload = {
        "version": 1,
        "workbook": os.path.abspath(wb_path),
        "sheet": EXCLUDE_RULES_SHEET_NAME,
        "column_e": c_e,
        "cells": cells,
    }
    path_sc = _exclude_rules_e_sidecar_path()
    try:
        os.makedirs(log_dir, exist_ok=True)
        with open(path_sc, "w", encoding="utf-8", newline="\n") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except OSError as ex:
        logging.warning("%s: E 列 JSON を書けません: %s", log_prefix, ex)
    _write_exclude_rules_e_vba_tsv_from_cells(wb_path, c_e, cells, log_prefix)
    _log_exclude_rules_sheet_debug(
        "E_APPLY_FILES_WRITTEN",
        log_prefix,
        "E 列を JSON と VBA 用 TSV に書き出しました（マクロで E 列を反映後、ファイル削除）。",
        details=f"cells={len(cells)}",
    )


def _try_apply_pending_exclude_rules_e_column(
    wb_path: str, ws, c_e: int, log_prefix: str
) -> int:
    """
    前回保存に失敗したとき書き出した JSON から E 列を復元する。
    ブックパスが一致しなければ何もしない。適用後はサイドカードを削除する。
    """
    path_sc = _exclude_rules_e_sidecar_path()
    if not os.path.isfile(path_sc):
        return 0
    try:
        with open(path_sc, "r", encoding="utf-8") as f:
            payload = json.load(f)
    except Exception:
        return 0
    if int(payload.get("version") or 0) != 1:
        return 0
    target = os.path.normcase(os.path.abspath(wb_path))
    if os.path.normcase(str(payload.get("workbook") or "")) != target:
        return 0
    if str(payload.get("sheet") or "") != EXCLUDE_RULES_SHEET_NAME:
        return 0
    cells = payload.get("cells")
    if not isinstance(cells, dict):
        return 0
    n = 0
    for rk, val in cells.items():
        try:
            ri = int(rk)
        except (TypeError, ValueError):
            continue
        if ri < 2:
            continue
        if isinstance(val, dict):
            sval = json.dumps(val, ensure_ascii=False)
        else:
            sval = "" if val is None else str(val).strip()
        if not sval:
            continue
        ws.cell(row=ri, column=c_e, value=sval)
        n += 1
    try:
        os.remove(path_sc)
    except OSError:
        pass
    if n:
        _log_exclude_rules_sheet_debug(
            "E_SIDECAR_APPLIED",
            log_prefix,
            f"未保存だった E 列をサイドカードから {n} セル復元しました。",
            details=path_sc,
        )
        logging.info(
            "%s: %s の内容をシートのロジック式列へ適用しました（続けて保存を試みます）。",
            log_prefix,
            path_sc,
        )
    return n


def _read_exclude_rules_d_cells_data_only_for_rows(
    wb_path: str, rows: list[int], c_d: int
) -> dict[int, object]:
    """
    D 列が数式のとき、openpyxl の通常読込では '=...' しか取れない。
    data_only=True でキャッシュ値を読む（Excel が一度でも保存・計算済みのブックで有効）。
    """
    out: dict[int, object] = {}
    if not rows or not os.path.isfile(wb_path):
        return out
    keep_vba = str(wb_path).lower().endswith(".xlsm")
    wbro = None
    try:
        wbro = load_workbook(
            wb_path,
            read_only=True,
            data_only=True,
            keep_vba=keep_vba,
        )
    except Exception:
        return out
    try:
        if EXCLUDE_RULES_SHEET_NAME not in wbro.sheetnames:
            return out
        wsro = wbro[EXCLUDE_RULES_SHEET_NAME]
        for r in rows:
            if r < 2:
                continue
            try:
                out[r] = wsro.cell(row=r, column=c_d).value
            except Exception:
                pass
    finally:
        if wbro is not None:
            try:
                wbro.close()
            except Exception:
                pass
    return out


def run_exclude_rules_sheet_maintenance(
    wb_path: str, pairs: list[tuple[str, str]], log_prefix: str
) -> None:
    """
    「設定_配台不要工程」の行同期・D→E の AI 補完・**openpyxl でブック保存**（A〜E）。

    ブックが Excel で開かれて保存できないときは ``log/exclude_rules_matrix_vba.tsv`` を残し、マクロ
    ``設定_配台不要工程_AからE_TSVから反映`` で A〜E を反映する。
    併せて従来どおり E 列のみの ``exclude_rules_e_column_vba.tsv`` も出力され得る（行列 TSV 優先で反映後は削除）。
    保存成功時は TSV/JSON は削除される。

    ``log/exclude_rules_e_column_pending.json`` は Python 次回起動時の E 列復元用。
    シートの新規作成と 1 行目見出しは VBA「設定_配台不要工程_シートを確保」。
    """
    if not wb_path:
        _log_exclude_rules_sheet_debug(
            "SKIP_NO_PATH",
            log_prefix,
            "TASK_INPUT_WORKBOOK が空のため設定シート処理をしません。",
        )
        return
    if not os.path.exists(wb_path):
        _log_exclude_rules_sheet_debug(
            "SKIP_NO_FILE",
            log_prefix,
            "ブックが存在しません。",
            details=f"path={wb_path}",
        )
        return

    _log_exclude_rules_sheet_debug(
        "START",
        log_prefix,
        "設定シート保守開始",
        details=f"path={wb_path} pairs={len(pairs)}",
    )
    global _exclude_rules_effective_read_path
    _exclude_rules_effective_read_path = None

    keep_vba = str(wb_path).lower().endswith(".xlsm")
    wb = None
    try:
        wb = load_workbook(wb_path, keep_vba=keep_vba, read_only=False, data_only=False)
    except Exception as e1:
        if keep_vba:
            _log_exclude_rules_sheet_debug(
                "OPEN_RETRY",
                log_prefix,
                "keep_vba=True でブックを開けず keep_vba=False で再試行します（マクロが失われる可能性）。",
                exc=e1,
            )
            try:
                wb = load_workbook(wb_path, keep_vba=False, read_only=False, data_only=False)
            except Exception as e2:
                _log_exclude_rules_sheet_debug(
                    "OPEN_FAIL",
                    log_prefix,
                    "ブックを開けません。シートは作成・保存されません。",
                    details=f"path={wb_path}",
                    exc=e2,
                )
                return
        else:
            _log_exclude_rules_sheet_debug(
                "OPEN_FAIL",
                log_prefix,
                "ブックを開けません。シートは作成・保存されません。",
                details=f"path={wb_path}",
                exc=e1,
            )
            return

    _log_exclude_rules_sheet_debug(
        "OPEN_OK",
        log_prefix,
        "ブックを開きました。",
        details=f"keep_vba={keep_vba} sheets={len(wb.sheetnames)}",
    )

    try:
        if EXCLUDE_RULES_SHEET_NAME not in wb.sheetnames:
            _log_exclude_rules_sheet_debug(
                "SKIP_NO_SHEET",
                log_prefix,
                "シートがありません。VBA の「設定_配台不要工程_シートを確保」を実行するか、段階1/2 をマクロから起動してください。",
                details=f"path={wb_path}",
            )
            logging.error(
                "%s: 「%s」がありません。Python ではシートを作成しません。",
                log_prefix,
                EXCLUDE_RULES_SHEET_NAME,
            )
            return

        ws = wb[EXCLUDE_RULES_SHEET_NAME]
        hm_before = _exclude_rules_sheet_header_map(ws)
        c_proc, c_mach, c_flag, c_d, c_e = _ensure_exclude_rules_sheet_headers_and_columns(
            ws, log_prefix
        )
        hm_after = _exclude_rules_sheet_header_map(ws)
        if tuple(hm_before.get(x) for x in (
            EXCLUDE_RULE_COL_PROCESS,
            EXCLUDE_RULE_COL_MACHINE,
            EXCLUDE_RULE_COL_FLAG,
            EXCLUDE_RULE_COL_LOGIC_JA,
            EXCLUDE_RULE_COL_LOGIC_JSON,
        )) != tuple(hm_after.get(x) for x in (
            EXCLUDE_RULE_COL_PROCESS,
            EXCLUDE_RULE_COL_MACHINE,
            EXCLUDE_RULE_COL_FLAG,
            EXCLUDE_RULE_COL_LOGIC_JA,
            EXCLUDE_RULE_COL_LOGIC_JSON,
        )):
            _log_exclude_rules_sheet_debug(
                "HEADER_FIX",
                log_prefix,
                "1行目に標準見出しを書き込みました（空シート・列名不一致の補正）。",
                details=f"cols=({c_proc},{c_mach},{c_flag},{c_d},{c_e})",
            )

        # 前回ブック保存に失敗したとき退避した E 列を、先にワークシートへ戻す（続く保存でディスクへ載る）
        _try_apply_pending_exclude_rules_e_column(wb_path, ws, c_e, log_prefix)

        existing_keys: set[tuple[str, str]] = set()
        max_r = max(2, int(ws.max_row or 2))
        for r in range(2, max_r + 1):
            pv = ws.cell(row=r, column=c_proc).value
            mv = ws.cell(row=r, column=c_mach).value
            p = str(pv).strip() if pv is not None and not (isinstance(pv, float) and pd.isna(pv)) else ""
            m = str(mv).strip() if mv is not None and not (isinstance(mv, float) and pd.isna(mv)) else ""
            if not p:
                continue
            existing_keys.add(
                (_normalize_process_name_for_rule_match(p), _normalize_equipment_match_key(m))
            )

        added = 0
        for p, m in pairs:
            key = (_normalize_process_name_for_rule_match(p), _normalize_equipment_match_key(m))
            if key in existing_keys:
                continue
            ws.append([p, m, None, None, None])
            existing_keys.add(key)
            added += 1
        if added:
            _log_exclude_rules_sheet_debug(
                "SYNC_ROWS",
                log_prefix,
                f"工程+機械の行を {added} 件追加しました。",
            )
            logging.info(
                "%s: 「%s」に工程+機械の組み合わせを %s 行追加しました。",
                log_prefix,
                EXCLUDE_RULES_SHEET_NAME,
                added,
            )

        # 加工計画からペアが1件も取れず、シートにもデータ行が無いときは例行のみ（従来の新規シート相当）
        if added == 0 and not existing_keys:
            ws.append(["梱包", "", "yes", "", ""])
            existing_keys.add(
                (_normalize_process_name_for_rule_match("梱包"), _normalize_equipment_match_key(""))
            )
            _log_exclude_rules_sheet_debug(
                "EXAMPLE_ROW",
                log_prefix,
                "データ行が無かったため例（梱包=yes）を1行追加。",
            )
            logging.info(
                "%s: 「%s」にデータ行が無かったため、例（梱包=yes）を1行追加しました。",
                log_prefix,
                EXCLUDE_RULES_SHEET_NAME,
            )

        # 空行詰めは AI より先に行う（後から詰めると、書き込んだ行番号と画面上の行がずれる）
        n_kept, n_removed_empty = _compact_exclude_rules_data_rows(
            ws, c_proc, c_mach, c_flag, c_d, c_e, log_prefix
        )
        if n_removed_empty:
            _log_exclude_rules_sheet_debug(
                "DATA_COMPACT",
                log_prefix,
                "空行を削除してデータ行を詰めました（並び順は維持）。AI 補完より前。",
                details=f"rows={n_kept} removed_empty={n_removed_empty}",
            )

        max_r = int(ws.max_row or 1)
        pending_rows: list[int] = []
        for r in range(2, max_r + 1):
            dv = ws.cell(row=r, column=c_d).value
            ev = ws.cell(row=r, column=c_e).value
            # C 列の有無に関係なく、D に説明があり E が空なら D→E を試す
            if _cell_is_blank_for_rule(dv):
                continue
            if not _cell_is_blank_for_rule(ev):
                continue
            pending_rows.append(r)

        # D が数式のときは通常読込では '=...' だけ取れる。data_only でキャッシュ表示値を補う。
        formula_rows = [
            r
            for r in pending_rows
            if isinstance(ws.cell(row=r, column=c_d).value, str)
            and str(ws.cell(row=r, column=c_d).value).strip().startswith("=")
        ]
        d_cached = (
            _read_exclude_rules_d_cells_data_only_for_rows(wb_path, formula_rows, c_d)
            if formula_rows
            else {}
        )
        pending_texts: list[str] = []
        filtered_rows: list[int] = []
        for r in pending_rows:
            dv = ws.cell(row=r, column=c_d).value
            blob = (
                ""
                if dv is None or (isinstance(dv, float) and pd.isna(dv))
                else str(dv).strip()
            )
            if blob.startswith("="):
                alt = d_cached.get(r)
                if alt is not None and not (isinstance(alt, float) and pd.isna(alt)):
                    blob = str(alt).strip()
                else:
                    logging.warning(
                        "%s: 「%s」%s 行目の D 列が数式で、キャッシュ値を読めませんでした（Excel で一度保存するか D を値にしてください）。",
                        log_prefix,
                        EXCLUDE_RULES_SHEET_NAME,
                        r,
                    )
                    continue
            if _cell_is_blank_for_rule(blob):
                continue
            filtered_rows.append(r)
            pending_texts.append(blob)
        pending_rows = filtered_rows

        ai_filled = 0
        ai_e_cell_addrs: list[str] = []
        if pending_texts:
            parsed_list = _ai_compile_exclude_rule_logics_batch(pending_texts)
            for r, parsed in zip(pending_rows, parsed_list):
                if not parsed:
                    logging.warning(
                        "%s: 「%s」%s 行目の D 列を JSON にできませんでした（APIキー・応答を確認）。",
                        log_prefix,
                        EXCLUDE_RULES_SHEET_NAME,
                        r,
                    )
                    continue
                jstr = json.dumps(parsed, ensure_ascii=False)
                ws.cell(row=r, column=c_e, value=jstr)
                cell_addr = f"{get_column_letter(c_e)}{r}"
                ai_e_cell_addrs.append(cell_addr)
                preview = jstr if len(jstr) <= 160 else (jstr[:160] + "…")
                logging.info(
                    "%s: 「%s」ロジック式列「%s」セル %s に JSON を書き込み: %s",
                    log_prefix,
                    EXCLUDE_RULES_SHEET_NAME,
                    EXCLUDE_RULE_COL_LOGIC_JSON,
                    cell_addr,
                    preview,
                )
                ai_filled += 1
        if ai_filled:
            _log_exclude_rules_sheet_debug(
                "AI_E_FILLED",
                log_prefix,
                f"D→E の AI 補完を {ai_filled} 行実施。",
                details="cells=" + ",".join(ai_e_cell_addrs),
            )
            logging.info(
                "%s: 「%s」で D→E の AI 補完を %s 行（セル: %s）。",
                log_prefix,
                EXCLUDE_RULES_SHEET_NAME,
                ai_filled,
                ",".join(ai_e_cell_addrs),
            )

        _er_test = os.environ.get("EXCLUDE_RULES_TEST_E1234", "").strip().lower()
        if _er_test in ("1", "yes", "true"):
            try:
                _er_row = int(os.environ.get("EXCLUDE_RULES_TEST_E1234_ROW", "9") or "9")
            except ValueError:
                _er_row = 9
            if _er_row < 2:
                _er_row = 9
            ws.cell(row=_er_row, column=c_e, value="1234")
            _e_addr = f"{get_column_letter(c_e)}{_er_row}"
            _log_exclude_rules_sheet_debug(
                "TEST_E1234",
                log_prefix,
                f'E列 {_e_addr} にテストで "1234" を書き込み',
                details=f"row={_er_row}",
            )
            logging.warning(
                '%s: 【テスト】%s に "1234" を書き込み（EXCLUDE_RULES_TEST_E1234）。',
                log_prefix,
                _e_addr,
            )

        _set_exclude_rules_snapshot_from_ws(
            wb_path, ws, c_proc, c_mach, c_flag, c_e
        )
        _write_exclude_rules_e_apply_artifacts(wb_path, ws, c_e, log_prefix)
        persisted = _persist_exclude_rules_workbook(wb, wb_path, ws, log_prefix)
        if not persisted:
            logging.warning(
                "%s: 設定シートの openpyxl 保存に失敗しました。"
                " log の行列 TSV をマクロ「設定_配台不要工程_AからE_TSVから反映」、"
                "または E 列のみ「設定_配台不要工程_E列_TSVから反映」で反映してください。",
                log_prefix,
            )
    except Exception as ex:
        _log_exclude_rules_sheet_debug(
            "FATAL",
            log_prefix,
            "設定シート処理中に未捕捉例外が発生しました。",
            exc=ex,
        )
        logging.exception("%s: 設定_配台不要工程の処理で例外", log_prefix)
    finally:
        if wb is not None:
            wb.close()
            _log_exclude_rules_sheet_debug("CLOSED", log_prefix, "ブックをクローズしました。")


def _resolve_exclude_rules_workbook_path_for_read(wb_path: str) -> str:
    """直前の保守で実効パスが変わったとき（通常は COM/保存成功後の元ブック）にそれを使う。"""
    p = _exclude_rules_effective_read_path
    if p and os.path.exists(p):
        return p
    return wb_path


def _load_exclude_rules_from_workbook(wb_path: str) -> list[dict]:
    """シートからルール行を読み、評価用リストを返す。"""
    if not wb_path:
        return []
    global _exclude_rules_rules_snapshot, _exclude_rules_snapshot_wb
    ap_arg = os.path.normcase(os.path.abspath(wb_path))
    if (
        _exclude_rules_rules_snapshot is not None
        and _exclude_rules_snapshot_wb == ap_arg
    ):
        snap = list(_exclude_rules_rules_snapshot)
        _exclude_rules_rules_snapshot = None
        _exclude_rules_snapshot_wb = None
        return snap
    path = _resolve_exclude_rules_workbook_path_for_read(wb_path)
    if not os.path.exists(path):
        return []
    try:
        df = pd.read_excel(path, sheet_name=EXCLUDE_RULES_SHEET_NAME)
    except Exception:
        return []
    df.columns = df.columns.str.strip()
    need = [EXCLUDE_RULE_COL_PROCESS, EXCLUDE_RULE_COL_MACHINE]
    for c in need:
        if c not in df.columns:
            return []
    rules = []
    for _, row in df.iterrows():
        proc = str(row.get(EXCLUDE_RULE_COL_PROCESS, "") or "").strip()
        if not proc:
            continue
        mach = str(row.get(EXCLUDE_RULE_COL_MACHINE, "") or "").strip()
        c_val = row.get(EXCLUDE_RULE_COL_FLAG)
        e_raw = row.get(EXCLUDE_RULE_COL_LOGIC_JSON)
        parsed = _parse_exclude_rule_json_cell(e_raw)
        rules.append(
            {
                "proc": proc,
                "mach": mach,
                "c_val": c_val,
                "parsed": parsed,
            }
        )
    return rules


def apply_exclude_rules_config_to_plan_df(
    df: pd.DataFrame, wb_path: str, log_prefix: str
) -> pd.DataFrame:
    """設定シートに基づき「配台不要」を設定（C=yes または E の JSON が真）。"""
    if df is None or df.empty:
        return df
    if TASK_COL_MACHINE not in df.columns or PLAN_COL_EXCLUDE_FROM_ASSIGNMENT not in df.columns:
        return df
    rules = _load_exclude_rules_from_workbook(wb_path)
    if not rules:
        return df
    df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT].astype(object)
    n = 0
    for i in df.index:
        try:
            row = df.loc[i]
        except Exception:
            continue
        tp = str(row.get(TASK_COL_MACHINE, "") or "").strip()
        tm = str(row.get(TASK_COL_MACHINE_NAME, "") or "").strip()
        if not tp:
            continue
        for ru in rules:
            if not _task_row_matches_exclude_rule_target(tp, tm, ru["proc"], ru["mach"]):
                continue
            if _exclude_rule_c_column_is_yes(ru["c_val"]):
                df.at[i, PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = "yes"
                n += 1
                break
            if ru.get("parsed") and evaluate_exclude_rule_json_for_row(ru["parsed"], row):
                df.at[i, PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = "yes"
                n += 1
                break
    if n:
        logging.info("%s: 設定「%s」により配台不要=yes を %s 行に設定しました。", log_prefix, EXCLUDE_RULES_SHEET_NAME, n)
    return df


# =============================================================================
# 段階1エントリ（task_extract_stage1.py → run_stage1_extract）
#   加工計画DATA 読取 → 配台不要自動処理 → 設定シート保守 → plan_input_tasks.xlsx 出力
# =============================================================================
def run_stage1_extract():
    """
    段階1: 加工計画DATA から配台用タスク一覧を抽出し output/plan_input_tasks.xlsx へ出力。
    同一依頼NOで同一機械名が複数行あるとき、工程名「分割」行の空の「配台不要」に yes を自動設定する。
    マクロブックの「設定_配台不要工程」で工程+機械ごとの配台不要・条件式（AI）を管理する（シート作成は VBA）。
    """
    if not TASKS_INPUT_WORKBOOK:
        logging.error("TASK_INPUT_WORKBOOK が未設定です。")
        return False
    if not os.path.exists(TASKS_INPUT_WORKBOOK):
        logging.error(f"TASK_INPUT_WORKBOOK が存在しません: {TASKS_INPUT_WORKBOOK}")
        return False
    reset_gemini_usage_tracker()
    df_src = load_tasks_df()
    try:
        _pm_pairs = _collect_process_machine_pairs_for_exclude_rules(df_src)
        run_exclude_rules_sheet_maintenance(TASKS_INPUT_WORKBOOK, _pm_pairs, "段階1")
    except Exception:
        logging.exception("段階1: 設定_配台不要工程の保守で例外（続行）")
    records = []
    for _, row in df_src.iterrows():
        if row_has_completion_keyword(row):
            continue
        task_id = str(row.get(TASK_COL_TASK_ID, "")).strip()
        machine = str(row.get(TASK_COL_MACHINE, "")).strip()
        machine_name = str(row.get(TASK_COL_MACHINE_NAME, "")).strip()
        qty_total = parse_float_safe(row.get(TASK_COL_QTY), 0.0)
        done_qty = calc_done_qty_equivalent_from_row(row)
        qty = max(0.0, qty_total - done_qty)
        if qty <= 0 or not machine or not task_id:
            continue
        rec = {c: row.get(c) for c in SOURCE_BASE_COLUMNS}
        # 工程名 + 機械名 を“因子”として表示用に追加（後段は計算キーにも使用）
        if machine_name:
            rec[PLAN_COL_PROCESS_FACTOR] = f"{machine}+{machine_name}"
        else:
            rec[PLAN_COL_PROCESS_FACTOR] = f"{machine}+"
        rec[PLAN_COL_REQUIRED_OP] = ""
        rec[PLAN_COL_SPEED_OVERRIDE] = ""
        rec[PLAN_COL_TASK_EFFICIENCY] = ""
        rec[PLAN_COL_PRIORITY] = ""
        rec[PLAN_COL_SPECIFIED_DUE_OVERRIDE] = ""
        rec[PLAN_COL_START_DATE_OVERRIDE] = ""
        rec[PLAN_COL_START_TIME_OVERRIDE] = ""
        rec[PLAN_COL_PREFERRED_OP] = ""
        rec[PLAN_COL_SPECIAL_REMARK] = ""
        rec[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = ""
        rec[PLAN_COL_AI_PARSE] = ""
        records.append(rec)
    if not records:
        logging.warning("段階1: 抽出対象タスクがありません。")
    order = plan_input_sheet_column_order()
    out_df = pd.DataFrame(records)
    if out_df.empty:
        out_df = pd.DataFrame(columns=order)
    else:
        out_df = out_df.reindex(columns=order).fillna("")
    if PLAN_COL_EXCLUDE_FROM_ASSIGNMENT in out_df.columns:
        out_df[PLAN_COL_EXCLUDE_FROM_ASSIGNMENT] = out_df[
            PLAN_COL_EXCLUDE_FROM_ASSIGNMENT
        ].astype(object)
    try:
        _, _, _, req_map, need_rules, _ = load_skills_and_needs()
    except Exception as e:
        logging.info("段階1: マスタ need を読めず元列は need なしで埋めます (%s)", e)
        req_map, need_rules = {}, []
    out_df = _merge_plan_sheet_user_overrides(out_df)
    _refresh_plan_reference_columns(out_df, req_map, need_rules)
    try:
        _apply_auto_exclude_bunkatsu_duplicate_machine(out_df, log_prefix="段階1")
    except Exception as ex:
        logging.exception("段階1: 分割行の配台不要自動設定で例外（出力は続行）: %s", ex)
    try:
        out_df = apply_exclude_rules_config_to_plan_df(out_df, TASKS_INPUT_WORKBOOK, "段階1")
    except Exception as ex:
        logging.warning("段階1: 設定シートによる配台不要適用で例外（続行）: %s", ex)
    out_path = os.path.join(output_dir, STAGE1_OUTPUT_FILENAME)
    out_df.to_excel(out_path, sheet_name="タスク一覧", index=False)
    _apply_excel_date_columns_date_only_display(out_path, "タスク一覧")
    _apply_plan_input_visual_format(out_path, "タスク一覧")
    logging.info(f"段階1完了: '{out_path}' を出力しました。マクロで '{PLAN_INPUT_SHEET_NAME}' に取り込んでください。")
    _try_write_main_sheet_gemini_usage_summary("段階1")
    return True


# 稼働ルール（デフォルト値・2026年3月基準）
TARGET_YEAR = 2026
TARGET_MONTH = 3
DEFAULT_START_TIME = time(8, 45)
DEFAULT_END_TIME = time(17, 0)
DEFAULT_BREAKS = [
    (time(12, 0), time(12, 50)),
    (time(14, 45), time(15, 0))
]

# =========================================================
# 1. コア計算ロジック (日時ベース)
#    休憩帯を挟んだ「実働分」換算・終了時刻の繰り上げ。割付ループの下回り。
# =========================================================
def merge_time_intervals(intervals):
    """時刻区間のリストをソートし、重なる区間を結合して返す。"""
    if not intervals:
        return []
    intervals.sort(key=lambda x: x[0])
    merged = [intervals[0]]
    for current_start, current_end in intervals[1:]:
        last_start, last_end = merged[-1]
        if current_start <= last_end:
            merged[-1] = (last_start, max(last_end, current_end))
        else:
            merged.append((current_start, current_end))
    return merged


def get_actual_work_minutes(start_dt, end_dt, breaks_dt):
    """
    start_dt から end_dt までの「休憩を除いた実働分数」。
    breaks_dt … (区間開始, 区間終了) の列（datetime または time。呼び出し元の勤怠イベントと整合）。
    """
    current = start_dt
    actual_mins = 0
    while current < end_dt:
        next_event = end_dt
        in_break = False
        b_end_time = None
        for b_s, b_e in breaks_dt:
            if b_s <= current < b_e:
                in_break = True
                b_end_time = b_e
                break
            elif current < b_s < next_event:
                next_event = b_s
        
        if in_break:
            current = b_end_time
        else:
            actual_mins += int((next_event - current).total_seconds() / 60)
            current = next_event
    return actual_mins


def calculate_end_time(start_dt, duration_minutes, breaks_dt, end_limit_dt):
    """
    start_dt から実働 duration_minutes 分進めた終了 datetime を求める（休憩はスキップ）。
    end_limit_dt を超えないよう打ち切り。戻り値: (終了時刻, 実際に進めた実働分, 残り未消化分)
    """
    current = start_dt
    remaining_work = duration_minutes
    actual_work_time = 0 

    while current < end_limit_dt and remaining_work > 0:
        next_event = end_limit_dt
        in_break = False
        break_end = None
        for b_start, b_end in breaks_dt:
            if b_start <= current < b_end:
                in_break = True
                break_end = b_end
                break
            elif current < b_start < next_event:
                next_event = b_start
                
        if in_break:
            current = break_end
            continue
            
        block_mins = int((next_event - current).total_seconds() / 60)
        if remaining_work <= block_mins:
            actual_work_time += remaining_work
            current += timedelta(minutes=remaining_work)
            remaining_work = 0
        else:
            actual_work_time += block_mins
            remaining_work -= block_mins
            current = next_event

    end_dt = min(current, end_limit_dt)
    return end_dt, actual_work_time, remaining_work

def match_need_sheet_condition(condition_raw: str, task_id: str) -> bool:
    """
    need シート「依頼NO条件」欄の解釈。
    空・*・全件 → 常にマッチ。
    prefix:ABC / 接頭辞:ABC → 依頼NO がその文字列で始まる
    regex:... / 正規表現:... → 正規表現（部分一致）
    それ以外の短文は接頭辞として扱う。従来の日本語例「依頼NOがJRで…」は JR を検出したら接頭辞JR扱い。
    """
    cond = (condition_raw or "").strip()
    tid = str(task_id).strip()
    if not cond or cond in ("*", "全件", "全て", "any", "ANY"):
        return True
    low = cond.lower()
    cn = cond.replace("：", ":")
    if low.startswith("prefix:") or low.startswith("接頭辞:"):
        pref = cn.split(":", 1)[1].strip() if ":" in cn else ""
        return bool(pref) and tid.startswith(pref)
    if low.startswith("regex:") or low.startswith("正規表現:"):
        pat = cn.split(":", 1)[1].strip() if ":" in cn else ""
        if not pat:
            return False
        try:
            return re.search(pat, tid) is not None
        except re.error:
            logging.warning(f"need 依頼NO条件の正規表現が無効です: {pat}")
            return False
    if "依頼" in cond and "JR" in cond.upper():
        return tid.upper().startswith("JR")
    return tid.startswith(cond)


def parse_need_sheet_special_rules(needs_df, label_col, equipment_list, cond_col):
    """特別指定1～99 行から、設備別の必要人数上書き（1～99）を抽出（先に定義された番号が優先）。"""
    rules = []
    for _, row in needs_df.iterrows():
        lab = str(row.get(label_col, "") or "").strip()
        m = re.match(r"特別指定\s*(\d+)", lab)
        if not m:
            continue
        order = int(m.group(1))
        if order < 1 or order > 99:
            continue
        cond = ""
        if cond_col is not None:
            v = row.get(cond_col)
            if v is not None and not (isinstance(v, float) and pd.isna(v)):
                cond = str(v).strip()
        overrides = {}
        for eq in equipment_list:
            v = row.get(eq)
            n = parse_optional_int(v)
            if n is not None and 1 <= n <= 99:
                overrides[str(eq).strip()] = n
        if not overrides:
            continue
        rules.append({"order": order, "condition": cond, "overrides": overrides})
    rules.sort(key=lambda r: r["order"])
    return rules


def resolve_need_required_op(process: str, machine_name: str, task_id: str, req_map: dict, need_rules: list) -> int:
    """
    need シートの「工程名 + 機械名」で必要OP人数を解決（特別指定1〜99は order が小さいほど優先）。

    req_map は
      - f\"{process}+{machine_name}\"（厳密キー）
      - machine_name（機械だけのフォールバック）
      - process（工程だけのフォールバック）
    のいずれかで base を引ける前提。
    need_rules の overrides も同様にキーを持つ。
    """
    p = str(process).strip()
    m = str(machine_name).strip()

    combo_key = f"{p}+{m}" if p and m else None

    base = None
    if combo_key and combo_key in req_map:
        base = req_map.get(combo_key)
    if base is None and m:
        base = req_map.get(m)
    if base is None and p:
        base = req_map.get(p)
    if base is None:
        base = 1

    for rule in need_rules:
        if not match_need_sheet_condition(rule["condition"], task_id):
            continue

        if combo_key and combo_key in rule["overrides"]:
            return int(rule["overrides"][combo_key])
        if m and m in rule["overrides"]:
            return int(rule["overrides"][m])
        if p and p in rule["overrides"]:
            return int(rule["overrides"][p])

    return int(base)


def _need_row_label_hints_surplus_add(label_a0: str) -> bool:
    """need シート A列: 基本必要人数の直下にある「配台結果で余剰が出たときの追加増員上限」行か。"""
    s = unicodedata.normalize("NFKC", str(label_a0 or "").strip())
    if not s or s.startswith("特別指定"):
        return False
    if "依頼" in s and "条件" in s:
        return False
    if "追加" in s and ("人数" in s or "人員" in s or "増員" in s):
        return True
    if "増員" in s or "余剰" in s:
        return True
    if "配台" in s and ("追加" in s or "増" in s or "余剰" in s):
        return True
    return False


def _find_need_surplus_add_row_index(
    needs_raw, base_row: int, col0: int, pm_cols: list
) -> int | None:
    """基本必要人数行の次行を優先。ラベルまたは数値で追加人数行と判定。"""
    r = base_row + 1
    if r >= needs_raw.shape[0]:
        return None
    v0 = needs_raw.iat[r, col0]
    s0 = "" if pd.isna(v0) else str(v0).strip()
    if s0.startswith("特別指定"):
        return None
    if _need_row_label_hints_surplus_add(s0):
        return r
    nz = 0
    for col_idx, _, _ in pm_cols:
        if parse_optional_int(needs_raw.iat[r, col_idx]) is not None:
            nz += 1
    if nz > 0 and not unicodedata.normalize("NFKC", s0).startswith("特別"):
        return r
    return None


def resolve_need_surplus_extra_max(
    process: str,
    machine_name: str,
    task_id: str,
    surplus_map: dict,
    need_rules: list,
) -> int:
    """
    need シート「配台時追加人数」行（工程×機械列）の値＝必要人数を満たしたうえで
    さらに割り当て可能な人数の上限（0 なら従来どおり必要人数ちょうどのみ）。
    need_rules は現状この行を上書きしない（将来拡張用に task_id を受け取る）。
    """
    _ = (task_id, need_rules)
    if not surplus_map:
        return 0
    p = str(process).strip()
    m = str(machine_name).strip()
    combo_key = f"{p}+{m}" if p and m else None
    v = None
    if combo_key and combo_key in surplus_map:
        v = surplus_map[combo_key]
    elif m and m in surplus_map:
        v = surplus_map[m]
    elif p and p in surplus_map:
        v = surplus_map[p]
    if v is None:
        return 0
    try:
        n = int(v)
    except (TypeError, ValueError):
        return 0
    return max(0, min(n, 50))


def _surplus_team_time_factor(
    rq_base: int, team_len: int, extra_max_allowed: int
) -> float:
    """
    必要人数を超えて入れたメンバーによる単位時間への係数（1.0＝短縮なし）。
    追加枠（extra_max_allowed）を使い切ったときでも、短縮は SURPLUS_TEAM_MAX_SPEEDUP_RATIO を上限とする線形モデル。
    """
    rq = max(1, int(rq_base))
    tl = int(team_len)
    em = max(0, int(extra_max_allowed))
    extra = max(0, tl - rq)
    if extra <= 0 or em <= 0:
        return 1.0
    frac = min(1.0, extra / float(em))
    return 1.0 - SURPLUS_TEAM_MAX_SPEEDUP_RATIO * frac


def _team_assignment_sort_tuple(
    team: tuple,
    team_start: datetime,
    units_today: int,
    team_prio_sum: int,
) -> tuple:
    """
    チーム候補の優劣用タプル（辞書式で小さい方が採用）。
    TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF のとき -len(team) を先頭に置き余剰活用を優先。
    """
    n = len(team)
    if TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF:
        return (-n, team_start, -units_today, team_prio_sum)
    return (team_start, -units_today, team_prio_sum)


# skills セル: OP / AS + 任意の優先度整数（例 OP1, AS 3）。数値が小さいほど割当で先に選ばれる。
_SKILL_OP_AS_CELL_RE = re.compile(r"^(OP|AS)(\d*)$", re.IGNORECASE)


def parse_op_as_skill_cell(cell_val):
    """
    master.xlsm「skills」のセル1つを解釈する。
    - 「OP」または「AS」の直後に優先度用の整数（空白は除去して解釈）。例: OP, OP1, AS3, AS 12
    - 優先度は小さいほど高優先（同一条件のチーム候補から先に選ばれる）。数字省略時は 1。
    - OP/AS で始まらない・空はスキルなし。
    """
    if cell_val is None or (isinstance(cell_val, float) and pd.isna(cell_val)):
        return None, 10**9
    s = str(cell_val).strip()
    if not s:
        return None, 10**9
    compact = re.sub(r"\s+", "", s).upper()
    m = _SKILL_OP_AS_CELL_RE.match(compact)
    if not m:
        return None, 10**9
    role = m.group(1).upper()
    tail = m.group(2) or ""
    if tail == "":
        pr = 1
    else:
        try:
            pr = int(tail)
        except ValueError:
            return None, 10**9
    if pr < 0:
        pr = 0
    return role, pr


def _normalize_person_name_for_match(s):
    """担当者指名のあいまい一致用（NFKC・富田/冨田の表記寄せ・空白除去・末尾敬称のみ除去）。"""
    if s is None:
        return ""
    t = unicodedata.normalize("NFKC", str(s).strip())
    if "富田" in t:
        t = t.replace("富田", "冨田")
    t = re.sub(r"[\s　]+", "", t)
    t = re.sub(r"(さん|様|氏)$", "", t)
    return t


def _resolve_preferred_op_to_member(raw, op_candidates):
    """
    自由記述の指名を、当日スキル上OPのメンバー名（skills シートの行キー）に解決する。
    op_candidates: その設備でOPのメンバー名リスト。
    """
    if not raw or not op_candidates:
        return None
    r0 = unicodedata.normalize("NFKC", str(raw).strip())
    r = _normalize_person_name_for_match(r0)
    if not r:
        return None
    for m in op_candidates:
        if _normalize_person_name_for_match(m) == r:
            return m
        if unicodedata.normalize("NFKC", str(m).strip()) == r0.strip():
            return m
    matches = []
    for m in op_candidates:
        mc = _normalize_person_name_for_match(m)
        if mc and (r in mc or mc in r):
            matches.append(m)
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        return min(matches, key=lambda x: len(str(x)))
    return None


# =========================================================
# 2. マスタデータ・出勤簿(カレンダー) と AI解析
#    master.xlsm の skills / need / 各メンバー勤怠シートを読み、
#    備考・休暇区分は必要に応じて Gemini で構造化する。
# =========================================================
def load_skills_and_needs():
    """
    統合ファイル(MASTER_FILE)からスキルと need を動的に読み込みます。

    今回の need は（Excel上で）
      工程名行・機械名行のあと「基本必要人数」行（A列に「必要人数」を含む）
      その直下: 配台で余剰人員があるときに追加で入れられる人数（工程×機械ごと。未設定は 0）
      以降: 特別指定1〜99
    という構造のため、必要OPは「工程名+機械名」で解決する。

    skills 交差セルは OP/AS の後に優先度整数（例 OP1, AS3）。数値が小さいほど当該工程への割当で優先。
    数字省略の OP/AS は優先度 1。
    """
    try:
        # skills は新仕様:
        #   1行目: 工程名
        #   2行目: 機械名
        #   A3以降: メンバー名
        #   交差セル: OP または AS の後に割当優先度の整数（例 OP1, AS3）。数値が小さいほど当該工程へ優先割当。
        #             数字省略の OP/AS は優先度 1（従来どおり最優先扱い）。
        # を基本としつつ、旧仕様（1行ヘッダ）にもフォールバック対応する。
        skills_raw = pd.read_excel(MASTER_FILE, sheet_name="skills", header=None)
        skills_dict = {}
        equipment_list = []
        members = []

        use_two_header = False
        if skills_raw.shape[0] >= 3 and skills_raw.shape[1] >= 2:
            non_empty_pm = 0
            for c in range(1, skills_raw.shape[1]):
                p = skills_raw.iat[0, c]
                m = skills_raw.iat[1, c]
                if pd.isna(p) or pd.isna(m):
                    continue
                p_s = str(p).strip()
                m_s = str(m).strip()
                if p_s and m_s and p_s.lower() != "nan" and m_s.lower() != "nan":
                    non_empty_pm += 1
            use_two_header = non_empty_pm > 0

        if use_two_header:
            pm_cols = []
            seen_eq = set()
            for c in range(1, skills_raw.shape[1]):
                p = skills_raw.iat[0, c]
                m = skills_raw.iat[1, c]
                if pd.isna(p) or pd.isna(m):
                    continue
                p_s = str(p).strip()
                m_s = str(m).strip()
                if not p_s or not m_s or p_s.lower() == "nan" or m_s.lower() == "nan":
                    continue
                combo = f"{p_s}+{m_s}"
                pm_cols.append((c, p_s, m_s, combo))
                if p_s not in seen_eq:
                    seen_eq.add(p_s)
                    equipment_list.append(p_s)

            for r in range(2, skills_raw.shape[0]):
                m_name_raw = skills_raw.iat[r, 0]
                if pd.isna(m_name_raw):
                    continue
                m_name = str(m_name_raw).strip()
                if not m_name or m_name.lower() in ("nan", "none", "null"):
                    continue
                row_skills = {}
                for c, p_s, m_s, combo in pm_cols:
                    v = skills_raw.iat[r, c] if c < skills_raw.shape[1] else None
                    sval = "" if pd.isna(v) else str(v).strip()
                    if not sval or sval.lower() in ("nan", "none", "null"):
                        continue
                    row_skills[combo] = sval
                    if m_s not in row_skills:
                        row_skills[m_s] = sval
                    if p_s not in row_skills:
                        row_skills[p_s] = sval
                skills_dict[m_name] = row_skills
            members = list(skills_dict.keys())
            logging.info(
                "skillsシート: 2段ヘッダ形式で読み込みました（工程+機械=%s列, メンバー=%s人）。",
                len(pm_cols),
                len(members),
            )
        else:
            skills_df = pd.read_excel(MASTER_FILE, sheet_name="skills")
            skills_df.columns = skills_df.columns.str.strip()
            skill_cols = [
                str(c).strip()
                for c in skills_df.columns
                if not str(c).startswith("Unnamed")
            ]

            member_col = None
            for c in skill_cols:
                if c in ("メンバー", "担当者", "氏名", "作業者"):
                    member_col = c
                    break
            if member_col is None and skill_cols:
                member_col = skill_cols[0]
                logging.warning(
                    "skillsシート: メンバー列名が標準と一致しないため、先頭列 '%s' をメンバー列として扱います。",
                    member_col,
                )

            seen_eq = set()
            for c in skill_cols:
                if c == member_col:
                    continue
                proc = c.split("+", 1)[0].strip() if "+" in c else c.strip()
                if proc and proc not in seen_eq:
                    seen_eq.add(proc)
                    equipment_list.append(proc)

            for _, row in skills_df.iterrows():
                m_name = str(row.get(member_col, "")).strip() if member_col else ""
                if not m_name or m_name.lower() == "nan":
                    continue
                row_skills = {}
                for c in skill_cols:
                    if c == member_col:
                        continue
                    sval = str(row.get(c, "")).strip()
                    if not sval or sval.lower() in ("nan", "none", "null"):
                        continue
                    row_skills[c] = sval
                    if "+" in c:
                        p, m = c.split("+", 1)
                        p = p.strip()
                        m = m.strip()
                        if m and m not in row_skills:
                            row_skills[m] = sval
                        if p and p not in row_skills:
                            row_skills[p] = sval
                skills_dict[m_name] = row_skills
            members = list(skills_dict.keys())
            logging.info(
                "skillsシート: 1行ヘッダ形式（旧互換）で読み込みました（メンバー=%s人）。",
                len(members),
            )

        if not members:
            logging.error("skillsシートからメンバーを読み込めませんでした。")

        # need は header=None で読み、先頭の複数行を“見出し行”として解釈
        needs_raw = pd.read_excel(MASTER_FILE, sheet_name="need", header=None)
        col0 = 0
        process_header_row = None
        machine_header_row = None
        base_row = None

        for r in range(needs_raw.shape[0]):
            v0 = needs_raw.iat[r, col0]
            if pd.isna(v0):
                continue
            s0 = str(v0).strip()
            if process_header_row is None and s0 == "工程名":
                process_header_row = r
            elif machine_header_row is None and s0 == "機械名":
                machine_header_row = r
            if base_row is None and "必要人数" in s0 and not s0.startswith("特別指定"):
                base_row = r
            if process_header_row is not None and machine_header_row is not None and base_row is not None:
                break

        if process_header_row is None or machine_header_row is None or base_row is None:
            raise ValueError("need シートのヘッダー行（工程名/機械名/基本必要人数）が見つかりません。")

        # 「依頼NO条件」列位置（デフォルトは 1列目）
        cond_col_idx = 1
        for r in range(needs_raw.shape[0]):
            c1 = needs_raw.iat[r, 1] if needs_raw.shape[1] > 1 else None
            c2 = needs_raw.iat[r, 2] if needs_raw.shape[1] > 2 else None
            if pd.isna(c1) or pd.isna(c2):
                continue
            if str(c1).strip() == NEED_COL_CONDITION and str(c2).strip() == NEED_COL_NOTE:
                cond_col_idx = 1
                break

        # 工程名×機械名 の列一覧（列番号は Excel上の実列を保持）
        pm_cols = []
        for col_idx in range(needs_raw.shape[1]):
            if col_idx < 3:
                continue
            p = needs_raw.iat[process_header_row, col_idx]
            m = needs_raw.iat[machine_header_row, col_idx]
            if pd.isna(p) or pd.isna(m):
                continue
            p_s = str(p).strip()
            m_s = str(m).strip()
            if not p_s or not m_s or p_s.lower() == "nan" or m_s.lower() == "nan":
                continue
            pm_cols.append((col_idx, p_s, m_s))

        req_map = {}
        # need_rules: [{'order': int, 'condition': str, 'overrides': {combo_key/machine/process: int}}]
        need_rules = []

        # 基本必要人数
        for col_idx, p_s, m_s in pm_cols:
            n = parse_optional_int(needs_raw.iat[base_row, col_idx])
            if n is None or n < 1:
                n = 1
            combo_key = f"{p_s}+{m_s}"
            req_map[combo_key] = n
            # フォールバック用（機械名 or 工程名だけで引けるようにする）
            if p_s not in req_map:
                req_map[p_s] = n
            if m_s not in req_map:
                req_map[m_s] = n

        surplus_map: dict[str, int] = {}
        surplus_row = _find_need_surplus_add_row_index(
            needs_raw, base_row, col0, pm_cols
        )
        if surplus_row is not None:
            for col_idx, p_s, m_s in pm_cols:
                raw_ex = parse_optional_int(needs_raw.iat[surplus_row, col_idx])
                ex = int(raw_ex) if raw_ex is not None and raw_ex >= 0 else 0
                ex = max(0, min(ex, 50))
                combo_key = f"{p_s}+{m_s}"
                surplus_map[combo_key] = ex
                if p_s not in surplus_map:
                    surplus_map[p_s] = ex
                if m_s not in surplus_map:
                    surplus_map[m_s] = ex
            logging.info(
                "need シート: 配台時追加人数行を検出（Excel行≈%s）。列ごとの上限を読み込みました。",
                surplus_row + 1,
            )
        else:
            logging.info(
                "need シート: 基本必要人数の直下に配台時追加人数行を検出できませんでした（省略可）。"
            )

        # 特別指定
        for r in range(needs_raw.shape[0]):
            v0 = needs_raw.iat[r, col0]
            if pd.isna(v0):
                continue
            lab = str(v0).strip()
            m = re.match(r"特別指定\s*(\d+)", lab)
            if not m:
                continue
            order = int(m.group(1))
            if order < 1 or order > 99:
                continue

            cond_raw = needs_raw.iat[r, cond_col_idx] if needs_raw.shape[1] > cond_col_idx else None
            cond = "" if pd.isna(cond_raw) else str(cond_raw).strip()

            overrides = {}
            for col_idx, p_s, m_s in pm_cols:
                v = needs_raw.iat[r, col_idx]
                n = parse_optional_int(v)
                if n is not None and 1 <= n <= 99:
                    combo_key = f"{p_s}+{m_s}"
                    overrides[combo_key] = n
                    # フォールバック用
                    overrides[p_s] = n
                    overrides[m_s] = n

            if overrides:
                need_rules.append({"order": order, "condition": cond, "overrides": overrides})

        need_rules.sort(key=lambda rr: rr["order"])
        logging.info(f"need 特別指定ルール: {len(need_rules)} 件（工程名+機械名キー）。")

        logging.info(f"『{MASTER_FILE}』からスキルと設備要件(need)を読み込みました。")
        return skills_dict, members, equipment_list, req_map, need_rules, surplus_map

    except Exception as e:
        logging.error(f"マスタファイル({MASTER_FILE})のスキル/need読み込みエラー: {e}")
        return {}, [], [], {}, [], {}

def generate_default_calendar_dates(year, month):
    cal = calendar.Calendar()
    return [d for d in cal.itermonthdates(year, month) if d.year == year and d.month == month and d.weekday() < 5]

def parse_time_str(time_str, default_time):
    if time_str is None or pd.isna(time_str) or not str(time_str).strip() or str(time_str).strip().lower() == 'null':
        return default_time
    try:
        if isinstance(time_str, time): return time_str
        if isinstance(time_str, datetime): return time_str.time()
        time_str = str(time_str).strip()
        if len(time_str.split(':')) == 3:
            return datetime.strptime(time_str, "%H:%M:%S").time()
        return datetime.strptime(time_str, "%H:%M").time()
    except:
        return default_time

def infer_mid_break_from_reason(reason_text, start_t, end_t):
    """
    備考から中抜け時間を推定するローカル補正。
    AIが中抜けを返さない場合のフェイルセーフとして使う。
    master.xlsm カレンダー由来の休暇区分: 前休=午前年休・午後のみ勤務、後休=午後年休・午前のみ勤務（出勤簿.txt と同義）。
    """
    if reason_text is None:
        return None, None
    txt = str(reason_text).strip()
    if not txt or txt.lower() in ("nan", "none", "null", "通常"):
        return None, None

    noon_end = time(12, 0)
    afternoon_start = time(13, 0)
    mae_kyuu_day_start = time(12, 45)  # 出勤簿.txt 前休の所定出勤（午前は年休）
    # カレンダー記号と一致させる（シフト時刻が誤っている場合の補完用。正しい行では区間が空になり追加されない）
    if txt == "前休":
        # 正しい行は出勤 12:45 以降で補完不要。全日シフトの誤入力時は 12:45 までを中抜け（午前年休相当）
        if start_t and start_t < mae_kyuu_day_start:
            return start_t, mae_kyuu_day_start
        return None, None
    if txt == "後休":
        if end_t and afternoon_start < end_t:
            return afternoon_start, end_t
        return None, None

    # 1) 明示的な時刻範囲（例: 11:00-14:00 / 11:00～14:00）
    m = re.search(r"(\d{1,2}[:：]\d{2})\s*[~〜\-－ー]\s*(\d{1,2}[:：]\d{2})", txt)
    if m:
        s = parse_time_str(m.group(1).replace("：", ":"), None)
        e = parse_time_str(m.group(2).replace("：", ":"), None)
        if s and e and s < e:
            return s, e

    # 2) あいまい語（午前/午後/終日） + 現場離脱・休暇系キーワード
    # 「午後休みです」等は「午後」を含むが、旧ロジックは「抜け」等のみ見ており中抜け推定に到達しなかった
    leave_keywords = (
        "事務所", "会議", "教育", "研修", "外出", "離れ", "抜け", "中抜け", "打合せ",
        "休み", "休暇", "欠勤",
    )
    has_leave_hint = any(k in txt for k in leave_keywords)
    if not has_leave_hint:
        return None, None

    if ("終日" in txt) or ("1日" in txt and "通常" not in txt):
        return start_t, end_t
    if ("午前中" in txt) or ("午前" in txt):
        return start_t, noon_end
    if ("午後" in txt):
        return afternoon_start, end_t

    return None, None


# 結果_カレンダー(出勤簿) の退勤表示。VBA 出勤簿「後休」（午後年休）と同様に実質 12:00 終了とみなす。
_AFTERNOON_OFF_DISPLAY_END = time(12, 0)


def _reason_is_afternoon_off(reason: str) -> bool:
    """後休（午後年休・午前のみ勤務）または備考の午後休系。"""
    r = str(reason or "")
    return ("午後" in r and ("休" in r or "休み" in r)) or ("後休" in r)


def _reason_is_morning_off(reason: str) -> bool:
    """前休（午前年休・午後のみ勤務）。カレンダー由来の略号のみ明示扱い（事務所勤務などと混同しない）。"""
    return "前休" in str(reason or "")


def _calendar_display_clock_out_for_calendar_sheet(entry: dict, day_date: date):
    """
    配台は breaks_dt の午後中抜けで正しくなる一方、end_dt が 17:00 のままだと結果カレンダーの退勤列だけ誤る。
    後休（午後年休）または備考が午後休み系で、定時まで続く午後の中抜けがあるときだけ退勤表示を 12:00 に揃える（end_dt 本体は変更しない）。
    """
    if not entry.get("is_working"):
        return None
    end_dt = entry.get("end_dt")
    if end_dt is None:
        return None
    reason = str(entry.get("reason") or "")
    afternoon_off = _reason_is_afternoon_off(reason)
    if not afternoon_off:
        return None
    breaks_dt = entry.get("breaks_dt") or []
    for b_s, b_e in breaks_dt:
        if b_s is None or b_e is None:
            continue
        bs = b_s.time() if isinstance(b_s, datetime) else b_s
        if isinstance(bs, datetime):
            bs = bs.time()
        if bs < time(12, 0):
            continue
        if isinstance(b_e, datetime):
            be_cmp = b_e
        elif isinstance(b_e, time):
            be_cmp = datetime.combine(day_date, b_e)
        else:
            continue
        if be_cmp >= end_dt - timedelta(seconds=1):
            return datetime.combine(day_date, _AFTERNOON_OFF_DISPLAY_END)
    return None


def _member_schedule_break_cell_label(grid_mid_dt, breaks_dt, shift_end_dt, reason):
    """
    個人_* スケジュールの10分枠が休憩帯に入るときの文言。
    昼食など通常休憩は「休憩」。後休（午後年休）で定時まで工場にいない午後帯は「休暇」。
    前休（午前年休）で午前の欠勤区間が休憩帯として入っている場合は「休暇」。
    """
    reason = str(reason or "")
    afternoon_off = _reason_is_afternoon_off(reason)
    morning_off = _reason_is_morning_off(reason)
    for b_s, b_e in breaks_dt:
        if b_s is None or b_e is None:
            continue
        if not (b_s <= grid_mid_dt < b_e):
            continue
        if isinstance(b_e, datetime) and isinstance(shift_end_dt, datetime):
            bs = b_s.time() if isinstance(b_s, datetime) else b_s
            if isinstance(bs, datetime):
                bs = bs.time()
            if afternoon_off and bs >= time(12, 0) and b_e >= shift_end_dt - timedelta(seconds=2):
                return "休暇"
            if morning_off and bs < time(12, 0):
                be_t = b_e.time() if isinstance(b_e, datetime) else b_e
                if be_t <= time(13, 0):
                    return "休暇"
        return "休憩"
    return None


def _member_schedule_off_shift_label(
    day_date: date,
    grid_mid_dt: datetime,
    d_start_dt: datetime,
    d_end_dt: datetime,
    reason: str,
) -> str:
    """
    個人_* シートで所定出退勤の外側の10分枠。
    前休の午前（工場日の所定開始～午後出勤まで）は年休、後休の午後は年休。それ以外のシフト外は勤務外。
    """
    r = str(reason or "")
    day_start = datetime.combine(day_date, DEFAULT_START_TIME)
    day_end = datetime.combine(day_date, DEFAULT_END_TIME)
    if grid_mid_dt < d_start_dt:
        if _reason_is_morning_off(r) and grid_mid_dt >= day_start:
            return "年休"
        return "勤務外"
    if grid_mid_dt >= d_end_dt:
        if _reason_is_afternoon_off(r) and grid_mid_dt < day_end:
            return "年休"
        return "勤務外"
    return "勤務外"


def _member_schedule_full_day_off_label(entry) -> str:
    """
    全日非勤務（is_working=False）の個人シート列の表示。
    休暇区分が年休（カレンダー *）のときは『年休』。工場休日などは『休』。
    """
    if not entry:
        return "休"
    r = str(entry.get("reason") or "").strip()
    if r == "年休" or r.startswith("年休 "):
        return "年休"
    return "休"


def _attendance_remark_text(row) -> str:
    """
    勤怠1行から「備考」列のテキストのみ取得する。
    勤怠AIの解析リストへの投入はこの列のみ。reason 文字列は load_attendance で備考と休暇区分を合成する。
    """
    if row is None:
        return ""
    try:
        v = row.get(ATT_COL_REMARK)
    except Exception:
        return ""
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return ""
    return s


def _attendance_leave_type_text(row) -> str:
    """勤怠1行から「休暇区分」列（カレンダー由来の 前休/後休 等）。"""
    if row is None:
        return ""
    try:
        v = row.get(ATT_COL_LEAVE_TYPE)
    except Exception:
        return ""
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return ""
    return s


def load_attendance_and_analyze(members):
    attendance_data = {}
    # ※「勤怠備考」は master 各メンバーシートの「備考」列のみ。メイン再優先・特別指定_備考は別API（generate_plan 側で追記）。
    ai_log = {
        "（注）このシートの見方": "先頭2行は勤怠「備考」の出退勤AIのみ。メイン再優先・特別指定は下段のJSONと「_*_AI_API」行。",
        "勤怠備考_AI_API": "なし",
        "勤怠備考_AI_詳細": "解析対象の備考行なし",
    }
    
    # 1. メンバー別シートからの読み込み
    all_records = []
    try:
        xls = pd.ExcelFile(MASTER_FILE)
        for sheet_name in xls.sheet_names:
            if "カレンダー" in sheet_name or sheet_name.lower() in ['skills', 'need', 'tasks']:
                continue 
                
            m_name = sheet_name.strip()
            if m_name not in members:
                continue 
                
            df_sheet = pd.read_excel(xls, sheet_name=sheet_name)
            df_sheet.columns = df_sheet.columns.str.strip()
            df_sheet['メンバー'] = m_name 
            all_records.append(df_sheet)
            
        if all_records:
            df = pd.concat(all_records, ignore_index=True)
            df['日付'] = pd.to_datetime(df['日付'], errors='coerce').dt.date
            df = df.dropna(subset=['日付'])
            logging.info(f"『{MASTER_FILE}』の各メンバーの勤怠シートを読み込みました。")
            _cols = {str(c).strip() for c in df.columns}
            if ATT_COL_REMARK in _cols and ATT_COL_LEAVE_TYPE in _cols:
                logging.info(
                    "勤怠列: AI 入力は「%s」のみ。備考が空の日は「%s」（前休・後休・他拠点勤務など）を reason に反映します。",
                    ATT_COL_REMARK,
                    ATT_COL_LEAVE_TYPE,
                )
            elif ATT_COL_REMARK not in _cols:
                logging.warning(
                    "勤怠データに「%s」列がありません。備考ベースの AI 解析は空扱いになります。",
                    ATT_COL_REMARK,
                )
        else:
            raise FileNotFoundError("有効なメンバー別勤怠シートが見つかりません。")
            
    except Exception as e:
        logging.warning(f"勤怠シート読み込みエラー: {e} デフォルトカレンダーを生成します。")
        default_dates = generate_default_calendar_dates(TARGET_YEAR, TARGET_MONTH)
        records = []
        for d in default_dates:
            for m in members: records.append({'日付': d, 'メンバー': m, '備考': '通常'})
        df = pd.DataFrame(records)

    # 2. AIによる特記事項の解析（備考に1文字でもあれば対象。「通常」明示も含む）
    remarks_to_analyze = []
    for _, row in df.iterrows():
        m = str(row.get('メンバー', '')).strip()
        rem = _attendance_remark_text(row)
        d_str = row['日付'].strftime("%Y-%m-%d") if pd.notna(row['日付']) else ""
        if m in members and rem:
            remarks_to_analyze.append(f"{d_str}_{m} の備考: {rem}")

    if remarks_to_analyze:
        remarks_blob = "\n".join(remarks_to_analyze)
        cache_key = hashlib.sha256(remarks_blob.encode("utf-8")).hexdigest()
        ai_cache = load_ai_cache()

        # 同一備考セットはキャッシュを優先利用し、APIコールを節約
        cached_data = get_cached_ai_result(ai_cache, cache_key)
        if cached_data is not None:
            ai_parsed = cached_data
            ai_log["勤怠備考_AI_API"] = "なし(キャッシュ使用)"
            ai_log["勤怠備考_AI_詳細"] = "キャッシュヒット"
        elif not API_KEY:
            ai_parsed = {}
            ai_log["勤怠備考_AI_API"] = "なし"
            ai_log["勤怠備考_AI_詳細"] = "GEMINI_API_KEY未設定のため勤怠備考AIをスキップ"
            logging.info("GEMINI_API_KEY 未設定のため備考AI解析をスキップしました。")
        else:
            logging.info("■ AIが複数日の特記事項を解析中...")
            ai_log["勤怠備考_AI_API"] = "あり"
            
            prompt = f"""
            以下の各日・メンバーの備考を読み取り、出退勤時刻の変更や中抜け、休日の判定を行い、JSON形式で出力してください。
            マークダウン記号(``` 等)は一切含めず、純粋なJSON文字列のみを返してください。

            【JSONの出力形式（キー名を厳密に守ること）】
            {{
              "YYYY-MM-DD_メンバー名": {{
                "出勤時刻": "HH:MM", 
                "退勤時刻": "HH:MM", 
                "中抜け開始": "HH:MM",
                "中抜け終了": "HH:MM",
                "作業効率": 1.0,     
                "is_holiday": true   
              }}
            }}
            ・出勤時刻/退勤時刻: 各日の「備考」欄のみを根拠に推測（休暇区分は入力に含まれません）。不明や変更なしなら null
            ・中抜け開始/終了: 備考に「11:00～14:00まで抜ける」など一時的な離脱（中抜け）がある場合、その開始・終了時刻。ない場合は null
            ・曖昧語の解釈例:
              - 「午前中は事務所で作業」=> 中抜け開始 "08:45", 中抜け終了 "12:00"
              - 「午後は会議」=> 中抜け開始 "13:00", 中抜け終了 "17:00"
            ・is_holiday: 「休み」「休む」「欠勤」などの場合は true、それ以外は false
            ・作業効率: 0.0〜1.0の数値
            
            【特記事項リスト】
            {chr(10).join(remarks_to_analyze)}
            """
            try:
                client = genai.Client(api_key=API_KEY)
                res = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
                record_gemini_response_usage(res, GEMINI_MODEL_FLASH)
                match = re.search(r'\{.*\}', res.text, re.DOTALL)
                if match:
                    ai_parsed = json.loads(match.group(0))
                    put_cached_ai_result(ai_cache, cache_key, ai_parsed)
                    save_ai_cache(ai_cache)
                    ai_log["勤怠備考_AI_詳細"] = "解析成功"
                else:
                    ai_parsed = {}
                    ai_log["勤怠備考_AI_詳細"] = "JSONパース失敗"
            except Exception as e:
                err_text = str(e)
                is_quota_or_rate = ("429" in err_text) or ("RESOURCE_EXHAUSTED" in err_text)
                retry_sec = extract_retry_seconds(err_text) if is_quota_or_rate else None

                if is_quota_or_rate and retry_sec is not None:
                    wait_sec = min(max(retry_sec, 1.0), 90.0)
                    logging.warning(f"AI通信 429/RESOURCE_EXHAUSTED。{wait_sec:.1f}秒待機して1回だけ再試行します。")
                    time_module.sleep(wait_sec)
                    try:
                        res = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
                        record_gemini_response_usage(res, GEMINI_MODEL_FLASH)
                        match = re.search(r'\{.*\}', res.text, re.DOTALL)
                        if match:
                            ai_parsed = json.loads(match.group(0))
                            put_cached_ai_result(ai_cache, cache_key, ai_parsed)
                            save_ai_cache(ai_cache)
                            ai_log["勤怠備考_AI_詳細"] = "再試行で解析成功"
                        else:
                            ai_parsed = {}
                            ai_log["勤怠備考_AI_詳細"] = "再試行後JSONパース失敗"
                    except Exception as e2:
                        ai_parsed = {}
                        logging.warning(f"AI再試行エラー: {e2}")
                        ai_log["勤怠備考_AI_詳細"] = f"429後再試行失敗: {e2}"
                else:
                    ai_parsed = {}
                    logging.warning(f"AI通信エラー: {e}")
                    ai_log["勤怠備考_AI_詳細"] = str(e)
    else:
        ai_parsed = {}

    # 3. 日付ごとの制約辞書を構築
    for _, row in df.iterrows():
        if pd.isna(row['日付']): continue
        curr_date = row['日付']
        m = str(row.get('メンバー', '')).strip()
        if m not in members: continue

        if curr_date not in attendance_data:
            attendance_data[curr_date] = {}

        key = f"{curr_date.strftime('%Y-%m-%d')}_{m}"
        ai_info = ai_parsed.get(key, {})
        
        is_empty_shift = pd.isna(row.get('出勤時間')) and pd.isna(row.get('退勤時間')) and not ai_info
        is_holiday = ai_info.get("is_holiday", False) or is_empty_shift
        
        ai_eff = ai_info.get("作業効率")
        excel_eff = row.get('作業効率')
        
        if ai_eff is not None:
            eff_val = ai_eff
        elif excel_eff is not None and not pd.isna(excel_eff):
            eff_val = excel_eff
        else:
            eff_val = 1.0
            
        try:
            efficiency = float(eff_val)
        except (ValueError, TypeError):
            efficiency = 1.0

        original_reason = _attendance_remark_text(row)
        leave_type = _attendance_leave_type_text(row)
        if original_reason:
            if (
                leave_type
                and leave_type not in ("通常", "")
                and leave_type not in original_reason
            ):
                reason = f"{leave_type} {original_reason}"
            else:
                reason = original_reason
        elif leave_type and leave_type not in ("通常", ""):
            reason = leave_type
        else:
            reason = '通常' if not is_empty_shift else '休日シフト'

        start_t = parse_time_str(ai_info.get("出勤時刻") or row.get('出勤時間'), DEFAULT_START_TIME)
        end_t = parse_time_str(ai_info.get("退勤時刻") or row.get('退勤時間'), DEFAULT_END_TIME)
        
        b1_s = parse_time_str(row.get('休憩時間1_開始'), DEFAULT_BREAKS[0][0])
        b1_e = parse_time_str(row.get('休憩時間1_終了'), DEFAULT_BREAKS[0][1])
        b2_s = parse_time_str(row.get('休憩時間2_開始'), DEFAULT_BREAKS[1][0])
        b2_e = parse_time_str(row.get('休憩時間2_終了'), DEFAULT_BREAKS[1][1])

        # ★追加: AIから中抜け時間を取得
        mid_break_s = parse_time_str(ai_info.get("中抜け開始"), None)
        mid_break_e = parse_time_str(ai_info.get("中抜け終了"), None)
        # AIが中抜けを返さなかった場合は、備考文言からローカル推定で補完
        if not (mid_break_s and mid_break_e):
            fb_s, fb_e = infer_mid_break_from_reason(reason, start_t, end_t)
            if fb_s and fb_e:
                mid_break_s, mid_break_e = fb_s, fb_e

        def combine_dt(t): return datetime.combine(curr_date, t) if t else None
        
        start_dt = combine_dt(start_t)
        end_dt = combine_dt(end_t)
        breaks_dt = []
        
        # 通常の休憩を追加
        if b1_s and b1_e: breaks_dt.append((combine_dt(b1_s), combine_dt(b1_e)))
        if b2_s and b2_e: breaks_dt.append((combine_dt(b2_s), combine_dt(b2_e)))
        
        # ★追加: 中抜け時間がある場合は、特別な「休憩」としてスケジュール計算に追加
        if mid_break_s and mid_break_e: breaks_dt.append((combine_dt(mid_break_s), combine_dt(mid_break_e)))
        
        attendance_data[curr_date][m] = {
            "is_working": not is_holiday,
            "start_dt": start_dt,
            "end_dt": end_dt,
            "breaks_dt": merge_time_intervals(breaks_dt),
            "efficiency": efficiency,
            "reason": reason
        }

    return attendance_data, ai_log

# =========================================================
# 3. メイン計画生成 (日毎ループ・持ち越し対応)
#    段階2の本体。plan_simulation_stage2 からのみ呼ばれる想定。
#    配台計画シート読込 → タスクキュー → 日付ごとに設備・OP割付 → 結果ブック出力。
# =========================================================
def generate_plan():
    """
    段階2のメイン処理。戻り値なし（ログ・Excel 出力で完結）。

    前提: 環境変数 TASK_INPUT_WORKBOOK、カレントディレクトリがスクリプトフォルダ。
    出力: ``get_dated_output_dir`` 配下の production_plan / member_schedule、および log/execution_log.txt。
    """
    skills_dict, members, equipment_list, req_map, need_rules, surplus_map = (
        load_skills_and_needs()
    )
    if not members:
        return
    reset_gemini_usage_tracker()

    # 段階2の基準日時は「マクロ実行時刻」ではなく「データ抽出日」を使用
    data_extract_dt = _extract_data_extraction_datetime()
    base_now_dt = data_extract_dt if data_extract_dt is not None else datetime.now()
    run_date = base_now_dt.date()
    data_extract_dt_str = (
        base_now_dt.strftime("%Y/%m/%d %H:%M:%S") if data_extract_dt is not None else "—"
    )
    logging.info(
        "計画基準日時: %s（%s）",
        base_now_dt.strftime("%Y/%m/%d %H:%M:%S"),
        "データ抽出日" if data_extract_dt is not None else "現在時刻フォールバック",
    )

    attendance_data, ai_log_data = load_attendance_and_analyze(members)
    global_priority_raw = load_main_sheet_global_priority_override_text()
    _factory_closure_dates = parse_factory_closure_dates_from_global_comment(
        global_priority_raw, run_date.year
    )
    if _factory_closure_dates:
        apply_factory_closure_dates_to_attendance(
            attendance_data, members, _factory_closure_dates
        )
        logging.info(
            "メイン・グローバルコメント: 工場休業扱いの日付 → %s",
            ", ".join(str(x) for x in sorted(_factory_closure_dates)),
        )
    ai_log_data["メイン_グローバル_工場休業日(解析)"] = (
        ", ".join(str(x) for x in sorted(_factory_closure_dates))
        if _factory_closure_dates
        else "（なし）"
    )

    sorted_dates = sorted(list(attendance_data.keys()))
    # 結果シートは「基準日（データ抽出日）」以降のみ表示・計画対象とする
    sorted_dates = [d for d in sorted_dates if d >= run_date]
    if not sorted_dates:
        logging.error("当日以降の処理対象日付がありません。")
        _try_write_main_sheet_gemini_usage_summary("段階2")
        return

    # タスク入力: ブック内「配台計画_タスク入力」（段階1で出力→取り込み後に編集）
    try:
        tasks_df = load_planning_tasks_df()
    except Exception as e:
        logging.error(f"配台計画タスクシート読み込みエラー: {e}")
        _try_write_main_sheet_gemini_usage_summary("段階2")
        return

    global_priority_override = analyze_global_priority_override_comment(
        global_priority_raw, members, run_date.year, ai_sheet_sink=ai_log_data
    )
    if global_priority_raw.strip():
        snip = global_priority_raw[:2500]
        if len(global_priority_raw) > 2500:
            snip += "…"
        ai_log_data["メイン_再優先特別記載(原文)"] = snip
    else:
        ai_log_data["メイン_再優先特別記載(原文)"] = (
            "（空、またはメインシートに「グローバルコメント」見出しが見つかりません）"
        )
    ai_log_data["メイン_再優先特別記載(AI)"] = json.dumps(
        global_priority_override, ensure_ascii=False
    )
    if global_priority_override.get("ignore_skill_requirements"):
        logging.warning(
            "メイン再優先特記: スキル要件を無視して配台します。%s",
            global_priority_override.get("interpretation_ja", ""),
        )
    if global_priority_override.get("ignore_need_minimum"):
        logging.warning(
            "メイン再優先特記: チーム人数を1名に固定します（need・行の必要OP上書きより優先）。%s",
            global_priority_override.get("interpretation_ja", ""),
        )
    if global_priority_override.get("abolish_all_scheduling_limits"):
        logging.warning(
            "メイン再優先特記: 設備専有・原反同日開始・指定開始時刻・マクロ実行時刻下限を適用しません。%s",
            global_priority_override.get("interpretation_ja", ""),
        )

    # 「当日」判定と最早開始時刻には基準日時（データ抽出日）を使う
    macro_now_dt = base_now_dt
    macro_run_date = macro_now_dt.date()
    ai_task_by_tid = analyze_task_special_remarks(
        tasks_df, reference_year=run_date.year, ai_sheet_sink=ai_log_data
    )
    task_queue = build_task_queue_from_planning_df(
        tasks_df, run_date, req_map, ai_task_by_tid, global_priority_override
    )
    # 開始日が非稼働日の場合は、直前の稼働日へ補正（例: 4/4, 4/5 が非稼働なら 4/3 へ）
    working_days = [
        d for d in sorted_dates
        if any(attendance_data[d][m]["is_working"] for m in attendance_data[d])
    ]
    if working_days:
        for t in task_queue:
            req_d = t.get("start_date_req")
            if not isinstance(req_d, date):
                continue
            if req_d in working_days:
                continue
            prev_work = None
            for wd in working_days:
                if wd <= req_d:
                    prev_work = wd
                else:
                    break
            if prev_work is not None:
                if str(t.get("task_id", "")).strip() == DEBUG_TASK_ID:
                    logging.info(
                        "DEBUG[task=%s] start_date_req を非稼働日補正: %s -> %s",
                        DEBUG_TASK_ID,
                        req_d,
                        prev_work,
                    )
                t["start_date_req"] = prev_work
    conflict_rows = collect_planning_conflicts_by_excel_row(tasks_df, ai_task_by_tid)
    try:
        apply_planning_sheet_conflict_styles(
            TASKS_INPUT_WORKBOOK,
            PLAN_INPUT_SHEET_NAME,
            len(tasks_df),
            conflict_rows,
        )
    except Exception as e:
        logging.warning(f"配台シート矛盾ハイライト適用をスキップ: {e}")

    if not task_queue:
        logging.warning(
            f"有効なタスクがありません。「{PLAN_INPUT_SHEET_NAME}」の「依頼NO」「工程名」「換算数量」、"
            "または完了区分・実出来高換算により残量が無い行のみの可能性があります。"
        )

    # 加工途中（実績あり）は優先。さらに同条件なら依頼NO（ハイフン後半の数値が小さい順）で先行。
    task_queue.sort(
        key=lambda x: (
            x.get("due_source_rank", 9),
            0 if x.get("has_done_deadline_override") else 1,
            0 if x.get("in_progress") else 1,
            x["priority"],
            0 if x["due_urgent"] else 1,
            x["due_basis_date"] or date.max,
            x["start_date_req"],
            _task_id_priority_key(x.get("task_id")),
            -resolve_need_required_op(
                x["machine"], x.get("machine_name", ""), x["task_id"], req_map, need_rules
            ),
        )
    )
    if DEBUG_TASK_ID:
        dbg_items = [t for t in task_queue if str(t.get("task_id", "")).strip() == DEBUG_TASK_ID]
        if dbg_items:
            t0 = dbg_items[0]
            logging.info(
                "DEBUG[task=%s] queue基準: start_date_req=%s due_basis=%s answer_due=%s specified_due=%s specified_due_ov=%s due_source=%s priority=%s in_progress=%s remark=%s",
                DEBUG_TASK_ID,
                t0.get("start_date_req"),
                t0.get("due_basis_date"),
                t0.get("answer_due_date"),
                t0.get("specified_due_date"),
                t0.get("specified_due_override"),
                t0.get("due_source"),
                t0.get("priority"),
                t0.get("in_progress"),
                t0.get("has_special_remark"),
            )
        else:
            logging.info("DEBUG[task=%s] task_queueに存在しません（完了/残量0/依頼NO不一致の可能性）。", DEBUG_TASK_ID)
    timeline_events = []

    # ---------------------------------------------------------
    # 日毎のスケジューリングループ
    # ---------------------------------------------------------
    for current_date in sorted_dates:
        daily_status = attendance_data[current_date]
        # 設備ごとの空き時刻（同一設備の同時並行割当を防止）
        machine_avail_dt = {}
        
        avail_dt = {}
        for m in members:
            if m in daily_status and daily_status[m]['is_working']:
                avail_dt[m] = daily_status[m]['start_dt']
        
        if not avail_dt:
            logging.info("DEBUG[day=%s] 稼働メンバー0のため割付スキップ", current_date)
            continue

        tasks_today = [t for t in task_queue if t['remaining_units'] > 0 and t['start_date_req'] <= current_date]
        pending_total = sum(1 for t in task_queue if t["remaining_units"] > 0)
        if not tasks_today:
            earliest_wait = min(
                [t["start_date_req"] for t in task_queue if t["remaining_units"] > 0],
                default=None,
            )
            logging.info(
                "DEBUG[day=%s] 割付対象タスク0件 pending_total=%s earliest_start_date_req=%s",
                current_date,
                pending_total,
                earliest_wait,
            )
        elif DEBUG_TASK_ID:
            has_dbg_today = any(str(t.get("task_id", "")).strip() == DEBUG_TASK_ID for t in tasks_today)
            if current_date.isoformat() == "2026-04-03" or has_dbg_today:
                logging.info(
                    "DEBUG[day=%s] avail_members=%s tasks_today=%s (task=%s 含む=%s)",
                    current_date,
                    len(avail_dt),
                    len(tasks_today),
                    DEBUG_TASK_ID,
                    has_dbg_today,
                )
        
        for task in tasks_today:
            if task['remaining_units'] <= 0: continue
            if DEBUG_TASK_ID and str(task.get("task_id", "")).strip() == DEBUG_TASK_ID:
                logging.info(
                    "DEBUG[task=%s] day=%s 開始判定: start_date_req=%s remaining_units=%s machine=%s",
                    DEBUG_TASK_ID,
                    current_date,
                    task.get("start_date_req"),
                    task.get("remaining_units"),
                    task.get("machine"),
                )
            if task.get("has_done_deadline_override"):
                logging.info(
                    "DEBUG[完了日指定] 依頼NO=%s 日付=%s start_date_req=%s due_basis=%s 指定納期(上書き)=%s 進捗=%s/%s",
                    task.get("task_id"),
                    current_date,
                    task.get("start_date_req"),
                    task.get("due_basis_date"),
                    task.get("specified_due_override"),
                    task.get("done_qty_reported"),
                    task.get("total_qty_m"),
                )
            
            machine = task['machine']
            ro = task.get("required_op")
            if ro is not None and int(ro) >= 1:
                req_num = int(ro)
            else:
                req_num = resolve_need_required_op(
                    machine, task.get("machine_name", ""), task["task_id"], req_map, need_rules
                )
            if global_priority_override.get("ignore_need_minimum"):
                req_num = 1

            # メンバー×設備スキル（parse_op_as_skill_cell: 小さい優先度ほど先にチーム候補へ採用）
            # skills 読込時に「機械名」単独キーへエイリアスするため、工程名+機械名が両方ある行では
            # 複合キー「工程名+機械名」のみを見る（別工程の同名機械の OP が流れ込まないようにする）。
            skill_meta_cache = {}
            machine_name = str(task.get("machine_name", "") or "").strip()
            machine_proc = str(machine or "").strip()
            _gpo = global_priority_override

            def skill_role_priority(mem):
                if _gpo.get("ignore_skill_requirements"):
                    return ("OP", 100)
                if mem not in skill_meta_cache:
                    srow = skills_dict.get(mem, {})
                    v = ""
                    if machine_proc and machine_name:
                        v = srow.get(f"{machine_proc}+{machine_name}", "")
                    elif machine_name:
                        v = srow.get(machine_name, "")
                    elif machine_proc:
                        v = srow.get(machine_proc, "")
                    skill_meta_cache[mem] = parse_op_as_skill_cell(v)
                return skill_meta_cache[mem]

            capable_members = [m for m in avail_dt.keys() if skill_role_priority(m)[0] in ("OP", "AS")]
            capable_members.sort(key=lambda mm: (skill_role_priority(mm)[1], mm))
            if task.get("has_done_deadline_override"):
                machine_free_dbg = machine_avail_dt.get(machine, datetime.combine(current_date, DEFAULT_START_TIME))
                logging.info(
                    "DEBUG[完了日指定] 依頼NO=%s 設備=%s req_num=%s capable_members=%s machine_free=%s",
                    task.get("task_id"),
                    machine,
                    req_num,
                    len(capable_members),
                    machine_free_dbg,
                )

            pref_raw = str(task.get("preferred_operator_raw") or "").strip()
            op_today = [m for m in capable_members if skill_role_priority(m)[0] == "OP"]
            pref_mem = _resolve_preferred_op_to_member(pref_raw, op_today) if pref_raw else None
            if pref_raw and pref_mem is None and op_today:
                logging.info(
                    "担当OP指名: 当日のOP候補に一致せず制約なし task=%s raw=%r",
                    task.get("task_id"),
                    pref_raw,
                )

            extra_max = resolve_need_surplus_extra_max(
                machine,
                machine_name,
                task["task_id"],
                surplus_map,
                need_rules,
            )
            max_team_size = min(req_num + extra_max, len(capable_members))
            if max_team_size < req_num:
                max_team_size = req_num
            rq_base = max(1, int(req_num))

            trace_assign = bool(TRACE_TEAM_ASSIGN_TASK_ID) and (
                str(task.get("task_id", "")).strip() == TRACE_TEAM_ASSIGN_TASK_ID
            )
            if trace_assign:
                logging.info(
                    "TRACE配台[%s] %s 工程/機械=%s / %s req_num=%s extra_max=%s → max_team=%s "
                    "capable(n=%s)=%s ignore_need1=%s ignore_skill=%s abolish=%s 担当OP指定=%r→%s",
                    task["task_id"],
                    current_date,
                    machine,
                    machine_name,
                    req_num,
                    extra_max,
                    max_team_size,
                    len(capable_members),
                    capable_members,
                    global_priority_override.get("ignore_need_minimum"),
                    global_priority_override.get("ignore_skill_requirements"),
                    global_priority_override.get("abolish_all_scheduling_limits"),
                    pref_raw,
                    pref_mem,
                )

            while task['remaining_units'] > 0:
                best_team = None
                best_info = {
                    "start_dt": datetime.max,
                    "units_today": 0,
                    "prio_sum": 10**9,
                }

                for tsize in range(req_num, max_team_size + 1):
                    sz_best_cand = None
                    sz_best_meta = None
                    if (
                        pref_mem is not None
                        and pref_mem in capable_members
                        and skill_role_priority(pref_mem)[0] == "OP"
                    ):
                        others = [m for m in capable_members if m != pref_mem]
                        if tsize == 1:
                            teams_iter = [(pref_mem,)]
                        elif len(others) >= tsize - 1:
                            teams_iter = [
                                tuple([pref_mem] + list(rest))
                                for rest in itertools.combinations(others, tsize - 1)
                            ]
                        else:
                            logging.info(
                                "担当OP指名: チーム人数を満たせないため指名を無視 task=%s size=%s raw=%r",
                                task.get("task_id"),
                                tsize,
                                pref_raw,
                            )
                            teams_iter = itertools.combinations(capable_members, tsize)
                    else:
                        teams_iter = itertools.combinations(capable_members, tsize)

                    for team in teams_iter:
                        op_list = [m for m in team if skill_role_priority(m)[0] == "OP"]
                        if not op_list:
                            continue

                        team_start = max(avail_dt[m] for m in team)
                        if not _gpo.get("abolish_all_scheduling_limits"):
                            # 同一設備は1時点で1タスクのみ（設備空き時刻を反映）
                            machine_free_dt = machine_avail_dt.get(
                                machine, datetime.combine(current_date, DEFAULT_START_TIME)
                            )
                            if team_start < machine_free_dt:
                                team_start = machine_free_dt
                            # 原反投入日と同日の開始は 10:30 以降
                            if task.get("same_day_raw_start_limit") and current_date == task["start_date_req"]:
                                min_start_dt = datetime.combine(
                                    current_date, task["same_day_raw_start_limit"]
                                )
                                if team_start < min_start_dt:
                                    team_start = min_start_dt
                            if current_date == task["start_date_req"] and task.get("earliest_start_time"):
                                min_user_t = datetime.combine(
                                    current_date, task["earliest_start_time"]
                                )
                                if team_start < min_user_t:
                                    team_start = min_user_t
                            # 当日は「マクロ実行した時刻」より前に開始できない
                            if current_date == macro_run_date and team_start < macro_now_dt:
                                team_start = macro_now_dt
                        team_end_limit = min(daily_status[m]['end_dt'] for m in team)

                        if team_start >= team_end_limit:
                            continue

                        team_breaks = []
                        for m in team:
                            team_breaks.extend(daily_status[m]['breaks_dt'])
                        team_breaks = merge_time_intervals(team_breaks)

                        avg_eff = sum(daily_status[m]['efficiency'] for m in team) / len(team)
                        if avg_eff <= 0:
                            avg_eff = 0.01
                        t_eff = parse_float_safe(task.get("task_eff_factor"), 1.0)
                        if t_eff <= 0:
                            t_eff = 1.0
                        # 追加増員による短縮は最大でも SURPLUS_TEAM_MAX_SPEEDUP_RATIO 程度（線形）
                        eff_time_per_unit = (
                            task["base_time_per_unit"]
                            / avg_eff
                            / t_eff
                            * _surplus_team_time_factor(rq_base, len(team), extra_max)
                        )

                        _, avail_mins, _ = calculate_end_time(team_start, 9999, team_breaks, team_end_limit)

                        units_can_do = int(avail_mins / eff_time_per_unit)
                        if units_can_do == 0:
                            continue

                        units_today = min(units_can_do, math.ceil(task['remaining_units']))
                        work_mins_needed = int(units_today * eff_time_per_unit)
                        actual_end_dt, _, _ = calculate_end_time(team_start, work_mins_needed, team_breaks, team_end_limit)

                        team_prio_sum = sum(skill_role_priority(m)[1] for m in team)
                        cand = _team_assignment_sort_tuple(
                            team, team_start, units_today, team_prio_sum
                        )
                        if trace_assign and (
                            sz_best_cand is None or cand < sz_best_cand
                        ):
                            sz_best_cand = cand
                            sz_best_meta = {
                                "team": team,
                                "team_start": team_start,
                                "units_today": units_today,
                                "prio_sum": team_prio_sum,
                                "eff_time_per_unit": eff_time_per_unit,
                            }
                        prev_best = (
                            None
                            if best_team is None
                            else _team_assignment_sort_tuple(
                                best_team,
                                best_info["start_dt"],
                                best_info["units_today"],
                                best_info["prio_sum"],
                            )
                        )
                        if best_team is None or cand < prev_best:
                            if pref_mem and pref_mem in op_list:
                                lead_op = pref_mem
                            else:
                                lead_op = min(op_list, key=lambda mm: (skill_role_priority(mm)[1], mm))
                            best_team = team
                            best_info = {
                                "start_dt": team_start,
                                "end_dt": actual_end_dt,
                                "op": lead_op,
                                "units_today": units_today,
                                "breaks": team_breaks,
                                "eff": avg_eff,
                                "prio_sum": team_prio_sum,
                            }

                    if trace_assign:
                        tid = task["task_id"]
                        if sz_best_meta is None:
                            logging.info(
                                "TRACE配台[%s] %s tsize=%s → この人数で成立するチームなし",
                                tid,
                                current_date,
                                tsize,
                            )
                        else:
                            sm = sz_best_meta
                            _tk = (
                                "(-人数, start, -units, prio_sum)"
                                if TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF
                                else "(start, -units, prio_sum)"
                            )
                            logging.info(
                                "TRACE配台[%s] %s tsize=%s 人数内最良: members=%s "
                                "start=%s units_today=%s prio_sum=%s eff_t/unit=%.6f "
                                "比較タプル=%s ※辞書式で小さい方が採用",
                                tid,
                                current_date,
                                tsize,
                                sm["team"],
                                sm["team_start"],
                                sm["units_today"],
                                sm["prio_sum"],
                                sm["eff_time_per_unit"],
                                _tk,
                            )

                if trace_assign and best_team is not None:
                    logging.info(
                        "TRACE配台[%s] %s ★採用 n=%s members=%s start=%s units_today=%s prio_sum=%s",
                        task["task_id"],
                        current_date,
                        len(best_team),
                        best_team,
                        best_info["start_dt"],
                        best_info["units_today"],
                        best_info["prio_sum"],
                    )
                    if len(best_team) == 1 and max_team_size > req_num:
                        if TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF:
                            logging.info(
                                "TRACE配台[%s] %s 1人採用（余剰優先モード）: より大きい人数で有効なチームなし"
                                "（OPが1人しかいない・組合せで時間内0単位・開始>=終了等）。",
                                task["task_id"],
                                current_date,
                            )
                        else:
                            logging.info(
                                "TRACE配台[%s] %s 1人採用の典型要因: team_start=max(メンバー空き) で人数が増えると開始が遅れやすい；"
                                "増員の速度効果が小さい(SURPLUS_TEAM_MAX_SPEEDUP_RATIO)と units_today がほぼ同じで開始の早い1人が勝つ。"
                                "必要人数が1(シート上書き・メイン再優先・need)のときは max_team=1 もあり得ます。",
                                task["task_id"],
                                current_date,
                            )

                if best_team:
                    sub_members = [m for m in best_team if m != best_info["op"]]
                    done_units = best_info["units_today"]
                    
                    total_u = math.ceil(task['total_qty_m'] / task['unit_m']) if task['unit_m'] else 0
                    rem_u_before = math.ceil(task['remaining_units'])
                    already_done = total_u - rem_u_before
                    
                    # 「マクロ実行時点」の完了率（予定の進捗ではなく、実加工数ベース）
                    try:
                        tot_qty = parse_float_safe(task.get('total_qty_m'), 0.0)
                        done_qty = parse_float_safe(task.get('done_qty_reported'), 0.0)
                        if tot_qty > 0:
                            pct_macro = max(0, min(100, int(round((done_qty / tot_qty) * 100))))
                        else:
                            pct_macro = 0
                    except Exception:
                        pct_macro = 0
                    
                    _te_disp = parse_float_safe(task.get("task_eff_factor"), 1.0)
                    if _te_disp <= 0:
                        _te_disp = 1.0
                    timeline_events.append({
                        "date": current_date, "task_id": task['task_id'], "machine": machine,
                        "op": best_info["op"], "sub": ", ".join(sub_members),
                        "start_dt": best_info["start_dt"], "end_dt": best_info["end_dt"],
                        "breaks": best_info["breaks"], "units_done": done_units,
                        "already_done_units": already_done,
                        "total_units": total_u,
                        "pct_macro": pct_macro,
                        "eff_time_per_unit": task["base_time_per_unit"]
                        / best_info["eff"]
                        / _te_disp
                        * _surplus_team_time_factor(rq_base, len(best_team), extra_max),
                        "unit_m": task['unit_m']
                    })
                    
                    task['remaining_units'] -= done_units
                    op_main = (best_info.get("op") or "").strip()
                    subs_part = ",".join(
                        s.strip() for s in sub_members if s and str(s).strip()
                    )
                    team_s = f"{op_main}, {subs_part}" if subs_part else op_main
                    task["assigned_history"].append(
                        {
                            "date": current_date.strftime("%m/%d"),
                            "team": team_s,
                            "done_m": int(done_units * task["unit_m"]),
                        }
                    )
                    
                    for m in best_team:
                        avail_dt[m] = best_info["end_dt"]
                    if not _gpo.get("abolish_all_scheduling_limits"):
                        machine_avail_dt[machine] = best_info["end_dt"]
                else:
                    if task.get("has_done_deadline_override"):
                        logging.info(
                            "DEBUG[完了日指定] 依頼NO=%s 日付=%s は割当不可（要員/設備空き条件でチーム不成立）。remaining_units=%s",
                            task.get("task_id"),
                            current_date,
                            task.get("remaining_units"),
                        )
                    break

    # タイムラインを日付別にインデックス化し、サブメンバー一覧を事前解析（以降の出力ループを高速化）
    events_by_date = defaultdict(list)
    for e in timeline_events:
        e["subs_list"] = [s.strip() for s in e["sub"].split(",")] if e.get("sub") else []
        events_by_date[e["date"]].append(e)

    # =========================================================
    # 4. Excel出力 (メイン計画)
    # =========================================================
    dated_output_dir = get_dated_output_dir(base_now_dt)
    output_filename = os.path.join(
        dated_output_dir, f'production_plan_multi_day_{base_now_dt.strftime("%Y%m%d_%H%M%S")}.xlsx'
    )
    all_eq_rows = []
    
    eq_empty_cols = {}
    for eq in equipment_list:
        eq_empty_cols[eq] = ""
        eq_empty_cols[f"{eq}進度"] = ""
    
    for d in sorted_dates:
        d_start = datetime.combine(d, DEFAULT_START_TIME)
        d_end = datetime.combine(d, DEFAULT_END_TIME)
        events_today = events_by_date[d]
        machine_to_events = defaultdict(list)
        for ev in events_today:
            machine_to_events[ev["machine"]].append(ev)
        
        is_anyone_working = any(daily_status['is_working'] for daily_status in attendance_data[d].values())
        if not events_today and not is_anyone_working: continue
        
        all_eq_rows.append({"日時帯": f"■ {d.strftime('%Y/%m/%d (%a)')} ■", **eq_empty_cols})
        
        curr_grid = d_start
        while curr_grid < d_end:
            next_grid = curr_grid + timedelta(minutes=10)
            if next_grid > d_end: next_grid = d_end
            
            mid_t = curr_grid + (next_grid - curr_grid) / 2
            row_data = {"日時帯": f"{curr_grid.strftime('%H:%M')}-{next_grid.strftime('%H:%M')}"}
            
            for eq in equipment_list:
                eq_text = ""
                progress_text = ""
                active_ev = None
                for ev in machine_to_events.get(eq, ()):
                    if ev["start_dt"] <= mid_t < ev["end_dt"]:
                        active_ev = ev
                        break
                
                if active_ev:
                    if any(b_s <= mid_t < b_e for b_s, b_e in active_ev['breaks']):
                        eq_text = "休憩"
                    else:
                        elapsed = get_actual_work_minutes(active_ev['start_dt'], min(next_grid, active_ev['end_dt']), active_ev['breaks'])
                        block_done_now = min(int(elapsed / active_ev['eff_time_per_unit']), active_ev['units_done'])
                        
                        cumulative_done = active_ev['already_done_units'] + block_done_now
                        total_u = active_ev['total_units']
                        
                        sub_text = f" 補:{active_ev['sub']}" if active_ev['sub'] else ""
                        eq_text = f"[{active_ev['task_id']}] 主:{active_ev['op']}{sub_text}"
                        progress_text = f"{cumulative_done}/{total_u}R"
                        
                row_data[eq] = eq_text
                row_data[f"{eq}進度"] = progress_text
            
            all_eq_rows.append(row_data)
            curr_grid = next_grid
        all_eq_rows.append({"日時帯": "", **eq_empty_cols})
        
    df_eq_schedule = pd.DataFrame(all_eq_rows)

    task_results = []
    max_history_len = max([len(t['assigned_history']) for t in task_queue] + [0])
    
    # ステータス（配台の可否・残）：完了相当=配台可／未割当=配台不可／一部のみ=配台残
    for t in sorted(task_queue, key=_result_task_sheet_sort_key):
        status = "配台可" if t['remaining_units'] <= 0 else "配台残"
        if not t['assigned_history'] and t['remaining_units'] > 0:
            status = "配台不可"
        
        total_r = int(t['total_qty_m'] / t['unit_m']) if t['unit_m'] else 0
        rem_r = int(t['remaining_units'])
        
        ans_s = t["answer_due_date"].strftime("%Y/%m/%d") if t.get("answer_due_date") else ""
        spec_s = t["specified_due_date"].strftime("%Y/%m/%d") if t.get("specified_due_date") else ""
        basis_s = t["due_basis_date"].strftime("%Y/%m/%d") if t.get("due_basis_date") else ""
        start_req = t["start_date_req"]
        start_req_s = start_req.strftime("%Y/%m/%d") if hasattr(start_req, "strftime") else str(start_req)
        rov = t.get("required_op")
        # 列順: A=ステータス → タスクID/工程/機械/優先度 → 履歴1..n → その他 → 最後に特別指定_AI
        row_status = {"ステータス": status}
        row_core = {
            "タスクID": t['task_id'],
            "工程名": t['machine'],
            "機械名": t.get("machine_name", ""),
            "優先度": t.get("priority", 999),
        }
        row_history = {}
        for i in range(max_history_len):
            if i < len(t['assigned_history']):
                h = t['assigned_history'][i]
                done_r = int(h['done_m'] / t['unit_m']) if t['unit_m'] else 0
                row_history[f"履歴{i+1}"] = (
                    f"・【{h['date']}】：{done_r}R ({h['done_m']}m) 担当[{h['team']}]"
                )
            else:
                row_history[f"履歴{i+1}"] = ""

        try:
            tot_qty = parse_float_safe(t.get("total_qty_m"), 0.0)
            done_qty = parse_float_safe(t.get("done_qty_reported"), 0.0)
            pct_macro = max(0, min(100, int(round((done_qty / tot_qty) * 100)))) if tot_qty > 0 else 0
        except Exception:
            pct_macro = 0

        row_tail = {
            "必要OP(上書)": rov if rov is not None else "",
            "タスク効率": parse_float_safe(t.get("task_eff_factor"), 1.0),
            "加工途中": "はい" if t.get("in_progress") else "いいえ",
            "特別指定あり": "はい" if t.get("has_special_remark") else "いいえ",
            "担当OP指名": (t.get("preferred_operator_raw") or "")[:120],
            "回答納期": ans_s,
            "指定納期": spec_s,
            "計画基準納期": basis_s,
            "納期緊急": "はい" if t.get("due_urgent") else "いいえ",
            "加工開始日": start_req_s,
            "総加工量": f"{total_r}R ({t['total_qty_m']}m)",
            "残加工量": f"{rem_r}R ({int(t['remaining_units'] * t['unit_m'])}m)",
            "完了率(実行時点)": f"{pct_macro}%",
        }
        row_ai_last = {"特別指定_AI": (t.get("task_special_ai_note") or "")[:300]}
        row_data = {**row_status, **row_core, **row_history, **row_tail, **row_ai_last}
        task_results.append(row_data)
        
    cal_rows = []
    for d in sorted_dates:
        for m in members:
            if m in attendance_data[d]:
                data = attendance_data[d][m]
                if data['is_working']:
                    cal_end = _calendar_display_clock_out_for_calendar_sheet(data, d)
                    end_disp = cal_end if cal_end is not None else data['end_dt']
                    clock_out_s = end_disp.strftime("%H:%M")
                else:
                    clock_out_s = "休"
                cal_rows.append({
                    "日付": d,
                    "メンバー": m,
                    "出勤": data['start_dt'].strftime("%H:%M") if data['is_working'] else "休",
                    "退勤": clock_out_s,
                    "効率": data['efficiency'],
                    "備考": data['reason'],
                })

    utilization_data = []
    for d in sorted_dates:
        row_data = {"年月日": d.strftime("%Y/%m/%d (%a)")}
        # その日のイベントからメンバー別作業分を一括集計（全メンバー×全イベントの二重ループを避ける）
        member_worked_mins = defaultdict(int)
        for ev in events_by_date[d]:
            mins = get_actual_work_minutes(ev["start_dt"], ev["end_dt"], ev["breaks"])
            member_worked_mins[ev["op"]] += mins
            for s in ev["subs_list"]:
                if s:
                    member_worked_mins[s] += mins
        for m in members:
            if m in attendance_data[d] and attendance_data[d][m]['is_working']:
                default_start = datetime.combine(d, DEFAULT_START_TIME)
                default_end = datetime.combine(d, DEFAULT_END_TIME)
                
                actual_start = attendance_data[d][m]['start_dt']
                actual_end = attendance_data[d][m]['end_dt']
                clip_start = max(actual_start, default_start)
                clip_end = min(actual_end, default_end)
                
                if clip_start >= clip_end:
                    total_avail_mins = 0
                else:
                    breaks_dt = attendance_data[d][m]['breaks_dt']
                    total_avail_mins = get_actual_work_minutes(clip_start, clip_end, breaks_dt)
                
                if total_avail_mins <= 0:
                    row_data[m] = "0.0%"
                    continue
                
                worked_mins = member_worked_mins.get(m, 0)
                ratio = (worked_mins / total_avail_mins) * 100
                row_data[m] = f"{ratio:.1f}% ({worked_mins}/{total_avail_mins}分)"
            else:
                row_data[m] = "休"
        utilization_data.append(row_data)
        
    df_utilization = pd.DataFrame(utilization_data)

    _usage_txt = build_gemini_usage_summary_text()
    if _usage_txt:
        ai_log_data["Gemini_トークン・料金サマリ"] = _usage_txt[:50000]

    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        df_eq_schedule.to_excel(writer, sheet_name='結果_設備毎の時間割', index=False)
        pd.DataFrame(cal_rows).to_excel(writer, sheet_name='結果_カレンダー(出勤簿)', index=False)
        df_utilization.to_excel(writer, sheet_name='結果_メンバー別作業割合', index=False)
        df_tasks = pd.DataFrame(task_results)
        df_tasks, task_column_order, _, vis_map = apply_result_task_sheet_column_order(
            df_tasks, max_history_len
        )
        pd.DataFrame(
            {
                "列名": task_column_order,
                "表示": [bool(vis_map.get(c, True)) for c in task_column_order],
            }
        ).to_excel(writer, sheet_name=COLUMN_CONFIG_SHEET_NAME, index=False)
        df_tasks.to_excel(writer, sheet_name=RESULT_TASK_SHEET_NAME, index=False)
        pd.DataFrame(list(ai_log_data.items()), columns=["項目", "内容"]).to_excel(writer, sheet_name='結果_AIログ', index=False)

        _write_results_equipment_gantt_sheet(
            writer,
            timeline_events,
            equipment_list,
            sorted_dates,
            attendance_data,
            data_extract_dt_str,
            base_now_dt,
        )

        for sheet_name, ws_out in writer.sheets.items():
            if sheet_name == RESULT_SHEET_GANTT_NAME:
                continue
            _apply_output_font_to_result_sheet(ws_out)

        ws_cfg = writer.sheets[COLUMN_CONFIG_SHEET_NAME]
        _add_column_config_sheet_helpers(ws_cfg, len(task_column_order))

        worksheet_tasks = writer.sheets[RESULT_TASK_SHEET_NAME]
        max_col = worksheet_tasks.max_column
        for row in worksheet_tasks.iter_rows(min_row=1, max_row=worksheet_tasks.max_row, max_col=max_col):
            for cell in row:
                cell.alignment = Alignment(wrap_text=False, vertical="top")

        _apply_result_task_sheet_column_visibility(
            worksheet_tasks, list(df_tasks.columns), vis_map
        )

        _apply_result_task_history_rich_text(worksheet_tasks, list(df_tasks.columns))

        # 未スケジュール行（配台不可・配台残）を目立たせる
        status_col_idx = None
        for col_idx, col_name in enumerate(df_tasks.columns, 1):
            if str(col_name) == "ステータス":
                status_col_idx = col_idx
                break
        if status_col_idx is not None:
            unscheduled_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            for r in range(2, worksheet_tasks.max_row + 1):
                st_val = worksheet_tasks.cell(row=r, column=status_col_idx).value
                st = str(st_val).strip() if st_val is not None else ""
                if st in ("配台不可", "配台残"):
                    for c in range(1, max_col + 1):
                        worksheet_tasks.cell(row=r, column=c).fill = unscheduled_fill

    try:
        _apply_excel_date_columns_date_only_display(
            output_filename, "結果_カレンダー(出勤簿)", frozenset({"日付"})
        )
    except Exception as e:
        logging.warning(f"結果_カレンダー(出勤簿)の日付列表示整形: {e}")

    logging.info(f"完了: '{output_filename}' を生成しました。")

    # =========================================================
    # 5. ★追加: メンバー毎の行動スケジュール (別ファイル) 出力
    # =========================================================
    member_output_filename = os.path.join(
        dated_output_dir, f'member_schedule_{base_now_dt.strftime("%Y%m%d_%H%M%S")}.xlsx'
    )
    
    # 時間帯は全メンバー共通で1回だけ生成（メンバー数分の重複計算を避ける）
    time_labels = []
    time_grids = []
    curr_dt = datetime.combine(run_date, DEFAULT_START_TIME)
    end_dt_grid = datetime.combine(run_date, DEFAULT_END_TIME)
    while curr_dt < end_dt_grid:
        next_dt = curr_dt + timedelta(minutes=10)
        if next_dt > end_dt_grid:
            next_dt = end_dt_grid
        time_labels.append(f"{curr_dt.strftime('%H:%M')}-{next_dt.strftime('%H:%M')}")
        time_grids.append((curr_dt.time(), next_dt.time()))
        curr_dt = next_dt
    
    with pd.ExcelWriter(member_output_filename, engine='openpyxl') as member_writer:
        for m in members:
            # 各行の辞書を初期化
            m_schedule = {t_label: {"時間帯": t_label} for t_label in time_labels}
            
            # 各日付のスケジュールを列として埋めていく
            for d in sorted_dates:
                d_str = d.strftime("%m/%d (%a)")
                
                # 全日非勤務: 年休（カレンダー *）は『年休』、工場休日などは『休』
                if m not in attendance_data[d] or not attendance_data[d][m]['is_working']:
                    off_label = _member_schedule_full_day_off_label(
                        attendance_data[d].get(m) if m in attendance_data[d] else None
                    )
                    for t_label in time_labels:
                        m_schedule[t_label][d_str] = off_label
                    continue
                
                daily_info = attendance_data[d][m]
                d_start_dt = daily_info['start_dt']
                d_end_dt = daily_info['end_dt']
                breaks_dt = daily_info['breaks_dt']
                
                events_today = events_by_date[d]
                
                for i, (t_start, t_end) in enumerate(time_grids):
                    t_label = time_labels[i]
                    
                    # 判定用の中間時刻を計算
                    grid_start_dt = datetime.combine(d, t_start)
                    grid_end_dt = datetime.combine(d, t_end)
                    grid_mid_dt = grid_start_dt + (grid_end_dt - grid_start_dt) / 2
                    
                    text = ""
                    if grid_mid_dt < d_start_dt or grid_mid_dt >= d_end_dt:
                        text = _member_schedule_off_shift_label(
                            d, grid_mid_dt, d_start_dt, d_end_dt, daily_info.get("reason")
                        )
                    else:
                        br_txt = _member_schedule_break_cell_label(
                            grid_mid_dt, breaks_dt, d_end_dt, daily_info.get("reason")
                        )
                        if br_txt is not None:
                            text = br_txt
                    if text == "":
                        # 該当するタスクを探す（subs_list は事前解析済み）
                        active_ev = next((e for e in events_today if e['start_dt'] <= grid_mid_dt < e['end_dt'] and (e['op'] == m or m in e.get('subs_list', []))), None)
                        if active_ev:
                            role = "主" if active_ev['op'] == m else "補"
                            text = f"[{active_ev['task_id']}] {active_ev['machine']}({role})"
                        else:
                            text = "" # 何も割り当てられていない空き時間
                    
                    m_schedule[t_label][d_str] = text
                    
            # データフレーム化してシートに書き込み
            df_m = pd.DataFrame(list(m_schedule.values()))
            cols = ["時間帯"] + [d.strftime("%m/%d (%a)") for d in sorted_dates]
            df_m = df_m[[c for c in cols if c in df_m.columns]]
            df_m.to_excel(member_writer, sheet_name=m, index=False)
            
            # --- 既定フォント・罫線・見出し背景（列幅は VBA 取り込み時の AutoFit） ---
            worksheet = member_writer.sheets[m]
            _apply_output_font_to_result_sheet(worksheet)
            header_fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
            for cell in worksheet[1]:
                cell.fill = header_fill
                
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, max_col=worksheet.max_column):
                for cell in row:
                    cell.border = thin_border
                    
    logging.info(f"完了: 個人別スケジュールを '{member_output_filename}' に出力しました。")
    _try_write_main_sheet_gemini_usage_summary("段階2")