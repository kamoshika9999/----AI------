# -*- coding: utf-8 -*-
"""
xlwings の「Show Console」＋ RunPython で段階1/2を動かす（cmd.exe 不要）。

前提
----
- Excel に **xlwings アドイン**を入れ、リボンの **Show Console** にチェック。
- 本モジュールは **マクロブック（.xlsm）と同じフォルダ直下の** ``python/`` **配下**
  （例: ``.../テストコード/python/xlwings_console_runner.py``）に置く想定。
  ``import xlwings_console_runner`` だけでは ``sys.path`` に ``python`` が入らないことが多いため、
  **VBA では ``runpy.run_path`` で本ファイルを直接実行する形式を推奨**（``配台デバッグ_VBA.txt`` の
  ``xlwings.RunPython "import os, runpy, ...`` と同じ）。補助として ``xlwings.conf.json`` の PYTHONPATH も可。
- VBA（xlwings 参照 ON）の推奨呼び出し::

    xlwings.RunPython "import os, runpy, xlwings as xw; wb=xw.Book.caller(); p=os.path.join(os.path.dirname(str(wb.fullname)), 'python', 'xlwings_console_runner.py'); ns=runpy.run_path(p); ns['run_stage1_for_xlwings']()"

  本番の ``段階1_コア実行`` / ``段階2_コア実行`` は定数 ``STAGE12_USE_XLWINGS_RUNPYTHON`` で上記経路を選べる。
  終了コードは ``log/stage_vba_exitcode.txt`` に1行で書く（VBA が cmd 経路と同様に読む）。

logging は planning_core が ``StreamHandler(sys.stdout)`` を付けているため、
print と同様に xlwings のコンソールへ流れる。

注意
----
- ``planning_core`` は import 時に TASK_INPUT_WORKBOOK 依存の初期化があるため、
  本モジュールでは **都度 ``planning_core`` を sys.modules から外して再 import** する。
- cmd 経路は ``task_extract_stage1.py`` が planning_core より前に ``execution_log.txt`` を作る。
  **本モジュールも** ``run_stage1_for_xlwings`` / ``run_stage2_for_xlwings`` で import 前に同ファイルへ1行追記し、
  例外時はトレースバックを追記する（VBA が LOG シート用にファイル存在を前提にできる）。
- ``runpy.run_path`` 実行時は **sys.path 先頭に本ファイルの親（python フォルダ）が入らない** ことがあり、
  ``import planning_core`` が ``ModuleNotFoundError`` になる。読み込み直後に ``_ensure_this_python_dir_on_syspath`` で補う。
"""
from __future__ import annotations

import logging
import os
import sys
import traceback
from datetime import datetime

STAGE_VBA_EXIT_CODE_FILE = "stage_vba_exitcode.txt"


def _ensure_this_python_dir_on_syspath() -> None:
    """planning_core 等と同じディレクトリを sys.path に必ず含める（run_path 経路対策）。"""
    here = os.path.dirname(os.path.abspath(__file__))
    here_n = os.path.normcase(here)
    for _p in sys.path:
        try:
            if os.path.normcase(os.path.abspath(_p)) == here_n:
                return
        except (OSError, ValueError):
            continue
    sys.path.insert(0, here)


_ensure_this_python_dir_on_syspath()


def _write_stage_vba_exit_code(code: int) -> None:
    """VBA の ReadStageVbaExitCodeFromFile と整合する1行ファイル。"""
    try:
        logd = os.path.join(os.getcwd(), "log")
        os.makedirs(logd, exist_ok=True)
        p = os.path.join(logd, STAGE_VBA_EXIT_CODE_FILE)
        with open(p, "w", encoding="utf-8", newline="\n") as f:
            f.write(str(int(code)))
            f.flush()
            os.fsync(f.fileno())
    except OSError:
        pass


def _purge_planning_core_modules() -> None:
    for k in list(sys.modules.keys()):
        if k == "planning_core" or k.startswith("planning_core."):
            del sys.modules[k]


def _prepare_from_caller_book() -> str:
    import xlwings as xw

    wb = xw.Book.caller()
    path = os.path.abspath(str(wb.fullname))
    root = os.path.dirname(path)
    os.chdir(root)
    os.environ["TASK_INPUT_WORKBOOK"] = path
    return path


def _execution_log_path() -> str:
    return os.path.join(os.getcwd(), "log", "execution_log.txt")


def _append_execution_log_line(level: str, msg: str) -> None:
    """
    cmd 経路の task_extract_stage1 と同様、planning_core より前に log を確保する（VBA が LOG シートへ読む）。
    """
    path = _execution_log_path()
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"{ts} - {level} - {msg}\n"
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "a", encoding="utf-8-sig", newline="\n") as f:
            f.write(line)
            f.flush()
            try:
                os.fsync(f.fileno())
            except OSError:
                pass
    except OSError:
        return
    try:
        import xlwings_splash_log as xsl

        if xsl.enabled():
            xsl.append_formatted_line(line)
    except Exception:
        pass


def _append_execution_log_traceback(title: str) -> None:
    """except ブロック内で呼ぶ（format_exc が有効なとき）。"""
    _append_execution_log_line("ERROR", title)
    tb = traceback.format_exc()
    try:
        path = _execution_log_path()
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "a", encoding="utf-8-sig", newline="\n") as f:
            f.write(tb)
            if not tb.endswith("\n"):
                f.write("\n")
            f.flush()
            try:
                os.fsync(f.fileno())
            except OSError:
                pass
    except OSError:
        pass


def run_stage1_for_xlwings() -> int:
    """
    段階1（run_stage1_extract）。戻り値: 0=成功, 1=失敗, 2=caller 不備。
    終了時に log/stage_vba_exitcode.txt を常に更新する。
    """
    rc = 1
    try:
        try:
            _prepare_from_caller_book()
        except Exception:
            logging.exception(
                "xlwings: Book.caller() を取得できません。"
                " マクロブック上のボタンから RunPython してください。"
            )
            rc = 2
            return rc
        _append_execution_log_line(
            "INFO",
            "段階1: xlwings run_stage1_for_xlwings 開始（planning_core 読み込み前）",
        )
        _purge_planning_core_modules()
        try:
            import planning_core as pc

            ok = pc.run_stage1_extract()
            rc = 0 if ok else 1
        except SystemExit as e:
            c = e.code
            if c is None:
                rc = 0
            elif isinstance(c, int):
                rc = 0 if c == 0 else c
            else:
                rc = 1
        except Exception:
            logging.exception("xlwings: 段階1で未捕捉例外")
            _append_execution_log_traceback("xlwings: 段階1で未捕捉例外")
            rc = 1
        return rc
    finally:
        _write_stage_vba_exit_code(rc)


def run_stage2_for_xlwings() -> int:
    """
    段階2（generate_plan）。戻り値: 0=正常終了, 1=例外, 2=caller 不備。
    終了時に log/stage_vba_exitcode.txt を常に更新する。
    """
    rc = 1
    try:
        try:
            _prepare_from_caller_book()
        except Exception:
            logging.exception("xlwings: Book.caller() を取得できません。")
            rc = 2
            return rc
        _append_execution_log_line(
            "INFO",
            "段階2: xlwings run_stage2_for_xlwings 開始（planning_core 読み込み前）",
        )
        _purge_planning_core_modules()
        try:
            import planning_core as pc

            pc.generate_plan()
            rc = 0
        except SystemExit as e:
            c = e.code
            if c is None:
                rc = 0
            elif isinstance(c, int):
                rc = 0 if c == 0 else c
            else:
                rc = 1
        except Exception:
            logging.exception("xlwings: 段階2で未捕捉例外")
            _append_execution_log_traceback("xlwings: 段階2で未捕捉例外")
            rc = 1
        return rc
    finally:
        _write_stage_vba_exit_code(rc)


def run_refresh_dispatch_roll_trace_sheet_for_xlwings() -> int:
    """
    ``log/dispatch_roll_trace.jsonl``（または DISPATCH_ROLL_TRACE_JSONL）を
    マクロブック内シート「配台_ロールトレース」に展開する。段階2実行後に任意で呼ぶ。

    RunPython 例::
        import xlwings_console_runner as _xw; _xw.run_refresh_dispatch_roll_trace_sheet_for_xlwings()
    """
    try:
        _prepare_from_caller_book()
    except Exception:
        logging.exception("xlwings: Book.caller() を取得できません。")
        return 2
    try:
        import planning_dispatch_debug as dd

        return dd.run_refresh_roll_trace_sheet_for_xlwings()
    except Exception:
        logging.exception("xlwings: ロールトレースシート更新")
        return 1


def run_stage2_with_roll_trace_jsonl_then_refresh_sheet_for_xlwings() -> int:
    """
    1 ロールごとの JSONL トレースを有効にして段階2を実行し、続けて「配台_ロールトレース」シートを更新する。

    環境変数 DISPATCH_ROLL_TRACE_JSONL が未設定のときは ``log/dispatch_roll_trace.jsonl`` を使用。
    早期終了のみ試す場合は DISPATCH_DEBUG_STOP_AFTER_ROLLS=1 等を xlwings.conf / システム環境に設定。

    RunPython 例::
        import xlwings_console_runner as _xw; _xw.run_stage2_with_roll_trace_jsonl_then_refresh_sheet_for_xlwings()
    """
    try:
        _prepare_from_caller_book()
    except Exception:
        logging.exception("xlwings: Book.caller() を取得できません。")
        return 2
    # setdefault は「キー未存在」のみ。空文字の DISPATCH_ROLL_TRACE_JSONL では既定が入らないため明示する。
    _dj = (os.environ.get("DISPATCH_ROLL_TRACE_JSONL") or "").strip()
    if not _dj:
        os.environ["DISPATCH_ROLL_TRACE_JSONL"] = "log/dispatch_roll_trace.jsonl"
    _purge_planning_core_modules()
    try:
        import planning_core as pc

        pc.generate_plan()
    except SystemExit as e:
        c = e.code
        if c is None:
            pass
        elif isinstance(c, int):
            if c != 0:
                return c
        else:
            return 1
    except Exception:
        logging.exception("xlwings: 段階2（ロールトレース付き）")
        return 1
    try:
        import planning_dispatch_debug as dd

        print(dd.refresh_roll_trace_sheet_from_jsonl())
    except Exception:
        logging.exception(
            "xlwings: ロールトレースシート更新のみ失敗（結果ブックは出力済みの可能性）"
        )
    return 0
