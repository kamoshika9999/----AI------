# -*- coding: utf-8 -*-
"""
段階1: 加工計画DATA から未完了タスクを抽出し output/plan_input_tasks.xlsx を生成する。
マクロで当ファイルをブックへ取り込み、「配台計画_タスク入力」で特別指定を編集した後、段階2を実行する。

環境: planning_core と同様、openpyxl（閉じたブックの I/O）と、保存ロック時の xlwings 同期を使用する。
初回はリポジトリの requirements.txt を pip し、Excel デスクトップ版を用意すること。
"""
import ctypes
import os
import sys
import traceback
from datetime import datetime

# planning_core は import 途中（FileHandler より前）で落ちると execution_log が作られない。
# VBA は「ログ無し」を判定するため、読み込み前に必ず log/execution_log.txt を用意する。
def _repo_root_for_stage1() -> str:
    d = os.path.dirname(os.path.abspath(__file__))
    return os.path.dirname(d) if os.path.basename(d).lower() == "python" else d


def _execution_log_paths_stage1() -> list[str]:
    """VBA が読む log と同じ候補（マクロブックのフォルダを最優先）。"""
    paths: list[str] = []
    wb = (os.environ.get("TASK_INPUT_WORKBOOK") or "").strip()
    if wb:
        paths.append(
            os.path.join(os.path.dirname(os.path.abspath(wb)), "log", "execution_log.txt")
        )
    paths.append(
        os.path.join(_repo_root_for_stage1(), "log", "execution_log.txt")
    )
    seen: set[str] = set()
    out: list[str] = []
    for p in paths:
        key = os.path.normcase(os.path.abspath(p))
        if key not in seen:
            seen.add(key)
            out.append(p)
    return out


def _append_execution_log_line(level: str, msg: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"{ts} - {level} - {msg}\n"
    last_err: OSError | None = None
    for path in _execution_log_paths_stage1():
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "a", encoding="utf-8-sig", newline="\n") as f:
                f.write(line)
                f.flush()
                try:
                    os.fsync(f.fileno())
                except OSError:
                    pass
            try:
                import xlwings_splash_log as _xsl

                if _xsl.enabled():
                    _xsl.append_formatted_line(line)
            except Exception:
                pass
            return
        except OSError as ex:
            last_err = ex
    print(f"log/execution_log.txt へ書けません: {last_err}", file=sys.stderr)


os.chdir(os.path.dirname(os.path.abspath(__file__)))

if os.name == "nt":
    hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    if hwnd:
        ctypes.windll.user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 3)

try:
    _append_execution_log_line(
        "INFO", "段階1: task_extract_stage1 起動（planning_core 読み込み前）"
    )
except OSError as ex:
    print(f"log/execution_log.txt を開けません: {ex}", file=sys.stderr)

try:
    import planning_core as pc
except Exception:
    _err_head = (
        f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - ERROR - "
        "planning_core の import に失敗しました\n"
    )
    _tb = traceback.format_exc()
    for path in _execution_log_paths_stage1():
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "a", encoding="utf-8-sig", newline="\n") as f:
                f.write(_err_head)
                f.write(_tb)
                f.flush()
                try:
                    os.fsync(f.fileno())
                except OSError:
                    pass
        except OSError:
            pass
    try:
        import xlwings_splash_log as _xsl

        if _xsl.enabled():
            _xsl.append_formatted_line(_err_head + _tb)
    except Exception:
        pass
    traceback.print_exc()
    sys.exit(1)


def main():
    if not pc.TASKS_INPUT_WORKBOOK:
        print("TASK_INPUT_WORKBOOK が未設定です。VBA からマクロ実行してください。", file=sys.stderr)
        sys.exit(2)
    try:
        ok = pc.run_stage1_extract()
    except pc.PlanningValidationError as e:
        msg = str(e).strip() or "マスタ skills の検証で中断しました。"
        if not os.path.isfile(pc.stage2_blocking_message_path):
            pc._write_stage2_blocking_message(msg)
        print(msg, file=sys.stderr)
        sys.exit(3)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
