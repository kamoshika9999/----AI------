# -*- coding: utf-8 -*-
"""
段階2: マクロブックの「配台計画_タスク入力」を読み、特別指定・備考AI反映後に計画シミュレーションを実行する。
"""
import os
import sys
import ctypes

os.chdir(os.path.dirname(os.path.abspath(__file__)))

if os.name == "nt":
    hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    if hwnd:
        ctypes.windll.user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 3)

import planning_core as pc


def main():
    try:
        pc.generate_plan()
    except pc.PlanningValidationError as e:
        msg = str(e).strip() or "配台計画の検証で中断しました。"
        if not os.path.isfile(pc.stage2_blocking_message_path):
            pc._write_stage2_blocking_message(msg)
        print(msg, file=sys.stderr)
        sys.exit(3)


if __name__ == "__main__":
    main()
