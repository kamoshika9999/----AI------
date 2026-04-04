# -*- coding: utf-8 -*-
"""
後方互換: 計画シミュレーション（段階2）へ委譲します。
タスク抽出のみ行う場合は task_extract_stage1.py を実行してください。
"""
import os
import ctypes

os.chdir(os.path.dirname(os.path.abspath(__file__)))

if os.name == "nt":
    hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    if hwnd:
        ctypes.windll.user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 3)

from plan_simulation_stage2 import main

if __name__ == "__main__":
    main()
