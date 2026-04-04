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
    pc.generate_plan()


if __name__ == "__main__":
    main()
