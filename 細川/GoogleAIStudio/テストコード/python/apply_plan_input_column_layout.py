# -*- coding: utf-8 -*-
"""
VBA から呼び出し: 配台計画_タスク入力の列順・表示のみ planning_core に任せる。
環境変数 TASK_INPUT_WORKBOOK にマクロブックのフルパスが必要（VBA が設定）。
"""
import os
import sys

os.chdir(os.path.dirname(os.path.abspath(__file__)))

import planning_core as pc


def main() -> int:
    ok = pc.apply_plan_input_column_layout_only()
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
