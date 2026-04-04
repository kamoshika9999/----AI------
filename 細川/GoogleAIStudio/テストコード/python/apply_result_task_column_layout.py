# -*- coding: utf-8 -*-
"""
VBA ボタンから起動: マクロブック内「列設定_結果_タスク一覧」に従い
「結果_タスク一覧」の列順・列非表示を更新する。

環境変数 TASK_INPUT_WORKBOOK にマクロブックのフルパスが必要（VBA が設定）。
Excel で当該ブックを開いたまま実行すること（xlwings が接続する）。
"""
import logging
import os
import sys

os.chdir(os.path.dirname(os.path.abspath(__file__)))

import planning_core as pc  # noqa: E402


def main() -> int:
    logging.info("apply_result_task_column_layout: 開始")
    ok = pc.apply_result_task_column_layout_only()
    logging.info("apply_result_task_column_layout: 終了 ok=%s", ok)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
