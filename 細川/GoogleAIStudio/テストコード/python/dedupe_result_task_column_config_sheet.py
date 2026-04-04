# -*- coding: utf-8 -*-
"""
VBA から起動: 「列設定_結果_タスク一覧」の重複列名を除き A:B を書き直す。
「結果_タスク一覧」は変更しない。環境変数 TASK_INPUT_WORKBOOK 必須。Excel でブックを開いたまま。
チェックボックスを列 B に付けている場合は、整理後に「列設定_結果_タスク一覧_チェックボックスを配置」を再実行。
"""
import logging
import os
import sys

os.chdir(os.path.dirname(os.path.abspath(__file__)))

import planning_core as pc  # noqa: E402


def main() -> int:
    logging.info("dedupe_result_task_column_config_sheet: 開始")
    ok = pc.dedupe_result_task_column_config_sheet_only()
    logging.info("dedupe_result_task_column_config_sheet: 終了 ok=%s", ok)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
