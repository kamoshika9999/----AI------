テスト用コード フォルダ
======================

このフォルダには、COM 検証・マスタ構造検証・段階1の PowerShell テストなど、
本番パイプライン（planning_core / task_extract_stage1 / plan_simulation_stage2）とは別に使うスクリプトをまとめています。

前提
----
- `master.xlsm`・`生産管理_AI配台テスト.xlsm`・`planning_core.py` は **親フォルダ（テストコード）** にあります。
- 各スクリプトは `_planning_repo_root()` で親を検出するため、このフォルダに置いたまま実行できます。

実行例（カレントを「テストコード」にしたうえで）
----------------------------------------------
  py テスト用コード\com_excel_com_sheet_verify.py
  py テスト用コード\com_exclude_rules_e9_a999.py
  py テスト用コード\_verify_master_structure.py
  powershell -ExecutionPolicy Bypass -File .\テスト用コード\run_stage1_exclude_rules_test.ps1

収録ファイル
------------
  com_excel_com_sheet_verify.py      … Excel COM でシート検証
  com_exclude_rules_e9_a999.py       … 設定シート E9 等への COM 書込テスト
  com_exclude_rules_column_e_test.py … E 列 COM 書込テスト
  _verify_master_structure.py        … master.xlsm と planning_core 想定の照合
  _verify_ae6_ooxml.py               … AE6 セル（OOXML）検証
  _verify_ae6_theme.py               … テーマ関連検証
  _verify_ae6_calendar.py            … カレンダー関連検証
  _inspect_g_output.py               … output 内の最新 production_plan をざっと確認 → log\_inspect_g_output.txt
  run_stage1_exclude_rules_test.ps1  … 段階1 + 配台不要 E 列テスト用ラッパー
