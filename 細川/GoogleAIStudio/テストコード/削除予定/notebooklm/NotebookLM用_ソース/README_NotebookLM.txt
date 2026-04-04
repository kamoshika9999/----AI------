NotebookLM 用ソース一式（このフォルダ）
=====================================

使い方
------
Google NotebookLM の「ソース」に、このフォルダ内のファイルをまとめて追加してください。
（.py はアップロードできないため、同内容の .txt に拡張子だけ変更してあります。文字コードは UTF-8 です。）

含まれるファイル
--------------
■ 説明・仕様（Markdown）
  - NotebookLM用_工程管理AI配台システム参照資料.md … 長文の解説（NotebookLM 向けに厚め）
  - 配台計画と関連ロジック_仕様書.md … 箇条書き中心の仕様サマリ

■ Python ソースのコピー（.txt）
  - planning_core.txt … 配台シミュレーションの中核（最大）
  - plan_simulation_stage2.txt … 段階2の起動
  - task_extract_stage1.txt … 段階1の起動
  - main.txt … エントリの一例
  - _verify_master_structure.txt … マスタ構造の簡易検証

含めていないもの
----------------
- `テスト用コード` フォルダ内の COM 検証・マスタ検証スクリプト（本体は `..\テスト用コード\*.py`。NotebookLM 用に .txt が要れば手動コピー）
- 優先順位仕様書_生成.py など周辺ツール

必要ならテストコード直下から手動で .py をコピーし、拡張子を .txt にリネームして同フォルダに追加してください。

更新手順（開発者向け）
----------------------
テストコードフォルダでソースを直したあと、次で再同期できます（PowerShell 例）。

  $base = "（テストコードのパス）"
  $d = Join-Path $base "NotebookLM用_ソース"
  Copy-Item (Join-Path $base "planning_core.py") (Join-Path $d "planning_core.txt") -Force
  （他ファイルも同様）
