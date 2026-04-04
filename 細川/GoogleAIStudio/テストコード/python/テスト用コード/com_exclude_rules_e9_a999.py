# -*- coding: utf-8 -*-
"""
起動中の Excel ブック（xlwings）で「設定_配台不要工程」シートの指定セルに文字列を書き込む最小テスト。
元のセル値へは戻しません（上書きしたまま終了します）。

【重要】画面の Excel に反映されないとき
  planning_core の _xlwings_attach_open_macro_workbook は、起動中にブックが見つからないと
  非表示の別プロセスで開き、終了時に保存せず閉じるため、変更が破棄され画面は変わりません。
  既定は _xlwings_attach_workbook_for_tests（起動中ブック優先、--dispatch-open で新規表示）です。

既定: 値は A999、Save はしない（メモリ上のブックのみ更新）。
  py com_exclude_rules_e9_a999.py

開いていないときだけ新しい Excel で開く（表示・別プロセスの可能性）:
  py com_exclude_rules_e9_a999.py --dispatch-open

planning_core と同じ「本番寄り」接続（非表示別インスタンスになり得る）:
  py com_exclude_rules_e9_a999.py --planning-attach

ディスクに保存も試す:
  py com_exclude_rules_e9_a999.py --save
"""

from __future__ import annotations

import argparse
import logging
import os
import sys


def _script_dir() -> str:
    return os.path.dirname(os.path.abspath(__file__))


def _planning_repo_root() -> str:
    here = _script_dir()
    parent = os.path.dirname(here)
    if os.path.isfile(os.path.join(parent, "planning_core.py")):
        return parent
    return here


def _refresh_excel_ui(book, sheet) -> None:
    try:
        book.app.screen_updating = True
    except Exception:
        pass
    try:
        book.app.visible = True
    except Exception:
        pass
    try:
        book.activate()
    except Exception:
        pass
    try:
        sheet.activate()
    except Exception:
        pass


def main() -> int:
    sd = _script_dir()
    repo = _planning_repo_root()
    for d in (repo, sd):
        if d not in sys.path:
            sys.path.insert(0, d)

    logging.disable(logging.WARNING)
    import planning_core as pc

    logging.disable(logging.NOTSET)
    logging.getLogger().setLevel(logging.CRITICAL)

    default_book = os.path.join(repo, "生産管理_AI配台テスト.xlsm")
    p = argparse.ArgumentParser(
        description="設定_配台不要工程 へ xlwings 書込（元値の復元はしない）"
    )
    p.add_argument("--book", default=default_book, help="マクロブックのフルパス")
    p.add_argument("--row", type=int, default=9, help="書き込み行（既定 9）")
    p.add_argument("--col", type=int, default=5, help="列番号（E=5。既定 5）")
    p.add_argument("--text", default="A999", help="セルに書く文字列（既定 A999）")
    p.add_argument(
        "--save",
        action="store_true",
        help="book.save() も試す（既定はメモリのみ。R/O や二重起動で失敗し得る）",
    )
    p.add_argument(
        "--dispatch-open",
        action="store_true",
        help="起動中 Excel にブックが無いときだけ、新規 Excel(表示)で開く",
    )
    p.add_argument(
        "--planning-attach",
        action="store_true",
        help="planning_core._xlwings_attach_open_macro_workbook を使う（非表示の別プロセスになり得る）",
    )
    args = p.parse_args()

    if args.dispatch_open and args.planning_attach:
        print("--dispatch-open と --planning-attach は同時に指定しないでください。", file=sys.stderr)
        return 2

    book_path = os.path.abspath(args.book)
    if not os.path.isfile(book_path):
        print(f"ブックがありません: {book_path}", file=sys.stderr)
        return 2

    xw_book = None
    info: dict
    attach_label = ""

    if args.planning_attach:
        attached = pc._xlwings_attach_open_macro_workbook(book_path, "XLWINGS_E9_A999")
        if attached is None:
            print(
                "xlwings でブックに接続できません（--planning-attach）。",
                file=sys.stderr,
            )
            return 3
        xw_book, info = attached
        attach_label = "planning_core._xlwings_attach_open_macro_workbook"
        if info.get("mode") == "quit_excel":
            print(
                "[警告] 非表示の別 Excel でブックを開いています。"
                " 終了時に Save しないと変更は破棄され、いま画面で見ている Excel には出ません。",
                file=sys.stderr,
            )
    else:
        attached = pc._xlwings_attach_workbook_for_tests(
            book_path,
            "XLWINGS_E9_A999",
            allow_dispatch_open=bool(args.dispatch_open),
        )
        if attached is None:
            print("xlwings で接続できませんでした。", file=sys.stderr)
            print(
                "ヒント: 対象を Excel で開き、--book を「名前を付けて保存」のフルパスと一致させてください。",
                file=sys.stderr,
            )
            print(
                "  試す: py com_exclude_rules_e9_a999.py --dispatch-open",
                file=sys.stderr,
            )
            return 3
        xw_book, info, attach_label = attached

    ok = False
    try:
        try:
            xw_book.app.display_alerts = False
        except Exception:
            pass

        try:
            print(f"接続: {attach_label}")
            print(f"ブック full_name: {xw_book.full_name}")
            print(f"Excel.visible: {xw_book.app.visible}")
        except Exception:
            pass

        sheet = xw_book.sheets[pc.EXCLUDE_RULES_SHEET_NAME]
        r, c = int(args.row), int(args.col)
        before = sheet.range((r, c)).value
        sheet.range((r, c)).value = args.text
        after = sheet.range((r, c)).value

        _refresh_excel_ui(xw_book, sheet)

        addr = f"{pc.EXCLUDE_RULES_SHEET_NAME} R{r}C{c}（列{c}がE列）"
        print(f"変更前: {before!r}")
        print(f"書込:   {args.text!r}")
        print(f"読取:   {after!r}")
        print(f"一致:   {str(after) == str(args.text)}  （{addr}）")

        if not args.save:
            print("メモリ上のみ更新しました（--save なしのため Save しません）。")
            ok = True
            return 0

        try:
            xw_book.save()
            print("[save] Save しました。")
            ok = True
            return 0
        except Exception as ex:
            print(f"[save] 失敗: {ex}", file=sys.stderr)
            ok = True
            return 5

    except Exception as ex:
        print(f"エラー: {ex}", file=sys.stderr)
        import traceback

        traceback.print_exc()
        return 4
    finally:
        if xw_book is not None:
            pc._xlwings_release_book_after_mutation(xw_book, info, ok)


if __name__ == "__main__":
    raise SystemExit(main())
