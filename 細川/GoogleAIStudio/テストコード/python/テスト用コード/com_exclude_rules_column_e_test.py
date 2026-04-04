# -*- coding: utf-8 -*-
"""
「設定_配台不要工程」シートの E 列（5 列目）が xlwings で書けるか検証する。

planning_core と同じ xlwings 接続（_xlwings_attach_open_macro_workbook）を使う。

前提:
  - xlwings / Excel
  - 対象 .xlsm を Excel で開いておくか、接続できないときは非表示 Excel が起動する（本番と同様）
  - 同じファイルを 2 つの Excel で開くと Save で失敗することがある → --no-save で書込検証のみ、または Excel は 1 つにする

例（cmd）:
  py com_exclude_rules_column_e_test.py
  py com_exclude_rules_column_e_test.py --bulk-block
  py com_exclude_rules_column_e_test.py --restore

PowerShell:
  - コマンドは必ず 1 行で（途中改行すると --row だけ実行されたり、--value の引数が壊れます）。
  - JSON を --value で渡すとき、'{"test":1}' でも内部の " が解釈で落ちて {test:1} になることがあります。
    → 推奨: --value-json test=1
    → または: $v = '{"test":1}'; py com_exclude_rules_column_e_test.py --row 5 --value $v --restore --no-save
"""

from __future__ import annotations

import argparse
import json
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


def _try_book_save(book, label: str, *, no_save: bool) -> bool:
    p = f"[{label}]"
    if no_save:
        print(f"{p} --no-save のため Save しません（メモリ上のブックには書込済み）。")
        return True
    try:
        book.save()
        print(f"{p} Save しました。")
        return True
    except Exception as ex:
        text = str(ex)
        print(f"{p} Save に失敗しました: {text}", file=sys.stderr)
        if "読み取り専用" in text or "read-only" in text.lower():
            print(
                f"{p} 別の Excel が同じファイルを開いていると R/O で開くことがあります。",
                file=sys.stderr,
            )
        print(
            f"{p} セルへの書込・読取の確認まで済んでいるなら --no-save で再実行できます。",
            file=sys.stderr,
        )
        return False


def _warn_if_value_looks_like_broken_json(s: str) -> None:
    t = str(s).strip()
    if not (t.startswith("{") and t.endswith("}")):
        return
    try:
        json.loads(t)
        return
    except json.JSONDecodeError:
        pass
    print(
        "[警告] --value が { ... } 形式ですが有効な JSON ではありません。"
        " PowerShell では `\"` が落ちて `{test:1}` のようになることがあります。",
        file=sys.stderr,
    )
    print(
        "  対策: 1 行で実行する／--value-json test=1 を使う／次のように変数で渡す: "
        '$v = \'{"test":1}\'; py com_exclude_rules_column_e_test.py --row 5 --value $v ...',
        file=sys.stderr,
    )


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
        description="設定_配台不要工程 の E 列 xlwings 書き込みテスト",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "PowerShell: JSON は --value-json を推奨。--value 使うときは 1 行で渡すか "
            "$v='{\"test\":1}'; py ... --value $v ... のように変数経由にしてください。"
        ),
    )
    p.add_argument("--book", default=default_book, help="マクロブックのフルパス")
    p.add_argument(
        "--row",
        type=int,
        default=9,
        help="単体テストで書く行（1 行目は見出しのため 2 以上推奨）",
    )
    p.add_argument(
        "--value",
        default=None,
        help="E 列に書き込む文字列（--value-json より優先）",
    )
    p.add_argument(
        "--value-json",
        action="append",
        metavar="KEY=VAL",
        help=(
            "JSON オブジェクトを組み立てて E 列に書く（複数回可）。例: --value-json test=1 --value-json name=foo "
            "→ {\"test\":1,\"name\":\"foo\"}（値は数値化できるなら数値）"
        ),
    )
    p.add_argument(
        "--no-save",
        action="store_true",
        help="Save しない（Excel 上の編集は試すがディスクに書かない）",
    )
    p.add_argument(
        "--restore",
        action="store_true",
        help="単体テスト後に元のセル値へ戻して Save（--bulk-block とは併用しない）",
    )
    p.add_argument(
        "--bulk-block",
        action="store_true",
        help="本番同様 A1:E3 に 2 次元配列を代入（先頭 3 行を上書き。注意）",
    )
    args = p.parse_args()

    if args.value_json:
        obj: dict = {}
        for pair in args.value_json:
            if "=" not in pair:
                p.error(f"--value-json は KEY=VAL 形式: {pair!r}")
            k, v = pair.split("=", 1)
            k = k.strip()
            v = v.strip()
            if not k:
                p.error(f"--value-json のキーが空: {pair!r}")
            try:
                obj[k] = json.loads(v)
            except json.JSONDecodeError:
                try:
                    obj[k] = int(v)
                except ValueError:
                    try:
                        obj[k] = float(v)
                    except ValueError:
                        obj[k] = v
        cell_value = json.dumps(obj, ensure_ascii=False)
    elif args.value is not None:
        cell_value = args.value
        _warn_if_value_looks_like_broken_json(cell_value)
    else:
        cell_value = '{"__xlwings_e_column_test__":true}'

    book_path = os.path.abspath(args.book)
    if not os.path.isfile(book_path):
        print(f"ブックがありません: {book_path}", file=sys.stderr)
        return 2

    if args.bulk_block and args.restore:
        print("--bulk-block と --restore は同時に使わないでください。", file=sys.stderr)
        return 2

    attached = pc._xlwings_attach_open_macro_workbook(book_path, "XLWINGS_E_COL_TEST")
    if attached is None:
        print(
            "xlwings でブックに接続できません。Excel で対象を開くか、planning_core と同じ環境か確認してください。",
            file=sys.stderr,
        )
        return 3

    xw_book, info = attached
    ok = False
    try:
        try:
            xw_book.app.display_alerts = False
        except Exception:
            pass

        sheet = xw_book.sheets[pc.EXCLUDE_RULES_SHEET_NAME]

        if args.bulk_block:
            hdr = [
                pc.EXCLUDE_RULE_COL_PROCESS,
                pc.EXCLUDE_RULE_COL_MACHINE,
                pc.EXCLUDE_RULE_COL_FLAG,
                pc.EXCLUDE_RULE_COL_LOGIC_JA,
                pc.EXCLUDE_RULE_COL_LOGIC_JSON,
            ]
            r1 = [
                "__XW_E_TEST_R1__",
                "__M_XW__",
                False,
                "xlwings一括テスト1行目",
                '{"__bulk_e_row2__":1}',
            ]
            r2 = ["__XW_E_TEST_R2__", "__M2__", False, "", ""]
            matrix = [hdr, r1, r2]
            sheet.range((1, 1), (3, 5)).value = matrix
            e2 = sheet.range((2, 5)).value
            e3 = sheet.range((3, 5)).value
            print("[一括] A1:E3 代入後:")
            print("  E2 =", repr(e2))
            print("  E3 =", repr(e3))
            if str(e2).find("__bulk_e_row2__") >= 0 or e2 == '{"__bulk_e_row2__":1}':
                print("  → E2 に期待文字列が見えます（OK の目安）")
            else:
                print("  → E2 が期待と異なる可能性があります（要確認）")
            print("  （E3 はテストデータで E 列を空にしているため None になります）")
            saved = _try_book_save(xw_book, "一括", no_save=args.no_save)
            if saved or args.no_save:
                print("[一括] A1:E3 はテスト用データに置き換わっています（Save できた場合はディスクにも反映）。")
            ok = True
            return 0 if (saved or args.no_save) else 5

        row = max(2, int(args.row))
        col_e = 5
        old = sheet.range((row, col_e)).value
        print(f"[単体] 変更前 E{row} = {old!r}")

        sheet.range((row, col_e)).value = cell_value
        back = sheet.range((row, col_e)).value
        print(f"[単体] 書込値     = {cell_value!r}")
        print(f"[単体] 直後の読取 = {back!r}")

        same = False
        try:
            same = str(back) == str(cell_value)
        except Exception:
            pass
        print(f"[単体] 文字列一致: {same}")

        saved_write = _try_book_save(xw_book, "単体", no_save=args.no_save)

        if args.restore:
            sheet.range((row, col_e)).value = old
            saved_restore = True
            if not args.no_save:
                saved_restore = _try_book_save(xw_book, "単体・復元", no_save=False)
            print(f"[単体] 元の値に戻しました（メモリ上）→ {old!r}")
            if args.no_save:
                pass
            elif not saved_restore:
                print(
                    "[単体] 復元後の Save に失敗しました（ディスクは未更新の可能性）。",
                    file=sys.stderr,
                )

        ok = True
        if args.no_save:
            return 0
        if not saved_write:
            return 5
        if args.restore and not args.no_save and not saved_restore:
            return 5
        return 0

    except Exception as ex:
        print(f"エラー: {ex}", file=sys.stderr)
        import traceback

        traceback.print_exc()
        return 4
    finally:
        pc._xlwings_release_book_after_mutation(xw_book, info, ok)


if __name__ == "__main__":
    raise SystemExit(main())
