# -*- coding: utf-8 -*-
"""
起動中の Excel に xlwings で接続し、シートごとに読み取り・書き込みが成功するか検証する。

前提:
  - 対象の .xlsm / .xlsx を先に Excel で開いておく（または --dispatch-open）
  - pip install xlwings
  - Windows で Excel インストール済み

使い方（このフォルダで）:
  py com_excel_com_sheet_verify.py
  py com_excel_com_sheet_verify.py --book "C:\\path\\生産管理_AI配台テスト.xlsm"

結果:
  - 標準出力に表形式
  - log/com_excel_com_verify.txt にも UTF-8 で追記（--no-file で抑止）

接続:
  - 起動中インスタンスから full_name が一致するブックを検索
  - 見つからず --dispatch-open のとき: 表示付きの新規 Excel で Workbooks.Open
"""

from __future__ import annotations

import argparse
import os
import sys
from datetime import datetime

SKIP_SHEET_NAMES = frozenset({"COM操作テストログ"})
TEST_ZZ = "ZZ1000"
TEST_ZZ_VAL = "__XW_TEST__"
TEST_A99 = "A99"
TEST_A99_VAL = "A666"


def visible_label(v) -> str:
    try:
        iv = int(v)
    except Exception:
        return str(v)
    if iv == -1:
        return "表示"
    if iv == 0:
        return "非表示"
    if iv == 2:
        return "VeryHidden"
    return str(iv)


def safe_str(x, max_len: int = 80) -> str:
    s = repr(x) if x is not None else ""
    if len(s) > max_len:
        return s[: max_len - 3] + "..."
    return s


def try_read_a1(sheet) -> tuple[str, str]:
    try:
        v = sheet.range("A1").value
        return "OK", safe_str(v)
    except Exception as e:
        return "NG", str(e)


def try_used_range(sheet) -> tuple[str, str]:
    try:
        ur = sheet.used_range
        addr = str(ur.address).replace("$", "")
        return "OK", addr
    except Exception as e:
        return "NG", str(e)


def try_zz_write(sheet) -> tuple[str, str]:
    try:
        sheet.range(TEST_ZZ).value = TEST_ZZ_VAL
        sheet.range(TEST_ZZ).clear_contents()
        return "OK", ""
    except Exception as e:
        return "NG", str(e)


def try_a99_roundtrip(sheet) -> tuple[str, str]:
    r = sheet.range(TEST_A99)
    try:
        old = r.value
    except Exception as e:
        return "NG(退避)", str(e)
    try:
        r.value = TEST_A99_VAL
        back = r.value
        if str(back) != TEST_A99_VAL:
            try:
                r.value = old
            except Exception:
                pass
            return "NG(不一致)", safe_str(back)
        r.value = old
        return "OK", ""
    except Exception as e:
        try:
            r.value = old
        except Exception:
            pass
        return "NG", str(e)


def verify_sheet(sheet) -> dict:
    name = str(sheet.name)
    out: dict = {"name": name}
    try:
        out["protect"] = bool(sheet.api.ProtectContents)
    except Exception as e:
        out["protect"] = f"? {e}"

    try:
        out["visible"] = visible_label(sheet.visible)
    except Exception as e:
        out["visible"] = str(e)

    st, msg = try_read_a1(sheet)
    out["a1"] = st
    out["a1_note"] = msg

    st, msg = try_used_range(sheet)
    out["used"] = st
    out["used_note"] = msg

    st, msg = try_zz_write(sheet)
    out["zz"] = st
    out["zz_note"] = msg

    st, msg = try_a99_roundtrip(sheet)
    out["a99"] = st
    out["a99_note"] = msg

    return out


def _planning_repo_root() -> str:
    here = os.path.dirname(os.path.abspath(__file__))
    parent = os.path.dirname(here)
    if os.path.isfile(os.path.join(parent, "planning_core.py")):
        return parent
    return here


def default_book_path(repo_root: str) -> str:
    cand = os.path.join(repo_root, "生産管理_AI配台テスト.xlsm")
    if os.path.isfile(cand):
        return cand
    return cand


def main() -> int:
    repo_root = _planning_repo_root()
    script_dir = os.path.dirname(os.path.abspath(__file__))
    for d in (repo_root, script_dir):
        if d not in sys.path:
            sys.path.insert(0, d)

    parser = argparse.ArgumentParser(
        description="Excel xlwings シート書き込み検証（起動中のブック）"
    )
    parser.add_argument(
        "--book",
        default=default_book_path(repo_root),
        help="検証するブックのパス（既定: リポジトリ直下の 生産管理_AI配台テスト.xlsm）",
    )
    parser.add_argument("--no-file", action="store_true", help="log ファイルへ書かない")
    parser.add_argument(
        "--dispatch-open",
        action="store_true",
        help="接続に失敗したときだけ Excel を新規起動しブックを開く（別プロセスになり得る）",
    )
    args = parser.parse_args()
    book_path = os.path.abspath(args.book)

    import planning_core as pc

    attached = pc._xlwings_attach_workbook_for_tests(
        book_path,
        "COM_EXCEL_SHEET_VERIFY",
        allow_dispatch_open=bool(args.dispatch_open),
    )
    if attached is None:
        print("xlwings でブックに接続できませんでした。", file=sys.stderr)
        print("", file=sys.stderr)
        print("確認:", file=sys.stderr)
        print("  ・--book は「名前を付けて保存」で見えるフルパスと一致させる。", file=sys.stderr)
        print("  ・pip install xlwings と Excel の有効化。", file=sys.stderr)
        if not args.dispatch_open:
            print("  ・試行: py com_excel_com_sheet_verify.py --dispatch-open", file=sys.stderr)
        return 3

    book, _, attach_info = attached
    rows: list[dict] = []
    try:
        for sheet in book.sheets:
            if sheet.name in SKIP_SHEET_NAMES:
                continue
            rows.append(verify_sheet(sheet))
    except Exception as ex:
        print(f"シート列挙エラー: {ex}", file=sys.stderr)
        return 4

    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines: list[str] = []
    lines.append(f"=== Excel xlwings 検証 {ts} ===")
    lines.append(f"ブック: {book_path}")
    lines.append(f"接続: {attach_info}")
    lines.append("")
    hdr = "シート名\t保護\t表示\tA1\tUsedRange\tZZ書込\tA99<->A666"
    lines.append(hdr)
    for r in rows:
        prot = r.get("protect")
        if isinstance(prot, bool):
            prot_s = "あり" if prot else "なし"
        else:
            prot_s = str(prot)
        line = "\t".join(
            [
                str(r.get("name", "")),
                prot_s,
                str(r.get("visible", "")),
                f"{r.get('a1', '')}"
                + (f" ({r.get('a1_note', '')})" if r.get("a1_note") else ""),
                f"{r.get('used', '')}"
                + (f" ({r.get('used_note', '')})" if r.get("used_note") else ""),
                f"{r.get('zz', '')}"
                + (f" ({r.get('zz_note', '')})" if r.get("zz_note") else ""),
                f"{r.get('a99', '')}"
                + (f" ({r.get('a99_note', '')})" if r.get("a99_note") else ""),
            ]
        )
        lines.append(line)

    text = "\n".join(lines) + "\n"
    try:
        sys.stdout.write(text)
    except UnicodeEncodeError:
        sys.stdout.buffer.write(text.encode("utf-8", errors="replace"))

    if not args.no_file:
        log_dir = os.path.join(repo_root, "log")
        os.makedirs(log_dir, exist_ok=True)
        log_path = os.path.join(log_dir, "com_excel_com_verify.txt")
        with open(log_path, "a", encoding="utf-8", newline="\n") as f:
            f.write(text)
        print(f"(ログ追記: {log_path})", file=sys.stderr)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
