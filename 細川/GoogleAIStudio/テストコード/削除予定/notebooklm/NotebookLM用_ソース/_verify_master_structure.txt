# -*- coding: utf-8 -*-
"""master.xlsm のシート・列を planning_core の想定と照合する（単発検証用）"""
import os
import sys

import pandas as pd


def _planning_repo_root() -> str:
    here = os.path.dirname(os.path.abspath(__file__))
    parent = os.path.dirname(here)
    if os.path.isfile(os.path.join(parent, "planning_core.py")):
        return parent
    return here


MASTER = os.path.join(_planning_repo_root(), "master.xlsm")

ATT_OPTIONAL = {"休暇区分"}
ATT_CORE = {
    "日付",
    "出勤時間",
    "退勤時間",
    "作業効率",
    "休憩時間1_開始",
    "休憩時間1_終了",
    "休憩時間2_開始",
    "休憩時間2_終了",
    "備考",
}


def main():
    if not os.path.isfile(MASTER):
        print("ERROR: master.xlsm が見つかりません:", MASTER)
        sys.exit(1)

    xls = pd.ExcelFile(MASTER)
    names = xls.sheet_names
    print("=== シート一覧 (%d) ===" % len(names))
    for i, n in enumerate(names, 1):
        print("  %2d. %s" % (i, n))

    issues = []
    warnings = []

    # ----- skills -----
    if "skills" not in names:
        issues.append('必須シート "skills" がありません')
        members = []
    else:
        raw = pd.read_excel(MASTER, sheet_name="skills", header=None)
        print("\n=== skills === shape=%s" % (raw.shape,))
        ne = 0
        if raw.shape[0] >= 3 and raw.shape[1] >= 2:
            for c in range(1, raw.shape[1]):
                p, m = raw.iat[0, c], raw.iat[1, c]
                if pd.isna(p) or pd.isna(m):
                    continue
                ps, ms = str(p).strip(), str(m).strip()
                if ps and ms and ps.lower() != "nan" and ms.lower() != "nan":
                    ne += 1
        print("  2段ヘッダ判定: 非空(工程+機械)列数 =", ne)
        if ne == 0:
            warnings.append(
                "skills: 2段ヘッダ未検出のため1行ヘッダ互換で読む可能性（列名に「工程+機械」形式が必要）"
            )

        members = []
        if raw.shape[0] >= 3 and ne > 0:
            for r in range(2, raw.shape[0]):
                mn = raw.iat[r, 0]
                if pd.isna(mn):
                    continue
                mname = str(mn).strip()
                if mname and mname.lower() not in ("nan", "none", "null"):
                    members.append(mname)

    # ----- need -----
    if "need" not in names:
        issues.append('必須シート "need" がありません')
    else:
        nr = pd.read_excel(MASTER, sheet_name="need", header=None)
        print("\n=== need === shape=%s" % (nr.shape,))
        ph = mh = br = None
        for r in range(nr.shape[0]):
            v0 = nr.iat[r, 0]
            if pd.isna(v0):
                continue
            s0 = str(v0).strip()
            if ph is None and s0 == "工程名":
                ph = r
            if mh is None and s0 == "機械名":
                mh = r
            if br is None and "必要人数" in s0 and not s0.startswith("特別指定"):
                br = r
        print("  工程名行:", ph, "機械名行:", mh, "基本必要人数行:", br)
        if ph is None or mh is None or br is None:
            issues.append("need: 「工程名」「機械名」「基本必要人数」を含む行が揃っていません")

    print("\n=== skills からのメンバー (%d) ===" % len(members))
    print(" ", members)

    skip_sub = {"skills", "need", "tasks"}
    print("\n=== メンバー別勤怠シート（列チェック）===")
    for sn in names:
        lsn = sn.lower()
        if "カレンダー" in sn or lsn in skip_sub:
            continue
        st = sn.strip()
        if st not in members:
            continue
        df = pd.read_excel(MASTER, sheet_name=sn)
        df.columns = df.columns.str.strip()
        cols = set(str(c).strip() for c in df.columns)
        need_cols = ATT_CORE
        missing = need_cols - cols
        if "備考" not in cols:
            issues.append("シート %r: 列「備考」がありません（AI・reason 用・必須）" % sn)
        if missing:
            m2 = missing - ATT_OPTIONAL
            if m2:
                warnings.append("シート %r: 想定列の欠け %s" % (sn, m2))
        print("  %s: %s" % (sn, sorted(cols)))

    for m in members:
        if m not in [x.strip() for x in names]:
            issues.append("メンバー %r が skills にいるが同名シートがありません" % m)

    for sn in names:
        if "カレンダー" in sn or sn.lower() in skip_sub:
            continue
        st = sn.strip()
        if st in members:
            continue
        if st and not st.startswith("_"):
            warnings.append(
                "シート %r は skills のメンバー名と一致しない（勤怠として読み飛ばされます）" % sn
            )

    print("\n=== 結果 ===")
    if issues:
        print("【不整合・要対応】")
        for x in issues:
            print(" -", x)
    else:
        print("致命的な不整合は検出されませんでした。")
    if warnings:
        print("【注意】")
        for w in warnings:
            print(" -", w)
    return 1 if issues else 0


if __name__ == "__main__":
    # コンソールが UTF-8 なら日本語が正しく表示されます（Windows: chcp 65001 等）
    rc = main()
    # planning_core の読込が通るか（例外のみ検知）
    try:
        import planning_core as pc

        os.chdir(_planning_repo_root())
        sd, mem, eq, req, rules, _surp = pc.load_skills_and_needs()
        if not mem:
            print("\n[ERROR] planning_core.load_skills_and_needs がメンバー0でした。")
            rc = 1
        else:
            att, _log = pc.load_attendance_and_analyze(mem)
            print(
                "\n=== planning_core 読込テスト OK === 勤怠日数=%d メンバー=%d"
                % (len(att), len(mem))
            )
    except Exception as e:
        print("\n[ERROR] planning_core 読込テスト:", e)
        rc = 1
    sys.exit(rc or 0)
