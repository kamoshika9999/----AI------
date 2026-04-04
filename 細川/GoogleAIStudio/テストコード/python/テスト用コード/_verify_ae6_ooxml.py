# -*- coding: utf-8 -*-
"""
master.xlsm を OOXML として解析し、AE6 の塗りと出勤簿.txt IsFactoryWorkingCell 判定を再現する。
"""
from __future__ import annotations

import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

def _planning_repo_root() -> Path:
    here = Path(__file__).resolve().parent
    parent = here.parent
    if (parent / "planning_core.py").is_file():
        return parent
    return here


REPO = _planning_repo_root()
MASTER = REPO / "master.xlsm"

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DRAW_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _ns(tag: str) -> str:
    return f"{{{MAIN_NS}}}{tag}"


def _parse_workbook_sheet_map(z: zipfile.ZipFile) -> dict[str, str]:
    """シート名 -> worksheet のパス（zip 内）"""
    wb_data = z.read("xl/workbook.xml")
    root = ET.fromstring(wb_data)
    rid_to_name: dict[str, str] = {}
    for sh in root.findall(_ns("sheets"))[0]:
        if sh.tag != _ns("sheet"):
            continue
        name = sh.get("name")
        rid = sh.get(f"{{{REL_NS}}}id")
        if name and rid:
            rid_to_name[rid] = name
    rel_data = z.read("xl/_rels/workbook.xml.rels")
    rel_root = ET.fromstring(rel_data)
    rid_to_target: dict[str, str] = {}
    for rel in rel_root:
        if "Relationship" not in rel.tag:
            continue
        rid = rel.get("Id")
        tgt = rel.get("Target")
        if rid and tgt:
            rid_to_target[rid] = "xl/" + tgt.replace("\\", "/").lstrip("/")
    out: dict[str, str] = {}
    for rid, name in rid_to_name.items():
        if rid in rid_to_target:
            out[name] = rid_to_target[rid]
    return out


def _cell_ref_col_row(ref: str) -> tuple[int, int]:
    m = re.match(r"^([A-Z]+)(\d+)$", ref.upper())
    if not m:
        raise ValueError(ref)
    letters, rs = m.group(1), int(m.group(2))
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch) - ord("A") + 1)
    return col, rs


def _col_row_to_ref(col: int, row: int) -> str:
    s = ""
    c = col
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return f"{s}{row}"


def _find_cell_in_sheet(sheet_xml: bytes, want_col: int, want_row: int) -> ET.Element | None:
    root = ET.fromstring(sheet_xml)
    ns = {"m": MAIN_NS}
    for row_el in root.findall(".//m:sheetData/m:row", ns):
        rnum = int(row_el.get("r", "0"))
        if rnum != want_row:
            continue
        for c in row_el.findall("m:c", ns):
            ref = c.get("r")
            if not ref:
                continue
            col, r = _cell_ref_col_row(ref)
            if col == want_col and r == want_row:
                return c
    return None


def _theme_colors_rgb(theme_xml: bytes) -> list[tuple[int, int, int] | None]:
    """clrScheme 子要素順（通常 dk1, lt1, dk2, lt2, accent1..6, ...）に sRGB を並べる"""
    root = ET.fromstring(theme_xml)
    scheme = root.find(f".//{{{DRAW_NS}}}clrScheme")
    if scheme is None:
        return []
    out: list[tuple[int, int, int] | None] = []
    for child in list(scheme):
        srgb = child.find(f".//{{{DRAW_NS}}}srgbClr")
        if srgb is not None and srgb.get("val"):
            v = srgb.get("val", "")
            if len(v) >= 6:
                r = int(v[0:2], 16)
                g = int(v[2:4], 16)
                b = int(v[4:6], 16)
                out.append((r, g, b))
                continue
        sysclr = child.find(f".//{{{DRAW_NS}}}sysClr")
        if sysclr is not None and sysclr.get("lastClr"):
            v = sysclr.get("lastClr", "")
            if len(v) >= 6:
                r = int(v[0:2], 16)
                g = int(v[2:4], 16)
                b = int(v[4:6], 16)
                out.append((r, g, b))
                continue
        out.append(None)
    return out


def _fill_rgb_from_ooxml(
    fill_el: ET.Element,
    theme_rgb: list[tuple[int, int, int] | None],
) -> tuple[int, int, int] | None:
    """patternFill fgColor theme/tint から RGB を推定（tint は簡易未対応で theme 基底のみ）"""
    pf = fill_el.find(f"{{{MAIN_NS}}}patternFill")
    if pf is None:
        return None
    fg = pf.find(f"{{{MAIN_NS}}}fgColor")
    if fg is None:
        return None
    rgb_attr = fg.get("rgb")
    if rgb_attr and len(rgb_attr) >= 6:
        # ARGB 8 hex
        s = rgb_attr[-6:]
        return (int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))
    th = fg.get("theme")
    if th is not None:
        idx = int(th)
        if 0 <= idx < len(theme_rgb) and theme_rgb[idx] is not None:
            return theme_rgb[idx]  # type: ignore[return-value]
    return None


def _cell_xfs_fill_index(styles_root: ET.Element, xf_index: int) -> int | None:
    cell_xfs = styles_root.find(f"{{{MAIN_NS}}}cellXfs")
    if cell_xfs is None:
        return None
    xfs = list(cell_xfs.findall(f"{{{MAIN_NS}}}xf"))
    if xf_index < 0 or xf_index >= len(xfs):
        return None
    fi = xfs[xf_index].get("fillId")
    return int(fi) if fi is not None else None


def _fill_element(fills_root: ET.Element, fill_id: int) -> ET.Element | None:
    fills = list(fills_root.findall(f"{{{MAIN_NS}}}fill"))
    if fill_id < 0 or fill_id >= len(fills):
        return None
    return fills[fill_id]


def is_factory_working(rr: int, gg: int, bb: int, pattern_none: bool) -> bool:
    if pattern_none:
        return True
    if (rr, gg, bb) == (255, 255, 255):
        return True
    if rr >= 235 and gg >= 220 and bb <= 190:
        return True
    return False


def main() -> int:
    if not MASTER.exists():
        print(f"not found: {MASTER}", file=sys.stderr)
        return 1

    z = zipfile.ZipFile(MASTER)
    smap = _parse_workbook_sheet_map(z)
    sheet_path = None
    sheet_name = None
    for name, path in smap.items():
        if "3" in name and "カレンダー" in name and not name.startswith("結果_"):
            sheet_path = path
            sheet_name = name
            break
    if not sheet_path:
        z.close()
        print("3月カレンダー シートが見つかりません", file=sys.stderr)
        return 1

    sheet_xml = z.read(sheet_path)
    AE_COL, ROW = 31, 6
    cell = _find_cell_in_sheet(sheet_xml, AE_COL, ROW)

    styles_xml = z.read("xl/styles.xml")
    styles_root = ET.fromstring(styles_xml)
    fills_root = styles_root.find(f"{{{MAIN_NS}}}fills")
    if fills_root is None:
        z.close()
        print("fills なし", file=sys.stderr)
        return 1

    theme_xml = z.read("xl/theme/theme1.xml")
    theme_rgb = _theme_colors_rgb(theme_xml)

    lines: list[str] = []
    lines.append(f"ファイル: {MASTER.name}")
    lines.append(f"シート: {sheet_name} -> {sheet_path}")
    lines.append(f"対象セル: AE6 (列={AE_COL}, 行={ROW})")
    lines.append("")

    if cell is None:
        lines.append("シート XML 上に AE6 の <c> 要素が見つかりませんでした（空セルで未保存の可能性）。")
        lines.append("→ Excel 上では既定スタイルの可能性あり。手元のブックで確認してください。")
    else:
        t = cell.get("t")
        s = cell.get("s")
        v_el = cell.find(f"{{{MAIN_NS}}}v")
        is_el = cell.find(f"{{{MAIN_NS}}}is")
        val = ""
        if v_el is not None and v_el.text is not None:
            val = v_el.text
        elif is_el is not None:
            t_el = is_el.find(f"{{{MAIN_NS}}}t")
            if t_el is not None and t_el.text:
                val = t_el.text
        lines.append(f"<c> t={t!r} s={s!r} 値(表示用要素)={val!r}")

        if s is None:
            lines.append("style s 属性なし -> fillId 0（styles の先頭 fill）を使う実装もあるが、Excel 既定は要確認")
            fill_id = 0
        else:
            xf_index = int(s)
            fill_id = _cell_xfs_fill_index(styles_root, xf_index)
            lines.append(f"cellXfs インデックス: {xf_index} -> fillId={fill_id}")

        if fill_id is not None:
            fel = _fill_element(fills_root, fill_id)
            if fel is not None:
                pattern_fill = fel.find(f"{{{MAIN_NS}}}patternFill")
                pat_type = None
                if pattern_fill is not None:
                    pat_type = pattern_fill.get("patternType")
                pattern_none = pat_type in (None, "none")
                lines.append(f"fill[{fill_id}] patternType={pat_type!r} -> xlPatternNone 相当: {pattern_none}")
                rgb = _fill_rgb_from_ooxml(fel, theme_rgb)
                if rgb:
                    rr, gg, bb = rgb
                    lines.append(f"推定 RGB: R={rr} G={gg} B={bb}")
                    lines.append(
                        f"IsFactoryWorkingCell: 白={rgb==(255,255,255)} 黄系={rr >= 235 and gg >= 220 and bb <= 190}"
                    )
                    w = is_factory_working(rr, gg, bb, pattern_none)
                    lines.append(f"→ 総合 isWorkingDay 相当: {w}")
                else:
                    lines.append("RGB を OOXML から確定できませんでした（tint のみ等）。")

    z.close()

    text = "\n".join(lines)
    outp = REPO / "output" / "verify_ae6_ooxml.txt"
    outp.parent.mkdir(parents=True, exist_ok=True)
    outp.write_text(text, encoding="utf-8")
    print(text)
    print(f"\n[UTF-8 保存] {outp}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
