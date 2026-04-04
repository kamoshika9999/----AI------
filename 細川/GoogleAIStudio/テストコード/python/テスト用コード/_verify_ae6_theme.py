# -*- coding: utf-8 -*-
"""xlsm を zip として theme + styles を読み AE6 の塗り RGB を推定"""
import re
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
}

def _planning_repo_root() -> Path:
    here = Path(__file__).resolve().parent
    parent = here.parent
    if (parent / "planning_core.py").is_file():
        return parent
    return here


REPO = _planning_repo_root()
MASTER = REPO / "master.xlsm"


def theme_index_to_srgb(theme_xml: str, idx: int) -> tuple[int, int, int] | None:
    """Excel theme color index 0..9 -> sRGB (theme1 の clrScheme 順)"""
    root = ET.fromstring(theme_xml)
    scheme = root.find(".//a:clrScheme", NS)
    if scheme is None:
        return None
    # Order: dk1, lt1, dk2, lt2, accent1..6, hyperlink, followedHyperlink
    children = list(scheme)
    names = [re.sub(r"\{[^}]+\}", "", c.tag) for c in children]
    # Find srgbClr or sysClr
    if idx < 0 or idx >= len(children):
        return None
    el = children[idx]
    srgb = el.find(".//a:srgbClr", NS)
    if srgb is not None:
        val = srgb.get("val")
        if val and len(val) == 6:
            r = int(val[0:2], 16)
            g = int(val[2:4], 16)
            b = int(val[4:6], 16)
            return (r, g, b)
    return None


def main() -> None:
    z = zipfile.ZipFile(MASTER)
    theme_xml = z.read("xl/theme/theme1.xml").decode("utf-8")
    # OOXML theme index 0 = lt1 相当が多い（Office の実装に依存）
    # openpyxl reported theme=0 for AE6
    for idx in range(10):
        rgb = theme_index_to_srgb(theme_xml, idx)
        print(f"theme index {idx}: {rgb}")
    z.close()


if __name__ == "__main__":
    main()
