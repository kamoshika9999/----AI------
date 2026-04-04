import glob
import os
from pathlib import Path

import pandas as pd


def _planning_repo_root() -> Path:
    here = Path(__file__).resolve().parent
    parent = here.parent
    if (parent / "planning_core.py").is_file():
        return parent
    return here


REPO = _planning_repo_root()
base = str(REPO / "output")
files = sorted(
    glob.glob(base + r"\**\production_plan_multi_day_*.xlsx", recursive=True),
    key=os.path.getmtime,
)

out_lines = []
out_lines.append(f"count={len(files)}")
if files:
    latest = files[-1]
    out_lines.append(f"latest={latest}")
    df = pd.read_excel(latest, sheet_name="結果_タスク一覧")
    out_lines.append("cols=" + ",".join([str(c) for c in list(df.columns)[:12]]))
else:
    out_lines.append("latest=NONE")

out_path = REPO / "log" / "_inspect_g_output.txt"
out_path.parent.mkdir(parents=True, exist_ok=True)
with open(out_path, "w", encoding="utf-8") as f:
    f.write("\n".join(out_lines))
