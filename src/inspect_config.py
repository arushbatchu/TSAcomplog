import glob
import os
import sys

from openpyxl import load_workbook


def dump(path, max_rows=40, max_cols=14):
    print("=" * 78)
    print(os.path.basename(path))
    print("=" * 78)
    wb = load_workbook(path, data_only=True)
    for ws in wb.worksheets:
        print(f"\n--- sheet: {ws.title!r}  ({ws.max_row} rows x {ws.max_column} cols)")
        for r, row in enumerate(ws.iter_rows(max_row=max_rows, max_col=max_cols,
                                             values_only=True), start=1):
            cells = ["" if c is None else str(c) for c in row]
            while cells and cells[-1] == "":
                cells.pop()
            if not cells:
                continue
            print(f"  r{r:<3}", " | ".join(c[:22] for c in cells))
        if ws.max_row > max_rows:
            print(f"  ... {ws.max_row - max_rows} more rows")


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else "config"
    files = sorted(glob.glob(os.path.join(target, "*.xlsx")))
    if not files:
        print(f"no .xlsx found in {target!r}")
        sys.exit(1)
    for f in files:
        dump(f)
