"""
Download every QC photo out of a Connecteam export BEFORE you cancel.
Connecteam's photo links live on their CDN and will likely stop working after cancellation.

Usage:
    pip install openpyxl requests
    python archive_connecteam.py "MDU Job Entry - Midwest.xlsx" connecteam_archive

Creates connecteam_archive/<entry#>_<project#>/photo_01.jpg ... plus index.csv
"""
import csv, os, re, sys
import openpyxl, requests

src, out = sys.argv[1], (sys.argv[2] if len(sys.argv) > 2 else "connecteam_archive")
os.makedirs(out, exist_ok=True)
ws = openpyxl.load_workbook(src).active
headers = [c.value for c in ws[1]]
photo_cols = [i for i, h in enumerate(headers, 1) if h and str(h).startswith("QC Photo")]

with open(os.path.join(out, "index.csv"), "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(headers[:9] + ["photo_folder", "photos_saved"] + headers[photo_cols[-1]:])
    for r in range(2, ws.max_row + 1):
        row = [ws.cell(r, c).value for c in range(1, len(headers) + 1)]
        if row[0] is None:
            continue
        folder = re.sub(r"[^\w\-]", "_", f"{row[0]}_{row[4]}")
        os.makedirs(os.path.join(out, folder), exist_ok=True)
        saved = 0
        for n, c in enumerate(photo_cols, 1):
            link = ws.cell(r, c).hyperlink
            if not link or not link.target:
                continue
            path = os.path.join(out, folder, f"photo_{n:02d}.jpg")
            if os.path.exists(path):
                saved += 1
                continue
            try:
                resp = requests.get(link.target, timeout=60)
                resp.raise_for_status()
                open(path, "wb").write(resp.content)
                saved += 1
            except Exception as e:
                print(f"  ! entry {row[0]} photo {n}: {e}")
        print(f"entry {row[0]}: {saved} photos")
        w.writerow(row[:9] + [folder, saved] + row[photo_cols[-1]:])
print("Done ->", out)
