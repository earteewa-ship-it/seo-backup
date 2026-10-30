"""Разбивает .xlsx на CSV по листам: python3 xlsx_to_csv.py файл.xlsx папка_вывода"""
import csv, os, re, sys, openpyxl
src, out = sys.argv[1], sys.argv[2]
os.makedirs(out, exist_ok=True)
wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
for ws in wb.worksheets:
    rows = [["" if v is None else v for v in r] for r in ws.iter_rows(values_only=True)]
    while rows and not any(str(c).strip() for c in rows[-1]): rows.pop()
    w = max((max((i+1 for i,c in enumerate(r) if str(c).strip()), default=0) for r in rows), default=0)
    name = re.sub(r'[\\/:*?"<>|]', '_', ws.title).strip() or "sheet"
    with open(os.path.join(out, name + ".csv"), "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(r[:w] for r in rows)
    print(name, len(rows), w)
