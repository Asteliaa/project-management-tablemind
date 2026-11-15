"""Сравнение сохранённых значений Excel (рабочая копия) и пересчёта LibreOffice (_recalc). Только чтение.
python3 -I compare_recalc.py <рабочая_папка> [имя_книги.xlsx ...]"""
import sys, os, glob, warnings, math
warnings.filterwarnings("ignore")
import openpyxl
root = sys.argv[1]; names = sys.argv[2:] or sorted(os.path.basename(p) for p in glob.glob(os.path.join(root, "*.xlsx")))
ERR = ("#NAME?", "#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NUM!", "#NULL!")
for n in names:
    a = os.path.join(root, n); b = os.path.join(root, "_recalc", n)
    if not os.path.exists(b): print(f"{n}: нет пересчёта"); continue
    A = openpyxl.load_workbook(a, data_only=True); B = openpyxl.load_workbook(b, data_only=True); F = openpyxl.load_workbook(a)
    same = diff = nocache = errs = 0; ex = []
    for ws in F.worksheets:
        if ws.title not in B.sheetnames: continue
        va, vb = A[ws.title], B[ws.title]
        for row in ws.iter_rows():
            for c in row:
                if not (isinstance(c.value, str) and c.value.startswith("=")): continue
                x, y = va[c.coordinate].value, vb[c.coordinate].value
                if isinstance(y, str) and (y in ERR or y.startswith("Err:")): errs += 1; ex.append(f"ERR {ws.title}!{c.coordinate}={y}")
                if x is None: nocache += 1; continue
                if isinstance(x, (int, float)) and isinstance(y, (int, float)):
                    if math.isclose(x, y, rel_tol=1e-9, abs_tol=1e-9): same += 1
                    else: diff += 1; ex.append(f"DIFF {ws.title}!{c.coordinate}: excel={x} lo={y}")
                elif x == y: same += 1
                else: diff += 1; ex.append(f"DIFF {ws.title}!{c.coordinate}: excel={str(x)[:25]} lo={str(y)[:25]}")
    print(f"{n}: формул совпало {same}, расходится {diff}, без сохранённого значения {nocache}, ошибок в пересчёте {errs}", ("| " + " ; ".join(ex[:4])) if ex else "")
