"""Preizkus pretvornika bilanc: izmišljena bilanca (format AJPES portal lp2025) mora dati
pričakovan Excel (tests/fixture_cbk_pricakovano.xlsx), celica za celico. Nato še podatki podjetja iz obrazca (D3 do D5)."""
import io, sys, urllib.request, uuid, pathlib
import openpyxl

T = pathlib.Path(__file__).parent
src = (T / "fixture_cbk_lp2025.xlsx").read_bytes()

def convert(extra=()):
    b = uuid.uuid4().hex
    parts = [f'--{b}\r\nContent-Disposition: form-data; name="files"; filename="lp2025_6736971000.xlsx"\r\nContent-Type: application/octet-stream\r\n\r\n'.encode() + src + b'\r\n']
    for k, v in (("tip_bilance", "8"), ("tip_subjekta", "1"), *extra):
        parts.append(f'--{b}\r\nContent-Disposition: form-data; name="{k}"\r\n\r\n{v}\r\n'.encode())
    body = b"".join(parts) + f"--{b}--\r\n".encode()
    req = urllib.request.Request("http://localhost:8781/cbk/convert", data=body, headers={"Content-Type": f"multipart/form-data; boundary={b}"})
    return urllib.request.urlopen(req, timeout=120).read()

got = openpyxl.load_workbook(io.BytesIO(convert())); exp = openpyxl.load_workbook(T / "fixture_cbk_pricakovano.xlsx")
diff = [(sn, i, j) for sn in exp.sheetnames for i, (ra, rb) in enumerate(zip(exp[sn].iter_rows(values_only=True), got[sn].iter_rows(values_only=True))) for j, (x, y) in enumerate(zip(ra, rb)) if x != y]
if diff:
    print("CBK: razlike", diff[:10]); sys.exit(1)
print("CBK: bilanca brez podatkov podjetja: enako pričakovanemu")

# podatki podjetja iz obrazca (CRMB / AJPES / HIBIS) pridejo v D3, D4, D5
ov = openpyxl.load_workbook(io.BytesIO(convert((("naziv", "ČŠŽ Testno podjetje d.o.o."), ("maticna", "1234567000"), ("partner", "F123456")))))
ws = ov["BS"]
vals = (ws["D3"].value, str(ws["D4"].value), ws["D5"].value)
if vals != ("ČŠŽ Testno podjetje d.o.o.", "1234567000", 123456):
    print("CBK: D3 do D5 narobe:", vals); sys.exit(1)
if (ov["IPI"]["D3"].value, str(ov["IPI"]["D4"].value), ov["IPI"]["D5"].value) != ("ČŠŽ Testno podjetje d.o.o.", "1234567000", 123456):
    print("CBK: IPI D3 do D5 niso vpisani neposredno"); sys.exit(1)
print("CBK: VSE V REDU")
