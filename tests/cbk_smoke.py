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

# DENAR | OBRAČUNI (VIZIJA): postavke z drugačnim besedilom kot uradni AOP nazivi se morajo ujeti
sys.path.insert(0, str(T.parent))
from cbk import pretvornik as P
cases = [
    (True, "F. KOSMATI DONOS IZ POSLOVANJA", None, ("126", False)),
    (True, "B. SPREMEMBA VREDNOSTI ZALOG PROIZVODOV IN NEDOKONČANE", None, ("121", True)),
    (True, "D. SUBVENCIJE, DOTACIJE, REGRESI, KOMPENZACIJE IN DRUGI PRIHODKI,", None, ("124", False)),
    (True, "c) povračila stroškov zaposlencem v zvezi z delom", None, ("137", False)),
    (True, "3. Prevrednotovalni poslovni odhodki pri obrat. sr.", None, ("147", False)),
    (True, "K. FINAČNI ODHODKI", None, ("166", False)),
    (True, "II. Finačni odhodki iz finančnih obveznosti", None, ("169", False)),
    (True, "III. Finačni odhodki iz poslovnih obveznosti", None, ("174", False)),
    (True, "II. Drugi finančni prihodki in ostali prihodki", None, ("180", False)),
    (True, "P. DAVEK IZ DOBIČKA", None, ("184", False)),
    (True, "1. Čisti prihodki od prodaje proizvodov in storitev", "eu", ("116", False)),
    (True, "1. Čisti prihodki od prodaje proizvodov in storitev", "izven_eu", ("119", False)),
    (False, "4. Druge naprave in oprema, drobni inventar in druga opredmetena osnovna", None, ("014", False)),
    (False, "IV. Presežek iz prevrednotenja", None, ("301", False)),
    (False, "Zabilančna sredstva", None, ("054", False)),
    (False, "Zabilančne obveznosti", None, ("096", False)),
]
bad = [(c[1], P._vizija_alias(c[1], c[2], c[0])) for c in cases if P._vizija_alias(c[1], c[2], c[0]) != c[3]]
if bad:
    print("CBK: VIZIJA aliasi narobe:", bad); sys.exit(1)
print("CBK: VIZIJA aliasi v redu")
