from flask import Flask, request, send_file, render_template_string, jsonify
import os, re, shutil, tempfile, uuid
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional

try:
    import pdfplumber
    import openpyxl
except ImportError:
    pass

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

TEMPLATE_PATH   = Path(__file__).parent / "template.xlsx"
UPLOAD_DIR      = Path(tempfile.gettempdir()) / "ajpes_uploads"
UPLOAD_DIR.mkdir(exist_ok=True)
ACCESS_PASSWORD = os.environ.get("ACCESS_PASSWORD", "bilance2025")

# ── JOLP → AOP mapping ────────────────────────────────────────────────────────
JOLP_TO_AOP = {
    # Bilanca stanja
    "SREDSTVA":                                                           "001",
    "A. DOLGOROČNA SREDSTVA":                                             "002",
    "I. Neopredmetena sredstva in dolgoročne aktivne časovne razmejitve": "003",
    "1. Neopredmetena sredstva":                                          "004",
    "2. Dolgoročne aktivne časovne razmejitve":                           "009",
    "II. Opredmetena osnovna sredstva":                                   "010",
    "III. Naložbene nepremičnine":                                        "018",
    "IV. Dolgoročne finančne naložbe":                                    "019",
    "1. Dolgoročne finančne naložbe, razen posojil":                      "020",
    "2. Dolgoročna posojila":                                             "024",
    "V. Dolgoročne poslovne terjatve":                                    "027",
    "VI. Odložene terjatve za davek":                                     "031",
    "B. KRATKOROČNA SREDSTVA":                                            "032",
    "I. Sredstva (skupine za odtujitev) za prodajo":                      "033",
    "II. Zaloge":                                                         "034",
    "III. Kratkoročne finančne naložbe":                                  "040",
    "1. Kratkoročne finančne naložbe, razen posojil":                     "041",
    "2. Kratkoročna posojila":                                            "045",
    "IV. Kratkoročne poslovne terjatve":                                  "048",
    "V. Denarna sredstva":                                                "052",
    "C. KRATKOROČNE AKTIVNE ČASOVNE RAZMEJITVE":                         "053",
    "Zunajbilančna sredstva":                                             "054",
    "OBVEZNOSTI DO VIROV SREDSTEV":                                       "055",
    "A. KAPITAL":                                                         "056",
    "I. Vpoklicani kapital":                                              "057",
    "1. Osnovni kapital":                                                 "058",
    "2. Nevpoklicani kapital (kot odbitna postavka)":                     "059",
    "II. Kapitalske rezerve":                                             "060",
    "III. Rezerve iz dobička":                                            "061",
    "IV. Revalorizacijske rezerve":                                       "067",
    "V. Rezerve, nastale zaradi vrednotenja po pošteni vrednosti":        "301",
    "VI. Preneseni čisti poslovni izid":                                  "068",
    "VII. Čisti poslovni izid poslovnega leta":                           "070",
    "B. REZERVACIJE IN DOLGOROČNE PASIVNE ČASOVNE RAZMEJITVE":           "072",
    "1. Rezervacije":                                                     "073",
    "2. Dolgoročne pasivne časovne razmejitve":                           "074",
    "C. DOLGOROČNE OBVEZNOSTI":                                           "075",
    "I. Dolgoročne finančne obveznosti":                                  "076",
    "II. Dolgoročne poslovne obveznosti":                                 "080",
    "III. Odložene obveznosti za davek":                                  "084",
    "Č. KRATKOROČNE OBVEZNOSTI":                                          "085",
    "I. Obveznosti, vključene v skupine za odtujitev":                    "086",
    "II. Kratkoročne finančne obveznosti":                                "087",
    "III. Kratkoročne poslovne obveznosti":                               "091",
    "D. KRATKOROČNE PASIVNE ČASOVNE RAZMEJITVE":                         "095",
    "Zunajbilančne obveznosti":                                           "096",
    # Izkaz poslovnega izida
    "1. ČISTI PRIHODKI OD PRODAJE":                                       "110",
    "2. SPREMEMBA VREDNOSTI ZALOG PROIZVODOV IN NEDOKONČANE":             "121",
    "3. USREDSTVENI LASTNI PROIZVODI IN LASTNE STORITVE":                 "123",
    "4. DRUGI POSLOVNI PRIHODKI":                                         "125",
    "5. Stroški blaga, materiala in storitev":                            "128",
    "a) Nabavna vrednost prodanega blaga in materiala ter stroški porabljenega": "129",
    "b) Stroški storitev":                                                "134",
    "6. Stroški dela":                                                    "139",
    "a) Stroški plač":                                                    "140",
    "b) Stroški pokojninskih zavarovanj":                                 "141",
    "c) Stroški drugih socialnih zavarovanj":                             "142",
    "č) Drugi stroški dela":                                              "143",
    "7. Odpisi vrednosti":                                                "144",
    "a) Amortizacija":                                                    "145",
    "b) Prevrednotovalni poslovni odhodki pri neopredmetenih sredstvih in":"146",
    "c) Prevrednotovalni poslovni odhodki pri obratnih sredstvih":        "147",
    "8. Drugi poslovni odhodki":                                          "148",
    "9. Finančni prihodki iz deležev":                                    "155",
    "10. Finančni prihodki iz danih posojil":                             "160",
    "11. Finančni prihodki iz poslovnih terjatev":                        "163",
    "12. Finančni odhodki iz oslabitve in odpisov finančnih naložb":      "168",
    "13. Finančni odhodki iz finančnih obveznosti":                       "169",
    "14. Finančni odhodki iz poslovnih obveznosti":                       "174",
    "15. DRUGI PRIHODKI":                                                 "178",
    "16. DRUGI ODHODKI":                                                  "181",
    "17. DAVEK IZ DOBIČKA":                                               "184",
    "18. ODLOŽENI DAVKI":                                                 "185",
    "19. ČISTI POSLOVNI IZID OBRAČUNSKEGA OBDOBJA":                      "186",
    "21. PRENESENI DOBIČEK/IZGUBA":                                       "202",
    "25. BILANČNI DOBIČEK/IZGUBA":                                        "215",
}

# ── Podatkovne strukture ──────────────────────────────────────────────────────
SIZE_MAP = {
    "Mikro podjetje":"1","Majhno podjetje":"2",
    "Srednje podjetje":"3","Veliko podjetje":"4",
}

@dataclass
class CompanyInfo:
    name:str=""; registration_number:str=""; tax_number:str=""
    period_from:str=""; period_to:str=""
    tip_bilance:str=""; tip_subjekta:str=""; obdobje_bilance:str=""

@dataclass
class AopEntry:
    aop:str; current_year:Optional[float]=None; previous_year:Optional[float]=None

@dataclass
class ParseResult:
    company:CompanyInfo = field(default_factory=CompanyInfo)
    aop_data:dict       = field(default_factory=dict)
    warnings:list       = field(default_factory=list)
    errors:list         = field(default_factory=list)
    pdf_format:str      = ""

# ── Skupne pomožne funkcije ───────────────────────────────────────────────────
def parse_si(raw):
    if not raw: return None
    try: return float(str(raw).strip().replace(".","").replace(",","."))
    except: return None

def finalize_company(c):
    if c.period_to:
        ye = c.period_to.startswith('31.12')
        c.tip_bilance     = "8" if ye else "7"
        c.obdobje_bilance = "ZR" if ye else "MR"
    return c

def validate(result):
    if not result.company.name:
        result.errors.append("Ime podjetja ni najdeno.")
    if not result.company.registration_number:
        result.errors.append("Matična številka ni najdena.")
    for code in ["001","055"]:
        if code not in result.aop_data:
            result.errors.append(f"Manjka AOP {code} — preverite format PDF.")
    a1 = result.aop_data.get("001"); a5 = result.aop_data.get("055")
    if a1 and a5:
        v1 = a1.current_year or 0; v2 = a5.current_year or 0
        if abs(v1-v2) > 0.02:
            result.warnings.append(f"AOP 001 ≠ AOP 055 — preverite bilanco.")
    return result

# ── FORMAT DETEKCIJA ──────────────────────────────────────────────────────────
def detect_pdf_format(pages_text):
    """Vrne 'AOP' ali 'JOLP' glede na vsebino prve strani."""
    first = pages_text[0] if pages_text else ""
    # AOP format: prva vrstica je bold (4x ponovitev znakov)
    first_line = first.split('\n')[0] if first else ""
    if len(first_line) > 10:
        sample = first_line[:20].replace(" ","")
        chunks = [sample[i:i+4] for i in range(0,len(sample)-3,4)]
        if sum(1 for c in chunks if len(c)==4 and len(set(c))==1) >= 2:
            return "AOP"
    # JOLP format: "Bilanca stanja na dan" ali "Podatki so v EUR" na strani 1
    if "Bilanca stanja na dan" in first or "Podatki so v EUR" in first:
        return "JOLP"
    return "AOP"  # privzeto

# ── PARSER 1: AOP format (lp2025_...) ────────────────────────────────────────
def decode_bold(line):
    result=[]; i=0
    while i < len(line):
        ch=line[i]
        if ch==' ':
            while i<len(line) and line[i]==' ': i+=1
            result.append(' ')
        elif i+3<len(line) and line[i]==line[i+1]==line[i+2]==line[i+3]:
            result.append(ch); i+=4
        else:
            result.append(ch); i+=1
    return ''.join(result).strip()

def is_bold(line):
    s=line.strip()[:32].replace(" ","")
    if len(s)<8: return False
    chunks=[s[i:i+4] for i in range(0,len(s)-3,4)]
    return sum(1 for c in chunks if len(c)==4 and len(set(c))==1)>=2

AOP_RE = re.compile(r'(\d{3})\s+([\d.]+,\d{2})(?:\s+([\d.]+,\d{2}))?')

def parse_aop_format(pages_text):
    result = ParseResult(pdf_format="AOP")
    seen = set()
    full = "\n".join(pages_text)

    # Podatki podjetja
    c = CompanyInfo()
    m = re.search(r'Ime poslovnega subjekta[:\s]+([A-ZŠĐŽČĆ][A-ZŠĐŽČĆ\s\.\,\-]{2,60}?)(?=\n|\s{2,}|$)', full)
    if m: c.name = m.group(1).strip()
    m = re.search(r'Mati[cč]na [sš]tevilka\s+(\d{10})', full)
    if m: c.registration_number = m.group(1).strip()
    m = re.search(r'Dav[cč]na [sš]tevilka\s+(\d{8})', full)
    if m: c.tax_number = m.group(1).strip()
    m = re.search(r'\bod\s+([\d]{1,2}\.[\d]{1,2}\.[\d]{4})', full)
    if m: c.period_from = m.group(1).strip()
    m = re.search(r'\bdo\s+([\d]{1,2}\.[\d]{1,2}\.[\d]{4})', full)
    if m: c.period_to = m.group(1).strip()
    m = re.search(r'Velikost\s+([\w\s]+?)(?=\n)', full)
    if m: c.tip_subjekta = SIZE_MAP.get(m.group(1).strip(), "")
    result.company = finalize_company(c)

    # AOP vrednosti
    for page_text in pages_text[1:]:
        for line in page_text.split('\n'):
            decoded = decode_bold(line) if is_bold(line) else line
            m = AOP_RE.search(decoded)
            if not m: continue
            aop = m.group(1)
            if not (1 <= int(aop) <= 310): continue
            if aop in seen: continue
            result.aop_data[aop] = AopEntry(
                aop=aop,
                current_year=parse_si(m.group(2)),
                previous_year=parse_si(m.group(3)) if m.group(3) else None,
            )
            seen.add(aop)
    return result

# ── PARSER 2: JOLP format ─────────────────────────────────────────────────────
NUM_RE = re.compile(r'^(.+?)\s+([\d.]+,\d{2})\s+([\d.]+,\d{2})\s*$')

def parse_jolp_format(pages_text):
    result = ParseResult(pdf_format="JOLP")
    full = "\n".join(pages_text)

    # Podatki podjetja — iz prve strani
    lines0 = pages_text[0].split('\n') if pages_text else []
    c = CompanyInfo()
    # Ime: prva vrstica
    if lines0: c.name = lines0[0].strip()
    # Matična
    m = re.search(r'Mati[cč]na [sš]tevilka:\s*(\d{10})', full)
    if m: c.registration_number = m.group(1).strip()
    # Obdobje — iz "Bilanca stanja na dan 31.12.2025"
    m = re.search(r'na dan\s+([\d]{1,2}\.[\d]{1,2}\.[\d]{4})', full)
    if m: c.period_to = m.group(1).strip()
    m = re.search(r'od\s+([\d]{1,2}\.[01]?\d\.[\d]{4})\s+do', full)
    if m: c.period_from = m.group(1).strip()
    # Če period_from ni, ga izpeljemo iz period_to (1.1.LETO)
    if not c.period_from and c.period_to:
        year = c.period_to.split('.')[-1]
        c.period_from = f"1.1.{year}"
    # Tip subjekta — JOLP nima velikosti, damo privzeto "2" (majhno)
    c.tip_subjekta = "2"
    result.company = finalize_company(c)

    # Spoji prelomljene vrstice in parsiraj vrednosti
    all_lines = []
    for page_text in pages_text:
        all_lines.extend(page_text.split('\n'))

    # Spoji vrstice kjer je besedilo prelomljeno (vrednosti na naslednji vrstici)
    joined = []
    for line in all_lines:
        line = line.strip()
        if not line or line.startswith('Stran') or line in ('2025 2024','Podatki so v EUR s centi'):
            continue
        m = NUM_RE.match(line)
        if m:
            joined.append(line)
        else:
            # Verjetno nadaljevanje prejšnje vrstice
            if joined:
                test = joined[-1] + " " + line
                if NUM_RE.match(test):
                    joined[-1] = test
                elif not NUM_RE.match(joined[-1]):
                    joined[-1] = test
                else:
                    joined.append(line)
            else:
                joined.append(line)

    # Matchiraj na AOP kode
    for line in joined:
        m = NUM_RE.match(line)
        if not m: continue
        label = m.group(1).strip()
        curr  = parse_si(m.group(2))
        prev  = parse_si(m.group(3))

        # Poišči AOP kodo — partial match od začetka labela
        aop = None
        for key, code in JOLP_TO_AOP.items():
            if label.startswith(key) or key.startswith(label[:50]):
                aop = code
                break

        if aop and aop not in result.aop_data:
            result.aop_data[aop] = AopEntry(aop=aop, current_year=curr, previous_year=prev)

    return result

# ── Glavni parser ─────────────────────────────────────────────────────────────
def parse_pdf_file(pdf_path):
    pages_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            pages_text.append(page.extract_text() or "")

    fmt = detect_pdf_format(pages_text)

    if fmt == "JOLP":
        result = parse_jolp_format(pages_text)
    else:
        result = parse_aop_format(pages_text)

    return validate(result)

# ── Exporter ──────────────────────────────────────────────────────────────────
BS_AOP  = set([str(i).zfill(3) for i in range(1,97)] + ["301"])
IPI_AOP = set([str(i).zfill(3) for i in range(110,190)])
FMT     = '#,##0.00'

def norm_aop(raw):
    if raw is None: return None
    s = str(raw).strip()
    if not s or s in ("/","—"): return None
    try: return str(int(s)).zfill(3)
    except: return None

def build_idx(ws):
    idx = {}
    for row in ws.iter_rows(min_row=12):
        k = norm_aop(row[0].value)
        if k and k not in idx: idx[k] = row[0].row
    return idx

def export_excel(result, output_path):
    shutil.copy2(TEMPLATE_PATH, output_path)
    wb = openpyxl.load_workbook(output_path)
    c  = result.company

    ws = wb["BS"]
    ws.cell(3,4).value  = c.name
    ws.cell(4,4).value  = c.registration_number
    ws.cell(6,4).value  = c.obdobje_bilance
    ws.cell(5,7).value  = c.tip_bilance
    ws.cell(6,7).value  = c.tip_subjekta
    ws.cell(11,4).value = c.period_to
    idx = build_idx(ws); bs_n = 0
    for code in BS_AOP:
        if code not in idx or code not in result.aop_data: continue
        val = result.aop_data[code].current_year
        if val is not None:
            cell = ws.cell(idx[code],4)
            cell.value = val; cell.number_format = FMT; bs_n += 1

    ws2 = wb["IPI"]
    ws2.cell(3,4).value = c.name
    ws2.cell(4,4).value = c.registration_number
    if c.period_from and c.period_to:
        ws2.cell(8,3).value = f"{c.period_from} - {c.period_to}"
    ws2.cell(11,4).value = c.period_to
    idx2 = build_idx(ws2); ipi_n = 0
    for code in IPI_AOP:
        if code not in idx2 or code not in result.aop_data: continue
        val = result.aop_data[code].current_year
        if val is not None:
            cell2 = ws2.cell(idx2[code],4)
            cell2.value = val; cell2.number_format = FMT; ipi_n += 1

    wb.save(output_path)
    return bs_n, ipi_n


# ── PARSER 3: JOLP Excel format (.xlsx) ──────────────────────────────────────
def parse_xlsx_file(xlsx_path):
    result = ParseResult(pdf_format="XLSX")
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    c = CompanyInfo()
    c.tip_subjekta = "2"

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows(values_only=True):
            col0 = str(row[0]).strip() if row[0] is not None else ""
            col1 = row[1]
            col2 = row[2] if len(row) > 2 else None

            # Podatki podjetja
            if not c.name and col0 and col1 is None:
                skip = ("Matična","Bilanca","Izkaz","Podatki","Jamova","Ljubljana","Domžale","Maribor")
                if len(col0) > 5 and not any(col0.startswith(s) for s in skip):
                    c.name = col0

            m = re.search(r'Mati[cč]na [sš]tevilka:\s*(\d{10})', col0)
            if m: c.registration_number = m.group(1)

            m = re.search(r'na dan\s+([\d]{1,2}\.[\d]{1,2}\.[\d]{4})', col0)
            if m: c.period_to = m.group(1)

            m = re.search(r'v obdobju od ([\d]{1,2}\.[01]?\d\.[\d]{4})', col0)
            if m: c.period_from = m.group(1)

            # AOP vrednosti
            if col0 and isinstance(col1, (int, float)):
                aop = JOLP_TO_AOP.get(col0)
                if not aop:
                    for key, code in JOLP_TO_AOP.items():
                        if col0.startswith(key) or key.startswith(col0[:60]):
                            aop = code; break
                if aop and aop not in result.aop_data:
                    result.aop_data[aop] = AopEntry(
                        aop=aop,
                        current_year=float(col1) if col1 is not None else None,
                        previous_year=float(col2) if isinstance(col2,(int,float)) else None,
                    )

    if (not c.period_from or len(c.period_from) < 6) and c.period_to:
        year = c.period_to.split(".")[-1]
        c.period_from = f"1.1.{year}"

    result.company = finalize_company(c)
    return validate(result)

# ── HTML ──────────────────────────────────────────────────────────────────────
HTML = """<!DOCTYPE html>
<html lang="sl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AJPES Bilance Converter</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');
  :root{--bg:#f5f4f1;--surface:#fff;--border:#e2e0d8;--text:#1a1917;--text2:#5c5a54;--text3:#9c9a94;--slate:#334155;--emerald:#059669;--rose:#e11d48;--amber:#d97706;}
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:'DM Sans',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:24px;}
  .card{background:var(--surface);border:1px solid var(--border);border-radius:16px;padding:40px;width:100%;max-width:480px;box-shadow:0 4px 24px rgba(0,0,0,.07);}
  .logo{display:flex;align-items:center;gap:12px;margin-bottom:28px;}
  .logo-icon{width:40px;height:40px;background:var(--slate);border-radius:10px;display:flex;align-items:center;justify-content:center;}
  .logo-icon svg{width:22px;height:22px;stroke:white;fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;}
  .logo-text{font-size:16px;font-weight:600;letter-spacing:-.02em;}
  .logo-sub{font-size:11px;color:var(--text3);font-weight:400;letter-spacing:.04em;text-transform:uppercase;}
  h1{font-size:22px;font-weight:600;letter-spacing:-.02em;margin-bottom:6px;}
  .sub{font-size:13.5px;color:var(--text3);margin-bottom:28px;}
  label{font-size:12.5px;font-weight:600;color:var(--text);display:block;margin-bottom:6px;}
  input[type=password]{width:100%;padding:10px 13px;border:1px solid var(--border);border-radius:8px;font-size:14px;font-family:'DM Sans',sans-serif;outline:none;transition:border-color .15s;margin-bottom:16px;}
  input[type=password]:focus{border-color:var(--slate);box-shadow:0 0 0 3px rgba(51,65,85,.08);}
  .drop-zone{border:2px dashed var(--border);border-radius:14px;padding:40px 24px;text-align:center;cursor:pointer;transition:all .2s;margin-bottom:16px;background:#fafaf8;}
  .drop-zone:hover,.drop-zone.drag{border-color:var(--slate);background:#f0eff8;}
  .drop-zone.ready{border-color:var(--emerald);border-style:solid;background:#f0fdf4;}
  .drop-icon{width:44px;height:44px;background:var(--border);border-radius:10px;display:flex;align-items:center;justify-content:center;margin:0 auto 14px;transition:.2s;}
  .drop-zone:hover .drop-icon,.drop-zone.drag .drop-icon{background:var(--slate);}
  .drop-zone.ready .drop-icon{background:var(--emerald);}
  .drop-icon svg{width:22px;height:22px;stroke:var(--text3);fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;}
  .drop-zone:hover .drop-icon svg,.drop-zone.drag .drop-icon svg,.drop-zone.ready .drop-icon svg{stroke:white;}
  .drop-title{font-size:14px;font-weight:600;margin-bottom:4px;}
  .drop-desc{font-size:12.5px;color:var(--text3);}
  .tags{display:flex;gap:6px;justify-content:center;margin-top:12px;flex-wrap:wrap;}
  .tag{padding:3px 10px;background:var(--border);border-radius:20px;font-size:11.5px;color:var(--text2);}
  input[type=file]{display:none;}
  .btn{width:100%;padding:12px;background:var(--slate);color:white;border:none;border-radius:10px;font-size:14px;font-weight:600;font-family:'DM Sans',sans-serif;cursor:pointer;transition:background .15s;display:flex;align-items:center;justify-content:center;gap:8px;}
  .btn:hover{background:#1e293b;}
  .btn:disabled{background:var(--text3);cursor:not-allowed;}
  .btn svg{width:16px;height:16px;stroke:currentColor;fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;}
  .msg{padding:12px 16px;border-radius:10px;font-size:13px;margin-top:14px;border:1px solid;display:none;line-height:1.5;}
  .msg.error{background:#fff1f2;border-color:#fecdd3;color:var(--rose);}
  .spinner{width:16px;height:16px;border:2px solid rgba(255,255,255,.3);border-top-color:white;border-radius:50%;animation:spin .7s linear infinite;display:none;}
  @keyframes spin{to{transform:rotate(360deg);}}
  .result-box{display:none;margin-top:16px;padding:16px;background:var(--bg);border:1px solid var(--border);border-radius:10px;}
  .result-row{display:flex;justify-content:space-between;padding:5px 0;border-bottom:1px solid var(--border);font-size:13px;}
  .result-row:last-child{border-bottom:none;}
  .result-label{color:var(--text2);}
  .result-val{font-weight:500;font-family:'DM Mono',monospace;font-size:12.5px;}
  .warn-note{margin-top:12px;font-size:12px;color:var(--amber);background:#fffbeb;border:1px solid #fcd34d;border-radius:8px;padding:8px 12px;}
  .footer{text-align:center;font-size:11.5px;color:var(--text3);margin-top:20px;}
</style>
</head>
<body>
<div class="card">
  <div class="logo">
    <div class="logo-icon">
      <svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
      <polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>
    </div>
    <div>
      <div class="logo-text">AJPES Bilance Converter</div>
      <div class="logo-sub">Bančni interni tool</div>
    </div>
  </div>

  <div id="login-section">
    <h1>Prijava</h1>
    <p class="sub">Dostop samo za interne uporabnike</p>
    <label>Geslo</label>
    <input type="password" id="password" placeholder="••••••••••" onkeydown="if(event.key==='Enter')login()">
    <button class="btn" onclick="login()">
      <svg viewBox="0 0 24 24"><path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"/>
      <polyline points="10 17 15 12 10 7"/><line x1="15" y1="12" x2="3" y2="12"/></svg>
      Prijava
    </button>
    <div class="msg error" id="login-error">Napačno geslo. Poskusi znova.</div>
  </div>

  <div id="app-section" style="display:none;">
    <h1>Nova pretvorba</h1>
    <p class="sub">Naloži AJPES letno poročilo — oba formata sta podprta</p>
    <div class="drop-zone" id="drop-zone" onclick="document.getElementById('file-input').click()"
         ondragover="ev(event,'drag')" ondragleave="ev(event,'')" ondrop="drop(event)">
      <div class="drop-icon">
        <svg viewBox="0 0 24 24"><polyline points="16 16 12 12 8 16"/>
        <line x1="12" y1="12" x2="12" y2="21"/>
        <path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3"/></svg>
      </div>
      <div class="drop-title" id="drop-title">Povleci PDF sem ali klikni</div>
      <div class="drop-desc" id="drop-desc">PDF ali Excel · lp2025, JOLP — vsi formati podprti</div>
      <div class="tags"><span class="tag">PDF</span><span class="tag">XLSX</span><span class="tag">lp2025 · JOLP</span></div>
    </div>
    <input type="file" id="file-input" accept=".pdf,.xlsx,.xls" onchange="fileSelected(this.files[0])">
    <button class="btn" id="convert-btn" onclick="convert()" disabled>
      <div class="spinner" id="spinner"></div>
      <svg id="btn-icon" viewBox="0 0 24 24"><polyline points="23 4 23 10 17 10"/>
      <polyline points="1 20 1 14 7 14"/>
      <path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
      <span id="btn-text">Pretvori</span>
    </button>
    <div class="msg error" id="error-msg"></div>
    <div class="result-box" id="result-box">
      <div class="result-row"><span class="result-label">Format</span><span class="result-val" id="r-fmt"></span></div>
      <div class="result-row"><span class="result-label">Podjetje</span><span class="result-val" id="r-name"></span></div>
      <div class="result-row"><span class="result-label">Matična</span><span class="result-val" id="r-reg"></span></div>
      <div class="result-row"><span class="result-label">Obdobje</span><span class="result-val" id="r-period"></span></div>
      <div class="result-row"><span class="result-label">Tip bilance</span><span class="result-val" id="r-tip"></span></div>
      <div class="result-row"><span class="result-label">BS vrstice</span><span class="result-val" id="r-bs"></span></div>
      <div class="result-row"><span class="result-label">IPI vrstice</span><span class="result-val" id="r-ipi"></span></div>
      <div class="warn-note">⚠ Ne pozabi ročno vnesti <strong>Št. partnerja</strong> (celica D5 v BS sheetu)</div>
    </div>
  </div>
  <div class="footer">AJPES Bilance Converter · Samo za interno uporabo</div>
</div>
<script>
let selectedFile = null;
function login() {
  const pw = document.getElementById('password').value.trim();
  fetch('/auth',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({password:pw})})
  .then(r=>r.json()).then(d=>{
    if(d.ok){
      document.getElementById('login-section').style.display='none';
      document.getElementById('app-section').style.display='block';
    } else {
      document.getElementById('login-error').style.display='block';
    }
  }).catch(e=>{
    document.getElementById('login-error').textContent='Napaka: '+e.message;
    document.getElementById('login-error').style.display='block';
  });
}
function ev(e,cls){e.preventDefault();const z=document.getElementById('drop-zone');z.className='drop-zone'+(cls?' '+cls:(selectedFile?' ready':''));}
function drop(e){e.preventDefault();const f=e.dataTransfer.files[0];if(f)fileSelected(f);}
function fileSelected(file){
  if(!file)return;
  const ok=file.name.toLowerCase().endsWith('.pdf')||file.name.toLowerCase().endsWith('.xlsx');if(!ok){
    document.getElementById('error-msg').textContent='Naložiti moraš PDF ali Excel datoteko.';
    document.getElementById('error-msg').style.display='block';return;
  }
  selectedFile=file;
  document.getElementById('drop-zone').className='drop-zone ready';
  document.getElementById('drop-title').textContent='✓ '+file.name;
  document.getElementById('drop-desc').textContent=(file.size/1024/1024).toFixed(1)+' MB';
  document.getElementById('convert-btn').disabled=false;
  document.getElementById('error-msg').style.display='none';
  document.getElementById('result-box').style.display='none';
}
async function convert(){
  if(!selectedFile)return;
  const btn=document.getElementById('convert-btn');
  const spinner=document.getElementById('spinner');
  const icon=document.getElementById('btn-icon');
  const btnText=document.getElementById('btn-text');
  btn.disabled=true;spinner.style.display='block';icon.style.display='none';btnText.textContent='Pretvarjam...';
  document.getElementById('error-msg').style.display='none';
  document.getElementById('result-box').style.display='none';
  const fd=new FormData();fd.append('file',selectedFile);
  try{
    const res=await fetch('/convert',{method:'POST',body:fd});
    if(!res.ok){
      const err=await res.json().catch(()=>({error:'Neznana napaka.'}));
      document.getElementById('error-msg').textContent=err.error||'Napaka pri pretvorbi.';
      document.getElementById('error-msg').style.display='block';return;
    }
    const fmt=res.headers.get('X-PDF-Format')||'';
    const name=decodeURIComponent(res.headers.get('X-Company-Name')||'');
    const reg=res.headers.get('X-Registration')||'';
    const period=res.headers.get('X-Period')||'';
    const tip=res.headers.get('X-Tip-Bilance')||'';
    const bs=res.headers.get('X-BS-Filled')||'';
    const ipi=res.headers.get('X-IPI-Filled')||'';
    document.getElementById('r-fmt').textContent=fmt==='XLSX'?'Excel JOLP':fmt==='JOLP'?'PDF JOLP':'PDF AOP (standardni)';
    document.getElementById('r-name').textContent=name;
    document.getElementById('r-reg').textContent=reg;
    document.getElementById('r-period').textContent=period;
    document.getElementById('r-tip').textContent=tip==='8'?'8 — Zaključena (ZR)':'7 — Preliminarna (MR)';
    document.getElementById('r-bs').textContent=bs;
    document.getElementById('r-ipi').textContent=ipi;
    document.getElementById('result-box').style.display='block';
    const blob=await res.blob();
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    const stem=selectedFile.name.replace(/\.pdf$/i,'');
    a.href=url;a.download=stem+'_CBK.xlsx';
    document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);
  }catch(e){
    document.getElementById('error-msg').textContent='Napaka: '+e.message;
    document.getElementById('error-msg').style.display='block';
  }finally{
    btn.disabled=false;spinner.style.display='none';icon.style.display='block';btnText.textContent='Pretvori';
  }
}
</script>
</body>
</html>"""


# ── PARSER 3: XLSX JOLP format ────────────────────────────────────────────────
def parse_xlsx_file(xlsx_path):
    result = ParseResult(pdf_format="XLSX")
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    c = CompanyInfo()
    all_rows = []  # (label, current, previous)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows(values_only=True):
            if not row or row[0] is None:
                continue
            label = str(row[0]).strip()
            col1  = row[1] if len(row) > 1 else None
            col2  = row[2] if len(row) > 2 else None

            # Header podatki
            if label.startswith("Matična številka:"):
                c.registration_number = label.replace("Matična številka:", "").strip()
            elif "Bilanca stanja na dan" in label:
                m = re.search(r"(\d{1,2}\.\d{1,2}\.\d{4})", label)
                if m: c.period_to = m.group(1)
            elif "obdobju od" in label:
                m = re.search(r"od\s+([\d.]+)\s+do", label)
                if m: c.period_from = m.group(1)
            elif col1 in (2025, 2024) or col1 == "2025":
                pass  # header vrstica z letnicami
            elif isinstance(col1, (int, float)) and isinstance(col2, (int, float)):
                all_rows.append((label, float(col1), float(col2)))

        # Ime podjetja — prva neprazna vrstica ki ni naslov dokumenta
        if not c.name:
            for row in ws.iter_rows(min_row=1, max_row=20, values_only=True):
                val = str(row[0]).strip() if row[0] else ""
                if val and not any(x in val for x in
                        ["Bilanca", "Izkaz", "Podatki", "Matična", "None", "obdobju"]):
                    c.name = val
                    break

    # Izpelji period_from če manjka
    if (not c.period_from or len(c.period_from) < 6) and c.period_to:
        year = c.period_to.split(".")[-1]
        c.period_from = f"1.1.{year}"

    c.tip_subjekta = "2"
    result.company = finalize_company(c)

    # Mapiranje na AOP
    for label, curr, prev in all_rows:
        aop = None
        for key, code in JOLP_TO_AOP.items():
            if label.startswith(key) or key.startswith(label[:50]):
                aop = code
                break
        if aop and aop not in result.aop_data:
            result.aop_data[aop] = AopEntry(aop=aop, current_year=curr, previous_year=prev)

    return validate(result)

# ── Routes ────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template_string(HTML)

@app.route("/health")
def health():
    return jsonify({
        "status": "ok",
        "template_exists": TEMPLATE_PATH.exists(),
        "password_set": bool(ACCESS_PASSWORD),
        "formats_supported": ["AOP (lp2025_...)", "JOLP (JOLP_...)"],
    })

@app.route("/auth", methods=["POST"])
def auth():
    try:
        data = request.get_json(force=True, silent=True) or {}
        ok = str(data.get("password","")).strip() == ACCESS_PASSWORD
        return jsonify({"ok": ok})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/convert", methods=["POST"])
def convert():
    if "file" not in request.files:
        return jsonify({"error": "Ni datoteke."}), 400
    file = request.files["file"]
    fname = file.filename.lower()
    if not (fname.endswith(".pdf") or fname.endswith(".xlsx") or fname.endswith(".xls")):
        return jsonify({"error": "Podprte datoteke: PDF ali Excel (XLSX)."}), 400
    if not TEMPLATE_PATH.exists():
        return jsonify({"error": "Template ni najden na strežniku."}), 500

    conv_id    = str(uuid.uuid4())
    fname_low  = file.filename.lower()
    is_excel   = fname_low.endswith(".xlsx") or fname_low.endswith(".xls")
    input_ext  = ".xlsx" if is_excel else ".pdf"
    input_path = UPLOAD_DIR / f"{conv_id}{input_ext}"
    xlsx_path  = UPLOAD_DIR / f"{conv_id}_out.xlsx"

    try:
        file.save(input_path)
        if is_excel:
            result = parse_xlsx_file(input_path)
        else:
            result = parse_pdf_file(input_path)
        if result.errors:
            return jsonify({"error": "\n".join(result.errors)}), 422
        bs_n, ipi_n = export_excel(result, xlsx_path)
        c = result.company

        response = send_file(
            xlsx_path, as_attachment=True,
            download_name=file.filename.replace('.pdf','').replace('.PDF','') + '_CBK.xlsx',
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        safe_name = c.name.encode('ascii', errors='replace').decode('ascii')
        response.headers["X-PDF-Format"]    = result.pdf_format
        response.headers["X-Company-Name"]  = safe_name
        response.headers["X-Registration"]  = c.registration_number
        response.headers["X-Period"]        = f"{c.period_from} - {c.period_to}"
        response.headers["X-Tip-Bilance"]   = c.tip_bilance
        response.headers["X-BS-Filled"]     = str(bs_n)
        response.headers["X-IPI-Filled"]    = str(ipi_n)
        response.headers["Access-Control-Expose-Headers"] = \
            "X-PDF-Format,X-Company-Name,X-Registration,X-Period,X-Tip-Bilance,X-BS-Filled,X-IPI-Filled"
        return response
    except Exception as e:
        return jsonify({"error": f"Napaka: {str(e)}"}), 500
    finally:
        for p in [input_path, xlsx_path]:
            try: p.unlink()
            except: pass

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
# ── XLSX JOLP PARSER (doda se na konec obstoječih parserjev) ──────────────────
