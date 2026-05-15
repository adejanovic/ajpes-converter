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
    "SREDSTVA": "001",
    "A. DOLGOROČNA SREDSTVA": "002",
    "I. Neopredmetena sredstva in dolgoročne aktivne časovne razmejitve": "003",
    "1. Neopredmetena sredstva": "004",
    "2. Dolgoročne aktivne časovne razmejitve": "009",
    "II. Opredmetena osnovna sredstva": "010",
    "III. Naložbene nepremičnine": "018",
    "IV. Dolgoročne finančne naložbe": "019",
    "1. Dolgoročne finančne naložbe, razen posojil": "020",
    "2. Dolgoročna posojila": "024",
    "V. Dolgoročne poslovne terjatve": "027",
    "VI. Odložene terjatve za davek": "031",
    "B. KRATKOROČNA SREDSTVA": "032",
    "I. Sredstva (skupine za odtujitev) za prodajo": "033",
    "II. Zaloge": "034",
    "III. Kratkoročne finančne naložbe": "040",
    "1. Kratkoročne finančne naložbe, razen posojil": "041",
    "2. Kratkoročna posojila": "045",
    "IV. Kratkoročne poslovne terjatve": "048",
    "V. Denarna sredstva": "052",
    "C. KRATKOROČNE AKTIVNE ČASOVNE RAZMEJITVE": "053",
    "Zunajbilančna sredstva": "054",
    "OBVEZNOSTI DO VIROV SREDSTEV": "055",
    "A. KAPITAL": "056",
    "I. Vpoklicani kapital": "057",
    "1. Osnovni kapital": "058",
    "2. Nevpoklicani kapital (kot odbitna postavka)": "059",
    "II. Kapitalske rezerve": "060",
    "III. Rezerve iz dobička": "061",
    "IV. Revalorizacijske rezerve": "067",
    "V. Rezerve, nastale zaradi vrednotenja po pošteni vrednosti": "301",
    "VI. Preneseni čisti poslovni izid (preneseni čisti dobiček/izguba)": "068",
    "VII. Čisti poslovni izid poslovnega leta (čisti dobiček/čista izguba poslovnega leta)": "070",
    "B. REZERVACIJE IN DOLGOROČNE PASIVNE ČASOVNE RAZMEJITVE": "072",
    "1. Rezervacije": "073",
    "2. Dolgoročne pasivne časovne razmejitve": "074",
    "C. DOLGOROČNE OBVEZNOSTI": "075",
    "I. Dolgoročne finančne obveznosti": "076",
    "II. Dolgoročne poslovne obveznosti": "080",
    "III. Odložene obveznosti za davek": "084",
    "Č. KRATKOROČNE OBVEZNOSTI": "085",
    "I. Obveznosti, vključene v skupine za odtujitev": "086",
    "II. Kratkoročne finančne obveznosti": "087",
    "III. Kratkoročne poslovne obveznosti": "091",
    "D. KRATKOROČNE PASIVNE ČASOVNE RAZMEJITVE": "095",
    "Zunajbilančne obveznosti": "096",
    "1. ČISTI PRIHODKI OD PRODAJE": "110",
    "2. SPREMEMBA VREDNOSTI ZALOG PROIZVODOV IN NEDOKONČANE PROIZVODNJE": "121",
    "3. USREDSTVENI LASTNI PROIZVODI IN LASTNE STORITVE": "123",
    "4. DRUGI POSLOVNI PRIHODKI": "125",
    "5. Stroški blaga, materiala in storitev": "128",
    "a) Nabavna vrednost prodanega blaga in materiala ter stroški porabljenega materiala": "129",
    "b) Stroški storitev": "134",
    "6. Stroški dela": "139",
    "a) Stroški plač": "140",
    "b) Stroški pokojninskih zavarovanj": "141",
    "c) Stroški drugih socialnih zavarovanj": "142",
    "č) Drugi stroški dela": "143",
    "7. Odpisi vrednosti": "144",
    "a) Amortizacija": "145",
    "b) Prevrednotovalni poslovni odhodki pri neopredmetenih sredstvih in opredmetenih osnovnih sredstvih": "146",
    "c) Prevrednotovalni poslovni odhodki pri obratnih sredstvih": "147",
    "8. Drugi poslovni odhodki": "148",
    "9. Finančni prihodki iz deležev": "155",
    "10. Finančni prihodki iz danih posojil": "160",
    "11. Finančni prihodki iz poslovnih terjatev": "163",
    "12. Finančni odhodki iz oslabitve in odpisov finančnih naložb": "168",
    "13. Finančni odhodki iz finančnih obveznosti": "169",
    "14. Finančni odhodki iz poslovnih obveznosti": "174",
    "15. DRUGI PRIHODKI": "178",
    "16. DRUGI ODHODKI": "181",
    "17. DAVEK IZ DOBIČKA": "184",
    "18. ODLOŽENI DAVKI": "185",
    "19. ČISTI POSLOVNI IZID OBRAČUNSKEGA OBDOBJA (ČISTI DOBIČEK/IZGUBA OBRAČUNSKEGA OBDOBJA)": "186",
    "21. PRENESENI DOBIČEK/IZGUBA": "202",
    "25. BILANČNI DOBIČEK/IZGUBA": "215",
}

SIZE_MAP = {
    "Mikro podjetje":"1","Majhno podjetje":"2",
    "Srednje podjetje":"3","Veliko podjetje":"4",
}

TIP_BILANCE_MAP = {
    "7": "7",   # Preliminarna
    "8": "8",   # Zaključena
    "3": "3",   # Revidirana
}

# ── Podatkovne strukture ──────────────────────────────────────────────────────
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

# ── Skupne funkcije ───────────────────────────────────────────────────────────
def parse_si(raw):
    if not raw: return None
    try: return float(str(raw).strip().replace(".","").replace(",","."))
    except: return None

def _infer_period_from(period_to: str) -> str:
    """
    Izpelje začetek obdobja iz konca obdobja.
    31.12.2025 → 1.1.2025  (standardno leto)
    31.03.2026 → 1.4.2025  (nestandardno: april-mart, prejšnje leto)
    31.03.2025 → 1.4.2024
    Splošno: vzame dan+1 in mesec+1, leto prilagodi.
    """
    try:
        parts = period_to.split('.')
        day, month, year = int(parts[0]), int(parts[1]), int(parts[2])
        # Standardno leto: konec = 31.12
        if month == 12:
            return f"1.1.{year}"
        # Nestandardno: začetek = 1.(month+1).(year-1)
        # npr. konec 31.03.2026 → začetek 1.4.2025
        return f"1.{month+1}.{year-1}"
    except Exception:
        year = period_to.split('.')[-1] if '.' in period_to else period_to[-4:]
        return f"1.1.{year}"

def finalize_company(c, tip_override=None):
    if c.period_to:
        ye = c.period_to.startswith('31.12')
        # Tip bilance: uporabi override če ga imamo, drugače avtomatsko
        if tip_override and tip_override in TIP_BILANCE_MAP:
            c.tip_bilance = tip_override
        else:
            c.tip_bilance = "8" if ye else "7"
        c.obdobje_bilance = "ZR" if c.tip_bilance == "8" else "MR"
    return c

def validate(result):
    if not result.company.name:
        result.errors.append("Ime podjetja ni najdeno.")
    if not result.company.registration_number:
        result.errors.append("Matična številka ni najdena.")
    # Preveri 001/055 samo če je to BS datoteka (IPI-only je ok brez njiju)
    has_ipi = any(110 <= int(k) <= 199 for k in result.aop_data)
    has_bs  = any(1   <= int(k) <= 109 for k in result.aop_data)
    if has_bs or not has_ipi:
        for code in ["001","055"]:
            if code not in result.aop_data:
                result.errors.append(f"Manjka AOP {code}.")
    a1 = result.aop_data.get("001"); a5 = result.aop_data.get("055")
    if a1 and a5:
        v1 = a1.current_year or 0; v2 = a5.current_year or 0
        if abs(v1-v2) > 0.02:
            result.warnings.append("AOP 001 ≠ AOP 055 — preverite bilanco.")
    return result

# ── FORMAT DETEKCIJA ──────────────────────────────────────────────────────────
def detect_pdf_format(pages_text):
    first = pages_text[0] if pages_text else ""
    lines = first.split('\n')
    # NAPOVED: ima "AOP Konto POSTAVKA" v headerju
    if any("AOP Konto" in l or "AOP" == l.strip()[:3] for l in lines[:8]):
        # Preverimo ali so vrstice ki začnejo z AOP kodo
        for line in lines[5:10]:
            if re.match(r'^\d{3}\s+', line.strip()):
                return "NAPOVED"
    # AOP (lp2025): bold tekst — 4x ponavljanje
    first_line = lines[0] if lines else ""
    sample = first_line[:20].replace(" ","")
    chunks = [sample[i:i+4] for i in range(0,len(sample)-3,4)]
    if sum(1 for c in chunks if len(c)==4 and len(set(c))==1) >= 2:
        return "AOP"
    # JOLP: ime podjetja prva vrstica, "Bilanca stanja na dan" v prvih vrsticah
    if "Bilanca stanja na dan" in first or "Podatki so v EUR" in first:
        return "JOLP"
    return "AOP"

def detect_napoved_type(pages_text):
    """Vrne 'BS' ali 'IPI' za NAPOVED format."""
    first = pages_text[0] if pages_text else ""
    if "BILANCA STANJA" in first: return "BS"
    if "IZKAZ POSLOVNEGA IZIDA" in first: return "IPI"
    return "BS"

# ── PARSER: AOP format (lp2025) ──────────────────────────────────────────────
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
NAPOVED_RE = re.compile(r'^(\d{3})\s+.+?\s+([\d.]+,\d{2})\s+([\d.]+,\d{2})\s*$')

def parse_aop_format(pages_text, tip_override=None):
    result = ParseResult(pdf_format="AOP")
    seen = set()
    full = "\n".join(pages_text)
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
    # Tip subjekta je vedno 1 (gospodarska družba) — banka to nastavi posebej
    c.tip_subjekta = "1"
    result.company = finalize_company(c, tip_override)
    for page_text in pages_text[1:]:
        for line in page_text.split('\n'):
            decoded = decode_bold(line) if is_bold(line) else line
            m = AOP_RE.search(decoded)
            if not m: continue
            aop = m.group(1)
            if not (1 <= int(aop) <= 310): continue
            if aop in seen: continue
            result.aop_data[aop] = AopEntry(aop=aop,
                current_year=parse_si(m.group(2)),
                previous_year=parse_si(m.group(3)) if m.group(3) else None)
            seen.add(aop)
    return result

def parse_jolp_format(pages_text, tip_override=None):
    result = ParseResult(pdf_format="JOLP")
    full = "\n".join(pages_text)
    lines0 = pages_text[0].split('\n') if pages_text else []
    c = CompanyInfo()
    if lines0: c.name = lines0[0].strip()
    m = re.search(r'Mati[cč]na [sš]tevilka:\s*(\d{10})', full)
    if m: c.registration_number = m.group(1).strip()
    m = re.search(r'na dan\s+([\d]{1,2}\.[\d]{1,2}\.[\d]{4})', full)
    if m: c.period_to = m.group(1).strip()
    m = re.search(r'od\s+([\d]{1,2}\.[01]?\d\.[\d]{4})', full)
    if m: c.period_from = m.group(1).strip()
    if (not c.period_from or len(c.period_from) < 6) and c.period_to:

        c.period_from = _infer_period_from(c.period_to)
    c.tip_subjekta = "1"
    result.company = finalize_company(c, tip_override)
    all_lines = []
    for pt in pages_text:
        all_lines.extend(pt.split('\n'))
    NUM_RE = re.compile(r'^(.+?)\s+([\d.]+,\d{2})\s+([\d.]+,\d{2})\s*$')
    joined = []
    for line in all_lines:
        line = line.strip()
        if not line or line.startswith('Stran') or line in ('2025 2024','Podatki so v EUR s centi'): continue
        if NUM_RE.match(line):
            joined.append(line)
        else:
            if joined and not NUM_RE.match(joined[-1]):
                joined[-1] = joined[-1] + " " + line
            else:
                joined.append(line)
    for line in joined:
        m = NUM_RE.match(line)
        if not m: continue
        label = m.group(1).strip()
        aop = JOLP_TO_AOP.get(label)
        if not aop:
            for key, code in JOLP_TO_AOP.items():
                if label.startswith(key) or key.startswith(label[:50]):
                    aop = code; break
        if aop and aop not in result.aop_data:
            result.aop_data[aop] = AopEntry(aop=aop,
                current_year=parse_si(m.group(2)),
                previous_year=parse_si(m.group(3)))
    return result

def parse_napoved_format(pages_text, tip_override=None):
    result = ParseResult(pdf_format="NAPOVED")
    full = "\n".join(pages_text)
    lines = pages_text[0].split('\n') if pages_text else []
    c = CompanyInfo()
    # Ime in matična sta v prvi vrstici skupaj
    if lines:
        m = re.search(r'^(.+?)\s+Matična številka:\s*(\d+)', lines[0])
        if m:
            c.name = m.group(1).strip()
            reg = m.group(2).strip()
            # NAPOVED format ima 7-mestno matično — dodaj 000 do 10 mest
            c.registration_number = reg.ljust(10, '0') if len(reg) < 10 else reg
    m = re.search(r'Davčna številka:\s*(\d+)', full)
    if m: c.tax_number = m.group(1).strip()
    m = re.search(r'na dan\s+([\d]{1,2}\.[\d]{1,2}\.[\d]{4})', full)
    if m: c.period_to = m.group(1).strip()
    # IPI ima 'v obdobju od DD.MM.YYYY do DD.MM.YYYY'
    m = re.search(r'do\s+([\d]{1,2}\.[\d]{1,2}\.[\d]{4})', full)
    if m and not c.period_to: c.period_to = m.group(1).strip()
    m = re.search(r'od\s+([\d]{1,2}\.[01]?\d\.[\d]{4})\s+do', full)
    if m: c.period_from = m.group(1).strip()
    if (not c.period_from or len(c.period_from) < 6) and c.period_to:

        c.period_from = _infer_period_from(c.period_to)
    c.tip_subjekta = "1"
    result.company = finalize_company(c, tip_override)
    seen = set()
    for page_text in pages_text:
        for line in page_text.split('\n'):
            m = NAPOVED_RE.match(line.strip())
            if not m: continue
            aop = m.group(1)
            if not (1 <= int(aop) <= 310): continue
            if aop in seen: continue
            result.aop_data[aop] = AopEntry(aop=aop,
                current_year=parse_si(m.group(2)),
                previous_year=parse_si(m.group(3)))
            seen.add(aop)
    return result

def parse_xlsx_file(xlsx_path, tip_override=None):
    result = ParseResult(pdf_format="XLSX")
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    c = CompanyInfo(); c.tip_subjekta = "1"
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows(values_only=True):
            col0 = str(row[0]).strip() if row[0] is not None else ""
            col1 = row[1]; col2 = row[2] if len(row) > 2 else None
            if not c.name and col0 and col1 is None:
                skip = ("Matična","Bilanca","Izkaz","Podatki","Jamova","Ljubljana","Domžale","Maribor","Mozirje")
                if len(col0) > 5 and not any(col0.startswith(s) for s in skip):
                    c.name = col0
            m = re.search(r'Mati[cč]na [sš]tevilka:\s*(\d+)', col0)
            if m: c.registration_number = m.group(1)
            m = re.search(r'na dan\s+([\d]{1,2}\.[\d]{1,2}\.[\d]{4})', col0)
            if m: c.period_to = m.group(1)
            m = re.search(r'od\s+([\d]{1,2}\.[01]?\d\.[\d]{4})', col0)
            if m: c.period_from = m.group(1)
            if col0 and isinstance(col1, (int,float)):
                aop = JOLP_TO_AOP.get(col0)
                if not aop:
                    for key, code in JOLP_TO_AOP.items():
                        if col0.startswith(key) or key.startswith(col0[:60]):
                            aop = code; break
                if aop and aop not in result.aop_data:
                    result.aop_data[aop] = AopEntry(aop=aop,
                        current_year=float(col1) if col1 is not None else None,
                        previous_year=float(col2) if isinstance(col2,(int,float)) else None)
    if (not c.period_from or len(c.period_from) < 6) and c.period_to:

        c.period_from = _infer_period_from(c.period_to)
    result.company = finalize_company(c, tip_override)
    return validate(result)

def parse_pdf_file(pdf_path, tip_override=None):
    pages_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            pages_text.append(page.extract_text() or "")
    fmt = detect_pdf_format(pages_text)
    if fmt == "JOLP":
        result = parse_jolp_format(pages_text, tip_override)
    elif fmt == "NAPOVED":
        result = parse_napoved_format(pages_text, tip_override)
    else:
        result = parse_aop_format(pages_text, tip_override)
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
    c = result.company
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

# ── HTML ──────────────────────────────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="sl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AJPES Bilance Converter</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');
  :root{--bg:#f5f4f1;--surface:#fff;--border:#e2e0d8;--border2:#d0cec4;--text:#1a1917;--text2:#5c5a54;--text3:#9c9a94;--slate:#334155;--emerald:#059669;--rose:#e11d48;--amber:#d97706;--blue:#2563eb;}
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:'DM Sans',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:24px;}
  .card{background:var(--surface);border:1px solid var(--border);border-radius:16px;padding:36px;width:100%;max-width:520px;box-shadow:0 4px 24px rgba(0,0,0,.07);}
  .logo{display:flex;align-items:center;gap:12px;margin-bottom:24px;}
  .logo-icon{width:40px;height:40px;background:var(--slate);border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0;}
  .logo-icon svg{width:22px;height:22px;stroke:white;fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;}
  .logo-text{font-size:16px;font-weight:600;letter-spacing:-.02em;}
  .logo-sub{font-size:11px;color:var(--text3);font-weight:400;letter-spacing:.04em;text-transform:uppercase;}
  h1{font-size:21px;font-weight:600;letter-spacing:-.02em;margin-bottom:5px;}
  .sub{font-size:13px;color:var(--text3);margin-bottom:24px;}
  label{font-size:12.5px;font-weight:600;color:var(--text);display:block;margin-bottom:6px;}
  input[type=password]{width:100%;padding:10px 13px;border:1px solid var(--border);border-radius:8px;font-size:14px;font-family:'DM Sans',sans-serif;outline:none;transition:border-color .15s;margin-bottom:16px;}
  input[type=password]:focus{border-color:var(--slate);box-shadow:0 0 0 3px rgba(51,65,85,.08);}

  /* Tip bilance selector */
  .section-label{font-size:12.5px;font-weight:600;color:var(--text);margin-bottom:8px;}
  .tip-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:20px;}
  .tip-btn{padding:10px 8px;border:1.5px solid var(--border);border-radius:10px;cursor:pointer;text-align:center;transition:all .15s;background:var(--surface);}
  .tip-btn:hover{border-color:var(--border2);background:var(--bg);}
  .tip-btn.selected{border-color:var(--slate);background:var(--slate);color:white;}
  .tip-btn .tip-code{font-size:20px;font-weight:700;font-family:'DM Mono',monospace;display:block;margin-bottom:2px;}
  .tip-btn .tip-name{font-size:11px;font-weight:500;opacity:.8;}
  .tip-btn.selected .tip-name{opacity:.7;}
  .obdobje-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:8px;margin-bottom:20px;}
  .obdobje-btn{padding:10px 12px;border:1.5px solid var(--border);border-radius:10px;cursor:pointer;transition:all .15s;background:var(--surface);display:flex;flex-direction:column;gap:2px;}
  .obdobje-btn:hover{border-color:var(--border2);background:var(--bg);}
  .obdobje-btn.selected{border-color:var(--slate);background:var(--slate);color:white;}
  .obdobje-icon{font-size:14px;}
  .obdobje-main{font-size:13.5px;font-weight:600;font-family:'DM Mono',monospace;}
  .obdobje-sub{font-size:11px;opacity:.7;}

  /* Upload zone */
  .drop-zone{border:2px dashed var(--border);border-radius:12px;padding:24px;text-align:center;cursor:pointer;transition:all .2s;margin-bottom:8px;background:#fafaf8;position:relative;}
  .drop-zone:hover,.drop-zone.drag{border-color:var(--slate);background:#f0eff8;}
  .drop-zone.has-file{border-color:var(--emerald);border-style:solid;background:#f0fdf4;}
  .drop-icon{width:40px;height:40px;background:var(--border);border-radius:9px;display:flex;align-items:center;justify-content:center;margin:0 auto 12px;transition:.2s;}
  .drop-zone:hover .drop-icon,.drop-zone.drag .drop-icon{background:var(--slate);}
  .drop-zone.has-file .drop-icon{background:var(--emerald);}
  .drop-icon svg{width:20px;height:20px;stroke:var(--text3);fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;}
  .drop-zone:hover .drop-icon svg,.drop-zone.drag .drop-icon svg,.drop-zone.has-file .drop-icon svg{stroke:white;}
  .drop-title{font-size:13.5px;font-weight:600;margin-bottom:3px;}
  .drop-desc{font-size:12px;color:var(--text3);}
  .drop-hint{font-size:11.5px;color:var(--text3);margin-bottom:16px;text-align:center;}
  .drop-hint a{color:var(--slate);text-decoration:none;font-weight:500;cursor:pointer;}
  .drop-hint a:hover{text-decoration:underline;}
  .tags{display:flex;gap:5px;justify-content:center;margin-top:10px;flex-wrap:wrap;}
  .tag{padding:2px 9px;background:var(--border);border-radius:20px;font-size:11px;color:var(--text2);}
  input[type=file]{display:none;}

  /* File list */
  .file-list{margin-bottom:16px;}
  .file-item{display:flex;align-items:center;gap:10px;padding:8px 12px;background:var(--bg);border:1px solid var(--border);border-radius:8px;margin-bottom:6px;font-size:13px;}
  .file-item-icon{width:28px;height:28px;background:var(--emerald);border-radius:6px;display:flex;align-items:center;justify-content:center;flex-shrink:0;}
  .file-item-icon svg{width:14px;height:14px;stroke:white;fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;}
  .file-item-name{flex:1;font-weight:500;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
  .file-item-size{color:var(--text3);font-size:12px;flex-shrink:0;}
  .file-item-remove{width:22px;height:22px;border-radius:50%;border:none;background:var(--border);cursor:pointer;display:flex;align-items:center;justify-content:center;flex-shrink:0;transition:.15s;}
  .file-item-remove:hover{background:var(--rose);color:white;}
  .file-item-remove svg{width:12px;height:12px;stroke:currentColor;fill:none;stroke-width:2.5;stroke-linecap:round;}

  /* Divider */
  .divider{display:flex;align-items:center;gap:10px;margin:16px 0;}
  .divider-line{flex:1;height:1px;background:var(--border);}
  .divider-text{font-size:11.5px;color:var(--text3);}

  /* Btn */
  .btn{width:100%;padding:12px;background:var(--slate);color:white;border:none;border-radius:10px;font-size:14px;font-weight:600;font-family:'DM Sans',sans-serif;cursor:pointer;transition:background .15s;display:flex;align-items:center;justify-content:center;gap:8px;}
  .btn:hover{background:#1e293b;}
  .btn:disabled{background:var(--text3);cursor:not-allowed;}
  .btn svg{width:16px;height:16px;stroke:currentColor;fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;}
  .spinner{width:16px;height:16px;border:2px solid rgba(255,255,255,.3);border-top-color:white;border-radius:50%;animation:spin .7s linear infinite;display:none;}
  @keyframes spin{to{transform:rotate(360deg);}}

  /* Messages */
  .msg{padding:12px 16px;border-radius:10px;font-size:13px;margin-top:14px;border:1px solid;display:none;line-height:1.5;}
  .msg.error{background:#fff1f2;border-color:#fecdd3;color:var(--rose);}

  /* Result */
  .result-box{display:none;margin-top:16px;background:var(--bg);border:1px solid var(--border);border-radius:12px;overflow:hidden;}
  .result-header{padding:12px 16px;background:var(--emerald);display:flex;align-items:center;gap:8px;}
  .result-header svg{width:16px;height:16px;stroke:white;fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;}
  .result-header-text{font-size:13.5px;font-weight:600;color:white;}
  .result-body{padding:14px 16px;}
  .result-row{display:flex;justify-content:space-between;align-items:center;padding:5px 0;border-bottom:1px solid var(--border);font-size:13px;}
  .result-row:last-child{border-bottom:none;}
  .result-label{color:var(--text2);}
  .result-val{font-weight:500;font-family:'DM Mono',monospace;font-size:12.5px;text-align:right;}
  .warn-note{margin-top:12px;font-size:12px;color:var(--amber);background:#fffbeb;border:1px solid #fcd34d;border-radius:8px;padding:9px 12px;line-height:1.5;}
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

  <!-- LOGIN -->
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
    <div class="msg error" id="login-error">Napačno geslo.</div>
  </div>

  <!-- APP -->
  <div id="app-section" style="display:none;">
    <h1>Nova pretvorba</h1>
    <p class="sub">Izberi tip bilance, naloži datoteko in pridobi izpolnjen Excel</p>

    <!-- KORAK 1: Tip bilance -->
    <div class="section-label">1. Tip bilance</div>
    <div class="tip-grid">
      <div class="tip-btn" onclick="selectTip(this,'7')">
        <span class="tip-code">7</span>
        <span class="tip-name">Preliminarna</span>
      </div>
      <div class="tip-btn selected" onclick="selectTip(this,'8')">
        <span class="tip-code">8</span>
        <span class="tip-name">Zaključena</span>
      </div>
      <div class="tip-btn" onclick="selectTip(this,'3')">
        <span class="tip-code">3</span>
        <span class="tip-name">Revidirana</span>
      </div>
    </div>

    <!-- Obdobje -->
    <div class="section-label">2. Poslovno obdobje</div>
    <div class="obdobje-grid">
      <div class="obdobje-btn selected" onclick="selectObdobje(this,'01.01-31.12')">
        <span class="obdobje-main">01.01. &#8211; 31.12.</span>
        <span class="obdobje-sub">Standardno poslovno leto</span>
      </div>
      <div class="obdobje-btn" onclick="selectObdobje(this,'01.04-31.03')">
        <span class="obdobje-main">01.04. &#8211; 31.03.</span>
        <span class="obdobje-sub">Nestandardno poslovno leto</span>
      </div>
    </div>

    <!-- KORAK 3: Datoteke -->
    <div class="section-label">3. Datoteka(-e)</div>
    <div class="drop-zone" id="drop-zone"
         onclick="document.getElementById('file-input').click()"
         ondragover="ev(event,'drag')" ondragleave="ev(event,'')" ondrop="drop(event)">
      <div class="drop-icon">
        <svg viewBox="0 0 24 24"><polyline points="16 16 12 12 8 16"/>
        <line x1="12" y1="12" x2="12" y2="21"/>
        <path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3"/></svg>
      </div>
      <div class="drop-title">Povleci datoteke sem ali klikni</div>
      <div class="drop-desc">En fajl ali oba (BS + IPI ločena) — PDF ali Excel</div>
      <div class="tags">
        <span class="tag">lp2025 PDF</span>
        <span class="tag">JOLP PDF</span>
        <span class="tag">JOLP Excel</span>
        <span class="tag">NAPOVED PDF</span>
      </div>
    </div>
    <input type="file" id="file-input" accept=".pdf,.xlsx" multiple onchange="filesSelected(this.files)">
    <div class="drop-hint">Preliminarne bilance imajo pogosto ločena PDF za BS in IPI — <a onclick="document.getElementById('file-input').click()">dodaj oba</a></div>

    <!-- Seznam datotek -->
    <div class="file-list" id="file-list"></div>

    <!-- KORAK 4: Pretvori -->
    <button class="btn" id="convert-btn" onclick="convert()" disabled>
      <div class="spinner" id="spinner"></div>
      <svg id="btn-icon" viewBox="0 0 24 24"><polyline points="23 4 23 10 17 10"/>
      <polyline points="1 20 1 14 7 14"/>
      <path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
      <span id="btn-text">Pretvori</span>
    </button>

    <div class="msg error" id="error-msg"></div>

    <!-- Rezultat -->
    <div class="result-box" id="result-box">
      <div class="result-header">
        <svg viewBox="0 0 24 24"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
        <polyline points="22 4 12 14.01 9 11.01"/></svg>
        <span class="result-header-text">Pretvorba uspešna — Excel se prenaša</span>
      </div>
      <div class="result-body">
        <div class="result-row"><span class="result-label">Podjetje</span><span class="result-val" id="r-name"></span></div>
        <div class="result-row"><span class="result-label">Matična</span><span class="result-val" id="r-reg"></span></div>
        <div class="result-row"><span class="result-label">Obdobje</span><span class="result-val" id="r-period"></span></div>
        <div class="result-row"><span class="result-label">Tip bilance</span><span class="result-val" id="r-tip"></span></div>
        <div class="result-row"><span class="result-label">Format vhoda</span><span class="result-val" id="r-fmt"></span></div>
        <div class="result-row"><span class="result-label">BS vrstice</span><span class="result-val" id="r-bs"></span></div>
        <div class="result-row"><span class="result-label">IPI vrstice</span><span class="result-val" id="r-ipi"></span></div>
        <div class="warn-note">⚠ Ne pozabi ročno vnesti <strong>Št. partnerja</strong> (celica D5 v BS sheetu)</div>
      </div>
    </div>
  </div>
  <div class="footer">AJPES Bilance Converter · Samo za interno uporabo</div>
</div>

<script>
let selectedFiles = [];
let selectedTip = '8';
let selectedObdobje = '01.01-31.12';

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

function selectObdobje(el, val) {
  selectedObdobje = val;
  document.querySelectorAll('.obdobje-btn').forEach(b => b.classList.remove('selected'));
  el.classList.add('selected');
}

function selectTip(el, val) {
  selectedTip = val;
  document.querySelectorAll('.tip-btn').forEach(b => b.classList.remove('selected'));
  el.classList.add('selected');
}

function ev(e, cls) {
  e.preventDefault();
  const z = document.getElementById('drop-zone');
  z.className = 'drop-zone' + (cls ? ' '+cls : (selectedFiles.length ? ' has-file' : ''));
}

function drop(e) {
  e.preventDefault();
  filesSelected(e.dataTransfer.files);
}

function filesSelected(fileList) {
  for (let f of fileList) {
    const ext = f.name.toLowerCase();
    if (!ext.endsWith('.pdf') && !ext.endsWith('.xlsx')) {
      showError('Dovoljeni so samo PDF ali Excel (.xlsx) fajli.');
      return;
    }
    // Ne dodaj duplikatov
    if (!selectedFiles.find(x => x.name === f.name)) {
      selectedFiles.push(f);
    }
  }
  renderFileList();
  document.getElementById('error-msg').style.display = 'none';
  document.getElementById('result-box').style.display = 'none';
  document.getElementById('drop-zone').className = 'drop-zone' + (selectedFiles.length ? ' has-file' : '');
  document.getElementById('convert-btn').disabled = selectedFiles.length === 0;
}

function removeFile(idx) {
  selectedFiles.splice(idx, 1);
  renderFileList();
  document.getElementById('drop-zone').className = 'drop-zone' + (selectedFiles.length ? ' has-file' : '');
  document.getElementById('convert-btn').disabled = selectedFiles.length === 0;
}

function renderFileList() {
  const list = document.getElementById('file-list');
  if (selectedFiles.length === 0) { list.innerHTML = ''; return; }
  list.innerHTML = selectedFiles.map((f, i) => `
    <div class="file-item">
      <div class="file-item-icon">
        <svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
        <polyline points="14 2 14 8 20 8"/></svg>
      </div>
      <span class="file-item-name" title="${f.name}">${f.name}</span>
      <span class="file-item-size">${(f.size/1024/1024).toFixed(1)} MB</span>
      <button class="file-item-remove" onclick="removeFile(${i})" title="Odstrani">
        <svg viewBox="0 0 24 24"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
      </button>
    </div>
  `).join('');
}

function showError(msg) {
  const e = document.getElementById('error-msg');
  e.textContent = msg; e.style.display = 'block';
}

const TIP_LABELS = {'7':'7 — Preliminarna (MR)','8':'8 — Zaključena (ZR)','3':'3 — Revidirana (ZR)'};

async function convert() {
  if (selectedFiles.length === 0) return;
  const btn = document.getElementById('convert-btn');
  const spinner = document.getElementById('spinner');
  const icon = document.getElementById('btn-icon');
  const btnText = document.getElementById('btn-text');
  btn.disabled=true; spinner.style.display='block'; icon.style.display='none'; btnText.textContent='Pretvarjam...';
  document.getElementById('error-msg').style.display='none';
  document.getElementById('result-box').style.display='none';

  const fd = new FormData();
  for (let f of selectedFiles) fd.append('files', f);
  fd.append('tip_bilance', selectedTip);
  fd.append('obdobje', selectedObdobje);

  try {
    const res = await fetch('/convert', {method:'POST', body:fd});
    if (!res.ok) {
      const err = await res.json().catch(()=>({error:'Neznana napaka.'}));
      showError(err.error || 'Napaka pri pretvorbi.'); return;
    }
    const name   = decodeURIComponent(res.headers.get('X-Company-Name')||'');
    const reg    = res.headers.get('X-Registration')||'';
    const period = res.headers.get('X-Period')||'';
    const tip    = res.headers.get('X-Tip-Bilance')||'';
    const fmt    = res.headers.get('X-PDF-Format')||'';
    const bs     = res.headers.get('X-BS-Filled')||'';
    const ipi    = res.headers.get('X-IPI-Filled')||'';
    document.getElementById('r-name').textContent   = name;
    document.getElementById('r-reg').textContent    = reg;
    document.getElementById('r-period').textContent = period;
    document.getElementById('r-tip').textContent    = TIP_LABELS[tip] || tip;
    document.getElementById('r-fmt').textContent    = fmt;
    document.getElementById('r-bs').textContent     = bs;
    document.getElementById('r-ipi').textContent    = ipi;
    document.getElementById('result-box').style.display = 'block';
    const blob = await res.blob();
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    const stem = selectedFiles[0].name.replace(/\.(pdf|xlsx)$/i,'');
    a.href=url; a.download=stem+'_CBK.xlsx';
    document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
  } catch(e) {
    showError('Napaka: '+e.message);
  } finally {
    btn.disabled=false; spinner.style.display='none'; icon.style.display='block'; btnText.textContent='Pretvori';
  }
}
</script>
</body>
</html>"""

# ── Routes ────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template_string(HTML)

@app.route("/health")
def health():
    return jsonify({"status":"ok","template_exists":TEMPLATE_PATH.exists(),
                    "formats":["AOP","JOLP","NAPOVED","XLSX"]})

@app.route("/auth", methods=["POST"])
def auth():
    try:
        data = request.get_json(force=True, silent=True) or {}
        ok = str(data.get("password","")).strip() == ACCESS_PASSWORD
        return jsonify({"ok": ok})
    except Exception as e:
        return jsonify({"ok":False,"error":str(e)}), 500

@app.route("/convert", methods=["POST"])
def convert():
    files = request.files.getlist("files")
    if not files or all(f.filename == "" for f in files):
        return jsonify({"error": "Ni datoteke."}), 400
    if not TEMPLATE_PATH.exists():
        return jsonify({"error": "Template ni najden na strežniku."}), 500

    tip_override = request.form.get("tip_bilance", "")
    conv_id = str(uuid.uuid4())
    saved_paths = []

    try:
        # Shrani vse naložene datoteke
        for f in files:
            if not f.filename: continue
            fname = f.filename.lower()
            if not (fname.endswith('.pdf') or fname.endswith('.xlsx')):
                return jsonify({"error": f"Nepodprt format: {f.filename}"}), 400
            ext = '.xlsx' if fname.endswith('.xlsx') else '.pdf'
            p = UPLOAD_DIR / f"{conv_id}_{len(saved_paths)}{ext}"
            f.save(p)
            saved_paths.append(p)

        # Parsiraj — združi rezultate če je več fajlov
        merged = ParseResult()
        fmt_used = ""

        for path in saved_paths:
            if path.suffix == '.xlsx':
                r = parse_xlsx_file(path, tip_override)
            else:
                r = parse_pdf_file(path, tip_override)

            if r.errors:
                return jsonify({"error": "\n".join(r.errors)}), 422

            # Prva datoteka dobi podatke podjetja
            if not merged.company.name:
                merged.company = r.company
                fmt_used = r.pdf_format

            # Združi AOP podatke (ne prepiši obstoječih)
            for code, entry in r.aop_data.items():
                if code not in merged.aop_data:
                    merged.aop_data[code] = entry

            merged.warnings.extend(r.warnings)

        if not merged.aop_data:
            return jsonify({"error": "Ni bilo mogoče prebrati podatkov iz datotek."}), 422

        # Izvozi
        out_path = UPLOAD_DIR / f"{conv_id}_out.xlsx"
        bs_n, ipi_n = export_excel(merged, out_path)
        c = merged.company

        stem = saved_paths[0].stem.split('_')[0] if saved_paths else "bilanca"
        response = send_file(out_path, as_attachment=True,
            download_name=f"{stem}_CBK.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        safe_name = c.name.encode('ascii', errors='replace').decode('ascii')
        response.headers["X-Company-Name"] = safe_name
        response.headers["X-Registration"] = c.registration_number
        response.headers["X-Period"]       = f"{c.period_from} - {c.period_to}"
        response.headers["X-Tip-Bilance"]  = c.tip_bilance
        response.headers["X-PDF-Format"]   = fmt_used
        response.headers["X-BS-Filled"]    = str(bs_n)
        response.headers["X-IPI-Filled"]   = str(ipi_n)
        response.headers["Access-Control-Expose-Headers"] = \
            "X-Company-Name,X-Registration,X-Period,X-Tip-Bilance,X-PDF-Format,X-BS-Filled,X-IPI-Filled"
        return response

    except Exception as e:
        return jsonify({"error": f"Napaka: {str(e)}"}), 500
    finally:
        for p in saved_paths:
            try: p.unlink()
            except: pass
        try: (UPLOAD_DIR / f"{conv_id}_out.xlsx").unlink()
        except: pass

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
