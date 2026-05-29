from flask import Flask, request, send_file, render_template_string, jsonify
import os, re, shutil, tempfile, uuid, subprocess
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional

try:
    import pdfplumber
    import openpyxl
    try:
        import xlrd
    except ImportError:
        xlrd = None
except ImportError:
    pass

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

TEMPLATE_PATH    = Path(__file__).parent / "template.xlsx"
TEMPLATE_SP_PATH = Path(__file__).parent / "template_sp.xlsx"
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
    "3": "3",   # Revidirana
    "4": "4",   # Zaključena nerevidirana
    "5": "5",   # Konsolidirana revidirana
    "6": "6",   # Konsolidirana nerevidirana
    "7": "7",   # Preliminarna
    "8": "8",   # Zaključena
}

TIP_BILANCE_LABELS = {
    "3": "REVIDIRANA",
    "4": "ZAKLJUCENA_NEREVIDIRANA",
    "5": "KONSOLIDIRANA_REVIDIRANA",
    "6": "KONSOLIDIRANA_NEREVIDIRANA",
    "7": "PRELIMINARNA",
    "8": "ZAKLJUCENA",
}

TIP_SUBJEKTA_MAP = {
    "1": "Gospodarske družbe",
    "2": "Samostojni podjetniki",
    "3": "Zadruge",
    "4": "Društva",
    "5": "Pravne osebe javnega prava",
    "6": "Pravne osebe zasebnega prava",
    "7": "Banke",
    "8": "Zavarovalnice",
    "9": "Druge osebe javnega prava",
}

DRUSTVO_KEYWORDS = [
    "SKLAD", "Društveni sklad", "PRESEŽEK POSLOVNIH PRIHODKOV",
    "PRESEŽEK POSLOVNIH ODHODKOV", "društvo", "Društvo", "DRUŠTVO"
]

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
    subject_type:str    = "GD"   # "GD" = Gospodarska družba, "DR" = Društvo

# ── Skupne funkcije ───────────────────────────────────────────────────────────
def parse_si(raw):
    if not raw: return None
    try: return float(str(raw).strip().replace(".","").replace(",","."))
    except: return None

def _detect_subject_type_from_text(text: str) -> str:
    """Zazna ali gre za društvo ali gospodarsko družbo iz vsebine Excel/PDF."""
    for kw in DRUSTVO_KEYWORDS:
        if kw in text:
            return "DR"
    return "GD"

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

def finalize_company(c, tip_override=None, tip_subjekta_override=None):
    if tip_subjekta_override and tip_subjekta_override in TIP_SUBJEKTA_MAP:
        c.tip_subjekta = tip_subjekta_override
    elif not c.tip_subjekta:
        c.tip_subjekta = "1"

    if c.period_to and (not c.period_from or len(str(c.period_from)) < 6):
        c.period_from = _infer_period_from(c.period_to)

    if c.period_to:
        ye = c.period_to.startswith('31.12')
        # Tip bilance: uporabi override če ga imamo, drugače avtomatsko
        if tip_override and tip_override in TIP_BILANCE_MAP:
            c.tip_bilance = tip_override
        else:
            c.tip_bilance = "8" if ye else "7"
        # 7 = preliminarna/medletna, ostali tipi iz template-a so zaključni računi
        c.obdobje_bilance = "MR" if c.tip_bilance == "7" else "ZR"
    return c

def _subject_type_from_tip_subjekta(tip_subjekta: str) -> str:
    return "SP" if str(tip_subjekta).strip() == "2" else "GD"

def is_cid_encoded(pages_text):
    """Zazna PDFje z pokvarjenim font encodingom (cid:XX) ki jih ne moremo brati."""
    if not pages_text: return False
    sample = " ".join(pages_text[:3])
    cid_count = sample.count("(cid:")
    total_chars = len(sample.replace(" ",""))
    return cid_count > 10 and (cid_count * 6) > (total_chars * 0.3)

def validate(result):
    if not result.company.name:
        result.errors.append("Ime podjetja ni najdeno.")
    if not result.company.registration_number:
        result.errors.append("Matična številka ni najdena.")
    # Preveri 001/055 samo če je to BS datoteka (IPI-only je ok brez njiju)
    def _safe_aop_int(k):
        try: return int(str(k).rstrip('abcdefghijklmnopqrstuvwxyz'))
        except: return 0
    has_ipi = any(110 <= _safe_aop_int(k) <= 199 for k in result.aop_data)
    has_bs  = any(1   <= _safe_aop_int(k) <= 109 for k in result.aop_data)
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

def parse_aop_format(pages_text, tip_override=None, tip_subjekta_override=None):
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
    result.company = finalize_company(c, tip_override, tip_subjekta_override)
    result.subject_type = _subject_type_from_tip_subjekta(result.company.tip_subjekta)
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

def parse_jolp_format(pages_text, tip_override=None, tip_subjekta_override=None):
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
    result.company = finalize_company(c, tip_override, tip_subjekta_override)
    result.subject_type = _subject_type_from_tip_subjekta(result.company.tip_subjekta)
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

def parse_napoved_format(pages_text, tip_override=None, tip_subjekta_override=None):
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
    result.company = finalize_company(c, tip_override, tip_subjekta_override)
    result.subject_type = _subject_type_from_tip_subjekta(result.company.tip_subjekta)
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

def _cell_to_text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()

def _cell_to_number(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    if isinstance(v, str):
        return parse_si(v)
    return None

def _aop_from_cell(v):
    if v is None:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if float(v).is_integer():
            n = int(v)
            if 100 <= n <= 499:   # prepreči headerje 1,2,3,4,5
                return str(n).zfill(3)
        return None
    s = str(v).strip()
    m = re.fullmatch(r'(\d{3})([a-z])?', s, flags=re.I)
    if not m:
        return None
    n = int(m.group(1))
    if not (1 <= n <= 499):
        return None
    return m.group(1).zfill(3) + (m.group(2).lower() if m.group(2) else "")

def _find_value_by_label(rows_by_sheet, label_patterns):
    patterns = [re.compile(p, re.I) for p in label_patterns]
    for rows in rows_by_sheet.values():
        for r_idx, row in enumerate(rows):
            for c_idx, cell in enumerate(row):
                txt = _cell_to_text(cell)
                if not txt:
                    continue
                if any(p.search(txt) for p in patterns):
                    # 1) desno v isti vrstici
                    for j in range(c_idx + 1, min(len(row), c_idx + 5)):
                        val = _cell_to_text(row[j])
                        if val:
                            return val
                    # 2) spodaj v istem stolpcu ali blizu stolpca
                    for rr in range(r_idx + 1, min(len(rows), r_idx + 5)):
                        for cc in range(max(0, c_idx - 1), min(len(rows[rr]), c_idx + 4)):
                            val = _cell_to_text(rows[rr][cc])
                            if val:
                                return val
    return ""

def _extract_company_from_excel_rows(rows_by_sheet):
    c = CompanyInfo(); c.tip_subjekta = "1"
    c.name = _find_value_by_label(rows_by_sheet, [r'ime poslovnega subjekta', r'ime podjetnika'])
    c.registration_number = re.sub(r'\D', '', _find_value_by_label(rows_by_sheet, [r'mati[cč]na [sš]tevilka']))
    c.tax_number = re.sub(r'\D', '', _find_value_by_label(rows_by_sheet, [r'dav[cč]na [sš]tevilka']))

    # Če label-finder pri AJPES LP2025 vrne header namesto vrednosti, popravi z znanim vzorcem.
    for rows in rows_by_sheet.values():
        for i, row in enumerate(rows[:-1]):
            texts = [_cell_to_text(x) for x in row]
            low = [x.lower() for x in texts]
            if any('ime poslovnega subjekta' in x for x in low) and i + 1 < len(rows):
                candidate = _cell_to_text(rows[i+1][0]) if rows[i+1] else ""
                if candidate and len(candidate) > 2:
                    c.name = candidate
            for j, x in enumerate(low):
                if 'matična številka' in x or 'maticna stevilka' in x:
                    # desno ali spodaj
                    right = ""
                    if j + 1 < len(row):
                        right = re.sub(r'\D', '', _cell_to_text(row[j+1]))
                    down = ""
                    if i + 1 < len(rows) and j < len(rows[i+1]):
                        down = re.sub(r'\D', '', _cell_to_text(rows[i+1][j]))
                    c.registration_number = right or down or c.registration_number
                if 'davčna številka' in x or 'davcna stevilka' in x:
                    right = ""
                    if j + 1 < len(row):
                        right = re.sub(r'\D', '', _cell_to_text(row[j+1]))
                    down = ""
                    if i + 1 < len(rows) and j < len(rows[i+1]):
                        down = re.sub(r'\D', '', _cell_to_text(rows[i+1][j]))
                    c.tax_number = right or down or c.tax_number

    all_text = "\n".join(
        " ".join(_cell_to_text(v) for v in row if _cell_to_text(v))
        for rows in rows_by_sheet.values() for row in rows
    )
    m = re.search(r'na dan\s+([\d]{1,2}\.[\d]{1,2}\.[\d]{4})', all_text, re.I)
    if m: c.period_to = m.group(1)
    m = re.search(r'od\s+([\d]{1,2}\.[\d]{1,2}\.[\d]{4})\s*(?:-|do)\s*([\d]{1,2}\.[\d]{1,2}\.[\d]{4})', all_text, re.I)
    if m:
        c.period_from = m.group(1); c.period_to = m.group(2)
    # SP Excel pogosto vsebuje "v obdobju od 1.1. do 31.12.2025"; period_from se varno izpelje spodaj.
    return c, all_text

def parse_excel_rows(rows_by_sheet, tip_override=None, tip_subjekta_override=None, source_format="XLSX"):
    result = ParseResult(pdf_format=source_format)
    c, all_cell_text = _extract_company_from_excel_rows(rows_by_sheet)
    c = finalize_company(c, tip_override, tip_subjekta_override)
    result.company = c
    result.subject_type = _subject_type_from_tip_subjekta(c.tip_subjekta)

    seen = set()
    for rows in rows_by_sheet.values():
        for row in rows:
            for idx, val in enumerate(row):
                aop = _aop_from_cell(val)
                if not aop or aop in seen:
                    continue
                # prva numerična vrednost desno od AOP = tekoče leto, naslednja = prejšnje leto
                nums = []
                for j in range(idx + 1, min(len(row), idx + 5)):
                    num = _cell_to_number(row[j])
                    if num is not None:
                        nums.append(num)
                if not nums:
                    continue
                result.aop_data[aop] = AopEntry(
                    aop=aop,
                    current_year=nums[0],
                    previous_year=nums[1] if len(nums) > 1 else None
                )
                seen.add(aop)
    return validate(result)

def parse_xlsx_file(xlsx_path, tip_override=None, tip_subjekta_override=None):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    rows_by_sheet = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows_by_sheet[sheet_name] = [tuple(row) for row in ws.iter_rows(values_only=True)]
    return parse_excel_rows(rows_by_sheet, tip_override, tip_subjekta_override, "XLSX")

def parse_xls_file(xls_path, tip_override=None, tip_subjekta_override=None):
    if xlrd is None:
        raise RuntimeError("Za uvoz .xls datotek manjka knjižnica xlrd. Dodaj xlrd==2.0.1 v requirements.txt in redeployaj aplikacijo.")
    book = xlrd.open_workbook(str(xls_path))
    rows_by_sheet = {}
    for sheet in book.sheets():
        rows = []
        for r in range(sheet.nrows):
            rows.append(tuple(sheet.cell_value(r, c) for c in range(sheet.ncols)))
        rows_by_sheet[sheet.name] = rows
    return parse_excel_rows(rows_by_sheet, tip_override, tip_subjekta_override, "XLS")


def parse_pdf_file(pdf_path, tip_override=None, tip_subjekta_override=None):
    pages_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            pages_text.append(page.extract_text() or "")
    # Preveri ali je PDF CID-encoded (pokvarjen font) - ni ga mogoče brati
    if is_cid_encoded(pages_text):
        result = ParseResult()
        result.errors.append(
            "Ta PDF ni prebcrljiv — vsebuje pokvarjen font (CID encoding).\n\n"
            "To se zgodi pri PDFjih iz računovodskih programov (Pantheon, Vasco, miniMax...) "
            "ki ne embedjirajo fontov pravilno.\n\n"
            "Prosimo stranko da pošlje podatke v enem od teh formatov:\n"
            "• Excel (.xlsx) izvoz iz računovodskega programa\n"
            "• AJPES PDF (lp2025_...) — prenesen direktno iz AJPES portala\n"
            "• AJPES Excel — izvoz iz AJPES portala"
        )
        return result
    fmt = detect_pdf_format(pages_text)
    if fmt == "JOLP":
        result = parse_jolp_format(pages_text, tip_override, tip_subjekta_override)
    elif fmt == "NAPOVED":
        result = parse_napoved_format(pages_text, tip_override, tip_subjekta_override)
    else:
        result = parse_aop_format(pages_text, tip_override, tip_subjekta_override)
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

def _to_excel_date(s):
    """Pretvori '31.12.2025' v Python date objekt za Excel."""
    from datetime import date
    if not s or not isinstance(s, str): return None
    try:
        parts = s.strip().split('.')
        if len(parts) == 3:
            return date(int(parts[2]), int(parts[1]), int(parts[0]))
    except: pass
    return s  # fallback na string če pretvorba ne uspe

def export_excel(result, output_path):
    # Tip subjekta 2 = s.p. → poseben template; ostali tipi uporabljajo standardni template
    tip_subjekta = str(getattr(result.company, 'tip_subjekta', '1') or '1')
    tmpl = TEMPLATE_SP_PATH if tip_subjekta == "2" and TEMPLATE_SP_PATH.exists() else TEMPLATE_PATH
    if tip_subjekta == "2" and not TEMPLATE_SP_PATH.exists():
        result.warnings.append("Izbran je tip subjekta 2 (s.p.), ampak template_sp.xlsx ni najden. Uporabljen je standardni template.")
    shutil.copy2(tmpl, output_path)
    wb = openpyxl.load_workbook(output_path)
    c = result.company
    ws = wb["BS"]
    ws.cell(3,4).value  = c.name
    ws.cell(4,4).value  = c.registration_number
    ws.cell(6,4).value  = c.obdobje_bilance
    ws.cell(5,7).value  = c.tip_bilance
    ws.cell(6,7).value  = c.tip_subjekta
    ws.cell(11,4).value = _to_excel_date(c.period_to)
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
    ws2.cell(11,4).value = _to_excel_date(c.period_to)
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
<title>CBK Pretvornik</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:ital,opsz,wght@0,14..32,300;0,14..32,400;0,14..32,500;0,14..32,600;0,14..32,700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
:root {
  --bg:        #f4f8ff;
  --surface:   #ffffff;
  --surface2:  #eef5ff;
  --border:    #dbe7f6;
  --border2:   #b7c9e2;
  --text:      #001b3d;
  --text2:     #4b607c;
  --text3:     #7d8fa8;
  --indigo:    #2f76f6;
  --indigo-lt: #e8f1ff;
  --indigo-md: #b8d2ff;
  --green:     #0f9f82;
  --green-lt:  #e9fbf7;
  --green-md:  #a8eadf;
  --amber:     #c47a00;
  --amber-lt:  #fff7e6;
  --red:       #d92d20;
  --red-lt:    #fff0f0;
  --radius-sm: 8px;
  --radius:    12px;
  --radius-lg: 16px;
  --shadow-sm: 0 1px 2px rgba(0,0,0,.05);
  --shadow:    0 1px 3px rgba(0,0,0,.07), 0 4px 16px rgba(0,0,0,.04);
  --shadow-lg: 0 4px 6px rgba(0,0,0,.04), 0 12px 32px rgba(0,0,0,.07);
}

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

body {
  font-family: 'Inter', -apple-system, sans-serif;
  background: var(--bg);
  color: var(--text);
  min-height: 100vh;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  padding: 24px;
  -webkit-font-smoothing: antialiased;
}

/* Subtle noise texture on bg */
body::before {
  content: '';
  position: fixed;
  inset: 0;
  background-image: url("data:image/svg+xml,%3Csvg viewBox='0 0 256 256' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.9' numOctaves='4' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)' opacity='0.018'/%3E%3C/svg%3E");
  pointer-events: none;
  z-index: 0;
}

.shell {
  width: 100%;
  max-width: 460px;
  position: relative;
  z-index: 1;
}

/* ── WORDMARK ── */
.wordmark {
  display: flex;
  align-items: center;
  gap: 10px;
  margin-bottom: 20px;
  padding: 0 4px;
}

.wm-logo {
  flex-shrink: 0;
  line-height: 0;
  filter: drop-shadow(0 3px 10px rgba(47,118,246,0.22));
}

.wm-logo svg { width: 36px; height: 36px; display: block; }

.wm-name {
  display: flex;
  flex-direction: column;
  gap: 1px;
}

.wm-text {
  font-size: 15px;
  font-weight: 700;
  color: var(--text);
  letter-spacing: -0.03em;
  line-height: 1;
}

.wm-sub {
  font-size: 10.5px;
  color: var(--text3);
  font-weight: 400;
  letter-spacing: 0.02em;
}

.wm-badge {
  margin-left: auto;
  font-size: 10.5px;
  font-weight: 500;
  color: var(--text3);
  background: var(--surface2);
  border: 1px solid var(--border);
  padding: 2px 8px;
  border-radius: 20px;
  letter-spacing: 0.02em;
  flex-shrink: 0;
}

/* ── CARD ── */
.card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius-lg);
  padding: 28px;
  box-shadow: var(--shadow-lg);
}

/* ── STEP HEADER ── */
.step {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-top: 20px;
  margin-bottom: 9px;
}

.step:first-of-type { margin-top: 0; }

.step-num {
  width: 18px; height: 18px;
  border-radius: 50%;
  background: var(--text);
  color: white;
  font-size: 10px;
  font-weight: 700;
  display: flex; align-items: center; justify-content: center;
  flex-shrink: 0;
  letter-spacing: 0;
}

.step-label {
  font-size: 11.5px;
  font-weight: 600;
  color: var(--text2);
  text-transform: uppercase;
  letter-spacing: 0.06em;
}

/* ── SEGMENTED CONTROL (tip + obdobje) ── */
.seg {
  display: flex;
  background: var(--surface2);
  border: 1px solid var(--border);
  border-radius: var(--radius-sm);
  padding: 3px;
  gap: 2px;
  margin-bottom: 4px;
}

.seg-btn {
  flex: 1;
  padding: 8px 6px;
  border-radius: 6px;
  border: none;
  background: transparent;
  cursor: pointer;
  text-align: center;
  transition: all .15s ease;
  font-family: 'Inter', sans-serif;
}

.seg-btn:hover { background: var(--border); }

.seg-btn.selected {
  background: var(--surface);
  box-shadow: var(--shadow-sm), 0 0 0 1px var(--border2);
}

.seg-code {
  display: block;
  font-family: 'JetBrains Mono', monospace;
  font-size: 18px;
  font-weight: 500;
  color: var(--text2);
  line-height: 1.1;
  margin-bottom: 3px;
  transition: color .15s;
}

.seg-btn.selected .seg-code { color: var(--text); }

.seg-name {
  font-size: 10px;
  font-weight: 500;
  color: var(--text3);
  letter-spacing: 0.02em;
  transition: color .15s;
}

.seg-btn.selected .seg-name { color: var(--text2); }

/* Obdobje segmented */
.seg-2 .seg-btn { padding: 10px 12px; text-align: left; }

.seg-main {
  display: block;
  font-family: 'JetBrains Mono', monospace;
  font-size: 12.5px;
  font-weight: 500;
  color: var(--text2);
  margin-bottom: 2px;
  transition: color .15s;
}

.seg-sub {
  display: block;
  font-size: 10.5px;
  color: var(--text3);
  transition: color .15s;
}

.seg-2 .seg-btn.selected .seg-main { color: var(--text); }
.seg-2 .seg-btn.selected .seg-sub  { color: var(--text2); }

/* ── DROP ZONE ── */
.drop-zone {
  border: 1.5px dashed var(--border2);
  border-radius: var(--radius);
  padding: 24px 20px;
  text-align: center;
  cursor: pointer;
  transition: all .2s ease;
  background: var(--surface2);
}

.drop-zone:hover, .drop-zone.drag {
  border-color: var(--indigo);
  background: var(--indigo-lt);
}

.drop-zone.has-file {
  border-color: var(--green);
  border-style: solid;
  background: var(--green-lt);
}

.drop-icon {
  width: 38px; height: 38px;
  border-radius: 9px;
  background: var(--border);
  display: flex; align-items: center; justify-content: center;
  margin: 0 auto 10px;
  transition: all .2s;
}

.drop-zone:hover .drop-icon, .drop-zone.drag .drop-icon {
  background: var(--indigo);
}

.drop-zone.has-file .drop-icon {
  background: var(--green);
}

.drop-icon svg {
  width: 17px; height: 17px;
  stroke: var(--text3); fill: none;
  stroke-width: 2; stroke-linecap: round; stroke-linejoin: round;
  transition: stroke .2s;
}

.drop-zone:hover .drop-icon svg,
.drop-zone.drag .drop-icon svg,
.drop-zone.has-file .drop-icon svg { stroke: white; }

.drop-title {
  font-size: 13.5px;
  font-weight: 600;
  color: var(--text);
  margin-bottom: 3px;
}

.drop-desc { font-size: 12px; color: var(--text3); }

.drop-hint {
  font-size: 11.5px;
  color: var(--text3);
  margin-top: 8px;
  margin-bottom: 12px;
}

.drop-hint a {
  color: var(--indigo);
  cursor: pointer;
  text-decoration: none;
  font-weight: 500;
}

.drop-hint a:hover { text-decoration: underline; }

.tags {
  display: flex;
  gap: 4px;
  justify-content: center;
  flex-wrap: wrap;
  margin-top: 9px;
}

.tag {
  padding: 2px 7px;
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 20px;
  font-size: 10.5px;
  color: var(--text3);
  font-family: 'JetBrains Mono', monospace;
}

input[type=file] { display: none; }

/* ── FILE LIST ── */
.file-list { margin-bottom: 14px; }

.file-item {
  display: flex;
  align-items: center;
  gap: 9px;
  padding: 8px 11px;
  background: var(--surface2);
  border: 1px solid var(--border);
  border-radius: var(--radius-sm);
  margin-bottom: 4px;
}

.file-item-icon {
  width: 26px; height: 26px;
  border-radius: 6px;
  background: var(--green-lt);
  border: 1px solid var(--green-md);
  display: flex; align-items: center; justify-content: center;
  flex-shrink: 0;
}

.file-item-icon svg {
  width: 12px; height: 12px;
  stroke: var(--green); fill: none;
  stroke-width: 2.2; stroke-linecap: round; stroke-linejoin: round;
}

.file-item-name {
  flex: 1;
  font-size: 12.5px;
  font-weight: 500;
  color: var(--text);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.file-item-size {
  font-size: 11px;
  color: var(--text3);
  font-family: 'JetBrains Mono', monospace;
  flex-shrink: 0;
}

.file-item-remove {
  width: 20px; height: 20px;
  border-radius: 50%;
  border: 1px solid var(--border);
  background: transparent;
  cursor: pointer;
  display: flex; align-items: center; justify-content: center;
  flex-shrink: 0;
  transition: all .15s;
}

.file-item-remove:hover { background: var(--red-lt); border-color: #f8b4b4; }
.file-item-remove:hover svg { stroke: var(--red); }

.file-item-remove svg {
  width: 10px; height: 10px;
  stroke: var(--text3); fill: none;
  stroke-width: 2.5; stroke-linecap: round;
}

/* ── DIVIDER ── */
.rule {
  height: 1px;
  background: var(--border);
  margin: 22px 0;
}

/* ── BTN ── */
.btn {
  width: 100%;
  padding: 12px;
  background: var(--text);
  color: white;
  border: none;
  border-radius: var(--radius-sm);
  font-size: 13.5px;
  font-weight: 600;
  font-family: 'Inter', sans-serif;
  cursor: pointer;
  transition: all .15s ease;
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 7px;
  letter-spacing: -0.01em;
}

.btn:hover { background: #002b5c; transform: translateY(-1px); box-shadow: 0 4px 14px rgba(47,118,246,0.22); }
.btn:active { transform: none; box-shadow: none; }
.btn:disabled { background: var(--border2); color: var(--text3); cursor: not-allowed; transform: none; box-shadow: none; }

.btn svg {
  width: 15px; height: 15px;
  stroke: currentColor; fill: none;
  stroke-width: 2; stroke-linecap: round; stroke-linejoin: round;
}

.spinner {
  width: 15px; height: 15px;
  border: 2px solid rgba(255,255,255,.25);
  border-top-color: white;
  border-radius: 50%;
  animation: spin .65s linear infinite;
  display: none;
}

@keyframes spin { to { transform: rotate(360deg); } }

/* ── LOGIN ── */
.login-eyebrow {
  font-size: 11px;
  font-weight: 600;
  color: var(--text3);
  text-transform: uppercase;
  letter-spacing: 0.08em;
  margin-bottom: 6px;
}

.login-title {
  font-size: 22px;
  font-weight: 700;
  color: var(--text);
  letter-spacing: -0.03em;
  margin-bottom: 4px;
}

.login-sub {
  font-size: 13px;
  color: var(--text3);
  margin-bottom: 22px;
  line-height: 1.5;
}

.field-label {
  font-size: 12px;
  font-weight: 600;
  color: var(--text);
  display: block;
  margin-bottom: 6px;
  letter-spacing: -0.01em;
}

input[type=password] {
  width: 100%;
  padding: 10px 13px;
  background: var(--surface);
  border: 1px solid var(--border2);
  border-radius: var(--radius-sm);
  font-size: 14px;
  font-family: 'Inter', sans-serif;
  color: var(--text);
  outline: none;
  transition: all .15s;
  margin-bottom: 14px;
  box-shadow: var(--shadow-sm);
}

input[type=password]:focus {
  border-color: var(--text);
  box-shadow: 0 0 0 3px rgba(47,118,246,.14);
}

select {
  width: 100%;
  padding: 10px 12px;
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius-sm);
  font-size: 12.5px;
  font-family: 'Inter', sans-serif;
  color: var(--text);
  outline: none;
  box-shadow: var(--shadow-sm);
}

select:focus {
  border-color: var(--text);
  box-shadow: 0 0 0 3px rgba(47,118,246,.14);
}

/* ── MESSAGES ── */
.msg {
  padding: 10px 13px;
  border-radius: var(--radius-sm);
  font-size: 12.5px;
  margin-top: 11px;
  border: 1px solid;
  display: none;
  line-height: 1.5;
}

.msg.error { background: var(--red-lt); border-color: #f8b4b4; color: var(--red); }

/* ── RESULT ── */
.result-box {
  display: none;
  margin-top: 14px;
  border: 1px solid var(--green-md);
  border-radius: var(--radius);
  overflow: hidden;
  box-shadow: var(--shadow-sm);
}

.result-header {
  padding: 11px 15px;
  background: var(--green-lt);
  border-bottom: 1px solid var(--green-md);
  display: flex;
  align-items: center;
  gap: 7px;
}

.result-header svg {
  width: 14px; height: 14px;
  stroke: var(--green); fill: none;
  stroke-width: 2.5; stroke-linecap: round; stroke-linejoin: round;
}

.result-header-text { font-size: 12.5px; font-weight: 600; color: var(--green); }

.result-body { padding: 13px 15px; background: var(--surface); }

.result-row {
  display: flex;
  justify-content: space-between;
  align-items: baseline;
  padding: 5px 0;
  border-bottom: 1px solid var(--border);
  font-size: 12.5px;
  gap: 12px;
}

.result-row:last-child { border-bottom: none; }
.result-label { color: var(--text3); flex-shrink: 0; }

.result-val {
  font-weight: 500;
  font-family: 'JetBrains Mono', monospace;
  font-size: 11.5px;
  color: var(--text);
  text-align: right;
  word-break: break-all;
}

.warn-note {
  margin-top: 11px;
  font-size: 12px;
  color: var(--amber);
  background: var(--amber-lt);
  border: 1px solid #ffd58a;
  border-radius: var(--radius-sm);
  padding: 8px 12px;
  line-height: 1.5;
}

/* ── FOOTER ── */
.footer {
  margin-top: 16px;
  text-align: center;
  font-size: 11px;
  color: var(--text3);
  letter-spacing: 0.02em;
}

@media (max-width: 500px) {
  body { padding: 16px; }
  .card { padding: 22px 18px; }
}
</style>
</head>
<body>
<div class="shell">

  <div class="wordmark">
    <div class="wm-logo">
      <!-- CBK logo mark: stylized chart bars with upward arrow -->
      <svg viewBox="0 0 36 36" fill="none" xmlns="http://www.w3.org/2000/svg">
        <!-- Background -->
        <rect width="36" height="36" rx="9" fill="#001b3d"/>
        <!-- Bar 1 -->
        <rect x="7" y="20" width="5" height="9" rx="1.5" fill="white" opacity="0.45"/>
        <!-- Bar 2 -->
        <rect x="15.5" y="14" width="5" height="15" rx="1.5" fill="white" opacity="0.75"/>
        <!-- Bar 3 (tallest) -->
        <rect x="24" y="8" width="5" height="21" rx="1.5" fill="white"/>
        <!-- Upward tick on bar 3 -->
        <path d="M25 11 L26.5 8.5 L28 11" stroke="white" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" fill="none"/>
      </svg>
    </div>
    <div class="wm-name">
      <span class="wm-text">CBK Pretvornik</span>
      <span class="wm-sub">Analiza bilanc</span>
    </div>
    <span class="wm-badge">Interno</span>
  </div>

  <div class="card">

    <!-- LOGIN -->
    <div id="login-section">
      <div class="login-eyebrow">CBK Pretvornik &middot; Interni dostop</div>
      <div class="login-title">Prijava</div>
      <div class="login-sub">Vpišite geslo za dostop do orodja za pretvorbo bilanc.</div>
      <label class="field-label">Geslo</label>
      <input type="password" id="password" placeholder="Vpiši geslo" onkeydown="if(event.key==='Enter')login()">
      <button class="btn" onclick="login()">
        <svg viewBox="0 0 24 24"><path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"/>
        <polyline points="10 17 15 12 10 7"/><line x1="15" y1="12" x2="3" y2="12"/></svg>
        Prijava
      </button>
      <div class="msg error" id="login-error">Napačno geslo — poskusite znova.</div>
    </div>

    <!-- APP -->
    <div id="app-section" style="display:none;">

      <div class="step">
        <div class="step-num">1</div>
        <div class="step-label">Tip bilance</div>
      </div>
      <select id="tip-bilance-select" onchange="selectTipValue(this.value)">
        <option value="3">3 — Revidirana</option>
        <option value="4">4 — Zaključena nerevidirana</option>
        <option value="5">5 — Konsolidirana revidirana</option>
        <option value="6">6 — Konsolidirana nerevidirana</option>
        <option value="7">7 — Preliminarna</option>
        <option value="8" selected>8 — Zaključena</option>
      </select>

      <div class="step">
        <div class="step-num">2</div>
        <div class="step-label">Tip subjekta</div>
      </div>
      <select id="tip-subjekta-select" onchange="selectTipSubjekta(this.value)">
        <option value="1" selected>1 — Gospodarske družbe</option>
        <option value="2">2 — Samostojni podjetniki</option>
        <option value="3">3 — Zadruge</option>
        <option value="4">4 — Društva</option>
        <option value="5">5 — Pravne osebe javnega prava</option>
        <option value="6">6 — Pravne osebe zasebnega prava</option>
        <option value="7">7 — Banke</option>
        <option value="8">8 — Zavarovalnice</option>
        <option value="9">9 — Druge osebe javnega prava</option>
      </select>

      <div class="step">
        <div class="step-num">3</div>
        <div class="step-label">Poslovno obdobje</div>
      </div>
      <div class="seg seg-2">
        <button class="seg-btn selected" onclick="selectObdobje(this,'01.01-31.12')">
          <span class="seg-main">01.01. &#8211; 31.12.</span>
          <span class="seg-sub">Standardno leto</span>
        </button>
        <button class="seg-btn" onclick="selectObdobje(this,'01.04-31.03')">
          <span class="seg-main">01.04. &#8211; 31.03.</span>
          <span class="seg-sub">Nestandardno leto</span>
        </button>
      </div>

      <div class="step">
        <div class="step-num">4</div>
        <div class="step-label">Datoteka(-e)</div>
      </div>
      <div class="drop-zone" id="drop-zone"
           onclick="document.getElementById('file-input').click()"
           ondragover="ev(event,'drag')" ondragleave="ev(event,'')" ondrop="drop(event)">
        <div class="drop-icon">
          <svg viewBox="0 0 24 24"><polyline points="16 16 12 12 8 16"/>
          <line x1="12" y1="12" x2="12" y2="21"/>
          <path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3"/></svg>
        </div>
        <div class="drop-title">Povleci datoteko sem</div>
        <div class="drop-desc">ali klikni za izbiro &mdash; PDF ali Excel</div>
        <div class="tags">
          <span class="tag">lp2025</span>
          <span class="tag">JOLP</span>
          <span class="tag">NAPOVED</span>
          <span class="tag">xlsx</span>
          <span class="tag">xls</span>
        </div>
      </div>
      <input type="file" id="file-input" accept=".pdf,.xlsx,.xls" multiple onchange="filesSelected(this.files)">
      <div class="drop-hint">Za ločena BS + IPI &mdash; <a onclick="document.getElementById('file-input').click()">dodaj oba hkrati</a></div>

      <div class="file-list" id="file-list"></div>

      <button class="btn" id="convert-btn" onclick="convert()" disabled>
        <div class="spinner" id="spinner"></div>
        <svg id="btn-icon" viewBox="0 0 24 24"><polyline points="23 4 23 10 17 10"/>
        <polyline points="1 20 1 14 7 14"/>
        <path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
        <span id="btn-text">Pretvori v Excel</span>
      </button>

      <div class="msg error" id="error-msg"></div>

      <div class="result-box" id="result-box">
        <div class="result-header">
          <svg viewBox="0 0 24 24"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
          <polyline points="22 4 12 14.01 9 11.01"/></svg>
          <span class="result-header-text">Pretvorba uspešna &mdash; prenos se začenja</span>
        </div>
        <div class="result-body">
          <div class="result-row"><span class="result-label">Podjetje</span><span class="result-val" id="r-name"></span></div>
          <div class="result-row"><span class="result-label">Matična</span><span class="result-val" id="r-reg"></span></div>
          <div class="result-row"><span class="result-label">Obdobje</span><span class="result-val" id="r-period"></span></div>
          <div class="result-row"><span class="result-label">Tip bilance</span><span class="result-val" id="r-tip"></span></div>
          <div class="result-row"><span class="result-label">Tip subjekta</span><span class="result-val" id="r-subj"></span></div>
          <div class="result-row"><span class="result-label">Format</span><span class="result-val" id="r-fmt"></span></div>
          <div class="result-row"><span class="result-label">BS vrstice</span><span class="result-val" id="r-bs"></span></div>
          <div class="result-row"><span class="result-label">IPI vrstice</span><span class="result-val" id="r-ipi"></span></div>
          <div class="warn-note" id="warn-partner">&#9888; Ročno vnesi <strong>Št. partnerja</strong> (celica D5 v BS sheetu)</div>
        <div class="warn-note" id="warn-drustvo" style="display:none;background:#fef3c7;border-color:#fcd34d;color:#92400e;">&#9888; Tip subjekta se zapiše v izhodni Excel; za s.p. se uporabi poseben template.</div>
        </div>
      </div>

    </div>
  </div>

  <div class="footer">CBK Pretvornik &middot; Samo za pooblaščene uporabnike</div>
</div>

<script>
let selectedFiles = [];
let selectedTip = '8';
let selectedTipSubjekta = '1';
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

function selectTip(el,val){ selectedTip=val; }
function selectTipValue(val){ selectedTip=val; }
function selectTipSubjekta(val){ selectedTipSubjekta=val; }

function selectObdobje(el,val){
  selectedObdobje=val;
  document.querySelectorAll('.seg-2 .seg-btn').forEach(b=>b.classList.remove('selected'));
  el.classList.add('selected');
}

function ev(e,cls){
  e.preventDefault();
  const z=document.getElementById('drop-zone');
  z.className='drop-zone'+(cls?' '+cls:(selectedFiles.length?' has-file':''));
}

function drop(e){
  e.preventDefault();
  filesSelected(e.dataTransfer.files);
}

function filesSelected(fileList){
  for(let f of fileList){
    const ext=f.name.toLowerCase();
    if(!ext.endsWith('.pdf')&&!ext.endsWith('.xlsx')&&!ext.endsWith('.xls')){
      showError('Dovoljeni so samo PDF ali Excel (.xlsx/.xls) fajli.'); return;
    }
    if(!selectedFiles.find(x=>x.name===f.name)) selectedFiles.push(f);
  }
  renderFileList();
  document.getElementById('error-msg').style.display='none';
  document.getElementById('result-box').style.display='none';
  document.getElementById('drop-zone').className='drop-zone'+(selectedFiles.length?' has-file':'');
  document.getElementById('convert-btn').disabled=selectedFiles.length===0;
}

function removeFile(idx){
  selectedFiles.splice(idx,1);
  renderFileList();
  document.getElementById('drop-zone').className='drop-zone'+(selectedFiles.length?' has-file':'');
  document.getElementById('convert-btn').disabled=selectedFiles.length===0;
}

function renderFileList(){
  const list=document.getElementById('file-list');
  if(!selectedFiles.length){list.innerHTML='';return;}
  list.innerHTML=selectedFiles.map((f,i)=>`
    <div class="file-item">
      <div class="file-item-icon">
        <svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
        <polyline points="14 2 14 8 20 8"/></svg>
      </div>
      <span class="file-item-name" title="${f.name}">${f.name}</span>
      <span class="file-item-size">${(f.size/1024/1024).toFixed(1)}&thinsp;MB</span>
      <button class="file-item-remove" onclick="removeFile(${i})">
        <svg viewBox="0 0 24 24"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
      </button>
    </div>`).join('');
}

function showError(msg){
  const e=document.getElementById('error-msg');
  e.textContent=msg; e.style.display='block';
}

const TIP_LABELS={'3':'3 — Revidirana (ZR)','4':'4 — Zaključena nerevidirana (ZR)','5':'5 — Konsolidirana revidirana (ZR)','6':'6 — Konsolidirana nerevidirana (ZR)','7':'7 — Preliminarna (MR)','8':'8 — Zaključena (ZR)'};
const SUBJEKT_LABELS={'1':'1 — Gospodarske družbe','2':'2 — Samostojni podjetniki','3':'3 — Zadruge','4':'4 — Društva','5':'5 — Pravne osebe javnega prava','6':'6 — Pravne osebe zasebnega prava','7':'7 — Banke','8':'8 — Zavarovalnice','9':'9 — Druge osebe javnega prava'};

async function convert(){
  if(!selectedFiles.length) return;
  const btn=document.getElementById('convert-btn');
  const spinner=document.getElementById('spinner');
  const icon=document.getElementById('btn-icon');
  const btnText=document.getElementById('btn-text');
  btn.disabled=true; spinner.style.display='block'; icon.style.display='none'; btnText.textContent='Pretvarjam...';
  document.getElementById('error-msg').style.display='none';
  document.getElementById('result-box').style.display='none';
  const fd=new FormData();
  for(let f of selectedFiles) fd.append('files',f);
  fd.append('tip_bilance',selectedTip);
  fd.append('tip_subjekta',selectedTipSubjekta);
  fd.append('obdobje',selectedObdobje);
  try{
    const res=await fetch('/convert',{method:'POST',body:fd});
    if(!res.ok){
      const err=await res.json().catch(()=>({error:'Neznana napaka.'}));
      showError(err.error||'Napaka pri pretvorbi.'); return;
    }
    const downloadName=res.headers.get('X-Download-Name')||selectedFiles[0].name.replace(/\.(pdf|xlsx|xls)$/i,'')+'_CBK.xlsx';
    const name=decodeURIComponent(res.headers.get('X-Company-Name')||'');
    const reg=res.headers.get('X-Registration')||'';
    const period=res.headers.get('X-Period')||'';
    const tip=res.headers.get('X-Tip-Bilance')||'';
    const fmt=res.headers.get('X-PDF-Format')||'';
    const bs=res.headers.get('X-BS-Filled')||'';
    const ipi=res.headers.get('X-IPI-Filled')||'';
    document.getElementById('r-name').textContent=name;
    document.getElementById('r-reg').textContent=reg;
    document.getElementById('r-period').textContent=period;
    document.getElementById('r-tip').textContent=TIP_LABELS[tip]||tip;
    const tipSubj=res.headers.get('X-Tip-Subjekta')||'1';
    const subj=res.headers.get('X-Subject-Type')||'GD';
    document.getElementById('r-subj').textContent=SUBJEKT_LABELS[tipSubj]||tipSubj;
    document.getElementById('r-fmt').textContent=fmt;
    document.getElementById('r-bs').textContent=bs;
    // Pokaži opozorilo za društvo če template ni dostopen
    document.getElementById('warn-drustvo').style.display='none';
    document.getElementById('r-ipi').textContent=ipi;
    document.getElementById('result-box').style.display='block';
    const blob=await res.blob();
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.href=url; a.download=downloadName;
    document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
  }catch(e){
    showError('Napaka: '+e.message);
  }finally{
    btn.disabled=false; spinner.style.display='none'; icon.style.display='block'; btnText.textContent='Pretvori v Excel';
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
    return jsonify({"status":"ok","template_exists":TEMPLATE_PATH.exists(),"template_sp_exists":TEMPLATE_SP_PATH.exists(),
                    "formats":["AOP","JOLP","NAPOVED","XLSX","XLS"]})

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
    tip_subjekta_override = request.form.get("tip_subjekta", "1")
    conv_id = str(uuid.uuid4())
    saved_paths = []

    try:
        # Shrani vse naložene datoteke
        for f in files:
            if not f.filename: continue
            fname = f.filename.lower()
            if not (fname.endswith('.pdf') or fname.endswith('.xlsx') or fname.endswith('.xls')):
                return jsonify({"error": f"Nepodprt format: {f.filename}"}), 400
            ext = '.xlsx' if fname.endswith('.xlsx') else ('.xls' if fname.endswith('.xls') else '.pdf')
            p = UPLOAD_DIR / f"{conv_id}_{len(saved_paths)}{ext}"
            f.save(p)
            saved_paths.append(p)

        # Parsiraj — združi rezultate če je več fajlov
        merged = ParseResult()
        fmt_used = ""

        for path in saved_paths:
            if path.suffix == '.xlsx':
                r = parse_xlsx_file(path, tip_override, tip_subjekta_override)
            elif path.suffix == '.xls':
                r = parse_xls_file(path, tip_override, tip_subjekta_override)
            else:
                r = parse_pdf_file(path, tip_override, tip_subjekta_override)

            if r.errors:
                return jsonify({"error": "\n".join(r.errors)}), 422

            # Prva datoteka dobi podatke podjetja
            if not merged.company.name:
                merged.company = r.company
                fmt_used = r.pdf_format
                merged.subject_type = getattr(r, "subject_type", "GD")

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

        # Ime datoteke: NAZIV_TIP_LETO_CBK.xlsx
        # Transliteracija slovenskih znakov
        _slo = str.maketrans('ČčŠšŽžĐđ', 'CcSsZzDd')
        _name_ascii = c.name.upper().translate(_slo)
        company_slug = re.sub(r'[^A-Za-z0-9]', '_', _name_ascii)
        company_slug = re.sub(r'_+', '_', company_slug).strip('_')
        tip_label = TIP_BILANCE_LABELS.get(c.tip_bilance, "BILANCA")
        leto = c.period_to.split('.')[-1] if c.period_to else "XXXX"
        download_name = f"{company_slug}_{tip_label}_{leto}_CBK.xlsx"

        response = send_file(out_path, as_attachment=True,
            download_name=download_name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        safe_name = c.name.encode('ascii', errors='replace').decode('ascii')
        response.headers["X-Download-Name"]  = download_name
        response.headers["X-Subject-Type"]   = getattr(merged, "subject_type", "GD")
        response.headers["X-Tip-Subjekta"]   = c.tip_subjekta
        response.headers["X-Company-Name"]   = safe_name
        response.headers["X-Registration"] = c.registration_number
        response.headers["X-Period"]       = f"{c.period_from} - {c.period_to}"
        response.headers["X-Tip-Bilance"]  = c.tip_bilance
        response.headers["X-PDF-Format"]   = fmt_used
        response.headers["X-BS-Filled"]    = str(bs_n)
        response.headers["X-IPI-Filled"]   = str(ipi_n)
        response.headers["Access-Control-Expose-Headers"] = \
            "X-Download-Name,X-Subject-Type,X-Tip-Subjekta,X-Company-Name,X-Registration,X-Period,X-Tip-Bilance,X-PDF-Format,X-BS-Filled,X-IPI-Filled"
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
