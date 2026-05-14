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
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB

TEMPLATE_PATH = Path(__file__).parent / "template.xlsx"
UPLOAD_DIR    = Path(tempfile.gettempdir()) / "ajpes_uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

ACCESS_PASSWORD = os.environ.get("ACCESS_PASSWORD", "bilance2025")

# ── HTML ──────────────────────────────────────────────────────────────────────

HTML = """<!DOCTYPE html>
<html lang="sl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AJPES Bilance Converter</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');
  :root {
    --bg: #f5f4f1; --surface: #fff; --border: #e2e0d8;
    --text: #1a1917; --text2: #5c5a54; --text3: #9c9a94;
    --slate: #334155; --emerald: #059669; --rose: #e11d48;
    --amber: #d97706;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'DM Sans', sans-serif; background: var(--bg); color: var(--text);
         min-height: 100vh; display: flex; align-items: center; justify-content: center; padding: 24px; }
  .card { background: var(--surface); border: 1px solid var(--border); border-radius: 16px;
          padding: 40px; width: 100%; max-width: 480px; box-shadow: 0 4px 24px rgba(0,0,0,0.07); }
  .logo { display: flex; align-items: center; gap: 12px; margin-bottom: 28px; }
  .logo-icon { width: 40px; height: 40px; background: var(--slate); border-radius: 10px;
               display: flex; align-items: center; justify-content: center; }
  .logo-icon svg { width: 22px; height: 22px; stroke: white; fill: none; stroke-width: 2;
                   stroke-linecap: round; stroke-linejoin: round; }
  .logo-text { font-size: 16px; font-weight: 600; letter-spacing: -0.02em; }
  .logo-sub { font-size: 11px; color: var(--text3); font-weight: 400; letter-spacing: 0.04em; text-transform: uppercase; }
  h1 { font-size: 22px; font-weight: 600; letter-spacing: -0.02em; margin-bottom: 6px; }
  .sub { font-size: 13.5px; color: var(--text3); margin-bottom: 28px; }
  label { font-size: 12.5px; font-weight: 600; color: var(--text); display: block; margin-bottom: 6px; }
  input[type=password], input[type=text] {
    width: 100%; padding: 10px 13px; border: 1px solid var(--border); border-radius: 8px;
    font-size: 14px; font-family: 'DM Sans', sans-serif; outline: none;
    transition: border-color .15s; margin-bottom: 16px; }
  input:focus { border-color: var(--slate); box-shadow: 0 0 0 3px rgba(51,65,85,.08); }
  .drop-zone {
    border: 2px dashed var(--border); border-radius: 14px; padding: 40px 24px;
    text-align: center; cursor: pointer; transition: all .2s; margin-bottom: 16px;
    background: #fafaf8; }
  .drop-zone:hover, .drop-zone.drag { border-color: var(--slate); background: #f0eff8; }
  .drop-zone.ready { border-color: var(--emerald); border-style: solid; background: #f0fdf4; }
  .drop-icon { width: 44px; height: 44px; background: var(--border); border-radius: 10px;
               display: flex; align-items: center; justify-content: center; margin: 0 auto 14px; transition: .2s; }
  .drop-zone:hover .drop-icon, .drop-zone.drag .drop-icon { background: var(--slate); }
  .drop-zone.ready .drop-icon { background: var(--emerald); }
  .drop-icon svg { width: 22px; height: 22px; stroke: var(--text3); fill: none; stroke-width: 2;
                   stroke-linecap: round; stroke-linejoin: round; }
  .drop-zone:hover .drop-icon svg, .drop-zone.drag .drop-icon svg,
  .drop-zone.ready .drop-icon svg { stroke: white; }
  .drop-title { font-size: 14px; font-weight: 600; margin-bottom: 4px; }
  .drop-desc { font-size: 12.5px; color: var(--text3); }
  .tags { display: flex; gap: 6px; justify-content: center; margin-top: 12px; flex-wrap: wrap; }
  .tag { padding: 3px 10px; background: var(--border); border-radius: 20px; font-size: 11.5px; color: var(--text2); }
  input[type=file] { display: none; }
  .btn {
    width: 100%; padding: 12px; background: var(--slate); color: white; border: none;
    border-radius: 10px; font-size: 14px; font-weight: 600; font-family: 'DM Sans', sans-serif;
    cursor: pointer; transition: background .15s; display: flex; align-items: center;
    justify-content: center; gap: 8px; }
  .btn:hover { background: #1e293b; }
  .btn:disabled { background: var(--text3); cursor: not-allowed; }
  .btn svg { width: 16px; height: 16px; stroke: currentColor; fill: none; stroke-width: 2;
             stroke-linecap: round; stroke-linejoin: round; }
  .msg { padding: 12px 16px; border-radius: 10px; font-size: 13px; margin-top: 14px;
         border: 1px solid; display: none; }
  .msg.error { background: #fff1f2; border-color: #fecdd3; color: var(--rose); }
  .msg.success { background: #f0fdf4; border-color: #a7f3d0; color: var(--emerald); }
  .msg.warn { background: #fffbeb; border-color: #fcd34d; color: var(--amber); }
  .spinner { width: 16px; height: 16px; border: 2px solid rgba(255,255,255,.3);
             border-top-color: white; border-radius: 50%; animation: spin .7s linear infinite; display: none; }
  @keyframes spin { to { transform: rotate(360deg); } }
  .result-box { display: none; margin-top: 16px; padding: 16px; background: var(--bg);
                border: 1px solid var(--border); border-radius: 10px; }
  .result-row { display: flex; justify-content: space-between; padding: 5px 0;
                border-bottom: 1px solid var(--border); font-size: 13px; }
  .result-row:last-child { border-bottom: none; }
  .result-label { color: var(--text2); }
  .result-val { font-weight: 500; font-family: 'DM Mono', monospace; font-size: 12.5px; }
  .warn-note { margin-top: 12px; font-size: 12px; color: var(--amber);
               background: #fffbeb; border: 1px solid #fcd34d; border-radius: 8px;
               padding: 8px 12px; }
  .divider { height: 1px; background: var(--border); margin: 20px 0; }
  .footer { text-align: center; font-size: 11.5px; color: var(--text3); margin-top: 20px; }
</style>
</head>
<body>
<div class="card">
  <div class="logo">
    <div class="logo-icon">
      <svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>
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
      <svg viewBox="0 0 24 24"><path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"/><polyline points="10 17 15 12 10 7"/><line x1="15" y1="12" x2="3" y2="12"/></svg>
      Prijava
    </button>
    <div class="msg error" id="login-error">Napačno geslo.</div>
  </div>

  <div id="app-section" style="display:none;">
    <h1>Nova pretvorba</h1>
    <p class="sub">Naloži AJPES letno poročilo in pridobi izpolnjen Excel template</p>

    <div class="drop-zone" id="drop-zone" onclick="document.getElementById('file-input').click()"
         ondragover="ev(event,'drag')" ondragleave="ev(event,'')" ondrop="drop(event)">
      <div class="drop-icon">
        <svg viewBox="0 0 24 24"><polyline points="16 16 12 12 8 16"/><line x1="12" y1="12" x2="12" y2="21"/><path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3"/></svg>
      </div>
      <div class="drop-title" id="drop-title">Povleci PDF sem ali klikni</div>
      <div class="drop-desc" id="drop-desc">AJPES letno poročilo v PDF formatu</div>
      <div class="tags"><span class="tag">PDF</span><span class="tag">maks. 50 MB</span></div>
    </div>
    <input type="file" id="file-input" accept=".pdf" onchange="fileSelected(this.files[0])">

    <button class="btn" id="convert-btn" onclick="convert()" disabled>
      <div class="spinner" id="spinner"></div>
      <svg id="btn-icon" viewBox="0 0 24 24"><polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
      <span id="btn-text">Pretvori</span>
    </button>

    <div class="msg error" id="error-msg"></div>

    <div class="result-box" id="result-box">
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
let authenticated = false;

function login() {
  const pw = document.getElementById('password').value;
  fetch('/auth', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({password: pw})})
  .then(r => r.json()).then(d => {
    if (d.ok) {
      authenticated = true;
      document.getElementById('login-section').style.display = 'none';
      document.getElementById('app-section').style.display = 'block';
    } else {
      const e = document.getElementById('login-error');
      e.style.display = 'block';
    }
  });
}

function ev(e, cls) {
  e.preventDefault();
  const z = document.getElementById('drop-zone');
  z.className = 'drop-zone' + (cls ? ' ' + cls : (selectedFile ? ' ready' : ''));
}

function drop(e) {
  e.preventDefault();
  const file = e.dataTransfer.files[0];
  if (file) fileSelected(file);
  document.getElementById('drop-zone').className = 'drop-zone' + (file ? ' ready' : '');
}

function fileSelected(file) {
  if (!file) return;
  if (!file.name.toLowerCase().endsWith('.pdf')) {
    showError('Naložiti moraš PDF datoteko.'); return;
  }
  selectedFile = file;
  document.getElementById('drop-zone').className = 'drop-zone ready';
  document.getElementById('drop-title').textContent = '✓ ' + file.name;
  document.getElementById('drop-desc').textContent = (file.size / 1024 / 1024).toFixed(1) + ' MB';
  document.getElementById('convert-btn').disabled = false;
  document.getElementById('error-msg').style.display = 'none';
  document.getElementById('result-box').style.display = 'none';
}

function showError(msg) {
  const e = document.getElementById('error-msg');
  e.textContent = msg; e.style.display = 'block';
}

async function convert() {
  if (!selectedFile) return;
  const btn = document.getElementById('convert-btn');
  const spinner = document.getElementById('spinner');
  const icon = document.getElementById('btn-icon');
  const btnText = document.getElementById('btn-text');

  btn.disabled = true;
  spinner.style.display = 'block';
  icon.style.display = 'none';
  btnText.textContent = 'Pretvarjam...';
  document.getElementById('error-msg').style.display = 'none';
  document.getElementById('result-box').style.display = 'none';

  const fd = new FormData();
  fd.append('file', selectedFile);

  try {
    const res = await fetch('/convert', {method: 'POST', body: fd});
    if (!res.ok) {
      const err = await res.json();
      showError(err.error || 'Napaka pri pretvorbi.');
      return;
    }

    // Preberi metadata iz headerjev
    const name   = decodeURIComponent(res.headers.get('X-Company-Name') || '');
    const reg    = res.headers.get('X-Registration') || '';
    const period = res.headers.get('X-Period') || '';
    const tip    = res.headers.get('X-Tip-Bilance') || '';
    const bs     = res.headers.get('X-BS-Filled') || '';
    const ipi    = res.headers.get('X-IPI-Filled') || '';

    // Prikaži rezultat
    document.getElementById('r-name').textContent = name;
    document.getElementById('r-reg').textContent = reg;
    document.getElementById('r-period').textContent = period;
    document.getElementById('r-tip').textContent = tip === '8' ? '8 — Zaključena (ZR)' : '7 — Preliminarna (MR)';
    document.getElementById('r-bs').textContent = bs;
    document.getElementById('r-ipi').textContent = ipi;
    document.getElementById('result-box').style.display = 'block';

    // Prenesi Excel
    const blob = await res.blob();
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    const stem = selectedFile.name.replace(/\\.pdf$/i, '');
    a.href = url; a.download = stem + '_CBK.xlsx';
    a.click(); URL.revokeObjectURL(url);

  } catch(e) {
    showError('Napaka pri povezavi s strežnikom: ' + e.message);
  } finally {
    btn.disabled = false;
    spinner.style.display = 'none';
    icon.style.display = 'block';
    btnText.textContent = 'Pretvori';
  }
}
</script>
</body>
</html>"""

# ── Parser ────────────────────────────────────────────────────────────────────

SIZE_MAP = {
    "Mikro podjetje": "1", "Majhno podjetje": "2",
    "Srednje podjetje": "3", "Veliko podjetje": "4",
}

@dataclass
class CompanyInfo:
    name: str = ""; registration_number: str = ""; tax_number: str = ""
    period_from: str = ""; period_to: str = ""
    tip_bilance: str = ""; tip_subjekta: str = ""; obdobje_bilance: str = ""

@dataclass
class AopEntry:
    aop: str; current_year: Optional[float] = None; previous_year: Optional[float] = None

@dataclass
class ParseResult:
    company: CompanyInfo = field(default_factory=CompanyInfo)
    aop_data: dict = field(default_factory=dict)
    warnings: list = field(default_factory=list)
    errors: list = field(default_factory=list)

def parse_si_number(raw):
    if not raw or str(raw).strip() in ("-","—",""): return None
    try: return float(str(raw).strip().replace(".","").replace(",","."))
    except: return None

def decode_bold(line):
    result = []; i = 0
    while i < len(line):
        ch = line[i]
        if ch == ' ':
            while i < len(line) and line[i] == ' ': i += 1
            result.append(' ')
        elif i+3 < len(line) and line[i]==line[i+1]==line[i+2]==line[i+3]:
            result.append(ch); i += 4
        else:
            result.append(ch); i += 1
    return ''.join(result).strip()

def is_bold(line):
    s = line.strip()[:32].replace(" ","")
    if len(s) < 8: return False
    chunks = [s[i:i+4] for i in range(0, len(s)-3, 4)]
    return sum(1 for c in chunks if len(c)==4 and len(set(c))==1) >= 2

AOP_RE = re.compile(r'(\d{3})\s+([\d.]+,\d{2})(?:\s+([\d.]+,\d{2}))?')

def extract_company(full_text):
    c = CompanyInfo()
    m = re.search(r'Ime poslovnega subjekta[:\s]+([A-ZŠĐŽČĆ][A-ZŠĐŽČĆ\s\.\,\-]{2,60}?)(?=\n|\s{2,}|$)', full_text)
    if m: c.name = m.group(1).strip()
    m = re.search(r'Mati[cč]na [sš]tevilka\s+(\d{10})', full_text)
    if m: c.registration_number = m.group(1).strip()
    m = re.search(r'Dav[cč]na [sš]tevilka\s+(\d{8})', full_text)
    if m: c.tax_number = m.group(1).strip()
    m = re.search(r'\bod\s+([\d]{1,2}\.[\d]{1,2}\.[\d]{4})', full_text)
    if m: c.period_from = m.group(1).strip()
    m = re.search(r'\bdo\s+([\d]{1,2}\.[\d]{1,2}\.[\d]{4})', full_text)
    if m: c.period_to = m.group(1).strip()
    m = re.search(r'Velikost\s+([\w\s]+?)(?=\n)', full_text)
    if m: c.tip_subjekta = SIZE_MAP.get(m.group(1).strip(), "")
    if c.period_to:
        ye = c.period_to.startswith('31.12')
        c.tip_bilance = "8" if ye else "7"
        c.obdobje_bilance = "ZR" if ye else "MR"
    return c

def parse_pdf(pdf_path):
    result = ParseResult(); seen = set(); parts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            parts.append(page.extract_text() or "")
    full = "\n".join(parts)
    result.company = extract_company(full)
    for page_text in parts[1:]:
        for line in page_text.split('\n'):
            decoded = decode_bold(line) if is_bold(line) else line
            m = AOP_RE.search(decoded)
            if not m: continue
            aop = m.group(1)
            if not (1 <= int(aop) <= 310): continue
            if aop in seen: continue
            result.aop_data[aop] = AopEntry(
                aop=aop,
                current_year=parse_si_number(m.group(2)),
                previous_year=parse_si_number(m.group(3)) if m.group(3) else None,
            )
            seen.add(aop)
    if not result.company.name: result.errors.append("Ime podjetja ni najdeno.")
    if not result.company.registration_number: result.errors.append("Matična ni najdena.")
    for code in ["001","055"]:
        if code not in result.aop_data:
            result.errors.append(f"Manjka AOP {code} — morda skenirani PDF?")
    return result

# ── Exporter ──────────────────────────────────────────────────────────────────

BS_AOP  = set([str(i).zfill(3) for i in range(1, 97)] + ["301"])
IPI_AOP = set([str(i).zfill(3) for i in range(110, 190)])
FMT     = '#,##0.00'

def norm(raw):
    if raw is None: return None
    s = str(raw).strip()
    if not s or s in ("/","—"): return None
    try: return str(int(s)).zfill(3)
    except: return None

def build_idx(ws):
    idx = {}
    for row in ws.iter_rows(min_row=12):
        k = norm(row[0].value)
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
            cell = ws.cell(idx[code], 4); cell.value = val; cell.number_format = FMT; bs_n += 1

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
            cell2 = ws2.cell(idx2[code], 4); cell2.value = val; cell2.number_format = FMT; ipi_n += 1

    wb.save(output_path)
    return bs_n, ipi_n

# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML)

@app.route("/auth", methods=["POST"])
def auth():
    data = request.get_json()
    ok = data.get("password","") == ACCESS_PASSWORD
    return jsonify({"ok": ok})

@app.route("/convert", methods=["POST"])
def convert():
    if "file" not in request.files:
        return jsonify({"error": "Ni datoteke."}), 400

    file = request.files["file"]
    if not file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Naložiti moraš PDF datoteko."}), 400

    if not TEMPLATE_PATH.exists():
        return jsonify({"error": "Template ni najden na strežniku."}), 500

    # Shrani PDF začasno
    conv_id   = str(uuid.uuid4())
    pdf_path  = UPLOAD_DIR / f"{conv_id}.pdf"
    xlsx_path = UPLOAD_DIR / f"{conv_id}.xlsx"

    try:
        file.save(pdf_path)

        # Parsaj
        result = parse_pdf(pdf_path)
        if result.errors:
            return jsonify({"error": "\n".join(result.errors)}), 422

        # Izvozi
        bs_n, ipi_n = export_excel(result, xlsx_path)

        c = result.company
        response = send_file(
            xlsx_path,
            as_attachment=True,
            download_name=f"{file.filename.replace('.pdf','')}_CBK.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Metadata v headerjih
        response.headers["X-Company-Name"]  = c.name.encode('latin-1', errors='replace').decode('latin-1')
        response.headers["X-Registration"]  = c.registration_number
        response.headers["X-Period"]        = f"{c.period_from} – {c.period_to}"
        response.headers["X-Tip-Bilance"]   = c.tip_bilance
        response.headers["X-BS-Filled"]     = str(bs_n)
        response.headers["X-IPI-Filled"]    = str(ipi_n)
        response.headers["Access-Control-Expose-Headers"] = \
            "X-Company-Name,X-Registration,X-Period,X-Tip-Bilance,X-BS-Filled,X-IPI-Filled"

        return response

    except Exception as e:
        return jsonify({"error": f"Napaka: {str(e)}"}), 500
    finally:
        # Počisti začasne datoteke
        for p in [pdf_path, xlsx_path]:
            try: p.unlink()
            except: pass

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
