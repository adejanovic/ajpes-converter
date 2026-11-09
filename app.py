import hashlib
import hmac
import os
import time

from flask import Flask, Response, jsonify, redirect, request, send_from_directory
from itsdangerous import BadSignature, URLSafeTimedSerializer

app = Flask(__name__, static_folder=None)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # bilance (PDF, XLSX) za CBK pretvornik

# CBK pretvornik (bilance AJPES -> CBK Excel): koda v mapi cbk/, obdelava na strežniku
try:
    from cbk import pretvornik as cbk
    CBK_ERR = ""
except Exception as e:  # manjkajoča knjižnica ne sme ustaviti obrazcev
    cbk, CBK_ERR = None, str(e)
STATIC = os.path.join(os.path.dirname(__file__), "static")

def _build_id():
    """Kratek odtis vseh knjižnic in obrazcev: ob vsaki spremembi datotek se spremeni,
    zato jih brskalnik lahko hrani dolgo in vseeno vedno dobi najnovejše."""
    h = hashlib.sha1()
    for sub in ("vendor", "forms"):
        d = os.path.join(STATIC, sub)
        for n in sorted(os.listdir(d)) if os.path.isdir(d) else []:
            with open(os.path.join(d, n), "rb") as f:
                h.update(n.encode()); h.update(f.read())
    for n in ("parser.js", "fill.js"):  # tudi programska koda: nova objava = nov naslov, brskalnik ne more ostati na stari
        fp = os.path.join(STATIC, n)
        if os.path.exists(fp):
            with open(fp, "rb") as f:
                h.update(n.encode()); h.update(f.read())
    return h.hexdigest()[:10]


BUILD = _build_id()

PASSWORD = os.environ.get("APP_PASS") or os.environ.get("ACCESS_PASSWORD", "")  # ACCESS_PASSWORD: ime iz starega pretvornika
ALLOW_NO_AUTH = os.environ.get("ALLOW_NO_AUTH") == "1"  # samo za lokalni razvoj
SESSION_HOURS = 12
COOKIE = "obrazci_session"
_signer = URLSafeTimedSerializer(hashlib.sha256(("obrazci-crmb|" + PASSWORD).encode()).hexdigest())

LOGIN_HTML = """<!DOCTYPE html>
<html lang="sl"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Prijava · Omni</title>
<link rel="icon" type="image/png" sizes="32x32" href="/wf-32.png"><link rel="icon" href="/wf.ico" sizes="any"><link rel="apple-touch-icon" href="/wf-180.png">
<meta name="robots" content="noindex, nofollow">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:"Segoe UI",Inter,system-ui,-apple-system,Arial,sans-serif;background:#F4F7FB;color:#14243A;min-height:100vh;display:flex;flex-direction:column;-webkit-font-smoothing:antialiased;font-size:13px}
.top{background:#020F26;color:#fff;height:56px;display:flex;align-items:center;gap:10px;padding:0 24px}
.mark{width:30px;height:30px;display:block}
.top b{font-size:19px;font-weight:700;letter-spacing:-.01em}
.top .sep{width:1px;height:20px;background:#25385A;margin:0 6px 0 8px}
.top span.t{font-size:13.5px;font-weight:600;color:#9FB0C8}
main{flex:1;display:flex;align-items:flex-start;justify-content:center;padding:96px 16px 40px}
.card{width:100%;max-width:380px;background:#fff;border:1px solid #DCE3EC;border-radius:12px;padding:24px;box-shadow:0 1px 2px rgba(10,35,69,.05),0 6px 18px rgba(10,35,69,.06)}
.kick{font-size:11px;font-weight:600;letter-spacing:.06em;color:#52627A}
h1{font-size:18px;font-weight:700;margin:2px 0 4px}
.sub{font-size:12.5px;color:#52627A;margin-bottom:18px}
label{display:block;font-size:11.5px;font-weight:600;color:#42536A;margin-bottom:6px}
input{width:100%;height:40px;padding:0 12px;border:1px solid #C7D5E5;border-radius:8px;font:14px inherit;font-family:inherit;color:#14243A;outline:none;margin-bottom:14px;background:#fff}
input:focus{border-color:#1F5FD6;box-shadow:0 0 0 3px rgba(31,95,214,.15)}
button{width:100%;height:40px;background:#1F5FD6;color:#fff;border:0;border-radius:8px;font:600 13.5px inherit;font-family:inherit;cursor:pointer}
button:hover{background:#174AA8}
button:focus-visible{outline:2px solid #1F5FD6;outline-offset:2px}
.err{display:none;margin-top:12px;padding:9px 12px;background:#FDECEC;border:1px solid #F3C3C3;color:#B42318;border-radius:8px;font-size:12.5px}
.foot{display:flex;align-items:center;justify-content:center;gap:6px;font-size:11.5px;color:#66768A;margin-top:14px}
</style></head><body>
<header class="top"><img class="mark" src="/wf-96.png" alt="" aria-hidden="true"><b>Omni</b><span class="sep" aria-hidden="true"></span><span class="t">Priprava dokumentacije za poslovne stranke</span></header>
<main><div>
 <div class="card">
  <div class="kick">INTERNI DOSTOP</div>
  <h1>Prijava</h1>
  <p class="sub">Za dostop do orodja vpiši geslo.</p>
  <label for="pw">Geslo</label>
  <input type="password" id="pw" autofocus autocomplete="current-password">
  <button id="go" type="button">Prijava</button>
  <div class="err" id="err" role="alert">Napačno geslo, poskusi znova.</div>
 </div>
 <div class="foot"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="4" y="11" width="16" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/></svg>Dostop je zaščiten z geslom.</div>
</div></main>
<script>
async function login(){
  const r = await fetch('/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({password:document.getElementById('pw').value})});
  if(r.ok){ location.reload(); } else { const e=document.getElementById('err'); e.style.display='block'; document.getElementById('pw').select(); }
}
document.getElementById('go').onclick = login;
document.getElementById('pw').addEventListener('keydown', e => { if(e.key==='Enter') login(); });
</script></body></html>"""


def authed():
    if ALLOW_NO_AUTH:
        return True
    if not PASSWORD:
        return False  # brez gesla je dostop zaprt (varno privzeto)
    tok = request.cookies.get(COOKIE)
    if not tok:
        return False
    try:
        _signer.loads(tok, max_age=SESSION_HOURS * 3600)
        return True
    except BadSignature:
        return False


@app.after_request
def headers(resp):
    resp.headers.setdefault("Cache-Control", "no-store")
    resp.headers["X-Robots-Tag"] = "noindex, nofollow"
    return resp


@app.route("/health")
def health():
    return "ok"


@app.route("/login", methods=["POST"])
def login():
    pw = (request.get_json(silent=True) or {}).get("password", "")
    if not PASSWORD or not hmac.compare_digest(pw.encode(), PASSWORD.encode()):
        time.sleep(1)  # upočasni ugibanje gesla
        return jsonify(ok=False), 401
    resp = jsonify(ok=True)
    resp.set_cookie(COOKIE, _signer.dumps("ok"), max_age=SESSION_HOURS * 3600,
                    httponly=True, samesite="Lax", secure=request.is_secure)
    return resp


@app.route("/logout")
def logout():
    resp = redirect("/")
    resp.delete_cookie(COOKIE)
    return resp


@app.route("/")
def index():
    if not authed():
        if not PASSWORD and not ALLOW_NO_AUTH:
            return Response("Aplikacija ni nastavljena: v Railway dodajte spremenljivko APP_PASS.", 503)
        return Response(LOGIN_HTML, mimetype="text/html")
    with open(os.path.join(STATIC, "index.html"), encoding="utf-8") as f:
        html = f.read().replace("__BUILD__", BUILD)
    r = Response(html, mimetype="text/html")
    r.headers["Cache-Control"] = "no-cache"
    return r


PUBLIC = {"favicon.ico", "icon-32.png", "icon-192.png", "apple-touch-icon.png", "wf.ico", "wf-32.png", "wf-180.png", "wf-96.png"}


@app.route("/cbk/convert", methods=["POST"])
def cbk_convert():
    if not authed():
        return jsonify({"error": "Prijava je potrebna."}), 401
    if cbk is None:
        return jsonify({"error": "CBK pretvornik na strežniku ni na voljo: " + CBK_ERR}), 500
    return cbk.convert()  # ista funkcija kot v samostojnem pretvorniku


@app.route("/<path:p>")
def files(p):
    if p in PUBLIC:  # ikona zavihka tudi na prijavni strani
        r = send_from_directory(STATIC, p)
        r.headers["Cache-Control"] = "public, max-age=86400"
        return r
    if not authed():
        return Response("Prijava je potrebna", 401)
    r = send_from_directory(STATIC, p)
    if (p.startswith(("vendor/", "forms/")) or p in ("parser.js", "fill.js")) and request.args.get("v") == BUILD:
        # naslov vsebuje odtis datotek: brskalnik jih hrani in ne sprašuje več strežnika
        r.headers["Cache-Control"] = "private, max-age=31536000, immutable"
    else:
        r.headers["Cache-Control"] = "no-cache"
    return r


if __name__ == "__main__":
    app.run(port=int(os.environ.get("PORT", 8765)))
