# Omni · Bilance za CBK

Pretvori bilanco iz AJPES (PDF, XLSX, XLS) v Excel za uvoz v CBK. Podatke podjetja (naziv, matična, št. partnerja)
stran prebere iz izreza CRMB, maske HIBIS (pravna oseba), besedila AJPES ePRS ali izpisa iz sodnega registra, zato jih ni treba
iskati v bilanci in vpisovati ročno.

- Vnos podjetja teče v brskalniku (OCR: tesseract.js), strežnik prejme samo naziv, matično in št. partnerja.
- Pretvorba bilance teče na strežniku (`POST /cbk/convert`, samo za prijavljene); naložene datoteke se po pretvorbi izbrišejo.
- Koda pretvornika je v `cbk/pretvornik.py`, predloge so `cbk/template*.xlsx`.
  Za skenirane PDF-je potrebuje strežnik Tesseract in Poppler (`nixpacks.toml`).

## Zagon

Lokalno: `pip install -r requirements.txt && ALLOW_NO_AUTH=1 python app.py` (brez gesla samo za razvoj).
Railway: nastavi spremenljivko `APP_PASS` (ali `ACCESS_PASSWORD` kot doslej). Brez gesla je dostop zaprt.

## Preizkus

`bash tests/run.sh` (Node + `playwright`): prilepi AJPES, naloži izmišljeno bilanco, preveri celice D3 do D5 v Excelu.
`python tests/cbk_smoke.py`: samo strežniška pretvorba, celica za celico proti pričakovanemu Excelu.
