# AJPES Bilance Converter — Deploy navodila

## Datoteke v tej mapi

```
ajpes_app/
  app.py              ← celotna aplikacija
  requirements.txt    ← Python knjižnice
  Procfile            ← za Railway/Render
  railway.toml        ← Railway nastavitve
  template.xlsx       ← !! DODAJ SAM (glej spodaj) !!
```

## NAJPREJ: Dodaj template

Kopiraj `CBK_-_template_za_uvoz_bilanc_14052026.xlsx` v to mapo
in ga **preimenuj** v `template.xlsx`.

---

## Deploy na Railway (PRIPOROČENO — brezplačno)

### 1. GitHub — naloži kodo

1. Pojdi na [github.com](https://github.com) → prijavi se (ali ustvari račun)
2. Klikni **New repository** → ime: `ajpes-converter` → **Create**
3. Klikni **uploading an existing file**
4. Naloži VSE datoteke iz te mape (app.py, requirements.txt, Procfile, railway.toml, template.xlsx)
5. Klikni **Commit changes**

### 2. Railway — deploy

1. Pojdi na [railway.app](https://railway.app) → **Login with GitHub**
2. Klikni **New Project** → **Deploy from GitHub repo**
3. Izberi `ajpes-converter`
4. Railway samodejno zazna Python in zažene app

### 3. Nastavi geslo (OBVEZNO)

1. V Railway projektu klikni na **Variables**
2. Dodaj: `ACCESS_PASSWORD` = `tvoje_geslo_tukaj`
3. Railway samodejno restarta aplikacijo

### 4. Pridobi URL

1. Klikni **Settings** → **Domains** → **Generate Domain**
2. Dobiš URL oblике: `ajpes-converter-production.up.railway.app`
3. Ta URL pošlji sodelavki

---

## Sprememba gesla

V Railway → Variables → spremeni `ACCESS_PASSWORD` → Save
Aplikacija se samodejno restarta.

## Brezplačna omejitev Railway

Brezplačni plan: 500 ur/mesec — za interno uporabo povsem dovolj.
Če bi rabili več: Hobby plan = $5/mesec.

---

## Kako deluje za uporabnika

1. Odpre URL v brskalniku
2. Vnese geslo
3. Naloži AJPES PDF (povleče ali klikne)
4. Klikne **Pretvori**
5. Excel se samodejno prenese
6. Ročno doda samo **Št. partnerja** (celica D5 v BS sheetu)
