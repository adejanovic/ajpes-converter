/* Izpolnjevanje PDF obrazcev (pdf-lib). Vse teče lokalno v brskalniku. */
(function (root) {
  const P = root.Parser;
  let fontBytes = null;
  const warnings = [];

  const LIMIT_DVIG = ['250,00 EUR', '500,00 EUR', '1.000,00 EUR', '2.000,00 EUR', '4.000,00 EUR', 'neomejeno', '0,00'];
  const LIMIT_NAKUP = ['200,00 EUR', '400,00 EUR', '600,00 EUR', '800,00 EUR', '1.000,00 EUR', '1.200,00 EUR', '1.400,00 EUR', '1.600,00 EUR', '1.800,00 EUR', '2.000,00 EUR', '2.200,00 EUR', '2.400,00 EUR', '2.600,00 EUR', '2.800,00 EUR', '3.000,00 EUR', '3.500,00 EUR', '4.000,00 EUR', '4.500,00 EUR', '5.000,00 EUR', 'neomejeno', '0,00'];

  function toTitle(s) {
    return (s || '').toLowerCase().replace(/(^|[\s-])(\S)/g, (m, a, b) => a + b.toUpperCase()).trim();
  }

  /* ---- izris polj enak kot CRMB: pisava iz DA polja (Arial 10), WinAnsi besedilo, znaki zunaj WinAnsi (č, ć, đ) iz vdelane pisave ---- */
  const WIN = { 'š': 0x9a, 'Š': 0x8a, 'ž': 0x9e, 'Ž': 0x8e, '€': 0x80, '–': 0x96, '—': 0x97, '„': 0x84, '“': 0x93, '”': 0x94, '’': 0x92, '‘': 0x91 };
  function ansiByte(ch, strict) { // strict: pisava Helv v teh obrazcih uporablja PDFDocEncoding (znaki 0x80-0x9F so drugačni), zato š/ž/Š/Ž pišemo iz vdelane pisave
    const c = ch.codePointAt(0);
    if (c >= 32 && c < 127) return c;
    if (!strict && WIN[ch] !== undefined) return WIN[ch];
    if (c >= 0xa0 && c <= 0xff) return c;
    return -1;
  }
  const esc = (bytes) => '(' + bytes.map((b) => (b === 40 || b === 41 || b === 92 ? '\\' + String.fromCharCode(b) : b < 32 || b > 126 ? '\\' + b.toString(8).padStart(3, '0') : String.fromCharCode(b))).join('') + ')';

  async function open(name, pre) {
    const { PDFDocument, PDFName, PDFString, PDFTextField, PDFDropdown, PDFCheckBox, PDFRadioGroup } = PDFLib;
    const E = (root.Fill && root.Fill.edits) || null; // ročni popravki iz predogleda (prebrani takoj ob klicu)
    const buf = await (await fetch('forms/' + name + '.pdf?v=' + (window.BUILD || ''))).arrayBuffer();
    const doc = await PDFDocument.load(buf, { ignoreEncryption: true });
    doc.registerFontkit(fontkit);
    if (!fontBytes) {
      const r = await fetch('vendor/LiberationSans-Regular.ttf?v=' + (window.BUILD || ''));
      if (!r.ok) throw new Error('Pisava ni naložena (' + r.status + '). Osvežite stran (Ctrl+F5) in poskusite znova.');
      fontBytes = await r.arrayBuffer();
    }
    const font = await doc.embedFont(fontBytes, { subset: true }); // metrično enaka Arialu; za znake zunaj WinAnsi
    if (pre) pre(doc);
    const form = doc.getForm();
    const miss = [], touched = new Set();
    const ctx = doc.context;

    // pisave iz AcroForm /DR (Arial, ArialNarrow, Calibri ... brez vdelave, kot jih uporablja CRMB)
    const drFonts = {};
    try {
      const dr = form.acroForm.dict.lookup(PDFName.of('DR'));
      const fd = dr.lookup(PDFName.of('Font'));
      for (const k of fd.keys()) drFonts[k.decodeText().replace(/^\//, '')] = fd.get(k);
    } catch (e) { /* ni DR */ }

    function width(text, fs) { try { return font.widthOfTextAtSize(text, fs); } catch (e) { return text.length * fs * 0.55; } }

    function runs(text, fname, fs) {
      let out = '', cur = [], mode = 'a';
      const flush = () => {
        if (!cur.length) return;
        if (mode === 'a') out += `/${fname} ${fs} Tf ` + esc(cur) + ' Tj ';
        else out += `/F1 ${fs} Tf ` + font.encodeText(cur.join('')).toString() + ' Tj ';
        cur = [];
      };
      for (const ch of text) {
        const b = ansiByte(ch, fname === 'Helv');
        const m = b >= 0 && drFonts[fname] ? 'a' : 'o'; // obrazec brez te pisave v /DR (npr. OB0012 nima Ariala): vse z vdelano pisavo
        if (m !== mode) { flush(); mode = m; }
        cur.push(m === 'a' ? b : ch);
      }
      flush();
      return out;
    }

    function wrap(text, maxW, fs) {
      const lines = []; let line = '';
      for (const w of text.split(/\s+/)) {
        const tryL = line ? line + ' ' + w : w;
        if (width(tryL, fs) <= maxW || !line) line = tryL; else { lines.push(line); line = w; }
      }
      if (line) lines.push(line);
      return lines;
    }

    // izbira v skupini izbirnih gumbov. Nekateri obrazci (OB0656 Podpisna kategorija) imajo /Opt s podvojenimi imeni
    // (Izbira1 dvakrat) in prekrivajoče se gumbe: pdf-lib.select() takrat izbere napačen (skriti) gumb, zato stanje nastavimo neposredno.
    function setRadio(fld, v) {
      const ws = fld.acroField.getWidgets(), os = fld.getOptions(), val = String(v);
      const idx = /^\d+$/.test(val) && os[+val] !== undefined ? +val : os.indexOf(val);
      if (idx < 0) { fld.select(val); return; }
      const on = PDFName.of(String(idx));
      const hit = ws.findIndex((w) => { const o = w.getOnValue && w.getOnValue(); return o && o.toString() === on.toString(); });
      if (hit < 0) { fld.select(os[idx]); return; }
      fld.acroField.dict.set(PDFName.of('V'), on);
      ws.forEach((w, i) => w.setAppearanceState(i === hit ? on : PDFName.of('Off')));
      // podvojeni gumbi na istem mestu: njihov »Off« videz (bela podlaga) bi prekril izbrani gumb -> skrijemo jih,
      // izbranemu pa narišemo jasno oznako (izvirna je drobcen križec, ki se pri tisku skoraj ne vidi)
      const R = ws[hit].getRectangle(), ov = (b) => !(b.x > R.x + R.width - 1 || b.x + b.width < R.x + 1 || b.y > R.y + R.height - 1 || b.y + b.height < R.y + 1);
      const dup = ws.filter((w, i) => i !== hit && ov(w.getRectangle()));
      if (dup.length) {
        dup.forEach((w) => w.dict.set(PDFName.of('F'), doc.context.obj(2))); // skrito, se ne natisne
        const W = R.width, H = R.height, ctx = doc.context;
        const q = Math.min(W, H), sq = (q - 1.2) * 0.56; // enak sredinski kvadratek kot na ostalih obrazcih
        const body = `1 g 0 0 ${W} ${H} re f 0 G 0.6 w 0.3 0.3 ${(W - 0.6).toFixed(2)} ${(H - 0.6).toFixed(2)} re S 0 g ${((W - sq) / 2).toFixed(2)} ${((H - sq) / 2).toFixed(2)} ${sq.toFixed(2)} ${sq.toFixed(2)} re f`;
        const st = ctx.register(ctx.stream(body, { Type: 'XObject', Subtype: 'Form', BBox: [0, 0, W, H] }));
        const N = ws[hit].dict.lookup(PDFName.of('AP')).lookup(PDFName.of('N')); N.set(on, st);
      }
    }
    // polje z več gumbi, ki ni označeno kot skupina izbirnih gumbov (npr. Halcom »50«: kartica / USB / Halcom One):
    // vklopi gumb z danim stanjem, ostale izklopi
    function setByOn(n, opt) {
      try {
        const fl = form.getField(n), ws = fl.acroField.getWidgets();
        const hit = ws.findIndex((w) => { const o = w.getOnValue && w.getOnValue(); return o && (o.decodeText ? o.decodeText() : String(o).replace(/^\//, '')) === String(opt); });
        if (hit < 0) return false;
        fl.acroField.dict.set(PDFName.of('V'), PDFName.of(String(opt)));
        ws.forEach((w, i) => w.setAppearanceState(i === hit ? PDFName.of(String(opt)) : PDFName.of('Off')));
        return true;
      } catch (e) { return false; }
    }
    // okras polja iz izvirnega videza (npr. črta pod poljem na Halcom obrazcih), da ob vpisu ne izgine
    function deco(wd, w, h) {
      try {
        const ap = wd.dict.lookup(PDFName.of('AP')), N = ap && ap.lookup ? ap.lookup(PDFName.of('N')) : null;
        if (N && N instanceof PDFLib.PDFRawStream) {
          const raw = PDFLib.decodePDFRawStream(N).decode(); let t = '';
          for (let i = 0; i < raw.length; i++) t += String.fromCharCode(raw[i]);
          t = t.replace(/\/Tx\s+BMC[\s\S]*?EMC/g, ' ').trim();
          if (t && !/\bBT\b|\bDo\b|\bTf\b/.test(t) && /\b(re|l|m)\b/.test(t)) return 'q ' + t + ' Q ';
          if (!t || /\bBT\b/.test(t)) return '';
        }
        const mk = wd.dict.lookup(PDFName.of('MK')), bs = wd.dict.lookup(PDFName.of('BS'));
        const bc = mk && mk.lookup ? mk.lookup(PDFName.of('BC')) : null, st = bs && bs.lookup ? bs.lookup(PDFName.of('S')) : null;
        if (bc && bc.size && bc.size() && st && st.toString() === '/U') {
          const c = bc.asArray().map((x) => x.asNumber()), bw = (bs.lookup(PDFName.of('W')) || { asNumber: () => 1 }).asNumber();
          const col = c.length === 1 ? `${c[0]} G` : c.length === 3 ? `${c.join(' ')} RG` : '0 G';
          return `q ${col} ${bw} w 0 ${bw / 2} m ${w.toFixed(2)} ${bw / 2} l S Q `;
        }
      } catch (e) { /* brez okrasa */ }
      return '';
    }
    function drawField(field, text, opt) {
      const acro = field.acroField;
      let da = acro.getDefaultAppearance() || '';
      let m = da.match(/\/(\S+)\s+([\d.]+)\s+Tf/);
      let fname = m ? m[1] : 'Arial', fs = m ? parseFloat(m[2]) : 10;
      if (!drFonts[fname] || fname === 'Helv' || !fs) { // samodejna velikost/Helvetica -> Arial 10 kot sosednja polja
        fname = drFonts['Arial'] ? 'Arial' : fname; if (!fs) fs = 10;
        acro.setDefaultAppearance(`/${fname} ${fs} Tf 0 g`);
      }
      if (opt && (opt.size || opt.fit) && (!opt.font || drFonts[opt.font] || !drFonts[fname])) { // posebna velikost (npr. polja v tabeli podpisnega kartona, kot jih izpolnjuje CRMB)
        if (opt.font && drFonts[opt.font]) fname = opt.font; if (opt.size) fs = opt.size; acro.setDefaultAppearance(`/${fname} ${fs} Tf 0 g`);
        if (opt.fit) { // predolgo besedilo pomanjšamo, da ostane v polju
          const wmax = acro.getWidgets()[0].getRectangle().width - 6;
          while (fs > (opt.min || 9) && width(text, fs) > wmax) fs -= 0.5;
          acro.setDefaultAppearance(`/${fname} ${fs} Tf 0 g`); // predogled (urejanje) z enako velikostjo kot tisk
        }
      }
      const q = (acro.getQuadding && acro.getQuadding()) || 0;
      const multi = field instanceof PDFTextField && field.isMultiline();
      const comb = field instanceof PDFTextField && field.isCombed();
      const maxLen = field instanceof PDFTextField ? field.getMaxLength() : undefined;
      let wi = -1;
      for (const wd of acro.getWidgets()) {
        wi++;
        const r = wd.getRectangle(); const w = r.width, h = r.height;
        let body = '';
        if (comb && maxLen) {
          const cw = w / maxLen; let i = 0;
          for (const ch of text.slice(0, maxLen)) {
            const x = i * cw + (cw - width(ch, fs)) / 2, y = h / 2 - 0.3815 * fs;
            const b = ansiByte(ch);
            body += `1 0 0 1 ${x.toFixed(2)} ${y.toFixed(2)} Tm ` + runs(ch, fname, fs);
            i++;
          }
          body = `0 g BT ${body}ET`;
        } else if (multi) {
          const lines = wrap(text, w - 4, fs);
          const lead = fs * 1.15; let s = '';
          lines.forEach((ln, i) => { s += `1 0 0 1 2 ${(h - 2 - fs * 0.905 - i * lead).toFixed(2)} Tm ` + runs(ln, fname, fs); });
          body = `q BT ${s}0 g ET Q`;
        } else {
          const tw = width(text, fs);
          const cen = q === 1 || (opt && (opt.center || (opt.centerIdx && opt.centerIdx.includes(wi))));
          if (cen && q !== 1) wd.dict.set(PDFName.of('Q'), ctx.obj(1)); // predogled in urejanje: besedilo na sredini
          const x = cen ? (w - tw) / 2 : q === 2 ? w - 2 - tw : 2;
          const y = h / 2 - 0.3815 * fs;
          body = `q BT 1 0 0 1 ${x.toFixed(2)} ${y.toFixed(2)} Tm ` + runs(text, fname, fs) + '0 g ET Q';
        }
        // spustni seznami imajo v izvirniku belo ozadje, ki prekrije pomožno besedilo pod poljem (»Izberite element.«)
        const bg = field instanceof PDFDropdown ? `1 g 0 2.2 ${w.toFixed(2)} ${(h - 2.2).toFixed(2)} re f 0 g ` : '';
        const content = `${deco(wd, w, h)}${bg}/Tx BMC q 0 0 ${w.toFixed(2)} ${h.toFixed(2)} re W n ${body} Q EMC`;
        const fontDict = { F1: font.ref };
        fontDict[fname] = drFonts[fname] || drFonts['Arial'] || font.ref;
        const stream = ctx.stream(content, { Type: 'XObject', Subtype: 'Form', FormType: 1, BBox: [0, 0, w, h], Resources: { Font: fontDict } });
        wd.dict.set(PDFName.of('AP'), ctx.obj({ N: ctx.register(stream) }));
      }
    }

    const api = {
      doc, form, font, miss,
      big: {},
      text(n, v) {
        if (v == null || v === '') return;
        try { form.getTextField(n).setText(String(v)); touched.add(n); } catch (e) { if (!(api.quiet && /no form field/i.test(e.message))) miss.push(n + ' (' + e.message.slice(0, 60) + ')'); }
      },
      check(n, on = true) {
        try { const c = form.getCheckBox(n); on ? c.check() : c.uncheck(); } catch (e) { miss.push(n); }
      },
      radio(n, opt) {
        try { setRadio(form.getRadioGroup(n), opt); }
        catch (e) { if (!setByOn(n, opt)) miss.push(n + '=' + opt); }
      },
      drop(n, v) {
        if (!v) return;
        try { form.getDropdown(n).select(v); touched.add(n); } catch (e) { miss.push(n + '=' + v); }
      },
      async finish() {
        if (E) for (const [n, v] of Object.entries(E)) {
          try {
            const fld = form.getField(n);
            if (fld instanceof PDFTextField) { fld.setText(fld.isMultiline() ? String(v) : String(v).trim()); touched.add(n); } // presledki na začetku/koncu (ročno »centriranje«) bi zamaknili besedilo
            else if (fld instanceof PDFDropdown) { if (v) fld.select(String(v)); touched.add(n); }
            else if (fld instanceof PDFCheckBox) { if (typeof v === 'string' && v !== 'true') { if (!setByOn(n, v)) fld.check(); } else v ? fld.check() : fld.uncheck(); }
            else if (fld instanceof PDFRadioGroup) { if (v === '__off') { fld.acroField.dict.set(PDFName.of('V'), PDFName.of('Off')); fld.acroField.getWidgets().forEach((w) => w.setAppearanceState(PDFName.of('Off'))); } else if (v) setRadio(fld, v); } // __off = ročno odznačeno
          } catch (e) { /* polje ne obstaja v tej različici */ }
        }
        for (const n of touched) {
          const fld = form.getField(n);
          const val = fld instanceof PDFDropdown ? (fld.getSelected()[0] || '') : fld.getText();
          if (val) drawField(fld, val, api.big[n]);
        }
        // polja s posebno velikostjo/poravnavo, ki so ostala prazna: nastavi enako velikost, da urejanje v predogledu ustreza tisku
        for (const n of Object.keys(api.big)) {
          if (touched.has(n)) continue;
          try { const fld = form.getField(n); if (fld instanceof PDFTextField && !fld.getText()) drawField(fld, '', api.big[n]); } catch (e) { /* ni polja */ }
        }
        return await doc.save({ updateFieldAppearances: false });
      },
    };
    return api;
  }

  const sedez = (k) => [k.ulica, [k.posta, k.kraj].filter(Boolean).join(' '), k.drzava].filter(Boolean).join(', ');
  const krajDatum = (m) => [m.kraj_prejema || 'Ljubljana', m.datum].filter(Boolean).join(', ');

  /* ---------- OB0006: poslovna debetna kartica (2 osebi na obrazec) ---------- */
  async function ob0006(m, chunk) {
    const f = await open('OB0006');
    const k = m.komitent;
    f.text('naziv', k.naziv);
    f.text('sedež', sedez(k));
    f.text('davčna št', k.ds);
    const d = P.trrDigits(k.trr);
    if (d.length === 15) {
      f.text('TRR-1-4', d.slice(0, 4)); f.text('TRR-2-4', d.slice(4, 8));
      f.text('TRR-3-4', d.slice(8, 12)); f.text('TRR-4-4', d.slice(12, 15));
    }
    const names = [
      { ime: 'PO 1', nas: 'naslov PO 1', ds: 'davčna št PO 1', em: 'e-naslov PO 1', gsm: 'GSM PO 1', da: 'Da', ne: 'Ne',
        ship: { subjekt: 'po pošti na naslov uporabnika poslovni subjekt', imetnik: 'po pošti na naslov imetnika kartice', drug: 'po pošti na drug naslov ima in priimek  naziv in naslov' },
        drug: 'drug naslov', lim: 'želim dnevni limit', dvig: 'Limit za dvige 1', nakup: 'limit za nakupe PO1' },
      { ime: 'PO 2', nas: 'naslov PO 2', ds: 'davčna št PO 2', em: 'e-naslov PO 2', gsm: 'GSM PO 2', da: 'Da_2', ne: 'Ne_2',
        ship: { subjekt: 'po pošti na naslov uporabnika poslovni subjekt_2', imetnik: 'po pošti na naslov imetnika kartice_2', drug: 'po pošti na drug naslov ima in priimek  naziv in naslov_2' },
        drug: 'drug-naslov PO 2', lim: 'želim dnevni limit_2', dvig: 'Limit za dvige 2', nakup: 'limit za nakupe PO2' },
    ];
    chunk.slice(0, 2).forEach((p, i) => {
      const n = names[i];
      f.text(n.ime, p.ime); f.text(n.nas, p.naslov); f.text(n.ds, p.ds); f.text(n.em, p.email); f.text(n.gsm, p.tel);
      if (p.sms === 'da') f.check(n.da); else if (p.sms === 'ne') f.check(n.ne);
      f.check(n.ship[p.ship || 'subjekt']);
      if (p.ship === 'drug') f.text(n.drug, p.drug_naslov);
      if (p.limit_dvig || p.limit_nakup) {
        f.check(n.lim);
        f.drop(n.dvig, p.limit_dvig); f.drop(n.nakup, p.limit_nakup);
      }
    });
    f.text('Naziv_uporabnika_na_kartici', (m.kartica_naziv || '').slice(0, 23));
    const zz = m.podpisnik || '';
    f.text('kraj in datum', krajDatum(m));
    f.text('zakoniti zastopnik', toTitle(zz));
    f.text('kraj in datum prejema', krajDatum(m));
    f.text('bančni delavec', m.delavec.ime);
    return { bytes: await f.finish(), miss: f.miss };
  }

  /* ---------- OB0021: podpisni karton (ena oseba na obrazec) ---------- */

  /* V izvirniku OB0021 sta »Novo pooblastilo« in »Preklic pooblastila« besedilni polji; pretvorimo ju v pravi potrditveni polji (klik vklopi/izklopi) */
  function textToCheckbox(doc, names) {
    const { PDFName, PDFString } = PDFLib;
    const ctx = doc.context;
    const annots = doc.getPage(0).node.Annots();
    if (!annots) return;
    for (let i = 0; i < annots.size(); i++) {
      const d = ctx.lookup(annots.get(i));
      const t = d && d.get(PDFName.of('T'));
      if (!t || !names.includes(t.decodeText())) continue;
      const r = d.get(PDFName.of('Rect')).asArray().map((x) => x.asNumber());
      const w = r[2] - r[0], h = r[3] - r[1];
      const mkAp = (body) => ctx.register(ctx.stream(body, { Type: 'XObject', Subtype: 'Form', FormType: 1, BBox: [0, 0, w, h] }));
      const yes = mkAp(`q 0 g 1.1 w 1 1 m ${(w - 1).toFixed(2)} ${(h - 1).toFixed(2)} l S 1 ${(h - 1).toFixed(2)} m ${(w - 1).toFixed(2)} 1 l S Q`);
      const off = mkAp('');
      d.set(PDFName.of('FT'), PDFName.of('Btn'));
      d.set(PDFName.of('Ff'), ctx.obj(0));
      d.set(PDFName.of('V'), PDFName.of('Off'));
      d.set(PDFName.of('AS'), PDFName.of('Off'));
      d.set(PDFName.of('AP'), ctx.obj({ N: { Yes: yes, Off: off } }));
      d.delete(PDFName.of('MaxLen')); d.delete(PDFName.of('Q'));
    }
  }
  // dva podpisnika banke v enem polju: velik razmik (presledki), da se oba lahko podpišeta
  function gapJoin(a, b, width) {
    if (!a || !b) return a || b || '';
    const n = Math.max(8, Math.min(24, Math.floor((width - 40 - 5.1 * (a.length + b.length)) / 2.78)));
    return a + ' '.repeat(n) + b;
  }
  async function ob0021(m, p) {
    const f = await open('OB0021', (doc) => textToCheckbox(doc, ['Novo pooblastilo', 'Preklic pooblastila']));
    const k = m.komitent;
    const d = P.trrDigits(k.trr);
    if (d.length === 15) f.text('št TRR', d.replace(/(\d{4})(?=\d)/g, '$1 ').trim());
    f.text('naziv', k.naziv);
    f.text('zakoniti zastopnik', toTitle(m.podpisnik));
    f.big['Podpis zakonitega zastopnika žig uporabnika'] = { font: 'Arial', size: 12.5, fit: true, center: true };
    f.text('Podpis zakonitega zastopnika žig uporabnika', toTitle(m.podpisnik));
    f.text('PO 1', p.ime);
    f.text('Datum rojstva', p.rojstvo);
    // polja v tabeli: enaka pisava in velikost kot v CRMB (Helvetica 16,66 pt)
    const BIG = { font: 'Helv', size: 16.657 };
    ['Datum rojstva', 'fill_12', 'fill_13', 'fill_14'].forEach(n => { f.big[n] = BIG; });
    f.big['PO 1'] = Object.assign({ fit: true }, BIG);
    f.text('fill_12', p.ds);
    f.text('fill_13', p.sifra);                 // Ključ P (šifra osebe)
    f.text('fill_14', m.komitent.sifra);        // Ključ S/F/T (šifra stranke)
    f.check('Novo pooblastilo');
    f.check(p.zastopanje === 'kolektivno' ? 'kolektivno' : 'toggle_1');
    f.check(p.zz ? 'zakoniti zastopnik1' : 'toggle_4');
    f.check(p.veljavnost === 'od-do' ? 'od' : 'do preklica');
    f.text('Kraj in datum', krajDatum(m));
    return { bytes: await f.finish(), miss: f.miss };
  }

  /* ---------- OB0024: identifikacija pravne osebe ---------- */
  async function ob0024(m) {
    const f = await open('OB0024');
    const k = m.komitent;
    f.text('naziv', k.naziv);
    f.text('sedež', sedez(k));
    f.text('država reg', k.drzava);
    f.text('naslov dostava', k.posiljanje || sedez(k));
    f.text('matična št', k.maticna);
    f.text('davčna št', k.ds);
    if (k.ddv === 'da') f.radio('Group1', 'Izbira1'); else if (k.ddv === 'ne') f.radio('Group1', 'Izbira2');
    f.text('davčno rezidentstvo1', k.drzava);
    f.text('davčna št 1', k.ds);
    // kontakt podjetja: če ni vnesen, se uporabi kontakt podpisnika (zakonitega zastopnika)
    const sp = m.podpisnikP || {};
    f.text('Kontaktni podatki  Tel', k.tel || sp.tel);
    f.text('Enaslov', k.email || sp.email);
    f.text('kraj in datum', krajDatum(m));
    f.text('zakoniti zastopnik', toTitle(m.podpisnik));
    return { bytes: await f.finish(), miss: f.miss };
  }

  /* ---------- OB0025: identifikacija fizične osebe (ena oseba na obrazec) ---------- */
  async function ob0025(m, p) {
    const f = await open('OB0025');
    const k = m.komitent;
    f.text('naziv', k.naziv);
    f.text('sedež', sedez(k));
    f.text('davčna št', k.ds);
    if (p.zz) { f.check('ZAKONITI ZASTOPNIK'); f.check('sklepanje depozitov'); f.check('razpolaganje s sredstvi na TRR'); }
    if (p.po && !p.zz) f.check('Poobl. oseba'); // namen pri pooblaščeni osebi se izpolni ročno
    if (p.dl) { f.check('OSTALO'); f.check('dejanski lastnik'); }
    f.text('PO 1', p.ime);
    f.big['PO 1'] = { centerIdx: [1] }; // ime pod podpisom (stran 2) na sredino
    if (p.dejavnost) f.text('dejavnost', p.dejavnost); // samo dejavnost (brez naziva podjetja)
    f.text('naslov PO 1', p.naslov);
    f.text('začasni naslov PO 1', p.zacasni);
    f.text('naslov za pošiljanje PO 1', p.posiljanje);
    f.text('rojstni podatki PO 1', [p.rojstvo, p.kraj_rojstva, p.drzava_rojstva].filter(Boolean).join(', '));
    f.text('državljanstvo PO 1', p.drzavljanstvo);
    f.text('osebni dokument PO 1', p.dokument);
    f.text('davčno rezidentstvo 1', p.davcno_rez);
    f.text('davčna št 1', p.davcno_rez ? p.ds : ''); // davčna sodi k izjavljenemu rezidentstvu (HIBIS), ne ugibamo
    f.text('Kontaktni podatki  Tel', p.tel);
    f.text('Enaslov', p.email);
    f.text('kraj in datum', krajDatum(m));
    // 2. in 3. stran: podpis (ime osebe) in del "Izpolni banka" (besedilna polja; izbire ostanejo banki)
    // Opomba: ime pod podpisom (stran 2) je isto polje kot "PO 1" - pdf-lib ga zapiše na vse pojavitve.
    f.text('naslov osebe', p.naslov);
    f.text('kraj rojstva', p.kraj_rojstva);
    const dm = (p.dokument || '').match(/št\.?\s*([A-Z0-9]+)/i) || (p.dokument || '').match(/,\s*([A-Z]{0,2}\d[A-Z0-9]{4,})\s*,/i);
    f.text('št', p.dok_st || (dm && dm[1]));
    f.text('datum izdaje', p.dok_izdan);
    f.text('izdajatelj', p.dok_izdajatelj);
    f.text('kraj in datum prejema', krajDatum(m));
    f.radio('Group1', 'Izbira1');   // s prebivališčem in krajem rojstva kot je na prvi strani
    f.radio('Dokument', 'Izbira3'); // osebnega dokumenta, navedenega na prvi strani
    f.text('bančni delavec', m.delavec.ime);
    return { bytes: await f.finish(), miss: f.miss };
  }


  /* ---------- tabelarno definirani obrazci ---------- */
  const digits = (m) => P.trrDigits(m.komitent.trr);
  const iban = (m) => { const d = digits(m); return d.length === 15 ? 'SI56 ' + d.replace(/(\d{4})(?=\d)/g, '$1 ') : ''; };
  const trr4 = (m) => { const d = digits(m); return d.length === 15 ? [d.slice(0, 4), d.slice(4, 8), d.slice(8, 12), d.slice(12, 15)] : ['', '', '', '']; };
  const common = (m) => {
    const k = m.komitent;
    return {
      'naziv': k.naziv, 'sedež': sedez(k), 'davčna št': k.ds, 'matična št': k.maticna,
      'kraj in datum': krajDatum(m), 'zakoniti zastopnik': toTitle(m.podpisnik),
      'kraj in datum prejema': krajDatum(m), 'bančni delavec': m.delavec.ime,
    };
  };
  const SPECS = {
    OB0015: { scope: 'company', fields: (m) => ({ ...common(m), 'TRR': iban(m), 'kraj in datum prejema2': krajDatum(m), 'bančni delavec3': m.vodja || m.delavec.ime }) }, // SZPP (zadnji podpis banke) izpolni podpora, ne mi
    OB0016: { scope: 'company', big: { 'banka': { font: 'Arial', size: 10, center: true, fit: true } }, fields: (m) => ({ ...common(m), 'kraj in datum banka': krajDatum(m), 'banka': gapJoin(m.delavec.ime, m.vodja, 263) }) },
    OB0026: { scope: 'company', fields: (m) => { const c = common(m); return { 'naziv': c.naziv, 'sedež': c['sedež'], 'davčna št': c['davčna št'], 'kraj in datum': c['kraj in datum'], 'zakoniti zastopnik': c['zakoniti zastopnik'] }; } },
    OB0027: { scope: 'company', fields: (m) => ({ ...common(m), 'sifra PO': m.komitent.sifra, 'sektor oz. poslovalnica': 'SPMSD' }) }, // »Dokumentacijo v OPPS pregledal« izpolni OPPS, ne mi
    OB0061: { scope: 'company', fields: (m) => {
      const out = { ...common(m), 'št TRR': iban(m), 'kraj in datum prejema': undefined };
      const SLOTS = [
        ['Ime in priimek', 'fill_8', 'fill_9', 'Datum in kraj rojstva', 'Državljanstvo', ' lastniškega deleža', 'fill_1'],
        ['Ime in priimek_2', 'fill_15', 'fill_16', 'Datum in kraj rojstva_2', 'Državljanstvo_2', ' lastniškega deleža_2', 'fill_3_2'],
        ['Ime in priimek_3', 'fill_22', 'fill_23', 'Datum in kraj rojstva_3', 'Državljanstvo_3', ' lastniškega deleža_3', 'fill_5'],
        ['Ime in priimek_4', 'fill_8_2', 'fill_9_2', 'Datum in kraj rojstva_4', 'Državljanstvo_4', ' lastniškega deleža_4', 'fill_5_2'],
      ];
      (m.owners || []).slice(0, 4).forEach((p, i) => {
        const [nm, st, za, rj, dr, de, ds] = SLOTS[i];
        out[nm] = (p.ime || '').toUpperCase(); /* ime z velikimi črkami kot ostali podatki */ out[st] = p.naslov; out[za] = p.zacasni;
        out[rj] = [p.rojstvo, p.kraj_rojstva].filter(Boolean).join(', ');
        out[dr] = p.drzavljanstvo; out[de] = p.delez ? p.delez + ' %' : '';
        out[ds] = p.ds && p.davcno_rez ? [p.davcno_rez, p.ds].join(' - ') : ''; // npr. SLOVENIJA - 71565884; brez rezidentstva ostane prazno
      });
      return out;
    } },
    OB0071: { scope: 'company', checks: ['Pristop k paketu'], big: { 'kraj in datum prejema': { font: 'Arial', size: 10, fit: true, min: 7 } }, fields: (m) => { const t = trr4(m); return { ...common(m), 'TRR-1-4': t[0], 'TRR-2-4': t[1], 'TRR-3-4': t[2], 'TRR-4-4': t[3] }; } }, // izpolni banka: samo vrstica »Pristopnico prejel« (ne »Vnos / prijava paketa«)
    OB0675: { scope: 'company', checks: ['Pristop k paketu'], big: { 'kraj in datum prejema': { font: 'Arial', size: 10, fit: true, min: 7 } }, fields: (m) => { const t = trr4(m); return { ...common(m), 'TRR-1-4': t[0], 'TRR-2-4': t[1], 'TRR-3-4': t[2], 'TRR-4-4': t[3] }; } }, // kot OB0071: banka izpolni samo »Pristopnico prejel«; paket (Digital START/PRO) se izbere v predogledu
    OB0012: { scope: 'company', fields: (m) => { const t = trr4(m), k = m.komitent, sp = m.podpisnikP || {}; return {
      'naziv': k.naziv, 'sedež': sedez(k), 'davčna št': k.ds,
      'Kontaktna oseba v povezavi z vlogo ime in priimek': toTitle(m.podpisnik), 'Enaslov': k.email || sp.email, 'Telefon GSM': k.tel || sp.tel,
      'TRR-1-4': t[0], 'TRR-2-4': t[1], 'TRR-3-4': t[2], 'TRR-4-4': t[3], 'Naziv uporabnika': m.kartica_naziv,
      'kraj in datum': krajDatum(m), 'zakoniti zastopnik': toTitle(m.podpisnik), 'bančni delavec': m.delavec.ime,
      'sifra PO': k.sifra }; } }, // namen, limit, poravnava in zavarovanje se izberejo v predogledu; odobritev izpolni banka
    OB0011: { scope: 'person', fields: (m, p) => { const k = m.komitent, sp = m.podpisnikP || {}; return {
      'naziv': k.naziv, 'sedež': sedez(k), 'davčna št': k.ds, 'kontaktna oseba': toTitle(m.podpisnik), 'Enaslov': k.email || sp.email, 'Telefon GSM': k.tel || sp.tel,
      'kraj in datum': krajDatum(m), 'zakoniti zastopnik': toTitle(m.podpisnik),
      'PO 1': p.ime, 'naslov PO 1': p.naslov, 'davčna št PO 1': p.ds, 'Enaslov_2': p.email, 'GSM': p.tel,
      'sifra PO': k.sifra, 'bančni delavec': m.delavec.ime, 'kraj in datum prejema': krajDatum(m) }; } }, // vrsta zahtevka in pogoji kartice se izberejo v predogledu; odobreno/zavrnjeno izpolni banka
    // Vloga za kredit in nalog za garancijo: PDF-ja sta bila brez polj; polja so dodana na črte in kvadratke (tools/vloga_polja.py)
    OB0051: { scope: 'company', big: vlogaBig('Kreditojemalec', 'str2 vrstica'), fields: (m) => vlogaGlava(m, 'Kreditojemalec', 'str2 vrstica') },
    OB0053: { scope: 'company', big: vlogaBig('Nalogodajalec', 'Nalogodajalec 2'), fields: (m) => vlogaGlava(m, 'Nalogodajalec', 'Nalogodajalec 2') },
    // Halcom: krovna naročilnica (podjetje, do 10 oseb v tabeli) in zahtevek za digitalno potrdilo (na osebo)
    HCKN: { scope: 'company', big: Object.assign({ 'legalRepresentative': { size: 10, fit: true, min: 7, center: true } }, ...['companyName', 'companyAddress', 'companyCity', 'companyId', 'companyTaxNumber', 'email', 'placeAndDate', '6', '7', '8', '9', '10', '16', '17', '18', '19', '20'].map((n) => ({ [n]: { size: 10, fit: true, min: 7 } }))), fields: (m) => { const k = m.komitent, sp = m.podpisnikP || {}, up = (x) => (x || '').toUpperCase();
      const out = { 'companyName': k.naziv, 'companyId': k.maticna, 'companyAddress': k.ulica, 'companyTaxNumber': k.ds,
        'companyCity': [k.posta, k.kraj].filter(Boolean).join(' '), 'email': k.email || sp.email,
        'legalRepresentative': up(m.podpisnik), 'placeAndDate': krajDatum(m) };
      (m.halcom || []).slice(0, 10).forEach((p, i) => { const zap = i < 5 ? i + 1 : i + 6, ime = i < 5 ? i + 6 : i + 11; out[String(zap)] = String(i + 1); out[String(ime)] = up(p.ime); });
      return out; } },
    HCZA: { scope: 'person', radios: [['50', 'Yesss']], big: Object.assign({}, ...['firstName', 'birthDate', 'taxNumber', 'address', 'phone', 'email', 'email2', 'type ID', 'issuer ID', 'number ID', 'validity ID', 'placeAndDate', 'ime in priimek', 'poslovni subjekt', 'placeAndDate1'].map((n) => ({ [n]: { size: 10, fit: true, min: 6.5 } })), { 'legalRepresentative': { size: 10, fit: true, min: 7, center: true } }), fields: (m, p) => { // medij: Halcom One (spremeniš v predogledu)
      const d = p.dokument || '', parts = d.split(/\s*,\s*/);
      const dates = d.match(/\d{1,2}\.\d{1,2}\.\d{4}/g) || [];
      return { 'firstName': (p.ime || '').toUpperCase(), 'birthDate': p.rojstvo, 'taxNumber': p.ds, 'address': (p.naslov || '').replace(/,\s*SLOVENIJA\s*$/i, '') /* polje je ozko: brez »SLOVENIJA« */, 'phone': p.tel, 'email': p.email, 'email2': p.email,
        'type ID': p.dok_vrsta || (/^[^\d]+$/.test(parts[0] || '') ? parts[0] : ''), 'number ID': p.dok_st || ((d.match(/št\.?\s*([A-Z0-9]+)/i) || d.match(/,\s*([A-Z]{0,2}\d[A-Z0-9]{4,})\s*,/i) || [])[1] || ''),
        'validity ID': [p.dok_izdan || dates[0], p.dok_velja || dates[1]].filter(Boolean).join(' - ') /* datum izdaje - datum veljavnosti */, 'issuer ID': p.dok_izdajatelj || (parts.length > 3 ? parts[parts.length - 1] : ''),
        'placeAndDate': krajDatum(m), 'legalRepresentative': toTitle(m.podpisnik),
        'ime in priimek': (m.delavec.ime || '').toUpperCase(), 'poslovni subjekt': 'SPMSD', 'placeAndDate1': m.datum }; } },
    OB0058: { scope: 'person', fields: (m, p) => ({ 'PO 1': p.ime, 'naslov PO 1': p.naslov, 'začasni naslov PO 1': p.zacasni, 'kraj in datum': krajDatum(m), 'kraj in datum prejema': krajDatum(m), 'bančni delavec': m.delavec.ime }) },
    OB0072: { scope: 'digital', fields: (m, p) => { const d = digits(m); const k = m.komitent; return {
      'naziv': k.naziv, 'sedež': sedez(k), 'davčna št': k.ds, 'Kontaktna_oseba': p.ime, 'e-pošta': p.email, 'telefon': p.tel,
      'TRR-2-4': d.length === 15 ? d.slice(4, 8) : '', 'TRR-3-4': d.length === 15 ? d.slice(8, 12) : '', 'TRR-4-4': d.length === 15 ? d.slice(12, 15) : '',
      'Ime in priimek': p.ime, 'davčna št_1': p.ds, 'eposta_za_prejem': p.email, 'Mobilna_st_za_prejem': p.tel,
      'kraj in datum': krajDatum(m), 'zakoniti zastopnik': toTitle(m.podpisnik), 'komitent PO': k.sifra, 'komitent FO': (p.sifra || '').replace(/^[A-Z]/, ''),
      'kraj in datum prejema': krajDatum(m), 'bančni delavec': m.delavec.ime }; } },
    OB0656: { scope: 'digital', fields: (m, p) => { const t = trr4(m); const k = m.komitent; return {
      'naziv': k.naziv, 'sedež': sedez(k), 'davčna št': k.ds,
      'PO 1': p.ime, 'davčna št PO 1': p.ds, 'GSM uporabnika': p.tel, 'enaslov': p.email,
      'št TRR': k.naziv, 'Besedilo6': p.ime, 'Besedilo7': p.naslov, 'Besedilo8': p.ime, 'Besedilo9': p.email, 'Besedilo10': p.tel,
      'Kraj in datum': krajDatum(m), 'zakoniti zastopnik': toTitle(m.podpisnik),
      'komitent': k.sifra, 'matična št': k.maticna, 'TRR-1-4': t[0], 'TRR-2-4': t[1], 'TRR-3-4': t[2], 'TRR-4-4': t[3] }; } },
  };


  /* Možnosti obrazcev: kljukice, ki jih uporabnik vnaprej nastavi (privzeto + zapomni v brskalniku) */
  const B = (id, label, name, def) => ({ id, label, type: 'bool', def: !!def, act: [['c', name]] });
  const YN = (id, label, da, ne, def) => ({ id, label, type: 'sel', def, opts: [['ne', 'Ne', [['c', ne], ['u', da]]], ['da', 'Da', [['c', da], ['u', ne]]], ['', '(ne obkljukaj)', []]] });
  const PEP = []; for (let i = 4; i <= 19; i++) PEP.push(['r', 'Group' + i, 'Izbira2']);
  const OPTS = {
    OB0027: { title: 'OB0027 Opomnik (priloge)', items: [
      B('c1', 'Vloga za odprtje transakcijskega računa', 'Check Box1', 1),
      B('c2', 'Namen in predvidena narava poslovnega razmerja', 'Check Box2', 1),
      B('c3', 'Pogodba o opravljanju plačilnih storitev (2 izvoda)', 'Check Box3', 1),
      B('c4', 'Podpisni karton pooblaščenih podpisnikov', 'Check Box4', 1),
      B('c5', 'Identifikacija stranke: pravne osebe', 'Check Box5', 1),
      B('c6', 'Interna identifikacija stranke: pravne osebe', 'Check Box6', 0),
      B('c7', 'Identifikacija stranke: fizične osebe (zakoniti zastopnik, pooblaščena oseba)', 'Check Box7', 1),
      B('c8', 'Interna identifikacija stranke: fizične osebe', 'Check Box8', 0),
      B('c9', 'Fotokopije uradnih osebnih dokumentov', 'Check Box9', 0),
      B('c10', 'Vprašalnik za politično izpostavljene osebe', 'Check Box10', 1),
      B('c11', 'Izpis iz Poslovnega registra Slovenije (Ajpes)', 'Check Box11', 1),
      B('c13', 'Izpis iz registra pravnih oseb, registriranih v tujini', 'Check Box13', 0),
      B('c12', 'Izpisi za pravne osebe v lastniški verigi', 'Check Box12', 0),
      B('c14', 'Izpis iz registra dejanskih lastnikov (Ajpes)', 'Check Box14', 1),
      B('c15', 'Ustanovitveni akt', 'Check Box15', 0),
      B('c16', 'Pisna obrazložitev poslovanja v Sloveniji (nerezidenti)', 'Check Box16', 0),
      B('c17', 'Odobritev direktorja sektorja za nerezidente in PPDFT', 'Check Box17', 0),
      B('c18', 'Preverjanje negativnih medijskih objav', 'Check Box18', 1),
      B('c19', 'Društva: izpis iz poslovnega registra', 'Check Box19', 0),
      B('c20', 'Društva: ustanovitvena dokumentacija', 'Check Box20', 0),
      B('c21', 'Društva: vprašalnik za neprofitne organizacije', 'Check Box21', 0),
    ] },
    OB0061: { title: 'OB0061 Dejanski lastniki (točka 2.5)', items: [
      YN('a', 'a) več kot 25 % v lasti tuje pravne osebe, ki ne opravlja dejavnosti v državi registracije', 'Da', 'Ne', 'ne'),
      YN('b', 'b) več kot 25 % v lasti fiduciarne ali podobne družbe z neznanimi lastniki', 'Da_2', 'Ne_2', 'ne'),
    ] },
    OB0058: { title: 'OB0058 Vprašalnik PEP (osebe)', items: [
      { id: 'ne', label: 'Vprašanja 1 do 11 odgovori z Ne', type: 'bool', def: true, act: PEP },
    ] },
    OB0072: { title: 'OB0072 Digitalni žeton (osebe)', items: [
      { id: 'v', label: 'Vrsta zahteve', type: 'sel', def: 'a', opts: [['a', 'Prijava pooblaščenca in naročilo aktivacijskih kod', [['m', 'Moznosti', 'Izbira1']]], ['b', 'Ponovno naročilo aktivacijskih kod', [['m', 'Moznosti', 'Izbira2']]], ['c', 'Sprememba podatkov pooblaščenca', [['m', 'Moznosti', 'Izbira3']]]] },
    ] },
    OB0656: { title: 'OB0656 LINK BIZ (osebe)', items: [
      { id: 'v', label: 'Vrsta zahteve', type: 'sel', def: 'a', opts: [['a', 'Prijava pooblaščenca', [['r', 'Group1', 'Izbira1']]], ['b', 'Preklic pooblaščenca', [['r', 'Group1', 'Izbira2']]], ['c', 'Sprememba podatkov pooblaščenca', [['r', 'Group1', 'Izbira3']]]] },
      { id: 'o', label: 'Oddaljeno podpisovanje', type: 'sel', def: 'ne', opts: [['ne', 'Ne', [['r', 'Oddaljeno_podpisovanje', 'Izbira4']]], ['da', 'Da', [['r', 'Oddaljeno_podpisovanje', 'Izbira3']]]] },
      B('web', 'Spletna aplikacija (Hal E-bank)', 'Spletna aplikacija', 1),
      B('mob', 'Mobilna aplikacija (Halcom MultiPay)', 'Mobilna aplikacija', 0),
      B('pl', 'Plačila: sklop', 'PLAČILA', 1), B('pl1', 'Plačila: priprava', 'Priprava', 1), B('pl2', 'Plačila: podpisovanje', 'Podpisovanje', 1),
      B('pl3', 'Plačila: pošiljanje', 'Pošiljanje', 1), B('pl4', 'Plačila: pregledi', 'Pregledi', 1),
      B('pl5', 'Plačila: priprava paketov', 'Priprava paketov', 0), B('pl6', 'Plačila: oddaljeno podpisovanje', 'Oddaljeno podpisovanje', 0),
      B('sdd', 'SDD: sklop', 'SDD', 0), B('er', 'E-računi: sklop', 'E_RAČUNI', 0), B('eri', 'E-računi: izdaja', 'toggle_15', 0), B('erp', 'E-računi: prejem', 'toggle_18', 0),
    ] },
  };
  function applyOpts(f, name, m, p) {
    const def = OPTS[name]; if (!def) return;
    const saved = (m.formOpts || {})[name] || {};
    const run = (acts) => acts.forEach((a) => {
      if (a[0] === 'c') f.check(a[1]); else if (a[0] === 'u') f.check(a[1], false);
      else if (a[0] === 'r') f.radio(a[1], a[2]); else if (a[0] === 'm') f.radio(a[1], a[2]);
    });
    def.items.forEach((it) => {
      const v = saved[it.id] === undefined ? it.def : saved[it.id];
      if (it.type === 'bool') { if (v) run(it.act); }
      else { const o = it.opts.find((x) => x[0] === v); if (o) run(o[2]); }
    });
    // Vrsta podpisovanja (samostojno/kolektivno) ni vnaprej označena: izbere se v predogledu
  }

  function vlogaBig(naziv, podpis) {
    const fit = { size: 10, fit: true, min: 6.5 }, o = {};
    [naziv, 'Sedež', 'Davčna številka', 'Kontaktna oseba', 'E-mail', 'Telefon', 'Zastopnik', 'v funkciji', 'Kraj, dne'].forEach((n) => { o[n] = fit; });
    o[podpis] = { size: 10, fit: true, min: 7, center: true };
    return o;
  }
  function vlogaGlava(m, naziv, podpis) {
    const k = m.komitent, sp = m.podpisnikP || {}, zz = toTitle(m.podpisnik);
    return { [naziv]: k.naziv, 'Sedež': sedez(k), 'Davčna številka': k.ds, 'Kontaktna oseba': zz, 'E-mail': k.email || sp.email, 'Telefon': k.tel || sp.tel,
      'Zastopnik': zz, 'v funkciji': (sp.dejavnost || '').toUpperCase(), 'Kraj, dne': krajDatum(m), [podpis]: zz };
  }
  async function spec(name, m, p) {
    const s = SPECS[name];
    const f = await open(name);
    f.quiet = true;
    Object.entries(s.big || {}).forEach(([n, o]) => { f.big[n] = o; });
    const fields = s.fields(m, p);
    for (const [n, v] of Object.entries(fields)) f.text(n, v);
    (s.checks || []).forEach((c) => f.check(c));
    (s.radios || []).forEach(([n, v]) => f.radio(n, v));
    applyOpts(f, name, m, p);
    return { bytes: await f.finish(), miss: f.miss };
  }

  root.Fill = { OPTS, spec, SPECS, ob0006, ob0021, ob0024, ob0025, toTitle, LIMIT_DVIG, LIMIT_NAKUP, sedez };
})(window);
