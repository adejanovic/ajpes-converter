/* Razčlenjevanje besedila, prebranega s posnetka zaslona "Plačilne storitve" (CRMB).
   Čista funkcija: besedilo (iz OCR ali prilepljeno) -> podatkovni model. Teče v brskalniku in v Node. */
(function (root) {
  const CODE = '([FP]\\d{5,7})';

  function cleanLine(s) {
    return s
      .replace(/[\s|\[\]]+\d{1,2}\.(?:\d\.)?\s*O[B8].{0,3}\d.*$/, '')
      .replace(/^\s*\d{1,2}\.(?:\d\.)?\s*O[B8].{0,3}\d.*$/, '')
      .replace(/\s+(Vse|Nič)(\s+(Vse|Nič))*\s*$/, '')
      .replace(/[|\[\]{}_()]/g, ' ')
      .replace(/\s[—–-]+\s/g, ' ')
      .replace(/[—–]+\s*$/g, '')
      .replace(/\s[-—–]+$/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function tidy(s) {
    return (s || '').replace(/^[\s.,;:'"`´]+|[\s,;:'"`´]+$/g, '').replace(/\s+/g, ' ').trim();
  }

  // Napake OCR pri naslovih: "15. 1230 DOMŽALE. SLOVENIJA" -> "15, 1230 DOMŽALE, SLOVENIJA"

  /* ---- slovenske poštne številke in kraji (za popravek napak OCR, npr. LIUELIANA -> LJUBLJANA) ---- */
  const POST_RAW = '1000 Ljubljana;1210 Ljubljana - Šentvid;1211 Ljubljana - Šmartno;1215 Medvode;1218 Komenda;1230 Domžale;1231 Ljubljana - Črnuče;1233 Dob;1234 Mengeš;1235 Radomlje;1236 Trzin;1241 Kamnik;1270 Litija;1275 Šmartno pri Litiji;1290 Grosuplje;1291 Škofljica;1292 Ig;1293 Šmarje - Sap;1294 Višnja Gora;1295 Ivančna Gorica;1296 Šentvid pri Stični;1310 Ribnica;1330 Kočevje;1360 Vrhnika;1370 Logatec;1380 Cerknica;1381 Rakek;1410 Zagorje ob Savi;1411 Izlake;1420 Trbovlje;1430 Hrastnik;1433 Radeče;2000 Maribor;2204 Miklavž na Dravskem polju;2230 Lenart v Slovenskih goricah;2250 Ptuj;2270 Ormož;2280 Gornja Radgona;2310 Slovenska Bistrica;2311 Hoče;2312 Orehova vas;2313 Fram;2314 Zgornja Polskava;2321 Makole;2322 Majšperk;2323 Ptujska Gora;2324 Lovrenc na Dravskem polju;2325 Kidričevo;2331 Pragersko;2341 Limbuš;2342 Ruše;2352 Selnica ob Dravi;2360 Radlje ob Dravi;2380 Slovenj Gradec;2390 Ravne na Koroškem;2391 Prevalje;2392 Mežica;3000 Celje;3201 Šmartno v Rožni dolini;3210 Slovenske Konjice;3212 Vojnik;3214 Zreče;3220 Štore;3230 Šentjur;3240 Šmarje pri Jelšah;3250 Rogaška Slatina;3252 Rogatec;3260 Kozje;3270 Laško;3301 Petrovče;3310 Žalec;3311 Šempeter v Savinjski dolini;3320 Velenje;3325 Šoštanj;3330 Mozirje;3331 Nazarje;3333 Ljubno ob Savinji;3342 Gornji Grad;4000 Kranj;4201 Zgornja Besnica;4202 Naklo;4203 Duplje;4204 Golnik;4205 Preddvor;4207 Cerklje na Gorenjskem;4208 Šenčur;4209 Žabnica;4211 Mavčiče;4212 Visoko;4220 Škofja Loka;4223 Poljane nad Škofjo Loko;4224 Gorenja vas;4226 Žiri;4227 Selca;4228 Železniki;4240 Radovljica;4243 Brezje;4244 Podnart;4245 Kropa;4246 Kamna Gorica;4247 Zgornje Gorje;4248 Lesce;4260 Bled;4264 Bohinjska Bistrica;4265 Bohinjsko jezero;4270 Jesenice;4274 Žirovnica;4275 Begunje na Gorenjskem;4276 Hrušica;4280 Kranjska Gora;4281 Mojstrana;4290 Tržič;4294 Križe;5000 Nova Gorica;5210 Deskle;5213 Kanal;5220 Tolmin;5230 Bovec;5250 Solkan;5270 Ajdovščina;5280 Idrija;5282 Cerkno;5290 Šempeter pri Gorici;5291 Miren;5292 Renče;5293 Volčja Draga;5294 Dornberk;5295 Branik;5297 Prvačina;6000 Koper;6210 Sežana;6215 Divača;6230 Postojna;6240 Kozina;6250 Ilirska Bistrica;6257 Pivka;6280 Ankaran;6310 Izola;6320 Portorož;6330 Piran;6333 Sečovlje;8000 Novo mesto;8210 Trebnje;8211 Dobrnič;8216 Mirna Peč;8222 Otočec;8230 Mokronog;8232 Šentrupert;8233 Mirna;8250 Brežice;8251 Čatež ob Savi;8257 Dobova;8259 Bizeljsko;8270 Krško;8273 Leskovec pri Krškem;8280 Brestanica;8281 Senovo;8290 Sevnica;8294 Boštanj;8310 Šentjernej;8311 Kostanjevica na Krki;8330 Metlika;8333 Semič;8340 Črnomelj;8344 Vinica;8350 Dolenjske Toplice;8360 Žužemberk;9000 Murska Sobota;9201 Puconci;9220 Lendava;9224 Turnišče;9226 Moravske Toplice;9231 Beltinci;9232 Črenšovci;9240 Ljutomer;9250 Gornja Radgona;9252 Radenci;9253 Apače';
  const POST = {}; POST_RAW.split(';').forEach((e) => { POST[e.slice(0, 4)] = e.slice(5); });
  const PLACES = Object.values(POST).filter((v, i, a) => a.indexOf(v) === i);
  const asciiUp = (t) => t.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().replace(/[^A-Z]/g, '');
  function levS(a, b) { const m = a.length, n = b.length; let pr = Array.from({ length: n + 1 }, (_, j) => j); for (let i = 1; i <= m; i++) { const cu = [i]; for (let j = 1; j <= n; j++) cu[j] = Math.min(pr[j] + 1, cu[j - 1] + 1, pr[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)); pr = cu; } return pr[n]; }
  function caseLike(src, name) { const letters = src.replace(/[^A-Za-zČŠŽĆĐčšžćđ]/g, ''); return letters && letters === letters.toUpperCase() ? name.toUpperCase() : name; }
  // kraj brez poštne številke: popravi samo, če je zelo podoben znanemu kraju
  function fixCity(city) {
    const c = tidy(city || ''); const k = asciiUp(c); if (k.length < 5) return c;
    let best = null, bd = 99;
    PLACES.forEach((n) => { const d = levS(k, asciiUp(n)); if (d < bd) { bd = d; best = n; } });
    return best && (bd > 0 || c.toUpperCase() !== best.toUpperCase()) && bd <= Math.max(1, Math.floor(k.length * 0.34)) ? caseLike(c, best) : c;
  }
  // naslov, ki se konča s »1234 KRAJ«: kraj preveri glede na poštno številko
  function fixPostCity(addr) {
    return addr.replace(/(\b\d{4})\s+([^\d,]{3,})$/, (all, code, city) => {
      const want = POST[code]; if (!want) { const f = fixCity(city); return f === city ? all : code + ' ' + f; }
      const a = asciiUp(city), b = asciiUp(want), b2 = asciiUp(want.split(' - ')[0]);
      const d = Math.min(levS(a, b), levS(a, b2));
      const tc = tidy(city).toUpperCase(); const same = tc === want.toUpperCase() || tc === want.split(' - ')[0].toUpperCase();
      if (a !== b && a !== b2 && (a.startsWith(b2) || b.startsWith(a))) return all; // npr. »Ljubljana - Šentvid«
      return !same && d <= Math.max(2, Math.floor(a.length * 0.45)) ? code + ' ' + caseLike(city, want) : all;
    });
  }
  function fixAddress(s) {
    s = tidy(s);
    s = s.replace(/(\d+[A-Za-z]?)[.,]?\s+(\d{4}\s)/, '$1, $2');
    s = s.replace(/([A-ZČŠŽĆĐ])[.,]\s+([A-ZČŠŽĆĐ]{4,})$/, '$1, $2');
    return fixPostCity(s);
  }

  function fixDoc(s) {
    return tidy(s).replace(/([a-zčšž])\.\s+(št\.)/i, '$1, $2').replace(/[,.\s]+ve$/i, '');
  }

  function fixCompany(s) {
    s = tidy(s);
    s = s.replace(/([A-ZČŠŽĆĐ0-9])\s*D\.?O\.?O\.?$/, '$1 D.O.O.');
    s = s.replace(/([A-ZČŠŽĆĐ0-9])\s*D\.?D\.?$/, '$1 D.D.');
    return s;
  }

  function fixDate(m) {
    return m ? `${m[1]}.${m[2]}.${m[3]}` : '';
  }

  function emptyPerson() {
    return { sifra: '', ime: '', ds: '', naslov: '', zacasni: '', posiljanje: '', rojstvo: '', kraj_rojstva: '', drzava_rojstva: '', drzavljanstvo: '', dokument: '', osr: '' };
  }

  function parseScreen(text) {
    const out = {
      komitent: { sifra: '', naziv: '', ds: '', maticna: '', trr: '', ulica: '', posta: '', kraj: '', drzava: '' },
      zast: [],            // [{sifra, ime}]
      po: [],              // do 4 oseb
      delavec: { koda: '', ime: '' },
      kraj_prejema: '',
      trrs: [],            // TRR pri Gorenjski banki (07000)
      opozorila: [],
    };
    const lines = text.split(/\r?\n/).map(cleanLine).filter(Boolean);
    let sec = 'top';
    let poIdx = -1;

    for (let L of lines) {
      let m;
      // --- vrhnji del ---
      if (sec === 'top') {
        if ((m = L.match(new RegExp('^Komitent\\s*:?\\s*' + CODE + '\\s+(.+)$', 'i')))) {
          out.komitent.sifra = m[1];
          out.komitent.naziv = fixCompany(m[2].replace(/(\s+(Vse|Nič))+$/i, ''));
          continue;
        }
        if ((m = L.match(/^D[ŠS5]\s*:?\s*(\d{8})/))) out.komitent.ds = m[1];
        if ((m = L.match(/Mati\S*\s*št\W*(\d{10})/i))) out.komitent.maticna = m[1];
        for (const t of L.matchAll(/(\d{5})\s*-\s*(\d{10})/g)) { // samo računi pri nas (07000)
          if (t[1] !== '07000') continue; const v = t[1] + '-' + t[2];
          if (!out.trrs.includes(v)) out.trrs.push(v); if (!out.komitent.trr) out.komitent.trr = v;
        }
        if ((m = L.match(/^Ulica\W*(.*?)(?:\s+Po[šs]tna.*)?$/i))) { out.komitent.ulica = tidy(m[1]); }
        if ((m = L.match(/Po[šs]tna\s*št\W*(\d{4})/i))) out.komitent.posta = m[1];
        if ((m = L.match(/^Kraj\W+(.+?)\s+Dr[žz]ava\W*(.*)$/i))) {
          out.komitent.kraj = tidy(m[1]);
          out.komitent.drzava = tidy(m[2]);
        }
      }
      if (/^Zakoniti\s+zastopnik/i.test(L)) { sec = 'zz'; continue; }

      // --- zakoniti zastopniki ---
      if (sec === 'zz' && (m = L.match(new RegExp('^Oseba\\s*([12])\\W*(?:' + CODE + '\\s+)?(.*)$', 'i')))) {
        const ime = tidy(m[3]);
        if (ime) out.zast.push({ sifra: m[2] || '', ime });
        continue;
      }

      // --- pooblaščene osebe ---
      if (/^Naziv\b/i.test(L)) {
        sec = 'po';
        poIdx++;
        if (poIdx > 3) continue;
        const p = emptyPerson();
        let rest = L.replace(/^Naziv\W*/i, '');
        if ((m = rest.match(new RegExp('^' + CODE + '\\s*')))) { p.sifra = m[1]; rest = rest.slice(m[0].length); }
        if ((m = rest.match(/(?:^|\s)D[ŠS5](?=[\s:.]|$)\s*:?\s*(\d{8})?.*$/))) { p.ds = m[1] || ''; rest = rest.slice(0, m.index); }
        p.ime = tidy(rest);
        out.po[poIdx] = p;
        continue;
      }
      if (sec === 'po' && poIdx >= 0 && poIdx <= 3) {
        const p = out.po[poIdx];
        if ((m = L.match(/^Naslov\s+za\s+po[šs]iljanje\W*(.*)$/i))) { p.posiljanje = fixAddress(m[1]); continue; }
        if ((m = L.match(/^Za[čc]asni\s+naslov\W*(.*)$/i))) { p.zacasni = fixAddress(m[1]); continue; }
        if ((m = L.match(/^Naslov\W*(.*)$/i))) { p.naslov = fixAddress(m[1]); continue; }
        if ((m = L.match(/^Datum\s+rojstva\W*(\d{2})\D{0,3}(\d{2})\D{0,3}(\d{4})(.*)$/i))) {
          p.rojstvo = fixDate(m);
          const k = m[4].match(/Kraj\s+rojstva\W*(.*)$/i);
          if (k) p.kraj_rojstva = fixCity(tidy(k[1]));
          continue;
        }
        if ((m = L.match(/^Datum\s+rojstva\W*(.*)$/i))) {
          const k = m[1].match(/Kraj\s+rojstva\W*(.*)$/i);
          if (k) p.kraj_rojstva = fixCity(tidy(k[1]));
          continue;
        }
        if ((m = L.match(/^OSR\W*(.*?)\s*Os\.?\s*dokument\W*(.*)$/i))) { p.osr = tidy(m[1]); p.dokument = fixDoc(m[2]); continue; }
        if ((m = L.match(/^Os\.?\s*dokument\W*(.*)$/i))) { p.dokument = fixDoc(m[1]); continue; }
      }

      // --- noga ---
      if ((m = L.match(/^Ban\S*\s+delavec\W*(\d{3,7})\s*(.*)$/i))) {
        out.delavec.koda = m[1];
        out.delavec.ime = tidy(m[2]);
        continue;
      }
      if ((m = L.match(/^Kraj\s+prejema\W*(.+)$/i))) { out.kraj_prejema = tidy(m[1]); continue; }
    }

    out.po = out.po.filter((p) => p && (p.sifra || p.ime || p.ds || p.naslov));
    return out;
  }

  // Združi zakonite zastopnike in pooblaščene osebe v en seznam oseb (ista šifra = ista oseba).
  function buildPersons(parsed) {
    const persons = [];
    const mk = (b) => Object.assign(emptyPerson(), b, {
      zz: false, po: false, tel: '', email: '',
      kartica: false, digital: false, sms: '', ship: 'subjekt', drug_naslov: '', limit_dvig: '', limit_nakup: '',
      zastopanje: 'posamično', veljavnost: 'do preklica',
    });
    const norm = (s) => (s || '').toUpperCase().replace(/[^A-ZČŠŽĆĐ]/g, '');
    for (const z of parsed.zast) {
      const p = mk({ sifra: z.sifra, ime: z.ime });
      p.zz = true;
      persons.push(p);
    }
    for (const q of parsed.po) {
      let p = persons.find((x) => (q.sifra && x.sifra === q.sifra) || (!q.sifra && norm(x.ime) === norm(q.ime)));
      if (!p) { p = mk({}); persons.push(p); }
      for (const k of Object.keys(q)) if (q[k]) p[k] = q[k];
      p.po = true;
      p.kartica = true;
      p.digital = true;
    }
    return persons;
  }


  // ---------- HIBIS: zavihka Identifikacija in Kontakti ----------
  const DATE = '(\\d{2})\\D{0,2}(\\d{2})\\D{0,2}(\\d{4})';
  // Pravila veljajo samo za slovenske dokumente (državljanstvo Slovenija): osebna izkaznica 00xxxxxxx ali IExxxxxxx, potni list PBxxxxxxx.
  // Tuji dokumenti imajo drugačne oblike, zato se pri njih številka samo očisti (velike črke, brez presledkov), sicer ostane, kot je prebrana.
  function fixDocNo(s, vrsta, drz) {
    s = (s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (!/^SLOVENIJ/i.test(drz || '') || !/^[A-Z0-9]{9}$/.test(s)) return s;
    const toD = { O: '0', I: '1', S: '5', B: '8', Z: '2', L: '1', D: '0', Q: '0' };
    const tail = s.slice(2).replace(/[A-Z]/g, (c) => toD[c] || c);
    let head = s.slice(0, 2);
    if (/potn/i.test(vrsta || '')) { if (/^[PF][B8R3E]$/.test(head)) head = 'PB'; }
    else if (/^[O0DQ]{2}$/.test(head)) head = '00';
    else if (/^[I1L|][E3F]$/.test(head)) head = 'IE';
    else head = head.replace(/[01528]/g, (c) => ({ '1': 'I', '0': 'O', '5': 'S', '8': 'B', '2': 'Z' })[c]);
    return head + tail;
  }
  function fixPhone(s) {
    const d = (s || '').replace(/[^\d+]/g, '');
    if (!d) return '';
    let n = d.replace(/^\+?386/, '').replace(/^0/, '');
    if (!/^\d{8}$/.test(n)) return d;
    return '+386 ' + n.slice(0, 2) + ' ' + n.slice(2, 5) + ' ' + n.slice(5);
  }
  const DOMAINS = ['gmail.com', 'siol.net', 'hotmail.com', 'yahoo.com', 'outlook.com', 't-2.net', 'telemach.net', 'amis.net', 'protonmail.com', 'icloud.com', 'live.com', 'msn.com', 'gmx.com', 'gmx.net', 'volja.net', 'yahoo.co.uk', 'outlook.si', 'email.si'];
  // OCR pogosto prebere @ kot O, G, E, 0: če @ manjka, ga vstavimo pred znano domeno.
  function fixEmail(s) {
    s = (s || '').toLowerCase().replace(/[,;\s]+$/g, '');
    // znaki, ki jih OCR pogosto vstavi namesto @ (£, €, §, ...), in vse, kar v e-naslovu ni dovoljeno
    s = s.replace(/[£€§©®¢¥#&$]/g, '@').replace(/\s+/g, '').replace(/[^a-z0-9._%+@-]/g, '').replace(/@{2,}/g, '@');
    if ((s.match(/@/g) || []).length > 1) { const i = s.lastIndexOf('@'); s = s.slice(0, i).replace(/@/g, '') + s.slice(i); }
    if (!s || s.includes('@')) return s;
    for (const d of DOMAINS) if (s.endsWith(d) && s.length > d.length + 1) return s.slice(0, -d.length - 1) + '@' + d;
    // @ prebran kot o/a/e/g/0/q: poišči znak, za katerim sledi videti veljavna domena (vsaj 3 znaki pred piko)
    for (const ch of ['o', '0', 'a', 'e', 'g', 'q', 'c']) {
      for (let i = 2; i < s.length - 6; i++) {
        if (s[i] !== ch) continue;
        const dom = s.slice(i + 1);
        if (/^[a-z0-9-]{3,}(\.[a-z0-9-]{2,})*\.[a-z]{2,}$/.test(dom)) return s.slice(0, i) + '@' + dom;
      }
    }
    return s;
  }
  // popravek e-naslova glede na znano domeno podjetja in ime osebe (OCR pogosto izgubi ali zamenja črko)
  function refineEmail(email, ctx) {
    const m = (email || '').match(/^([^@]+)@(.+)$/);
    if (!m) return email || '';
    let local = m[1], dom = m[2];
    const cd = ((ctx && ctx.domain) || '').toLowerCase();
    if (cd && dom !== cd && lev(dom, cd) <= 2) dom = cd;
    const plain = (x) => (x || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z]/g, '');
    for (const t of ((ctx && ctx.names) || []).map(plain)) if (t.length >= 4 && local !== t && lev(local, t) <= 1) { local = t; break; }
    return local + '@' + dom;
  }
  const clean2 = (s) => tidy((s || '').replace(/^[a-z∨⌄v<>.\-]\s+/i, '').replace(/^[\W_]+/, ''));
  const cap = (s) => (s ? s.charAt(0).toUpperCase() + s.slice(1).toLowerCase() : '');

  function parseHibis(text) {
    const h = { ime: '', priimek: '', ds: '', rojstvo: '', kraj_rojstva: '', drzava_rojstva: '', drzavljanstvo: '', davcno_rez: '',
      sifra: '', dok_vrsta: '', dok_st: '', dok_izdajatelj: '', dok_izdan: '', dok_velja: '', tel: '', email: '', naslov: '', zacasni: '', posiljanje: '' };
    let partner = '';
    let asec = '', ul = {}, hs = {};
    for (const L of text.split(/\r?\n/).map(cleanLine).filter(Boolean)) {
      let m;
      if (/^Stalni\s+naslov/i.test(L)) { asec = 'naslov'; continue; }
      if (/^Za[čc]asno\s+prebivali/i.test(L)) { asec = 'zacasni'; continue; }
      if (/^Naslov\s+za\s+po[šs]iljanje/i.test(L)) { asec = 'posiljanje'; continue; }
      if (asec) {
        const a = (ul[asec] = ul[asec] || { ulica: '', hs: '', posta: '', kraj: '' });
        if ((m = L.match(/^Ulica\s+(.*?)(?:\s+Hi[šs]na\s+[šs]tevilka\s*(.*))?$/i))) { a.ulica = /^Hi[šs]na\b/i.test(m[1]) ? '' : tidy(m[1]); if (m[2]) a.hs = tidy(m[2]); continue; }
        if ((m = L.match(/^Po[šs]ta\s+\D{0,3}?(\d{4})\D{0,3}\s+(\S.*)$/i))) { a.posta = m[1]; a.kraj = tidy(m[2].replace(/^[^A-Za-zČŠŽĆĐčšžćđ]+/, '')); continue; }
        if ((m = L.match(/^Po[šs]ta\s+\D{0,3}?(\d{4})\b/i))) { a.posta = m[1]; continue; }
      }
      if ((m = L.match(/^Partner\s+(\d{5,9})\s+(.+)$/i))) { partner = tidy(m[2]); h.sifra = sifraFO(m[1]); } // šifra komitenta FO: prva števka (4) -> P
      else if ((m = L.match(/^Partner\s+\d*\s*(.+)$/i))) partner = tidy(m[1]);
      if ((m = L.match(/^Ime\s+(\S.*)$/))) h.ime = tidy(m[1]);
      if ((m = L.match(/^Priimek\s+(\S.*)$/i))) h.priimek = tidy(m[1]);
      if ((m = L.match(/Dav[čc]na\s+(\d{8})\b/))) h.ds = m[1];
      if ((m = L.match(new RegExp('Datum\\s+rojstva\\s+' + DATE, 'i')))) h.rojstvo = m[1] + '.' + m[2] + '.' + m[3];
      if ((m = L.match(/Kraj\s+rojstva\s+(.+)$/i))) h.kraj_rojstva = fixCity(clean2(m[1]));
      if ((m = L.match(/Dr[žz]ava\s+rojstva\s+\d*\s*(.+)$/i))) h.drzava_rojstva = clean2(m[1]);
      if ((m = L.match(/rezidentstva\s+\d{3}\s*(.+)$/i))) h.davcno_rez = clean2(m[1]);
      if ((m = L.match(/^Dr[žz]avljanstvo\s+\d{3}\s*(.+)$/i))) h.drzavljanstvo = clean2(m[1]);
      if ((m = L.match(/Vrsta\s+osebnega\s+dokumenta\s+\d*\s*(.+?)(?:\s+[ŠS]tevilka.*)?$/i))) h.dok_vrsta = cap(clean2(m[1]));
      if ((m = L.match(/[ŠS]tevilka\s+oseb\.?\s*dokumenta\s+([A-Za-z0-9]{6,12})/i))) h.dok_st = m[1];
      if ((m = L.match(/Izdajatelj\s+oseb\.?\s*dok\.?\s+(.+?)(?:\s{2,}.*)?$/i))) h.dok_izdajatelj = clean2(m[1]);
      if ((m = L.match(new RegExp('Dat\\.?\\s*izd\\.?\\s+osebnega\\s+dokumenta\\s+' + DATE, 'i')))) h.dok_izdan = m[1] + '.' + m[2] + '.' + m[3];
      if ((m = L.match(/Datum\s+velj\.?\s+os\.?\s*dok\.?\s+(\d{2})\D{0,2}(\d{2})[^\d]{0,2}(\d{3,4})/i))) h.dok_velja = m[1] + '.' + m[2] + '.' + (m[3].length === 3 ? '2' + m[3] : m[3]);
      if ((m = L.match(/Mobilni\s+tele\w+\s+([\d\s+]{8,})/i))) h.tel = fixPhone(m[1]);
      if ((m = L.match(/^E-?MAIL\s+(\S{6,})/i))) h.email = fixEmail(m[1]);
    }
    if (h.dok_st) h.dok_st = fixDocNo(h.dok_st, h.dok_vrsta, h.drzavljanstvo);
    if (!h.priimek && partner) { const w = partner.split(' '); h.priimek = w[0]; h.ime = h.ime || w.slice(1).join(' '); }
    for (const k of ['naslov', 'zacasni', 'posiljanje']) {
      const a = ul[k];
      if (a && a.ulica) h[k] = fixPostCity([[a.ulica, a.hs].filter(Boolean).join(' '), [a.posta, a.kraj].filter(Boolean).join(' ')].filter(Boolean).join(', '));
    }
    h.imepriimek = tidy((h.priimek + ' ' + h.ime));
    return h;
  }
  const isHibis = (t) => !/Komitent\s*:/i.test(t) && /(Vrsta\s+partnerja|Mobilni\s+tele|Ime\s+\S+\s*\n.*Priimek|rezidentstva|Priimek|Stalni\s+naslov|Za[čc]asno\s+prebivali)/i.test(t);

  /* ---- HIBIS maska pravne osebe (podjetje): zgornji del, zavihki Identifikacija, Naslovi, Kontakti ---- */
  const PO_MARK = /(PRAVNE\s+OSEBE|Dodatni\s+naziv|Mati[čc]na\s+[šs]tevilka|Kontaktna\s+oseba|Pro[ka]urist|Direktor|SWIFT|Dodatne\s+opombe|GVIN|registra\s+tujcev)/i;
  const FO_MARK = /(Priimek|Datum\s+rojstva|Stalni\s+naslov|osebnega\s+dok|Dr[žz]avljanstvo|Kraj\s+rojstva|FIZI[ČC]NE\s+OSEBE|Vrsta\s+partnerja|EM[ŠS]O|Dr[žz]ava\s+izvora|Telefon\s+v\s+slu[žz]bi|Zaposlitev|Sociodemograf)/i;
  const isHibisPO = (t) => !/Komitent\s*:/i.test(t) && !FO_MARK.test(t) && (PO_MARK.test(t) || (/Naslov\s+za\s+po[šs]iljanje/i.test(t) && /^\s*Naziv\b/im.test(t)));
  function streetCase(s) { // »KOLODVORSKA CESTA 4« -> »Kolodvorska cesta 4«
    const small = /^(cesta|ulica|ul\.?|c\.?|pot|trg|naselje|nabrežje|steza|cesti|na|pri|v|ob|pod|nad|za|in|k|s|z|ga|gasilska)$/i;
    return tidy(s).toLowerCase().split(' ').map((w, i) => (i > 0 && small.test(w)) ? w : w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
  }
  function coName(s) { return fixCompany(s).replace(/\s*D\.O\.O\.$/, ' d.o.o.').replace(/\s*D\.D\.$/, ' d.d.').replace(/\s*S\.P\.$/, ' s.p.'); }
  // šifra komitenta iz številke partnerja v HIBIS (7 števk): prva števka odpade, spredaj črka
  // fizična oseba -> P, samostojni podjetnik (s.p.) -> S, druge pravne osebe -> F; že vpisana črka ostane
  const isSP = (naziv) => /\bs\.\s*p\.?\s*$|samostojn\w*\s+podjetni/i.test(naziv || '');
  function sifraFO(v) { const t = String(v || '').trim().toUpperCase().replace(/\s+/g, ''); const m = t.match(/^\d?(\d{6})$/); return m ? 'P' + m[1] : t; }
  function sifraPO(v, naziv) { const t = String(v || '').trim().toUpperCase().replace(/\s+/g, ''); const m = t.match(/^\d?(\d{6})$/); return m ? (isSP(naziv) ? 'S' : 'F') + m[1] : t; }
  function parseHibisPO(text) {
    const h = { sifra: '', partner: '', naziv: '', popolno: '', ds: '', maticna: '', ulica: '', posta: '', kraj: '', drzava: '', tel: '', mob: '', email: '', kontakt: '', posiljanje: '', skrbnik_koda: '', skrbnik_ime: '' };
    let sec = '', a = {}, ad = { '': {}, zac: {}, pos: {} };
    for (const L of text.split(/\r?\n/).map(cleanLine).filter(Boolean)) {
      let m;
      if (/^(Identifikacija|Naslovi|Kontakti)\b.*(Naslovi|Kontakti|Ostali)/i.test(L)) { sec = 'tab'; continue; }
      if (/^Za[čc]asno\s+prebivali/i.test(L)) { sec = 'zac'; continue; }
      if (/^Naslov\s+za\s+po[šs]iljanje/i.test(L)) { sec = 'pos'; continue; }
      a = ad[sec === 'tab' ? 'x' : sec] || (ad.x = ad.x || {});
      if ((m = L.match(/^Skrbnik\s+(\d{3,6})\W*(?:v\s+)?([A-Za-zČŠŽĆĐčšžćđ].+)$/))) { h.skrbnik_koda = m[1]; h.skrbnik_ime = tidy(m[2]); continue; }
      if ((m = L.match(/^Partner\s+(\d{7})\b/i))) { h.partner = m[1]; continue; } // 2602101 -> F602101 (črka po pravni obliki)
      if (sec === '' && (m = L.match(/^Dr[žz]ava\s+\d{3}\s*([A-ZČŠŽĆĐ][A-ZČŠŽĆĐ ]+?)(?:\s+[A-Za-z]{2})?$/))) { h.drzava = tidy(m[1]); continue; }
      if (!h.maticna && (m = L.match(/Mati[čc]na\s+[šs]tevilka\s+(\d{7,10})\b/i))) h.maticna = m[1];
      if (!h.maticna && (m = L.match(/poro[čc]ila\.?\s+(\d{7,10})\b/i))) h.maticna = m[1];
      if ((m = L.match(/^Dodatni\s+naziv\s+(.+)$/i))) { h.popolno = tidy(m[1]); continue; }
      if ((m = L.match(/^Naziv\s+(.+)$/i))) { if (sec === 'pos') a.naziv = tidy(m[1]); else if (sec === '') h.naziv = coName(m[1]); continue; }
      if ((m = L.match(/^Ulica\s+(.*?)(?:\s+Hi[šs]na\s+[šs]t\S*\s*(\S+)?)?$/i))) { if (!/^Hi[šs]na\b/i.test(m[1])) a.ulica = tidy(m[1]); if (m[2]) a.hs = m[2].replace(/[^\dA-Za-z/]/g, ''); continue; }
      if ((m = L.match(/^Po[šs]ta\s+\D{0,3}?(\d{4})\b\s*(.*)$/i))) { a.posta = m[1]; a.kraj = tidy(m[2].replace(/\s+Dodatni\s+izbor.*$/i, '').replace(/^[^A-Za-zČŠŽĆĐčšžćđ]+/, '')); continue; }
      if ((m = L.match(/^Dav[čc]na\s+(\d{8})\b/i))) { h.ds = m[1]; continue; }
      if ((m = L.match(/^Telefon\s+([\d+][\d\s/+-]{5,})/i))) { h.tel = fixPhone(m[1]); continue; }
      if ((m = L.match(/^Mobilni\s+tele\w+\s+([\d+][\d\s/+-]{5,})/i))) { h.mob = fixPhone(m[1]); continue; }
      if ((m = L.match(/^E-?MAIL\s+(\S{6,})/i))) { h.email = m[1]; continue; }
      if ((m = L.match(/^Kontaktna\s+oseba\s+(\S.+)$/i))) { h.kontakt = tidy(m[1]); continue; }
    }
    if (h.email) { // @ prebran kot e/a/o: če je pred njim ime kontaktne osebe, je @ takoj za imenom
      const raw = h.email.toLowerCase(), plain = (x) => x.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z]/g, '');
      let fixed = '';
      if (!raw.includes('@')) for (const t of h.kontakt.split(/\s+/).map(plain).filter((x) => x.length >= 3).reverse()) {
        const j = raw.indexOf(t); if (j < 0) continue; const e = j + t.length;
        if ('oaegq0c'.includes(raw[e]) && /^[a-z0-9-]{2,}(\.[a-z0-9-]{2,})*\.[a-z]{2,}$/.test(raw.slice(e + 1))) { fixed = raw.slice(0, e) + '@' + raw.slice(e + 1); break; }
      }
      h.email = fixed || fixEmail(h.email);
    }
    if (h.partner) h.sifra = sifraPO(h.partner, h.naziv);
    const city = (x) => (POST[x.posta] ? POST[x.posta].split(' - ')[0] : (x.kraj ? fixCity(x.kraj) : ''));
    const top = ad[''];
    if (top.ulica) h.ulica = streetCase([top.ulica, top.hs].filter(Boolean).join(' '));
    if (top.posta) { h.posta = top.posta; h.kraj = city(top); }
    const pz = ad.pos;
    if (pz.ulica) h.posiljanje = [pz.naziv, streetCase([pz.ulica, pz.hs].filter(Boolean).join(' ')), [pz.posta, pz.posta ? city(pz) : ''].filter(Boolean).join(' ')].filter(Boolean).join(', ');
    return h;
  }
  function mergeHibisPO(list) {
    const out = {};
    for (const k of Object.keys(list[0])) out[k] = vote(list.map((r) => r[k]), k === 'ds' ? checkDS : null);
    return out;
  }

  function mergeHibis(list) {
    const out = {};
    const preds = { ds: checkDS };
    for (const k of Object.keys(list[0])) out[k] = vote(list.map((r) => r[k]), preds[k]);
    return out;
  }

  // --- preverjanja ---
  function checkDS(ds) {
    if (!/^\d{8}$/.test(ds || '')) return false;
    const w = [8, 7, 6, 5, 4, 3, 2];
    let s = 0;
    for (let i = 0; i < 7; i++) s += w[i] * +ds[i];
    let c = 11 - (s % 11);
    if (c === 11) return false;
    if (c === 10) c = 0;
    return c === +ds[7];
  }
  function trrDigits(trr) { return (trr || '').replace(/\D/g, ''); }
  function ibanOf(trr) { const d = trrDigits(trr); return d.length === 15 ? 'SI56' + d : ''; }
  function checkIBAN(iban) {
    if (!/^SI\d{17}$/.test(iban || '')) return false;
    const r = iban.slice(4) + '2818' + iban.slice(2, 4); // S=28, I=18
    let rem = 0;
    for (const ch of r) rem = (rem * 10 + +ch) % 97;
    return rem === 1;
  }


  // Glasovanje med več branji (različne velikosti/pragovi): za vsako polje izberemo najpogostejšo vrednost,
  // veljavna kontrolna številka (DŠ, IBAN) ima prednost.
  function lev(a, b) {
    const m = a.length, n = b.length;
    let prev = Array.from({ length: n + 1 }, (_, j) => j);
    for (let i = 1; i <= m; i++) {
      const cur = [i];
      for (let j = 1; j <= n; j++) cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      prev = cur;
    }
    return prev[n];
  }
  function vote(vals, pred) {
    const all = vals.filter((v) => v && v.length > 1);
    if (all.length < Math.min(2, vals.length)) return '';
    const cnt = new Map();
    all.forEach((v) => cnt.set(v, (cnt.get(v) || 0) + 1));
    let best = '', bs = -Infinity;
    for (const [v, c] of cnt) {
      const dist = all.reduce((s, o) => s + lev(v, o), 0);
      const s = 10 * c - dist + (pred && pred(v) ? 1000 : 0);
      if (s > bs) { best = v; bs = s; }
    }
    return best;
  }
  function mergeParsed(list) {
    const first = list[0];
    const out = JSON.parse(JSON.stringify(first));
    const preds = { ds: checkDS, trr: (v) => checkIBAN(ibanOf(v)) };
    for (const k of Object.keys(out.komitent)) out.komitent[k] = vote(list.map((r) => r.komitent[k]), preds[k]);
    out.delavec.koda = vote(list.map((r) => r.delavec.koda));
    out.delavec.ime = vote(list.map((r) => r.delavec.ime)); // skrbnik iz CRMB (siva vrstica »Bančni delavec«)
    out.kraj_prejema = vote(list.map((r) => r.kraj_prejema));
    out.trrs = [...new Set(list.flatMap((r) => r.trrs || []))].filter((v) => checkIBAN(ibanOf(v))); // veljavni naši TRR iz vseh branj
    if (out.komitent.trr && !checkIBAN(ibanOf(out.komitent.trr)) && out.trrs.length) out.komitent.trr = out.trrs[0];
    const n = (key) => Math.max(...list.map((r) => r[key].length));
    out.zast = [];
    for (let i = 0; i < n('zast'); i++) {
      const g = list.map((r) => r.zast[i]).filter(Boolean);
      out.zast.push({ sifra: vote(g.map((x) => x.sifra)), ime: vote(g.map((x) => x.ime)) });
    }
    out.po = [];
    for (let i = 0; i < n('po'); i++) {
      const g = list.map((r) => r.po[i]).filter(Boolean);
      const q = emptyPerson();
      for (const k of Object.keys(q)) q[k] = vote(g.map((x) => x[k]), preds[k]);
      if (q.sifra || q.ds || q.ime.length > 3 || q.naslov.length > 5) out.po.push(q);
    }
    return out;
  }



  // Sodni register (redni izpis iz sodnega/poslovnega registra, PDF): družbeniki, deleži, zastopniki
  function parseSodni(text) {
    const lines = (text || '').replace(/\r/g, '').replace(/\u00a0/g, ' ').split('\n').map((l) => l.replace(/\s+$/, '').replace(/^\s+/, ''));
    const kvOf = (l) => { const m = l.match(/^(.+?)\s{2,}(\S.*)$/); return m ? [m[1].trim(), m[2].trim()] : null; };
    const top = {};
    let sec = '', rec = null;
    const owners = [], zast = [], shares = [];
    const flush = () => { if (!rec) return; (rec.t === 'o' ? owners : rec.t === 'z' ? zast : shares).push(rec); rec = null; };
    for (const l of lines) {
      if (!l) continue;
      if (/^OSNOVNI PODATKI O SUBJEKTU/.test(l)) { flush(); sec = 'top'; continue; }
      if (/^DRUŽBENIKI IN POSLOVNI DELEŽI/.test(l)) { flush(); sec = 'dr'; continue; }
      if (/^DRUŽBENIKI$/.test(l)) { flush(); sec = 'dr'; continue; }
      if (/^POSLOVNI DELEŽI$/.test(l)) { flush(); sec = 'del'; continue; }
      if (/^OSEBE, POOBLAŠČENE ZA ZASTOPANJE/.test(l)) { flush(); sec = 'zast'; continue; }
      if (/^(ČLANI ORGANA NADZORA|SKUPŠČINSKI SKLEPI|RAZNO|SPREMEMBA DRUŽBENE|PROKURA|ZASTAVA)/.test(l)) { flush(); sec = 'x'; continue; }
      if (/^Stran \d+ od \d+/.test(l)) continue;
      const kv = kvOf(l); if (!kv) continue;
      const [k, v] = kv;
      if (sec === 'top') top[k] = v;
      else if (sec === 'dr' || sec === 'del' || sec === 'zast') {
        if (/^Zap\. ?št\. družbenika/i.test(k)) { flush(); rec = { t: 'o', zap: v }; continue; }
        if (/^Zap\. ?št\. zastopnika/i.test(k)) { flush(); rec = { t: 'z', zap: v }; continue; }
        if (/^Zap\. ?št\. deleža/i.test(k)) { flush(); rec = { t: 's', zap: v }; continue; }
        if (!rec) continue;
        if (/^Osebno ime|^Firma/i.test(k)) rec.ime = v;
        else if (/^Naslov|^Sedež/i.test(k)) rec.naslov = v;
        else if (/^Vrsta zastopnika/i.test(k)) rec.vrsta = v.toLowerCase();
        else if (/^Način zastopanja/i.test(k)) rec.nacin = v.toLowerCase();
        else if (/^Omejitve/i.test(k)) rec.omejitve = v;
        else if (/^Delež v odstotku/i.test(k)) rec.delez = v;
        else if (/^Imetniki/i.test(k)) rec.imetniki = v;
      }
    }
    flush();
    const up = (s) => (s || '').trim().toUpperCase();
    const pct = (s) => { const m = (s || '').match(/(\d+(?:[.,]\d+)?)\s*%/); if (m) return parseFloat(m[1].replace(',', '.')); const f = (s || '').match(/(\d+)\s*\/\s*(\d+)/); return f ? 100 * f[1] / f[2] : 0; };
    const k = { sifra: '', naziv: (/^[\s/\-–.]*$/.test(top['Skrajšana firma'] || '') ? '' : top['Skrajšana firma']) || top['Firma'] || '', ds: '', maticna: top['Matična številka'] || '', trr: '', ulica: '', posta: '', kraj: '', drzava: 'SLOVENIJA', ddv: '', tel: '', email: '' };
    const dd = (top['Ident. št. za DDV in davčna številka'] || '').match(/^(SI)?\s*(\d{8})/i);
    if (dd) { k.ds = dd[2]; k.ddv = dd[1] ? 'da' : 'ne'; }
    const adr = top['Poslovni naslov'] || '';
    const am = adr.match(/^(.*),\s*(\d{4})\s+(.+)$/);
    if (am) { k.ulica = am[1].trim(); k.posta = am[2]; k.kraj = am[3].trim(); } else { k.ulica = adr; k.kraj = top['Sedež'] || ''; }
    const byZap = {};
    owners.forEach((o) => { o.delez = 0; o.pravna = /d\.?\s?o\.?\s?o|d\.\s?d\.|s\.\s?p\.|zavod|društvo|sklad/i.test(o.ime || '') && !/EMŠO/i.test(''); byZap[o.zap] = o; });
    shares.forEach((s) => {
      const p = pct(s.delez);
      const ids = (s.imetniki || '').match(/\d{4,}/g) || [];
      ids.forEach((id) => { if (byZap[id]) byZap[id].delez += p / (ids.length || 1); });
    });
    return {
      komitent: k, found: !!(k.maticna || k.naziv),
      owners: owners.map((o) => ({ ime: up(o.ime), naslov: o.naslov || '', delez: Math.round(o.delez * 100) / 100 })),
      zast: zast.map((z) => ({ ime: up(z.ime), naslov: z.naslov || '', vrsta: z.vrsta || '', nacin: z.nacin || '', omejitve: z.omejitve || '' })),
    };
  }
  const isSodni = (t) => /Redni izpis iz sodnega|DRUŽBENIKI IN POSLOVNI DELEŽI|OSEBE, POOBLAŠČENE ZA ZASTOPANJE/.test(t || '');

  // AJPES ePRS: besedilo strani (Ctrl+A, kopiraj) -> isti model kot parseScreen
  /* ---- spletni sodni register (e-sodstvo, »Osnovni podatki o subjektu«): oznaka in vrednost v vrsticah ali stolpcih ---- */
  const ES_LABELS = ['status subjekta', 'datum vpisa subjekta', 'matična številka', 'ident. št. za ddv', 'firma', 'skrajšana firma', 'sedež', 'poslovni naslov', 'pravnoorganizacijska oblika', 'osnovni kapital', 'število delnic', 'vrsta organa nadzora', 'register'];
  const isEsodni = (t) => /Skraj[šs]ana\s+firma|Poslovni\s+naslov|Osnovni\s+podatki\s+o\s+subjektu/i.test(t || '') && /Mati[čc]na\s+[šs]tevilka/i.test(t || '');
  function parseEsodni(text) {
    const tok = String(text || '').replace(/\r/g, '').replace(/\u00a0/g, ' ').split(/\n|\t/).map((x) => x.replace(/\s+/g, ' ').trim()).filter(Boolean);
    const isLab = (x) => { const l = x.toLowerCase().replace(/:$/, ''); return ES_LABELS.some((L) => l === L || l.startsWith(L + ' ') && L.length > 8); };
    const lab = (L) => { // vrednost: v isti vrstici za oznako ali naslednji del, ki ni oznaka
      for (let i = 0; i < tok.length; i++) {
        const l = tok[i].toLowerCase();
        if (l === L || l === L + ':') { const v = tok[i + 1]; return v && !isLab(v) ? v : ''; }
        if (l.startsWith(L + ' ') || l.startsWith(L + ':')) { const rest = tok[i].slice(L.length).replace(/^[:\s]+/, ''); if (L.length > 8 && /^(in davčna številka|subjekta v sodni register)/i.test(rest)) { const v = tok[i + 1]; return v && !isLab(v) ? v : ''; } if (rest && !/^(firma)/i.test(rest)) return rest; }
      }
      return '';
    };
    const k = { sifra: '', naziv: '', ds: '', maticna: '', trr: '', ulica: '', posta: '', kraj: '', drzava: 'SLOVENIJA', ddv: '', tel: '', email: '' };
    const firma = lab('firma'), skr = lab('skrajšana firma');
    k.naziv = (skr || firma).trim();
    k.maticna = (lab('matična številka').match(/\d{10}/) || [''])[0];
    const dm = lab('ident. št. za ddv').match(/^(SI)?\s*(\d{8})/i) || (String(text).match(/Ident\.\s*št\.\s*za\s+DDV[^\d]*?(SI)?\s*(\d{8})/i));
    if (dm) { k.ds = dm[2]; k.ddv = dm[1] ? 'da' : 'ne'; }
    const nas = lab('poslovni naslov'), am = nas.match(/^(.*?),\s*(\d{4})\s+(.+)$/);
    if (am) { k.ulica = am[1].trim(); k.posta = am[2]; k.kraj = am[3].trim(); } else { k.ulica = nas; k.kraj = lab('sedež'); }
    const sp = /samostojni\s+podjetnik/i.test(lab('pravnoorganizacijska oblika'));
    return { komitent: k, popolno: firma || k.naziv, zast: [], trrs: [], trrsGB: [], sp, found: !!(k.naziv && (k.maticna || k.ds)) };
  }

  function parseAjpes(text) {
    if (isEsodni(text) && !/popolno ime/i.test(text)) return parseEsodni(text);
    const t = (text || '').replace(/\r/g, '').replace(/\u00a0/g, ' ');
    const lines = t.split('\n').map((l) => l.trim());
    const LABELS = ['popolno ime', 'kratko ime', 'ulica', 'naselje', 'pošta', 'občina', 'regija', 'izpostava ajpes', 'telefon/telefaks/e-pošta', 'matična številka', 'ident. št. za ddv', 'transakcijski računi', 'datum vpisa', 'registrski organ', 'poreklo', 'vrsta lastnine', 'ustanovitelji', 'zastopniki', 'pravnoorganizacijska', 'podoblika', 'sektorska', 'velikost', 'glavna dejavnost', 'druge dejavnosti', 'drugi podatki', 'identifikacijski podatki', 'podatki o ustanovitvi', 'vir:', 'podatki iz drugih'];
    const isLabel = (l) => { const x = l.toLowerCase(); return LABELS.some((p) => x.startsWith(p)); };
    const idx = (label) => lines.findIndex((l) => l.toLowerCase().startsWith(label.toLowerCase()) && (isLabel(l)));
    const inline = (l) => { const c = l.indexOf(':'); return c < 0 ? '' : l.slice(c + 1).trim(); };
    const rest = (label) => {
      const i = idx(label); const out = [];
      if (i < 0) return out;
      const f = inline(lines[i]); if (f) out.push(f);
      for (let j = i + 1; j < lines.length; j++) { if (isLabel(lines[j])) break; if (lines[j]) out.push(lines[j]); }
      return out;
    };
    const val = (label) => rest(label)[0] || '';
    const k = { sifra: '', naziv: '', ds: '', maticna: '', trr: '', ulica: '', posta: '', kraj: '', drzava: 'SLOVENIJA', ddv: '', tel: '', email: '' };
    const real = (v) => (/^[\s/\-–.]*$/.test(v || '') ? '' : v); // »/« ali »-« pomeni, da kratkega imena ni
    k.naziv = (real(val('kratko ime')) || val('popolno ime')).replace(/\.{2,}$/, '.');
    k.ulica = val('ulica');
    const posta = val('pošta');
    const m = posta.match(/^(\d{4})\s+(.+)$/);
    if (m) { k.posta = m[1]; k.kraj = m[2]; } else k.kraj = val('naselje');
    k.maticna = (val('matična številka').match(/\d{10}/) || [''])[0] || ((t.match(/matična številka\s+(\d{10})/) || [])[1] || '');
    const dd = val('ident. št. za DDV');
    const dm = dd.match(/^(SI)?\s*(\d{8})/i);
    if (dm) { k.ds = dm[2]; k.ddv = dm[1] ? 'da' : 'ne'; }
    const trrs = [];
    for (const l of rest('transakcijski računi')) { const x = l.match(/IBAN\s+(SI\d{2}(?:\s?\d{4}){3}\s?\d{3})/i); if (x) trrs.push(x[1].replace(/\s+/g, '')); }
    // obrazci so za Gorenjsko banko: vpiše se samo TRR pri nas (07000), računi drugih bank nikoli
    const trrsGB = trrs.filter((x) => x.startsWith('SI560700')).map((x) => x.slice(4, 9) + '-' + x.slice(9));
    if (trrsGB.length) k.trr = trrsGB[0];
    const ct = rest('telefon/telefaks/e-pošta');
    for (const l of ct) {
      const e = l.match(/[^\s@]+@[^\s@]+\.[a-z]{2,}/i); if (e && !k.email) k.email = e[0];
      const p = l.match(/^[+\d][\d\s\/-]{6,}(?=\s*\((?:tel|mobil|gsm|telefaks))/i); if (p && !k.tel) k.tel = p[0].trim();
    }
    const zast = [];
    for (const l of rest('zastopniki')) {
      const x = l.match(/^(.+?),\s*(.+)$/); if (!x) continue;
      zast.push({ ime: x[1].trim().toUpperCase(), vloga: x[2].trim().toLowerCase() });
    }
    // s.p.: zastopnikov AJPES ne navaja; nosilec (ustanovitelj) je zakoniti zastopnik in podpisnik
    const sp = /samostojni\s+podjetnik/i.test(rest('pravnoorganizacijska').join(' ')) || /\bs\.\s*p\.?\s*$/i.test(k.naziv || '');
    if (!zast.length && sp) for (const l of rest('ustanovitelji')) { const nm = l.replace(/,.*$/, '').trim(); if (/^[A-ZČŠŽĆĐ][^\d]+\s+\S+/.test(nm)) zast.push({ ime: nm.toUpperCase(), vloga: 'nosilec dejavnosti' }); }
    return { komitent: k, popolno: val('popolno ime').replace(/\.{2,}$/, '.'), zast, trrs, trrsGB, found: !!(k.naziv || k.maticna) };
  }

  const api = { isEsodni, parseEsodni, sifraFO, sifraPO, parseHibisPO, isHibisPO, mergeHibisPO, parseSodni, isSodni, refineEmail, parseAjpes, fixEmail, parseHibis, isHibis, mergeHibis, mergeParsed, parseScreen, buildPersons, emptyPerson, checkDS, checkIBAN, ibanOf, trrDigits, fixAddress, fixCompany, fixCity, fixPostCity };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.Parser = api;
})(typeof window !== 'undefined' ? window : globalThis);
