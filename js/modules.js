/* Cel voor Cel – inhoud van de modules: lessen, Excel-opdrachten, hulplijnen en controles.
   Elke opdrachtenreeks wordt gemaakt uit een 'seed', zodat dezelfde seed altijd dezelfde opdrachten geeft
   (de leerkracht ziet zo precies welke opdrachten een leerling kreeg). */
(function (root) {
  'use strict';
  const X = (typeof module !== 'undefined' && module.exports) ? require('./xlsx-check.js') : root.XlsxCheck;
  const { addr, numToCol, parseAddr, eachCell, norm, colorNames, fitCheck } = X;

  // ---------- toeval met seed ----------
  function hashStr(s) { let h = 1779033703 ^ s.length; for (let i = 0; i < s.length; i++) { h = Math.imul(h ^ s.charCodeAt(i), 3432918353); h = h << 13 | h >>> 19; } return h >>> 0; }
  function rng(seed) {
    let a = hashStr(String(seed));
    const next = () => { a |= 0; a = a + 0x6D2B79F5 | 0; let t = Math.imul(a ^ a >>> 15, 1 | a); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; };
    next.int = (lo, hi) => lo + Math.floor(next() * (hi - lo + 1));
    next.pick = arr => arr[Math.floor(next() * arr.length)];
    next.shuffle = arr => { const a = arr.slice(); for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(next() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; };
    return next;
  }
  function newSeed() { return Math.random().toString(36).slice(2, 10); }

  const ok = msg => ({ ok: true, msg: msg || 'Juist!' });
  const no = msg => ({ ok: false, msg });
  const list = arr => arr.length <= 1 ? arr.join('') : arr.slice(0, -1).join(', ') + ' en ' + arr[arr.length - 1];
  const cleanName = s => norm(s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]/g, '');

  /* Een vrije cel kiezen die nog niet gebruikt is. */
  function freeCell(r, used, c1, c2, r1, r2) {
    for (let i = 0; i < 200; i++) {
      const a = addr(r.int(c1, c2), r.int(r1, r2));
      if (!used.has(a)) { used.add(a); return a; }
    }
    throw new Error('geen vrije cel');
  }

  const KLEUREN = ['geel', 'groen', 'blauw', 'rood', 'oranje', 'paars'];

  // ======================================================================
  // MODULE 1 – Rijen, kolommen en cellen
  // ======================================================================
  const m1 = {
    id: 'm1', nr: 1,
    title: 'Rijen, kolommen en cellen',
    short: 'Je leert hoe Excel in elkaar zit: kolommen, rijen, cellen en celadressen. Je opent Excel en bewaart je eerste bestand in OneDrive.',
    skills: ['werkmap en werkblad', 'kolom, rij en cel', 'celadres', 'bereik', 'opslaan in OneDrive'],
    lessons: [
      { id: 'l1', title: 'Excel is een groot rooster',
        body: `<p>Een Excel-bestand heet een <b>werkmap</b>. In een werkmap zitten één of meer <b>werkbladen</b>. Onderaan zie je ze als tabbladen.</p>
<p>Elk werkblad is een groot rooster. De vakjes van dat rooster heten <b>cellen</b>. In een cel kan je tekst, een getal of een berekening zetten.</p>`,
        widget: { type: 'quiz', questions: [
          { q: 'Hoe heet een Excel-bestand?', options: ['Een werkmap', 'Een werkblad', 'Een cel'], correct: 0, explain: 'Het bestand is de werkmap. Daarin zitten de werkbladen.' },
          { q: 'Waar zie je de werkbladen van een werkmap?', options: ['Bovenaan, in het lint', 'Onderaan, als tabbladen', 'Je kan ze niet zien'], correct: 1, explain: 'Onderaan het scherm staan tabbladen, bv. "Blad1".' },
          { q: 'Hoe heten de vakjes in het rooster?', options: ['Kolommen', 'Cellen', 'Rijen'], correct: 1, explain: 'Elk vakje is een cel.' }
        ] } },
      { id: 'l2', title: 'Kolommen en rijen',
        body: `<p>Een <b>kolom</b> loopt van boven naar onder. Kolommen hebben een <b>letter</b>: A, B, C …</p>
<p>Een <b>rij</b> loopt van links naar rechts. Rijen hebben een <b>getal</b>: 1, 2, 3 …</p>
<p>Klik in het rooster hieronder op de letter of het getal dat gevraagd wordt.</p>`,
        widget: { type: 'gridHeader', rounds: 4 } },
      { id: 'l3', title: 'Elke cel heeft een adres',
        body: `<p>Een cel ligt altijd in één kolom én in één rij. Daarom heeft elke cel een <b>celadres</b>: eerst de letter van de kolom, dan het getal van de rij.</p>
<p>De cel in kolom <b>D</b> en rij <b>15</b> heet dus <span class="addr">D15</span>. In Excel zie je het adres van de gekozen cel linksboven, in het <b>naamvak</b>.</p>`,
        widget: { type: 'gridClick', rounds: 4 } },
      { id: 'l4', title: 'Lees het adres',
        body: `<p>Nu andersom: één cel is groen omrand. Typ het adres van die cel. Eerst de letter, dan het getal.</p>`,
        widget: { type: 'gridRead', rounds: 3 } },
      { id: 'l5', title: 'Een bereik: meerdere cellen samen',
        body: `<p>Soms wil je meerdere cellen tegelijk kiezen. Dat heet een <b>bereik</b>. Je schrijft de cel linksboven en de cel rechtsonder, met een dubbelpunt ertussen.</p>
<p><span class="addr">B2:D4</span> betekent: alle cellen van B2 tot en met D4. Dat zijn er 9.</p>
<p>Sleep met de muis over het rooster om het gevraagde bereik te kiezen.</p>`,
        widget: { type: 'gridRange', rounds: 2 } },
      { id: 'l6', title: 'Excel openen en bewaren in OneDrive',
        body: `<p>Zo werk je in Excel. Doe dit elke keer op dezelfde manier.</p>`,
        widget: { type: 'steps', steps: [
          'Start <b>Excel</b> via het startmenu van de computer.',
          'Kijk rechtsboven: staat daar je <b>schoolaccount</b>? Dan is het goed.',
          'Kies <b>Lege werkmap</b>.',
          'Klik op <b>Bestand</b> → <b>Opslaan als</b> → <b>OneDrive</b> (van je school).',
          'Typ een goede <b>bestandsnaam</b>, bv. <span class="addr">Lotte_module1</span>, en klik op <b>Opslaan</b>.',
          'De volgende keer: start Excel en open je bestand onder <b>Recent</b>.'
        ], confirm: 'Ik weet hoe ik Excel open en bewaar' } }
    ],
    newFile: true,
    gen(seed, mode, student) {
      const r = rng(seed + '|m1|' + mode);
      const used = new Set();
      const c1 = freeCell(r, used, 1, 12, 1, 20), c2 = freeCell(r, used, 1, 12, 1, 20);
      const c3 = freeCell(r, used, 1, 10, 1, 20);
      // bereik van 2 kolommen x 3 rijen
      let rng0;
      for (let i = 0; i < 100; i++) {
        const cc = r.int(2, 10), rr = r.int(2, 16);
        const cells = []; for (let a = 0; a < 2; a++) for (let b = 0; b < 3; b++) cells.push(addr(cc + a, rr + b));
        if (cells.every(x => !used.has(x))) { cells.forEach(x => used.add(x)); rng0 = addr(cc, rr) + ':' + addr(cc + 1, rr + 2); break; }
      }
      const n = r.int(2, 9);
      const dirs = [
        { naam: 'rechts naast', dc: 1, dr: 0 }, { naam: 'onder', dc: 0, dr: 1 },
        { naam: 'links naast', dc: -1, dr: 0 }, { naam: 'boven', dc: 0, dr: -1 }
      ];
      const dir = mode === 'test' ? r.pick(dirs.slice(2)) : r.pick(dirs.slice(0, 2));
      let c4, buurAddr;
      for (let i = 0; i < 200; i++) {
        const p = { c: r.int(2, 11), r: r.int(2, 19) };
        const a = addr(p.c, p.r), b = addr(p.c + dir.dc, p.r + dir.dr);
        if (!used.has(a) && !used.has(b)) { used.add(a); used.add(b); c4 = a; buurAddr = b; break; }
      }
      const woord = r.pick(['buur', 'hier', 'ok', 'top']);
      const bladNaam = r.pick(mode === 'test' ? ['Test', 'Mijn test', 'Proef'] : ['Oefening', 'Mijn blad', 'Start']);
      const fileWord = mode === 'test' ? 'test1' : 'module1';
      const vn = (student && student.voornaam) || 'Voornaam';
      const fileSuggest = vn.replace(/\s+/g, '') + '_' + fileWord;

      return { startFile: null, tasks: [
        { id: 't1', text: `Typ een <b>x</b> in cel <span class="addr">${c1}</span> en in cel <span class="addr">${c2}</span>. Nergens anders een x.`,
          hints: ['Zoek eerst de kolom (de letter), ga dan naar beneden tot de rij (het getal).', `Klik op een cel en kijk linksboven in het naamvak: daar staat het adres van de cel waar je staat. Staat er ${c1}? Typ dan x en druk op Enter.`],
          check: ({ sheet }) => {
            const missing = [c1, c2].filter(a => norm(sheet.text(a)) !== 'x');
            const extra = sheet.find(v => norm(v) === 'x').filter(a => a !== c1 && a !== c2);
            if (missing.length && extra.length) return no(`In ${list(missing)} staat nog geen x. Er staat wel een x in ${list(extra)}. Controleer de letter en het getal.`);
            if (missing.length) return no(`In ${list(missing)} staat nog geen x.`);
            if (extra.length) return no(`Goed in ${c1} en ${c2}, maar er staat ook een x in ${list(extra)}. Die moet weg.`);
            return ok();
          } },
        { id: 't2', text: `Typ je <b>voornaam</b> in cel <span class="addr">${c3}</span>.`,
          hints: ['Klik op de cel, typ je voornaam en druk op Enter.', `Kolom ${c3.replace(/\d+/, '')} vind je bovenaan, rij ${c3.replace(/[A-Z]+/, '')} vind je links.`],
          check: ({ sheet }) => {
            if (norm(sheet.text(c3)) === norm(vn)) return ok();
            const elders = sheet.findText(vn);
            if (elders.length) return no(`Je voornaam staat in ${list(elders)}, niet in ${c3}.`);
            if (!sheet.isEmpty(c3)) return no(`In ${c3} staat "${sheet.text(c3)}". Typ daar je voornaam: ${vn}.`);
            return no(`Cel ${c3} is nog leeg.`);
          } },
        { id: 't3', text: `Typ het getal <b>${n}</b> in <b>elke</b> cel van het bereik <span class="addr">${rng0}</span>.`,
          hints: ['Een bereik gaat van de cel linksboven tot de cel rechtsonder. Hoeveel cellen zijn dat?', `${rng0} bestaat uit 6 cellen: 2 kolommen breed en 3 rijen hoog.`],
          check: ({ sheet }) => {
            const wrong = []; eachCell(rng0, a => { if (sheet.value(a) !== n && norm(sheet.text(a)) !== String(n)) wrong.push(a); });
            if (!wrong.length) return ok();
            if (wrong.length === 6) return no(`In het bereik ${rng0} staat nog niet het getal ${n}.`);
            return no(`Bijna! In ${list(wrong)} staat nog niet ${n}.`);
          } },
        { id: 't4', text: `Typ het woord <b>${woord}</b> in de cel <b>${dir.naam}</b> cel <span class="addr">${c4}</span>.`,
          hints: [`Zoek eerst ${c4}. Welke cel ligt daar ${dir.naam}?`, dir.dc ? `${dir.naam === 'rechts naast' ? 'Rechts' : 'Links'} naast een cel verandert de letter, het getal blijft hetzelfde.` : `${dir.naam === 'onder' ? 'Onder' : 'Boven'} een cel blijft de letter hetzelfde, het getal verandert.`],
          check: ({ sheet }) => {
            if (norm(sheet.text(buurAddr)) === woord) return ok();
            const elders = sheet.findText(woord);
            if (elders.length) return no(`Je typte "${woord}" in ${list(elders)}. Dat is niet de cel ${dir.naam} ${c4}.`);
            return no(`Het woord "${woord}" staat nog nergens.`);
          } },
        { id: 't5', text: `Geef het werkblad de naam <b>${bladNaam}</b>.`,
          hints: ['Het werkblad zie je onderaan als tabblad. Nu heet het waarschijnlijk "Blad1".', 'Dubbelklik op het tabblad onderaan, typ de nieuwe naam en druk op Enter.'],
          check: ({ wb }) => {
            if (wb.sheet(bladNaam)) return ok();
            return no(`Het werkblad heet nu "${wb.sheetNames().join('", "')}". Het moet "${bladNaam}" heten.`);
          } },
        { id: 't6', text: `Bewaar het bestand met de naam <span class="addr">${fileSuggest}</span>.`,
          hints: ['Bewaren doe je via Bestand → Opslaan als.', `Typ bij bestandsnaam: ${fileSuggest}. Gebruik een liggend streepje ( _ ) tussen je naam en ${fileWord}.`],
          check: ({ fileName }) => {
            const f = cleanName((fileName || '').replace(/\.xlsx$/i, ''));
            if (f.includes(cleanName(vn)) && f.includes(fileWord)) return ok();
            return no(`Je bestand heet "${fileName}". Het moet "${fileSuggest}" heten.`);
          } }
      ] };
    }
  };

  // ======================================================================
  // MODULE 2 – Typen, verbeteren en kolombreedte
  // ======================================================================
  const PRODUCTEN = [
    ['appels', 'apels'], ['bananen', 'banaanen'], ['tomaten', 'tomatten'], ['komkommer', 'komkomer'],
    ['wortelen', 'wortellen'], ['aardbeien', 'aarbeien'], ['broccoli', 'brocoli'], ['spinazie', 'spinazi'],
    ['melk', 'mellek'], ['kaas', 'kaaz'], ['eieren', 'eiren'], ['boter', 'botter'], ['yoghurt', 'yogurt'], ['peren', 'perren']
  ];
  const NIEUW = ['brood', 'rijst', 'pasta', 'soep', 'honing', 'water'];
  const LANG = ['verjaardagsfeest', 'boterhammendoos', 'schoolvoetbaltoernooi', 'wiskundeleerkracht', 'speelplaatsfeest', 'zwembadabonnement'];
  const TWEE = [['Excel', 'is handig'], ['Hallo', 'wereld'], ['Goed', 'gedaan'], ['Lijn 1', 'lijn 2'], ['Stap', 'voor stap'], ['Cel', 'voor cel']];

  function m2data(seed, mode) {
    const r = rng(seed + '|m2|' + mode);
    const prods = r.shuffle(PRODUCTEN).slice(0, 7);
    const typoIdx = r.int(0, 6);
    const amounts = prods.map(() => r.int(1, 9));
    const insertIdx = r.int(1, 6);              // nooit boven het eerste product
    const nieuw = r.pick(NIEUW), nieuwN = r.int(2, 8);
    const rommelIdx = r.int(0, 6);
    const cols = r.shuffle(['E', 'F', 'G', 'H']);
    const lang = r.pick(LANG), twee = r.pick(TWEE);
    const langCell = cols[0] + '1', tweeCell = cols[1] + '3', kopieCell = cols[2] + '2';
    return { prods, typoIdx, amounts, insertIdx, nieuw, nieuwN, rommelIdx, lang, twee, langCell, tweeCell, kopieCell };
  }

  const m2 = {
    id: 'm2', nr: 2,
    title: 'Typen, verbeteren en kolombreedte',
    short: 'Je typt en verbetert in cellen, maakt een rij bij, zet twee lijnen in één cel en laat Excel de kolombreedte zelf aanpassen.',
    skills: ['typen en bevestigen', 'verbeteren en wissen', 'rij invoegen', 'kopiëren en plakken', 'Alt + Enter', 'kolombreedte automatisch'],
    lessons: [
      { id: 'l1', title: 'Typen en bevestigen',
        body: `<p>Klik op een cel en begin te typen. Wat je typt, zie je ook bovenaan in de <b>formulebalk</b>.</p>
<p>Klaar? Druk op <span class="key">Enter</span>: je gaat naar de cel <b>eronder</b>. Met <span class="key">Tab</span> ga je naar de cel <b>rechts</b>. Vergist? Met <span class="key">Esc</span> stop je zonder iets te bewaren.</p>`,
        widget: { type: 'quiz', questions: [
          { q: 'Je typt in cel B3 en drukt op Enter. Waar sta je dan?', options: ['B4', 'C3', 'B2'], correct: 0, explain: 'Enter brengt je naar de cel eronder: B4.' },
          { q: 'Je typt in cel B3 en drukt op Tab. Waar sta je dan?', options: ['B4', 'C3', 'A3'], correct: 1, explain: 'Tab brengt je naar de cel rechts: C3.' },
          { q: 'Je bent iets fout aan het typen en wil stoppen zonder te bewaren. Welke toets?', options: ['Enter', 'Delete', 'Esc'], correct: 2, explain: 'Esc annuleert wat je aan het typen was.' }
        ] } },
      { id: 'l2', title: 'Verbeteren, wissen en ongedaan maken',
        body: `<p><b>Alles opnieuw typen?</b> Klik op de cel en typ gewoon: het oude verdwijnt.</p>
<p><b>Een klein stukje verbeteren?</b> Dubbelklik op de cel of druk op <span class="key">F2</span>. Nu kan je in de tekst klikken en één letter aanpassen.</p>
<p><b>Een cel leegmaken?</b> Klik op de cel en druk op <span class="key">Delete</span>.</p>
<p><b>Oei, foutje?</b> Met <span class="key">Ctrl</span> + <span class="key">Z</span> maak je de laatste stap ongedaan.</p>`,
        widget: { type: 'keys', rounds: [
          { ask: 'Met welke toets kan je de tekst in een cel aanpassen zonder alles opnieuw te typen?', answer: ['F2'] },
          { ask: 'Met welke toetsen maak je een fout ongedaan?', answer: ['Ctrl', 'Z'] },
          { ask: 'Met welke toets maak je een cel leeg?', answer: ['Delete'] }
        ] } },
      { id: 'l3', title: 'Twee lijnen in één cel',
        body: `<p>Als je in een cel op <span class="key">Enter</span> drukt, spring je naar de volgende cel. Maar soms wil je <b>in dezelfde cel</b> een nieuwe lijn beginnen.</p>
<p>Dat doe je met <span class="key">Alt</span> + <span class="key">Enter</span>: hou <b>Alt</b> ingedrukt en druk op <b>Enter</b>. Excel zet de tekst dan op twee lijnen en maakt de rij vanzelf hoger.</p>`,
        widget: { type: 'keys', rounds: [
          { ask: 'Welke toetsen druk je in om in dezelfde cel een nieuwe lijn te beginnen?', answer: ['Alt', 'Enter'] },
          { ask: 'En welke toets brengt je naar de cel eronder?', answer: ['Enter'] }
        ] } },
      { id: 'l4', title: 'De kolom laten passen',
        body: `<p>Is een woord te lang voor de cel? Dan loopt het door in de buurcel, of je ziet maar een stukje.</p>
<p>Je kan de kolom breder slepen, maar het kan slimmer: <b>dubbelklik op de lijn rechts van de kolomletter</b>. Excel maakt de kolom dan precies zo breed als het langste woord in die kolom.</p>
<p>Probeer het hieronder: dubbelklik op de lijn rechts van de letter van elke kolom met een te lang woord.</p>`,
        widget: { type: 'autofit' } },
      { id: 'l5', title: 'Een rij of kolom bijmaken',
        body: `<p>Vergeten iets te typen in het midden van een lijst? Dan maak je een <b>nieuwe rij</b>.</p>
<p><b>Rechtsklik op het rijnummer</b> (bv. 5) en kies <b>Invoegen</b>. Er komt een lege rij <b>boven</b> rij 5. Alles daaronder schuift één rij naar beneden.</p>
<p>Voor een kolom doe je hetzelfde: rechtsklik op de kolomletter en kies <b>Invoegen</b>. De nieuwe kolom komt <b>links</b>.</p>`,
        widget: { type: 'quiz', questions: [
          { q: 'Je rechtsklikt op rijnummer 6 en kiest Invoegen. Waar komt de nieuwe rij?', options: ['Boven rij 6', 'Onder rij 6', 'Helemaal onderaan'], correct: 0, explain: 'De nieuwe rij komt boven de rij waarop je klikte. De oude rij 6 wordt rij 7.' },
          { q: 'In A7 staat "kaas". Je voegt een rij in boven rij 5. Waar staat "kaas" nu?', options: ['A7', 'A8', 'A6'], correct: 1, explain: 'Alles onder de nieuwe rij schuift één plaats naar beneden: A7 wordt A8.' },
          { q: 'Waar klik je met de rechtermuisknop om een kolom in te voegen?', options: ['Op een cel in het midden', 'Op de kolomletter', 'Op het tabblad onderaan'], correct: 1, explain: 'Rechtsklik op de letter bovenaan en kies Invoegen.' }
        ] } },
      { id: 'l6', title: 'Kopiëren en plakken',
        body: `<p>Wil je dezelfde inhoud in een andere cel? Typ het niet opnieuw.</p>
<p>Klik op de cel en druk op <span class="key">Ctrl</span> + <span class="key">C</span> (kopiëren). Er komt een bewegende stippellijn rond de cel. Klik op de nieuwe cel en druk op <span class="key">Ctrl</span> + <span class="key">V</span> (plakken).</p>`,
        widget: { type: 'keys', rounds: [
          { ask: 'Met welke toetsen kopieer je een cel?', answer: ['Ctrl', 'C'] },
          { ask: 'Met welke toetsen plak je?', answer: ['Ctrl', 'V'] }
        ] } }
    ],
    startFile(seed, mode) {
      const d = m2data(seed, mode);
      const cells = { A1: { v: 'Boodschappenlijst', bold: true, size: 14 }, A2: { v: 'Product', bold: true }, B2: { v: 'Aantal', bold: true } };
      d.prods.forEach((p, i) => { cells['A' + (3 + i)] = i === d.typoIdx ? p[1] : p[0]; cells['B' + (3 + i)] = d.amounts[i]; });
      cells['D' + (3 + d.rommelIdx)] = 'WEG MET MIJ';
      return { name: mode === 'test' ? 'Test2_boodschappen.xlsx' : 'Module2_boodschappen.xlsx', spec: { sheetName: 'Boodschappen', cells, colWidths: { A: 16 } } };
    },
    gen(seed, mode) {
      const d = m2data(seed, mode);
      const typo = d.prods[d.typoIdx], ins = d.prods[d.insertIdx], before = d.prods[d.insertIdx - 1];
      const col = s => X.colToNum(s.replace(/\d+/, ''));
      return { startFile: true, tasks: [
        { id: 't1', text: `In de lijst staat een schrijffout: <b>${typo[1]}</b>. Verbeter het naar <b>${typo[0]}</b>.`,
          hints: ['Klik op de cel met de fout. Je kan alles opnieuw typen, of je dubbelklikt om één letter aan te passen.', 'Druk op F2 om in de cel te verbeteren. Druk op Enter als je klaar bent.'],
          check: ({ sheet }) => {
            const fout = sheet.findText(typo[1]), goed = sheet.findText(typo[0]);
            if (fout.length) return no(`In ${fout[0]} staat nog altijd "${typo[1]}".`);
            if (!goed.length) return no(`Ik vind "${typo[0]}" niet in de lijst. Controleer je spelling.`);
            return ok();
          } },
        { id: 't2', text: `In kolom D staat <b>WEG MET MIJ</b>. Maak die cel helemaal leeg.`,
          hints: ['Een cel leegmaken gaat met één toets.', 'Klik op de cel en druk op Delete. Spatiebalk is niet goed: dan staat er een onzichtbare spatie in.'],
          check: ({ sheet }) => {
            const nog = sheet.find(v => norm(v).includes('weg met mij'));
            if (nog.length) return no(`In ${nog[0]} staat nog "${sheet.text(nog[0])}".`);
            const spatie = Object.keys(sheet.cells).some(a => { const v = sheet.value(a); return /^D\d+$/.test(a) && typeof v === 'string' && v.length > 0 && v.trim() === ''; });
            if (spatie) return no('De tekst is weg, maar er staat nog een spatie in de cel. Klik erop en druk op Delete.');
            return ok();
          } },
        { id: 't3', text: `Voeg een nieuwe rij in <b>boven ${ins[0]}</b>. Typ daar <b>${d.nieuw}</b> in kolom A en <b>${d.nieuwN}</b> in kolom B.`,
          hints: [`Zoek de rij van "${ins[0]}". Je maakt de nieuwe rij met de rechtermuisknop.`, `Rechtsklik op het rijnummer van "${ins[0]}" (links) en kies Invoegen. Typ dan in de lege rij.`],
          check: ({ sheet }) => {
            const insName = d.insertIdx === d.typoIdx ? [ins[0], ins[1]] : [ins[0]];
            const pos = sheet.find((v, a) => /^A\d+$/.test(a) && insName.includes(norm(v)));
            const newPos = sheet.find((v, a) => /^A\d+$/.test(a) && norm(v) === d.nieuw);
            if (!newPos.length) return no(`Ik vind "${d.nieuw}" niet in kolom A.`);
            if (!pos.length) return no(`Ik vind "${ins[0]}" niet meer in kolom A. Heb je het overschreven?`);
            const rIns = +pos[0].slice(1), rNew = +newPos[0].slice(1);
            if (rNew !== rIns - 1) return no(`"${d.nieuw}" staat in rij ${rNew}, maar het moet in de rij net boven "${ins[0]}" (rij ${rIns}) staan.`);
            const bName = d.insertIdx - 1 === d.typoIdx ? [before[0], before[1]] : [before[0]];
            if (!bName.includes(norm(sheet.text('A' + (rNew - 1))))) return no(`Boven "${d.nieuw}" moet nog altijd "${before[0]}" staan. Heb je een rij overschreven in plaats van er een in te voegen?`);
            if (sheet.value('B' + rNew) !== d.nieuwN && norm(sheet.text('B' + rNew)) !== String(d.nieuwN)) return no(`Typ ${d.nieuwN} in B${rNew}, naast "${d.nieuw}".`);
            return ok();
          } },
        { id: 't4', text: `Typ <b>${d.lang}</b> in cel <span class="addr">${d.langCell}</span>. Maak kolom ${d.langCell[0]} daarna <b>automatisch</b> zo breed als het woord.`,
          hints: ['Slepen mag, maar Excel kan het precies voor jou doen.', `Dubbelklik bovenaan op de lijn tussen de letters ${d.langCell[0]} en ${numToCol(col(d.langCell) + 1)}.`],
          check: ({ sheet }) => {
            if (norm(sheet.text(d.langCell)) !== d.lang) {
              const el = sheet.findText(d.lang);
              return no(el.length ? `"${d.lang}" staat in ${el[0]}, niet in ${d.langCell}.` : `In ${d.langCell} staat nog niet "${d.lang}". Let op de spelling.`);
            }
            const f = fitCheck(sheet, col(d.langCell), d.lang);
            if (f.ok) return ok();
            if (f.why === 'smal') return no(`Kolom ${d.langCell[0]} is nog te smal: het woord past er niet in. Dubbelklik op de lijn rechts van de letter ${d.langCell[0]}.`);
            return no(`Kolom ${d.langCell[0]} is veel breder dan nodig. Dubbelklik op de lijn rechts van de letter ${d.langCell[0]}: dan wordt hij precies passend.`);
          } },
        { id: 't5', text: `Typ in cel <span class="addr">${d.tweeCell}</span> op <b>twee lijnen</b>: bovenaan <b>${d.twee[0]}</b>, daaronder <b>${d.twee[1]}</b>.`,
          hints: ['In dezelfde cel een nieuwe lijn beginnen gaat met twee toetsen samen.', `Typ "${d.twee[0]}", hou Alt ingedrukt en druk op Enter. Typ dan "${d.twee[1]}" en druk op Enter.`],
          check: ({ sheet }) => {
            const t = sheet.text(d.tweeCell);
            if (!t.trim()) return no(`Cel ${d.tweeCell} is nog leeg.`);
            const lines = t.split('\n').map(norm).filter(Boolean);
            if (lines.length < 2) return no(`De tekst staat op één lijn. Gebruik Alt + Enter tussen "${d.twee[0]}" en "${d.twee[1]}".`);
            if (lines[0] !== norm(d.twee[0]) || lines[1] !== norm(d.twee[1])) return no(`Er staan twee lijnen, goed! Maar de woorden kloppen niet helemaal. Bovenaan "${d.twee[0]}", daaronder "${d.twee[1]}".`);
            return ok();
          } },
        { id: 't6', text: `Kopieer het woord <b>Aantal</b> (cel <span class="addr">B2</span>) en plak het in cel <span class="addr">${d.kopieCell}</span>.`,
          hints: ['Kopiëren en plakken gaat met twee sneltoetsen.', `Klik op B2, druk Ctrl + C. Klik op ${d.kopieCell}, druk Ctrl + V.`],
          check: ({ sheet }) => {
            if (norm(sheet.text(d.kopieCell)) !== 'aantal') return no(`In ${d.kopieCell} staat nog niet "Aantal".`);
            if (norm(sheet.text('B2')) !== 'aantal') return no('Het woord "Aantal" moet ook nog in B2 blijven staan. Kopiëren, niet knippen.');
            return ok();
          } }
      ] };
    }
  };

  // ======================================================================
  // MODULE 3 – Cellen opmaken
  // ======================================================================
  const TABELLEN = [
    { titel: 'Onze lievelingsdieren', kop: ['Naam', 'Dier', 'Kleur'], rijen: [['Lotte', 'kat', 'grijs'], ['Milan', 'hond', 'bruin'], ['Noor', 'paard', 'wit'], ['Yasin', 'konijn', 'zwart']] },
    { titel: 'Sporten in onze klas', kop: ['Naam', 'Sport', 'Uren'], rijen: [['Emma', 'voetbal', 3], ['Rayan', 'judo', 2], ['Fien', 'turnen', 4], ['Lucas', 'zwemmen', 2]] },
    { titel: 'Onze uitstap', kop: ['Wat', 'Waar', 'Prijs'], rijen: [['bus', 'school', 6], ['toegang', 'museum', 8], ['picknick', 'park', 4], ['ijsje', 'kraam', 2]] },
    { titel: 'Planeten', kop: ['Planeet', 'Manen', 'Plaats'], rijen: [['Mercurius', 0, 1], ['Venus', 0, 2], ['Aarde', 1, 3], ['Mars', 2, 4]] }
  ];

  function m3data(seed, mode) {
    const r = rng(seed + '|m3|' + mode);
    const tab = r.pick(TABELLEN);
    const cell = addr(r.int(6, 11), r.int(9, 16));     // 'leren'-cel, buiten de tabel
    const k = r.shuffle(KLEUREN);
    const size = r.pick([16, 18, 20]);
    const woord = mode === 'test' ? r.pick(['oefenen', 'proberen', 'durven']) : 'leren';
    return { tab, cell, fontKleur: k[0], vulKleur: k[1], kopKleur: k[2], size, woord, row: +cell.replace(/[A-Z]+/, '') };
  }

  const m3 = {
    id: 'm3', nr: 3,
    title: 'Cellen opmaken',
    short: 'Je maakt tekst vet, schuin en onderstreept, kiest kleuren, zet tekst in het midden, voegt cellen samen en omlijnt een tabel.',
    skills: ['vet, schuin, onderstreept', 'letterkleur en opvulkleur', 'uitlijnen', 'rijhoogte', 'samenvoegen en centreren', 'lettergrootte', 'randen'],
    lessons: [
      { id: 'l1', title: 'Het lint: tab Start',
        body: `<p>Bovenaan in Excel zie je het <b>lint</b> met tabbladen: Bestand, Start, Invoegen … Bijna alles voor opmaak staat op de tab <b>Start</b>.</p>
<p>De knoppen staan in <b>groepen</b>. In de groep <b>Lettertype</b> vind je vet (<b>B</b>), schuin (<i>I</i>), onderstreept (<u>U</u>), lettergrootte, letterkleur (een <b>A</b> met een kleurstreep) en opvulkleur (een <b>verfemmer</b>). In de groep <b>Uitlijning</b> zet je tekst links, in het midden of rechts, en boven, in het midden of onder.</p>
<p>Belangrijk: <b>eerst de cel aanklikken, dan pas de knop</b>.</p>`,
        widget: { type: 'quiz', questions: [
          { q: 'Op welke tab vind je de meeste knoppen voor opmaak?', options: ['Invoegen', 'Start', 'Bestand'], correct: 1, explain: 'De tab Start heeft de groepen Lettertype en Uitlijning.' },
          { q: 'Welke knop maakt tekst vet?', options: ['<b>B</b>', '<i>I</i>', '<u>U</u>'], correct: 0, explain: 'B komt van "bold", het Engelse woord voor vet.' },
          { q: 'Je wil de achtergrond van een cel geel kleuren. Welke knop?', options: ['De A met de kleurstreep', 'De verfemmer', 'De B'], correct: 1, explain: 'De verfemmer is de opvulkleur. De A verandert de kleur van de letters.' },
          { q: 'Wat doe je eerst?', options: ['Op de knop klikken', 'De cel aanklikken'], correct: 1, explain: 'Excel moet weten welke cel je wil opmaken.' }
        ] } },
      { id: 'l2', title: 'Oefen de opmaakknoppen',
        body: `<p>Maak jouw cel precies hetzelfde als het voorbeeld. Gebruik de knoppen, net zoals in Excel.</p>`,
        widget: { type: 'formatSim', rounds: 3 } },
      { id: 'l3', title: 'Samenvoegen en centreren',
        body: `<p>Een titel boven een tabel staat mooi in het midden van de tabel. Daarvoor maak je van een paar cellen <b>één grote cel</b>.</p>
<p>Selecteer de cellen (bv. <span class="addr">B2:D2</span>) en klik in de groep Uitlijning op <b>Samenvoegen en centreren</b>.</p>
<p>Een rij hoger maken? Sleep de lijn onder het rijnummer naar beneden. Daarna kan je tekst ook <b>verticaal</b> in het midden zetten.</p>`,
        widget: { type: 'quiz', questions: [
          { q: 'Je voegt B2:D2 samen. Hoeveel cellen zie je daarna in B2 tot D2?', options: ['3', '1', '2'], correct: 1, explain: 'Samenvoegen maakt er één grote cel van.' },
          { q: 'Je wil tekst ook van boven naar onder in het midden. Wat heb je nodig?', options: ['De knop "Midden uitlijnen"', 'Vet maken', 'De letterkleur'], correct: 0, explain: 'In de groep Uitlijning staan knoppen voor boven, midden en onder.' }
        ] } },
      { id: 'l4', title: 'Randen rond een tabel',
        body: `<p>De grijze lijntjes die je in Excel ziet, worden <b>niet afgedrukt</b>. Wil je echte lijnen, dan zet je <b>randen</b>.</p>
<p>Selecteer de tabel en klik in de groep Lettertype op het pijltje naast de <b>randen-knop</b> (een vierkantje met lijntjes). Kies <b>Alle randen</b>: dan krijgt elke cel een lijn rondom. <b>Buitenrand</b> zet alleen een lijn rond het geheel.</p>`,
        widget: { type: 'quiz', questions: [
          { q: 'Je kiest "Buitenrand" voor B4:D8. Wat zie je?', options: ['Een lijn rond elke cel', 'Een lijn rond de hele tabel, niet ertussen', 'Geen lijnen'], correct: 1, explain: 'Buitenrand tekent alleen de rand rond het geheel.' },
          { q: 'Je wil een lijn rond elke cel van de tabel. Wat kies je?', options: ['Alle randen', 'Buitenrand', 'Geen rand'], correct: 0, explain: 'Alle randen zet een lijn rond elke cel.' },
          { q: 'Worden de grijze lijntjes van Excel afgedrukt?', options: ['Ja', 'Nee'], correct: 1, explain: 'Die lijntjes zijn er alleen om te helpen. Echte lijnen maak je met randen.' }
        ] } },
      { id: 'l5', title: 'Slim: de opmaak kopiëren',
        body: `<p>Heb je één cel mooi opgemaakt en wil je een andere cel net zo? Klik op de mooie cel, klik op het <b>penseel</b> (Opmaak kopiëren/plakken) in de groep Klembord, en klik dan op de andere cel.</p>
<p>Dubbelklik je op het penseel, dan kan je de opmaak op veel cellen na elkaar plakken. Druk op <span class="key">Esc</span> om te stoppen.</p>`,
        widget: { type: 'read', confirm: 'Begrepen' } }
    ],
    startFile(seed, mode) {
      const d = m3data(seed, mode);
      const cells = { B2: d.tab.titel };
      ['B', 'C', 'D'].forEach((c, i) => { cells[c + '4'] = d.tab.kop[i]; d.tab.rijen.forEach((row, j) => { cells[c + (5 + j)] = row[i]; }); });
      return { name: mode === 'test' ? 'Test3_opmaak.xlsx' : 'Module3_opmaak.xlsx', spec: { sheetName: 'Opmaak', cells, colWidths: { B: 13, C: 13, D: 13 } } };
    },
    gen(seed, mode) {
      const d = m3data(seed, mode);
      const c = d.cell, w = d.woord;
      const kleurMsg = (hex, wat) => hex ? `De ${wat} is nu ${colorNames(hex)[0] || 'een andere kleur'}.` : `Er is nog geen ${wat} gekozen.`;
      return { startFile: true, tasks: [
        { id: 't1', text: `Typ het woord <b>${w}</b> in cel <span class="addr">${c}</span>.`,
          hints: ['Zoek eerst de kolom, dan de rij.', `Kijk in het naamvak linksboven of je echt in ${c} staat.`],
          check: ({ sheet }) => {
            if (norm(sheet.text(c)) === w) return ok();
            const el = sheet.findText(w);
            return no(el.length ? `"${w}" staat in ${el[0]}, niet in ${c}.` : `In ${c} staat nog niet "${w}".`);
          } },
        { id: 't2', text: `Maak <b>${w}</b> <b>vet</b>, <i>schuin</i> én <u>onderstreept</u>.`,
          hints: ['Drie knoppen naast elkaar in de groep Lettertype.', 'Klik op de cel en klik dan op B, I en U. Of gebruik Ctrl + B, Ctrl + I en Ctrl + U.'],
          check: ({ sheet }) => {
            const s = sheet.style(c), mis = [];
            if (!s.bold) mis.push('vet'); if (!s.italic) mis.push('schuin'); if (!s.underline) mis.push('onderstreept');
            return mis.length ? no(`Het woord is nog niet ${list(mis)}.`) : ok();
          } },
        { id: 't3', text: `Geef de letters van <b>${w}</b> de kleur <b>${d.fontKleur}</b>.`,
          hints: ['De letterkleur is de A met een gekleurde streep eronder.', 'Klik op het pijltje naast de A en kies de kleur.'],
          check: ({ sheet }) => {
            const s = sheet.style(c);
            return colorNames(s.fontColor).includes(d.fontKleur) ? ok() : no(kleurMsg(s.fontColor, 'letterkleur') + ` Kies ${d.fontKleur}.`);
          } },
        { id: 't4', text: `Kleur de cel <span class="addr">${c}</span> <b>${d.vulKleur}</b> (de achtergrond van de cel).`,
          hints: ['De achtergrond kleur je met de verfemmer.', 'Klik op het pijltje naast de verfemmer en kies de kleur.'],
          check: ({ sheet }) => {
            const s = sheet.style(c);
            return colorNames(s.fill).includes(d.vulKleur) ? ok() : no(kleurMsg(s.fill, 'opvulkleur') + ` Kies ${d.vulKleur}.`);
          } },
        { id: 't5', text: `Maak rij <b>${d.row}</b> ongeveer <b>twee keer zo hoog</b>.`,
          hints: ['De hoogte van een rij pas je aan bij de rijnummers, links.', `Sleep de lijn onder het rijnummer ${d.row} naar beneden.`],
          check: ({ sheet }) => {
            const h = sheet.rowHeight(d.row);
            if (h >= 24 && h <= 60) return ok();
            if (h > 60) return no(`Rij ${d.row} is nu veel te hoog. Ongeveer twee keer zo hoog is genoeg.`);
            return no(`Rij ${d.row} is nog niet hoger gemaakt.`);
          } },
        { id: 't6', text: `Zet <b>${w}</b> helemaal <b>in het midden</b> van de cel: van links naar rechts én van boven naar onder.`,
          hints: ['In de groep Uitlijning heb je twee rijen knoppen: één voor boven/midden/onder en één voor links/midden/rechts.', 'Klik op "Midden uitlijnen" (bovenste rij) en op "Centreren" (onderste rij).'],
          check: ({ sheet }) => {
            const s = sheet.style(c), mis = [];
            if (!['center', 'centerContinuous'].includes(s.h)) mis.push('van links naar rechts');
            if (s.v !== 'center') mis.push('van boven naar onder');
            return mis.length ? no(`Het woord staat nog niet in het midden ${list(mis)}.`) : ok();
          } },
        { id: 't7', text: `Voeg de cellen <span class="addr">B2:D2</span> samen en zet de titel in het midden.`,
          hints: ['Er is één knop die beide tegelijk doet.', 'Selecteer B2 tot D2 en klik op "Samenvoegen en centreren".'],
          check: ({ sheet }) => {
            const m = sheet.merges.find(x => x.toUpperCase() === 'B2:D2');
            if (!m) {
              const other = sheet.mergeAt('B2');
              return no(other ? `Je voegde ${other} samen. Het moet B2:D2 zijn.` : 'B2:D2 is nog niet samengevoegd.');
            }
            if (!['center', 'centerContinuous'].includes(sheet.style('B2').h)) return no('Samengevoegd, goed! Zet de titel nu ook in het midden.');
            return ok();
          } },
        { id: 't8', text: `Geef de titel lettergrootte <b>${d.size}</b>.`,
          hints: ['De lettergrootte staat in de groep Lettertype, naast de naam van het lettertype.', `Klik op de titel en kies ${d.size} in het lijstje met getallen.`],
          check: ({ sheet }) => {
            const s = sheet.style('B2');
            return s.size === d.size ? ok() : no(`De titel heeft nu lettergrootte ${s.size || 11}. Kies ${d.size}.`);
          } },
        { id: 't9', text: `Maak de kopjes <span class="addr">B4:D4</span> <b>vet</b> en kleur ze <b>${d.kopKleur}</b>.`,
          hints: ['Selecteer de drie cellen samen: sleep van B4 tot D4.', 'Klik daarna op B (vet) en kies de kleur bij de verfemmer.'],
          check: ({ sheet }) => {
            const notBold = [], notKleur = [];
            eachCell('B4:D4', a => { const s = sheet.style(a); if (!s.bold) notBold.push(a); if (!colorNames(s.fill).includes(d.kopKleur)) notKleur.push(a); });
            if (!notBold.length && !notKleur.length) return ok();
            if (notBold.length === 3 && notKleur.length === 3) return no(`De kopjes zijn nog niet vet en nog niet ${d.kopKleur}.`);
            const p = [];
            const wie = arr => arr.length === 3 ? 'De kopjes zijn' : `${list(arr)} ${arr.length > 1 ? 'zijn' : 'is'}`;
            if (notBold.length) p.push(`${wie(notBold)} nog niet vet`);
            if (notKleur.length) p.push(`${p.length ? wie(notKleur).toLowerCase() : wie(notKleur)} nog niet ${d.kopKleur}`);
            return no(p.join('; ') + '.');
          } },
        { id: 't10', text: `Zet <b>alle randen</b> rond de tabel <span class="addr">B4:D8</span>: een lijn rond elke cel.`,
          hints: ['Selecteer eerst de hele tabel, van B4 tot D8.', 'Klik op het pijltje naast de randen-knop en kies "Alle randen".'],
          check: ({ sheet }) => {
            const miss = [];
            eachCell('B4:D8', a => { if (!['left', 'right', 'top', 'bottom'].every(sd => sheet.hasEdge(a, sd))) miss.push(a); });
            if (!miss.length) return ok();
            let anyEdge = false;
            eachCell('B4:D8', a => { if (['left', 'right', 'top', 'bottom'].some(sd => sheet.hasEdge(a, sd))) anyEdge = true; });
            if (!anyEdge) return no('De tabel heeft nog geen randen.');
            if (miss.length === 15 || miss.includes('C6')) return no('Er staan al lijnen, maar niet rond elke cel. Kies "Alle randen", niet "Buitenrand".');
            return no(`Bijna! ${list(miss.slice(0, 4))}${miss.length > 4 ? ' en nog andere cellen' : ''} ${miss.length > 1 ? 'hebben' : 'heeft'} nog niet rondom een lijn. Kies "Alle randen", niet "Buitenrand".`);
          } }
      ] };
    }
  };

  const MODULES = [m1, m2, m3];

  /* Volgende modules (nog te bouwen) – getoond als vooruitblik. */
  const SOON = [
    { nr: 4, title: 'Excel vult zelf aan', short: 'Reeksen met de vulgreep: 1 tot 20, maanden, dagen, een 100-veld.' },
    { nr: 5, title: 'Getallen tonen', short: 'Procent, euro, decimalen en datums.' },
    { nr: 6, title: 'Je eerste formules', short: 'Rekenen met celadressen en formules doortrekken.' },
    { nr: 7, title: 'Functies', short: 'SOM, GEMIDDELDE, MEDIAAN, MIN, MAX en AFRONDEN.' },
    { nr: 8, title: 'Project: mijn maandrapport', short: 'Je bouwt je eigen rapport voor dagelijks werk.', project: true },
    { nr: 9, title: 'Excel kiest: ALS', short: 'Excel zegt zelf "juist" of "probeer opnieuw".' },
    { nr: 10, title: 'Keuzelijsten en kleuren die reageren', short: 'Gegevensvalidatie en voorwaardelijke opmaak.' },
    { nr: 11, title: 'Klaar voor een gebruiker', short: 'Score tellen, afbeeldingen, blad beveiligen.' },
    { nr: 12, title: 'Project: mijn interactief werkblad', short: 'Je maakt een werkblad voor je klas of voor een andere klas.', project: true }
  ];

  /* Alle opdrachten van een poging nakijken. */
  function runChecks(mod, gen, wb, ctx) {
    const sheet = wb.sheet();
    return gen.tasks.map(t => {
      let r;
      try { r = t.check(Object.assign({ wb, sheet }, ctx)); }
      catch (e) { r = no('Deze opdracht kon ik niet controleren. Vraag het aan je leerkracht.'); }
      return { id: t.id, ok: !!r.ok, msg: r.msg };
    });
  }

  const PASS = 0.8;
  const api = { MODULES, SOON, rng, newSeed, runChecks, PASS, KLEUREN };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.CvcModules = api;
})(typeof window !== 'undefined' ? window : globalThis);
