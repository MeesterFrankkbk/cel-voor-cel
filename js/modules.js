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
            if (colorNames(s.fontColor).includes(d.fontKleur)) return ok();
            if (colorNames(s.fill).includes(d.fontKleur)) return no(`Je kleurde de achtergrond ${d.fontKleur}. Hier moeten de letters zelf ${d.fontKleur} worden: gebruik de A met de kleurstreep.`);
            return no(kleurMsg(s.fontColor, 'letterkleur') + ` Kies ${d.fontKleur}.`);
          } },
        { id: 't4', text: `Kleur de cel <span class="addr">${c}</span> <b>${d.vulKleur}</b> (de achtergrond van de cel).`,
          hints: ['De achtergrond kleur je met de verfemmer.', 'Klik op het pijltje naast de verfemmer en kies de kleur.'],
          check: ({ sheet }) => {
            const s = sheet.style(c);
            if (colorNames(s.fill).includes(d.vulKleur)) return ok();
            if (!s.fill && colorNames(s.fontColor).includes(d.vulKleur)) return no(`Je kleurde de letters ${d.vulKleur}. Hier moet de achtergrond van de cel ${d.vulKleur} worden: gebruik de verfemmer.`);
            return no(kleurMsg(s.fill, 'opvulkleur') + ` Kies ${d.vulKleur}.`);
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
        { id: 't9', text: `Maak de kopjes <span class="addr">B4:D4</span> <b>vet</b> en geef de cellen een <b>${d.kopKleur} achtergrond</b> (opvulkleur).`,
          hints: ['Selecteer de drie cellen samen: sleep van B4 tot D4.', 'Klik daarna op B (vet). De achtergrond kleur je met de verfemmer, niet met de A.'],
          check: ({ sheet }) => {
            const notBold = [], notKleur = [];
            eachCell('B4:D4', a => { const s = sheet.style(a); if (!s.bold) notBold.push(a); if (!colorNames(s.fill).includes(d.kopKleur)) notKleur.push(a); });
            if (!notBold.length && !notKleur.length) return ok();
            const fontOnly = notKleur.length && notKleur.every(a => !sheet.style(a).fill && colorNames(sheet.style(a).fontColor).includes(d.kopKleur));
            if (fontOnly) return no(`Je gaf de letters de kleur ${d.kopKleur}. De opdracht vraagt de achtergrond van de cellen: gebruik de verfemmer (opvulkleur).${notBold.length ? ' Maak ze ook nog vet.' : ''}`);
            if (notBold.length === 3 && notKleur.length === 3) return no(`De kopjes zijn nog niet vet en hebben nog geen ${d.kopKleur} achtergrond.`);
            const p = [];
            const wie = arr => arr.length === 3 ? 'De kopjes zijn' : `${list(arr)} ${arr.length > 1 ? 'zijn' : 'is'}`;
            if (notBold.length) p.push(`${wie(notBold)} nog niet vet`);
            if (notKleur.length) p.push(`${p.length ? wie(notKleur).toLowerCase() : wie(notKleur)} nog niet ${d.kopKleur} gekleurd (achtergrond)`);
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

  // ======================================================================
  // Doorvullen zoals Excel (vulgreep) – gebruikt door de simulator en de controles
  // ======================================================================
  const MAANDEN = ['januari', 'februari', 'maart', 'april', 'mei', 'juni', 'juli', 'augustus', 'september', 'oktober', 'november', 'december'];
  const MAANDEN_KORT = ['jan', 'feb', 'mrt', 'apr', 'mei', 'jun', 'jul', 'aug', 'sep', 'okt', 'nov', 'dec'];
  const DAGEN = ['maandag', 'dinsdag', 'woensdag', 'donderdag', 'vrijdag', 'zaterdag', 'zondag'];
  const DAGEN_KORT = ['ma', 'di', 'wo', 'do', 'vr', 'za', 'zo'];
  const LIJSTEN = [MAANDEN, DAGEN, MAANDEN_KORT, DAGEN_KORT];
  const isNum = v => typeof v === 'number' && isFinite(v);
  const round9 = x => Math.round(x * 1e9) / 1e9;
  function likeCase(src, w) {
    const s = String(src);
    if (s.length > 1 && s === s.toUpperCase()) return w.toUpperCase();
    if (s[0] === s[0].toUpperCase() && s[0] !== s[0].toLowerCase()) return w[0].toUpperCase() + w.slice(1);
    return w;
  }
  /* Geeft 'count' nieuwe waarden na de startwaarden 'seeds', zoals Excel ze met de vulgreep maakt.
     ctrl = Ctrl ingedrukt (bij één getal: tellen in plaats van kopiëren). */
  function fillSeries(seeds, count, ctrl) {
    const n = seeds.length, out = [];
    if (!n || seeds.every(v => v == null || v === '')) return Array(count).fill(null);
    if (seeds.every(isNum)) {
      if (n === 1) { for (let i = 1; i <= count; i++) out.push(ctrl ? seeds[0] + i : seeds[0]); return out; }
      // lineaire trend (kleinste kwadraten), zoals Excel
      const xm = (n - 1) / 2, ym = seeds.reduce((a, b) => a + b, 0) / n;
      let sxy = 0, sxx = 0; seeds.forEach((y, x) => { sxy += (x - xm) * (y - ym); sxx += (x - xm) * (x - xm); });
      const slope = sxy / sxx, icpt = ym - slope * xm;
      for (let i = 0; i < count; i++) out.push(round9(icpt + slope * (n + i)));
      return out;
    }
    const lower = seeds.map(v => norm(v));
    const L = LIJSTEN.find(l => lower.every(v => l.includes(v)));
    if (L) {
      const idx = lower.map(v => L.indexOf(v));
      const step = n > 1 ? ((idx[n - 1] - idx[n - 2]) % L.length + L.length) % L.length || L.length : 1;
      let cur = idx[n - 1];
      for (let i = 0; i < count; i++) { cur = (cur + step) % L.length; out.push(likeCase(seeds[0], L[cur])); }
      return out;
    }
    const m = seeds.map(v => /^(.*?)(\d+)$/.exec(String(v)));
    if (m.every(Boolean) && m.every(x => x[1] === m[0][1])) {
      const nums = m.map(x => +x[2]);
      const step = n > 1 ? (nums[n - 1] - nums[0]) / (n - 1) : 1;
      if (Number.isInteger(step)) {
        for (let i = 1; i <= count; i++) out.push(m[0][1] + (nums[n - 1] + step * i));
        return out;
      }
    }
    for (let i = 0; i < count; i++) out.push(seeds[i % n]);
    return out;
  }

  /* Leest een getal uit een cel (ook als het als tekst met komma getypt is). */
  function numOf(v) {
    if (isNum(v)) return v;
    if (typeof v === 'string' && /^\s*-?\d+([.,]\d+)?\s*$/.test(v)) return parseFloat(v.replace(',', '.'));
    return null;
  }

  /* Controleert een reeks die in kolom 'col' begint in rij 'row0'. */
  function checkSeries(sheet, col, row0, expected) {
    const numeric = isNum(expected[0]);
    const show = v => v == null || v === '' ? 'niets' : (isNum(v) ? String(v).replace('.', ',') : `"${v}"`);
    const ENG = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'];
    const eq = (got, exp) => numeric ? numOf(got) === exp : norm(got) === norm(exp);
    const at = r => col + r;
    const got = expected.map((_, i) => sheet.value(at(row0 + i)));
    const filled = got.filter(v => v != null && String(v).trim() !== '');
    if (!filled.length) {
      for (let r = 1; r <= row0 + 6; r++) if (r !== row0 && eq(sheet.value(at(r)), expected[0]) && eq(sheet.value(at(r + 1)), expected[1])) {
        return no(`Je reeks begint in ${at(r)}. Ze moet in ${at(row0)} beginnen.`);
      }
      return no(`Kolom ${col} is nog leeg vanaf ${at(row0)}.`);
    }
    const bad = expected.findIndex((e, i) => !eq(got[i], e));
    if (bad >= 0) {
      if (!numeric) {
        const kort = expected.map(e => { const l = LIJSTEN.find(x => x.includes(norm(e))); return l === MAANDEN ? MAANDEN_KORT[MAANDEN.indexOf(norm(e))] : l === DAGEN ? DAGEN_KORT[DAGEN.indexOf(norm(e))] : null; });
        if (kort[bad] && norm(got[bad]) === kort[bad]) return no(`Schrijf de namen voluit: typ "${expected[0]}", niet "${kort[0]}".`);
        if (ENG.includes(norm(got[bad]))) return no(`In ${at(row0 + bad)} staat een Engels woord: "${got[bad]}". Typ de Nederlandse naam: "${expected[0]}".`);
      }
      if (!eq(got[0], expected[0])) {
        for (let r = 1; r <= row0 + 6; r++) if (r !== row0 && eq(sheet.value(at(r)), expected[0]) && eq(sheet.value(at(r + 1)), expected[1])) {
          return no(`Je reeks begint in ${at(r)}. Ze moet in ${at(row0)} beginnen.`);
        }
        return no(`In ${at(row0)} moet ${expected[0]} staan. Nu staat daar ${show(got[0])}.`);
      }
      const same = filled.length > 2 && filled.every(v => norm(v) === norm(filled[0]));
      if (same) {
        if (numeric) return no(`Excel kopieerde het getal ${filled[0]}. Typ de eerste twee getallen, selecteer ze allebei en sleep dan pas.`);
        return no(`In elke cel staat "${filled[0]}". Excel herkent het woord niet als lijst. Staat Excel in het Nederlands? Vraag het aan je leerkracht.`);
      }
      if (got[bad] == null || String(got[bad]).trim() === '') return no(`Je reeks stopt in ${at(row0 + bad - 1)}. Ga verder tot ${expected[expected.length - 1]} in ${at(row0 + expected.length - 1)}.`);
      return no(`In ${at(row0 + bad)} verwacht ik ${expected[bad]}, maar daar staat ${show(got[bad])}.`);
    }
    const after = sheet.value(at(row0 + expected.length));
    if (after != null && String(after).trim() !== '') return no(`Je sleepte te ver. Na ${expected[expected.length - 1]} moet je stoppen: ${at(row0 + expected.length)} moet leeg zijn.`);
    return ok();
  }

  // ======================================================================
  // MODULE 4 – Excel vult zelf aan
  // ======================================================================
  function m4data(seed, mode) {
    const r = rng(seed + '|m4|' + mode);
    const test = mode === 'test';
    return {
      N: r.int(30, 50), k: test ? r.int(6, 12) : r.int(3, 9), S: r.int(15, 25),
      ms: test ? r.int(1, 11) : 0, ds: test ? r.int(1, 6) : 0, base: test ? 100 : 0
    };
  }
  const range = (a, n, step) => Array.from({ length: n }, (_, i) => round9(a + i * (step == null ? 1 : step)));
  const cyc = (L, start, n) => Array.from({ length: n }, (_, i) => L[(start + i) % L.length]);

  const m4 = {
    id: 'm4', nr: 4,
    title: 'Excel vult zelf aan',
    short: 'Met de vulgreep laat je Excel reeksen maken: tellen, tafels, maanden, dagen en een honderdveld in een paar seconden.',
    skills: ['de vulgreep', 'reeksen met een stap', 'aftellen', 'maanden en dagen', 'kopie of reeks', 'honderdveld'],
    lessons: [
      { id: 'l1', title: 'De vulgreep',
        body: `<p>Klik op een cel. Rechtsonder in het groene kader zie je een klein <b>vierkantje</b>. Dat is de <b>vulgreep</b>.</p>
<p>Ga met de muis op de vulgreep staan: de muis wordt een dun <b>zwart plusje</b>. Hou de muisknop ingedrukt en sleep naar beneden of naar rechts. Excel vult de cellen voor je in.</p>
<p>Probeer het hieronder. Het werkt net zoals in Excel.</p>`,
        widget: { type: 'fillSim', rounds: [
          { rows: 8, cols: 4, cells: { A1: 'ja' }, goal: { A1: 'ja', A2: 'ja', A3: 'ja', A4: 'ja', A5: 'ja', A6: 'ja' }, ask: 'Klik op <span class="addr">A1</span> en sleep de vulgreep naar beneden tot <span class="addr">A6</span>.' },
          { rows: 11, cols: 4, cells: { A1: 1, A2: 2 }, goalSeries: { col: 'A', from: 1, values: range(1, 10) }, ask: 'Selecteer <span class="addr">A1</span> én <span class="addr">A2</span> samen. Sleep dan de vulgreep naar beneden tot <span class="addr">A10</span>.' },
          { rows: 10, cols: 4, cells: { B1: 5, B2: 10 }, goalSeries: { col: 'B', from: 1, values: range(5, 8, 5) }, ask: 'Maak 5, 10, 15 … tot <b>40</b> in kolom B.' }
        ] } },
      { id: 'l2', title: 'Excel zoekt het patroon',
        body: `<p>Excel kijkt naar de cellen die je selecteert en zoekt het <b>patroon</b>. Met twee getallen toon je de <b>stap</b>: van 5 naar 10 is +5, dus Excel gaat verder met 15, 20, 25 …</p>
<p>Een stap kan ook naar beneden gaan: 20, 18 … wordt 16, 14, 12.</p>`,
        widget: { type: 'quiz', questions: [
          { q: 'Je selecteert 5 en 10 en sleept naar beneden. Wat komt er na 10?', options: ['11', '15', '20'], correct: 1, explain: 'De stap is +5, dus 15.' },
          { q: 'Je selecteert 20 en 18 en sleept. Wat komt er na 18?', options: ['16', '17', '19'], correct: 0, explain: 'De stap is −2, dus 16.' },
          { q: 'Je typt "Les 1" in een cel en sleept naar beneden. Wat krijg je?', options: ['Les 1, Les 1, Les 1 …', 'Les 2, Les 3, Les 4 …', 'Les 11, Les 111 …'], correct: 1, explain: 'Staat er een getal achter een woord, dan telt Excel verder.' },
          { q: 'Je typt alleen 7 in één cel en sleept naar beneden. Wat krijg je?', options: ['7, 8, 9 …', '7, 14, 21 …', '7, 7, 7 …'], correct: 2, explain: 'Met één getal ziet Excel geen stap. Het kopieert dan gewoon.' }
        ] } },
      { id: 'l3', title: 'Eén getal: kopie of reeks?',
        body: `<p>Sleep je met maar één getal, dan <b>kopieert</b> Excel het. Wil je toch tellen? Dan heb je twee manieren:</p>
<p>1. Hou <span class="key">Ctrl</span> ingedrukt terwijl je sleept.</p>
<p>2. Of sleep gewoon. Na het slepen verschijnt rechtsonder een klein knopje: <b>Opvulopties</b>. Klik erop en kies <b>Reeks doorvoeren</b>.</p>`,
        widget: { type: 'fillSim', rounds: [
          { rows: 10, cols: 4, cells: { A1: 1 }, goalSeries: { col: 'A', from: 1, values: range(1, 8) }, options: true, ask: 'Maak 1 tot <b>8</b> in kolom A met maar één getal. Gebruik <span class="key">Ctrl</span> of het knopje <b>Opvulopties</b>.' },
          { rows: 10, cols: 4, cells: { C1: 100 }, goalSeries: { col: 'C', from: 1, values: range(100, 6) }, options: true, ask: 'Maak 100, 101, 102 … tot <b>105</b> in kolom C.' }
        ] } },
      { id: 'l4', title: 'Lijsten die Excel kent',
        body: `<p>Excel kent de <b>maanden</b> en de <b>dagen van de week</b>, ook afgekort (jan, feb … en ma, di …). Eén naam typen is genoeg: sleep en Excel gaat verder. Na december komt weer januari.</p>
<p>Belangrijk: dit werkt alleen als <b>Excel in het Nederlands</b> staat. Krijg je overal hetzelfde woord? Vraag het dan aan je leerkracht.</p>`,
        widget: { type: 'fillSim', rounds: [
          { rows: 13, cols: 4, cells: { A1: 'januari' }, goalSeries: { col: 'A', from: 1, values: MAANDEN }, ask: 'Maak alle maanden in kolom A: van januari tot <b>december</b>.' },
          { rows: 13, cols: 4, cells: { B1: 'maandag' }, goalSeries: { col: 'B', from: 1, values: DAGEN }, ask: 'Maak de dagen van de week in kolom B: van maandag tot <b>zondag</b>.' },
          { rows: 13, cols: 4, cells: { C1: 'Ronde 1' }, goalSeries: { col: 'C', from: 1, values: range(1, 6).map(i => 'Ronde ' + i) }, ask: 'Maak Ronde 1 tot <b>Ronde 6</b> in kolom C.' }
        ] } },
      { id: 'l5', title: 'Een honderdveld in een paar seconden',
        body: `<p>Een honderdveld met de hand typen duurt lang. Met een slim plan doet Excel het voor jou:</p>
<p><b>1.</b> In kolom A: typ 1 en 11 onder elkaar en sleep naar beneden. Zo krijg je 1, 11, 21 … 91.</p>
<p><b>2.</b> In kolom B: hetzelfde met 2 en 12.</p>
<p><b>3.</b> Selecteer kolom A en B samen en sleep naar <b>rechts</b>. Elke rij ziet zijn eigen stap (+1) en telt verder.</p>`,
        widget: { type: 'fillSim', rounds: [
          { rows: 10, cols: 10, cells: { A1: 1, A2: 11 }, goalSeries: { col: 'A', from: 1, values: range(1, 10, 10) }, ask: 'Stap 1: maak 1, 11, 21 … tot <b>91</b> in kolom A.' },
          { rows: 10, cols: 10, cells: Object.assign({ B1: 2, B2: 12 }, ...range(1, 10, 10).map((v, i) => ({ ['A' + (i + 1)]: v }))), goalSeries: { col: 'B', from: 1, values: range(2, 10, 10) }, keep: true, ask: 'Stap 2: maak 2, 12, 22 … tot <b>92</b> in kolom B.' },
          { rows: 10, cols: 10, cells: Object.assign({}, ...range(1, 10, 10).map((v, i) => ({ ['A' + (i + 1)]: v, ['B' + (i + 1)]: v + 1 }))), goalBlock: { base: 0 }, ask: 'Stap 3: selecteer <span class="addr">A1:B10</span> en sleep de vulgreep naar rechts tot kolom <b>J</b>.' }
        ] } }
    ],
    startFile(seed, mode) {
      const d = m4data(seed, mode);
      const b = { bold: true };
      const cells = { A1: { v: 'Tellen', ...b }, C1: { v: 'Tafel van ' + d.k, ...b }, E1: { v: 'Aftellen', ...b }, G1: { v: 'Maanden', ...b }, I1: { v: 'Dagen', ...b }, K1: { v: 'Honderdveld', ...b } };
      return { name: mode === 'test' ? 'Test4_reeksen.xlsx' : 'Module4_reeksen.xlsx', spec: { sheetName: 'Reeksen', cells, colWidths: { C: 12, G: 12, I: 12 } } };
    },
    gen(seed, mode) {
      const d = m4data(seed, mode);
      const months = cyc(MAANDEN, d.ms, 12), days = cyc(DAGEN, d.ds, 14), b = d.base;
      return { startFile: true, tasks: [
        { id: 't1', text: `Zet in kolom A, vanaf <span class="addr">A2</span>, de getallen <b>1 tot ${d.N}</b> onder elkaar.`,
          hints: ['Typ de eerste twee getallen, selecteer ze allebei en sleep de vulgreep naar beneden.', `Tijdens het slepen toont Excel in een klein kadertje tot welk getal je komt. Laat los bij ${d.N}.`],
          check: ({ sheet }) => checkSeries(sheet, 'A', 2, range(1, d.N)) },
        { id: 't2', text: `Zet in kolom C, vanaf <span class="addr">C2</span>, de <b>tafel van ${d.k}</b>: ${d.k}, ${2 * d.k}, ${3 * d.k} … tot ${10 * d.k}.`,
          hints: ['Welke twee getallen moet Excel zien om de stap te begrijpen?', `Typ ${d.k} in C2 en ${2 * d.k} in C3. Selecteer ze samen en sleep naar beneden tot ${10 * d.k}.`],
          check: ({ sheet }) => checkSeries(sheet, 'C', 2, range(d.k, 10, d.k)) },
        { id: 't3', text: `Tel in kolom E, vanaf <span class="addr">E2</span>, <b>af van ${d.S} tot 1</b>.`,
          hints: ['Excel kan ook naar beneden tellen. Toon de stap met twee getallen.', `Typ ${d.S} en ${d.S - 1}, selecteer ze samen en sleep tot je bij 1 bent.`],
          check: ({ sheet }) => checkSeries(sheet, 'E', 2, range(d.S, d.S, -1)) },
        { id: 't4', text: mode === 'test' ? `Zet in kolom G, vanaf <span class="addr">G2</span>, <b>twaalf maanden</b> na elkaar, te beginnen met <b>${months[0]}</b>.` : `Zet in kolom G, vanaf <span class="addr">G2</span>, de <b>twaalf maanden</b>: januari tot december.`,
          hints: ['Excel kent de maanden. Eén maand typen is genoeg.', `Typ ${months[0]} in G2 en sleep de vulgreep 11 cellen naar beneden, tot ${months[11]}.`],
          check: ({ sheet }) => checkSeries(sheet, 'G', 2, months) },
        { id: 't5', text: `Zet in kolom I, vanaf <span class="addr">I2</span>, de dagen van <b>twee weken</b> (14 dagen), te beginnen met <b>${days[0]}</b>.`,
          hints: ['Na zondag begint Excel vanzelf weer met maandag.', `Typ ${days[0]} in I2 en sleep tot I15.`],
          check: ({ sheet }) => checkSeries(sheet, 'I', 2, days) },
        { id: 't6', text: `Maak in <span class="addr">K2:T11</span> een <b>honderdveld van ${b + 1} tot ${b + 100}</b>: ${b + 1} tot ${b + 10} naast elkaar in rij 2, ${b + 11} tot ${b + 20} in rij 3, enzovoort.`,
          hints: [`Begin met de eerste twee kolommen. Welk getal staat onder ${b + 1}? En onder ${b + 2}?`, `Typ ${b + 1} en ${b + 11} in K2 en K3 en sleep naar beneden tot K11. Doe hetzelfde met ${b + 2} en ${b + 12} in kolom L. Selecteer dan K2:L11 samen en sleep naar rechts tot kolom T.`],
          check: ({ sheet }) => {
            const val = (c, r) => numOf(sheet.value(addr(11 + c, 2 + r)));
            let wrong = null, empty = 0, turned = true;
            for (let r = 0; r < 10; r++) for (let c = 0; c < 10; c++) {
              const v = val(c, r), e = b + r * 10 + c + 1;
              if (v == null) empty++;
              if (v !== b + c * 10 + r + 1) turned = false;
              if (v !== e && !wrong) wrong = { a: addr(11 + c, 2 + r), e, v };
            }
            if (!wrong) {
              const extra = [addr(21, 2), addr(11, 12)].filter(a => !sheet.isEmpty(a));
              return extra.length ? no(`Je honderdveld klopt, maar je ging te ver: ${list(extra)} moet leeg zijn.`) : ok();
            }
            if (empty === 100) return no('Het honderdveld is nog leeg. Begin in K2.');
            if (turned) return no(`Je honderdveld staat gekanteld: ${b + 1} tot ${b + 10} moet naast elkaar in rij 2 staan, niet onder elkaar.`);
            if (empty > 0 && wrong.v == null) return no(`Je honderdveld is nog niet af. In ${wrong.a} moet ${wrong.e} komen.`);
            return no(`In ${wrong.a} verwacht ik ${wrong.e}, maar daar staat ${wrong.v == null ? 'niets' : wrong.v}.`);
          } }
      ] };
    }
  };

  // ======================================================================
  // MODULE 5 – Getallen tonen
  // ======================================================================
  const nlNum = (v, d) => (d == null ? String(+(+v).toPrecision(10)) : (+v).toFixed(d)).replace('.', ',');
  const pctStr = (v, d) => nlNum(v * 100, d) + '%';
  const euroStr = (v, d) => '€ ' + nlNum(v, d);
  const MAAND_LANG = MAANDEN;
  const close = (a, b) => isNum(a) && Math.abs(a - b) < 1e-9;
  /* een breuk kiezen waarvan het procent niet op een mooi getal of op ,5 uitkomt (zo zie je het afronden) */
  function pickFraction(r, dec) {
    for (let i = 0; i < 200; i++) {
      const max = r.pick([15, 30, 35, 45, 60, 70]), score = r.int(Math.ceil(max * 0.5), max - 1);
      const x = score / max * 100 * Math.pow(10, dec), frac = x - Math.floor(x);
      if (frac > 0.08 && frac < 0.92 && Math.abs(frac - 0.5) > 0.08) return { score, max, v: score / max };
    }
    return { score: 23, max: 30, v: 23 / 30 };
  }
  function m5data(seed, mode) {
    const r = rng(seed + '|m5|' + mode);
    const test = mode === 'test';
    const p1 = pickFraction(r, 0), p2 = pickFraction(r, 1);
    const day = r.int(1, 28), month = r.int(1, 12), year = 2027;
    return {
      test, p1, p2,
      prijs1: r.pick([2.5, 1.75, 3.2, 0.9, 4.6]), prijs2: r.pick([12, 8, 15, 20, 6]),
      lang: r.pick([3.14159, 2.71828, 1.41421, 9.87654, 6.28318]), langDec: test ? 1 : 2,
      date: { y: year, m: month, d: day }, lengte: r.pick([1.38, 1.42, 1.45, 1.51, 1.56, 1.62]),
      labels: test
        ? ['Toets Frans', 'Toets WO', 'Prijs pizza', 'Prijs schrift', 'Getal met veel decimalen', 'Datum van het schoolfeest', 'Lengte van mijn broer (m)']
        : ['Toets wiskunde', 'Toets taal', 'Prijs ijsje', 'Prijs boek', 'Getal met veel decimalen', 'Datum van de uitstap', 'Mijn lengte (m)']
    };
  }
  function fmtDiag(sheet, a, want, wantDec, original, origLabel) {
    const orig = origLabel || nlNum(original);
    const info = sheet.numFmt(a), v = sheet.value(a);
    const names = { percent: 'procent', currency: 'euro', number: 'getal met vaste decimalen', date: 'datum', general: 'standaard', text: 'tekst' };
    if (original != null && !close(v, original)) {
      if (isNum(v) && Math.abs(v - original) < 0.5 * Math.pow(10, -(wantDec || 0)) * (want === 'percent' ? 0.01 : 1) + 1e-9) return no(`Je typte het getal in ${a} opnieuw. Dat mag niet: het getal zelf moet ${orig} blijven. Verander alleen hoe het eruitziet.`);
      if (v == null) return no(`Cel ${a} is leeg. Er moet ${orig} in staan.`);
      return no(`In ${a} staat niet meer het oorspronkelijke getal (${orig}). Klik op Ctrl + Z of download het startbestand opnieuw.`);
    }
    if (info.kind !== want) return no(`${a} is nog ${info.kind === 'general' ? 'niet opgemaakt' : 'opgemaakt als ' + names[info.kind]}. Het moet ${names[want]} worden.`);
    if (wantDec != null && info.dec !== wantDec) return no(`${a} is ${names[want]}, goed! Maar met ${info.dec} ${info.dec === 1 ? 'decimaal' : 'decimalen'}. Het moet ${wantDec} ${wantDec === 1 ? 'decimaal' : 'decimalen'} zijn: klik op ${info.dec > wantDec ? 'Minder decimalen' : 'Meer decimalen'}.`);
    return ok();
  }

  const m5 = {
    id: 'm5', nr: 5,
    title: 'Getallen tonen',
    short: 'Excel bewaart een getal en kiest apart hoe het eruitziet. Je leert procent, euro, decimalen, datums en de komma.',
    skills: ['waarde en weergave', 'procent', 'euro', 'meer en minder decimalen', 'datum typen', 'komma, geen punt'],
    lessons: [
      { id: 'l1', title: 'Wat Excel bewaart en wat je ziet',
        body: `<p>In een cel kan <b>0,675</b> staan, terwijl je <b>68%</b> ziet. Het getal blijft hetzelfde. Alleen de <b>weergave</b> verandert.</p>
<p>Wat er echt in de cel zit, zie je bovenaan in de <b>formulebalk</b>. De knoppen om de weergave te kiezen, staan op de tab Start in de groep <b>Getal</b>: <b>%</b> (procent), <b>€</b> (valuta), <b>Meer decimalen</b> en <b>Minder decimalen</b>.</p>
<p>Maak de cel hieronder hetzelfde als het voorbeeld. Kijk telkens ook naar de formulebalk.</p>`,
        widget: { type: 'numSim', rounds: [
          { v: 0.675, target: '68%' }, { v: 2.5, target: '€ 2,50' }, { v: 3.14159, target: '3,14' }, { v: 0.8333, target: '83,3%' }
        ] } },
      { id: 'l2', title: 'Komma of punt?',
        body: `<p>In België schrijf je een kommagetal met een <b>komma</b>: twee en een half is <b>2,5</b>. Excel op een Belgische computer verwacht dat ook.</p>
<p>Typ je <b>2.5</b> met een punt, dan begrijpt Excel het niet als getal. Het wordt <b>tekst</b>, of soms zelfs een <b>datum</b>.</p>
<p>Zo zie je het verschil: een <b>getal</b> staat rechts in de cel, <b>tekst</b> staat links.</p>`,
        widget: { type: 'quiz', questions: [
          { q: 'Je wil twee en een half typen in Excel. Wat typ je?', options: ['2,5', '2.5', '2½'], correct: 0, explain: 'In België gebruik je een komma.' },
          { q: 'Je typt 1.45 in een cel. Wat gebeurt er?', options: ['Excel maakt er 1,45 van', 'Excel ziet het als tekst of maakt er een datum van', 'Excel maakt er 145 van'], correct: 1, explain: 'Met een punt herkent Excel het niet als getal.' },
          { q: 'Hoe zie je dat Excel iets als getal begrijpt?', options: ['Getallen staan rechts in de cel, tekst links', 'Getallen zijn vet', 'Getallen zijn blauw'], correct: 0, explain: 'Staat je getal links, dan is het tekst. Rekenen lukt dan niet.' }
        ] } },
      { id: 'l3', title: 'Procent',
        body: `<p>Procent betekent "op honderd". <b>1</b> is 100%, <b>0,5</b> is 50% en <b>0,75</b> is 75%.</p>
<p>Klik je op de knop <b>%</b>, dan vermenigvuldigt Excel het getal niet echt. Het <b>toont</b> 0,75 gewoon als 75%.</p>
<p>Dat is handig voor je rapport: 27 op 40 is 27 gedeeld door 40 = 0,675. Met de knop % zie je 68%.</p>`,
        widget: { type: 'quiz', questions: [
          { q: 'In een cel staat 0,25. Je klikt op %. Wat zie je?', options: ['0,25%', '25%', '2,5%'], correct: 1, explain: '0,25 is 25 op 100.' },
          { q: 'Een cel toont 68%. Wat kan er echt in de cel staan?', options: ['68', '0,675', '6,8'], correct: 1, explain: '0,675 is 67,5%. Zonder decimalen rondt Excel af naar 68%. In de formulebalk zie je het echte getal.' },
          { q: 'Je typt 75% in een cel. Wat bewaart Excel?', options: ['0,75', '75', '7,5'], correct: 0, explain: 'Excel bewaart 0,75 en toont 75%.' }
        ] } },
      { id: 'l4', title: 'Datums',
        body: `<p>Een datum typ je met schuine strepen: <b>15/3/2027</b> is 15 maart 2027. Excel herkent de datum en zet hem <b>rechts</b> in de cel.</p>
<p>Weetje: Excel telt de <b>dagen</b>. Dag 1 is 1 januari 1900. Voor Excel is 1 januari 2027 dag <b>46388</b>. Daarom kan Excel ook rekenen met datums, bv. hoeveel dagen het nog is tot de vakantie.</p>`,
        widget: { type: 'quiz', questions: [
          { q: 'Hoe typ je 15 maart 2027?', options: ['15/3/2027', '15 maart twintig zevenentwintig', '2027.3.15'], correct: 0, explain: 'Dag, maand en jaar met schuine strepen.' },
          { q: 'Je typt een datum en hij staat links in de cel. Wat betekent dat?', options: ['Excel ziet het als tekst, niet als datum', 'Het is een extra mooie datum', 'Excel heeft hem afgerond'], correct: 0, explain: 'Een echte datum is een getal en staat rechts.' },
          { q: 'Je zet de datum 2/1/1900 om naar Standaard. Welk getal zie je?', options: ['2', '1900', '21'], correct: 0, explain: '2 januari 1900 is dag 2 voor Excel.' }
        ] } }
    ],
    startFile(seed, mode) {
      const d = m5data(seed, mode), L = d.labels;
      const cells = { A1: { v: 'Getallen tonen', bold: true, size: 14 },
        A3: L[0], B3: d.p1.v, A4: L[1], B4: d.p2.v, A6: L[2], B6: d.prijs1, A7: L[3], B7: d.prijs2,
        A9: L[4], B9: d.lang, A11: L[5], A12: L[6] };
      return { name: mode === 'test' ? 'Test5_getallen.xlsx' : 'Module5_getallen.xlsx', spec: { sheetName: 'Getallen', cells, colWidths: { A: 28, B: 16 } } };
    },
    gen(seed, mode) {
      const d = m5data(seed, mode), L = d.labels;
      const dt = d.date, serial = X.excelSerial(dt.y, dt.m, dt.d);
      const dateTxt = `${dt.d} ${MAAND_LANG[dt.m - 1]} ${dt.y}`, dateTyped = `${dt.d}/${dt.m}/${dt.y}`;
      return { startFile: true, tasks: [
        { id: 't1', text: `${L[0]} haalde ${d.p1.score} op ${d.p1.max}. In <span class="addr">B3</span> staat de uitkomst van ${d.p1.score} gedeeld door ${d.p1.max}. Toon het als <b>procent zonder decimalen</b>: je ziet dan <b>${pctStr(d.p1.v, 0)}</b>.`,
          hints: ['Er is een knop die een getal als procent toont.', 'Klik op B3 en dan op de knop % (tab Start, groep Getal).'],
          check: ({ sheet }) => fmtDiag(sheet, 'B3', 'percent', 0, d.p1.v, `de uitkomst van ${d.p1.score} gedeeld door ${d.p1.max}`) },
        { id: 't2', text: `${L[1]}: ${d.p2.score} op ${d.p2.max}. Toon <span class="addr">B4</span> als <b>procent met 1 decimaal</b>: je ziet dan <b>${pctStr(d.p2.v, 1)}</b>.`,
          hints: ['Eerst procent, dan één cijfer na de komma erbij.', 'Klik op %, daarna één keer op "Meer decimalen" (de knop met ←,0).'],
          check: ({ sheet }) => fmtDiag(sheet, 'B4', 'percent', 1, d.p2.v, `de uitkomst van ${d.p2.score} gedeeld door ${d.p2.max}`) },
        { id: 't3', text: `Toon de twee prijzen in <span class="addr">B6:B7</span> in <b>euro met 2 decimalen</b>: ${euroStr(d.prijs1, 2)} en ${euroStr(d.prijs2, 2)}.`,
          hints: ['Selecteer de twee cellen samen, dan hoef je maar één keer te klikken.', 'Sleep over B6 en B7 en klik op de knop met het geldbiljet (€). Staat er een dollar? Klik op het pijltje ernaast en kies € Euro.'],
          check: ({ sheet }) => { const a = fmtDiag(sheet, 'B6', 'currency', 2, d.prijs1); if (!a.ok) return a; return fmtDiag(sheet, 'B7', 'currency', 2, d.prijs2); } },
        { id: 't4', text: `In <span class="addr">B9</span> staat een getal met veel decimalen. Toon het met <b>${d.langDec} ${d.langDec === 1 ? 'decimaal' : 'decimalen'}</b> (${nlNum(d.lang, d.langDec)}), <b>zonder</b> het opnieuw te typen.`,
          hints: ['Excel kan minder cijfers tonen zonder het getal te veranderen.', 'Klik op B9 en klik een paar keer op "Minder decimalen" (de knop met ,00→,0).'],
          check: ({ sheet }) => fmtDiag(sheet, 'B9', 'number', d.langDec, d.lang) },
        { id: 't5', text: `Typ in <span class="addr">B11</span> de datum <b>${dateTxt}</b>.`,
          hints: ['Dag, maand en jaar met schuine strepen.', `Typ ${dateTyped} en druk op Enter. Staat het rechts in de cel? Dan is het goed.`],
          check: ({ sheet }) => {
            const v = sheet.value('B11'), info = sheet.numFmt('B11');
            if (v == null || String(v).trim() === '') return no('B11 is nog leeg.');
            if (typeof v === 'string') return no(`Excel ziet "${v}" als tekst, niet als datum. Typ het zo: ${dateTyped}`);
            if (info.kind !== 'date') return no(`In B11 staat het getal ${nlNum(v)}, geen datum. Typ de datum zo: ${dateTyped}`);
            if (Math.round(v) !== serial) { const g = X.serialToDate(v); if (g.d === dt.m && g.m === dt.d && g.y === dt.y) return no(`Dag en maand zijn omgewisseld: Excel leest ${g.d} ${MAAND_LANG[g.m - 1]}. Typ eerst de dag, dan de maand: ${dateTyped}. Gebeurt het opnieuw? Dan staat Excel misschien in het Engels. Vraag het aan je leerkracht.`); return no(`In B11 staat de datum ${g.d} ${MAAND_LANG[g.m - 1]} ${g.y}. Het moet ${dateTxt} zijn.`); }
            return ok();
          } },
        { id: 't6', text: `Typ in <span class="addr">B12</span> de lengte <b>${nlNum(d.lengte)} meter</b> als kommagetal.`,
          hints: ['Op een Belgische computer gebruik je een komma, geen punt.', `Typ ${nlNum(d.lengte)} en druk op Enter. Het getal moet rechts in de cel staan.`],
          check: ({ sheet }) => {
            const v = sheet.value('B12'), info = sheet.numFmt('B12');
            if (v == null || String(v).trim() === '') return no('B12 is nog leeg.');
            if (typeof v === 'string') return /\./.test(v) ? no(`Je typte een punt: "${v}". Excel ziet dat als tekst. Typ ${nlNum(d.lengte)} met een komma.`) : no(`Excel ziet "${v}" als tekst. Typ alleen het getal: ${nlNum(d.lengte)}`);
            if (info.kind === 'date') return no(`Excel maakte er een datum van. Maak de cel leeg (Delete), zet de notatie op Standaard en typ ${nlNum(d.lengte)} met een komma.`);
            if (close(v, d.lengte * 100)) return no(`Je typte ${nlNum(v)}: dat zijn centimeters. Typ de lengte in meter: ${nlNum(d.lengte)}`);
            if (!close(v, d.lengte)) return no(`In B12 staat ${nlNum(v)}. Het moet ${nlNum(d.lengte)} zijn.`);
            return ok();
          } }
      ] };
    }
  };

  // ======================================================================
  // Formules controleren (stap 6 en verder)
  // ======================================================================
  const FX = (typeof module !== 'undefined' && module.exports) ? require('./formula.js') : root.CvcFormula;
  const showV = v => FX.display(v, 'nl');
  const dutch = f => FX.toDutch(f);
  /* Controleert een formulecel. opt: { expect, refs: [...], label, funcs: [...] } */
  function checkFormula(sheet, a, opt) {
    const c = sheet.cell(a), f = sheet.formula(a);
    const raw = c ? c.v : null;
    if (!f) {
      if (raw == null || String(raw).trim() === '') return no(`Cel ${a} is nog leeg.`);
      if (typeof raw === 'string' && /^\s*=/.test(raw)) return no(`Excel ziet je formule in ${a} als tekst. Maak de cel leeg, zet de notatie op Standaard en typ de formule opnieuw.`);
      if (typeof raw === 'string' && /[A-Za-z]+\d+\s*[-+*/x×:]\s*[A-Za-z]+\d+/.test(raw)) return no(`In ${a} staat "${raw}" als gewone tekst: je vergat het =-teken. Een formule begint altijd met =.`);
      if (isNum(raw) && opt.expect != null && Math.abs(raw - opt.expect) < 1e-6) return no(`In ${a} staat het getal ${showV(raw)}. Dat klopt, maar het is geen formule. Laat Excel rekenen: typ een formule die met = begint.`);
      return no(`In ${a} staat ${showV(raw)}, geen formule. Een formule begint met =.`);
    }
    const inf = FX.info(f, 'en'), v = sheet.calc(a);
    if (/[A-Za-z]+\$?\d+\s*[xX×]\s*\$?[A-Za-z]/.test(f)) return no('Vermenigvuldigen doe je met * (sterretje), niet met x.');
    if (FX.isErr(v)) {
      const bron = [...new Set(inf.cells)].find(x => x !== a && FX.isErr(sheet.calc(x)));
      if (bron) return no(`${a} toont een foutmelding omdat ${bron} zelf een fout bevat. Los eerst ${bron} op.`);
      if (v.error === '#DIV/0!') return no(`${a} toont #DEEL/0!: je formule deelt door 0 of door een lege cel. Kijk welke cel je achter de / zette.`);
      if (v.error === '#NAME?') {
        const onbekend = inf.funcs.filter(x => !FX.EN_NL[x]);
        return no(onbekend.length ? `${a} toont #NAAM?: Excel kent de functie ${onbekend[0]} niet. Controleer de schrijfwijze.` : `${a} toont #NAAM?: Excel begrijpt een woord in je formule niet. Controleer de schrijfwijze.`);
      }
      return no(`${a} toont een foutmelding (${showV(v)}). Kijk je formule na: ${dutch(f)}`);
    }
    if (!inf.hasRef) return no(`Je formule (${dutch(f)}) rekent met getallen die je zelf typte. Gebruik de celadressen, dan rekent Excel vanzelf opnieuw als er iets verandert.`);
    if (opt.funcs && !opt.funcs.some(fn => inf.funcs.includes(fn))) return opt.funcMsg ? no(opt.funcMsg) : no(`Gebruik de functie ${opt.funcs.map(x => FX.EN_NL[x] || x).join(' of ')}.`);
    if (opt.refs) {
      const missing = opt.refs.filter(r => !inf.cells.includes(r));
      if (missing.length) {
        const used = [...new Set(inf.cells)].filter(x => !opt.refs.includes(x));
        if (opt.allRefsMsg) return no(opt.allRefsMsg(missing, used));
        return no(`Je formule (${dutch(f)}) gebruikt ${used.length ? list(used) : 'andere cellen'}${missing.length ? `, maar niet ${list(missing)}` : ''}.`);
      }
    }
    if (opt.expect != null && !(isNum(v) && Math.abs(v - opt.expect) < 1e-6)) {
      if (opt.wrongMsg) { const m = opt.wrongMsg(v, f); if (m) return no(m); }
      return no(`Je formule geeft ${showV(v)}, maar het moet ${showV(opt.expect)} zijn. Kijk je formule na: ${dutch(f)}`);
    }
    return ok();
  }

  // ======================================================================
  // MODULE 6 – Je eerste formules
  // ======================================================================
  const ARTIKELS = [['Chips', 0.85], ['Frisdrank', 1.29], ['Water', 0.65], ['Koekjes', 1.49], ['Bekertjes', 1.15], ['Servetten', 0.99], ['Snoep', 2.35], ['Fruitsap', 1.79], ['Popcorn', 1.25], ['Slingers', 3.5], ['Ballonnen', 2.75], ['Pistolets', 0.45]];
  function m6data(seed, mode) {
    const r = rng(seed + '|m6|' + mode);
    const items = r.shuffle(ARTIKELS).slice(0, 6).map(([n, p]) => ({ n, p, a: r.int(2, 12) }));
    const total = round9(items.reduce((s, x) => s + x.p * x.a, 0));
    const budget = Math.ceil(total / 5) * 5 + 5 * r.int(1, 3);
    const kids = r.int(18, 26);
    const kids2 = kids + r.pick([-3, -2, 2, 3, 4]);
    return { items, total, budget, kids, kids2, titel: mode === 'test' ? 'Inkopen voor de sportdag' : 'Inkopen voor het klasfeest' };
  }

  const m6 = {
    id: 'm6', nr: 6,
    title: 'Je eerste formules',
    short: 'Je laat Excel rekenen met formules en celadressen. Verander je een getal, dan rekent Excel vanzelf opnieuw.',
    skills: ['het =-teken', '+ − * /', 'celadressen in formules', 'formules doortrekken', 'voorrang en haakjes', 'foutmeldingen'],
    lessons: [
      { id: 'l1', title: 'Excel als rekenmachine',
        body: `<p>Een <b>formule</b> begint altijd met een <b>=</b>. Daarna schrijf je de berekening. Excel toont in de cel de <b>uitkomst</b>; de formule zelf zie je in de formulebalk.</p>
<p>Excel rekent met deze tekens: <span class="key">+</span> optellen, <span class="key">-</span> aftrekken, <span class="key">*</span> vermenigvuldigen (geen ×) en <span class="key">/</span> delen (geen :).</p>`,
        widget: { type: 'sheetSim', rows: 4, cols: 3, rounds: [
          { select: 'A1', ask: 'Klik op <span class="addr">A1</span> en typ <b>=6*7</b>. Druk op Enter.',
            check: (s, ch) => ch !== 'A1' ? null : s.isFormula('A1') ? (s.num('A1') === 42 ? ok('Juist: 42. Excel deed de som voor jou.') : no(`Je formule geeft ${s.show('A1')}. Typ precies =6*7`)) : no('Je typte geen formule. Een formule begint met =.') },
          { select: 'A2', ask: 'Bereken in <span class="addr">A2</span>: 125 gedeeld door 5.',
            check: (s, ch) => ch !== 'A2' ? null : /:/.test(String(s.raw('A2'))) ? no('Delen doe je met / (schuine streep), niet met :.') : !s.isFormula('A2') ? no('Begin met =.') : s.num('A2') === 25 ? ok('Juist: 25.') : no(`Je formule geeft ${s.show('A2')}. Het moet 25 zijn.`) },
          { select: 'A3', ask: 'Typ in <span class="addr">A3</span> eens <b>6*7</b>, maar <b>zonder</b> =. Wat zie je?',
            check: (s, ch) => ch !== 'A3' ? null : s.isFormula('A3') ? no('Typ het nu eens zonder =.') : String(s.raw('A3')).replace(/\s/g, '') === '6*7' ? ok('Zie je? Zonder = is het gewone tekst. Excel rekent niet.') : no('Typ 6*7 zonder =.') }
        ] } },
      { id: 'l2', title: 'Rekenen met celadressen',
        body: `<p>Je kan in een formule ook <b>celadressen</b> gebruiken: <b>=B1*B2</b> betekent "wat in B1 staat, maal wat in B2 staat".</p>
<p>Dat is slim: verandert er een getal in B1 of B2, dan rekent Excel de uitkomst <b>vanzelf opnieuw</b>.</p>
<p>Tip: je hoeft het adres niet te typen. Typ <b>=</b> en <b>klik</b> op de cel. Excel zet het adres in je formule.</p>`,
        widget: { type: 'sheetSim', rows: 5, cols: 3, widths: [110, 70, 70], cells: { A1: 'Prijs', B1: 2.5, A2: 'Aantal', B2: 4, A3: 'Totaal', A5: 'Per persoon' }, rounds: [
          { select: 'B3', ask: 'Bereken in <span class="addr">B3</span> het totaal: typ <b>=</b>, klik op B1, typ <b>*</b>, klik op B2 en druk op Enter.',
            check: (s, ch) => {
              if (ch !== 'B3') return null;
              const inf = s.info('B3');
              if (!inf) return no('Een formule begint met =.');
              if (!inf.hasRef) return no('Je rekende met getallen. Gebruik de celadressen B1 en B2.');
              if (!inf.cells.includes('B1') || !inf.cells.includes('B2')) return no('Gebruik B1 en B2 in je formule.');
              return s.num('B3') === 10 ? ok('Juist: 10.') : no(`Je formule geeft ${s.show('B3')}. Het moet 10 zijn.`);
            } },
          { select: 'B2', ask: 'Er worden 6 stuks gekocht. Verander <span class="addr">B2</span> in <b>6</b>. Kijk wat er met B3 gebeurt!',
            check: (s, ch) => ch !== 'B2' ? null : s.num('B2') !== 6 ? no('Typ 6 in B2.') : s.isFormula('B3') ? ok(`B3 rekende vanzelf opnieuw: ${s.show('B3')}. Daarom gebruik je celadressen!`) : no('B3 is geen formule meer. Typ eerst weer =B1*B2 in B3.') },
          { select: 'B5', ask: 'Twee personen delen alles. Bereken in <span class="addr">B5</span> hoeveel iedereen betaalt: het totaal gedeeld door 2.',
            check: (s, ch) => ch !== 'B5' ? null : !s.info('B5') ? no('Begin met =.') : !s.info('B5').cells.includes('B3') ? no('Gebruik het totaal uit B3 in je formule.') : s.num('B5') === s.num('B3') / 2 ? ok(`Juist: ${s.show('B5')}.`) : no(`Je formule geeft ${s.show('B5')}. Deel B3 door 2.`) }
        ] } },
      { id: 'l3', title: 'Formules doortrekken',
        body: `<p>Moet je in elke rij dezelfde berekening maken? Typ de formule één keer en <b>trek ze door</b> met de vulgreep, net zoals een reeks.</p>
<p>Excel past de adressen vanzelf aan: <b>=A2*B2</b> in rij 2 wordt <b>=A3*B3</b> in rij 3, <b>=A4*B4</b> in rij 4 …</p>
<p>Klik daarna op een cel: in de formulebalk zie je welke formule Excel maakte.</p>`,
        widget: { type: 'fillSim', formulaBar: true, keep: 'Klik nu op <span class="addr">C4</span> en op <span class="addr">C6</span>. Welke formule zie je in de formulebalk?', doneMsg: 'Juist! Excel trok de formule door en paste de adressen aan.', rounds: [
          { rows: 6, cols: 3, cells: { A1: 'Aantal', B1: 'Prijs', C1: 'Totaal', A2: 3, B2: 2, C2: '=A2*B2', A3: 5, B3: 1.5, A4: 2, B4: 4, A5: 10, B5: 0.5, A6: 4, B6: 3 },
            goal: { C3: 7.5, C4: 8, C5: 5, C6: 12 }, ask: 'In <span class="addr">C2</span> staat <b>=A2*B2</b>. Klik op C2 en sleep de vulgreep naar beneden tot <span class="addr">C6</span>.' }
        ] } },
      { id: 'l4', title: 'Wat doet Excel bij doortrekken?',
        body: `<p>Excel onthoudt niet "B2", maar "de cel links van mij". Daarom schuiven de adressen mee als je een formule doortrekt.</p>`,
        widget: { type: 'quiz', questions: [
          { q: 'In C2 staat =A2*B2. Je trekt de formule door tot C5. Welke formule staat in C5?', options: ['=A5*B5', '=A2*B2', '=A5*B2'], correct: 0, explain: 'Je ging 3 rijen naar beneden, dus alle rijnummers gaan 3 omhoog.' },
          { q: 'In C2 staat =A2+B2. Je trekt de formule naar rechts, naar D2. Wat staat er in D2?', options: ['=A2+B2', '=B2+C2', '=A3+B3'], correct: 1, explain: 'Eén kolom naar rechts: alle letters schuiven één plaats op.' },
          { q: 'Waarom trek je een formule door?', options: ['Het is sneller en Excel past de adressen zelf aan', 'Anders werkt de formule niet', 'Dan wordt de tekst vet'], correct: 0, explain: 'Eén keer goed denken, Excel doet de rest.' }
        ] } },
      { id: 'l5', title: 'Voorrang en haakjes',
        body: `<p>Excel rekent zoals in de wiskundeles: eerst <b>*</b> en <b>/</b>, dan pas <b>+</b> en <b>-</b>. Wil je iets eerst laten uitrekenen, zet het dan tussen <b>haakjes</b>.</p>`,
        widget: { type: 'quiz', questions: [
          { q: 'Wat is de uitkomst van =2+3*4?', options: ['20', '14', '24'], correct: 1, explain: 'Eerst 3*4 = 12, dan 2+12 = 14.' },
          { q: 'Wat is de uitkomst van =(2+3)*4?', options: ['20', '14', '11'], correct: 0, explain: 'Haakjes eerst: 2+3 = 5, dan 5*4 = 20.' },
          { q: 'Je wil het gemiddelde van A1 en A2. Welke formule klopt?', options: ['=A1+A2/2', '=(A1+A2)/2', '=A1/2+A2'], correct: 1, explain: 'Eerst samen optellen (haakjes), dan delen door 2.' }
        ] } },
      { id: 'l6', title: 'Foutmeldingen',
        body: `<p>Soms toont Excel een foutmelding in plaats van een uitkomst:</p>
<p><b>#DEEL/0!</b> – je deelt door 0 of door een lege cel.<br><b>#NAAM?</b> – Excel begrijpt een woord in je formule niet (bv. een x in plaats van *).<br><b>#####</b> – geen fout: de kolom is te smal voor het getal. Maak de kolom breder.</p>`,
        widget: { type: 'sheetSim', rows: 5, cols: 3, widths: [110, 70, 70], cells: { A1: 'Totaal', B1: 36, A2: 'Kinderen', A3: 'Per kind', B3: '=B1/B2' }, rounds: [
          { select: 'B2', ask: 'B3 toont <b>#DEEL/0!</b> omdat B2 leeg is. Los het op: er zijn <b>9</b> kinderen.',
            check: (s, ch) => ch !== 'B2' ? null : s.num('B3') === 4 ? ok('Opgelost! 36 gedeeld door 9 is 4.') : no('Typ 9 in B2.') },
          { select: 'B5', ask: 'Typ in <span class="addr">B5</span> eens <b>=B1xB2</b> (met een x). Welke foutmelding zie je?',
            check: (s, ch) => ch !== 'B5' ? null : s.isErr('B5') && s.show('B5') === '#NAAM?' ? ok('Juist: #NAAM?. Excel kent x niet als teken. Vermenigvuldigen doe je met *.') : no('Typ precies =B1xB2 in B5.') }
        ] } }
    ],
    startFile(seed, mode) {
      const d = m6data(seed, mode), b = { bold: true };
      const cells = { A1: { v: d.titel, bold: true, size: 14 }, A3: { v: 'Artikel', ...b }, B3: { v: 'Prijs per stuk', ...b }, C3: { v: 'Aantal', ...b }, D3: { v: 'Totaal', ...b },
        A11: { v: 'Totaal', ...b }, F3: { v: 'Budget', ...b }, G3: d.budget, F4: { v: 'Over', ...b }, F6: { v: 'Aantal leerlingen', ...b }, G6: d.kids, F7: { v: 'Per leerling', ...b } };
      d.items.forEach((x, i) => { cells['A' + (4 + i)] = x.n; cells['B' + (4 + i)] = x.p; cells['C' + (4 + i)] = x.a; });
      return { name: mode === 'test' ? 'Test6_formules.xlsx' : 'Module6_formules.xlsx', spec: { sheetName: 'Inkopen', cells, colWidths: { A: 14, B: 14, D: 11, F: 18 } } };
    },
    gen(seed, mode) {
      const d = m6data(seed, mode), it = d.items;
      const rowMsg = r => (missing, used) => `De formule in D${r} moet met rij ${r} rekenen (B${r} en C${r}), maar ze gebruikt ${list(used) || 'andere cellen'}.`;
      return { startFile: true, tasks: [
        { id: 't1', text: `Bereken in <span class="addr">D4</span> het totaal voor <b>${it[0].n.toLowerCase()}</b>: prijs per stuk × aantal. Gebruik een formule met de celadressen.`,
          hints: ['Een formule begint met =. Vermenigvuldigen doe je met *.', 'Typ = , klik op B4, typ * , klik op C4 en druk op Enter.'],
          check: ({ sheet }) => checkFormula(sheet, 'D4', { expect: round9(it[0].p * it[0].a), refs: ['B4', 'C4'] }) },
        { id: 't2', text: `<b>Trek de formule</b> van D4 door tot <span class="addr">D9</span>.`,
          hints: ['Je hoeft de formule niet opnieuw te typen.', 'Klik op D4 en sleep de vulgreep naar beneden tot D9.'],
          check: ({ sheet }) => {
            for (let k = 1; k < it.length; k++) {
              const r = 4 + k, a = 'D' + r;
              if (!sheet.formula(a) && isNum(sheet.value(a))) return no(`In ${a} staat een getal, geen formule. Trek de formule van D4 door met de vulgreep.`);
              const res = checkFormula(sheet, a, { expect: round9(it[k].p * it[k].a), refs: ['B' + r, 'C' + r], allRefsMsg: rowMsg(r) });
              if (!res.ok) return res;
            }
            return ok();
          } },
        { id: 't3', text: `Tel in <span class="addr">D11</span> alle totalen op, van D4 tot D9, met een formule.`,
          hints: ['Je kan cellen optellen met +.', 'Typ =D4+D5+D6+D7+D8+D9 (of klik telkens op de cel) en druk op Enter.'],
          check: ({ sheet }) => checkFormula(sheet, 'D11', { expect: d.total, refs: ['D4', 'D5', 'D6', 'D7', 'D8', 'D9'], allRefsMsg: miss => `Je formule telt ${list(miss)} nog niet mee.` }) },
        { id: 't4', text: `Bereken in <span class="addr">G4</span> hoeveel geld er <b>overblijft</b> van het budget: budget min het totaal.`,
          hints: ['Welke twee cellen heb je nodig? Het budget en het totaal.', 'Typ =G3-D11 en druk op Enter.'],
          check: ({ sheet }) => {
            const exp = numOf(sheet.calc('G3')) - numOf(sheet.calc('D11'));
            return checkFormula(sheet, 'G4', { expect: round9(exp), refs: ['G3', 'D11'], wrongMsg: v => isNum(v) && Math.abs(v + exp) < 1e-6 ? 'Je uitkomst is negatief: je trok af in de verkeerde volgorde. Begin met het budget: =G3-D11.' : null });
          } },
        { id: 't5', text: `Bereken in <span class="addr">G7</span> hoeveel het kost <b>per leerling</b>: het totaal gedeeld door het aantal leerlingen.`,
          hints: ['Delen doe je met /.', 'Typ =D11/G6 en druk op Enter.'],
          check: ({ sheet }) => {
            const exp = numOf(sheet.calc('D11')) / numOf(sheet.calc('G6'));
            return checkFormula(sheet, 'G7', { expect: round9(exp), refs: ['D11', 'G6'], wrongMsg: v => isNum(v) && Math.abs(v - 1 / exp) < 1e-6 ? 'Je deelde omgekeerd. Eerst het totaal, dan delen door het aantal leerlingen: =D11/G6.' : null });
          } },
        { id: 't6', text: `Toon <span class="addr">G7</span> in <b>euro met 2 decimalen</b>.`,
          hints: ['Dat leerde je in stap 5.', 'Klik op G7 en klik op de knop met het geldbiljet (€).'],
          check: ({ sheet }) => { const i = sheet.numFmt('G7'); return i.kind === 'currency' && i.dec === 2 ? ok() : i.kind === 'currency' ? no(`G7 is in euro, maar met ${i.dec} ${i.dec === 1 ? 'decimaal' : 'decimalen'}. Het moeten er 2 zijn.`) : no('G7 is nog niet in euro.'); } },
        { id: 't7', text: `Er komen <b>${d.kids2}</b> leerlingen in plaats van ${d.kids}. Verander het getal in <span class="addr">G6</span> en kijk wat er gebeurt met G7.`,
          hints: ['Je hoeft G7 niet te veranderen.', `Klik op G6, typ ${d.kids2} en druk op Enter. G7 rekent vanzelf opnieuw.`],
          check: ({ sheet }) => {
            if (numOf(sheet.value('G6')) !== d.kids2) return no(`In G6 moet ${d.kids2} staan.`);
            if (!sheet.formula('G7')) return no('G7 is geen formule meer. Typ opnieuw =D11/G6.');
            return ok();
          } }
      ] };
    }
  };

  // ======================================================================
  // MODULE 7 – Functies
  // ======================================================================
  const NAMEN = ['Lotte', 'Milan', 'Noor', 'Yasin', 'Fien', 'Rayan', 'Emma', 'Lucas', 'Lina', 'Adam', 'Mila', 'Jules', 'Zoë', 'Nathan', 'Elif', 'Wout'];
  function m7data(seed, mode) {
    const r = rng(seed + '|m7|' + mode);
    const names = r.shuffle(NAMEN).slice(0, 10);
    let scores;
    for (let t = 0; t < 500; t++) {
      const pool = r.shuffle(Array.from({ length: mode === 'test' ? 200 : 84 }, (_, i) => i + (mode === 'test' ? 180 : 12))).slice(0, 9);
      const avg = pool.reduce((a, b) => a + b, 0) / 9;
      const f1 = avg * 10 - Math.floor(avg * 10), f0 = avg - Math.floor(avg);
      if (Math.abs(f0 - 0.5) > 0.06 && f0 > 0.05 && Math.abs(f1 - 0.5) > 0.06) { scores = pool; break; }
    }
    const absent = r.int(1, 8);
    const col = names.map((n, i) => i === absent ? 'afwezig' : scores[i < absent ? i : i - 1]);
    return { names, col, nums: scores, dec: mode === 'test' ? 1 : 0, titel: mode === 'test' ? 'Sportdag: verspringen (cm)' : 'Sportdag: touwtjespringen' };
  }
  const B_RANGE = Array.from({ length: 10 }, (_, i) => 'B' + (4 + i));

  function checkFunc(sheet, a, d, fn, expect, extra) {
    const needed = B_RANGE.filter(c => isNum(sheet.value(c)));
    const NL = { SUM: 'SOM', AVERAGE: 'GEMIDDELDE', MAX: 'MAX', MIN: 'MIN', MEDIAN: 'MEDIAAN', COUNT: 'AANTAL', ROUND: 'AFRONDEN' };
    const f = sheet.formula(a);
    if (f && FX.isErr(sheet.calc(a))) return checkFormula(sheet, a, {});
    if (f) {
      const inf = FX.info(f, 'en');
      const buiten = [...new Set(inf.cells)].filter(c => !/^B(3|4|5|6|7|8|9|10|11|12|13)$/.test(c) && !(extra && extra.allowRefs && extra.allowRefs.includes(c)));
      const anyFn = ['SUM', 'AVERAGE', 'MAX', 'MIN', 'MEDIAN', 'COUNT', 'COUNTA'].some(x => inf.funcs.includes(x));
      if (!anyFn && inf.hasRef) return no(`Je rekende het zelf uit met ${dutch(f)}. Gebruik de functie ${NL[fn]}: =${NL[fn]}(B4:B13).`);
      if (buiten.length && fn !== 'ROUND') return no(`Je formule gebruikt ook ${list(buiten)}. Gebruik alleen de scores: B4:B13.`);
      if (fn === 'COUNT' && inf.funcs.includes('COUNTA')) return no('AANTALARG telt ook de tekst "afwezig" mee. Gebruik AANTAL: die telt alleen getallen.');
      const others = ['SUM', 'AVERAGE', 'MAX', 'MIN', 'MEDIAN', 'COUNT'].filter(x => x !== fn && inf.funcs.includes(x) && !(extra && extra.allow && extra.allow.includes(x)));
      if (!inf.funcs.includes(fn) && others.length) return no(`Je gebruikte ${others.map(x => NL[x]).join(' en ')}. Hier heb je ${NL[fn]} nodig.`);
      if (!inf.funcs.includes(fn) && fn !== 'ROUND' && inf.hasRef) return no(`Je rekende het zelf uit met ${dutch(f)}. Gebruik de functie ${NL[fn]}: =${NL[fn]}(B4:B13).`);
    }
    return checkFormula(sheet, a, Object.assign({ expect, funcs: [fn], funcMsg: `Gebruik de functie ${NL[fn]}.`, refs: extra && extra.noRefs ? null : needed,
      allRefsMsg: miss => `Je bereik is niet volledig: ${list(miss)} ${miss.length > 1 ? 'tellen' : 'telt'} niet mee. Het bereik is B4:B13.` }, extra || {}));
  }

  const m7 = {
    id: 'm7', nr: 7,
    title: 'Functies',
    short: 'Excel heeft ingebouwde rekenhulpjes: SOM, GEMIDDELDE, MIN, MAX, MEDIAAN, AANTAL en AFRONDEN.',
    skills: ['SOM', 'GEMIDDELDE', 'MIN en MAX', 'MEDIAAN', 'AANTAL', 'AFRONDEN', 'de puntkomma'],
    lessons: [
      { id: 'l1', title: 'Een functie: een rekenhulpje',
        body: `<p>Weet je nog hoe lang de formule <b>=D4+D5+D6+D7+D8+D9</b> was? Met een <b>functie</b> gaat het korter: <b>=SOM(D4:D9)</b>.</p>
<p>Een functie heeft een <b>naam</b> (SOM), dan <b>haakjes</b>, en daartussen een <b>bereik</b> (D4:D9). Het bereik schrijf je met een <b>dubbelpunt</b>, zoals in stap 1.</p>
<p>Snel: klik onder je getallen en klik op de knop <b>AutoSom Σ</b> (tab Start), of druk <span class="key">Alt</span> + <span class="key">=</span>.</p>`,
        widget: { type: 'sheetSim', rows: 6, cols: 2, widths: [90, 90], cells: { A1: 'Week', B1: 'Gespaard', A2: 'week 1', B2: 3, A3: 'week 2', B3: 5, A4: 'week 3', B4: 2, A5: 'week 4', B5: 6, A6: 'Totaal' }, rounds: [
          { select: 'B6', ask: 'Bereken in <span class="addr">B6</span> het totaal met de functie <b>=SOM(B2:B5)</b>.',
            check: (s, ch) => {
              if (ch !== 'B6') return null;
              const i = s.info('B6');
              if (!i) return no('Een formule begint met =.');
              if (i.error === '#SEP') return no('Tussen de eerste en de laatste cel van een bereik zet je een dubbelpunt: B2:B5.');
              if (!i.funcs.includes('SUM')) return s.num('B6') === 16 ? no('Dat klopt, maar gebruik nu de functie SOM: =SOM(B2:B5).') : no('Gebruik de functie SOM: =SOM(B2:B5).');
              return s.num('B6') === 16 ? ok('Juist: 16. Kort en krachtig!') : no(`Je formule geeft ${s.show('B6')}. Controleer het bereik: van B2 tot B5.`);
            } },
          { select: 'B2', ask: 'In week 1 werd er toch 10 euro gespaard. Verander <span class="addr">B2</span> in <b>10</b>. Kijk naar het totaal.',
            check: (s, ch) => ch !== 'B2' ? null : s.num('B2') === 10 && s.num('B6') === 23 ? ok('Het totaal rekende vanzelf mee: 23.') : no('Typ 10 in B2. In B6 moet de formule met SOM blijven staan.') }
        ] } },
      { id: 'l2', title: 'SOM, GEMIDDELDE, MIN en MAX',
        body: `<p><b>SOM</b> telt alles op. <b>GEMIDDELDE</b> telt alles op en deelt door het aantal getallen. <b>MAX</b> zoekt het grootste getal, <b>MIN</b> het kleinste.</p>
<p>Ze werken allemaal op dezelfde manier: <b>=GEMIDDELDE(B2:B6)</b>, <b>=MAX(B2:B6)</b> …</p>`,
        widget: { type: 'sheetSim', rows: 6, cols: 5, widths: [70, 50, 20, 110, 70], cells: { A1: 'Naam', B1: 'Punten', A2: 'Lotte', B2: 8, A3: 'Milan', B3: 3, A4: 'Noor', B4: 10, A5: 'Yasin', B5: 5, A6: 'Fien', B6: 9, D2: 'Gemiddelde', D3: 'Hoogste', D4: 'Laagste' }, rounds: [
          { select: 'E2', ask: 'Bereken in <span class="addr">E2</span> het gemiddelde van de punten.', check: (s, ch) => ch !== 'E2' ? null : simFunc(s, 'E2', 'AVERAGE', 7) },
          { select: 'E3', ask: 'Zoek in <span class="addr">E3</span> de hoogste score.', check: (s, ch) => ch !== 'E3' ? null : simFunc(s, 'E3', 'MAX', 10) },
          { select: 'E4', ask: 'Zoek in <span class="addr">E4</span> de laagste score.', check: (s, ch) => ch !== 'E4' ? null : simFunc(s, 'E4', 'MIN', 3) }
        ] } },
      { id: 'l3', title: 'MEDIAAN en AANTAL',
        body: `<p><b>MEDIAAN</b> zoekt het <b>middelste</b> getal. Zet de getallen in volgorde: het getal in het midden is de mediaan. Bij een even aantal neemt Excel het gemiddelde van de twee middelste.</p>
<p><b>AANTAL</b> telt hoeveel <b>getallen</b> er in een bereik staan. Tekst en lege cellen telt het niet mee. (<b>AANTALARG</b> telt alles wat niet leeg is.)</p>`,
        widget: { type: 'quiz', questions: [
          { q: 'Wat is de mediaan van 3, 9 en 4?', options: ['4', '9', '5,33'], correct: 0, explain: 'In volgorde: 3, 4, 9. Het middelste getal is 4.' },
          { q: 'Wat is de mediaan van 2, 4, 6 en 100?', options: ['28', '5', '4'], correct: 1, explain: 'De twee middelste zijn 4 en 6. Hun gemiddelde is 5. Het rare getal 100 telt nauwelijks mee: daarom is de mediaan handig.' },
          { q: 'In B2:B5 staat: 5, "afwezig", 7 en een lege cel. Wat geeft =AANTAL(B2:B5)?', options: ['4', '3', '2'], correct: 2, explain: 'Alleen 5 en 7 zijn getallen.' }
        ] } },
      { id: 'l4', title: 'AFRONDEN en de puntkomma',
        body: `<p>Een gemiddelde heeft vaak veel cijfers na de komma. Met <b>AFRONDEN</b> rond je het getal echt af: <b>=AFRONDEN(B6;1)</b> rondt B6 af op 1 decimaal.</p>
<p>Let op de <b>puntkomma ;</b> tussen de twee delen. In een Belgische Excel gebruik je altijd een puntkomma tussen de delen van een functie.</p>
<p>Verschil met stap 5: "Minder decimalen" verandert alleen wat je <b>ziet</b>. AFRONDEN verandert het <b>getal zelf</b>.</p>`,
        widget: { type: 'sheetSim', rows: 7, cols: 2, widths: [110, 90], cells: { A1: 'Toets', B1: 'Punten', A2: 'toets 1', B2: 7, A3: 'toets 2', B3: 8, A4: 'toets 3', B4: 8, A6: 'Gemiddelde', A7: 'Afgerond' }, rounds: [
          { select: 'B6', ask: 'Bereken in <span class="addr">B6</span> het gemiddelde van de drie toetsen.', check: (s, ch) => ch !== 'B6' ? null : simFunc(s, 'B6', 'AVERAGE', 23 / 3) },
          { select: 'B7', ask: 'Rond het gemiddelde in <span class="addr">B7</span> af op <b>1 decimaal</b>: =AFRONDEN(B6;1)',
            check: (s, ch) => {
              if (ch !== 'B7') return null;
              const i = s.info('B7');
              if (!i) return no('Een formule begint met =.');
              if (i.error === '#SEP' || /AFRONDEN\([^;]*,/i.test(String(s.raw('B7')))) return no('Gebruik een puntkomma tussen de delen: =AFRONDEN(B6;1).');
              if (!i.funcs.includes('ROUND')) return no('Gebruik de functie AFRONDEN.');
              return s.num('B7') === 7.7 ? ok('Juist: 7,7.') : no(`Je formule geeft ${s.show('B7')}. Het moet 7,7 zijn.`);
            } }
        ] } },
      { id: 'l5', title: 'Tikfouten en andere talen',
        body: `<p>Typ je de naam van een functie fout, bv. <b>=SOMM(B2:B5)</b>, dan toont Excel <b>#NAAM?</b>.</p>
<p>Kleine letters mogen: typ je <b>=som(b2:b5)</b>, dan maakt Excel er zelf hoofdletters van.</p>
<p>Werk je ooit in een Engelse Excel? Dan hebben de functies Engelse namen: SOM is <b>SUM</b>, GEMIDDELDE is <b>AVERAGE</b>, en je gebruikt een komma in plaats van een puntkomma.</p>`,
        widget: { type: 'quiz', questions: [
          { q: 'Je typt =SOMM(B2:B5). Wat toont Excel?', options: ['De som', '#NAAM?', '#DEEL/0!'], correct: 1, explain: 'SOMM bestaat niet. Excel begrijpt de naam niet.' },
          { q: 'Je typt =gemiddelde(b2:b5) in kleine letters. Werkt dat?', options: ['Ja, Excel maakt er hoofdletters van', 'Nee, je krijgt een fout'], correct: 0, explain: 'Kleine letters zijn geen probleem.' },
          { q: 'Hoe heet SOM in een Engelse Excel?', options: ['TOTAL', 'SUM', 'PLUS'], correct: 1, explain: 'SUM. Een bestand dat je in een Nederlandse Excel maakt, werkt ook in een Engelse: Excel vertaalt de namen zelf.' }
        ] } }
    ],
    startFile(seed, mode) {
      const d = m7data(seed, mode), b = { bold: true };
      const cells = { A1: { v: d.titel, bold: true, size: 14 }, A3: { v: 'Naam', ...b }, B3: { v: mode === 'test' ? 'Afstand' : 'Sprongen', ...b },
        D4: 'Alles samen', D5: 'Gemiddelde', D6: d.dec ? 'Gemiddelde (1 decimaal)' : 'Gemiddelde afgerond', D7: 'Hoogste', D8: 'Laagste', D9: 'Mediaan', D10: 'Aantal deelnemers' };
      d.names.forEach((n, i) => { cells['A' + (4 + i)] = n; cells['B' + (4 + i)] = d.col[i]; });
      return { name: mode === 'test' ? 'Test7_functies.xlsx' : 'Module7_functies.xlsx', spec: { sheetName: 'Sportdag', cells, colWidths: { A: 12, B: 11, D: 26, E: 12 } } };
    },
    gen(seed, mode) {
      const d = m7data(seed, mode), v = d.nums, n = v.length;
      const sum = v.reduce((a, b) => a + b, 0), avg = sum / n, sorted = v.slice().sort((a, b) => a - b);
      const rnd = FX.evaluate(`ROUND(${avg},${d.dec})`, { locale: 'en', get: () => null });
      return { startFile: true, tasks: [
        { id: 't1', text: `Bereken in <span class="addr">E4</span> ${mode === 'test' ? 'alle afstanden' : 'alle sprongen'} samen met de functie <b>SOM</b>.`,
          hints: ['Een functie: naam, haakjes, bereik.', 'Typ =SOM(B4:B13) of gebruik de knop AutoSom Σ en controleer het bereik.'],
          check: ({ sheet }) => checkFunc(sheet, 'E4', d, 'SUM', sum) },
        { id: 't2', text: `Bereken in <span class="addr">E5</span> het <b>gemiddelde</b>.`,
          hints: ['De functie heet GEMIDDELDE.', 'Typ =GEMIDDELDE(B4:B13). De tekst "afwezig" telt Excel niet mee.'],
          check: ({ sheet }) => checkFunc(sheet, 'E5', d, 'AVERAGE', avg) },
        { id: 't3', text: `Toon in <span class="addr">E6</span> het gemiddelde, <b>afgerond op ${d.dec ? '1 decimaal' : 'een geheel getal (0 decimalen)'}</b>, met de functie AFRONDEN.`,
          hints: [`AFRONDEN heeft twee delen: het getal en het aantal decimalen, met een puntkomma ertussen.`, `Typ =AFRONDEN(E5;${d.dec}) of =AFRONDEN(GEMIDDELDE(B4:B13);${d.dec}).`],
          check: ({ sheet }) => {
            const f = sheet.formula('E6');
            if (f && !FX.info(f, 'en').funcs.includes('ROUND')) return no('Gebruik de functie AFRONDEN. Met "Minder decimalen" verandert alleen de weergave, niet het getal zelf.');
            return checkFormula(sheet, 'E6', { expect: rnd, funcs: ['ROUND'], funcMsg: 'Gebruik de functie AFRONDEN.',
              wrongMsg: (val) => isNum(val) && Math.abs(val - FX.evaluate(`ROUND(${avg},${d.dec ? 0 : 1})`, { locale: 'en', get: () => null })) < 1e-9 ? `Je rondde af op ${d.dec ? 0 : 1} ${d.dec ? 'decimalen' : 'decimaal'}. Het tweede deel van AFRONDEN is het aantal decimalen: ${d.dec}.` : null });
          } },
        { id: 't4', text: `Zoek in <span class="addr">E7</span> de <b>hoogste</b> score met een functie.`,
          hints: ['Het grootste getal zoek je met MAX.', 'Typ =MAX(B4:B13).'],
          check: ({ sheet }) => checkFunc(sheet, 'E7', d, 'MAX', sorted[n - 1]) },
        { id: 't5', text: `Zoek in <span class="addr">E8</span> de <b>laagste</b> score met een functie.`,
          hints: ['Het kleinste getal zoek je met MIN.', 'Typ =MIN(B4:B13). Afwezig telt niet mee.'],
          check: ({ sheet }) => checkFunc(sheet, 'E8', d, 'MIN', sorted[0]) },
        { id: 't6', text: `Zoek in <span class="addr">E9</span> de <b>mediaan</b>: de middelste score.`,
          hints: ['De functie heet MEDIAAN.', 'Typ =MEDIAAN(B4:B13).'],
          check: ({ sheet }) => checkFunc(sheet, 'E9', d, 'MEDIAN', sorted[(n - 1) / 2]) },
        { id: 't7', text: `Tel in <span class="addr">E10</span> hoeveel leerlingen <b>meededen</b> met de functie AANTAL.`,
          hints: ['AANTAL telt alleen de getallen in een bereik.', 'Typ =AANTAL(B4:B13). De afwezige leerling telt dan niet mee.'],
          check: ({ sheet }) => checkFunc(sheet, 'E10', d, 'COUNT', n) }
      ] };
    }
  };
  /* controle in het mini-rekenblad: juiste functie en uitkomst */
  function simFunc(s, a, fn, expect) {
    const NL = { SUM: 'SOM', AVERAGE: 'GEMIDDELDE', MAX: 'MAX', MIN: 'MIN', MEDIAN: 'MEDIAAN', COUNT: 'AANTAL' };
    const i = s.info(a);
    if (!i) return no('Een formule begint met =.');
    if (i.error === '#SEP') return no('Tussen de eerste en de laatste cel van een bereik zet je een dubbelpunt, bv. B2:B6.');
    if (s.isErr(a)) return no(`${a} toont ${s.show(a)}. Controleer de naam van de functie.`);
    if (!i.funcs.includes(fn)) { const other = i.funcs.find(x => NL[x]); return no(other ? `Je gebruikte ${NL[other]}. Hier heb je ${NL[fn]} nodig.` : `Gebruik de functie ${NL[fn]}.`); }
    const v = s.num(a);
    return v != null && Math.abs(v - expect) < 1e-9 ? ok(`Juist: ${s.show(a)}.`) : no(`Je formule geeft ${s.show(a)}. Controleer het bereik.`);
  }

  // ======================================================================
  // MODULE 8 – Project: mijn maandrapport
  // ======================================================================
  const SCHOOLMAANDEN = ['september', 'oktober', 'november', 'december', 'januari', 'februari', 'maart', 'april', 'mei', 'juni'];
  const ALIAS = {
    vak: ['vak', 'vakken', 'onderdeel', 'toets', 'toetsen', 'leergebied'],
    score: ['score', 'scores', 'punten', 'behaald', 'resultaat', 'mijnpunten', 'mijnscore', 'punt'],
    max: ['maximum', 'max', 'op', 'maximaal', 'maxpunten', 'maximumscore', 'opdeel', 'totaalpunten'],
    pct: ['procent', '%', 'percent', 'percentage', 'procenten', 'procent%']
  };
  const key = t => norm(t).replace(/[^a-z%]/g, '');
  /* Zoekt de rapporttabel, waar die ook staat. */
  function findReport(sheet) {
    if (sheet._rep !== undefined) return sheet._rep;
    let rep = null;
    const cand = sheet.filled().filter(a => ALIAS.vak.includes(key(sheet.text(a)))).map(parseAddr).sort((p, q) => p.r - q.r || p.c - q.c);
    for (const p of cand) {
      const cols = { vak: p.c };
      for (let c = p.c + 1; c <= p.c + 8; c++) {
        const k = key(sheet.text(addr(c, p.r)));
        for (const nm of ['score', 'max', 'pct']) if (!cols[nm] && ALIAS[nm].includes(k)) cols[nm] = c;
      }
      if (cols.score && cols.max && cols.pct) {
        const rows = [];
        let r = p.r + 1, tot = null, avg = null;
        const label = rr => { for (let c = Math.max(1, p.c - 1); c <= p.c; c++) { const t = key(sheet.text(addr(c, rr))); if (t) return t; } return ''; };
        for (; r <= p.r + 60; r++) {
          const t = label(r);
          if (!t) { if (label(r + 1) || label(r + 2)) continue; break; }
          if (t.startsWith('totaal')) { tot = tot || r; continue; }
          if (t.startsWith('gemiddel')) { avg = avg || r; continue; }
          if (tot || avg) break;
          rows.push(r);
        }
        rep = { hdr: p.r, cols, rows, tot, avg, A: (k, rr) => addr(cols[k], rr) };
        break;
      }
    }
    sheet._rep = rep;
    return rep;
  }
  const noTable = () => no('Ik vind je tabel nog niet. Zet de kopjes Vak, Score, Maximum en Procent naast elkaar op één rij.');
  function pctFmt(sheet, a) { const k = sheet.numFmt(a).kind; return k === 'percent' ? null : `${a} rekent goed, maar toont nog geen procent. Klik op %.`; }
  function rowRatio(sheet, a, sA, mA, label) {
    const f = sheet.formula(a);
    const s = numOf(sheet.calc(sA)), m = numOf(sheet.calc(mA));
    if (!f) {
      const v = sheet.value(a);
      if (v == null || String(v).trim() === '') return no(`${a} (procent ${label}) is nog leeg.`);
      return no(`In ${a} staat een getal, geen formule. Laat Excel het procent berekenen: =${sA}/${mA}`);
    }
    const inf = FX.info(f, 'en'), v = sheet.calc(a);
    if (FX.isErr(v)) return checkFormula(sheet, a, {});
    if (!inf.cells.includes(sA) || !inf.cells.includes(mA)) {
      const used = [...new Set(inf.cells)];
      return no(`De formule in ${a} moet rekenen met ${sA} en ${mA} (${label}), maar ze gebruikt ${list(used) || 'geen cellen'}.`);
    }
    if (!(isNum(v) && s != null && m ? Math.abs(v - s / m) < 0.005 : false)) {
      if (isNum(v) && s != null && m && Math.abs(v - m / s) < 1e-6) return no(`In ${a} deelde je omgekeerd. Het is score gedeeld door maximum: =${sA}/${mA}`);
      return no(`De formule in ${a} geeft ${showV(v)}. Score gedeeld door maximum is ${showV(s / m)}.`);
    }
    return null;
  }
  function m8data(seed) {
    const r = rng(seed + '|m8');
    const pool = [['Wiskunde 1'], ['Wiskunde 2'], ['Wiskunde 3'], ['Taal 1'], ['Taal 2'], ['Spelling 1'], ['Frans 1'], ['Frans 2'], ['WO 1']];
    const n = r.int(6, 8);
    const keep = r.shuffle(pool.map((_, i) => i)).slice(0, n).sort((a, b) => a - b);
    const vakken = keep.map(i => { const max = r.pick([10, 15, 20, 25, 30, 40, 50]); return { n: pool[i][0], max, s: r.int(Math.ceil(max * 0.45), max) }; });
    return { maand: r.pick(SCHOOLMAANDEN), vakken };
  }

  const m8 = {
    id: 'm8', nr: 8, project: true,
    title: 'Project: mijn maandrapport',
    short: 'Je bouwt je eigen rapport voor dagelijks werk, van een lege werkmap tot een verzorgde tabel die zelf rekent. Eerst een oefenrapport, dan je eigen rapport met je echte punten.',
    skills: ['plannen', 'procent per vak', 'SOM en GEMIDDELDE', 'procentnotatie', 'opmaak', 'eigen punten'],
    tabLabels: { practice: 'Oefenrapport', test: 'Mijn eigen rapport' },
    testGate: `<ul class="list-plain">
<li>Nu maak je <b>je eigen rapport</b> met je <b>echte punten</b> van deze maand (minstens 5 vakken).</li>
<li>Je mag je oefenrapport gebruiken als vertrekpunt: <b>Bestand → Opslaan als</b> met een nieuwe naam, en vervang de vakken en punten.</li>
<li>Er zijn geen hulplijnen. De app kijkt de opbouw na; je leerkracht kijkt mee naar je punten.</li>
<li>Je hebt 80% nodig.</li></ul>`,
    lessons: [
      { id: 'l1', title: 'Een plan maken',
        body: `<p>Een goed rapport begin je met een <b>plan</b>. Kijk naar het voorbeeld: wat staat er allemaal in?</p>
<div class="grid-wrap"><table class="mini sample"><tbody>
<tr><td colspan="4"><b>Rapport dagelijks werk – oktober</b></td></tr>
<tr><td colspan="4">Lotte Peeters</td></tr>
<tr><td><b>Vak</b></td><td><b>Score</b></td><td><b>Maximum</b></td><td><b>Procent</b></td></tr>
<tr><td>Wiskunde 1</td><td class="n">27</td><td class="n">40</td><td class="n">68%</td></tr>
<tr><td>Taal 1</td><td class="n">15</td><td class="n">25</td><td class="n">60%</td></tr>
<tr><td>Frans 1</td><td class="n">25</td><td class="n">30</td><td class="n">83%</td></tr>
<tr><td><b>Totaal</b></td><td class="n">67</td><td class="n">95</td><td class="n">71%</td></tr>
<tr><td><b>Gemiddeld</b></td><td class="n">22,3</td><td class="n">31,7</td><td class="n">71%</td></tr>
</tbody></table></div>
<p>Alles wat grijs zou zijn op papier, laat je Excel berekenen: het procent, het totaal en het gemiddelde.</p>`,
        widget: { type: 'steps', steps: ['Een <b>titel</b> met het woord rapport en de maand', 'Je <b>naam</b>', 'De <b>kopjes</b> Vak, Score, Maximum en Procent naast elkaar', 'Eén rij per <b>vak</b> met je score en het maximum', 'Het <b>procent</b> per vak: een formule', 'Een rij <b>Totaal</b> met SOM', 'Een rij <b>Gemiddeld</b> met GEMIDDELDE', 'Mooie <b>opmaak</b>: vet, randen, procent'], confirm: 'Mijn plan is klaar' } },
      { id: 'l2', title: 'Procent per vak',
        body: `<p>27 op 40: hoeveel procent is dat? Je deelt de score door het maximum: 27 / 40 = 0,675. Met de knop % (stap 5) zie je 68%.</p>
<p>In Excel schrijf je dat met celadressen, zodat het rapport vanzelf klopt als je een punt verandert.</p>`,
        widget: { type: 'sheetSim', rows: 3, cols: 4, widths: [100, 60, 80, 80], cells: { A1: 'Vak', B1: 'Score', C1: 'Maximum', D1: 'Procent', A2: 'Wiskunde 1', B2: 27, C2: 40, A3: 'Taal 1', B3: 15, C3: 25 }, rounds: [
          { select: 'D2', ask: 'Bereken in <span class="addr">D2</span> het procent voor wiskunde: score gedeeld door maximum.',
            check: (s, ch) => ch !== 'D2' ? null : !s.info('D2') ? no('Een formule begint met =.') : !(s.info('D2').cells.includes('B2') && s.info('D2').cells.includes('C2')) ? no('Gebruik B2 en C2 in je formule.') : s.num('D2') === 0.675 ? ok('Juist: 0,675. In Excel klik je daarna op % en zie je 68%.') : s.num('D2') === 40 / 27 ? no('Omgekeerd: eerst de score, dan delen door het maximum.') : no(`Je formule geeft ${s.show('D2')}.`) },
          { select: 'D3', ask: 'Doe hetzelfde voor taal in <span class="addr">D3</span>.',
            check: (s, ch) => ch !== 'D3' ? null : s.info('D3') && s.info('D3').cells.includes('B3') && s.info('D3').cells.includes('C3') && s.num('D3') === 0.6 ? ok('Juist: 0,6 is 60%.') : no('Typ =B3/C3.') }
        ] } },
      { id: 'l3', title: 'Denkvraag: twee soorten gemiddelde',
        body: `<p>Lotte heeft drie toetsen: <b>4 op 5</b> (80%), <b>30 op 50</b> (60%) en <b>9 op 10</b> (90%).</p>
<p>Manier 1: alles samen. 4 + 30 + 9 = 43 op 5 + 50 + 10 = 65. Dat is <b>66%</b>.<br>Manier 2: het gemiddelde van de procenten. (80 + 60 + 90) / 3 = <b>77%</b>.</p>
<p>Twee keer een "gemiddelde", maar een andere uitkomst! Denk even na voor je de vragen beantwoordt.</p>`,
        widget: { type: 'quiz', questions: [
          { q: 'Waarom geven de twee manieren een andere uitkomst?', options: ['Excel maakt een rekenfout', 'Bij manier 1 telt de toets op 50 veel zwaarder mee dan de toets op 5', 'Bij manier 2 tel je de toetsen dubbel'], correct: 1, explain: 'Bij manier 1 telt elk punt even zwaar. Een grote toets heeft dus meer invloed. Bij manier 2 telt elke toets even zwaar, groot of klein.' },
          { q: 'Je rapport gebruikt manier 1: totaal van de scores gedeeld door totaal van de maxima. Welke formule hoort bij de rij Totaal?', options: ['=B8/C8 (totaal score / totaal maximum)', '=GEMIDDELDE(D2:D7)', '=B8*C8'], correct: 0, explain: 'Zo telt elk punt dat je haalde even zwaar.' },
          { q: 'En in de rij Gemiddeld: gemiddelde score / gemiddeld maximum. Wat komt daar uit?', options: ['Hetzelfde procent als bij Totaal', 'Altijd 50%', 'Het gemiddelde van de procenten'], correct: 0, explain: 'Teller en noemer worden allebei door hetzelfde aantal vakken gedeeld, dus het procent blijft gelijk. Knap als je dat zelf zag!' }
        ] } },
      { id: 'l4', title: 'Een rapport dat je graag leest',
        body: `<p>Een goed rapport lees je in één oogopslag. Dit helpt:</p>
<p>• een <b>titel</b> die groter of vet is (stap 3)<br>• <b>vette kopjes</b> en eventueel een opvulkleur<br>• <b>alle randen</b> rond de tabel<br>• kolommen die <b>passen</b>: dubbelklik tussen de kolomletters (stap 2)<br>• procenten <b>zonder of met 1 decimaal</b>, niet met tien (stap 5)<br>• de rijen <b>Totaal</b> en <b>Gemiddeld</b> vet, zodat ze opvallen</p>`,
        widget: { type: 'read', confirm: 'Ik let erop' } },
      { id: 'l5', title: 'Je eigen rapport',
        body: `<p>Na het oefenrapport maak je <b>je eigen rapport</b> met je echte punten van deze maand. Je vindt ze in je agenda, op je toetsen of in je map.</p>
<p>Je mag zoveel vakken toevoegen als je wil, maar zorg dat <b>SOM en GEMIDDELDE alle vakken</b> meetellen. Voeg je een vak toe in het midden van de tabel (rij invoegen, stap 2), dan past Excel de formules vanzelf aan.</p>
<p>De app kent jouw punten niet. Ze kijkt na of je rapport goed <b>rekent</b> en goed <b>opgebouwd</b> is. Je leerkracht ziet je bestand.</p>`,
        widget: { type: 'read', confirm: 'Begrepen' } }
    ],
    newFile: true,
    gen(seed, mode, student) {
      const own = mode === 'test';
      const d = m8data(seed);
      const vn = (student && student.voornaam) || '';
      const tabel = `<div class="grid-wrap"><table class="mini sample"><tbody><tr><td><b>Vak</b></td><td><b>Score</b></td><td><b>Maximum</b></td></tr>${d.vakken.map(v => `<tr><td>${v.n}</td><td class="n">${v.s}</td><td class="n">${v.max}</td></tr>`).join('')}</tbody></table></div>`;
      const needRep = fn => ({ sheet }) => { const rep = findReport(sheet); return rep ? fn(sheet, rep) : noTable(); };
      const allRowsMsg = (what, miss) => `${what} telt ${list(miss)} nog niet mee. Neem alle vakken op in het bereik.`;
      return { startFile: null, tasks: [
        { id: 't1', text: own ? 'Typ bovenaan een <b>titel</b> met het woord <b>rapport</b> en de <b>maand</b>.' : `Typ bovenaan een <b>titel</b> met het woord <b>rapport</b> en de maand <b>${d.maand}</b>, bv. "Rapport dagelijks werk – ${d.maand}".`,
          hints: ['Typ de titel in A1.', 'Maak de titel vet of groter, dat komt later ook aan bod.'],
          check: ({ sheet }) => {
            const rep = findReport(sheet), hdr = rep ? rep.hdr : 999;
            const cells = sheet.filled().filter(a => parseAddr(a).r < hdr && /rapport/.test(norm(sheet.text(a))));
            if (!cells.length) return no('Ik vind nog geen titel met het woord "rapport" boven je tabel.');
            const okMonth = cells.some(a => own ? MAANDEN.some(m => norm(sheet.text(a)).includes(m)) : norm(sheet.text(a)).includes(d.maand));
            return okMonth ? ok() : no(own ? 'Zet ook de maand in je titel.' : `Zet ook de maand ${d.maand} in je titel.`);
          } },
        { id: 't2', text: 'Zet je <b>naam</b> onder de titel.',
          hints: ['Een aparte cel onder de titel is genoeg.', 'Typ je voornaam en naam in A2.'],
          check: ({ sheet }) => sheet.filled().some(a => norm(sheet.text(a)).includes(norm(vn))) ? ok() : no(`Ik vind je naam (${vn}) nog niet in je rapport.`) },
        { id: 't3', text: 'Maak een tabel met de <b>kopjes</b> Vak, Score, Maximum en Procent naast elkaar.',
          hints: ['Vier kopjes op dezelfde rij, bv. in A4 tot D4.', 'Typ Vak, Score, Maximum en Procent in vier cellen naast elkaar.'],
          check: ({ sheet }) => findReport(sheet) ? ok() : noTable() },
        own
          ? { id: 't4', text: 'Vul je <b>eigen vakken en punten</b> van deze maand in: minstens 5 vakken, met je score en het maximum.',
            hints: [], check: needRep((sheet, rep) => {
              if (rep.rows.length < 5) return no(`Ik vind ${rep.rows.length} vak${rep.rows.length === 1 ? '' : 'ken'}. Er moeten er minstens 5 zijn.`);
              for (const r of rep.rows) {
                const s = numOf(sheet.calc(rep.A('score', r))), m = numOf(sheet.calc(rep.A('max', r)));
                if (s == null) return no(`In ${rep.A('score', r)} staat geen score (getal).`);
                if (m == null || m <= 0) return no(`In ${rep.A('max', r)} staat geen maximum (getal groter dan 0).`);
                if (s > m) return no(`In rij ${r} is de score (${showV(s)}) hoger dan het maximum (${showV(m)}). Kijk je punten na.`);
              }
              return ok();
            }) }
          : { id: 't4', text: `Typ de <b>vakken en punten</b> over in je tabel, onder de kopjes:${tabel}`,
            hints: ['Eén rij per vak, in dezelfde volgorde.', 'Typ het vak in de kolom Vak, de score onder Score en het maximum onder Maximum.'],
            check: needRep((sheet, rep) => {
              if (rep.rows.length !== d.vakken.length) return no(`Ik vind ${rep.rows.length} vakken in je tabel. Het moeten er ${d.vakken.length} zijn.`);
              for (let k = 0; k < d.vakken.length; k++) {
                const r = rep.rows[k], v = d.vakken[k];
                if (norm(sheet.text(rep.A('vak', r))) !== norm(v.n)) return no(`In rij ${r} verwacht ik het vak ${v.n}, maar daar staat "${sheet.text(rep.A('vak', r))}".`);
                if (numOf(sheet.value(rep.A('score', r))) !== v.s) return no(`De score van ${v.n} (${rep.A('score', r)}) moet ${v.s} zijn.`);
                if (numOf(sheet.value(rep.A('max', r))) !== v.max) return no(`Het maximum van ${v.n} (${rep.A('max', r)}) moet ${v.max} zijn.`);
              }
              return ok();
            }) },
        { id: 't5', text: 'Bereken het <b>procent per vak</b> met een formule (score gedeeld door maximum) en toon het als <b>procent</b>.',
          hints: ['Typ de formule één keer en trek ze door naar beneden.', 'Bv. =B5/C5 in de eerste rij, doortrekken, en dan de knop %.'],
          check: needRep((sheet, rep) => {
            if (!rep.rows.length) return no('Er staan nog geen vakken in je tabel.');
            for (const r of rep.rows) {
              const a = rep.A('pct', r), res = rowRatio(sheet, a, rep.A('score', r), rep.A('max', r), sheet.text(rep.A('vak', r)));
              if (res) return res;
              const f = pctFmt(sheet, a); if (f) return no(f);
            }
            return ok();
          }) },
        { id: 't6', text: 'Maak onder de vakken een rij <b>Totaal</b>: tel de scores en de maxima op met <b>SOM</b>, en bereken het procent.',
          hints: ['Typ "Totaal" in de kolom Vak, onder het laatste vak.', 'Gebruik =SOM(...) onder Score en onder Maximum. Het procent is totaal score gedeeld door totaal maximum.'],
          check: needRep((sheet, rep) => {
            if (!rep.tot) return no('Ik vind nog geen rij met het woord "Totaal" onder je vakken.');
            const t = rep.tot, sc = rep.rows.map(r => rep.A('score', r)), mx = rep.rows.map(r => rep.A('max', r));
            const sumOf = arr => arr.reduce((a, c) => a + (numOf(sheet.calc(c)) || 0), 0);
            for (const [k, cells, word] of [['score', sc, 'scores'], ['max', mx, 'maxima']]) {
              const a = rep.A(k, t);
              const res = checkFormula(sheet, a, { expect: round9(sumOf(cells)), funcs: ['SUM'], funcMsg: `Tel de ${word} op met de functie SOM (stap 7).`, refs: cells, allRefsMsg: miss => allRowsMsg(`De SOM in ${a}`, miss) });
              if (!res.ok) return res;
            }
            const pa = rep.A('pct', t), res = rowRatio(sheet, pa, rep.A('score', t), rep.A('max', t), 'totaal');
            if (res) return res;
            const f = pctFmt(sheet, pa); return f ? no(f) : ok();
          }) },
        { id: 't7', text: 'Maak een rij <b>Gemiddeld</b>: het gemiddelde van de scores en van de maxima met <b>GEMIDDELDE</b>, en het procent.',
          hints: ['Typ "Gemiddeld" in de kolom Vak, onder Totaal.', 'Gebruik =GEMIDDELDE(...) onder Score en onder Maximum. Het procent is gemiddelde score gedeeld door gemiddeld maximum.'],
          check: needRep((sheet, rep) => {
            if (!rep.avg) return no('Ik vind nog geen rij met het woord "Gemiddeld" onder je vakken.');
            const g = rep.avg, sc = rep.rows.map(r => rep.A('score', r)), mx = rep.rows.map(r => rep.A('max', r));
            const avgOf = arr => arr.reduce((a, c) => a + (numOf(sheet.calc(c)) || 0), 0) / arr.length;
            for (const [k, cells, word] of [['score', sc, 'scores'], ['max', mx, 'maxima']]) {
              const a = rep.A(k, g);
              const res = checkFormula(sheet, a, { expect: round9(avgOf(cells)), funcs: ['AVERAGE'], funcMsg: `Bereken het gemiddelde van de ${word} met de functie GEMIDDELDE.`, refs: cells, allRefsMsg: miss => allRowsMsg(`GEMIDDELDE in ${a}`, miss) });
              if (!res.ok) return res;
            }
            const pa = rep.A('pct', g), f = sheet.formula(pa);
            if (f && FX.info(f, 'en').funcs.includes('AVERAGE')) return no(`In ${pa} nam je het gemiddelde van de procenten. Weet je nog, de denkvraag? Hier deel je de gemiddelde score door het gemiddelde maximum: =${rep.A('score', g)}/${rep.A('max', g)}`);
            const res = rowRatio(sheet, pa, rep.A('score', g), rep.A('max', g), 'gemiddeld');
            if (res) return res;
            const ff = pctFmt(sheet, pa); return ff ? no(ff) : ok();
          }) },
        { id: 't8', text: 'Werk je rapport mooi af: <b>kopjes vet</b>, <b>alle randen</b> rond de tabel (van de kopjes tot Gemiddeld) en een <b>titel</b> die vet of groter is.',
          hints: ['Selecteer de hele tabel en kies "Alle randen".', 'Kopjes: selecteer ze en klik op B. Titel: klik erop en kies een grotere lettergrootte.'],
          check: needRep((sheet, rep) => {
            const hdrCells = ['vak', 'score', 'max', 'pct'].map(k => rep.A(k, rep.hdr));
            const notBold = hdrCells.filter(a => !sheet.style(a).bold);
            if (notBold.length) return no(`De kopjes ${list(notBold)} zijn nog niet vet.`);
            const last = Math.max(rep.avg || 0, rep.tot || 0, rep.rows[rep.rows.length - 1] || rep.hdr);
            const c1 = Math.min(...Object.values(rep.cols)), c2 = Math.max(...Object.values(rep.cols));
            const miss = [];
            for (let r = rep.hdr; r <= last; r++) for (let c = c1; c <= c2; c++) { const a = addr(c, r); if (!['left', 'right', 'top', 'bottom'].every(sd => sheet.hasEdge(a, sd))) miss.push(a); }
            if (miss.length) return no(`Nog niet alle cellen van je tabel hebben randen, bv. ${list(miss.slice(0, 3))}. Selecteer de tabel van ${addr(c1, rep.hdr)} tot ${addr(c2, last)} en kies "Alle randen".`);
            const title = sheet.filled().find(a => parseAddr(a).r < rep.hdr && /rapport/.test(norm(sheet.text(a))));
            if (title) { const st = sheet.style(title); if (!st.bold && !(st.size > 11)) return no('Maak je titel vet of groter.'); }
            return ok();
          }) }
      ] };
    }
  };

  // ======================================================================
  // MODULE 9 – Excel kiest: ALS
  // ======================================================================
  /* Rekent de formule in een cel uit alsof sommige cellen een andere inhoud hadden. */
  function tryWith(sheet, a, over) {
    const f = sheet.formula(a);
    return FX.evaluate(f, { locale: 'en', get: (ref, sh) => (!sh && Object.prototype.hasOwnProperty.call(over, ref)) ? over[ref] : (sh ? (sheet.wb && sheet.wb.sheet(sh) ? sheet.wb.sheet(sh).calc(ref) : new FX.Err('#REF!')) : sheet.calc(ref)) });
  }
  const said = v => norm(showV(v));
  function ifBasics(sheet, a, refs) {
    const f = sheet.formula(a);
    if (!f) { const r = checkFormula(sheet, a, {}); return r.ok ? no(`In ${a} moet een formule met ALS staan.`) : r; }
    const inf = FX.info(f, 'en');
    const words = inf.badNames.filter(x => /^[A-Z]{2,}$/i.test(x) && !FX.NL_EN[x] && !FX.EN_NL[x]);
    if (words.length) return no(`Excel kent het woord ${words[0].toLowerCase()} niet. Tekst zet je tussen aanhalingstekens: "${words[0].toLowerCase()}".`);
    const v = sheet.calc(a);
    if (FX.isErr(v)) return checkFormula(sheet, a, {});
    if (!inf.funcs.includes('IF')) return no(`Gebruik de functie ALS in ${a}.`);
    const miss = refs.filter(r => !inf.cells.includes(r));
    if (miss.length) return no(`De formule in ${a} moet ${list(refs)} gebruiken. ${list(miss)} ${miss.length > 1 ? 'ontbreken' : 'ontbreekt'}.`);
    return null;
  }
  /* Is de juist/jammer-formule logisch goed? (de app probeert zelf antwoorden uit) */
  function ifAnswerCheck(sheet, a, ans, sol, W, needEmpty) {
    const base = ifBasics(sheet, a, [ans, sol]);
    if (base) return base;
    const s = sheet.calc(sol);
    const wrong = isNum(s) ? s + 1 : String(s) + 'x';
    const rOk = said(tryWith(sheet, a, { [ans]: s })), rBad = said(tryWith(sheet, a, { [ans]: wrong })), rEmpty = said(tryWith(sheet, a, { [ans]: null }));
    const okW = norm(W.ok), badW = norm(W.bad), emW = norm(W.empty);
    if (rOk === badW && rBad === okW) return no(`In ${a} staan "${W.ok}" en "${W.bad}" omgewisseld. Eerst wat Excel toont als het klopt, dan wat het toont als het niet klopt.`);
    if (rOk !== okW) return no(`Als het antwoord juist is (${showV(s)}), toont je formule in ${a} "${rOk}". Dat moet "${W.ok}" zijn.`);
    if (rBad !== badW) return no(`Als het antwoord fout is, toont je formule in ${a} "${rBad}". Dat moet "${W.bad}" zijn.`);
    if (needEmpty && rEmpty !== emW) {
      if (rEmpty === badW) return no(`Als ${ans} leeg is, toont ${a} nog "${W.bad}". Er moet "${W.empty}" staan. Begin met ALS(${ans}="";"${W.empty}";…)`);
      return no(`Als ${ans} leeg is, toont ${a} "${rEmpty}". Er moet "${W.empty}" staan.`);
    }
    return ok();
  }
  const VRAGEN = [['6 x 7 =', 42], ['Hoofdstad van Italië?', 'Rome'], ['Verkleinwoord van bal?', 'balletje'], ['72 : 6 =', 12], ['De helft van 36?', 18], ['Planeet het dichtst bij de zon?', 'Mercurius'],
    ['4 m = ... cm (alleen het getal)', 400], ['Meervoud van kind?', 'kinderen'], ['1000 - 632 =', 368], ['Hoeveel poten heeft een spin?', 8], ['Hoofdstad van Spanje?', 'Madrid'], ['Wat komt na goed, beter, ...?', 'best']];
  function m9data(seed, mode) {
    const r = rng(seed + '|m9|' + mode);
    const test = mode === 'test';
    const q = r.shuffle(VRAGEN).slice(0, 6);
    const pattern = r.shuffle(['ok', 'ok', 'bad', 'bad', 'empty', 'empty']);
    const answers = q.map((x, i) => pattern[i] === 'ok' ? x[1] : pattern[i] === 'empty' ? null : (isNum(x[1]) ? x[1] + r.pick([-2, -1, 1, 10]) : x[1].slice(0, -1)));
    return {
      q, pattern, answers,
      W: test ? { ok: 'goed', bad: 'fout', empty: 'vul in' } : { ok: 'juist', bad: 'jammer', empty: 'vul in' },
      S: test ? { t: r.int(5, 6), pass: 'Proficiat!', fail: 'Probeer opnieuw' } : { t: 4, pass: 'Goed gedaan!', fail: 'Oefen nog wat' },
      score: r.int(2, 6)
    };
  }

  const m9 = {
    id: 'm9', nr: 9,
    title: 'Excel kiest: ALS',
    short: 'Met ALS laat je Excel zelf beslissen: "juist" of "jammer", "geslaagd" of "nog oefenen". Het hart van een interactief werkblad.',
    skills: ['de functie ALS', 'vergelijken', 'juist of jammer', 'lege cel', 'ALS in een ALS', 'tekst tussen aanhalingstekens'],
    lessons: [
      { id: 'l1', title: 'Excel beslist',
        body: `<p>Met <b>ALS</b> laat je Excel kiezen wat er in een cel komt. De functie heeft drie delen, met een puntkomma ertussen:</p>
<p><b>=ALS(</b>voorwaarde<b>;</b> wat als het klopt<b>;</b> wat als het niet klopt<b>)</b></p>
<p>Voorbeeld: <b>=ALS(B1&gt;=5;"geslaagd";"nog oefenen")</b>. Tekst zet je altijd tussen <b>aanhalingstekens</b>.</p>`,
        widget: { type: 'sheetSim', rows: 3, cols: 2, widths: [100, 160], cells: { A1: 'Punten', B1: 7, A2: 'Boodschap' }, rounds: [
          { select: 'B2', ask: 'Typ in <span class="addr">B2</span>: <b>=ALS(B1&gt;=5;"geslaagd";"nog oefenen")</b>',
            check: (s, ch) => {
              if (ch !== 'B2') return null;
              if (s.isErr('B2') && s.show('B2') === '#NAAM?') return no('Excel begrijpt een woord niet. Staat de tekst tussen aanhalingstekens "…"?');
              if (!s.info('B2') || !s.info('B2').funcs.includes('IF')) return no('Gebruik de functie ALS.');
              return norm(s.show('B2')) === 'geslaagd' ? ok('Juist! 7 is minstens 5, dus Excel kiest "geslaagd".') : no(`B2 toont "${s.show('B2')}". Kijk je formule na.`);
            } },
          { select: 'B1', ask: 'Verander de punten in <span class="addr">B1</span> in <b>3</b>. Wat kiest Excel nu?',
            check: (s, ch) => ch !== 'B1' ? null : s.num('B1') === 3 && norm(s.show('B2')) === 'nog oefenen' ? ok('Excel koos zelf "nog oefenen". Jij hoeft niets te veranderen!') : no('Typ 3 in B1.') }
        ] } },
      { id: 'l2', title: 'Vergelijken',
        body: `<p>De voorwaarde is een vergelijking. Excel antwoordt met <b>WAAR</b> of <b>ONWAAR</b>.</p>
<p><span class="key">=</span> is gelijk aan · <span class="key">&lt;&gt;</span> is niet gelijk aan · <span class="key">&lt;</span> kleiner dan · <span class="key">&gt;</span> groter dan · <span class="key">&lt;=</span> kleiner dan of gelijk aan · <span class="key">&gt;=</span> groter dan of gelijk aan</p>`,
        widget: { type: 'quiz', questions: [
          { q: 'Wat geeft =7>=7?', options: ['WAAR', 'ONWAAR'], correct: 0, explain: '>= betekent groter dan of gelijk aan. 7 is gelijk aan 7.' },
          { q: 'Je wil "geslaagd" vanaf 10 punten (10 is ook geslaagd). Welke voorwaarde?', options: ['B2>10', 'B2>=10', 'B2=10'], correct: 1, explain: 'Met > zou 10 punten net niet geslaagd zijn.' },
          { q: 'Wat betekent <>?', options: ['is niet gelijk aan', 'is ongeveer gelijk aan', 'is kleiner of groter dan 0'], correct: 0, explain: '=ALS(C2<>B2;"jammer";"juist") werkt dus ook.' }
        ] } },
      { id: 'l3', title: 'Juist of jammer?',
        body: `<p>Zo werkt een interactief werkblad: in kolom B staat de <b>oplossing</b>, in kolom C typt de leerling zijn <b>antwoord</b>, en in kolom D vergelijkt Excel de twee:</p>
<p><b>=ALS(C2=B2;"juist";"jammer")</b></p>`,
        widget: { type: 'sheetSim', rows: 2, cols: 4, widths: [90, 80, 80, 90], cells: { A1: 'Vraag', B1: 'Oplossing', C1: 'Antwoord', D1: 'Feedback', A2: '6 x 7 =', B2: 42 }, rounds: [
          { select: 'D2', ask: 'Typ in <span class="addr">D2</span> de formule die "juist" of "jammer" toont.',
            check: (s, ch) => {
              if (ch !== 'D2') return null;
              const i = s.info('D2');
              if (!i || !i.funcs.includes('IF')) return no('Gebruik de functie ALS.');
              if (s.isErr('D2')) return no('Excel begrijpt een woord niet. Staat de tekst tussen aanhalingstekens?');
              if (!i.cells.includes('C2') || !i.cells.includes('B2')) return no('Vergelijk het antwoord (C2) met de oplossing (B2).');
              return norm(s.show('D2')) === 'jammer' ? ok('Je formule werkt! C2 is nog leeg, dus toont Excel "jammer". Typ nu een antwoord.') : no(`D2 toont "${s.show('D2')}". Bij een leeg antwoord verwacht ik "jammer".`);
            } },
          { select: 'C2', ask: 'Typ het juiste antwoord <b>42</b> in <span class="addr">C2</span>.',
            check: (s, ch) => ch !== 'C2' ? null : norm(s.show('D2')) === 'juist' ? ok('"juist"! Excel kijkt het antwoord voor jou na.') : no('Typ 42 in C2.') },
          { select: 'C2', ask: 'Typ nu een fout antwoord in <span class="addr">C2</span>, bv. 40.',
            check: (s, ch) => ch !== 'C2' ? null : norm(s.show('D2')) === 'jammer' ? ok('"jammer". Je werkblad werkt in beide richtingen.') : no('Typ een fout antwoord in C2.') }
        ] } },
      { id: 'l4', title: 'Eerst kijken of er iets staat',
        body: `<p>Probleem: is het antwoord nog <b>leeg</b>, dan zegt het werkblad al "jammer". Dat is niet eerlijk!</p>
<p>Oplossing: eerst kijken of de cel leeg is. <b>""</b> (twee aanhalingstekens zonder iets ertussen) betekent "leeg". Je zet een ALS in een ALS:</p>
<p><b>=ALS(C2="";"vul in";ALS(C2=B2;"juist";"jammer"))</b></p>`,
        widget: { type: 'sheetSim', rows: 2, cols: 4, widths: [90, 80, 80, 90], cells: { A1: 'Vraag', B1: 'Oplossing', C1: 'Antwoord', D1: 'Feedback', A2: '6 x 7 =', B2: 42, D2: '=ALS(C2=B2;"juist";"jammer")' }, rounds: [
          { select: 'D2', ask: 'Verbeter de formule in <span class="addr">D2</span> zodat er "vul in" staat als C2 leeg is.',
            check: (s, ch) => {
              if (ch !== 'D2') return null;
              const i = s.info('D2');
              if (!i || i.funcs.filter(x => x === 'IF').length < 2) return no('Gebruik twee keer ALS: =ALS(C2="";"vul in";ALS(C2=B2;"juist";"jammer"))');
              if (s.isErr('D2')) return no('Er loopt iets mis. Tel de haakjes en kijk of alle tekst tussen aanhalingstekens staat.');
              return norm(s.show('D2')) === 'vul in' ? ok('Juist: C2 is leeg, dus "vul in".') : no(`D2 toont "${s.show('D2')}". Er moet "vul in" staan.`);
            } },
          { select: 'C2', ask: 'Typ <b>42</b> in <span class="addr">C2</span>. Wat zegt het werkblad nu?',
            check: (s, ch) => ch !== 'C2' ? null : norm(s.show('D2')) === 'juist' ? ok('"juist"! Je werkblad is nu eerlijk.') : no('Typ 42 in C2.') }
        ] } },
      { id: 'l5', title: 'Tekst vergelijken',
        body: `<p>Bij tekst let Excel op elk teken. Hoofdletters tellen <b>niet</b> mee, maar een <b>spatie</b> wel.</p>`,
        widget: { type: 'quiz', questions: [
          { q: 'De oplossing is "Rome". Een leerling typt "rome". Wat zegt =ALS(C2=B2;"juist";"jammer")?', options: ['juist', 'jammer'], correct: 0, explain: 'Hoofdletters en kleine letters zijn voor deze vergelijking hetzelfde.' },
          { q: 'Een leerling typt "Rome " (met een spatie erachter). Wat zegt de formule?', options: ['juist', 'jammer'], correct: 1, explain: 'De spatie telt mee. In stap 11 leer je SPATIES.WISSEN, die zulke spaties weghaalt.' },
          { q: 'De oplossing is 400. Een leerling typt "400 cm". Wat zegt de formule?', options: ['juist', 'jammer'], correct: 1, explain: '"400 cm" is tekst, geen getal. Zet daarom bij de vraag: typ alleen het getal.' }
        ] } }
    ],
    startFile(seed, mode) {
      const d = m9data(seed, mode), b = { bold: true };
      const cells = { A1: { v: 'Mijn eerste interactief werkblad', bold: true, size: 14 }, A3: { v: 'Vraag', ...b }, B3: { v: 'Oplossing', ...b }, C3: { v: 'Antwoord', ...b }, D3: { v: 'Feedback', ...b },
        A11: { v: 'Mijn score', ...b }, B11: d.score, A12: { v: 'Boodschap', ...b } };
      d.q.forEach((x, i) => { cells['A' + (4 + i)] = x[0]; cells['B' + (4 + i)] = x[1]; if (d.answers[i] != null) cells['C' + (4 + i)] = d.answers[i]; });
      return { name: mode === 'test' ? 'Test9_als.xlsx' : 'Module9_als.xlsx', spec: { sheetName: 'Werkblad', cells, colWidths: { A: 34, B: 13, C: 13, D: 14 } } };
    },
    gen(seed, mode) {
      const d = m9data(seed, mode), W = d.W, S = d.S;
      const emptyRows = d.pattern.map((p, i) => p === 'empty' ? 4 + i : null).filter(Boolean);
      return { startFile: true, tasks: [
        { id: 't1', text: `Maak in <span class="addr">D4</span> een formule met <b>ALS</b>: is het antwoord (C4) gelijk aan de oplossing (B4), dan toont Excel <b>"${W.ok}"</b>, anders <b>"${W.bad}"</b>.`,
          hints: ['Drie delen: de voorwaarde, wat als het klopt, wat als het niet klopt. Tekst tussen aanhalingstekens.', `Typ =ALS(C4=B4;"${W.ok}";"${W.bad}")`],
          check: ({ sheet }) => ifAnswerCheck(sheet, 'D4', 'C4', 'B4', W, false) },
        { id: 't2', text: `Trek de formule door tot <span class="addr">D9</span>.`,
          hints: ['Met de vulgreep, net als in stap 6.', 'Klik op D4 en sleep de vulgreep naar beneden tot D9.'],
          check: ({ sheet }) => { for (let r = 5; r <= 9; r++) { const res = ifAnswerCheck(sheet, 'D' + r, 'C' + r, 'B' + r, W, false); if (!res.ok) return res; } return ok(); } },
        { id: 't3', text: `Is een antwoord nog leeg, dan moet er <b>"${W.empty}"</b> staan in plaats van "${W.bad}". Verbeter de formule in D4 en trek ze opnieuw door tot D9.`,
          hints: ['Zet een ALS in een ALS: eerst kijken of C4 leeg is.', `Typ =ALS(C4="";"${W.empty}";ALS(C4=B4;"${W.ok}";"${W.bad}")) en trek opnieuw door.`],
          check: ({ sheet }) => {
            const first = ifAnswerCheck(sheet, 'D4', 'C4', 'B4', W, true);
            if (!first.ok) return first;
            for (let r = 5; r <= 9; r++) { const res = ifAnswerCheck(sheet, 'D' + r, 'C' + r, 'B' + r, W, true); if (!res.ok) return no(`D4 is goed! Trek de verbeterde formule opnieuw door tot D9: in D${r} staat nog de oude formule.`); }
            return ok();
          } },
        { id: 't4', text: `In <span class="addr">B11</span> staat de score. Laat in <span class="addr">B12</span> <b>"${S.pass}"</b> verschijnen als de score <b>${S.t} of meer</b> is, en anders <b>"${S.fail}"</b>.`,
          hints: ['"… of meer" schrijf je als >= (groter dan of gelijk aan).', `Typ =ALS(B11>=${S.t};"${S.pass}";"${S.fail}")`],
          check: ({ sheet }) => {
            const base = ifBasics(sheet, 'B12', ['B11']);
            if (base) return base;
            const at = v => said(tryWith(sheet, 'B12', { B11: v }));
            const P = norm(S.pass), Fa = norm(S.fail);
            if (at(S.t + 2) === Fa && at(0) === P) return no(`"${S.pass}" en "${S.fail}" staan omgewisseld.`);
            if (at(S.t) === Fa && at(S.t + 1) === P) return no(`Met precies ${S.t} punten toont je formule "${S.fail}". "${S.t} of meer" schrijf je met >= in plaats van >.`);
            if (at(S.t) !== P || at(S.t + 3) !== P) return no(`Met ${S.t} punten of meer moet B12 "${S.pass}" tonen. Nu toont het "${at(S.t)}".`);
            if (at(S.t - 1) !== Fa || at(0) !== Fa) return no(`Met minder dan ${S.t} punten moet B12 "${S.fail}" tonen. Nu toont het "${at(S.t - 1)}".`);
            return ok();
          } },
        { id: 't5', text: `Test je werkblad: typ het <b>juiste antwoord</b> in de lege antwoordcellen (${list(emptyRows.map(r => 'C' + r))}). Toont kolom D nu "${W.ok}"?`,
          hints: ['Kijk in kolom B wat de oplossing is.', 'Typ precies wat er in kolom B staat.'],
          check: ({ sheet }) => {
            for (const r of emptyRows) {
              const c = sheet.value('C' + r), sol = sheet.value('B' + r);
              if (c == null || String(c).trim() === '') return no(`C${r} is nog leeg.`);
              if (said(sheet.calc('D' + r)) !== norm(W.ok) && norm(c) === norm(sol) && String(c) === String(c).trim()) return no(`Het antwoord in C${r} klopt, maar D${r} toont "${showV(sheet.calc('D' + r))}". Je formule werkt nog niet goed: kijk opdracht 1 tot 3 na.`);
              if (said(sheet.calc('D' + r)) !== norm(W.ok)) return no(`In C${r} staat "${showV(c)}", maar D${r} toont "${showV(sheet.calc('D' + r))}". ${typeof c === 'string' && c !== c.trim() ? 'Er staat een spatie te veel.' : `De oplossing is ${showV(sol)}.`}`);
            }
            return ok();
          } }
      ] };
    }
  };

  // ======================================================================
  // MODULE 10 – Keuzelijsten en kleuren die reageren
  // ======================================================================
  const WW = [['ik … (worden, tegenwoordige tijd)', 'word', ['word', 'wordt', 'werd', 'worden']], ['hij … (vinden, tegenwoordige tijd)', 'vindt', ['vind', 'vindt', 'vond', 'vinden']], ['wij … (lopen, verleden tijd)', 'liepen', ['lopen', 'liep', 'liepen', 'loopt']], ['jij … (antwoorden, tegenwoordige tijd)', 'antwoordt', ['antwoord', 'antwoordt', 'antwoordde', 'antwoorden']]];
  const HOOFDSTEDEN = [['Hoofdstad van Italië?', 'Rome', ['Parijs', 'Rome', 'Brussel', 'Riga']], ['Hoofdstad van Spanje?', 'Madrid', ['Lissabon', 'Madrid', 'Rome', 'Berlijn']], ['Hoofdstad van Frankrijk?', 'Parijs', ['Parijs', 'Londen', 'Wenen', 'Praag']], ['Hoofdstad van Duitsland?', 'Berlijn', ['München', 'Berlijn', 'Hamburg', 'Bern']]];
  const REKEN = [['6 x 7 =', 42], ['72 : 6 =', 12], ['1000 - 632 =', 368], ['De helft van 36?', 18], ['125 + 375 =', 500]];
  const OVERIG = [['Verkleinwoord van bal?', 'balletje'], ['Planeet het dichtst bij de zon?', 'Mercurius'], ['Meervoud van kind?', 'kinderen'], ['Hoeveel poten heeft een spin?', 8], ['Wat komt na goed, beter, …?', 'best'], ['Hoeveel dagen heeft een week?', 7]];
  function m10data(seed, mode) {
    const r = rng(seed + '|m10|' + mode);
    const test = mode === 'test';
    const ww = r.pick(WW), hs = r.pick(HOOFDSTEDEN), rk = r.pick(REKEN), ov = r.shuffle(OVERIG).slice(0, 3);
    const W = test ? { ok: 'goed', bad: 'fout' } : { ok: 'juist', bad: 'jammer' };
    const lo = test ? 0.6 : 0.5, hi = test ? 0.85 : 0.8;
    // rapportpunten: zorg voor procenten onder, tussen en boven de grenzen
    const vakken = ['Wiskunde 1', 'Wiskunde 2', 'Taal 1', 'Taal 2', 'Frans 1', 'WO 1'];
    const targets = r.shuffle([lo - 0.12, lo - 0.05, (lo + hi) / 2, (lo + hi) / 2 + 0.04, hi + 0.06, hi + 0.12]);
    const rap = vakken.map((n, k) => { const max = r.pick([20, 25, 40, 50]); return { n, max, s: Math.max(1, Math.min(max, Math.round(targets[k] * max))) }; });
    return { ww, hs, rk, ov, W, lo, hi, rap, maand: r.pick(['november', 'december', 'maart', 'mei']), test };
  }
  const pctTxt = v => nlNum(Math.round(v * 100)) + '%';

  const m10 = {
    id: 'm10', nr: 10,
    title: 'Keuzelijsten en kleuren die reageren',
    short: 'Je maakt keuzelijsten en laat alleen goede antwoorden toe. Met voorwaardelijke opmaak kleurt Excel zelf: groen bij juist, rood bij jammer. Ook je rapport krijgt kleuren.',
    skills: ['keuzelijst', 'alleen getallen toelaten', 'eigen foutmelding', 'voorwaardelijke opmaak', 'kleuren bij juist en jammer', 'kleuren in je rapport'],
    lessons: [
      { id: 'l1', title: 'Een keuzelijst',
        body: `<p>Met een <b>keuzelijst</b> kiest de gebruiker het antwoord met een pijltje. Zo kan hij zich niet vertypen.</p>
<p>Klik op de cel. Ga naar het tabblad <b>Gegevens</b> en klik op <b>Gegevensvalidatie</b>. Kies bij <b>Toestaan</b>: <b>Lijst</b>. Typ bij <b>Bron</b> de keuzes met een <b>puntkomma</b> ertussen, bv. <b>ja;nee</b>.</p>
<p>Staan de keuzes al ergens in cellen? Dan kan je als bron ook dat <b>bereik</b> selecteren, bv. K6:N6.</p>`,
        widget: { type: 'dvSim', rounds: [
          { label: 'ik … (worden, tegenwoordige tijd)', ask: 'Maak in de cel een keuzelijst met de keuzes <b>word, wordt, werd</b> en <b>worden</b>.', goal: { type: 'list', items: ['word', 'wordt', 'werd', 'worden'] }, use: { pick: 'word' } }
        ] } },
      { id: 'l2', title: 'Alleen goede antwoorden toelaten',
        body: `<p>Bij een rekenvraag wil je alleen <b>getallen</b> toelaten. In Gegevensvalidatie kies je bij Toestaan: <b>Geheel getal</b>, en daarna een minimum en een maximum.</p>
<p>Op het tabblad <b>Foutmelding</b> typ je wat Excel moet zeggen als iemand iets anders typt, bv. "Typ een getal tussen 0 en 100."</p>`,
        widget: { type: 'dvSim', rounds: [
          { label: '6 x 7 =', ask: 'Laat in de cel alleen <b>gehele getallen tussen 0 en 100</b> toe, met een eigen foutmelding.', goal: { type: 'whole', min: 0, max: 100 }, use: { bad: 150 } }
        ] } },
      { id: 'l3', title: 'Kleuren die reageren',
        body: `<p>Met <b>voorwaardelijke opmaak</b> kleurt Excel een cel <b>vanzelf</b>, afhankelijk van wat erin staat.</p>
<p>Selecteer de cellen. Klik op het tabblad Start op <b>Voorwaardelijke opmaak</b> → <b>Markeringsregels voor cellen</b> en kies bv. <b>Kleiner dan</b>. Typ de waarde en kies een kleur.</p>`,
        widget: { type: 'cfSim', rounds: [
          { head: 'Punten', values: [8, 3, 10, 5, 9, 4], ask: 'Kleur alle punten <b>kleiner dan 5</b> met een <b>lichtrode</b> opvulling.', goal: v => v < 5 ? 'rood' : null }
        ] } },
      { id: 'l4', title: 'Groen voor juist, rood voor jammer',
        body: `<p>In je interactief werkblad kan Excel ook de feedback kleuren. Gebruik <b>Tekst die bevat…</b>: één regel voor "juist" (groen) en één regel voor "jammer" (rood).</p>
<p>Je mag meerdere regels op dezelfde cellen zetten.</p>`,
        widget: { type: 'cfSim', rounds: [
          { head: 'Feedback', values: ['juist', 'jammer', 'vul in', 'juist', 'jammer', 'juist'], ask: 'Laat <b>juist</b> groen en <b>jammer</b> rood kleuren. "vul in" blijft zonder kleur.', goal: v => v === 'juist' ? 'groen' : v === 'jammer' ? 'rood' : null }
        ] } },
      { id: 'l5', title: 'Kleuren in je rapport',
        body: `<p>In je rapport zie je zo in één oogopslag waar het goed ging en waar je nog moet oefenen.</p>
<p>Let op: 68% is voor Excel eigenlijk 0,68. Typ in de regel dus <b>50%</b> of <b>0,5</b>, niet 50.</p>
<p>Extra: neem je eigen rapport uit stap 8 erbij en geef het ook kleuren!</p>`,
        widget: { type: 'cfSim', rounds: [
          { head: 'Procent', pct: true, values: [0.68, 0.45, 0.83, 0.92, 0.5, 0.38], ask: 'Kleur procenten <b>lager dan 50%</b> rood en procenten <b>hoger dan 80%</b> groen.', goal: v => v < 0.5 ? 'rood' : v > 0.8 ? 'groen' : null, okMsg: 'Juist! Rood en groen springen meteen in het oog.' }
        ] } }
    ],
    startFile(seed, mode) {
      const d = m10data(seed, mode), b = { bold: true }, W = d.W;
      const ws = { A1: { v: 'Werkblad met keuzelijsten en kleuren', bold: true, size: 14 }, A3: { v: 'Vraag', ...b }, B3: { v: 'Oplossing', ...b }, C3: { v: 'Antwoord', ...b }, D3: { v: 'Feedback', ...b }, K3: { v: 'Keuzes', ...b } };
      const qs = [[d.rk[0], d.rk[1]], [d.ww[0], d.ww[1]], [d.hs[0], d.hs[1]], ...d.ov];
      qs.forEach((q, k) => {
        const r = 4 + k;
        ws['A' + r] = q[0]; ws['B' + r] = q[1];
        ws['D' + r] = { f: `IF(C${r}="","vul in",IF(C${r}=B${r},"${W.ok}","${W.bad}"))` };
      });
      d.hs[2].forEach((x, k) => { ws[String.fromCharCode(75 + k) + '6'] = x; });
      const rp = { A1: { v: 'Rapport dagelijks werk – ' + d.maand, bold: true, size: 14 }, A3: { v: 'Vak', ...b }, B3: { v: 'Score', ...b }, C3: { v: 'Maximum', ...b }, D3: { v: 'Procent', ...b }, A10: { v: 'Totaal', ...b } };
      d.rap.forEach((x, k) => { const r = 4 + k; rp['A' + r] = x.n; rp['B' + r] = x.s; rp['C' + r] = x.max; rp['D' + r] = { f: `B${r}/C${r}`, fmt: '0%' }; });
      rp.B10 = { f: 'SUM(B4:B9)' }; rp.C10 = { f: 'SUM(C4:C9)' }; rp.D10 = { f: 'B10/C10', fmt: '0%' };
      return { name: mode === 'test' ? 'Test10_kleuren.xlsx' : 'Module10_kleuren.xlsx', spec: { sheets: [
        { name: 'Werkblad', cells: ws, colWidths: { A: 36, B: 12, C: 14, D: 12, K: 11, L: 11, M: 11, N: 11 } },
        { name: 'Rapport', cells: rp, colWidths: { A: 14, B: 9, C: 10, D: 10 } }] } };
    },
    gen(seed, mode) {
      const d = m10data(seed, mode), W = d.W;
      const items = d.ww[2];
      const sheetOf = (wb, n) => wb.sheet(n) || wb.sheets[n === 'Rapport' ? 1 : 0];
      const dvList = (sheet, a, want, needRange) => {
        const dvs = sheet.dvAt(a);
        if (!dvs.length) {
          const elders = sheet.dataValidations.filter(x => x.type === 'list').map(x => x.sqref);
          return no(elders.length ? `Je keuzelijst staat op ${[...new Set(elders)].join(', ')}, niet op ${a}.` : `${a} heeft nog geen keuzelijst. Gebruik Gegevens → Gegevensvalidatie → Lijst.`);
        }
        const dv = dvs[dvs.length - 1];
        if (dv.type !== 'list') return no(`Bij ${a} koos je bij Toestaan niet "Lijst".`);
        const got = (sheet.dvListItems(dv) || []).map(x => norm(x));
        if (got.length === 1 && want.length > 1 && /[ ,]/.test(got[0])) return no(`Je keuzelijst heeft maar één keuze: "${got[0]}". Zet een puntkomma tussen de keuzes.`);
        const w = want.map(norm), miss = w.filter(x => !got.includes(x)), extra = got.filter(x => !w.includes(x));
        if (needRange && !sheet.dvIsRange(dv)) return no(`Gebruik als bron het bereik (de cellen ${needRange}), niet losse woorden. Klik in het vak Bron en selecteer de cellen.`);
        if (miss.length) return no(`In de keuzelijst van ${a} ontbreekt nog: ${miss.join(', ')}.`);
        if (extra.length) return no(`In de keuzelijst van ${a} staat iets wat er niet bij hoort: ${extra.join(', ')}.`);
        return ok();
      };
      const colorTask = (sheetName, range, tests) => ({ wb }) => {
        const sh = sheetOf(wb, sheetName);
        const cells = []; eachCell(range, a => cells.push(a));
        if (!sh.conditionalFormats.length) return no(`Er staat nog geen voorwaardelijke opmaak op ${range}${sheetName === 'Rapport' ? ' (blad Rapport)' : ''}.`);
        const notCovered = cells.filter(a => !sh.cfAt(a).length);
        if (notCovered.length === cells.length) return no(`Je regel staat op ${[...new Set(sh.conditionalFormats.map(c => c.sqref))].join(', ')}. Ze moet op ${range} staan.`);
        if (notCovered.length) return no(`Je regel geldt nog niet voor ${list(notCovered.slice(0, 3))}. Selecteer ${range} helemaal.`);
        if (tests.some(t => typeof t.v === 'number') && cells.some(a => sh.cfAt(a).some(r => r.type === 'cellIs' && r.formulas.some(f => numOf(f) != null && numOf(f) >= 1)))) return no('Je procenten zijn eigenlijk getallen tussen 0 en 1. Typ in de regel 50% of 0,5, niet 50.');
        for (const t of tests) {
          for (const a of cells) {
            const st = sh.cfStyle(a, t.v);
            const cols = [...colorNames(st.fill), ...colorNames(st.font)];
            const shown = typeof t.v === 'number' ? pctTxt(t.v) : `"${t.v}"`;
            if (t.want && !cols.includes(t.want)) {
              if (!cols.length) {
                return no(`Bij ${shown} krijgt ${a} nog geen kleur. Het moet ${t.want} worden.`);
              }
              return no(`Bij ${shown} wordt ${a} ${cols[0]}. Het moet ${t.want} worden.`);
            }
            if (t.not && cols.includes(t.not)) return no(`Bij ${shown} wordt ${a} ook ${t.not}. Dat mag niet: kijk je regel na (${typeof t.v === 'number' ? 'kleiner dan of groter dan?' : 'welke tekst?'}).`);
          }
        }
        return ok();
      };
      return { startFile: true, tasks: [
        { id: 't1', text: `Maak in <span class="addr">C5</span> een <b>keuzelijst</b> met de keuzes ${items.map(x => `<b>${x}</b>`).join(', ')}.`,
          hints: ['Tabblad Gegevens → Gegevensvalidatie → Toestaan: Lijst.', `Typ bij Bron: ${items.join(';')}`],
          check: ({ wb }) => dvList(sheetOf(wb, 'Werkblad'), 'C5', items, null) },
        { id: 't2', text: `Maak in <span class="addr">C6</span> een keuzelijst met als bron het <b>bereik K6:N6</b>.`,
          hints: ['Ook hier: Gegevensvalidatie → Lijst.', 'Klik in het vak Bron en selecteer met de muis de cellen K6 tot N6.'],
          check: ({ wb }) => dvList(sheetOf(wb, 'Werkblad'), 'C6', d.hs[2], 'K6:N6') },
        { id: 't3', text: `Laat in <span class="addr">C4</span> alleen <b>gehele getallen tussen 0 en 1000</b> toe, met een <b>eigen foutmelding</b>.`,
          hints: ['Gegevensvalidatie → Toestaan: Geheel getal, tussen 0 en 1000.', 'Typ op het tabblad Foutmelding een bericht, bv. "Typ een getal."'],
          check: ({ wb }) => {
            const sh = sheetOf(wb, 'Werkblad'), dv = sh.dvAt('C4')[0];
            if (!dv) return no('C4 heeft nog geen gegevensvalidatie.');
            if (dv.type === 'decimal') return no('Kies Geheel getal in plaats van Decimaal: er horen geen kommagetallen bij deze vraag.');
            if (dv.type !== 'whole') return no('Kies bij Toestaan: Geheel getal.');
            if (dv.operator !== 'between') return no('Kies bij Gegevens: tussen.');
            if (numOf(dv.formula1) !== 0 || numOf(dv.formula2) !== 1000) return no(`Het minimum moet 0 zijn en het maximum 1000. Nu is het ${dv.formula1 || '?'} en ${dv.formula2 || '?'}.`);
            if (!dv.showErrorMessage || !(dv.error || '').trim()) return no('Typ op het tabblad Foutmelding een eigen foutbericht.');
            return ok();
          } },
        { id: 't4', text: `Laat de feedback in <span class="addr">D4:D9</span> <b>groen</b> kleuren als er <b>"${W.ok}"</b> staat.`,
          hints: ['Selecteer D4:D9. Start → Voorwaardelijke opmaak → Markeringsregels voor cellen.', `Kies "Tekst die bevat…", typ ${W.ok} en kies "Groene opvulling met donkergroene tekst".`],
          check: colorTask('Werkblad', 'D4:D9', [{ v: W.ok, want: 'groen' }, { v: W.bad, not: 'groen' }, { v: 'vul in', not: 'groen' }]) },
        { id: 't5', text: `Laat de feedback in <span class="addr">D4:D9</span> <b>rood</b> kleuren als er <b>"${W.bad}"</b> staat.`,
          hints: ['Een tweede regel op dezelfde cellen.', `"Tekst die bevat…", typ ${W.bad} en kies "Lichtrode opvulling met donkerrode tekst".`],
          check: colorTask('Werkblad', 'D4:D9', [{ v: W.bad, want: 'rood' }, { v: W.ok, not: 'rood' }, { v: 'vul in', not: 'rood' }]) },
        { id: 't6', text: `Ga naar het blad <b>Rapport</b>. Kleur de procenten in <span class="addr">D4:D9</span> die <b>lager dan ${pctTxt(d.lo)}</b> zijn <b>rood</b>.`,
          hints: ['Selecteer D4:D9 op het blad Rapport. Markeringsregels voor cellen → Kleiner dan.', `Typ ${pctTxt(d.lo)} (of ${nlNum(d.lo)}) en kies de lichtrode opvulling.`],
          check: colorTask('Rapport', 'D4:D9', [{ v: d.lo - 0.05, want: 'rood' }, { v: d.lo + 0.05, not: 'rood' }]) },
        { id: 't7', text: `Kleur op het blad Rapport de procenten <b>hoger dan ${pctTxt(d.hi)}</b> <b>groen</b>.`,
          hints: ['Een tweede regel op D4:D9: Groter dan.', `Typ ${pctTxt(d.hi)} (of ${nlNum(d.hi)}) en kies de groene opvulling.`],
          check: colorTask('Rapport', 'D4:D9', [{ v: d.hi + 0.05, want: 'groen' }, { v: d.hi - 0.05, not: 'groen' }, { v: d.lo - 0.05, not: 'groen' }]) },
        { id: 't8', text: `Test je werkblad: beantwoord de vragen in <span class="addr">C4</span>, <span class="addr">C5</span> en <span class="addr">C6</span> juist. Gebruik de keuzelijsten. Wordt de feedback groen?`,
          hints: ['Klik op het pijltje naast C5 en C6.', 'De oplossing staat in kolom B.'],
          check: ({ wb }) => {
            const sh = sheetOf(wb, 'Werkblad');
            for (const r of [4, 5, 6]) {
              const v = sh.value('C' + r);
              if (v == null || String(v).trim() === '') return no(`C${r} is nog leeg.`);
              if (norm(showV(sh.calc('D' + r))) !== norm(W.ok)) return no(`In C${r} staat "${showV(v)}", maar dat is niet het juiste antwoord.`);
            }
            return ok();
          } }
      ] };
    }
  };

  // ======================================================================
  // MODULE 11 – Klaar voor een gebruiker
  // ======================================================================
  function m11data(seed, mode) {
    const r = rng(seed + '|m11|' + mode);
    const test = mode === 'test';
    const nums = r.shuffle(REKEN).slice(0, 2), txt = r.shuffle([...HOOFDSTEDEN.map(x => [x[0], x[1]]), ...OVERIG.filter(x => typeof x[1] === 'string')]).slice(0, 4);
    const q = r.shuffle([...nums, ...txt]);
    const ti = q.findIndex(x => typeof x[1] === 'string'), ni = q.findIndex(x => typeof x[1] === 'number');
    return { q, ti, ni, W: test ? { ok: 'goed', bad: 'fout', empty: 'vul in' } : { ok: 'juist', bad: 'jammer', empty: 'vul in' } };
  }
  /* SPATIES.WISSEN-controle: werkt de formule ook met een spatie te veel, en ook bij getallen? */
  function trimCheck(sheet, a, ans, sol, W) {
    const base = ifAnswerCheck(sheet, a, ans, sol, W, true);
    if (!base.ok) {
      const s = sheet.calc(sol);
      if (isNum(s) && FX.info(sheet.formula(a) || '', 'en').funcs.includes('TRIM') && said(tryWith(sheet, a, { [ans]: s })) !== norm(W.ok))
        return no(`In ${a} geeft een juist getal (${showV(s)}) nu "${said(tryWith(sheet, a, { [ans]: s }))}". SPATIES.WISSEN maakt van een getal tekst, en "${showV(s)}" als tekst is niet hetzelfde als het getal ${showV(s)}. Zet SPATIES.WISSEN rond allebei: SPATIES.WISSEN(${ans})=SPATIES.WISSEN(${sol})`);
      return base;
    }
    const f = sheet.formula(a), s = sheet.calc(sol);
    if (!FX.info(f, 'en').funcs.includes('TRIM')) return no(`Gebruik SPATIES.WISSEN in ${a}, zodat een spatie te veel geen "${W.bad}" geeft.`);
    const withSpace = said(tryWith(sheet, a, { [ans]: showV(s) + ' ' }));
    if (withSpace !== norm(W.ok)) return no(`Typt de gebruiker "${showV(s)} " (met een spatie erachter), dan toont ${a} "${withSpace}". Kijk na waar je SPATIES.WISSEN zette.`);
    return ok();
  }

  const m11 = {
    id: 'm11', nr: 11,
    title: 'Klaar voor een gebruiker',
    short: 'Je maakt je werkblad klaar voor iemand anders: de score tellen, spaties wegwerken, een afbeelding, oplossingen verbergen en het blad beveiligen.',
    skills: ['AANTAL.ALS', 'SPATIES.WISSEN', 'score in procent', 'afbeelding', 'kolom verbergen', 'rasterlijnen', 'cellen ontgrendelen', 'blad beveiligen'],
    lessons: [
      { id: 'l1', title: 'De score tellen',
        body: `<p>Hoeveel antwoorden zijn juist? Dat telt Excel met <b>AANTAL.ALS</b>: het telt hoe vaak iets voorkomt in een bereik.</p>
<p><b>=AANTAL.ALS(D2:D6;"juist")</b> telt hoe vaak "juist" in D2:D6 staat.</p>`,
        widget: { type: 'sheetSim', rows: 8, cols: 4, widths: [110, 20, 20, 80], cells: { D1: 'Feedback', D2: 'juist', D3: 'jammer', D4: 'juist', D5: 'vul in', D6: 'juist', A8: 'Aantal juist' }, rounds: [
          { select: 'B8', ask: 'Tel in <span class="addr">B8</span> hoeveel keer "juist" in D2:D6 staat.',
            check: (s, ch) => {
              if (ch !== 'B8') return null;
              const i = s.info('B8');
              if (!i) return no('Een formule begint met =.');
              if (i.error === '#SEP') return no('Tussen de delen zet je een puntkomma: =AANTAL.ALS(D2:D6;"juist")');
              if (!i.funcs.includes('COUNTIF')) return no(i.funcs.includes('COUNTA') ? 'AANTALARG telt alle gevulde cellen. Gebruik AANTAL.ALS.' : 'Gebruik de functie AANTAL.ALS.');
              return s.num('B8') === 3 ? ok('Juist: 3 keer "juist".') : no(`Je formule geeft ${s.show('B8')}. Er staat 3 keer "juist".`);
            } },
          { select: 'D3', ask: 'De gebruiker verbetert zijn antwoord: typ <b>juist</b> in <span class="addr">D3</span>. Kijk naar de score.',
            check: (s, ch) => ch !== 'D3' ? null : s.num('B8') === 4 ? ok('De score telt vanzelf mee: 4.') : no('Typ juist in D3.') }
        ] } },
      { id: 'l2', title: 'Spaties weghalen',
        body: `<p>Typt de gebruiker per ongeluk een spatie achter zijn antwoord, dan zegt je werkblad "jammer". Niet eerlijk! <b>SPATIES.WISSEN</b> haalt overbodige spaties weg.</p>
<p>Maar let op: SPATIES.WISSEN maakt van alles <b>tekst</b>, ook van een getal. Voor Excel is de tekst "42" niet hetzelfde als het getal 42. Daarom zet je SPATIES.WISSEN rond <b>allebei</b>:</p>
<p><b>=ALS(SPATIES.WISSEN(C2)=SPATIES.WISSEN(B2);"juist";"jammer")</b></p>`,
        widget: { type: 'sheetSim', rows: 3, cols: 4, widths: [150, 80, 80, 90], cells: { A1: 'Vraag', B1: 'Oplossing', C1: 'Antwoord', D1: 'Feedback', A2: 'Hoofdstad van Italië?', B2: 'Rome', C2: 'Rome ', D2: '=ALS(C2=B2;"juist";"jammer")', A3: '6 x 7 =', B3: 42, C3: '42 ', D3: '=ALS(C3=B3;"juist";"jammer")' }, rounds: [
          { select: 'D2', ask: 'D2 zegt "jammer", maar Rome is juist: er staat een spatie achter. Verbeter D2 met <b>SPATIES.WISSEN</b>.',
            check: (s, ch) => {
              if (ch !== 'D2') return null;
              const i = s.info('D2');
              if (!i || !i.funcs.includes('TRIM')) return no('Gebruik SPATIES.WISSEN rond C2.');
              return norm(s.show('D2')) === 'juist' ? ok('Juist! De spatie telt niet meer mee.') : no(`D2 toont "${s.show('D2')}". Kijk je formule na.`);
            } },
          { select: 'D3', ask: 'Ook in C3 staat een spatie te veel. Verbeter <span class="addr">D3</span> op dezelfde manier. Wat gebeurt er?',
            check: (s, ch) => {
              if (ch !== 'D3') return null;
              const i = s.info('D3');
              if (!i || !i.funcs.includes('TRIM')) return no('Gebruik SPATIES.WISSEN.');
              if (norm(s.show('D3')) === 'juist') return ok('Juist! Met SPATIES.WISSEN aan beide kanten werkt het ook voor getallen.');
              return no('Nog altijd "jammer"! SPATIES.WISSEN maakt van 42 de tekst "42", en die is voor Excel niet gelijk aan het getal 42. Zet SPATIES.WISSEN rond allebei: =ALS(SPATIES.WISSEN(C3)=SPATIES.WISSEN(B3);"juist";"jammer")');
            } }
        ] } },
      { id: 'l3', title: 'Een afbeelding',
        body: `<p>Een afbeelding maakt je werkblad aantrekkelijker, bv. een prentje bij een vraag of een logo naast de titel.</p>`,
        widget: { type: 'steps', steps: [
          'Klik op het tabblad <b>Invoegen</b> → <b>Afbeeldingen</b>.',
          'Kies <b>Dit apparaat</b> (een afbeelding op je computer) of <b>Stockafbeeldingen</b> (afbeeldingen van Excel zelf).',
          'De afbeelding <b>zweeft</b> boven de cellen. Sleep ze naar de juiste plaats en maak ze kleiner met de <b>hoekjes</b> (dan blijft ze mooi in verhouding).',
          'In Excel 365 kan het ook <b>in een cel</b>: Invoegen → Afbeeldingen → <b>Plaatsen in cel</b>. De afbeelding wordt dan even groot als de cel.',
          'Denk aan de gebruiker: een afbeelding moet <b>helpen</b>, niet de vragen bedekken.'
        ], confirm: 'Ik weet hoe ik een afbeelding invoeg' } },
      { id: 'l4', title: 'Oplossingen verbergen',
        body: `<p>De gebruiker mag de oplossingen in kolom B niet zien. Rechtsklik op de kolomletter <b>B</b> en kies <b>Verbergen</b>. De formules werken gewoon verder.</p>
<p>Zonder de grijze <b>rasterlijnen</b> lijkt je werkblad meer op een echt werkblad: tabblad <b>Beeld</b> → vinkje <b>Rasterlijnen</b> weg.</p>`,
        widget: { type: 'quiz', questions: [
          { q: 'Hoe verberg je kolom B?', options: ['Rechtsklik op de letter B → Verbergen', 'Kolom B leegmaken', 'De letters in kolom B wit kleuren'], correct: 0, explain: 'Leegmaken kan niet: dan zijn de oplossingen weg. Wit kleuren is te makkelijk te zien als je de cel aanklikt.' },
          { q: 'Hoe maak je kolom B weer zichtbaar?', options: ['Selecteer kolom A tot C, rechtsklik → Zichtbaar maken', 'Dat kan niet meer', 'Druk een week later op Ctrl + Z'], correct: 0, explain: 'Een verborgen kolom is dus niet echt geheim…' },
          { q: 'Een slimme gebruiker kan een verborgen kolom zichtbaar maken. Hoe voorkom je dat?', options: ['Het blad beveiligen', 'Een kleinere letter gebruiken', 'Het bestand een andere naam geven'], correct: 0, explain: 'Op een beveiligd blad kan je geen kolommen zichtbaar maken. Dat zie je in de volgende les.' }
        ] } },
      { id: 'l5', title: 'Beveiligen',
        body: `<p>Op een beveiligd blad kan de gebruiker niets kapotmaken: geen formule wissen, geen kolom zichtbaar maken. Alleen de cellen die jij <b>ontgrendelt</b>, kan hij invullen.</p>
<p><b>1.</b> Selecteer de antwoordcellen → rechtsklik → <b>Celeigenschappen</b> (of <span class="key">Ctrl</span> + <span class="key">1</span>) → tabblad <b>Bescherming</b> → vinkje <b>Geblokkeerd</b> weg.<br>
<b>2.</b> Tabblad <b>Controleren</b> → <b>Blad beveiligen</b> → OK. Een wachtwoord mag, maar vergeet het niet!</p>
<p>De volgorde is belangrijk: eerst ontgrendelen, dan beveiligen. Op een beveiligd blad kan je niets meer ontgrendelen.</p>`,
        widget: { type: 'protectSim', rows: 4, cols: 4, open: ['C2', 'C3', 'C4'], ask: 'Zorg dat de gebruiker alleen in <span class="addr">C2:C4</span> kan typen: ontgrendel eerst C2:C4, en beveilig daarna het blad.',
          cells: { A1: 'Vraag', B1: 'Oplossing', C1: 'Antwoord', D1: 'Feedback', A2: '6 x 7 =', B2: 42, D2: 'vul in', A3: 'Hoofdstad van Italië?', B3: 'Rome', D3: 'vul in', A4: 'Meervoud van kind?', B4: 'kinderen', D4: 'vul in' } } }
    ],
    startFile(seed, mode) {
      const d = m11data(seed, mode), b = { bold: true }, W = d.W;
      const cells = { A1: { v: 'Mijn werkblad – klaar voor een gebruiker', bold: true, size: 14 }, A3: { v: 'Vraag', ...b }, B3: { v: 'Oplossing', ...b }, C3: { v: 'Antwoord', ...b }, D3: { v: 'Feedback', ...b },
        A11: { v: 'Aantal juist', ...b }, A12: { v: 'Score', ...b } };
      d.q.forEach((x, k) => {
        const r = 4 + k;
        cells['A' + r] = x[0]; cells['B' + r] = x[1];
        cells['D' + r] = { f: `IF(C${r}="","${W.empty}",IF(C${r}=B${r},"${W.ok}","${W.bad}"))` };
      });
      cells['C' + (4 + d.ti)] = d.q[d.ti][1] + ' ';           // een juist antwoord met een spatie te veel
      return { name: mode === 'test' ? 'Test11_gebruiker.xlsx' : 'Module11_gebruiker.xlsx', spec: { sheetName: 'Werkblad', cells, colWidths: { A: 34, B: 12, C: 13, D: 12 } } };
    },
    gen(seed, mode) {
      const d = m11data(seed, mode), W = d.W, D = ['D4', 'D5', 'D6', 'D7', 'D8', 'D9'];
      return { startFile: true, tasks: [
        { id: 't1', text: `Tel in <span class="addr">B11</span> hoeveel keer <b>"${W.ok}"</b> in D4:D9 staat, met <b>AANTAL.ALS</b>.`,
          hints: ['AANTAL.ALS heeft twee delen: het bereik en wat je telt.', `Typ =AANTAL.ALS(D4:D9;"${W.ok}")`],
          check: ({ sheet }) => {
            const f = sheet.formula('B11');
            if (!f) { const r0 = checkFormula(sheet, 'B11', {}); return r0.ok ? no('Typ in B11 een formule met AANTAL.ALS.') : r0; }
            const inf = FX.info(f, 'en');
            if (FX.isErr(sheet.calc('B11'))) return checkFormula(sheet, 'B11', {});
            if (!inf.funcs.includes('COUNTIF')) return no(inf.funcs.includes('COUNTA') ? 'AANTALARG telt alle gevulde cellen, ook "jammer". Gebruik AANTAL.ALS.' : 'Gebruik de functie AANTAL.ALS.');
            const miss = D.filter(a => !inf.cells.includes(a));
            if (miss.length) return no(`Je bereik telt ${list(miss)} niet mee. Gebruik D4:D9.`);
            const all = {}, none = {}, three = {};
            D.forEach((a, k) => { all[a] = W.ok; none[a] = W.bad; three[a] = k % 2 ? W.ok : W.empty; });
            const t = o => numOf(tryWith(sheet, 'B11', o));
            if (t(all) !== 6 || t(none) !== 0 || t(three) !== 3) return no(`Je formule telt niet hoe vaak "${W.ok}" voorkomt. Kijk na wat er tussen de aanhalingstekens staat.`);
            return ok();
          } },
        { id: 't2', text: `Toon in <span class="addr">B12</span> de <b>score als procent</b>: het aantal juiste antwoorden gedeeld door het aantal vragen (6).`,
          hints: ['Delen doe je met /. Daarna de knop %.', 'Typ =B11/6 en klik op %.'],
          check: ({ sheet }) => {
            const f = sheet.formula('B12');
            if (!f) { const r0 = checkFormula(sheet, 'B12', {}); return r0.ok ? no('Typ in B12 een formule.') : r0; }
            if (!FX.info(f, 'en').cells.includes('B11')) return no('Gebruik het aantal juist uit B11 in je formule.');
            const v = numOf(tryWith(sheet, 'B12', { B11: 3 }));
            if (v == null || Math.abs(v - 0.5) > 1e-9) return no(`Je formule rekent niet goed: bij 3 juist moet het 50% zijn, nu is het ${v == null ? 'geen getal' : pctTxt(v)}.`);
            return sheet.numFmt('B12').kind === 'percent' ? ok() : no('B12 rekent goed. Toon het nog als procent met de knop %.');
          } },
        { id: 't3', text: `Verbeter de formules in <span class="addr">D4:D9</span> met <b>SPATIES.WISSEN</b>, zodat een spatie te veel geen "${W.bad}" meer geeft. Het moet ook werken bij getallen!`,
          hints: ['Zet SPATIES.WISSEN rond het antwoord én rond de oplossing.', `=ALS(C4="";"${W.empty}";ALS(SPATIES.WISSEN(C4)=SPATIES.WISSEN(B4);"${W.ok}";"${W.bad}")) en doortrekken.`],
          check: ({ sheet }) => {
            const first = trimCheck(sheet, 'D4', 'C4', 'B4', W);
            if (!first.ok) return first;
            for (let r = 5; r <= 9; r++) { const res = trimCheck(sheet, 'D' + r, 'C' + r, 'B' + r, W); if (!res.ok) return /^In D\d+ geeft een juist getal/.test(res.msg) ? res : no(`D4 is goed! Trek de verbeterde formule door tot D9: D${r} werkt nog niet zoals D4.`); }
            return ok();
          } },
        { id: 't4', text: 'Voeg een <b>afbeelding</b> toe die past bij je werkblad.',
          hints: ['Tabblad Invoegen → Afbeeldingen.', 'Kies Stockafbeeldingen of een afbeelding van je computer. Zet ze naast de titel.'],
          check: ({ sheet }) => { const im = sheet.imageInfo(); return im.floating || im.inCell.length ? ok() : no('Ik vind nog geen afbeelding op je werkblad.'); } },
        { id: 't5', text: 'Verberg kolom <b>B</b> met de oplossingen.',
          hints: ['Dat doe je bij de kolomletter.', 'Rechtsklik op de letter B en kies Verbergen.'],
          check: ({ sheet }) => {
            if (sheet.colHidden(2)) { const andere = [1, 3, 4].filter(c => sheet.colHidden(c)).map(numToCol); return andere.length ? no(`Je verborg ook kolom ${list(andere)}. Die moet de gebruiker kunnen zien.`) : ok(); }
            const andere = [1, 3, 4, 5].filter(c => sheet.colHidden(c)).map(numToCol);
            return no(andere.length ? `Je verborg kolom ${list(andere)}. Verberg kolom B: daar staan de oplossingen.` : 'Kolom B is nog zichtbaar.');
          } },
        { id: 't6', text: 'Zet de <b>rasterlijnen</b> uit.',
          hints: ['Dat vind je op het tabblad Beeld.', 'Tabblad Beeld → haal het vinkje bij Rasterlijnen weg.'],
          check: ({ sheet }) => sheet.showGrid ? no('De rasterlijnen zijn nog zichtbaar.') : ok() },
        { id: 't7', text: 'Ontgrendel de antwoordcellen <span class="addr">C4:C9</span>. Alle andere cellen blijven geblokkeerd.',
          hints: ['Selecteer C4:C9 en open Celeigenschappen (Ctrl + 1).', 'Tabblad Bescherming → haal het vinkje bij Geblokkeerd weg → OK.'],
          check: ({ sheet }) => {
            const C = ['C4', 'C5', 'C6', 'C7', 'C8', 'C9'];
            const still = C.filter(a => sheet.style(a).locked);
            if (still.length === 6) return no('De antwoordcellen zijn nog geblokkeerd.');
            if (still.length) return no(`${list(still)} ${still.length > 1 ? 'zijn' : 'is'} nog geblokkeerd. Selecteer C4:C9 helemaal.`);
            const open = ['B4', 'D4', 'A4', 'B11'].filter(a => !sheet.style(a).locked);
            if (open.length) return no(`Ook ${list(open)} ${open.length > 1 ? 'zijn' : 'is'} ontgrendeld. Dan kan de gebruiker daar iets veranderen. Ontgrendel alleen C4:C9.`);
            return ok();
          } },
        { id: 't8', text: '<b>Beveilig</b> het blad.',
          hints: ['Dat vind je op het tabblad Controleren.', 'Controleren → Blad beveiligen → OK. Gebruik je een wachtwoord, schrijf het dan op.'],
          check: ({ sheet }) => {
            if (!sheet.isProtected) return no('Het blad is nog niet beveiligd.');
            if (['C4', 'C5', 'C6'].every(a => sheet.style(a).locked)) return no('Het blad is beveiligd, maar de antwoordcellen zijn nog geblokkeerd. De gebruiker kan dan niets invullen! Hef de beveiliging op, ontgrendel C4:C9 en beveilig opnieuw.');
            return ok();
          } }
      ] };
    }
  };

  // ======================================================================
  // MODULE 12 – Project: mijn interactief werkblad
  // ======================================================================
  const NIVEAUS = [
    { n: 1, stars: '★', name: 'Basis', short: 'Een werkblad dat werkt.',
      items: ['een titel en een uitleg', 'minstens 10 vragen: rekenen én taal of WO', 'feedback met ALS, ook als de cel leeg is', 'een score met AANTAL.ALS', 'groen bij juist, rood bij jammer', 'een vraag met een afbeelding', 'de oplossingen verborgen'] },
    { n: 2, stars: '★★', name: 'Gevorderd', short: 'Alles van ★, en klaar voor elke gebruiker.',
      items: ['SPATIES.WISSEN', 'minstens 2 keuzelijsten', 'de score als procent', 'geen rasterlijnen', 'een beveiligd blad'] },
    { n: 3, stars: '★★★', name: 'Expert', short: 'Alles van ★★, en nog slimmer.',
      items: ['een eindboodschap die afhangt van de score', 'een score die kleurt', 'een rekenvraag met een eigen foutmelding', 'een keuzelijst met de keuzes uit een bereik'] }
  ];
  const levelOf = seed => { const m = /^([123])~/.exec(String(seed || '')); return m ? +m[1] : 1; };
  const withLevel = (lvl, s) => lvl + '~' + (s || newSeed());

  /* Wat bewaart Excel als de gebruiker dit typt? Een getal wordt een getal (ook met spaties of een komma). */
  function typedAs(v) {
    if (isNum(v)) return v;
    const s = String(v).trim();
    if (/^-?\d+(,\d+)?$/.test(s)) return parseFloat(s.replace(',', '.'));
    return String(v);
  }
  const JUNK = 'qzx', JUNK2 = 'wvb';
  const isErrV = v => FX.isErr(v);
  const blank = v => v == null || String(v).trim() === '';
  const rowOf = a => parseAddr(a).r, colOf = a => parseAddr(a).c;
  const cut = (s, n) => { s = String(s); return s.length > (n || 30) ? s.slice(0, (n || 30) - 1) + '…' : s; };
  const nlFn = fn => FX.EN_NL[fn] || fn;
  const colsOfStyle = st => [...colorNames(st.fill), ...colorNames(st.font)];
  const hiddenAt = (sh, a) => sh.state !== 'visible' || sh.colHidden(colOf(a)) || sh.rowHidden(rowOf(a));
  const lists = (sh, a) => sh.dvAt(a).filter(d => d.type === 'list');

  /* Zoekt de vragen op één blad: feedbackcel (ALS) + antwoordcel + juist antwoord. */
  function analyzeSheet(sh) {
    const wb = sh.wb;
    const res = { sheet: sh, qs: [], unresolved: [], unknown: [], helpers: [] };
    const byAns = {};
    for (const F of Object.keys(sh.cells).filter(a => sh.formula(a))) {
      const inf = FX.info(sh.formula(F), 'en');
      if (!inf.funcs.includes('IF') && !inf.funcs.includes('IFS')) continue;
      const own = [...new Set(inf.refs.filter(r => !r.sheet).map(r => r.a))].filter(a => a !== F && !sh.formula(a));
      if (!own.length) continue;                           // bv. een eindboodschap die naar de score kijkt
      const unk = inf.funcs.filter(x => !FX.EN_NL[x]);
      if (unk.length) { res.unknown.push({ F, fn: unk[0] }); continue; }
      const partners = inf.refs.map(r => ({ a: r.a, sh: r.sheet ? (wb && wb.sheet(r.sheet)) : sh })).filter(r => r.sh);
      let best = null, questionLike = false;
      for (const A of own.slice(0, 8)) {
        const sim = v => sh.calcWith(F, { [A]: v });
        const oNull = sim(null), oJunk = sim(JUNK);
        if (isErrV(oJunk) || said(oJunk) !== said(sim(JUNK2))) continue;   // geen vraag (bv. "Dag " & naam)
        questionLike = true;
        const cands = [];
        partners.filter(p => !(p.a === A && p.sh === sh)).forEach(p => cands.push({ v: p.sh.calc(p.a), ref: p }));
        inf.numbers.forEach(v => cands.push({ v }));
        inf.strings.forEach(v => cands.push({ v }));
        let hit = null;
        const tryIt = (c, t, flag) => {
          const o = sim(t);
          if (!isErrV(o) && said(o) !== said(oJunk) && said(o) !== said(oNull)) hit = { sol: c.v, typed: t, solRef: c.ref || null, okRaw: o, textNumber: !!flag };
        };
        for (const c of cands) { if (blank(c.v) || isErrV(c.v) || typeof c.v === 'boolean') continue; tryIt(c, typedAs(c.v)); if (hit) break; }
        if (!hit) for (const c of cands) {    // een getal dat als tekst in de oplossing staat ('42)
          if (typeof c.v === 'string' && isNum(typedAs(c.v))) { tryIt(c, c.v, true); if (hit) break; }
        }
        if (!hit) continue;
        let score = 0;
        if (!sh.style(A).locked) score += 2;
        if (!hiddenAt(sh, A)) score += 1;
        if (hit.solRef && hiddenAt(hit.solRef.sh, hit.solRef.a)) score += 1;
        if (said(oNull) !== said(oJunk)) score += 2;
        if (lists(sh, A).length) score += 2;
        if (blank(sh.value(A))) score += 1;
        if (hit.solRef && hit.solRef.sh === sh && colOf(A) > colOf(hit.solRef.a)) score += 0.5;
        if (!best || score > best.score) best = Object.assign(hit, { score, A, oNull, oJunk });
      }
      if (!best) {
        if (questionLike) res.unresolved.push({ F, trim: inf.funcs.includes('TRIM'), numPartner: partners.some(p => isNum(p.sh.calc(p.a))) || inf.numbers.length > 0 });
        continue;
      }
      const q = { F, ans: best.A, sol: best.sol, typed: best.typed, solRef: best.solRef, textNumber: best.textNumber,
        okRaw: best.okRaw, okW: said(best.okRaw), badRaw: best.oJunk, badW: said(best.oJunk), emptyRaw: best.oNull, emptyW: isErrV(best.oNull) ? '#' : said(best.oNull),
        hasList: lists(sh, best.A).length > 0 };
      (byAns[q.ans] = byAns[q.ans] || []).push(q);
    }
    Object.values(byAns).forEach(arr => {
      arr.sort((x, y) => (numOf(x.okRaw) == null ? 0 : 1) - (numOf(y.okRaw) == null ? 0 : 1));   // woorden ("juist") eerst, punten (1/0) als hulpcel
      res.qs.push(arr[0]); res.helpers.push(...arr.slice(1));
    });
    res.qs.sort((x, y) => rowOf(x.ans) - rowOf(y.ans) || colOf(x.ans) - colOf(y.ans));
    return res;
  }
  /* Het blad met de meeste vragen (bij gelijkspel: het blad met de meeste lege antwoordcellen, zoals een blad "Blanco"). */
  function analyze(wb) {
    if (wb._m12) return wb._m12;
    const all = wb.sheets.map(analyzeSheet);
    const empties = r => r.qs.filter(q => blank(r.sheet.value(q.ans))).length;
    const pick = all.slice().sort((x, y) => y.qs.length - x.qs.length || empties(y) - empties(x) || (x.sheet.state === 'visible' ? 0 : 1) - (y.sheet.state === 'visible' ? 0 : 1))[0];
    const an = pick && pick.qs.length ? pick : (all.find(r => r.sheet === wb.sheet()) || all[0] || { sheet: wb.sheet(), qs: [], unresolved: [], unknown: [], helpers: [] });
    an.unresolved = an.qs.length ? an.unresolved : all.flatMap(r => r.unresolved.map(u => Object.assign({ sheet: r.sheet }, u)));
    an.unknown = an.qs.length ? an.unknown : all.flatMap(r => r.unknown);
    wb._m12 = an;
    return an;
  }
  /* De gebruiker naspelen: alle antwoordcellen tegelijk invullen. */
  const fillIn = (qs, pick) => { const o = {}; qs.forEach((q, i) => { const s = pick(q, i); o[q.ans] = s === 'ok' ? q.typed : s === 'bad' ? JUNK : null; }); return o; };
  function scoreInfo(an) {
    if (an._score) return an._score;
    const sh = an.sheet, qs = an.qs, n = qs.length;
    const skip = new Set([...qs.map(q => q.F), ...an.helpers.map(q => q.F)]);
    const ALL = fillIn(qs, () => 'ok'), NONE = fillIn(qs, () => 'bad'), EMPTY = fillIn(qs, () => 'empty');
    const half = Math.ceil(n / 2), HALF = fillIn(qs, (q, i) => i < half ? 'ok' : 'bad');
    const cells = Object.keys(sh.cells).filter(a => sh.formula(a) && !skip.has(a));
    const list = [];
    for (const a of cells) {
      const vA = sh.calcWith(a, ALL), vN = sh.calcWith(a, NONE), vE = sh.calcWith(a, EMPTY);
      if (!isNum(vA) || !isNum(vN) || !(vA > vN)) continue;
      const cover = qs.filter((q, i) => { const v = sh.calcWith(a, fillIn(qs, (x, j) => j === i ? 'ok' : 'bad')); return isNum(v) && v > vN; });
      list.push({ a, vA, vN, vE, vH: sh.calcWith(a, HALF), cover, countif: FX.info(sh.formula(a), 'en').funcs.includes('COUNTIF') });
    }
    list.sort((x, y) => y.cover.length - x.cover.length || (y.countif ? 1 : 0) - (x.countif ? 1 : 0) || (Number.isInteger(y.vA) ? 1 : 0) - (Number.isInteger(x.vA) ? 1 : 0));
    const pct = list.find(x => Math.abs(x.vA - 1) < 1e-9 && x.cover.length === n) || null;
    const msgs = [];
    for (const a of cells) {
      if (list.some(x => x.a === a) || !FX.info(sh.formula(a), 'en').funcs.some(f => f === 'IF' || f === 'IFS')) continue;
      const oA = sh.calcWith(a, ALL), oN = sh.calcWith(a, NONE);
      if (isErrV(oA) || isErrV(oN) || blank(oA) || blank(oN)) continue;
      msgs.push({ a, oA, oN, differs: said(oA) !== said(oN) });
    }
    an._score = { list, best: list[0] || null, pct, msgs, ALL, NONE, half, cells };
    return an._score;
  }
  function hasQuestion(sh, q, an) {
    const p = parseAddr(q.ans), others = new Set([...an.qs.map(x => x.ans), ...an.qs.map(x => x.F)]);
    for (let c = 1; c < p.c; c++) {
      const a = addr(c, p.r);
      if (q.solRef && q.solRef.sh === sh && q.solRef.a === a) continue;
      if (others.has(a) || sh.colHidden(c)) continue;
      if (!sh.isEmpty(a) || (sh.cell(a) && sh.cell(a).vm != null)) return true;
    }
    const up = addr(p.c, p.r - 1);
    return p.r > 1 && !others.has(up) && !sh.isEmpty(up);
  }
  function unresolvedMsg(an) {
    const u = an.unresolved[0];
    if (!u) return '';
    if (u.trim && u.numPartner) return ` In ${u.F} werkt het juiste antwoord niet: SPATIES.WISSEN maakt van een getal tekst. Zet SPATIES.WISSEN rond het antwoord én rond de oplossing.`;
    return ` In ${u.F} staat een ALS-formule, maar ik vind geen antwoord dat "juist" geeft. Probeer het zelf: typ het juiste antwoord. Klopt je oplossing?`;
  }
  function dvNum(sh, f) {
    const s = String(f || '').trim().replace(/^=/, '');
    if (s === '') return null;
    const n = numOf(s); if (n != null) return n;
    const v = FX.evaluate(s, { locale: 'en', get: (r, sn) => sn ? (sh.wb && sh.wb.sheet(sn) ? sh.wb.sheet(sn).calc(r) : null) : sh.calc(r) });
    return isNum(v) ? v : null;
  }
  function dvAllows(sh, dv, v) {
    if (!isNum(v)) return false;
    if (dv.type === 'whole' && !Number.isInteger(v)) return false;
    const x = dvNum(sh, dv.formula1), y = dvNum(sh, dv.formula2);
    switch (dv.operator) {
      case 'between': return x != null && y != null && v >= Math.min(x, y) && v <= Math.max(x, y);
      case 'notBetween': return x != null && y != null && !(v >= Math.min(x, y) && v <= Math.max(x, y));
      case 'equal': return v === x;
      case 'notEqual': return v !== x;
      case 'greaterThan': return v > x;
      case 'lessThan': return v < x;
      case 'greaterThanOrEqual': return v >= x;
      case 'lessThanOrEqual': return v <= x;
    }
    return true;
  }
  const ruleTxt = (sh, dv) => {
    const x = dvNum(sh, dv.formula1), y = dvNum(sh, dv.formula2), s = v => v == null ? '?' : showV(v);
    return { between: `tussen ${s(x)} en ${s(y)}`, notBetween: `niet tussen ${s(x)} en ${s(y)}`, equal: `gelijk aan ${s(x)}`, notEqual: `niet gelijk aan ${s(x)}`, greaterThan: `groter dan ${s(x)}`, lessThan: `kleiner dan ${s(x)}`, greaterThanOrEqual: `groter dan of gelijk aan ${s(x)}`, lessThanOrEqual: `kleiner dan of gelijk aan ${s(x)}` }[dv.operator] || '';
  };
  const showAns = q => isNum(q.typed) ? showV(q.typed) : String(q.sol);
  const few = (arr, k) => arr.length > (k || 4) ? arr.slice(0, k || 4).join(', ') + ' …' : list(arr);
  const swappedQ = qs => qs.find(q => NEG.includes(q.okW) && POS.includes(q.badW));
  const SWAP_MSG = 'Wissel eerst de woorden voor juist en fout om in je ALS-formule (zie de opdracht over feedback).';
  const NEG = ['jammer', 'fout', 'onjuist', 'niet juist', 'probeer opnieuw', 'nee', 'helaas', 'mis', 'oei'], POS = ['juist', 'goed', 'super', 'prima', 'ja', 'top', 'bravo', 'correct', 'proficiat'];

  /* De opdrachten. lvl = vanaf welk niveau. */
  function m12tasks(mode) {
    const need = fn => ({ wb }) => {
      const an = analyze(wb);
      if (!an.qs.length) {
        if (an.unknown.length) return no(`In ${an.unknown[0].F} gebruik je de functie ${nlFn(an.unknown[0].fn)}. Die kent de app (nog) niet. Gebruik ALS, zoals in stap 9.`);
        return no('Ik vind nog geen vragen. Een vraag heeft een antwoordcel en een feedbackcel met ALS, bv. =ALS(C4="";"vul in";ALS(C4=B4;"juist";"jammer")).' + unresolvedMsg(an));
      }
      return fn(an, an.sheet, an.qs);
    };
    const T = [
      { id: 'b1', lvl: 1, text: 'Zet bovenaan een <b>titel</b> (vet of groter) en een korte <b>uitleg</b> voor de gebruiker: wat moet hij doen?',
        hints: ['De uitleg mag kort zijn, maar wel een zin (minstens 4 woorden).', 'Bv. in A1 "Rekenen en taal voor L3" (vet, groot) en in A2 "Typ je antwoord in de gele vakjes."'],
        check: ({ wb }) => {
          const an = analyze(wb), sh = an.sheet;
          const first = an.qs.length ? Math.min(...an.qs.map(q => rowOf(q.ans))) : Infinity;
          const qRows = new Set(an.qs.map(q => rowOf(q.ans)));
          const texts = sh.filled().filter(a => !sh.formula(a) && typeof sh.value(a) === 'string').sort((x, y) => rowOf(x) - rowOf(y) || colOf(x) - colOf(y));
          const perRow = {}; sh.filled().forEach(a => { perRow[rowOf(a)] = (perRow[rowOf(a)] || 0) + 1; });
          const above = texts.filter(a => rowOf(a) < first && perRow[rowOf(a)] < 3);   // een rij met 3 of meer kopjes is geen titel
          const title = above.find(a => { const st = sh.style(a); return st.bold || st.size >= 12; });
          if (!title) return no(above.length ? `Is "${cut(sh.text(above[0]))}" je titel? Maak ze dan vet of groter, zodat ze opvalt. Of zet er een titel boven.` : 'Zet bovenaan een titel, bv. "Werkblad voor het derde leerjaar", en maak ze vet of groter.');
          const uitleg = texts.find(a => a !== title && !qRows.has(rowOf(a)) && sh.text(a).trim().split(/\s+/).length >= 4);
          return uitleg ? ok() : no('Schrijf ook een korte uitleg voor de gebruiker, een zin van minstens 4 woorden. Bv. "Typ je antwoord in de gele vakjes."');
        } },
      { id: 'b2', lvl: 1, text: 'Maak <b>minstens 10 vragen</b>. Elke vraag heeft een <b>antwoordcel</b> en een <b>feedbackcel</b> met ALS.',
        hints: ['Eén rij per vraag: de vraag, de oplossing, de antwoordcel en de feedback.', 'Maak de ALS-formule één keer en trek ze door naar beneden.'],
        check: need((an, sh, qs) => {
          if (qs.length < 10) return no(`Ik vind ${qs.length} vra${qs.length === 1 ? 'ag' : 'gen'} met feedback (antwoordcel${qs.length === 1 ? '' : 'len'} ${few(qs.map(q => q.ans))}). Er moeten er minstens 10 zijn.` + unresolvedMsg(an));
          const noText = qs.filter(q => !hasQuestion(sh, q, an));
          if (noText.length) return no(`Bij de antwoordcel ${noText[0].ans} vind ik geen vraag. Zet de vraag links van de antwoordcel.`);
          return ok(`In orde: ik vind ${qs.length} vragen (antwoordcellen ${qs[0].ans} tot ${qs[qs.length - 1].ans}).`);
        }) },
      { id: 'b3', lvl: 1, text: 'Zorg voor <b>verschillende vakken</b>: minstens 2 vragen met een <b>getal</b> als antwoord (rekenen) en minstens 2 met een <b>woord</b> (taal, WO, Frans …).',
        hints: ['Een rekenvraag heeft een getal als oplossing, bv. 6 x 7 = 42.', 'Een taal- of WO-vraag heeft een woord als oplossing, bv. "Rome".'],
        check: need((an, sh, qs) => {
          const num = qs.filter(q => isNum(q.typed)).length, txt = qs.length - num;
          if (num < 2) return no(`Je werkblad heeft ${num} vra${num === 1 ? 'ag' : 'gen'} met een getal als antwoord. Zet er minstens 2 rekenvragen in.` + unresolvedMsg(an));
          if (txt < 2) return no(`Je werkblad heeft ${txt} vra${txt === 1 ? 'ag' : 'gen'} met een woord als antwoord. Zet er minstens 2 vragen over taal, WO of Frans in.` + unresolvedMsg(an));
          return ok(`In orde: ${num} vragen met een getal en ${txt} met een woord.`);
        }) },
      { id: 'b4', lvl: 1, text: 'Elke vraag geeft <b>feedback</b> met ALS: een woord als het juist is, een ander woord als het fout is, en iets anders als de cel nog <b>leeg</b> is (bv. "vul in").',
        hints: ['Eerst kijk je of de cel leeg is, dan of het antwoord klopt.', 'Bv. =ALS(C4="";"vul in";ALS(C4=B4;"juist";"jammer")) en doortrekken.'],
        check: need((an, sh, qs) => {
          for (const q of qs) {
            if (q.textNumber) return no(`In ${q.solRef ? q.solRef.a : q.F} staat ${q.sol} als tekst. Typt de gebruiker ${q.sol}, dan zegt ${q.F} "${showV(q.badRaw)}". Typ de oplossing opnieuw, gewoon als getal.`);
            if (NEG.includes(q.okW) && POS.includes(q.badW)) return no(`Als het antwoord juist is, toont ${q.F} "${showV(q.okRaw)}", en als het fout is "${showV(q.badRaw)}". Wissel de twee woorden om in je ALS-formule.`);
            if (numOf(q.okRaw) != null && numOf(q.badRaw) != null) return no(`${q.F} toont ${showV(q.okRaw)} of ${showV(q.badRaw)}. Laat woorden zien, bv. "juist" en "jammer".`);
            if (isErrV(q.emptyRaw)) return no(`Als ${q.ans} leeg is, toont ${q.F} een fout (${showV(q.emptyRaw)}). Begin met ALS(${q.ans}="";"vul in";…)`);
            if (q.emptyW === q.badW) return no(`Als ${q.ans} nog leeg is, toont ${q.F} al "${showV(q.badRaw)}". Laat dan iets anders zien, bv. "vul in": =ALS(${q.ans}="";"vul in";…)`);
          }
          if (an.unresolved.length) return no(unresolvedMsg(an).trim());
          return ok();
        }) },
      { id: 'b5', lvl: 1, text: 'Maak een <b>score</b>: tel hoeveel antwoorden juist zijn met <b>AANTAL.ALS</b>.',
        hints: ['AANTAL.ALS telt hoe vaak een woord voorkomt in je feedbackcellen.', 'Bv. =AANTAL.ALS(D4:D13;"juist"). Gebruik hetzelfde woord als in je ALS-formule.'],
        check: need((an, sh, qs) => {
          if (swappedQ(qs)) return no(SWAP_MSG);
          const si = scoreInfo(an), b = si.best;
          if (!b) {
            const all = Object.keys(sh.cells).filter(a => sh.formula(a));
            const ca = all.find(a => FX.info(sh.formula(a), 'en').funcs.includes('COUNTA'));
            if (ca) return no(`AANTALARG in ${ca} telt alle gevulde cellen, ook "${showV(qs[0].badRaw)}". Gebruik AANTAL.ALS.`);
            const ci = all.find(a => FX.info(sh.formula(a), 'en').funcs.includes('COUNTIF'));
            if (ci) return no(`AANTAL.ALS in ${ci} telt niet hoeveel antwoorden juist zijn. Kijk na wat er tussen de aanhalingstekens staat: bij jou is dat "${showV(qs[0].okRaw)}".`);
            return no(`Ik vind nog geen score. Tel het aantal juiste antwoorden, bv. =AANTAL.ALS(${qs[0].F}:${qs[qs.length - 1].F};"${showV(qs[0].okRaw)}").`);
          }
          if (b.cover.length < qs.length) { const miss = qs.filter(q => !b.cover.includes(q)).map(q => q.F); return no(`Je score in ${b.a} telt ${few(miss, 3)} nog niet mee. Neem al je feedbackcellen op in het bereik.`); }
          if (b.vN !== 0 || (isNum(b.vE) && b.vE !== 0)) return no(`Als alle antwoorden fout zijn, toont je score in ${b.a} ${showV(b.vN)}. Dat moet 0 zijn: tel alleen "${showV(qs[0].okRaw)}".`);
          return ok(`In orde: je score staat in ${b.a}.` + (b.countif ? '' : ' Je telde zonder AANTAL.ALS. Dat werkt ook!'));
        }) },
      { id: 'b6', lvl: 1, text: 'Laat de feedback <b>kleuren</b> met voorwaardelijke opmaak: <b>groen</b> als het juist is, <b>rood</b> als het fout is.',
        hints: ['Selecteer al je feedbackcellen. Start → Voorwaardelijke opmaak → Markeringsregels voor cellen → Tekst die bevat…', 'Eén regel voor "juist" (groene opvulling) en één regel voor "jammer" (lichtrode opvulling).'],
        check: need((an, sh, qs) => {
          if (swappedQ(qs)) return no(SWAP_MSG);
          const none = qs.filter(q => !sh.cfAt(q.F).length);
          if (none.length === qs.length) return no('Er staat nog geen voorwaardelijke opmaak op je feedbackcellen. Start → Voorwaardelijke opmaak → Markeringsregels voor cellen → Tekst die bevat…');
          if (none.length) return no(`De feedbackcel${none.length > 1 ? 'len' : ''} ${few(none.map(q => q.F), 3)} ${none.length > 1 ? 'kleuren' : 'kleurt'} nog niet. Selecteer al je feedbackcellen voor je de regel maakt.`);
          for (const q of qs) {
            const cO = colsOfStyle(sh.cfStyle(q.F, q.okRaw)), cB = colsOfStyle(sh.cfStyle(q.F, q.badRaw)), cE = colsOfStyle(sh.cfStyle(q.F, q.emptyRaw));
            if (!cO.includes('groen')) return no(`Bij "${showV(q.okRaw)}" wordt ${q.F} ${cO.length ? cO[0] : 'niet gekleurd'}. Het moet groen worden.`);
            if (cB.includes('groen') && q.badW.includes(q.okW)) return no(`"${showV(q.badRaw)}" bevat het woord "${showV(q.okRaw)}". Daardoor kleurt de regel "Tekst die bevat ${showV(q.okRaw)}" ook "${showV(q.badRaw)}" groen. Kies een ander woord, bv. "jammer", of gebruik de regel "Gelijk aan".`);
            if (!cB.includes('rood')) return no(`Bij "${showV(q.badRaw)}" wordt ${q.F} ${cB.length ? cB[0] : 'niet gekleurd'}. Het moet rood worden.`);
            if (cO.includes('rood') || cB.includes('groen')) return no(`${q.F} krijgt bij juist en bij fout dezelfde kleur. Kijk je regels na.`);
            if (!blank(q.emptyRaw) && q.emptyW !== q.badW && q.emptyW !== q.okW && (cE.includes('groen') || cE.includes('rood'))) return no(`Bij "${showV(q.emptyRaw)}" kleurt ${q.F} al ${cE.includes('groen') ? 'groen' : 'rood'}. Dat mag pas als er een antwoord staat. Kijk na welke tekst je regel zoekt.`);
          }
          return ok();
        }) },
      { id: 'b7', lvl: 1, text: 'Maak een vraag met een <b>afbeelding</b>, bv. "Welk dier zie je?"',
        hints: ['Tabblad Invoegen → Afbeeldingen.', 'Kies Stockafbeeldingen of een afbeelding van je computer, en zet ze bij de vraag. In Excel 365 kan ook "Plaatsen in cel".'],
        check: ({ wb }) => {
          const an = analyze(wb), im = an.sheet.imageInfo();
          if (im.floating || im.inCell.length) return ok();
          const elders = wb.sheets.find(s => s !== an.sheet && (s.imageInfo().floating || s.imageInfo().inCell.length));
          return no(elders ? `Je afbeelding staat op het blad "${elders.name}". Zet ze bij een vraag op je werkblad ("${an.sheet.name}").` : 'Ik vind nog geen afbeelding op je werkblad.');
        } },
      { id: 'b8', lvl: 1, text: '<b>Verberg</b> de oplossingen, zodat de gebruiker ze niet ziet.',
        hints: ['Staan je oplossingen in een kolom? Dan verberg je die kolom.', 'Rechtsklik op de kolomletter → Verbergen.'],
        check: need((an, sh, qs) => {
          const hidA = qs.find(q => hiddenAt(sh, q.ans));
          if (hidA) return no(`De antwoordcel ${hidA.ans} is verborgen. Dan kan de gebruiker niets invullen! Verberg alleen de oplossingen.`);
          const hidF = qs.find(q => hiddenAt(sh, q.F));
          if (hidF) return no(`De feedbackcel ${hidF.F} is verborgen. Dan ziet de gebruiker niet of zijn antwoord juist is.`);
          const withRef = qs.filter(q => q.solRef);
          const vis = withRef.filter(q => !hiddenAt(q.solRef.sh, q.solRef.a));
          if (vis.length) {
            const r = vis[0].solRef;
            return no(r.sh === sh ? `De oplossing in ${r.a} is nog zichtbaar. Verberg kolom ${numToCol(colOf(r.a))}: rechtsklik op de letter ${numToCol(colOf(r.a))} → Verbergen.` : `Je oplossingen staan op het blad "${r.sh.name}". Verberg dat blad: rechtsklik op het bladtabje → Verbergen.`);
          }
          return ok(withRef.length ? '' : 'In orde: je oplossingen zitten in je formules.');
        }) },
      ...(mode === 'test' ? [{ id: 'b9', lvl: 1, text: 'Maak de <b>antwoordcellen leeg</b> voor je inlevert, zodat een nieuwe gebruiker kan beginnen.',
        hints: [], check: need((an, sh, qs) => {
          const vol = qs.filter(q => !blank(sh.value(q.ans))).map(q => q.ans);
          return vol.length ? no(`Maak de antwoordcellen leeg (nog ingevuld: ${few(vol)}), bewaar en laad opnieuw op.`) : ok();
        }) }] : []),

      { id: 'g1', lvl: 2, text: 'Gebruik <b>SPATIES.WISSEN</b>: een spatie te veel achter een woord mag geen fout geven.',
        hints: ['Zet SPATIES.WISSEN rond het antwoord én rond de oplossing (stap 11).', '=ALS(C4="";"vul in";ALS(SPATIES.WISSEN(C4)=SPATIES.WISSEN(B4);"juist";"jammer"))'],
        check: need((an, sh, qs) => {
          const todo = qs.filter(q => !q.hasList && !isNum(q.typed));
          if (!todo.length) return ok('In orde: je woordvragen hebben allemaal een keuzelijst.');
          const bad = todo.map(q => ({ q, o: sh.calcWith(q.F, { [q.ans]: String(q.sol) + ' ' }) })).filter(x => said(x.o) !== x.q.okW);
          if (!bad.length) return ok();
          const { q, o } = bad[0], uses = FX.info(sh.formula(q.F), 'en').funcs.includes('TRIM');
          return no(`Typt de gebruiker "${q.sol} " (met een spatie erachter), dan toont ${q.F} "${showV(o)}". ${uses ? 'Kijk na waar je SPATIES.WISSEN zette.' : 'Gebruik SPATIES.WISSEN.'}${bad.length > 1 ? ` Dat geldt ook voor ${bad.length - 1} andere vra${bad.length === 2 ? 'ag' : 'gen'}.` : ''}`);
        }) },
      { id: 'g2', lvl: 2, text: 'Maak bij <b>minstens 2 vragen een keuzelijst</b>. Het juiste antwoord staat er natuurlijk tussen!',
        hints: ['Klik op de antwoordcel. Gegevens → Gegevensvalidatie → Toestaan: Lijst.', 'Typ bij Bron de keuzes met een puntkomma ertussen, bv. Parijs;Rome;Brussel.'],
        check: need((an, sh, qs) => {
          const lq = qs.filter(q => q.hasList);
          if (lq.length < 2) {
            const elders = sh.dataValidations.filter(d => d.type === 'list' && !qs.some(q => lists(sh, q.ans).includes(d))).map(d => d.sqref);
            if (elders.length) return no(`Je keuzelijst staat op ${elders[0]}, maar dat is geen antwoordcel. Zet de keuzelijst op de cel waar de gebruiker zijn antwoord kiest.`);
            return no(lq.length ? `Je hebt 1 keuzelijst (bij ${lq[0].ans}). Maak er minstens 2.` : 'Maak bij minstens 2 vragen een keuzelijst: Gegevens → Gegevensvalidatie → Lijst.');
          }
          for (const q of lq) {
            const items = sh.dvListItems(lists(sh, q.ans)[0]);
            if (!items) return no(`De keuzes van de keuzelijst in ${q.ans} kan ik niet lezen. Typ ze met een puntkomma ertussen, of kies een bereik op hetzelfde blad.`);
            if (items.length < 2) return no(`De keuzelijst in ${q.ans} heeft maar één keuze${items[0] ? ` ("${items[0]}")` : ''}. Zet een puntkomma tussen de keuzes.`);
            const want = norm(showAns(q));
            if (!items.some(x => norm(x) === want || (isNum(q.typed) && typedAs(x) === q.typed))) return no(`In de keuzelijst van ${q.ans} staat het juiste antwoord (${showAns(q)}) niet. Zo kan de gebruiker nooit juist antwoorden!`);
          }
          return ok(`In orde: ${lq.length} keuzelijsten.`);
        }) },
      { id: 'g3', lvl: 2, text: 'Toon de <b>score als procent</b>: het aantal juist gedeeld door het aantal vragen.',
        hints: ['Delen doe je met /. Daarna de knop %.', 'Bv. =B16/10 als je score in B16 staat en je 10 vragen hebt.'],
        check: need((an, sh, qs) => {
          const si = scoreInfo(an), n = qs.length, p = si.pct;
          if (!p) {
            const near = si.list.find(x => !Number.isInteger(x.vA) && x.vA > 0 && x.vA < 3);
            if (near && an.unresolved.length) return no('Eerst moeten al je vragen werken.' + unresolvedMsg(an));
            if (near) return no(`Als alles juist is, toont ${near.a} ${pctTxt(near.vA)} in plaats van 100%. Deel je door het juiste aantal vragen (${n})?`);
            return no(`Toon de score ook als procent: het aantal juist gedeeld door het aantal vragen (${n})${si.best ? `, bv. =${si.best.a}/${n}` : ''}.`);
          }
          if (!isNum(p.vH) || Math.abs(p.vH - si.half / n) > 1e-9) return no(`Met ${si.half} van de ${n} juist toont ${p.a} ${isNum(p.vH) ? pctTxt(p.vH) : showV(p.vH)}. Dat moet ${pctTxt(si.half / n)} zijn.`);
          return sh.numFmt(p.a).kind === 'percent' ? ok() : no(`${p.a} rekent goed, maar toont nog geen procent. Klik op %.`);
        }) },
      { id: 'g4', lvl: 2, text: 'Zet de <b>rasterlijnen</b> uit.',
        hints: ['Dat vind je op het tabblad Beeld.', 'Tabblad Beeld → haal het vinkje bij Rasterlijnen weg.'],
        check: ({ wb }) => analyze(wb).sheet.showGrid ? no('De rasterlijnen zijn nog zichtbaar.') : ok() },
      { id: 'g5', lvl: 2, text: '<b>Beveilig</b> het blad. Eerst ontgrendel je de antwoordcellen, de rest blijft geblokkeerd.',
        hints: ['Selecteer de antwoordcellen → Ctrl + 1 → Bescherming → vinkje Geblokkeerd weg.', 'Daarna: Controleren → Blad beveiligen → OK.'],
        check: need((an, sh, qs) => {
          if (!sh.isProtected) return no('Het blad is nog niet beveiligd. Controleren → Blad beveiligen.');
          const lockedA = qs.filter(q => sh.style(q.ans).locked).map(q => q.ans);
          if (lockedA.length === qs.length) return no('Het blad is beveiligd, maar de antwoordcellen zijn nog geblokkeerd. De gebruiker kan niets invullen! Hef de beveiliging op, ontgrendel de antwoordcellen en beveilig opnieuw.');
          if (lockedA.length) return no(`De antwoordcel${lockedA.length > 1 ? 'len' : ''} ${few(lockedA, 3)} ${lockedA.length > 1 ? 'zijn' : 'is'} nog geblokkeerd. Daar kan de gebruiker niets typen.`);
          const si = scoreInfo(an);
          const must = [...qs.map(q => q.F), ...qs.filter(q => q.solRef && q.solRef.sh === sh).map(q => q.solRef.a), ...(si.best ? [si.best.a] : [])];
          const open = [...new Set(must.filter(a => !sh.style(a).locked))];
          if (open.length) return no(`Ook ${few(open, 3)} ${open.length > 1 ? 'zijn' : 'is'} ontgrendeld. Dan kan de gebruiker daar iets veranderen of wissen. Ontgrendel alleen de antwoordcellen.`);
          const inF = qs.find(q => !q.solRef && !sh.style(q.F).hiddenFormula);
          if (inF) return no(`De oplossing van ${inF.ans} staat in de formule van ${inF.F}. Die kan de gebruiker lezen in de formulebalk. Hef de beveiliging op, selecteer je feedbackcellen → Ctrl + 1 → Bescherming → vink ook "Verborgen" aan, en beveilig opnieuw.`);
          return ok();
        }) },

      { id: 'e1', lvl: 3, text: 'Maak een <b>eindboodschap</b> die afhangt van de score, bv. "Super gedaan!" of "Oefen nog wat".',
        hints: ['Gebruik ALS met je score: is de score hoog genoeg?', 'Bv. =ALS(B16>=8;"Super gedaan!";"Oefen nog wat")'],
        check: need((an, sh, qs) => {
          const si = scoreInfo(an), m = si.msgs.find(x => x.differs);
          if (m) return ok(`In orde: bij alles juist zegt ${m.a} "${showV(m.oA)}", bij niets juist "${showV(m.oN)}".`);
          if (si.msgs.length) return no(`Je boodschap in ${si.msgs[0].a} zegt altijd "${showV(si.msgs[0].oA)}", ook als alles fout is. Kijk na met welke cel en welk getal je ALS-formule vergelijkt.`);
          return no(`Maak een eindboodschap met ALS die naar je score kijkt${si.best ? `, bv. =ALS(${si.best.a}>=${si.best.vA === 1 ? '80%' : Math.ceil(si.best.vA * 0.8)};"Super gedaan!";"Oefen nog wat")` : ''}.`);
        }) },
      { id: 'e2', lvl: 3, text: 'Laat de <b>score kleuren</b>: bv. groen als de score hoog is, rood als ze laag is.',
        hints: ['Voorwaardelijke opmaak op de cel met je score of je procent.', 'Markeringsregels voor cellen → Groter dan / Kleiner dan.'],
        check: need((an, sh, qs) => {
          const si = scoreInfo(an);
          const targets = [...new Set([si.best, si.pct].filter(Boolean))];
          const withCf = targets.filter(t => sh.cfAt(t.a).length);
          for (const t of withCf) {
            const cA = colsOfStyle(sh.cfStyle(t.a, t.vA)), cN = colsOfStyle(sh.cfStyle(t.a, t.vN));
            if ((cA.length || cN.length) && cA.join() !== cN.join()) return ok();
          }
          const mc = si.msgs.find(m => m.differs && sh.cfAt(m.a).length && colsOfStyle(sh.cfStyle(m.a, m.oA)).join() !== colsOfStyle(sh.cfStyle(m.a, m.oN)).join());
          if (mc) return ok();
          if (withCf.length) return no(`Je score in ${withCf[0].a} krijgt bij alles juist en bij niets juist dezelfde kleur. Kijk de getallen in je regel na.`);
          return no(`Zet voorwaardelijke opmaak op je score${targets.length ? ` (${list(targets.map(t => t.a))})` : ''}: bv. groen bij een hoge score.`);
        }) },
      { id: 'e3', lvl: 3, text: 'Laat bij een <b>rekenvraag alleen getallen</b> toe, met een <b>eigen foutmelding</b>.',
        hints: ['Gegevens → Gegevensvalidatie → Toestaan: Geheel getal (of Decimaal), met een minimum en een maximum.', 'Typ op het tabblad Foutmelding een eigen bericht, bv. "Typ een getal tussen 0 en 100."'],
        check: need((an, sh, qs) => {
          const nq = qs.filter(q => isNum(q.typed)).map(q => ({ q, dv: sh.dvAt(q.ans).find(d => d.type === 'whole' || d.type === 'decimal') })).filter(x => x.dv);
          if (!nq.length) return no('Geef een rekenvraag een gegevensvalidatie: Toestaan: Geheel getal, met een eigen foutmelding.');
          for (const { q, dv } of nq) if (!dvAllows(sh, dv, q.typed)) return no(`Het juiste antwoord van ${q.ans} (${showV(q.typed)}) past niet in je regel (${dv.type === 'whole' && !Number.isInteger(q.typed) ? 'geheel getal' : ruleTxt(sh, dv)}). Dan kan de gebruiker het juiste antwoord niet eens typen!`);
          const good = nq.find(x => x.dv.showErrorMessage && (x.dv.error || '').trim());
          return good ? ok() : no(`Typ bij de gegevensvalidatie van ${nq[0].q.ans} op het tabblad Foutmelding een eigen bericht.`);
        }) },
      { id: 'e4', lvl: 3, text: 'Laat minstens één keuzelijst de keuzes uit een <b>bereik</b> halen (bv. uit een verborgen kolom).',
        hints: ['Typ de keuzes in cellen naast elkaar, bv. K9:N9.', 'Gegevensvalidatie → Lijst → klik in het vak Bron en selecteer die cellen. Verberg daarna die kolommen.'],
        check: need((an, sh, qs) => {
          const lq = qs.filter(q => q.hasList).map(q => ({ q, dv: lists(sh, q.ans)[0] }));
          const r = lq.find(x => sh.dvIsRange(x.dv));
          if (!r) return no(lq.length ? 'Je keuzelijsten hebben losse woorden als bron. Laat minstens één keuzelijst de keuzes uit cellen halen: klik in het vak Bron en selecteer de cellen.' : 'Maak een keuzelijst met als bron een bereik: typ de keuzes in cellen en selecteer die cellen in het vak Bron.');
          const items = sh.dvListItems(r.dv);
          if (!items || items.length < 2) return no(`De bron van de keuzelijst in ${r.q.ans} (${r.dv.formula1}) bevat geen keuzes. Staan de keuzes in die cellen?`);
          return ok();
        }) }
    ];
    return T;
  }
  function starsOf(tasks, results) {
    const okIds = new Set((results || []).filter(r => r.ok).map(r => r.id));
    let s = 0;
    for (const L of [1, 2, 3]) {
      const t = tasks.filter(x => x.lvl <= L);
      if (!tasks.some(x => x.lvl === L)) break;
      if (t.every(x => okIds.has(x.id))) s = L; else break;
    }
    return s;
  }

  const m12 = {
    id: 'm12', nr: 12, project: true,
    title: 'Project: mijn interactief werkblad',
    short: 'Je maakt een echt werkblad voor je klas of voor een andere klas. Je kiest zelf je niveau: ★, ★★ of ★★★. Een klasgenoot test het, en daarna lever je het in.',
    skills: ['plannen', 'voor iemand anders', '10 vragen', 'ALS en AANTAL.ALS', 'kleuren', 'afbeelding', 'beveiligen', 'testen'],
    tabLabels: { practice: 'Tussentijds nakijken', test: 'Inleveren' },
    tabShort: { practice: 'Nakijken' },
    noun: 'werkblad', fileWord: 'werkblad',
    levels: NIVEAUS, levelOf, withLevel, starsOf,
    passCheck: (tasks, results) => starsOf(tasks, results) >= 1,   // in orde = alles van ★ klopt; elke ster meer is een extra
    route: {
      practice: ['Werk in Excel aan je werkblad. Begin met een <b>lege werkmap</b>, of open je bestand van stap 11 en kies <b>Bestand → Opslaan als</b> met een nieuwe naam.',
        'Bewaar het in <b>OneDrive</b> met de naam <span class="addr">{fname}</span>.',
        'Laad het tussendoor op: de app kijkt na wat al in orde is. Dat mag zo vaak als je wil!'],
      test: ['Maak de <b>antwoordcellen leeg</b>, zodat een nieuwe gebruiker kan beginnen.', 'Bewaar je werkblad (<b>Ctrl + S</b>).', 'Laad het hieronder op. De app kijkt alles na en stuurt het naar je leerkracht.']
    },
    testIntro: 'Geen tips meer. De app speelt gebruiker en kijkt alles na. Je werkblad is goedgekeurd als alle opdrachten van ★ in orde zijn. Elke ster meer is een extra.',
    gate: {
      intro: '<p>Je werkblad is bijna klaar. Eerst nog dit:</p>',
      checks: ['Een klasgenoot heeft mijn werkblad getest met de testkaart.', 'Ik heb de tip van mijn tester bekeken en mijn werkblad verbeterd.', 'De antwoordcellen zijn weer leeg en mijn werkblad is bewaard.'],
      fields: [{ id: 'tester', label: 'Wie testte je werkblad?' }, { id: 'tip', label: 'Welke tip kreeg je?', area: true }]
    },
    lessons: [
      { id: 'l1', title: 'Voor wie maak je het?',
        body: `<p>Dit is het grote project: je maakt een <b>echt werkblad</b> dat iemand anders gaat gebruiken. Voor je eigen klas, of voor een andere klas, bv. het derde leerjaar.</p>
<p>Denk eerst aan je <b>gebruiker</b>. Wat kan hij al? Welke woorden kent hij? Hoe weet hij wat hij moet doen?</p>`,
        widget: { type: 'quiz', questions: [
          { q: 'Je maakt een werkblad voor het derde leerjaar. Wat is het beste?', options: ['Korte vragen, eenvoudige woorden en een duidelijke uitleg', 'Zo moeilijk mogelijk, dan leren ze meer', 'Geen uitleg: ze zien wel hoe het werkt'], correct: 0, explain: 'Een goed werkblad past bij de gebruiker. Wat voor jou makkelijk is, kan voor een jongere leerling moeilijk zijn.' },
          { q: 'Waarom zet je een afbeelding bij een vraag?', options: ['Ze helpt de gebruiker de vraag te begrijpen, en het is leuker', 'Om de vragen te bedekken', 'Excel werkt niet zonder afbeelding'], correct: 0, explain: 'Een afbeelding kan zelf de vraag zijn: "Welk dier zie je?"' },
          { q: 'Hoe weet je gebruiker waar hij zijn antwoord typt?', options: ['De antwoordcellen hebben een kleur en de uitleg zegt het', 'Hij klikt overal tot het lukt', 'Dat raadt hij wel'], correct: 0, explain: 'Een opvulkleur (stap 3) en één zin uitleg helpen veel. Bv. "Typ je antwoord in de gele vakjes."' }
        ] } },
      { id: 'l2', title: 'Kies je niveau',
        body: `<p>Iedereen maakt <b>minstens 10 vragen</b>, met <b>rekenen én taal of WO</b>. Daarna kies je hoe ver je gaat:</p>
<p><b>★ Basis</b> – een werkblad dat werkt: feedback, een score, kleuren, een afbeelding en verborgen oplossingen.<br>
<b>★★ Gevorderd</b> – alles van ★, plus SPATIES.WISSEN, keuzelijsten, de score in procent, geen rasterlijnen en een beveiligd blad.<br>
<b>★★★ Expert</b> – alles van ★★, plus een eindboodschap, een score die kleurt, een rekenvraag met een eigen foutmelding en een keuzelijst uit een bereik.</p>
<p>Kies wat bij je past. Je mag later altijd <b>een niveau hoger</b> proberen (of lager).</p>`,
        widget: { type: 'levelPick' } },
      { id: 'l3', title: 'Een plan op papier',
        body: `<p><b>Eerst goed denken, dan pas bouwen.</b> Maak je plan op papier, voor je Excel opent.</p>
<p>Kijk je oplossingen goed na! Een fout in je oplossing betekent dat je gebruiker "jammer" krijgt terwijl hij juist antwoordt.</p>
<div class="grid-wrap"><table class="mini sample"><tbody>
<tr><td><b>Vraag</b></td><td><b>Oplossing</b></td><td><b>Antwoord</b></td><td><b>Feedback</b></td></tr>
<tr><td>6 x 7 =</td><td class="n">42</td><td></td><td>vul in</td></tr>
<tr><td>Hoofdstad van Italië?</td><td>Rome</td><td></td><td>vul in</td></tr>
<tr><td>Welk dier zie je? 🐘</td><td>olifant</td><td></td><td>vul in</td></tr>
</tbody></table></div>`,
        widget: { type: 'steps', steps: [
          'Ik weet <b>voor wie</b> ik het werkblad maak.',
          'Ik schrijf minstens <b>10 vragen</b> met hun <b>oplossing</b> op papier.',
          'Er zijn <b>rekenvragen</b> (het antwoord is een getal) én vragen over <b>taal, WO of Frans</b> (het antwoord is een woord).',
          'Ik weet welke vraag een <b>afbeelding</b> krijgt.',
          'Ik teken mijn <b>kolommen</b>: vraag, oplossing, antwoord en feedback.',
          'Bij ★★ en ★★★: ik weet welke vragen een <b>keuzelijst</b> krijgen.'
        ], confirm: 'Mijn plan is klaar' } },
      { id: 'l4', title: 'Zo kijkt de app na',
        body: `<p>Jij kiest hoe je werkblad eruitziet. De app zoekt zelf je vragen: elke <b>ALS-formule</b> die naar een antwoordcel kijkt, is een vraag.</p>
<p>Daarna <b>speelt de app gebruiker</b>. In gedachten vult ze je werkblad in: alles juist, alles fout, alles leeg, een spatie te veel … en ze kijkt wat je werkblad dan zegt. Klopt je score? Verandert de kleur?</p>
<p>Tip: laad je werkblad <b>tussendoor</b> op bij Tussentijds nakijken. Zo zie je meteen wat al in orde is.</p>`,
        widget: { type: 'quiz', questions: [
          { q: 'Je antwoordcellen staan in kolom E in plaats van C. Vindt de app je vragen?', options: ['Ja: de app zoekt de cellen waar je ALS-formule naar kijkt', 'Nee, het moet altijd kolom C zijn'], correct: 0, explain: 'Je bent vrij in je opbouw. Zet wel de vraag links van de antwoordcel, dan leest het makkelijk.' },
          { q: 'Je feedback zegt "goed" en "probeer opnieuw". Mag dat?', options: ['Ja, als je score dan ook "goed" telt en je kleuren "goed" zoeken', 'Nee, alleen juist en jammer'], correct: 0, explain: 'Je mag je eigen woorden kiezen. Gebruik ze dan overal hetzelfde: in ALS, in AANTAL.ALS en in je kleurregels.' },
          { q: 'De app typt in gedachten het juiste antwoord, maar je werkblad zegt "jammer". Wat is er mis?', options: ['Waarschijnlijk klopt mijn oplossing of mijn formule niet', 'De app is stuk'], correct: 0, explain: 'Vaak is het een tikfout in de oplossing, een spatie te veel, of een getal dat als tekst getypt is. Probeer het zelf in Excel!' }
        ] } },
      { id: 'l5', title: 'Testen door een klasgenoot',
        body: `<p>Een echte maker laat zijn werk <b>testen</b>. Vraag een klasgenoot om aan jouw laptop je werkblad te proberen, met deze <b>testkaart</b>. Jij kijkt mee en zegt niets!</p>
<p>Daarna verbeter je wat niet goed werkte, en maak je de antwoordcellen weer leeg. Bij <b>Inleveren</b> vul je in wie testte en welke tip je kreeg.</p>`,
        widget: { type: 'steps', steps: [
          'Lees de uitleg. Weet je meteen wat je moet doen?',
          'Vul alles <b>juist</b> in. Staat de score op het maximum?',
          'Typ bij een vraag iets <b>fouts</b>. Wordt het rood?',
          'Laat een vraag <b>leeg</b>. Zegt het werkblad "vul in"?',
          'Typ een woord met een <b>spatie</b> erachter. Blijft het juist? (★★)',
          'Probeer een <b>oplossing</b> te zien of een <b>formule</b> te wissen. Lukt het niet? Prima! (★★)',
          'Zeg de maker <b>één top</b> (wat goed is) en <b>één tip</b> (wat beter kan).'
        ], confirm: 'Ik ken de testkaart' } }
    ],
    newFile: true,
    gen(seed, mode) {
      const L = levelOf(seed);
      return { startFile: null, level: L, tasks: m12tasks(mode).filter(t => t.lvl <= L) };
    }
  };

  const MODULES = [m1, m2, m3, m4, m5, m6, m7, m8, m9, m10, m11, m12];

  /* Volgende modules (nog te bouwen) – getoond als vooruitblik. */
  const SOON = [];

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
  const api = { MODULES, SOON, rng, newSeed, runChecks, PASS, KLEUREN, fillSeries, numOf, MAANDEN, DAGEN, nlNum };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.CvcModules = api;
})(typeof window !== 'undefined' ? window : globalThis);
