/* Cel voor Cel – leerling-app */
(function () {
  'use strict';
  const { MODULES, SOON, newSeed, runChecks, PASS } = window.CvcModules;
  const X = window.XlsxCheck, Api = window.CvcApi;
  const $app = document.getElementById('app');

  // ---------------- kleine DOM-hulp ----------------
  function h(tag, attrs, ...kids) {
    const el = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (v == null || v === false) continue;
      if (k === 'class') el.className = v;
      else if (k === 'html') el.innerHTML = v;
      else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
      else if (v === true) el.setAttribute(k, '');
      else el.setAttribute(k, v);
    }
    for (const c of kids.flat(Infinity)) if (c != null && c !== false) el.append(c.nodeType ? c : document.createTextNode(String(c)));
    return el;
  }
  const clear = el => { while (el.firstChild) el.removeChild(el.firstChild); return el; };
  const fmtDate = d => new Date(d).toLocaleString('nl-BE', { day: 'numeric', month: 'long', hour: '2-digit', minute: '2-digit' });

  // ---------------- toestand ----------------
  const S = { student: null, view: 'home', mid: null, tab: 'learn', lesson: 0, busy: false, uploadRoute: 'A', results: {} };
  const prog = mid => (S.student && S.student.progress && S.student.progress[mid]) || {};
  const modIndex = mid => MODULES.findIndex(m => m.id === mid);
  const isUnlocked = i => i === 0 || !!prog(MODULES[i - 1].id).test?.passed;
  const lessonsDone = m => m.lessons.filter(l => prog(m.id).lessons?.[l.id]).length;

  async function save(mid, patch) {
    try { const d = await Api.progress(mid, patch); S.student = d.student; }
    catch (e) { if (e.status === 401) return logout(e.message); toast(e.message, true); }
  }
  function toast(msg, bad) {
    const t = h('div', { class: 'msg ' + (bad ? 'bad' : 'ok'), role: 'status', style: 'position:fixed;left:50%;transform:translateX(-50%);bottom:calc(16px + env(safe-area-inset-bottom,0px));z-index:20;max-width:min(560px,calc(100vw - 32px));box-shadow:var(--shadow)' }, msg);
    document.body.append(t); setTimeout(() => t.remove(), 4200);
  }

  // ---------------- voorlezen ----------------
  function speak(text) {
    if (!('speechSynthesis' in window)) return toast('Voorlezen werkt niet in deze browser.', true);
    speechSynthesis.cancel();
    const u = new SpeechSynthesisUtterance(text);
    const voices = speechSynthesis.getVoices();
    u.voice = voices.find(v => v.lang === 'nl-BE') || voices.find(v => /^nl/i.test(v.lang)) || null;
    u.lang = u.voice ? u.voice.lang : 'nl-BE'; u.rate = 0.95;
    speechSynthesis.speak(u);
  }

  // ---------------- bovenbalk ----------------
  function renderWho() {
    const w = clear(document.getElementById('who'));
    if (!S.student) return;
    w.append(h('div', { class: 'btn-row', style: 'flex-wrap:nowrap' },
      h('div', { class: 'namebox', title: 'Aangemeld als' }, h('b', {}, S.student.klas), h('span', {}, S.student.voornaam + ' ' + S.student.naam)),
      h('button', { class: 'btn ghost small', onclick: () => logout() }, 'Afmelden')));
  }
  function logout(msg) {
    Api.logout(); S.student = null; S.view = 'home'; renderWho(); renderLogin(msg);
  }
  document.getElementById('logo').addEventListener('click', e => { e.preventDefault(); if (S.student) { S.view = 'home'; render(); } });

  // ---------------- aanmelden ----------------
  function renderLogin(msg) {
    clear($app);
    const err = h('div', { class: 'msg bad', hidden: !msg }, msg || '');
    const f = {
      voornaam: h('input', { class: 'in', id: 'f-voornaam', autocomplete: 'given-name', required: true }),
      naam: h('input', { class: 'in', id: 'f-naam', autocomplete: 'family-name', required: true }),
      klas: h('input', { class: 'in', id: 'f-klas', placeholder: 'bv. L5', required: true }),
      code: h('input', { class: 'in', id: 'f-code', inputmode: 'numeric', type: 'password', pattern: '[0-9]*', maxlength: 8, required: true })
    };
    const newBox = h('div', { class: 'confirm-box', hidden: true });
    const go = h('button', { class: 'btn primary', type: 'submit' }, 'Aanmelden');
    const form = h('form', { class: 'form sheet pad', novalidate: true },
      h('h2', {}, 'Aanmelden'),
      h('div', { class: 'two' },
        h('label', { class: 'f', for: 'f-voornaam' }, 'Voornaam', f.voornaam),
        h('label', { class: 'f', for: 'f-naam' }, 'Naam', f.naam)),
      h('div', { class: 'two' },
        h('label', { class: 'f', for: 'f-klas' }, 'Klas', f.klas),
        h('label', { class: 'f', for: 'f-code' }, h('span', {}, 'Jouw code ', h('small', {}, '(4 tot 8 cijfers)')), f.code)),
      err, newBox, h('div', { class: 'btn-row' }, go));
    const send = async nieuw => {
      err.hidden = true; go.disabled = true;
      try {
        const d = await Api.login({ voornaam: f.voornaam.value, naam: f.naam.value, klas: f.klas.value, code: f.code.value, nieuw });
        if (d.nieuw) {
          clear(newBox).append(
            h('p', {}, h('b', {}, 'Ben je hier voor de eerste keer?'), ' Er is nog geen account voor ', h('b', {}, `${f.voornaam.value.trim()} ${f.naam.value.trim()}`), ' in ', h('b', {}, f.klas.value.trim().toUpperCase()), '.'),
            h('p', { class: 'small' }, 'Onthoud je code goed: je hebt hem elke keer nodig. Klopt je naam niet? Verbeter hem dan eerst.'),
            h('div', { class: 'btn-row' }, h('button', { class: 'btn primary', type: 'button', onclick: () => send(true) }, 'Ja, maak mijn account'), h('button', { class: 'btn', type: 'button', onclick: () => { newBox.hidden = true; } }, 'Nee, verbeteren')));
          newBox.hidden = false;
        } else { S.student = d.student; S.view = 'home'; renderWho(); render(); }
      } catch (e) { err.textContent = e.message; err.hidden = false; }
      go.disabled = false;
    };
    form.addEventListener('submit', e => { e.preventDefault(); send(false); });

    const hg = h('div', { class: 'hero-grid', 'aria-hidden': 'true' });
    const cols = ['', 'A', 'B', 'C', 'D', 'E'];
    cols.forEach(c => hg.append(h('div', { class: 'h' }, c)));
    for (let r = 1; r <= 3; r++) cols.forEach((c, i) => hg.append(h('div', { class: i === 0 ? 'h' : (r === 2 && c === 'C' ? 'sel' : '') }, i === 0 ? r : (r === 2 && c === 'C' ? 'C2' : ''))));

    $app.append(h('div', { class: 'login-wrap' },
      h('section', { class: 'hero' },
        h('h1', { html: 'Excel leren, <em>cel voor cel</em>' }),
        h('p', {}, 'Je werkt op je eigen tempo. Eerst leer je hoe het werkt, dan oefen je in het echte Excel, en daarna toon je in een test wat je kan. Stap voor stap bouw je naar je eigen maandrapport en een interactief werkblad.'),
        hg),
      form));
    f.voornaam.focus();
  }

  // ---------------- leerlijn ----------------
  const levelNow = m => { const p = prog(m.id); return m.levels && p.practice?.seed ? m.levelOf(p.practice.seed) : null; };
  const levelInfo = (m, n) => m.levels.find(L => L.n === n) || m.levels[0];
  function moduleChips(m) {
    const p = prog(m.id), chips = [];
    const ld = lessonsDone(m);
    chips.push(h('span', { class: 'chip' + (ld === m.lessons.length ? ' ok' : '') }, `Leren ${ld}/${m.lessons.length}`));
    if (m.levels) {
      const lv = levelNow(m);
      if (lv) chips.push(h('span', { class: 'chip' }, 'Niveau ' + levelInfo(m, lv).stars));
      if (p.practice?.attempts) chips.push(h('span', { class: 'chip go' }, `Nagekeken: ${p.practice.attempts}×`));
      if (p.test?.passed) chips.push(h('span', { class: 'chip ok' }, `Ingeleverd ✓ ${'★'.repeat(p.test.stars || 0)}`));
      else if (p.test?.attempts) chips.push(h('span', { class: 'chip bad' }, 'Inleveren: nog niet in orde'));
      return chips;
    }
    if (p.practice?.passed) chips.push(h('span', { class: 'chip ok' }, 'Oefenen ✓'));
    else if (p.practice?.attempts) chips.push(h('span', { class: 'chip go' }, `Oefenen: ${p.practice.attempts}× opgeladen`));
    if (p.test?.passed) chips.push(h('span', { class: 'chip ok' }, `Test ${p.test.best}% ✓`));
    else if (p.test?.attempts) chips.push(h('span', { class: 'chip bad' }, `Test ${p.test.best || 0}%`));
    return chips;
  }
  function renderHome() {
    clear($app);
    const cur = MODULES.findIndex((m, i) => isUnlocked(i) && !prog(m.id).test?.passed);
    const allDone = cur === -1;
    const rows = MODULES.map((m, i) => {
      const open = isUnlocked(i), done = !!prog(m.id).test?.passed;
      return h('div', { class: 'mrow' + (i === cur ? ' current' : '') + (open ? '' : ' locked') + (m.project ? ' project' : '') },
        h('div', { class: 'rn' }, m.nr),
        h('div', { class: 'mb' }, h('h3', {}, m.title), h('p', { class: 'small' }, open ? m.short : 'Dit deel gaat open als je de test van de vorige stap haalt.'), open ? h('div', { class: 'chips' }, moduleChips(m)) : null),
        h('div', { class: 'ma' }, open ? h('button', { class: 'btn' + (i === cur ? ' primary' : ''), onclick: () => openModule(m.id) }, done ? 'Bekijk' : (i === cur && (lessonsDone(m) || prog(m.id).practice) ? 'Verder' : 'Start')) : h('span', { class: 'muted small' }, 'Op slot')));
    });
    const soon = SOON.map(s => h('div', { class: 'mrow soon' + (s.project ? ' project' : '') }, h('div', { class: 'rn' }, s.nr), h('div', { class: 'mb' }, h('h3', {}, s.title), h('p', { class: 'small' }, s.short))));
    $app.append(
      h('div', { class: 'home-head' },
        h('div', {}, h('h1', {}, `Dag ${S.student.voornaam}!`), h('p', { class: 'muted' }, allDone ? (SOON.length ? 'Je hebt alle stappen die nu klaarstaan gehaald. Knap!' : `Je hebt alle ${MODULES.length} stappen gehaald. Proficiat, je bent een echte Excel-kenner!`) : 'Dit is jouw leerlijn. Werk de stappen van boven naar onder af.'))),
      h('div', { class: 'path' }, rows),
      ...(SOON.length ? [h('p', { class: 'sep-label' }, 'Komt later'), h('div', { class: 'path' }, soon)] : []));
  }

  // ---------------- module ----------------
  function openModule(mid, tab) {
    S.view = 'module'; S.mid = mid;
    const m = MODULES[modIndex(mid)], p = prog(mid);
    S.tab = tab || (p.test?.open ? 'test' : (lessonsDone(m) === m.lessons.length ? 'practice' : 'learn'));
    const firstOpen = m.lessons.findIndex(l => !p.lessons?.[l.id]);
    S.lesson = firstOpen < 0 ? 0 : firstOpen;
    render(); window.scrollTo(0, 0);
  }
  function renderModule() {
    const m = MODULES[modIndex(S.mid)], p = prog(m.id);
    clear($app);
    const TL = m.tabLabels || {};
    const tabs = [['learn', 'Leren', lessonsDone(m) === m.lessons.length], ['practice', TL.practice || 'Oefenen in Excel', !!p.practice?.passed], ['test', TL.test || 'Test', !!p.test?.passed]];
    const tabBar = h('div', { class: 'tabs', role: 'tablist' },
      tabs.map(([id, label, done], i) => h('button', { class: 'tab', role: 'tab', 'aria-selected': String(S.tab === id), onclick: () => { S.tab = id; render(); } }, h('span', { class: 'n' }, i + 1), m.tabShort && m.tabShort[id] ? [h('span', { class: 'tl-long' }, label), h('span', { class: 'tl-short' }, m.tabShort[id])] : label, done ? h('span', { class: 'done', 'aria-label': 'klaar' }, '✓') : null)));
    const skip = S.tab !== 'test' && !p.test?.passed ? h('button', { class: 'btn ghost small skip', onclick: () => { S.tab = 'test'; render(); } }, m.project ? `Meteen naar ${(TL.test || 'de test').toLowerCase()} →` : 'Ik kan dit al → naar de test') : null;
    $app.append(
      h('div', { class: 'crumb' }, h('button', { class: 'btn ghost small', onclick: () => { S.view = 'home'; render(); } }, '← Leerlijn')),
      h('div', { class: 'mod-head' }, h('span', { class: 'eyebrow' }, 'Stap ' + m.nr), h('h1', {}, m.title), h('div', { class: 'chips' }, m.skills.map(s => h('span', { class: 'chip' }, s)))),
      h('div', { class: 'tabs-row' }, tabBar, skip));
    const body = h('div', {});
    $app.append(body);
    if (S.tab === 'learn') renderLearn(m, body);
    else if (S.tab === 'practice') renderWork(m, body, 'practice');
    else renderWork(m, body, 'test');
  }

  // ---------------- leren ----------------
  function renderLearn(m, root) {
    const p = prog(m.id);
    const nav = h('nav', { class: 'lnav', 'aria-label': 'Lessen' }, m.lessons.map((l, i) =>
      h('button', { 'aria-current': String(i === S.lesson), onclick: () => { S.lesson = i; render(); } },
        h('span', { class: 'ck' + (p.lessons?.[l.id] ? ' on' : '') }, p.lessons?.[l.id] ? '✓' : ''), h('span', {}, l.title))));
    const l = m.lessons[S.lesson];
    const done = !!p.lessons?.[l.id];
    const band = h('div', { class: 'done-band', hidden: !done });
    const fillBand = () => {
      const last = S.lesson === m.lessons.length - 1;
      const allDone = m.lessons.every(x => prog(m.id).lessons?.[x.id]);
      clear(band).append(h('b', {}, allDone ? 'Alle lessen zijn klaar. Tijd om te oefenen in Excel!' : 'Deze les is klaar.'),
        allDone && last ? h('button', { class: 'btn primary', onclick: () => { S.tab = 'practice'; render(); window.scrollTo(0, 0); } }, 'Naar Oefenen in Excel →')
          : h('button', { class: 'btn primary', onclick: () => { S.lesson = last ? m.lessons.findIndex(x => !prog(m.id).lessons?.[x.id]) : S.lesson + 1; render(); } }, 'Volgende les →'));
    };
    if (done) fillBand();
    const onDone = async () => {
      if (!prog(m.id).lessons?.[l.id]) await save(m.id, { lessons: { [l.id]: true } });
      nav.querySelectorAll('.ck')[S.lesson].className = 'ck on'; nav.querySelectorAll('.ck')[S.lesson].textContent = '✓';
      fillBand(); band.hidden = false; band.classList.add('pop');
    };
    const bodyEl = h('div', { class: 'body', html: l.body });
    const card = h('article', { class: 'sheet pad lesson' },
      h('div', { class: 'lesson-top' }, h('h2', {}, l.title), h('button', { class: 'btn small tts', onclick: () => speak(l.title + '. ' + bodyEl.textContent) }, '🔊 Lees voor')),
      bodyEl, h('div', { class: 'widget' }, WIDGETS[l.widget.type](l.widget, onDone)), band);
    root.append(h('div', { class: 'learn' }, nav, card));
  }

  // ---------------- widgets ----------------
  const COLS = 'ABCDEFGHIJ';
  const rnd = (a, b) => a + Math.floor(Math.random() * (b - a + 1));

  function miniGrid(opt) {
    const cols = opt.cols || 10, rows = opt.rows || 12;
    const t = h('table', { class: 'mini' });
    const head = h('tr', {}, h('th', { class: 'corner' }));
    for (let c = 0; c < cols; c++) head.append(h('th', { class: opt.headClick ? 'click' : '', 'data-col': COLS[c] }, COLS[c]));
    t.append(h('thead', {}, head));
    const tb = h('tbody');
    for (let r = 1; r <= rows; r++) {
      const tr = h('tr', {}, h('th', { class: opt.headClick ? 'click' : '', 'data-row': r }, r));
      for (let c = 0; c < cols; c++) tr.append(h('td', { class: opt.cellClick ? 'click' : '', 'data-a': COLS[c] + r }));
      tb.append(tr);
    }
    t.append(tb);
    const nb = h('b', {}, opt.namebox || '');
    const wrap = h('div', {}, h('div', { class: 'mini-namebox' }, 'Naamvak', nb), h('div', { class: 'grid-wrap' }, t));
    return { wrap, table: t, nb, cell: a => t.querySelector(`td[data-a="${a}"]`), clearMarks: () => t.querySelectorAll('.hl,.hl-strong,.target,.wrong,.right,.insel').forEach(e => e.classList.remove('hl', 'hl-strong', 'target', 'wrong', 'right', 'insel')) };
  }
  function roundsUI(total) {
    const ask = h('p', { class: 'ask' }), fb = h('p', { class: 'fb' }), cnt = h('span', { class: 'qcount' });
    return { ask, fb, cnt, set(i, html) { ask.innerHTML = html; cnt.textContent = `Vraag ${i + 1} van ${total}`; } };
  }
  function good(fb, txt) { fb.className = 'fb ok'; fb.textContent = txt || 'Juist!'; }
  function bad(fb, txt, el) { fb.className = 'fb bad'; fb.textContent = txt; if (el) { el.classList.remove('shake'); void el.offsetWidth; el.classList.add('shake'); } }

  const WIDGETS = {
    quiz(w, onDone) {
      let i = 0;
      const box = h('div', { class: 'q' });
      const show = () => {
        const q = w.questions[i];
        const fb = h('p', { class: 'fb' }), next = h('div', { class: 'btn-row' });
        const opts = h('div', { class: 'opts' }, q.options.map((o, k) => h('button', { html: o, onclick: e => {
          if (k === q.correct) {
            e.currentTarget.classList.add('right'); opts.querySelectorAll('button').forEach(b => { b.disabled = true; });
            good(fb, 'Juist! ' + q.explain);
            if (i < w.questions.length - 1) next.append(h('button', { class: 'btn', onclick: () => { i++; show(); } }, 'Volgende vraag →'));
            else onDone();
          } else { e.currentTarget.classList.add('wrong'); bad(fb, 'Nog niet. Probeer opnieuw.'); }
        } })));
        clear(box).append(h('span', { class: 'qcount' }, `Vraag ${i + 1} van ${w.questions.length}`), h('p', { class: 'ask' }, q.q), opts, fb, next);
      };
      show();
      return box;
    },
    gridHeader(w, onDone) {
      const g = miniGrid({ headClick: true }); const ui = roundsUI(w.rounds);
      let i = 0, want;
      const next = () => {
        if (i >= w.rounds) { good(ui.fb, 'Super! Je kent kolommen en rijen.'); return onDone(); }
        want = i % 2 === 0 ? { col: COLS[rnd(0, 9)] } : { row: rnd(1, 12) };
        ui.set(i, want.col ? `Klik op kolom <span class="addr">${want.col}</span>.` : `Klik op rij <span class="addr">${want.row}</span>.`);
      };
      g.table.addEventListener('click', e => {
        const th = e.target.closest('th'); if (!th || i >= w.rounds) return;
        g.clearMarks();
        const col = th.dataset.col, row = th.dataset.row && +th.dataset.row;
        if (!col && !row) return;
        g.table.querySelectorAll(col ? `td[data-a^="${col}"]` : `tr:nth-child(${row}) td`).forEach(td => { if (!col || td.dataset.a.replace(/\d+/, '') === col) td.classList.add('hl'); });
        if ((want.col && col === want.col) || (want.row && row === want.row)) { good(ui.fb, `Juist, dat is ${col ? 'kolom ' + col : 'rij ' + row}.`); i++; setTimeout(next, 700); }
        else bad(ui.fb, col ? `Dat is kolom ${col}${want.row ? '. Rijen hebben een getal.' : '.'}` : `Dat is rij ${row}${want.col ? '. Kolommen hebben een letter.' : '.'}`, th);
      });
      next();
      return h('div', { class: 'widget-in' }, ui.cnt, ui.ask, g.wrap, ui.fb);
    },
    gridClick(w, onDone) {
      const g = miniGrid({ cellClick: true }); const ui = roundsUI(w.rounds);
      let i = 0, want;
      const next = () => { if (i >= w.rounds) { good(ui.fb, 'Top! Je vindt elke cel.'); return onDone(); } want = COLS[rnd(0, 9)] + rnd(1, 12); ui.set(i, `Klik op cel <span class="addr">${want}</span>.`); };
      g.table.addEventListener('mouseover', e => { const td = e.target.closest('td'); if (td) g.nb.textContent = td.dataset.a; });
      g.table.addEventListener('click', e => {
        const td = e.target.closest('td'); if (!td || i >= w.rounds) return;
        g.clearMarks(); g.nb.textContent = td.dataset.a;
        if (td.dataset.a === want) { td.classList.add('right'); good(ui.fb, `Juist: ${want}.`); i++; setTimeout(() => { g.clearMarks(); next(); }, 800); }
        else { td.classList.add('wrong'); bad(ui.fb, `Dit is ${td.dataset.a}. Zoek eerst kolom ${want.replace(/\d+/, '')}, dan rij ${want.replace(/[A-Z]+/, '')}.`, td); }
      });
      next();
      return h('div', {}, ui.cnt, ui.ask, g.wrap, ui.fb);
    },
    gridRead(w, onDone) {
      const g = miniGrid({}); const ui = roundsUI(w.rounds);
      const inp = h('input', { class: 'in', id: 'read-' + Math.random().toString(36).slice(2, 7), 'aria-label': 'Celadres', autocomplete: 'off', maxlength: 4 });
      const btn = h('button', { class: 'btn primary' }, 'Controleer');
      let i = 0, want;
      const next = () => {
        g.clearMarks(); inp.value = '';
        if (i >= w.rounds) { good(ui.fb, 'Knap! Je kan celadressen lezen.'); inp.disabled = btn.disabled = true; return onDone(); }
        want = COLS[rnd(0, 9)] + rnd(1, 12); g.cell(want).classList.add('target'); ui.set(i, 'Welk adres heeft de groen omrande cel?');
      };
      const check = () => {
        const v = inp.value.trim().toUpperCase().replace(/\s/g, '');
        if (!v) return;
        if (v === want) { good(ui.fb, `Juist: ${want}.`); i++; setTimeout(next, 800); }
        else if (/^\d+[A-Z]+$/.test(v)) bad(ui.fb, 'Let op de volgorde: eerst de letter, dan het getal.', inp);
        else bad(ui.fb, 'Nog niet. Kijk naar de letter bovenaan en het getal links.', inp);
      };
      btn.addEventListener('click', check); inp.addEventListener('keydown', e => { if (e.key === 'Enter') check(); });
      next();
      return h('div', {}, ui.cnt, ui.ask, g.wrap, h('div', { class: 'read-row' }, inp, btn), ui.fb);
    },
    gridRange(w, onDone) {
      const g = miniGrid({ cellClick: true }); const ui = roundsUI(w.rounds);
      let i = 0, want, start = null, cur = null;
      const P = a => ({ c: COLS.indexOf(a[0]), r: +a.slice(1) });
      const mark = () => {
        g.clearMarks(); if (!start || !cur) return;
        const p = P(start), q = P(cur);
        const c1 = Math.min(p.c, q.c), c2 = Math.max(p.c, q.c), r1 = Math.min(p.r, q.r), r2 = Math.max(p.r, q.r);
        for (let c = c1; c <= c2; c++) for (let r = r1; r <= r2; r++) g.cell(COLS[c] + r).classList.add('insel');
        g.nb.textContent = COLS[c1] + r1 + (c1 !== c2 || r1 !== r2 ? ':' + COLS[c2] + r2 : '');
      };
      const next = () => {
        g.clearMarks(); g.nb.textContent = '';
        if (i >= w.rounds) { good(ui.fb, 'Goed gesleept! Een bereik schrijf je met een dubbelpunt.'); return onDone(); }
        const c = rnd(0, 6), r = rnd(1, 8), cw = rnd(1, 3), rh = rnd(1, 3);
        want = COLS[c] + r + ':' + COLS[c + cw] + (r + rh);
        ui.set(i, `Sleep om het bereik <span class="addr">${want}</span> te kiezen.`);
      };
      const at = e => { const el = document.elementFromPoint(e.clientX, e.clientY); return el && el.closest && el.closest('table.mini td'); };
      g.table.addEventListener('pointerdown', e => { const td = e.target.closest('td'); if (!td || i >= w.rounds) return; e.preventDefault(); start = cur = td.dataset.a; mark(); });
      window.addEventListener('pointermove', e => { if (!start) return; const td = at(e); if (td && g.table.contains(td) && td.dataset.a !== cur) { cur = td.dataset.a; mark(); } });
      window.addEventListener('pointerup', () => {
        if (!start) return;
        const got = g.nb.textContent; start = null;
        if (got === want) { good(ui.fb, `Juist: ${want}.`); i++; setTimeout(next, 900); }
        else bad(ui.fb, got.includes(':') ? `Je koos ${got}. Begin linksboven in ${want.split(':')[0]} en sleep tot rechtsonder in ${want.split(':')[1]}.` : 'Hou de muisknop ingedrukt en sleep over de cellen.', g.table);
      });
      next();
      return h('div', {}, ui.cnt, ui.ask, h('p', { class: 'small muted' }, 'Op een tablet: hou je vinger op de eerste cel en schuif.'), g.wrap, ui.fb);
    },
    keys(w, onDone) {
      const KEYS = ['Ctrl', 'Shift', 'Alt', 'Enter', 'Tab', 'Esc', 'Delete', 'F2', 'C', 'V', 'Z'];
      const ui = roundsUI(w.rounds.length);
      let i = 0; let sel = new Set();
      const combo = h('p', { class: 'combo' });
      const kb = h('div', { class: 'keys' }, KEYS.map(k => h('button', { 'aria-pressed': 'false', onclick: e => { sel.has(k) ? sel.delete(k) : sel.add(k); e.currentTarget.setAttribute('aria-pressed', String(sel.has(k))); combo.textContent = [...sel].join(' + '); } }, k)));
      const btn = h('button', { class: 'btn primary', onclick: () => {
        const want = w.rounds[i].answer;
        if (sel.size === want.length && want.every(k => sel.has(k))) {
          good(ui.fb, 'Juist: ' + want.join(' + ') + '.'); i++;
          setTimeout(() => { if (i >= w.rounds.length) { btn.disabled = true; kb.querySelectorAll('button').forEach(b => { b.disabled = true; }); onDone(); } else show(); }, 900);
        } else bad(ui.fb, sel.size ? 'Nog niet. Lees de uitleg hierboven nog eens.' : 'Klik eerst op de toets(en).', kb);
      } }, 'Controleer');
      const show = () => { sel = new Set(); combo.textContent = ''; kb.querySelectorAll('button').forEach(b => b.setAttribute('aria-pressed', 'false')); ui.set(i, w.rounds[i].ask); ui.fb.textContent = ''; };
      show();
      return h('div', {}, ui.cnt, ui.ask, kb, combo, h('div', { class: 'btn-row' }, btn), ui.fb);
    },
    autofit(w, onDone) {
      const data = [['Naam', 'Feest', 'Aantal', 'Waar?'], ['Lotte', 'verjaardagsfeest', '12', 'in de turnzaal van de school'], ['Milan', 'carnaval', '25', 'op de speelplaats']];
      const long = [1, 3];
      const widths = [70, 70, 70, 70];
      const fb = h('p', { class: 'fb' });
      const canvas = document.createElement('canvas').getContext('2d');
      const grid = h('div', { class: 'af' });
      let fitW = [];
      const measure = () => { const cs = getComputedStyle(grid); canvas.font = `${cs.fontSize} ${cs.fontFamily}`; fitW = [0, 1, 2, 3].map(c => Math.ceil(Math.max(...data.map(r => canvas.measureText(r[c]).width))) + 16); };
      let dragged = false;
      const draw = () => {
        clear(grid);
        const head = h('div', { class: 'af-row af-head' }, h('div', { class: 'af-cell af-rn' }));
        widths.forEach((wd, c) => {
          const handle = h('span', { class: 'af-handle' + (long.includes(c) ? ' pulse' : ''), title: 'Dubbelklik hier' });
          let lastUp = 0;
          handle.addEventListener('dblclick', () => { measure(); widths[c] = fitW[c]; draw(); after(c, true); });
          handle.addEventListener('pointerdown', e => {
            e.preventDefault(); const x0 = e.clientX, w0 = widths[c]; let moved = false;
            const mv = ev => { const d = ev.clientX - x0; if (Math.abs(d) > 3) moved = true; widths[c] = Math.max(24, w0 + d); grid.querySelectorAll('.af-row').forEach(row => { row.children[c + 1].style.width = widths[c] + 'px'; }); };
            const up = () => { window.removeEventListener('pointermove', mv); window.removeEventListener('pointerup', up);
              const now = Date.now();
              if (!moved && now - lastUp < 400 && e.pointerType !== 'mouse') { measure(); widths[c] = fitW[c]; draw(); after(c, true); return; }
              lastUp = now; if (moved) { dragged = true; draw(); after(c, false); } };
            window.addEventListener('pointermove', mv); window.addEventListener('pointerup', up);
          });
          head.append(h('div', { class: 'af-cell', style: `width:${wd}px` }, COLS[c], handle));
        });
        grid.append(head);
        data.forEach((row, r) => grid.append(h('div', { class: 'af-row' }, h('div', { class: 'af-cell af-rn' }, r + 1),
          row.map((v, c) => h('div', { class: 'af-cell' + (fitW[c] && long.includes(c) && widths[c] >= fitW[c] - 2 ? ' fit' : ''), style: `width:${widths[c]}px` }, v)))));
      };
      const after = (c, dbl) => {
        measure();
        const okAll = long.every(k => widths[k] >= fitW[k] - 2);
        if (okAll) { good(fb, dragged && !dbl ? 'Alles past! Slepen werkt, maar dubbelklikken is sneller en precies.' : 'Alles past! Dubbelklikken maakt de kolom precies breed genoeg.'); onDone(); }
        else if (dbl) good(fb, `Kolom ${COLS[c]} past nu precies. Is er nog een kolom met een te lang woord?`);
        else if (long.includes(c) && widths[c] < fitW[c] - 2) bad(fb, `Kolom ${COLS[c]} is nog te smal. Probeer eens te dubbelklikken op de lijn.`);
        else fb.textContent = '';
      };
      draw(); setTimeout(() => { measure(); }, 50);
      return h('div', {}, h('p', { class: 'ask' }, 'Maak kolom B en kolom D passend. Dubbelklik op de lijn rechts van de letter.'), h('div', { class: 'grid-wrap' }, grid), fb);
    },
    formatSim(w, onDone) {
      const FONT = [['automatisch', '#15241D'], ['rood', '#E00000'], ['blauw', '#0070C0'], ['groen', '#00A050'], ['paars', '#7030A0']];
      const FILL = [['geen', null], ['geel', '#FFE14D'], ['groen', '#9ED97A'], ['blauw', '#8EC9F2'], ['oranje', '#FFB766']];
      const H_AL = [['links', 'flex-start'], ['midden', 'center'], ['rechts', 'flex-end']];
      const V_AL = [['boven', 'flex-start'], ['midden', 'center'], ['onder', 'flex-end']];
      const words = ['leren', 'Excel', 'samen', 'proberen', 'cel'];
      const ui = roundsUI(w.rounds);
      let i = 0, target, mine;
      const base = () => ({ b: false, i: false, u: false, font: 0, fill: 0, h: 0, v: 2 });
      const tCell = h('div', { class: 'fcell', 'aria-label': 'Voorbeeld' }), mCell = h('div', { class: 'fcell', 'aria-label': 'Jouw cel' });
      const paint = (el, s, word) => {
        el.textContent = word;
        el.style.fontWeight = s.b ? '700' : '400'; el.style.fontStyle = s.i ? 'italic' : 'normal'; el.style.textDecoration = s.u ? 'underline' : 'none';
        el.style.color = FONT[s.font][1]; el.style.background = FILL[s.fill][1] || '#FFFFFF';
        el.style.justifyContent = H_AL[s.h][1]; el.style.alignItems = V_AL[s.v][1];
      };
      let word = '';
      const tog = (k, label, inner) => h('button', { class: 'rb', 'aria-pressed': 'false', 'aria-label': label, title: label, 'data-k': k, onclick: () => { mine[k] = !mine[k]; sync(); } }, inner);
      const sw = (k, idx, name, col) => h('button', { class: 'rb swatch' + (col ? '' : ' none'), 'aria-pressed': 'false', 'aria-label': name, title: name, 'data-k': k, 'data-i': idx, style: col ? `background:${col}` : null, onclick: () => { mine[k] = idx; sync(); } });
      const al = (k, idx, name, label) => h('button', { class: 'rb txt', 'aria-pressed': 'false', 'aria-label': name, title: name, 'data-k': k, 'data-i': idx, onclick: () => { mine[k] = idx; sync(); } }, label);
      const ribbon = h('div', { class: 'ribbon' },
        h('div', { class: 'rgroup' }, h('span', {}, 'Lettertype'), h('div', { class: 'rbtns' }, tog('b', 'Vet', h('b', {}, 'B')), tog('i', 'Schuin', h('i', { style: 'font-family:Georgia,serif' }, 'I')), tog('u', 'Onderstrepen', h('u', {}, 'U')))),
        h('div', { class: 'rgroup' }, h('span', {}, 'Letterkleur'), h('div', { class: 'rbtns' }, FONT.map(([n, c], k) => sw('font', k, 'Letterkleur ' + n, c)))),
        h('div', { class: 'rgroup' }, h('span', {}, 'Opvulkleur'), h('div', { class: 'rbtns' }, FILL.map(([n, c], k) => sw('fill', k, 'Opvulkleur ' + n, c)))),
        h('div', { class: 'rgroup' }, h('span', {}, 'Uitlijning'),
          h('div', { class: 'rbtns' }, al('v', 0, 'Boven uitlijnen', '⤒ boven'), al('v', 1, 'Midden uitlijnen', '↕ midden'), al('v', 2, 'Onder uitlijnen', '⤓ onder')),
          h('div', { class: 'rbtns' }, al('h', 0, 'Links uitlijnen', '⇤ links'), al('h', 1, 'Centreren', '↔ midden'), al('h', 2, 'Rechts uitlijnen', 'rechts ⇥'))));
      const sync = () => {
        paint(mCell, mine, word);
        ribbon.querySelectorAll('[data-k]').forEach(b => { const k = b.dataset.k; b.setAttribute('aria-pressed', String(b.dataset.i != null ? mine[k] === +b.dataset.i : !!mine[k])); });
      };
      const next = () => {
        if (i >= w.rounds) { good(ui.fb, 'Prachtig! Je kent de opmaakknoppen.'); chk.disabled = true; return onDone(); }
        word = words[rnd(0, words.length - 1)];
        target = { b: Math.random() < .6, i: Math.random() < .5, u: Math.random() < .4, font: rnd(0, 4), fill: rnd(0, 4), h: rnd(0, 2), v: rnd(0, 2) };
        if (!target.b && !target.i && !target.u) target.b = true;
        mine = base(); paint(tCell, target, word); sync(); ui.set(i, 'Maak jouw cel hetzelfde als het voorbeeld.'); ui.fb.textContent = '';
      };
      const chk = h('button', { class: 'btn primary', onclick: () => {
        const diff = [];
        if (mine.b !== target.b) diff.push('vet'); if (mine.i !== target.i) diff.push('schuin'); if (mine.u !== target.u) diff.push('onderstreept');
        if (mine.font !== target.font) diff.push('letterkleur'); if (mine.fill !== target.fill) diff.push('opvulkleur');
        if (mine.h !== target.h) diff.push('links-midden-rechts'); if (mine.v !== target.v) diff.push('boven-midden-onder');
        if (!diff.length) { good(ui.fb, 'Precies hetzelfde!'); i++; setTimeout(next, 900); }
        else bad(ui.fb, 'Kijk nog eens naar: ' + diff.join(', ') + '.', mCell);
      } }, 'Klaar?');
      next();
      return h('div', { class: 'fsim' }, ui.cnt, ui.ask, h('div', { class: 'fs-cells' }, h('figure', {}, h('figcaption', {}, 'Voorbeeld'), tCell), h('figure', {}, h('figcaption', {}, 'Jouw cel'), mCell)), ribbon, h('div', { class: 'btn-row' }, chk), ui.fb);
    },
    fillSim(w, onDone) {
      const { fillSeries } = window.CvcModules;
      const ui = roundsUI(w.rounds.length);
      const box = h('div', {});
      const restart = h('button', { class: 'btn small', onclick: () => startRound() }, '↺ Opnieuw');
      const ctrlNote = h('span', { class: 'small muted' });
      let i = 0, R, vals, sel, prev, lastFill, busy = false, ctrlDown = false;
      let table, frame, handle, optsBtn, optsMenu, nb, wrap, fxv = null;
      const A = (c, r) => COLS[c] + r;
      const F = window.CvcFormula;
      const show = v => F.display(v, 'nl');
      const expected = () => {
        const e = Object.assign({}, R.cells);
        if (R.goal) Object.assign(e, R.goal);
        if (R.goalSeries) R.goalSeries.values.forEach((v, k) => { e[R.goalSeries.col + (R.goalSeries.from + k)] = v; });
        if (R.goalBlock) for (let r = 0; r < 10; r++) for (let c = 0; c < 10; c++) e[A(c, r + 1)] = R.goalBlock.base + r * 10 + c + 1;
        return e;
      };
      const same = (a, b) => (a == null || a === '') ? (b == null || b === '') : String(a).toLowerCase() === String(b).toLowerCase();

      function build() {
        table = h('table', { class: 'mini' });
        const head = h('tr', {}, h('th', { class: 'corner' }));
        for (let c = 0; c < R.cols; c++) head.append(h('th', {}, COLS[c]));
        table.append(h('thead', {}, head));
        const tb = h('tbody');
        for (let r = 1; r <= R.rows; r++) { const tr = h('tr', {}, h('th', {}, r)); for (let c = 0; c < R.cols; c++) tr.append(h('td', { class: 'click', 'data-c': c, 'data-r': r })); tb.append(tr); }
        table.append(tb);
        frame = h('div', { class: 'sel-frame', hidden: true });
        handle = h('div', { class: 'fill-handle', hidden: true, title: 'Vulgreep', 'aria-label': 'Vulgreep' });
        optsMenu = h('div', { class: 'fill-menu', hidden: true },
          h('button', { onclick: () => reapply(false) }, '○ Cellen kopiëren'),
          h('button', { onclick: () => reapply(true) }, '○ Reeks doorvoeren'));
        optsBtn = h('button', { class: 'fill-opts', hidden: true, title: 'Opvulopties', onclick: () => { optsMenu.hidden = !optsMenu.hidden; } }, '⊞ ▾');
        wrap = h('div', { class: 'fs-wrap' }, table, frame, handle, optsBtn, optsMenu);
        nb = h('b', {});
        fxv = R.formulaBar || w.formulaBar ? h('span', { class: 'fx-v' }) : null;
        clear(box).append(h('div', { class: 'mini-namebox' }, 'Naamvak', nb, fxv ? h('span', { class: 'fx sim-fx' }, h('span', { class: 'fx-l' }, 'fx'), fxv) : null), h('div', { class: 'grid-wrap' }, wrap));
        table.addEventListener('pointerdown', e => {
          const td = e.target.closest('td'); if (!td || busy) return;
          e.preventDefault(); optsMenu.hidden = true;
          const c = +td.dataset.c, r = +td.dataset.r;
          sel = { c1: c, r1: r, c2: c, r2: r, ac: c, ar: r }; selecting = true; paint();
        });
        handle.addEventListener('pointerdown', e => { e.preventDefault(); e.stopPropagation(); if (!sel || busy || i >= w.rounds.length) return; filling = true; optsMenu.hidden = true; });
      }
      let selecting = false, filling = false;
      const tdAt = e => { const el = document.elementFromPoint(e.clientX, e.clientY); const td = el && el.closest && el.closest('td'); return td && table.contains(td) ? td : null; };
      // automatisch scrollen als je tijdens het slepen de rand van het scherm nadert
      let lastPt = null, scrollTimer = null;
      const autoScroll = () => {
        if (!(selecting || filling) || !lastPt) { scrollTimer = null; return; }
        const edge = 60, y = lastPt.clientY;
        const dy = y > innerHeight - edge ? 14 : y < edge + 60 ? -14 : 0;
        if (dy) { window.scrollBy(0, dy); onMove(lastPt); }
        scrollTimer = requestAnimationFrame(autoScroll);
      };
      function onMove(e) {
        if (!table || !document.body.contains(table)) return;
        if (selecting || filling) { lastPt = { clientX: e.clientX, clientY: e.clientY, ctrlKey: e.ctrlKey, metaKey: e.metaKey }; if (!scrollTimer) scrollTimer = requestAnimationFrame(autoScroll); }
        ctrlDown = e.ctrlKey || e.metaKey;
        const td = tdAt(e); if (!td) return;
        const c = +td.dataset.c, r = +td.dataset.r;
        if (selecting) { sel.c1 = Math.min(sel.ac, c); sel.c2 = Math.max(sel.ac, c); sel.r1 = Math.min(sel.ar, r); sel.r2 = Math.max(sel.ar, r); paint(); }
        else if (filling) {
          const dr = r - sel.r2, dc = c - sel.c2;
          prev = dr > 0 && dr >= dc ? { dir: 'down', n: dr } : dc > 0 ? { dir: 'right', n: dc } : null;
          paint();
          if (prev) { const out = compute(prev.dir, prev.n, ctrlDown); const lastKey = Object.keys(out).pop(); nb.textContent = show(out[lastKey]); }
        }
      }
      function onUp(e) {
        if (!table || !document.body.contains(table)) return;
        if (selecting) { selecting = false; paint(); return; }
        if (filling) { filling = false; ctrlDown = e.ctrlKey || e.metaKey; const p = prev; prev = null; if (p) apply(p.dir, p.n, ctrlDown); else paint(); }
      }
      window.addEventListener('pointermove', onMove);
      window.addEventListener('pointerup', onUp);

      const isF = v => typeof v === 'string' && v.startsWith('=');
      // reeks met formules: Excel kopieert het patroon en schuift de adressen mee
      const fillLine = (seeds, n, ctrl, horiz) => seeds.some(isF)
        ? Array.from({ length: n }, (_, t) => { const k = t % seeds.length, off = seeds.length + t - k, v = seeds[k]; return isF(v) ? F.shift(v, horiz ? 0 : off, horiz ? off : 0) : v; })
        : fillSeries(seeds, n, ctrl);
      function compute(dir, n, ctrl) {
        const out = {};
        if (dir === 'down') for (let c = sel.c1; c <= sel.c2; c++) {
          const seeds = []; for (let r = sel.r1; r <= sel.r2; r++) seeds.push(vals[A(c, r)] ?? null);
          fillLine(seeds, n, ctrl, false).forEach((v, k) => { if (sel.r2 + 1 + k <= R.rows) out[A(c, sel.r2 + 1 + k)] = v; });
        } else for (let r = sel.r1; r <= sel.r2; r++) {
          const seeds = []; for (let c = sel.c1; c <= sel.c2; c++) seeds.push(vals[A(c, r)] ?? null);
          fillLine(seeds, n, ctrl, true).forEach((v, k) => { if (sel.c2 + 1 + k < R.cols) out[A(sel.c2 + 1 + k, r)] = v; });
        }
        return out;
      }
      // uitrekenen van formules in de simulator
      const calc = (a, seen) => {
        const v = vals[a];
        if (!isF(v)) return v ?? null;
        seen = seen || new Set(); if (seen.has(a)) return new F.Err('#REF!'); seen.add(a);
        return F.evaluate(v, { locale: 'nl', get: x => calc(x, new Set(seen)) });
      };
      function apply(dir, n, ctrl) {
        lastFill = { dir, n, before: Object.assign({}, vals), sel: Object.assign({}, sel) };
        Object.entries(compute(dir, n, ctrl)).forEach(([k, v]) => { if (v == null) delete vals[k]; else vals[k] = v; });
        if (dir === 'down') sel.r2 = Math.min(R.rows, sel.r2 + n); else sel.c2 = Math.min(R.cols - 1, sel.c2 + n);
        optsBtn.hidden = false;
        paint(); evaluate(ctrl);
      }
      function reapply(ctrl) {
        if (!lastFill) return;
        vals = Object.assign({}, lastFill.before); sel = Object.assign({}, lastFill.sel); optsMenu.hidden = true;
        apply(lastFill.dir, lastFill.n, ctrl);
      }
      function paint() {
        table.querySelectorAll('td').forEach(td => {
          const c = +td.dataset.c, r = +td.dataset.r, a = A(c, r);
          const cv = calc(a);
          td.textContent = show(cv);
          td.style.textAlign = typeof cv === 'number' ? 'right' : (F.isErr(cv) ? 'center' : 'left');
          td.classList.toggle('err', F.isErr(cv));
          td.style.paddingInline = '5px';
          const inSel = sel && c >= sel.c1 && c <= sel.c2 && r >= sel.r1 && r <= sel.r2;
          td.classList.toggle('insel', !!inSel && !(sel.c1 === sel.c2 && sel.r1 === sel.r2));
          const inPrev = prev && sel && (prev.dir === 'down' ? c >= sel.c1 && c <= sel.c2 && r > sel.r2 && r <= sel.r2 + prev.n : r >= sel.r1 && r <= sel.r2 && c > sel.c2 && c <= sel.c2 + prev.n);
          td.classList.toggle('fillprev', !!inPrev);
        });
        if (!sel) { frame.hidden = handle.hidden = true; nb.textContent = ''; if (fxv) fxv.textContent = ''; return; }
        if (!filling) nb.textContent = A(sel.c1, sel.r1) + (sel.c1 !== sel.c2 || sel.r1 !== sel.r2 ? ':' + A(sel.c2, sel.r2) : '');
        if (fxv) { const raw = vals[A(sel.c1, sel.r1)]; fxv.textContent = raw == null ? '' : (typeof raw === 'number' ? show(raw) : raw); }
        const tl = table.querySelector(`td[data-c="${sel.c1}"][data-r="${sel.r1}"]`), br = table.querySelector(`td[data-c="${sel.c2}"][data-r="${sel.r2}"]`);
        const wr = wrap.getBoundingClientRect(), a = tl.getBoundingClientRect(), b = br.getBoundingClientRect();
        Object.assign(frame.style, { left: (a.left - wr.left - 1) + 'px', top: (a.top - wr.top - 1) + 'px', width: (b.right - a.left + 1) + 'px', height: (b.bottom - a.top + 1) + 'px' });
        frame.hidden = false;
        const hs = handle.offsetWidth || 10;
        Object.assign(handle.style, { left: (b.right - wr.left - hs / 2 - 1) + 'px', top: (b.bottom - wr.top - hs / 2 - 1) + 'px' });
        handle.hidden = false;
        Object.assign(optsBtn.style, { left: (b.right - wr.left + 6) + 'px', top: (b.bottom - wr.top + 4) + 'px' });
        Object.assign(optsMenu.style, { left: (b.right - wr.left + 6) + 'px', top: (b.bottom - wr.top + 34) + 'px' });
      }
      function evaluate(ctrl) {
        const exp = expected();
        const keys = new Set([...Object.keys(exp), ...Object.keys(vals)]);
        let wrong = 0, missing = 0, extra = 0;
        keys.forEach(k => {
          const e = exp[k], v = isF(exp[k]) ? vals[k] : calc(k);
          if (e == null && v != null && v !== '') extra++;
          else if (e != null && (v == null || v === '')) missing++;
          else if (e != null && !same(v, e)) wrong++;
        });
        const seeds = lastFill ? (() => { const s = lastFill.sel; const arr = []; for (let r = s.r1; r <= s.r2; r++) for (let c = s.c1; c <= s.c2; c++) arr.push(lastFill.before[A(c, r)]); return arr; })() : [];
        if (!wrong && !missing && !extra) {
          good(ui.fb, 'Juist! Excel deed het werk voor jou.'); busy = true; i++;
          setTimeout(() => {
            busy = false;
            if (i >= w.rounds.length) {
              if (w.keep) { restart.hidden = true; ui.cnt.textContent = ''; ui.ask.innerHTML = w.keep; good(ui.fb, w.doneMsg || 'Juist!'); onDone(); return; }
              clear(box).append(h('p', { class: 'fb ok' }, w.doneMsg || 'Super! Je kan de vulgreep gebruiken.')); restart.hidden = true; ui.ask.textContent = ''; ui.cnt.textContent = ''; onDone();
            } else startRound();
          }, 1100);
        } else if (wrong) {
          const nums = seeds.filter(v => typeof v === 'number');
          if (R.options && nums.length === 1 && seeds.length === 1 && !ctrl) bad(ui.fb, 'Excel kopieerde het getal. Klik op het knopje Opvulopties ⊞ en kies "Reeks doorvoeren", of klik op Opnieuw en hou Ctrl ingedrukt.', table);
          else if (seeds.length === 1 && typeof seeds[0] === 'number') bad(ui.fb, 'Excel kopieerde het getal. Selecteer de eerste twee getallen samen voor je sleept.', table);
          else bad(ui.fb, 'Dat is niet de reeks die we zoeken. Kijk welke cellen je selecteerde en klik op Opnieuw.', table);
        } else if (extra && !missing) bad(ui.fb, 'Je sleepte te ver. Klik op Opnieuw en laat eerder los.', table);
        else if (missing && !extra) { ui.fb.className = 'fb'; ui.fb.textContent = 'Goed bezig! Sleep nog verder.'; }
        else bad(ui.fb, 'Dat klopt nog niet. Klik op Opnieuw.', table);
      }
      function startRound() {
        R = w.rounds[i]; vals = Object.assign({}, R.cells); sel = null; prev = null; lastFill = null; busy = false;
        build(); paint(); optsBtn.hidden = true;
        ui.set(i, R.ask); ui.fb.className = 'fb'; ui.fb.textContent = '';
        ctrlNote.textContent = R.options ? 'Op een tablet: gebruik het knopje Opvulopties ⊞ dat na het slepen verschijnt.' : '';
      }
      startRound();
      return h('div', { class: 'q' }, ui.cnt, ui.ask, h('p', { class: 'small muted' }, 'Selecteer eerst de cel(len). Sleep dan het groene vierkantje rechtsonder.'), box, h('div', { class: 'btn-row' }, restart, ctrlNote), ui.fb);
    },
    sheetSim(w, onDone) {
      // Een klein rekenblad: cel aanklikken, typen in de formulebalk, Enter. Formules rekenen mee.
      const F = window.CvcFormula;
      const ui = roundsUI(w.rounds.length);
      const ROWS = w.rows || 6, NC = w.cols || 4;
      let i = 0, vals = {}, cur = 'A1', R;
      const isF = v => typeof v === 'string' && v.startsWith('=');
      const memo = {};
      const calc = (a, seen) => {
        const v = vals[a];
        if (!isF(v)) return v ?? null;
        seen = seen || new Set(); if (seen.has(a)) return new F.Err('#REF!'); seen.add(a);
        return F.evaluate(v, { locale: 'nl', get: x => calc(x, new Set(seen)) });
      };
      const raw = a => { const v = vals[a]; return v == null ? '' : typeof v === 'number' ? F.display(v, 'nl') : String(v); };
      const nb = h('b', {}, 'A1');
      const fx = h('input', { class: 'fx-in', id: 'fx-' + Math.random().toString(36).slice(2, 7), autocomplete: 'off', spellcheck: 'false', 'aria-label': 'Formulebalk' });
      const enter = h('button', { class: 'btn small', title: 'Bevestigen (Enter)' }, '↵ Enter');
      const table = h('table', { class: 'mini sheet-sim' });
      const head = h('tr', {}, h('th', { class: 'corner' }));
      for (let c = 0; c < NC; c++) head.append(h('th', { style: w.widths && w.widths[c] ? `min-width:${w.widths[c]}px` : null }, COLS[c]));
      table.append(h('thead', {}, head));
      const tb = h('tbody');
      for (let r = 1; r <= ROWS; r++) { const tr = h('tr', {}, h('th', {}, r)); for (let c = 0; c < NC; c++) tr.append(h('td', { class: 'click', 'data-a': COLS[c] + r })); tb.append(tr); }
      table.append(tb);
      const paint = () => {
        const refs = isF(fx.value) && document.activeElement === fx ? new Set((F.info(fx.value, 'nl').cells) || []) : new Set();
        table.querySelectorAll('td').forEach(td => {
          const a = td.dataset.a, v = calc(a);
          td.textContent = F.display(v, 'nl');
          td.style.textAlign = typeof v === 'number' ? 'right' : F.isErr(v) ? 'center' : 'left';
          td.style.paddingInline = '6px';
          td.classList.toggle('err', F.isErr(v));
          td.classList.toggle('cur', a === cur);
          td.classList.toggle('refhl', refs.has(a));
          td.classList.toggle('isformula', isF(vals[a]));
        });
        nb.textContent = cur;
      };
      const select = a => { cur = a; fx.value = raw(a); paint(); };
      table.addEventListener('pointerdown', e => {
        const td = e.target.closest('td'); if (!td) return;
        // tijdens het typen van een formule: adres invoegen, zoals in Excel
        if (document.activeElement === fx && isF(fx.value) && td.dataset.a !== cur && /[=+\-*/^(;:&<>]$/.test(fx.value.slice(0, fx.selectionStart).trim())) {
          e.preventDefault();
          const p = fx.selectionStart; fx.value = fx.value.slice(0, p) + td.dataset.a + fx.value.slice(fx.selectionEnd);
          fx.setSelectionRange(p + td.dataset.a.length, p + td.dataset.a.length); paint(); return;
        }
        e.preventDefault(); select(td.dataset.a); fx.focus();
      });
      fx.addEventListener('input', paint);
      fx.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); commit(); } else if (e.key === 'Escape') { fx.value = raw(cur); paint(); } });
      enter.addEventListener('click', commit);
      function commit() {
        const t = fx.value.trim(), a = cur;
        if (t === '') delete vals[a];
        else if (t.startsWith('=')) vals[a] = t.split('"').map((part, k) => k % 2 ? part : part.replace(/\s+/g, '')).join('"');
        else if (/^-?\d+(,\d+)?$/.test(t)) vals[a] = parseFloat(t.replace(',', '.'));
        else if (/^-?\d+(,\d+)?%$/.test(t)) vals[a] = parseFloat(t.replace(',', '.')) / 100;
        else vals[a] = t;
        const p = /([A-Z])(\d+)/.exec(a); const below = p[1] + Math.min(ROWS, +p[2] + 1);
        select(below === a ? a : below);
        check(a);
      }
      const api = {
        val: a => calc(a), raw: a => vals[a], isFormula: a => isF(vals[a]),
        info: a => isF(vals[a]) ? F.info(vals[a], 'nl') : null, num: a => { const v = calc(a); return typeof v === 'number' ? v : null; },
        isErr: a => F.isErr(calc(a)), show: a => F.display(calc(a), 'nl')
      };
      function check(changed) {
        if (!R) return;
        let r;
        try { r = R.check(api, changed); } catch (e) { r = null; }
        if (!r) { ui.fb.className = 'fb'; ui.fb.textContent = ''; return; }
        if (r.ok) {
          good(ui.fb, r.msg || 'Juist!'); const was = i; i++;
          setTimeout(() => { if (i !== was + 1) return; if (i >= w.rounds.length) { ui.ask.textContent = ''; ui.cnt.textContent = ''; good(ui.fb, w.doneMsg || 'Klaar! Je rekent als een echte Excel-kenner.'); onDone(); } else start(); }, 1300);
        } else bad(ui.fb, r.msg, table);
      }
      function start() {
        R = w.rounds[i];
        if (i === 0 || R.cells) { vals = {}; Object.assign(vals, w.cells || {}, R.cells || {}); }
        ui.set(i, R.ask); ui.fb.textContent = ''; ui.fb.className = 'fb';
        select(R.select || 'A1');
      }
      start();
      return h('div', { class: 'q' }, ui.cnt, ui.ask,
        h('p', { class: 'small muted' }, 'Klik op een cel en typ bovenaan in de formulebalk. Druk op Enter om te bevestigen.'),
        h('div', { class: 'fxrow' }, h('span', { class: 'mini-namebox' }, nb), h('span', { class: 'fx-l' }, 'fx'), fx, enter),
        h('div', { class: 'grid-wrap' }, table), ui.fb);
    },
    dvSim(w, onDone) {
      // Nagebootst venster "Gegevensvalidatie" + de cel om het resultaat uit te proberen.
      const ui = roundsUI(w.rounds.length);
      let i = 0, R, stage = 'setup';
      const cellBox = h('div', { class: 'dv-cell' });
      const label = h('div', { class: 'dv-label' });
      const allow = h('select', { class: 'in', id: 'dv-allow-' + Math.random().toString(36).slice(2, 6) },
        ['Alle waarden', 'Geheel getal', 'Decimaal', 'Lijst'].map(t => h('option', {}, t)));
      const oper = h('select', { class: 'in' }, ['tussen', 'groter dan', 'kleiner dan'].map(t => h('option', {}, t)));
      const min = h('input', { class: 'in', 'aria-label': 'Minimum' }), max = h('input', { class: 'in', 'aria-label': 'Maximum' });
      const bron = h('input', { class: 'in', 'aria-label': 'Bron', placeholder: 'bv. ja;nee' });
      const showErr = h('input', { type: 'checkbox', checked: true });
      const errTitle = h('input', { class: 'in', 'aria-label': 'Titel' }), errMsg = h('textarea', { class: 'in', rows: 2, 'aria-label': 'Foutbericht' });
      const numRow = h('div', { class: 'dv-grid' }, h('label', { class: 'f' }, 'Gegevens', oper), h('label', { class: 'f' }, 'Minimum', min), h('label', { class: 'f' }, 'Maximum', max));
      const listRow = h('label', { class: 'f' }, 'Bron (keuzes met een puntkomma ertussen)', bron);
      const tabA = h('div', { class: 'dv-pane' }, h('label', { class: 'f' }, 'Toestaan', allow), numRow, listRow);
      const tabB = h('div', { class: 'dv-pane', hidden: true }, h('label', { class: 'dv-check' }, showErr, ' Foutmelding weergeven als ongeldige gegevens zijn ingevoerd'), h('label', { class: 'f' }, 'Titel', errTitle), h('label', { class: 'f' }, 'Foutbericht', errMsg));
      const tA = h('button', { role: 'tab', 'aria-selected': 'true', onclick: () => { tabA.hidden = false; tabB.hidden = true; tA.setAttribute('aria-selected', 'true'); tB.setAttribute('aria-selected', 'false'); } }, 'Instellingen');
      const tB = h('button', { role: 'tab', 'aria-selected': 'false', onclick: () => { tabA.hidden = true; tabB.hidden = false; tA.setAttribute('aria-selected', 'false'); tB.setAttribute('aria-selected', 'true'); } }, 'Foutmelding');
      const okBtn = h('button', { class: 'btn primary', onclick: () => apply() }, 'OK');
      const dialog = h('div', { class: 'dv-dialog' }, h('div', { class: 'dv-title' }, 'Gegevensvalidatie'), h('div', { class: 'route-tabs dv-tabs' }, tA, tB), tabA, tabB, h('div', { class: 'btn-row' }, okBtn));
      const popup = h('div', { class: 'dv-popup', hidden: true });
      const sync = () => { const v = allow.value; numRow.hidden = !(v === 'Geheel getal' || v === 'Decimaal'); listRow.hidden = v !== 'Lijst'; };
      allow.addEventListener('change', sync);
      function apply() {
        const g = R.goal, v = allow.value;
        if (g.type === 'list') {
          if (v !== 'Lijst') return bad(ui.fb, 'Kies bij Toestaan: Lijst.', dialog);
          const raw = bron.value.trim();
          if (!raw) return bad(ui.fb, 'Vul de keuzes in bij Bron.', dialog);
          if (!raw.includes(';') && raw.includes(',')) return bad(ui.fb, 'In een Belgische Excel zet je een puntkomma tussen de keuzes, geen komma.', dialog);
          const items = raw.split(';').map(x => x.trim()).filter(Boolean);
          const want = g.items.map(x => x.toLowerCase()), got = items.map(x => x.toLowerCase());
          const miss = want.filter(x => !got.includes(x)), extra = got.filter(x => !want.includes(x));
          if (miss.length) return bad(ui.fb, `Deze keuze ontbreekt nog: ${miss.join(', ')}.`, dialog);
          if (extra.length) return bad(ui.fb, `Deze keuze hoort er niet bij: ${extra.join(', ')}.`, dialog);
          stage = 'use'; cellBox.replaceChildren(h('select', { class: 'in dv-select', onchange: e => { if (e.target.value.toLowerCase() === R.use.pick.toLowerCase()) { good(ui.fb, 'Juist! Met een keuzelijst kan je je niet meer vergissen bij het typen.'); next(); } else if (e.target.value) bad(ui.fb, 'Dat is een keuze, maar niet het juiste antwoord. Probeer nog eens.'); } }, h('option', { value: '' }, '▾ kies'), items.map(x => h('option', {}, x))));
          good(ui.fb, 'Goed ingesteld! Klik nu op het pijltje in de cel en kies het juiste antwoord.');
        } else {
          if (v !== 'Geheel getal') return bad(ui.fb, v === 'Decimaal' ? 'Kies Geheel getal: er mogen geen kommagetallen in.' : 'Kies bij Toestaan: Geheel getal.', dialog);
          if (oper.value !== 'tussen') return bad(ui.fb, 'Kies bij Gegevens: tussen.', dialog);
          if (+min.value.replace(',', '.') !== g.min || +max.value.replace(',', '.') !== g.max) return bad(ui.fb, `Het minimum moet ${g.min} zijn en het maximum ${g.max}.`, dialog);
          if (!showErr.checked || !errMsg.value.trim()) return bad(ui.fb, 'Ga naar het tabblad Foutmelding en typ een eigen foutbericht.', dialog);
          stage = 'use';
          const inp = h('input', { class: 'in dv-input', placeholder: 'typ en druk Enter' });
          inp.addEventListener('keydown', e => {
            if (e.key !== 'Enter') return;
            const n = Number(inp.value.replace(',', '.'));
            const okv = inp.value.trim() !== '' && Number.isInteger(n) && n >= g.min && n <= g.max;
            if (!okv) {
              popup.replaceChildren(h('b', {}, errTitle.value || 'Microsoft Excel'), h('p', {}, errMsg.value), h('button', { class: 'btn small', onclick: () => { popup.hidden = true; inp.value = ''; inp.focus(); } }, 'Opnieuw'));
              popup.hidden = false; R.sawError = true;
              good(ui.fb, 'Zie je? Excel toont jouw eigen foutmelding. Typ nu een goed antwoord.');
            } else if (!R.sawError) { ui.fb.className = 'fb'; ui.fb.textContent = `Dat getal mag. Typ eerst eens ${R.use.bad} om je foutmelding te zien.`; inp.value = ''; }
            else { good(ui.fb, 'Juist! Excel laat alleen goede getallen toe.'); next(); }
          });
          cellBox.replaceChildren(inp); inp.focus();
          good(ui.fb, `Goed ingesteld! Typ nu eens ${R.use.bad} in de cel en druk op Enter.`);
        }
      }
      function next() { i++; setTimeout(() => { if (i >= w.rounds.length) { okBtn.disabled = true; ui.ask.textContent = ''; ui.cnt.textContent = ''; good(ui.fb, 'Super! Je kan gegevensvalidatie instellen.'); onDone(); } else start(); }, 1400); }
      function start() {
        R = w.rounds[i]; stage = 'setup'; popup.hidden = true;
        allow.value = 'Alle waarden'; bron.value = ''; min.value = ''; max.value = ''; errTitle.value = ''; errMsg.value = ''; showErr.checked = true; oper.value = 'tussen';
        tA.click(); sync();
        label.textContent = R.label; cellBox.replaceChildren(h('span', { class: 'muted small' }, '(leeg)'));
        ui.set(i, R.ask); ui.fb.textContent = ''; ui.fb.className = 'fb';
      }
      start();
      return h('div', { class: 'q' }, ui.cnt, ui.ask,
        h('div', { class: 'dv-wrap' }, h('div', { class: 'dv-sheet' }, label, h('div', { class: 'dv-cellwrap' }, cellBox, popup)), dialog), ui.fb);
    },
    cfSim(w, onDone) {
      // Regelbouwer voor voorwaardelijke opmaak, met meteen zichtbare kleuren.
      const F = window.CvcFormula;
      const PRESETS = [
        { name: 'Lichtrode opvulling met donkerrode tekst', fill: '#FFC7CE', font: '#9C0006', col: 'rood' },
        { name: 'Gele opvulling met donkergele tekst', fill: '#FFEB9C', font: '#9C5700', col: 'geel' },
        { name: 'Groene opvulling met donkergroene tekst', fill: '#C6EFCE', font: '#006100', col: 'groen' },
        { name: 'Rode tekst', fill: null, font: '#E00000', col: 'rood' }
      ];
      const TYPES = ['Groter dan…', 'Kleiner dan…', 'Gelijk aan…', 'Tekst die bevat…'];
      const ui = roundsUI(w.rounds.length);
      let i = 0, R, rules = [];
      const typeSel = h('select', { class: 'in' }, TYPES.map(t => h('option', {}, t)));
      const valIn = h('input', { class: 'in', 'aria-label': 'Waarde', placeholder: 'waarde' });
      const fmtSel = h('select', { class: 'in' }, PRESETS.map((p, k) => h('option', { value: k }, p.name)));
      const list = h('ul', { class: 'cf-rules' });
      const table = h('table', { class: 'mini cf-table' });
      const parseVal = t => { t = String(t).trim(); if (/^-?\d+(,\d+)?%$/.test(t)) return parseFloat(t.replace(',', '.')) / 100; if (/^-?\d+(,\d+)?$/.test(t)) return parseFloat(t.replace(',', '.')); return t; };
      const showVal = (v, pct) => typeof v === 'number' ? (pct ? F.display(+(v * 100).toFixed(4), 'nl') + '%' : F.display(v, 'nl')) : v;
      const match = (r, v) => {
        if (r.type === 'Tekst die bevat…') return String(v).toLowerCase().includes(String(r.val).toLowerCase());
        const c = F.compare(v, r.val);
        return r.type === 'Groter dan…' ? (typeof v === typeof r.val && c > 0) : r.type === 'Kleiner dan…' ? (typeof v === typeof r.val && c < 0) : c === 0;
      };
      const colorOf = v => { const r = rules.find(x => match(x, v)); return r ? PRESETS[r.fmt] : null; };
      function paint() {
        clear(table).append(h('thead', {}, h('tr', {}, h('th', { class: 'corner' }), h('th', {}, 'A'))));
        const tb = h('tbody');
        tb.append(h('tr', {}, h('th', {}, 1), h('td', { class: 'cf-head' }, R.head)));
        R.values.forEach((v, k) => { const p = colorOf(v); tb.append(h('tr', {}, h('th', {}, k + 2), h('td', { style: `text-align:${typeof v === 'number' ? 'right' : 'left'};padding-inline:8px;${p ? `background:${p.fill || 'transparent'};color:${p.font};` : ''}` }, showVal(v, R.pct)))); });
        table.append(tb);
        clear(list).append(...rules.map((r, k) => h('li', {}, h('span', {}, `${r.type.replace('…', '')} ${showVal(r.val, R.pct && typeof r.val === 'number')} → ${PRESETS[r.fmt].name}`), h('button', { class: 'btn ghost small', onclick: () => { rules.splice(k, 1); paint(); } }, '✕'))));
        if (!rules.length) list.append(h('li', { class: 'muted small' }, 'Nog geen regels.'));
      }
      const add = h('button', { class: 'btn', onclick: () => {
        if (!valIn.value.trim()) return bad(ui.fb, 'Vul een waarde in.', valIn);
        rules.push({ type: typeSel.value, val: parseVal(valIn.value), fmt: +fmtSel.value }); valIn.value = ''; ui.fb.textContent = ''; paint();
      } }, 'Regel toevoegen');
      const chk = h('button', { class: 'btn primary', onclick: () => {
        const wrong = [];
        R.values.forEach((v, k) => { const p = colorOf(v), want = R.goal(v); const got = p ? p.col : null; if (got !== want) wrong.push({ v, want, got }); });
        if (!wrong.length) { good(ui.fb, R.okMsg || 'Juist! De kleuren tonen meteen wat je moet weten.'); i++; setTimeout(() => { if (i >= w.rounds.length) { chk.disabled = add.disabled = true; ui.ask.textContent = ''; ui.cnt.textContent = ''; good(ui.fb, 'Knap! Je kan voorwaardelijke opmaak gebruiken.'); onDone(); } else start(); }, 1400); return; }
        const x = wrong[0];
        const allNum = R.pct && rules.some(r => typeof r.val === 'number' && r.val >= 1);
        if (allNum) return bad(ui.fb, 'Je procenten zijn eigenlijk getallen tussen 0 en 1. Typ in de regel 50% of 0,5, niet 50.', table);
        bad(ui.fb, `${showVal(x.v, R.pct)} ${x.got ? 'wordt nu ' + x.got : 'krijgt nog geen kleur'}${x.want ? `, maar moet ${x.want} worden.` : ', maar moet geen kleur krijgen.'}`, table);
      } }, 'Klaar?');
      function start() { R = w.rounds[i]; rules = []; paint(); ui.set(i, R.ask); ui.fb.textContent = ''; ui.fb.className = 'fb'; }
      start();
      return h('div', { class: 'q' }, ui.cnt, ui.ask,
        h('div', { class: 'cf-wrap' }, h('div', { class: 'grid-wrap' }, table),
          h('div', { class: 'cf-builder' }, h('div', { class: 'dv-title' }, 'Voorwaardelijke opmaak › Markeringsregels voor cellen'),
            h('label', { class: 'f' }, 'Regel', typeSel), h('label', { class: 'f' }, 'Waarde', valIn), h('label', { class: 'f' }, 'met', fmtSel),
            h('div', { class: 'btn-row' }, add), h('div', { class: 'small' }, h('b', {}, 'Regels:')), list)),
        h('div', { class: 'btn-row' }, chk), ui.fb);
    },
    protectSim(w, onDone) {
      // Cellen ontgrendelen, blad beveiligen en uitproberen.
      const ui = roundsUI(1);
      const ROWS = w.rows, NC = w.cols;
      const vals = Object.assign({}, w.cells);
      const unlocked = new Set();
      let prot = false, sel = null, anchor = null, dragging = false, stage = 'setup', triedLocked = false, typedOpen = false;
      const A = (c, r) => COLS[c] + r;
      const table = h('table', { class: 'mini prot-table' });
      const head = h('tr', {}, h('th', { class: 'corner' }));
      for (let c = 0; c < NC; c++) head.append(h('th', {}, COLS[c]));
      table.append(h('thead', {}, head));
      const tb = h('tbody');
      for (let r = 1; r <= ROWS; r++) { const tr = h('tr', {}, h('th', {}, r)); for (let c = 0; c < NC; c++) tr.append(h('td', { class: 'click', 'data-c': c, 'data-r': r })); tb.append(tr); }
      table.append(tb);
      const nb = h('b', {});
      const lockBox = h('input', { type: 'checkbox', id: 'lock-' + Math.random().toString(36).slice(2, 6) });
      const protBtn = h('button', { class: 'btn', onclick: () => toggleProt() }, 'Blad beveiligen');
      const status = h('span', { class: 'chip' }, 'Blad niet beveiligd');
      const typeIn = h('input', { class: 'in', placeholder: 'typ in de gekozen cel en druk Enter', 'aria-label': 'Typ in de cel' });
      const popup = h('div', { class: 'dv-popup prot-pop', hidden: true });
      const selCells = () => { const out = []; if (!sel) return out; for (let r = sel.r1; r <= sel.r2; r++) for (let c = sel.c1; c <= sel.c2; c++) out.push(A(c, r)); return out; };
      function paint() {
        table.querySelectorAll('td').forEach(td => {
          const c = +td.dataset.c, r = +td.dataset.r, a = A(c, r);
          td.textContent = vals[a] ?? '';
          td.style.textAlign = typeof vals[a] === 'number' ? 'right' : 'left'; td.style.paddingInline = '6px';
          td.classList.toggle('insel', !!sel && c >= sel.c1 && c <= sel.c2 && r >= sel.r1 && r <= sel.r2);
          td.classList.toggle('unlocked', unlocked.has(a));
        });
        const cs = selCells();
        nb.textContent = cs.length ? (cs.length > 1 ? cs[0] + ':' + cs[cs.length - 1] : cs[0]) : '';
        lockBox.checked = cs.length ? cs.every(a => !unlocked.has(a)) : true;
        lockBox.disabled = prot || !cs.length;
        protBtn.textContent = prot ? 'Beveiliging opheffen' : 'Blad beveiligen';
        status.textContent = prot ? '🔒 Blad beveiligd' : 'Blad niet beveiligd';
        status.className = 'chip' + (prot ? ' ok' : '');
      }
      const tdAt = e => { const el = document.elementFromPoint(e.clientX, e.clientY); const td = el && el.closest && el.closest('td'); return td && table.contains(td) ? td : null; };
      table.addEventListener('pointerdown', e => { const td = e.target.closest('td'); if (!td) return; e.preventDefault(); anchor = { c: +td.dataset.c, r: +td.dataset.r }; sel = { c1: anchor.c, c2: anchor.c, r1: anchor.r, r2: anchor.r }; dragging = true; popup.hidden = true; paint(); });
      window.addEventListener('pointermove', e => { if (!dragging || !document.body.contains(table)) return; const td = tdAt(e); if (!td) return; const c = +td.dataset.c, r = +td.dataset.r; sel = { c1: Math.min(anchor.c, c), c2: Math.max(anchor.c, c), r1: Math.min(anchor.r, r), r2: Math.max(anchor.r, r) }; paint(); });
      window.addEventListener('pointerup', () => { dragging = false; });
      lockBox.addEventListener('change', () => { selCells().forEach(a => lockBox.checked ? unlocked.delete(a) : unlocked.add(a)); paint(); ui.fb.textContent = ''; });
      function toggleProt() {
        if (prot) { prot = false; stage = 'setup'; paint(); ui.fb.className = 'fb'; ui.fb.textContent = 'De beveiliging is opgeheven. Je kan nu weer cellen ontgrendelen.'; return; }
        prot = true; paint();
        const want = new Set(w.open), got = [...unlocked];
        const extra = got.filter(a => !want.has(a)), miss = [...want].filter(a => !unlocked.has(a));
        if (!got.length) return bad(ui.fb, 'Het blad is beveiligd, maar ALLE cellen zijn nog geblokkeerd. De gebruiker kan zelfs geen antwoord typen! Hef de beveiliging op en ontgrendel eerst de antwoordcellen.', table);
        if (miss.length) return bad(ui.fb, `Nog niet alle antwoordcellen zijn ontgrendeld: ${miss.join(', ')}. Hef de beveiliging op en ontgrendel ze.`, table);
        if (extra.length) return bad(ui.fb, `Ook ${extra.join(', ')} ${extra.length > 1 ? 'zijn' : 'is'} ontgrendeld. Dan kan de gebruiker daar iets veranderen. Hef de beveiliging op en blokkeer ze weer.`, table);
        stage = 'try'; good(ui.fb, `Goed beveiligd! Probeer het uit: klik op een antwoordcel (${w.open[0]}) en typ iets. Probeer daarna iets te typen in een oplossing (kolom B).`);
      }
      typeIn.addEventListener('keydown', e => {
        if (e.key !== 'Enter') return;
        const cs = selCells(); if (cs.length !== 1) return bad(ui.fb, 'Klik eerst op één cel.', table);
        const a = cs[0];
        if (prot && !unlocked.has(a)) {
          popup.replaceChildren(h('b', {}, 'Microsoft Excel'), h('p', {}, 'De cel die je probeert te wijzigen, staat op een beveiligd blad.'), h('button', { class: 'btn small', onclick: () => { popup.hidden = true; } }, 'OK'));
          popup.hidden = false; triedLocked = true;
        } else { const t = typeIn.value.trim(); vals[a] = /^-?\d+(,\d+)?$/.test(t) ? parseFloat(t.replace(',', '.')) : t; if (w.open.includes(a)) typedOpen = true; paint(); }
        typeIn.value = '';
        if (stage === 'try') {
          if (typedOpen && triedLocked) { good(ui.fb, 'Precies zo! De gebruiker kan antwoorden typen, maar niets anders kapotmaken.'); stage = 'done'; protBtn.disabled = true; onDone(); }
          else if (typedOpen) good(ui.fb, 'Antwoorden typen lukt. Probeer nu iets te typen in een oplossing (kolom B).');
          else if (triedLocked) good(ui.fb, 'Zie je? Excel beschermt die cel. Typ nu een antwoord in een antwoordcel.');
        }
      });
      paint(); ui.set(0, w.ask);
      return h('div', { class: 'q' }, ui.ask,
        h('p', { class: 'small muted' }, 'Selecteer cellen door erover te slepen. Ontgrendelde cellen krijgen een stippellijn.'),
        h('div', { class: 'ribbon prot-ribbon' },
          h('div', { class: 'rgroup' }, h('span', {}, '1. Celeigenschappen › Bescherming'), h('label', { class: 'dv-check', for: lockBox.id }, lockBox, ' Geblokkeerd')),
          h('div', { class: 'rgroup' }, h('span', {}, '2. Controleren'), h('div', { class: 'btn-row' }, protBtn, status))),
        h('div', { style: 'display:grid;gap:8px;min-width:0' }, h('div', { class: 'mini-namebox' }, 'Naamvak', nb), h('div', { class: 'grid-wrap' }, table),
          h('div', { class: 'dv-cellwrap prot-type' }, typeIn, popup)),
        ui.fb);
    },
    numSim(w, onDone) {
      const nl = window.CvcModules.nlNum;
      const ui = roundsUI(w.rounds.length);
      let i = 0, R, st;
      const decShown = v => { const t = String(+(+v).toPrecision(10)); return t.includes('.') ? t.split('.')[1].length : 0; };
      const disp = (v, s) => s.kind === 'percent' ? nl(v * 100, s.dec) + '%' : s.kind === 'currency' ? '€ ' + nl(v, s.dec) : s.kind === 'number' ? nl(v, s.dec) : nl(v);
      const bar = h('div', { class: 'fx' }, h('span', { class: 'fx-l' }, 'fx'), h('span', { class: 'fx-v' }));
      const tBar = h('div', { class: 'fx' }, h('span', { class: 'fx-l' }, 'fx'), h('span', { class: 'fx-v' }));
      const tCell = h('div', { class: 'ncell' }), mCell = h('div', { class: 'ncell' });
      const kindLbl = h('span', { class: 'nkind' });
      const NAMES = { general: 'Standaard', number: 'Getal', percent: 'Percentage', currency: 'Valuta' };
      const sync = () => {
        mCell.textContent = disp(R.v, st);
        bar.lastChild.textContent = st.kind === 'percent' ? nl(R.v * 100) + '%' : nl(R.v);
        kindLbl.textContent = NAMES[st.kind];
      };
      const act = f => () => { f(); sync(); ui.fb.textContent = ''; };
      const btn = (label, title, f) => h('button', { class: 'rb txt', title, 'aria-label': title, onclick: act(f) }, label);
      const group = h('div', { class: 'ribbon' }, h('div', { class: 'rgroup' }, h('span', {}, 'Getal'),
        h('div', { class: 'rbtns' }, kindLbl),
        h('div', { class: 'rbtns' },
          btn('€', 'Valuta (euro)', () => { st = { kind: 'currency', dec: 2 }; }),
          btn('%', 'Procent', () => { st = { kind: 'percent', dec: 0 }; }),
          btn('←,0', 'Meer decimalen', () => { if (st.kind === 'general') st = { kind: 'number', dec: Math.min(9, decShown(R.v) + 1) }; else st.dec = Math.min(9, st.dec + 1); }),
          btn(',00→', 'Minder decimalen', () => { if (st.kind === 'general') st = { kind: 'number', dec: Math.max(0, decShown(R.v) - 1) }; else st.dec = Math.max(0, st.dec - 1); }),
          btn('Standaard', 'Standaard', () => { st = { kind: 'general', dec: null }; }))));
      const chk = h('button', { class: 'btn primary', onclick: () => {
        if (mCell.textContent === R.target) { good(ui.fb, `Juist! Je ziet ${R.target}, maar in de cel zit nog altijd ${nl(R.v)}.`); i++; setTimeout(next, 1400); }
        else bad(ui.fb, `Je cel toont ${mCell.textContent}. Het moet ${R.target} zijn.`, mCell);
      } }, 'Klaar?');
      function next() {
        if (i >= w.rounds.length) { good(ui.fb, 'Knap! Je weet nu dat weergave en waarde twee verschillende dingen zijn.'); chk.disabled = true; return onDone(); }
        R = w.rounds[i]; st = { kind: 'general', dec: null };
        tCell.textContent = R.target; tBar.lastChild.textContent = /%$/.test(R.target) ? nl(R.v * 100) + '%' : nl(R.v); sync(); ui.set(i, 'Laat jouw cel er net zo uitzien als het voorbeeld.'); ui.fb.textContent = '';
      }
      next();
      return h('div', { class: 'fsim' }, ui.cnt, ui.ask,
        h('div', { class: 'fs-cells' }, h('figure', {}, h('figcaption', {}, 'Voorbeeld'), tBar, tCell), h('figure', {}, h('figcaption', {}, 'Jouw cel'), bar, mCell)),
        group, h('div', { class: 'btn-row' }, chk), ui.fb);
    },
    steps(w, onDone) {
      const on = new Set();
      const btn = h('button', { class: 'btn primary', disabled: true, onclick: () => { btn.disabled = true; onDone(); } }, w.confirm);
      const list = h('ol', { class: 'steps' }, w.steps.map((s, k) => h('li', { tabindex: '0', role: 'checkbox', 'aria-checked': 'false', onclick: e => tog(e.currentTarget, k), onkeydown: e => { if (e.key === ' ' || e.key === 'Enter') { e.preventDefault(); tog(e.currentTarget, k); } } }, h('span', { class: 'bx' }), h('span', { html: s }))));
      const tog = (li, k) => { on.has(k) ? on.delete(k) : on.add(k); li.classList.toggle('on', on.has(k)); li.querySelector('.bx').textContent = on.has(k) ? '✓' : ''; li.setAttribute('aria-checked', String(on.has(k))); btn.disabled = on.size !== w.steps.length; };
      return h('div', { class: 'q' }, h('p', { class: 'small muted' }, 'Vink elke stap aan als je hem begrijpt.'), list, h('div', { class: 'btn-row' }, btn));
    },
    levelPick(w, onDone) {
      const m = MODULES[modIndex(S.mid)];
      const box = h('div', { style: 'display:grid;gap:10px' });
      const draw = () => {
        const cur = levelNow(m);
        clear(box).append(levelPicker(m, cur, async n => { await setLevel(m, n); draw(); onDone(); }),
          h('p', { class: 'small muted' }, cur ? `Je koos ${levelInfo(m, cur).stars} ${levelInfo(m, cur).name}. Wijzigen kan later nog, bij ${m.tabLabels?.practice || 'Oefenen'}.` : 'Klik op het niveau dat bij jou past.'));
      };
      draw();
      return box;
    },
    read(w, onDone) {
      const btn = h('button', { class: 'btn primary', onclick: () => { btn.disabled = true; onDone(); } }, w.confirm || 'Gelezen');
      return h('div', { class: 'btn-row' }, btn);
    }
  };

  // ---------------- oefenen en test ----------------
  function levelPicker(m, cur, onPick) {
    return h('div', { class: 'levels', role: 'radiogroup', 'aria-label': 'Kies je niveau' }, m.levels.map(L =>
      h('button', { class: 'level' + (L.n === cur ? ' on' : ''), role: 'radio', 'aria-checked': String(L.n === cur), onclick: () => onPick(L.n) },
        h('span', { class: 'lv-stars', 'aria-hidden': 'true' }, L.stars), h('b', {}, L.name), h('span', { class: 'small' }, L.short),
        h('ul', {}, L.items.map(x => h('li', {}, x))))));
  }
  async function setLevel(m, n) {
    const p = prog(m.id);
    if (p.practice?.seed && m.levelOf(p.practice.seed) === n) return;
    await save(m.id, { practice: { seed: m.withLevel(n), started: p.practice?.started || new Date().toISOString(), passed: false, lastSeed: null, lastResults: null } });
  }
  const passedOf = (m, mode, tasks, results) => {
    const sc = results.filter(r => r.ok).length;
    if (mode === 'practice') return sc === results.length;
    return m.passCheck ? m.passCheck(tasks, results) : Math.round(sc / results.length * 100) >= PASS * 100;
  };
  async function ensureSeed(m, mode) {
    const p = prog(m.id);
    if (mode === 'practice' && !p.practice?.seed && m.levels) return null;     // eerst een niveau kiezen
    if (mode === 'practice' && !p.practice?.seed) await save(m.id, { practice: { seed: newSeed(), started: new Date().toISOString() } });
    return mode === 'practice' ? prog(m.id).practice.seed : p.test?.seed;
  }
  function fileWord(m, mode) { if (m.fileWord) return m.fileWord; return m.project && mode === 'test' ? 'rapport' : mode === 'test' ? 'test' + m.nr : 'module' + m.nr; }

  async function downloadStart(m, mode, seed) {
    try {
      const sf = m.startFile(seed, mode);
      const blob = await window.XlsxWrite.buildXlsx(sf.spec, window.JSZip);
      Api.download(blob, sf.name);
    } catch (e) { toast('Het startbestand kon niet gemaakt worden: ' + e.message, true); }
  }
  async function downloadLast(m) {
    try { const d = await Api.lastFile(m.id); Api.download(Api.b64ToBlob(d.b64), d.name || (m.id + '.xlsx')); }
    catch (e) { toast(e.message, true); }
  }

  async function renderWork(m, root, mode) {
    const p = prog(m.id);
    if (mode === 'test' && !p.test?.open) return renderTestGate(m, root);
    const seed = await ensureSeed(m, mode);
    if (!seed && m.levels) {
      root.append(h('section', { class: 'sheet pad', style: 'display:grid;gap:12px' }, h('h2', {}, 'Kies eerst je niveau'),
        h('p', {}, 'Welk werkblad ga je maken? Je kan later nog wisselen.'),
        levelPicker(m, null, async n => { await setLevel(m, n); render(); })));
      return;
    }
    const gen = m.gen(seed, mode, S.student);
    const st = mode === 'practice' ? (prog(m.id).practice || {}) : (prog(m.id).test || {});
    const last = S.results[m.id + mode + seed] || (st.lastSeed === seed ? st.lastResults : null);
    const resMap = {}; (last || []).forEach(r => { resMap[r.id] = r; });
    const vn = S.student.voornaam.replace(/\s+/g, '');
    const fname = `${vn}_${fileWord(m, mode)}`;

    // --- niveau ---
    let lvSec = null;
    if (m.levels) {
      const L = levelInfo(m, m.levelOf(seed));
      const pick = h('div', { hidden: true });
      lvSec = h('section', { class: 'sheet pad lv-now', style: 'display:grid;gap:10px' },
        h('div', { class: 'lv-row' }, h('span', { class: 'lv-stars', 'aria-hidden': 'true' }, L.stars), h('div', {}, h('b', {}, (mode === 'test' ? 'Je levert in op niveau ' : 'Mijn niveau: ') + L.name), h('p', { class: 'small muted' }, L.short)),
          mode === 'practice' ? h('button', { class: 'btn small', onclick: e => { pick.hidden = !pick.hidden; e.currentTarget.textContent = pick.hidden ? 'Ander niveau' : 'Sluiten'; } }, 'Ander niveau') : null),
        mode === 'practice' ? pick : null);
      if (mode === 'practice') pick.append(levelPicker(m, L.n, async n => { await setLevel(m, n); render(); }), h('p', { class: 'small muted' }, 'Na een wissel laad je je werkblad gewoon opnieuw op.'));
    }
    // --- zo begin je ---
    const start = h('section', { class: 'sheet pad', style: 'display:grid;gap:12px' },
      h('h2', {}, mode === 'test' ? (m.tabLabels?.test || 'Je test') : 'Zo begin je'),
      mode === 'test' ? h('p', { class: 'msg info' }, m.testIntro ? m.testIntro : [m.project ? 'Geen hulplijnen. De app kijkt na of je rapport goed rekent en goed opgebouwd is. Je hebt ' : 'Dit is de test: er zijn geen hulplijnen. Lees elke opdracht goed. Je hebt ', h('b', {}, Math.round(PASS * 100) + '%'), ' nodig.']) : null,
      m.route ? h('ol', { class: 'route' }, m.route[mode].map(x => h('li', { html: x.replace('{fname}', fname) })))
      : m.newFile
        ? h('ol', { class: 'route' }, h('li', { html: 'Start <b>Excel</b> en kies <b>Lege werkmap</b>.' }), h('li', { html: 'Doe de opdrachten hieronder.' }), h('li', { html: `Bewaar via <b>Bestand → Opslaan als → OneDrive</b> met de naam <span class="addr">${fname}</span>.` }))
        : h('div', { style: 'display:grid;gap:10px' },
          h('ol', { class: 'route' },
            h('li', { html: 'Download het startbestand.' }),
            h('li', { html: 'Open het in Excel. Zie je bovenaan een gele balk? Klik op <b>Bewerken inschakelen</b>.' }),
            h('li', { html: `Bewaar het meteen via <b>Bestand → Opslaan als → OneDrive</b> met de naam <span class="addr">${fname}</span>. Werk verder in dat bestand.` })),
          h('div', { class: 'btn-row' }, h('button', { class: 'btn primary', onclick: () => downloadStart(m, mode, seed) }, '⤓ Download het startbestand'))),
      p.lastFile ? h('p', { class: 'small' }, 'Op een ander toestel? ', h('button', { class: 'btn ghost small', onclick: () => downloadLast(m) }, 'Download je laatst opgeladen bestand'), h('span', { class: 'muted' }, ` (${fmtDate(p.lastFile.at)})`)) : null);

    // --- opdrachten ---
    const hintsUsed = (prog(m.id).practice || {}).hints || {};
    const tasks = h('ol', { class: 'tasks', 'aria-label': 'Opdrachten' }, gen.tasks.map((t, k) => {
      const r = resMap[t.id];
      const hintBox = h('div', { style: 'display:grid;gap:6px' });
      const lvl = mode === 'practice' ? (hintsUsed[t.id] || 0) : 0;
      const drawHints = level => {
        clear(hintBox);
        t.hints.slice(0, level).forEach((ht, j) => hintBox.append(h('p', { class: 'hint' }, h('b', {}, j === 0 ? 'Tip: ' : 'Waar klik ik? '), ht)));
        if (mode === 'practice' && level < t.hints.length) {
          hintBox.append(h('div', { class: 'hints' }, h('button', { class: 'btn small', onclick: async () => {
            const nl = level + 1; drawHints(nl);
            await save(m.id, { practice: { hints: { [t.id]: nl } } });
          } }, level === 0 ? '💡 Tip' : '👆 Waar klik ik?')));
        }
      };
      drawHints(lvl);
      return h('li', { class: 'task' + (r ? (r.ok ? ' ok' : ' bad') : '') },
        h('div', { class: 'tn' }, r ? (r.ok ? '✓' : '✗') : k + 1),
        h('div', { class: 'tb' }, t.lvl > 1 ? h('span', { class: 'lvl-tag', title: 'vanaf niveau ' + '★'.repeat(t.lvl) }, '★'.repeat(t.lvl)) : null, h('div', { class: 'ttext', html: t.text }), r ? h('p', { class: 'res' }, r.ok ? (m.project ? (r.msg && r.msg !== 'Juist!' ? r.msg : 'In orde!') : 'Juist!') : r.msg) : null, hintBox));
    }));

    // --- opladen ---
    const side = h('aside', { class: 'side' });
    const total = gen.tasks.length;
    if (last) {
      const sc = last.filter(r => r.ok).length, pct = Math.round(sc / total * 100);
      const passed = passedOf(m, mode, gen.tasks, last);
      const noun = m.noun || 'rapport', TLt = (m.tabLabels?.test || 'de test').toLowerCase();
      const stars = m.levels ? m.starsOf(gen.tasks, last) : null;
      side.append(h('section', { class: 'sheet pad pop', style: 'display:grid;gap:10px' },
        h('div', { class: 'score' }, h('b', {}, `${sc}/${total}`), h('span', {}, mode === 'test' && !m.levels ? pct + '%' : (m.project ? 'in orde' : 'juist'))),
        h('div', { class: 'bar' }, h('i', { style: `width:${pct}%` })),
        m.levels ? h('p', { class: 'stars-got' }, stars ? h('span', {}, 'Je haalt ', h('b', { class: 'lv-stars' }, '★'.repeat(stars)), stars < m.levelOf(seed) ? ` van de ${'★'.repeat(m.levelOf(seed))}` : '') : 'Nog geen ster: zorg eerst dat alle opdrachten zonder ★ in orde zijn.') : null,
        passed
          ? h('p', { class: 'msg ok' }, mode === 'test' ? h('span', {}, h('b', {}, m.levels ? 'Goedgekeurd! ' : 'Geslaagd! '), modIndex(m.id) < MODULES.length - 1 ? 'De volgende stap is open.' : (m.levels ? 'Je werkblad is ingeleverd. Je hebt alle stappen gehaald!' : 'Je hebt alle stappen gehaald.')) : h('span', {}, h('b', {}, 'Alles in orde! '), m.levels ? 'Laat een klasgenoot je werkblad testen en lever het daarna in.' : 'Je bent klaar voor de test.'))
          : h('p', { class: 'msg info' }, mode === 'test' ? (m.project ? `Nog niet in orde. Verbeter de onderdelen met een ✗ in je ${noun} en laad het opnieuw op.` : 'Nog niet geslaagd. Oefen de opdrachten met een ✗ nog eens en start daarna een nieuwe test.') : 'Verbeter de opdrachten met een ✗ in Excel, bewaar en laad opnieuw op.'),
        passed && mode === 'practice' ? h('div', { class: 'btn-row' }, h('button', { class: 'btn primary', onclick: () => { S.tab = 'test'; render(); window.scrollTo(0, 0); } }, `Naar ${TLt} →`), m.levels ? null : h('button', { class: 'btn', onclick: async () => { await save(m.id, { practice: { seed: newSeed(), started: new Date().toISOString(), lastSeed: null, lastResults: null } }); render(); } }, 'Nieuwe oefeningen')) : null,
        passed && mode === 'test' ? reflectBlock(m) : null,
        !passed && mode === 'test' ? h('div', { class: 'btn-row' }, h('button', { class: 'btn', onclick: () => { S.tab = 'practice'; render(); } }, m.tabLabels?.practice || 'Oefenen'), h('button', { class: 'btn primary', onclick: () => startTest(m) }, m.project ? 'Opnieuw opladen' : 'Nieuwe test')) : null));
    }
    const testFinished = mode === 'test' && last;
    if (!testFinished) side.append(uploadPanel(m, mode, seed, gen, fname));

    root.append(h('div', { class: 'work' }, h('div', { style: 'display:grid;gap:16px;min-width:0' }, lvSec, start, h('section', { style: 'display:grid;gap:10px' }, h('h2', {}, 'Opdrachten'), tasks)), side));
  }

  function renderGateChecks(m, root) {
    const p = prog(m.id).test || {}, peer = prog(m.id).peer || {};
    const lv = levelNow(m);
    const box = h('section', { class: 'sheet pad gate', style: 'display:grid;gap:12px' });
    root.append(box);
    box.append(h('h2', {}, p.passed ? `Je ${m.noun} is ingeleverd ${'★'.repeat(p.stars || 0)}` : `Klaar om je ${m.noun} in te leveren?`));
    if (!lv) { box.append(h('p', {}, 'Kies eerst je niveau.'), levelPicker(m, null, async n => { await setLevel(m, n); render(); })); return; }
    const L = levelInfo(m, lv);
    box.append(h('div', { class: 'lv-row' }, h('span', { class: 'lv-stars', 'aria-hidden': 'true' }, L.stars), h('div', {}, h('b', {}, 'Niveau ' + L.name), h('p', { class: 'small muted' }, `Wil je een ander niveau? Dat kies je bij ${m.tabLabels?.practice || 'Oefenen'}.`))));
    if (m.gate.intro) box.append(h('div', { html: m.gate.intro }));
    const btn = h('button', { class: 'btn primary', disabled: true }, p.passed ? 'Opnieuw inleveren' : `Lever mijn ${m.noun} in`);
    const boxes = m.gate.checks.map((c, k) => h('input', { type: 'checkbox', id: 'gate-c' + k }));
    const fields = m.gate.fields.map(f => f.area
      ? h('textarea', { class: 'in', id: 'gate-' + f.id, rows: 2, maxlength: 300 })
      : h('input', { class: 'in', id: 'gate-' + f.id, maxlength: 60 }));
    m.gate.fields.forEach((f, k) => { fields[k].value = peer[f.id] || ''; });
    const upd = () => { btn.disabled = !(boxes.every(b => b.checked) && fields.every(f => f.value.trim())); };
    box.append(h('ul', { class: 'list-plain gate-checks' }, m.gate.checks.map((c, k) => h('li', {}, h('label', { class: 'chk', for: 'gate-c' + k }, boxes[k], h('span', {}, c))))),
      h('div', { class: 'form', style: 'display:grid;gap:10px' }, m.gate.fields.map((f, k) => h('label', { class: 'f', for: 'gate-' + f.id }, f.label, fields[k]))),
      h('p', { class: 'small muted' }, 'Je leerkracht ziet wie testte en welke tip je kreeg.'),
      h('div', { class: 'btn-row' }, btn, !p.passed ? h('button', { class: 'btn', onclick: () => { S.tab = 'practice'; render(); } }, `Eerst nog ${(m.tabLabels?.practice || 'oefenen').toLowerCase()}`) : null));
    [...boxes, ...fields].forEach(el => el.addEventListener('input', upd));
    boxes.forEach(b => b.addEventListener('change', upd));
    btn.addEventListener('click', async () => {
      btn.disabled = true;
      const data = { at: new Date().toISOString() };
      m.gate.fields.forEach((f, k) => { data[f.id] = fields[k].value.trim(); });
      await save(m.id, { peer: data });
      startTest(m);
    });
  }
  // ---------------- terugblik (na een geslaagde test) ----------------
  const FEELS = [['easy', '😊', 'makkelijk'], ['ok', '🙂', 'goed te doen'], ['hard', '😓', 'moeilijk']];
  const DEFAULT_REFLECT = [{ id: 'nieuw', label: 'Wat was nieuw voor jou?', optional: true }];
  function reflectBlock(m) {
    const r = prog(m.id).reflect;
    const toHome = h('div', { class: 'btn-row' }, h('button', { class: 'btn primary', onclick: () => { S.view = 'home'; render(); } }, 'Naar de leerlijn →'));
    if (r && r.feel) {
      const f = FEELS.find(x => x[0] === r.feel) || FEELS[1];
      const Q = m.reflect || DEFAULT_REFLECT;
      return h('div', { class: 'reflect done' },
        h('p', {}, h('b', {}, 'Je terugblik: '), f[1] + ' ' + f[2]),
        Q.filter(q => (r[q.id] || '').trim()).map(q => h('p', { class: 'small' }, h('b', {}, q.label + ' '), r[q.id])),
        h('div', { class: 'btn-row' }, toHome.firstChild, h('button', { class: 'btn ghost small', onclick: e => { e.currentTarget.closest('.reflect').replaceWith(reflectForm(m, r)); } }, 'Wijzig')));
    }
    return reflectForm(m, null);
  }
  function reflectForm(m, r) {
    const Q = m.reflect || DEFAULT_REFLECT;
    let feel = r ? r.feel : null;
    const btn = h('button', { class: 'btn primary', disabled: true }, 'Bewaar en ga verder');
    const fields = Q.map(q => h('textarea', { class: 'in', id: 'refl-' + q.id, rows: 2, maxlength: 400 }));
    Q.forEach((q, k) => { fields[k].value = (r && r[q.id]) || ''; });
    const upd = () => { btn.disabled = !feel || Q.some((q, k) => !q.optional && !fields[k].value.trim()); };
    const feelBtns = FEELS.map(([id, emo, label]) => h('button', { class: 'feel' + (feel === id ? ' on' : ''), role: 'radio', 'aria-checked': String(feel === id), onclick: e => {
      feel = id; feelBtns.forEach(b => { const on = b === e.currentTarget; b.classList.toggle('on', on); b.setAttribute('aria-checked', String(on)); }); upd();
    } }, h('span', { class: 'emo', 'aria-hidden': 'true' }, emo), label));
    fields.forEach(f => f.addEventListener('input', upd));
    btn.addEventListener('click', async () => {
      btn.disabled = true;
      const data = { feel, at: new Date().toISOString() };
      Q.forEach((q, k) => { data[q.id] = fields[k].value.trim(); });
      await save(m.id, { reflect: data });
      render();
    });
    upd();
    return h('div', { class: 'reflect' },
      h('b', {}, m.levels ? 'Terugblik op je eindproject' : 'Even terugblikken'),
      h('p', { class: 'small' }, m.levels ? 'Je hebt het hele traject afgewerkt! Hoe ging het eindproject?' : 'Hoe ging deze stap voor jou?'),
      h('div', { class: 'feels', role: 'radiogroup', 'aria-label': 'Hoe ging het?' }, feelBtns),
      Q.map((q, k) => h('label', { class: 'f', for: 'refl-' + q.id }, h('span', {}, q.label, q.optional ? h('small', {}, ' (mag leeg blijven)') : null), fields[k])),
      h('p', { class: 'small muted' }, 'Je leerkracht leest je terugblik.'),
      h('div', { class: 'btn-row' }, btn));
  }

  function renderTestGate(m, root) {
    if (m.gate) return renderGateChecks(m, root);
    const p = prog(m.id).test || {};
    const prev = (S.student.attempts || []).filter(a => a.moduleId === m.id && a.mode === 'test');
    root.append(h('section', { class: 'sheet pad gate' },
      h('h2', {}, p.passed ? (m.project ? 'Je eigen rapport is goedgekeurd' : 'Je haalde deze test al') : (m.project ? 'Klaar voor je eigen rapport?' : 'Klaar voor de test?')),
      m.testGate ? h('div', { html: m.testGate }) : h('ul', { class: 'list-plain' },
        h('li', {}, 'Je krijgt ', h('b', {}, 'nieuwe opdrachten'), ', een beetje anders dan bij het oefenen.'),
        h('li', {}, 'Er zijn ', h('b', {}, 'geen hulplijnen'), '.'),
        h('li', {}, 'Je hebt ', h('b', {}, Math.round(PASS * 100) + '%'), ' juist nodig om naar de volgende stap te gaan.'),
        h('li', {}, 'Lukt het niet? Dan oefen je nog wat en start je een nieuwe test.')),
      prev.length ? h('p', { class: 'muted small' }, `Je deed deze test al ${prev.length}×. Beste score: ${p.best || 0}%.`) : null,
      h('div', { class: 'btn-row' }, h('button', { class: 'btn primary', onclick: () => startTest(m) }, m.project ? (p.passed ? 'Opnieuw opladen' : 'Start mijn eigen rapport') : (p.passed ? 'Doe de test opnieuw' : 'Start de test')), !p.passed ? h('button', { class: 'btn', onclick: () => { S.tab = 'practice'; render(); } }, 'Eerst nog oefenen') : null)));
  }
  async function startTest(m) {
    await save(m.id, { test: { seed: m.levels ? m.withLevel(levelNow(m) || 1) : newSeed(), open: true, started: new Date().toISOString(), lastSeed: null, lastResults: null } });
    S.tab = 'test'; render(); window.scrollTo(0, 0);
  }

  function uploadPanel(m, mode, seed, gen, fname) {
    const box = h('section', { class: 'sheet pad', style: 'display:grid;gap:12px' });
    const routeBody = h('div', {});
    const drawRoute = () => {
      clear(routeBody).append(S.uploadRoute === 'A'
        ? h('ol', { class: 'route' }, h('li', { html: 'Bewaar je werk in Excel (<b>Ctrl + S</b>).' }), h('li', { html: 'Klik hieronder op <b>Kies je bestand</b>.' }), h('li', { html: `Kies links de map <b>OneDrive</b> van je school en daarin <span class="addr">${fname}</span>.` }))
        : h('ol', { class: 'route' }, h('li', { html: '<b>Excel op de computer:</b> Bestand → <b>Een kopie opslaan</b> → kies de map <b>Downloads</b>.' }), h('li', { html: '<b>Excel in de browser:</b> Bestand → <b>Een kopie downloaden</b>.' }), h('li', { html: 'Klik hieronder op <b>Kies je bestand</b> en kies het in de map <b>Downloads</b>.' }), h('li', { class: 'muted', html: 'Werk daarna gewoon verder in je bestand in OneDrive.' })));
      tabs.querySelectorAll('button').forEach(b => b.setAttribute('aria-selected', String(b.dataset.r === S.uploadRoute)));
    };
    const tabs = h('div', { class: 'route-tabs', role: 'tablist' },
      h('button', { role: 'tab', 'data-r': 'A', onclick: () => { S.uploadRoute = 'A'; drawRoute(); } }, 'Ik zie de map OneDrive'),
      h('button', { role: 'tab', 'data-r': 'B', onclick: () => { S.uploadRoute = 'B'; drawRoute(); } }, 'Ik zie geen OneDrive'));
    const input = h('input', { type: 'file', id: 'file-' + mode, accept: '.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const status = h('div', {});
    const drop = h('label', { class: 'drop', for: 'file-' + mode },
      input, h('span', { style: 'font-size:1.6rem' }, '⤒'), h('b', {}, 'Kies je bestand'), h('span', { class: 'small muted' }, 'of sleep het hierheen'));
    ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, () => drop.classList.remove('over')));
    drop.addEventListener('drop', e => { e.preventDefault(); if (e.dataTransfer.files[0]) handle(e.dataTransfer.files[0]); });
    input.addEventListener('change', () => { if (input.files[0]) handle(input.files[0]); input.value = ''; });

    async function handle(file, force) {
      clear(status);
      if (!/\.xlsx$/i.test(file.name)) return status.append(h('p', { class: 'msg bad' }, `"${file.name}" is geen Excel-bestand (.xlsx). Kies het bestand dat je in Excel bewaarde.`));
      if (file.size > 3 * 1024 * 1024) return status.append(h('p', { class: 'msg bad' }, 'Dit bestand is te groot (meer dan 3 MB). Staan er grote afbeeldingen in?'));
      status.append(h('p', { class: 'muted' }, 'Ik kijk je bestand na …'));
      let wb;
      try { wb = await X.readWorkbook(await file.arrayBuffer(), { JSZip: window.JSZip, DOMParser: window.DOMParser }); }
      catch (e) { clear(status); return status.append(h('p', { class: 'msg bad' }, 'Dit bestand kan ik niet openen. Bewaar het opnieuw in Excel en probeer nog eens.')); }
      const meta = wb.meta;
      if (!force) {
        if (!meta.application && !meta.lastModifiedBy && !m.newFile) {
          clear(status);
          return status.append(h('div', { class: 'confirm-box' }, h('p', {}, h('b', {}, 'Dit lijkt nog het startbestand zonder jouw werk.'), ' Open het in Excel, doe de opdrachten en bewaar het. Laad daarna dat bestand op.'),
            h('div', { class: 'btn-row' }, h('button', { class: 'btn small', onclick: () => handle(file, true) }, 'Toch nakijken'))));
        }
        const started = new Date((mode === 'test' ? prog(m.id).test?.started : prog(m.id).practice?.started) || 0);
        if (meta.modified && Date.now() - meta.modified.getTime() > 36 * 3600 * 1000 && meta.modified < started) {
          clear(status);
          return status.append(h('div', { class: 'confirm-box' }, h('p', {}, h('b', {}, `Dit bestand werd laatst bewaard op ${fmtDate(meta.modified)}.`), ' Is dit wel je nieuwste versie?'),
            h('div', { class: 'btn-row' }, h('button', { class: 'btn small primary', onclick: () => handle(file, true) }, 'Ja, kijk het na'), h('button', { class: 'btn small', onclick: () => { clear(status); input.click(); } }, 'Nee, ander bestand kiezen'))));
        }
      }
      const results = runChecks(m, gen, wb, { student: S.student, fileName: file.name });
      const score = results.filter(r => r.ok).length, total = results.length, pct = Math.round(score / total * 100);
      const hints = Object.values((prog(m.id).practice || {}).hints || {}).reduce((a, b) => a + b, 0);
      const st = mode === 'practice' ? (prog(m.id).practice || {}) : (prog(m.id).test || {});
      const okNow = passedOf(m, mode, gen.tasks, results);
      const patch = mode === 'practice'
        ? { practice: { attempts: (st.attempts || 0) + 1, passed: !!st.passed || okNow, lastSeed: seed, lastResults: results } }
        : { test: { attempts: (st.attempts || 0) + 1, passed: !!st.passed || okNow, best: Math.max(st.best || 0, pct), lastSeed: seed, lastResults: results, lastPct: pct } };
      if (m.levels) { const stars = m.starsOf(gen.tasks, results); patch[mode].lastStars = stars; if (mode === 'test') patch.test.stars = Math.max(st.stars || 0, stars); }
      let b64 = null;
      try { b64 = await Api.fileToB64(file); } catch (e) { /* zonder bestand */ }
      try {
        const d = await Api.attempt({ moduleId: m.id, mode, seed, score, total, results, hints: mode === 'practice' ? hints : 0, fileName: file.name, fileB64: b64, fileModified: meta.modified ? meta.modified.toISOString() : null, patch });
        S.student = d.student;
      } catch (e) {
        if (e.status === 401) return logout(e.message);
        S.student.progress[m.id] = Object.assign(S.student.progress[m.id] || {}, patch);
        toast('Nagekeken, maar niet doorgestuurd: ' + e.message, true);
      }
      S.results[m.id + mode + seed] = results;
      render();
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }
    box.append(h('h2', {}, 'Laad je bestand op'), h('p', { class: 'small muted' }, 'De app kijkt je werk meteen na en stuurt het naar je leerkracht.'), tabs, routeBody, drop, status);
    drawRoute();
    return box;
  }

  // ---------------- start ----------------
  function render() {
    renderWho();
    if (!S.student) return renderLogin();
    if (S.view === 'module') renderModule(); else renderHome();
  }
  (async function boot() {
    const { demo } = await Api.init();
    document.getElementById('demoBand').hidden = !demo;
    if (Api.hasSession()) {
      try { const d = await Api.me(); S.student = d.student; } catch (e) { Api.logout(); }
    }
    render();
  })();
})();
