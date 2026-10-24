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
  function moduleChips(m) {
    const p = prog(m.id), chips = [];
    const ld = lessonsDone(m);
    chips.push(h('span', { class: 'chip' + (ld === m.lessons.length ? ' ok' : '') }, `Leren ${ld}/${m.lessons.length}`));
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
      return h('div', { class: 'mrow' + (i === cur ? ' current' : '') + (open ? '' : ' locked') },
        h('div', { class: 'rn' }, m.nr),
        h('div', { class: 'mb' }, h('h3', {}, m.title), h('p', { class: 'small' }, open ? m.short : 'Dit deel gaat open als je de test van de vorige stap haalt.'), open ? h('div', { class: 'chips' }, moduleChips(m)) : null),
        h('div', { class: 'ma' }, open ? h('button', { class: 'btn' + (i === cur ? ' primary' : ''), onclick: () => openModule(m.id) }, done ? 'Bekijk' : (i === cur && (lessonsDone(m) || prog(m.id).practice) ? 'Verder' : 'Start')) : h('span', { class: 'muted small' }, 'Op slot')));
    });
    const soon = SOON.map(s => h('div', { class: 'mrow soon' + (s.project ? ' project' : '') }, h('div', { class: 'rn' }, s.nr), h('div', { class: 'mb' }, h('h3', {}, s.title), h('p', { class: 'small' }, s.short))));
    $app.append(
      h('div', { class: 'home-head' },
        h('div', {}, h('h1', {}, `Dag ${S.student.voornaam}!`), h('p', { class: 'muted' }, allDone ? 'Je hebt alle stappen die nu klaarstaan gehaald. Knap!' : 'Dit is jouw leerlijn. Werk de stappen van boven naar onder af.'))),
      h('div', { class: 'path' }, rows),
      h('p', { class: 'sep-label' }, 'Komt later'),
      h('div', { class: 'path' }, soon));
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
    const tabs = [['learn', 'Leren', lessonsDone(m) === m.lessons.length], ['practice', 'Oefenen in Excel', !!p.practice?.passed], ['test', 'Test', !!p.test?.passed]];
    const tabBar = h('div', { class: 'tabs', role: 'tablist' },
      tabs.map(([id, label, done], i) => h('button', { class: 'tab', role: 'tab', 'aria-selected': String(S.tab === id), onclick: () => { S.tab = id; render(); } }, h('span', { class: 'n' }, i + 1), label, done ? h('span', { class: 'done', 'aria-label': 'klaar' }, '✓') : null)));
    const skip = S.tab !== 'test' && !p.test?.passed ? h('button', { class: 'btn ghost small skip', onclick: () => { S.tab = 'test'; render(); } }, 'Ik kan dit al → naar de test') : null;
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
    steps(w, onDone) {
      const on = new Set();
      const btn = h('button', { class: 'btn primary', disabled: true, onclick: () => { btn.disabled = true; onDone(); } }, w.confirm);
      const list = h('ol', { class: 'steps' }, w.steps.map((s, k) => h('li', { tabindex: '0', role: 'checkbox', 'aria-checked': 'false', onclick: e => tog(e.currentTarget, k), onkeydown: e => { if (e.key === ' ' || e.key === 'Enter') { e.preventDefault(); tog(e.currentTarget, k); } } }, h('span', { class: 'bx' }), h('span', { html: s }))));
      const tog = (li, k) => { on.has(k) ? on.delete(k) : on.add(k); li.classList.toggle('on', on.has(k)); li.querySelector('.bx').textContent = on.has(k) ? '✓' : ''; li.setAttribute('aria-checked', String(on.has(k))); btn.disabled = on.size !== w.steps.length; };
      return h('div', { class: 'q' }, h('p', { class: 'small muted' }, 'Vink elke stap aan als je hem begrijpt.'), list, h('div', { class: 'btn-row' }, btn));
    },
    read(w, onDone) {
      const btn = h('button', { class: 'btn primary', onclick: () => { btn.disabled = true; onDone(); } }, w.confirm || 'Gelezen');
      return h('div', { class: 'btn-row' }, btn);
    }
  };

  // ---------------- oefenen en test ----------------
  async function ensureSeed(m, mode) {
    const p = prog(m.id);
    if (mode === 'practice' && !p.practice?.seed) await save(m.id, { practice: { seed: newSeed(), started: new Date().toISOString() } });
    return mode === 'practice' ? prog(m.id).practice.seed : p.test?.seed;
  }
  function fileWord(m, mode) { return mode === 'test' ? 'test' + m.nr : 'module' + m.nr; }

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
    const gen = m.gen(seed, mode, S.student);
    const st = mode === 'practice' ? (prog(m.id).practice || {}) : (prog(m.id).test || {});
    const last = S.results[m.id + mode + seed] || (st.lastSeed === seed ? st.lastResults : null);
    const resMap = {}; (last || []).forEach(r => { resMap[r.id] = r; });
    const vn = S.student.voornaam.replace(/\s+/g, '');
    const fname = `${vn}_${fileWord(m, mode)}`;

    // --- zo begin je ---
    const start = h('section', { class: 'sheet pad', style: 'display:grid;gap:12px' },
      h('h2', {}, mode === 'test' ? 'Je test' : 'Zo begin je'),
      mode === 'test' ? h('p', { class: 'msg info' }, 'Dit is de test: er zijn geen hulplijnen. Lees elke opdracht goed. Je hebt ', h('b', {}, Math.round(PASS * 100) + '%'), ' nodig.') : null,
      m.newFile
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
        h('div', { class: 'tb' }, h('p', { html: t.text }), r ? h('p', { class: 'res' }, r.ok ? 'Juist!' : r.msg) : null, hintBox));
    }));

    // --- opladen ---
    const side = h('aside', { class: 'side' });
    const total = gen.tasks.length;
    if (last) {
      const sc = last.filter(r => r.ok).length, pct = Math.round(sc / total * 100);
      const passed = mode === 'test' ? pct >= PASS * 100 : sc === total;
      side.append(h('section', { class: 'sheet pad pop', style: 'display:grid;gap:10px' },
        h('div', { class: 'score' }, h('b', {}, `${sc}/${total}`), h('span', {}, mode === 'test' ? pct + '%' : 'juist')),
        h('div', { class: 'bar' }, h('i', { style: `width:${pct}%` })),
        passed
          ? h('p', { class: 'msg ok' }, mode === 'test' ? h('span', {}, h('b', {}, 'Geslaagd! '), modIndex(m.id) < MODULES.length - 1 ? 'De volgende stap is open.' : 'Je hebt alle stappen gehaald.') : h('span', {}, h('b', {}, 'Alles juist! '), 'Je bent klaar voor de test.'))
          : h('p', { class: 'msg info' }, mode === 'test' ? 'Nog niet geslaagd. Oefen de opdrachten met een ✗ nog eens en start daarna een nieuwe test.' : 'Verbeter de opdrachten met een ✗ in Excel, bewaar en laad opnieuw op.'),
        passed && mode === 'practice' ? h('div', { class: 'btn-row' }, h('button', { class: 'btn primary', onclick: () => { S.tab = 'test'; render(); window.scrollTo(0, 0); } }, 'Naar de test →'), h('button', { class: 'btn', onclick: async () => { await save(m.id, { practice: { seed: newSeed(), started: new Date().toISOString(), lastSeed: null, lastResults: null } }); render(); } }, 'Nieuwe oefeningen')) : null,
        passed && mode === 'test' ? h('div', { class: 'btn-row' }, h('button', { class: 'btn primary', onclick: () => { S.view = 'home'; render(); } }, 'Naar de leerlijn →')) : null,
        !passed && mode === 'test' ? h('div', { class: 'btn-row' }, h('button', { class: 'btn', onclick: () => { S.tab = 'practice'; render(); } }, 'Oefenen'), h('button', { class: 'btn primary', onclick: () => startTest(m) }, 'Nieuwe test')) : null));
    }
    const testFinished = mode === 'test' && last;
    if (!testFinished) side.append(uploadPanel(m, mode, seed, gen, fname));

    root.append(h('div', { class: 'work' }, h('div', { style: 'display:grid;gap:16px;min-width:0' }, start, h('section', { style: 'display:grid;gap:10px' }, h('h2', {}, 'Opdrachten'), tasks)), side));
  }

  function renderTestGate(m, root) {
    const p = prog(m.id).test || {};
    const prev = (S.student.attempts || []).filter(a => a.moduleId === m.id && a.mode === 'test');
    root.append(h('section', { class: 'sheet pad gate' },
      h('h2', {}, p.passed ? 'Je haalde deze test al' : 'Klaar voor de test?'),
      h('ul', { class: 'list-plain' },
        h('li', {}, 'Je krijgt ', h('b', {}, 'nieuwe opdrachten'), ', een beetje anders dan bij het oefenen.'),
        h('li', {}, 'Er zijn ', h('b', {}, 'geen hulplijnen'), '.'),
        h('li', {}, 'Je hebt ', h('b', {}, Math.round(PASS * 100) + '%'), ' juist nodig om naar de volgende stap te gaan.'),
        h('li', {}, 'Lukt het niet? Dan oefen je nog wat en start je een nieuwe test.')),
      prev.length ? h('p', { class: 'muted small' }, `Je deed deze test al ${prev.length}×. Beste score: ${p.best || 0}%.`) : null,
      h('div', { class: 'btn-row' }, h('button', { class: 'btn primary', onclick: () => startTest(m) }, p.passed ? 'Doe de test opnieuw' : 'Start de test'), !p.passed ? h('button', { class: 'btn', onclick: () => { S.tab = 'practice'; render(); } }, 'Eerst nog oefenen') : null)));
  }
  async function startTest(m) {
    await save(m.id, { test: { seed: newSeed(), open: true, started: new Date().toISOString(), lastSeed: null, lastResults: null } });
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
      const patch = mode === 'practice'
        ? { practice: { attempts: (st.attempts || 0) + 1, passed: !!st.passed || score === total, lastSeed: seed, lastResults: results } }
        : { test: { attempts: (st.attempts || 0) + 1, passed: !!st.passed || pct >= PASS * 100, best: Math.max(st.best || 0, pct), lastSeed: seed, lastResults: results, lastPct: pct } };
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
