/* Cel voor Cel – leerkrachtoverzicht */
(function () {
  'use strict';
  const { MODULES, PASS } = window.CvcModules;
  const Api = window.CvcApi;
  const $app = document.getElementById('app');
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
  const fmt = d => d ? new Date(d).toLocaleString('nl-BE', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }) : '–';
  const strip = s => String(s).replace(/<[^>]+>/g, '');

  const T = { students: [], klas: '', zoek: '', detail: null };

  function renderWho() {
    const w = clear(document.getElementById('who'));
    if (Api.tHasSession()) w.append(h('button', { class: 'btn ghost small', onclick: () => { Api.tLogout(); renderLogin(); renderWho(); } }, 'Afmelden'));
  }
  document.getElementById('logo').addEventListener('click', e => { e.preventDefault(); if (Api.tHasSession()) { T.detail = null; renderOverview(); } });

  function renderLogin(msg) {
    clear($app);
    const pw = h('input', { class: 'in', type: 'password', id: 't-pw', autocomplete: 'current-password' });
    const err = h('p', { class: 'msg bad', hidden: !msg }, msg || '');
    const form = h('form', { class: 'sheet pad form', style: 'max-width:420px' }, h('h1', {}, 'Leerkracht'), h('label', { class: 'f', for: 't-pw' }, 'Wachtwoord', pw), err, h('div', { class: 'btn-row' }, h('button', { class: 'btn primary', type: 'submit' }, 'Aanmelden')));
    form.addEventListener('submit', async e => {
      e.preventDefault(); err.hidden = true;
      try { await Api.tLogin(pw.value); renderWho(); load(); } catch (x) { err.textContent = x.message; err.hidden = false; }
    });
    $app.append(form); pw.focus();
  }

  async function load() {
    clear($app).append(h('p', { class: 'muted' }, 'Laden …'));
    try { const d = await Api.tOverview(); T.students = d.students.sort((a, b) => (a.klas + a.naam + a.voornaam).localeCompare(b.klas + b.naam + b.voornaam, 'nl')); }
    catch (e) { if (e.status === 401) { Api.tLogout(); return renderLogin(e.message); } clear($app).append(h('p', { class: 'msg bad' }, e.message)); return; }
    if (T.detail) openStudent(T.detail); else renderOverview();
  }

  function modStatus(s, m) {
    const p = (s.progress || {})[m.id] || {};
    const ld = m.lessons.filter(l => p.lessons?.[l.id]).length;
    const prA = (s.attempts || []).filter(a => a.moduleId === m.id && a.mode === 'practice');
    const hints = Object.values(p.practice?.hints || {}).reduce((a, b) => a + b, 0);
    const seed = p.practice?.seed || p.test?.seed;
    const level = m.levels && seed ? m.levelOf(seed) : null;
    return { p, ld, prA, hints, testBest: p.test?.best, testPassed: !!p.test?.passed, testN: p.test?.attempts || 0, practicePassed: !!p.practice?.passed, skipped: !ld && !prA.length && (p.test?.attempts || 0) > 0,
      level, stars: p.test?.stars || 0, peer: p.peer || null };
  }
  const FEELS = { easy: ['😊', 'makkelijk', 'ok'], ok: ['🙂', 'goed te doen', 'mark'], hard: ['😓', 'moeilijk', 'bad'] };
  const reflectQ = m => m.reflect || [{ id: 'nieuw', label: 'Wat was nieuw?' }];
  const reflectTxt = (m, r) => reflectQ(m).filter(q => (r[q.id] || '').trim()).map(q => `${q.label} ${r[q.id]}`).join(' · ');
  const TL = (m, mode) => (m.tabLabels && m.tabLabels[mode]) || (mode === 'test' ? 'Test' : 'Oefenen');
  const starTxt = n => n ? '★'.repeat(n) : 'geen ster';

  function renderOverview() {
    clear($app);
    const klassen = [...new Set(T.students.map(s => s.klas))].sort();
    const sel = h('select', { class: 'in', id: 't-klas', style: 'width:auto', onchange: e => { T.klas = e.target.value; drawTable(); } }, h('option', { value: '' }, 'Alle klassen'), klassen.map(k => h('option', { value: k, selected: T.klas === k }, k)));
    const zoek = h('input', { class: 'in', id: 't-zoek', placeholder: 'Zoek een leerling', style: 'width:220px;max-width:100%', value: T.zoek, oninput: e => { T.zoek = e.target.value; drawTable(); } });
    const wrap = h('div', { class: 't-table-wrap' });
    const drawTable = () => {
      const list = T.students.filter(s => (!T.klas || s.klas === T.klas) && (!T.zoek || (s.voornaam + ' ' + s.naam).toLowerCase().includes(T.zoek.toLowerCase())));
      clear(wrap);
      if (!list.length) return wrap.append(h('p', { class: 'pad muted' }, T.students.length ? 'Geen leerlingen gevonden.' : 'Nog geen leerlingen. Een leerling verschijnt hier zodra hij of zij zich voor het eerst aanmeldt.'));
      wrap.append(h('table', { class: 't' },
        h('thead', {}, h('tr', {}, h('th', {}, 'Klas'), h('th', {}, 'Leerling'), MODULES.map(m => h('th', {}, `${m.nr}. ${m.title}`)), h('th', {}, 'Laatst actief'))),
        h('tbody', {}, list.map(s => h('tr', { onclick: () => openStudent(s.id), tabindex: '0', onkeydown: e => { if (e.key === 'Enter') openStudent(s.id); } },
          h('td', {}, s.klas), h('td', {}, h('b', {}, s.voornaam + ' ' + s.naam)),
          MODULES.map(m => {
            const st = modStatus(s, m);
            return h('td', {}, h('div', { class: 'mcell' },
              h('div', { class: 'dots', title: `Leren ${st.ld}/${m.lessons.length}` }, m.lessons.map(l => h('span', { class: 'dot' + (st.p.lessons?.[l.id] ? ' on' : '') }))),
              st.p.reflect?.feel ? h('span', { class: 'small feel-tag ' + FEELS[st.p.reflect.feel][2], title: reflectTxt(m, st.p.reflect) || 'geen extra uitleg' }, FEELS[st.p.reflect.feel][0] + ' ' + FEELS[st.p.reflect.feel][1]) : null,
              st.level ? h('span', { class: 'small' }, 'Niveau ', h('b', {}, '★'.repeat(st.level))) : null,
              h('span', { class: 'small' }, m.levels ? 'Tussentijds: ' : 'Oefenen: ', st.practicePassed ? h('b', { style: 'color:var(--ok)' }, '✓') : (st.prA.length ? st.prA.length + '×' : '–'), st.hints ? h('span', { class: 'muted' }, ` · ${st.hints} hulp`) : null),
              h('span', { class: 'small' }, m.levels ? 'Ingeleverd: ' : 'Test: ', st.testN ? h('b', { style: `color:var(--${st.testPassed ? 'ok' : 'bad'})` }, m.levels ? (st.testPassed ? '✓ ' + starTxt(st.stars) : 'nog niet in orde') : `${st.testBest || 0}%`) : '–', st.testN > 1 ? h('span', { class: 'muted' }, ` (${st.testN}×)`) : null, st.skipped ? h('span', { class: 'chip', style: 'margin-left:4px' }, 'kon het al') : null)));
          }),
          h('td', { class: 'num small' }, fmt(s.lastSeen)))))));
    };
    $app.append(
      h('div', { class: 'home-head' }, h('div', {}, h('h1', {}, 'Overzicht'), h('p', { class: 'muted' }, `${T.students.length} leerling${T.students.length === 1 ? '' : 'en'}. Klik op een leerling voor alle pogingen en bestanden.`))),
      h('div', { class: 't-tools' }, sel, zoek, h('span', { class: 'spacer' }), h('button', { class: 'btn small', onclick: load }, '↻ Vernieuwen'), h('button', { class: 'btn small', onclick: exportCsv }, '⤓ Exporteer (CSV)')),
      wrap);
    drawTable();
  }

  function exportCsv() {
    const rows = [['Klas', 'Voornaam', 'Naam'].concat(...MODULES.map(m => [`M${m.nr} lessen`, `M${m.nr} oefenen klaar`, `M${m.nr} oefenpogingen`, `M${m.nr} hulplijnen`, `M${m.nr} test beste %`, `M${m.nr} test geslaagd`, `M${m.nr} testpogingen`, `M${m.nr} gevoel`, `M${m.nr} terugblik`].concat(m.levels ? [`M${m.nr} niveau`, `M${m.nr} sterren`, `M${m.nr} getest door`, `M${m.nr} tip`] : [])), ['Laatst actief'])];
    T.students.filter(s => !T.klas || s.klas === T.klas).forEach(s => rows.push([s.klas, s.voornaam, s.naam].concat(...MODULES.map(m => {
      const st = modStatus(s, m);
      return [`${st.ld}/${m.lessons.length}`, st.practicePassed ? 'ja' : 'nee', st.prA.length, st.hints, st.testBest ?? '', st.testPassed ? 'ja' : 'nee', st.testN, st.p.reflect?.feel ? FEELS[st.p.reflect.feel][1] : '', st.p.reflect ? reflectTxt(m, st.p.reflect) : ''].concat(m.levels ? [st.level || '', st.stars || 0, st.peer?.tester || '', st.peer?.tip || ''] : []);
    }), [s.lastSeen ? new Date(s.lastSeen).toLocaleString('nl-BE') : ''])));
    const csv = '﻿' + rows.map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(';')).join('\r\n');
    Api.download(new Blob([csv], { type: 'text/csv;charset=utf-8' }), `celvoorcel-overzicht${T.klas ? '-' + T.klas : ''}.csv`);
  }

  async function openStudent(id) {
    T.detail = id;
    clear($app).append(h('p', { class: 'muted' }, 'Laden …'));
    let s;
    try { s = (await Api.tStudent(id)).student; } catch (e) { clear($app).append(h('p', { class: 'msg bad' }, e.message)); return; }
    clear($app);
    const actions = h('div', { style: 'display:grid;gap:10px' });
    const drawActions = () => {
      clear(actions).append(h('div', { class: 'btn-row' },
        h('button', { class: 'btn small', onclick: () => {
          const inp = h('input', { class: 'in', id: 't-newcode', inputmode: 'numeric', placeholder: 'nieuwe code', style: 'width:150px', maxlength: 8 });
          clear(actions).append(h('div', { class: 'confirm-box' }, h('label', { class: 'f', for: 't-newcode' }, `Nieuwe code voor ${s.voornaam} (4 tot 8 cijfers)`, inp),
            h('div', { class: 'btn-row' }, h('button', { class: 'btn small primary', onclick: async () => { try { await Api.tReset(s.id, inp.value.trim()); drawActions(); actions.prepend(h('p', { class: 'msg ok' }, `De code is nu ${inp.value.trim()}. Geef die door aan ${s.voornaam}.`)); } catch (e) { actions.append(h('p', { class: 'msg bad' }, e.message)); } } }, 'Bewaar code'), h('button', { class: 'btn small', onclick: drawActions }, 'Annuleer'))));
          inp.focus();
        } }, 'Code opnieuw instellen'),
        h('button', { class: 'btn small', onclick: () => {
          clear(actions).append(h('div', { class: 'inline-confirm' }, h('span', {}, `${s.voornaam} ${s.naam} en alle bestanden definitief verwijderen?`),
            h('button', { class: 'btn small', style: 'border-color:var(--bad);color:var(--bad)', onclick: async () => { await Api.tDelete(s.id); T.detail = null; load(); } }, 'Ja, verwijder'), h('button', { class: 'btn small', onclick: drawActions }, 'Annuleer')));
        } }, 'Leerling verwijderen')));
    };
    drawActions();

    const mods = MODULES.map(m => {
      const st = modStatus(s, m);
      const atts = (s.attempts || []).filter(a => a.moduleId === m.id).sort((a, b) => b.at.localeCompare(a.at));
      return h('section', { class: 'sheet pad', style: 'display:grid;gap:12px' },
        h('div', { class: 'att-head' }, h('h2', {}, `${m.nr}. ${m.title}`), h('span', { class: 'spacer' }),
          (st.ld || atts.length) ? h('button', { class: 'btn ghost small', onclick: async e => {
            const b = e.currentTarget;
            b.replaceWith(h('span', { class: 'inline-confirm' }, 'Deze stap voor deze leerling helemaal opnieuw laten starten?',
              h('button', { class: 'btn small', onclick: async () => { await Api.tModuleReset(s.id, m.id); openStudent(s.id); } }, 'Ja'), h('button', { class: 'btn small', onclick: () => openStudent(s.id) }, 'Nee')));
          } }, 'Stap opnieuw openzetten') : null),
        h('div', { class: 'chips' },
          h('span', { class: 'chip' + (st.ld === m.lessons.length ? ' ok' : '') }, `Leren ${st.ld}/${m.lessons.length}`),
          st.level ? h('span', { class: 'chip go' }, `Niveau ${'★'.repeat(st.level)}`) : null,
          h('span', { class: 'chip' + (st.practicePassed ? ' ok' : '') }, `${TL(m, 'practice')} ${st.practicePassed ? '✓' : (st.prA.length ? st.prA.length + '× opgeladen' : '–')}`),
          h('span', { class: 'chip' }, `${st.hints} hulplijn${st.hints === 1 ? '' : 'en'}`),
          m.levels
            ? h('span', { class: 'chip' + (st.testPassed ? ' ok' : st.testN ? ' bad' : '') }, `${TL(m, 'test')} ${st.testN ? (st.testPassed ? '✓ ' + starTxt(st.stars) : 'nog niet in orde') : '–'}`)
            : h('span', { class: 'chip' + (st.testPassed ? ' ok' : st.testN ? ' bad' : '') }, `Test ${st.testN ? (st.testBest || 0) + '%' : '–'}${st.testPassed ? ' ✓' : ''}`)),
        st.p.reflect?.feel ? h('div', { class: 'peer refl' }, h('b', {}, 'Terugblik: '), FEELS[st.p.reflect.feel][0] + ' ' + FEELS[st.p.reflect.feel][1],
          reflectQ(m).filter(q => (st.p.reflect[q.id] || '').trim()).map(q => [h('br'), h('b', {}, q.label + ' '), st.p.reflect[q.id]])) : null,
        st.peer ? h('p', { class: 'peer' }, h('b', {}, 'Getest door: '), st.peer.tester || '–', st.peer.tip ? [h('br'), h('b', {}, 'Tip van de tester: '), st.peer.tip] : null) : null,
        atts.length ? h('div', { style: 'display:grid;gap:10px' }, atts.map(a => attemptCard(s, m, a))) : h('p', { class: 'muted small' }, 'Nog niets opgeladen.'));
    });

    $app.append(
      h('div', { class: 'crumb' }, h('button', { class: 'btn ghost small', onclick: () => { T.detail = null; renderOverview(); } }, '← Overzicht')),
      h('div', { class: 'drawer' },
        h('div', { class: 'mod-head' }, h('span', { class: 'eyebrow' }, s.klas), h('h1', {}, `${s.voornaam} ${s.naam}`), h('p', { class: 'muted small' }, `Account sinds ${fmt(s.created)} · laatst actief ${fmt(s.lastSeen)}`)),
        actions, mods));
    window.scrollTo(0, 0);
  }

  function attemptCard(s, m, a) {
    let tasks = [];
    try { tasks = m.gen(a.seed, a.mode, s).tasks; } catch (e) { /* oude seed */ }
    const pct = a.total ? Math.round(a.score / a.total * 100) : 0;
    const passed = a.mode === 'test' ? (m.passCheck && tasks.length ? m.passCheck(tasks, a.results || []) : pct >= PASS * 100) : a.score === a.total;
    const stars = m.levels && tasks.length ? m.starsOf(tasks, a.results || []) : null;
    const details = h('ul', { hidden: true }, (a.results || []).map(r => {
      const t = tasks.find(x => x.id === r.id);
      return h('li', {}, h('span', { class: r.ok ? 'y' : 'n' }, r.ok ? '✓' : '✗'), h('span', {}, t ? strip(t.text) : r.id, !r.ok && r.msg ? h('span', { class: 'muted' }, ' — ' + r.msg) : null));
    }));
    const dl = a.fileKey ? h('button', { class: 'btn small', onclick: async () => {
      try { const d = await Api.tFile(a.fileKey); Api.download(Api.b64ToBlob(d.b64), `${s.klas}_${s.voornaam}_${s.naam}_${m.id}_${a.mode === 'test' ? (m.levels ? 'ingeleverd' : 'test') : (m.levels ? 'tussentijds' : 'oefenen')}.xlsx`.replace(/\s+/g, '-')); }
      catch (e) { dl.after(h('span', { class: 'small', style: 'color:var(--bad)' }, ' ' + e.message)); }
    } }, '⤓ Bestand') : null;
    return h('div', { class: 'att' },
      h('div', { class: 'att-head' },
        h('span', { class: 'chip' + (a.mode === 'test' ? ' go' : '') }, TL(m, a.mode)),
        m.levels ? h('span', { class: 'chip' }, 'Niveau ' + '★'.repeat(m.levelOf(a.seed))) : null,
        h('b', { style: `color:var(--${passed ? 'ok' : 'bad'})` }, `${a.score}/${a.total}${a.mode === 'test' && !m.levels ? ` (${pct}%)` : ''}`),
        stars != null ? h('span', { class: 'small' }, 'behaald: ', h('b', {}, starTxt(stars))) : null,
        h('span', { class: 'small muted' }, fmt(a.at)),
        a.mode === 'practice' && a.hints ? h('span', { class: 'small muted' }, `${a.hints} hulplijnen gebruikt`) : null,
        a.fileName ? h('span', { class: 'small addr' }, a.fileName) : null,
        h('span', { class: 'spacer' }),
        (a.results || []).length ? h('button', { class: 'btn ghost small', onclick: e => { details.hidden = !details.hidden; e.currentTarget.textContent = details.hidden ? 'Toon opdrachten' : 'Verberg opdrachten'; } }, 'Toon opdrachten') : null,
        dl),
      details);
  }

  (async function boot() {
    const { demo } = await Api.init();
    document.getElementById('demoBand').hidden = !demo;
    renderWho();
    if (Api.tHasSession()) load(); else renderLogin();
  })();
})();
