/* Cel voor Cel – praat met de opslag op Netlify, of werkt in demomodus (alles in deze browser). */
(function (root) {
  'use strict';
  const TKEY = 'cvc-token';
  const ss = {
    get(k) { try { return sessionStorage.getItem(k); } catch (e) { return null; } },
    set(k, v) { try { sessionStorage.setItem(k, v); } catch (e) { /* geen opslag */ } },
    del(k) { try { sessionStorage.removeItem(k); } catch (e) { /* geen opslag */ } }
  };

  // ---------------- echte opslag ----------------
  function Remote() {}
  Remote.prototype.call = async function (path, opts) {
    opts = opts || {};
    const headers = { 'content-type': 'application/json' };
    const t = opts.token || ss.get(opts.teacher ? TKEY + '-t' : TKEY);
    if (t) headers.authorization = 'Bearer ' + t;
    const res = await fetch('/api/' + path, { method: opts.body ? 'POST' : 'GET', headers, body: opts.body ? JSON.stringify(opts.body) : undefined });
    let data = {};
    try { data = await res.json(); } catch (e) { throw new Error('De server antwoordt niet goed. Probeer het opnieuw.'); }
    if (!res.ok) { const err = new Error(data.error || 'Er ging iets mis.'); err.status = res.status; throw err; }
    return data;
  };

  // ---------------- demomodus (localStorage) ----------------
  const DKEY = 'cvc-demo-v1';
  const demoLoad = () => { try { return JSON.parse(localStorage.getItem(DKEY)) || { students: {}, files: {} }; } catch (e) { return { students: {}, files: {} }; } };
  const demoSave = db => { try { localStorage.setItem(DKEY, JSON.stringify(db)); } catch (e) { /* vol of geblokkeerd */ } };
  const slug = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
  const pub = s => { if (!s) return s; const { code, ...r } = s; return JSON.parse(JSON.stringify(r)); };
  function merge(t, p) { for (const [k, v] of Object.entries(p || {})) { if (v && typeof v === 'object' && !Array.isArray(v)) t[k] = merge(t[k] && typeof t[k] === 'object' ? t[k] : {}, v); else t[k] = v; } return t; }

  function Demo() {}
  Demo.prototype.call = async function (path, opts) {
    opts = opts || {};
    const b = opts.body || {};
    const db = demoLoad();
    const err = (m, st) => { const e = new Error(m); e.status = st || 400; throw e; };
    const [route, q] = path.split('?');
    const qs = new URLSearchParams(q || '');
    const me = () => { const id = ss.get(TKEY); return id && db.students[id.replace('demo:', '')]; };

    if (route === 'login') {
      const clean = s => String(s || '').replace(/\s+/g, ' ').trim();
      const voornaam = clean(b.voornaam), naam = clean(b.naam), klas = clean(b.klas).toUpperCase(), code = String(b.code || '').trim();
      if (!voornaam || !naam || !klas) err('Vul je voornaam, naam en klas in.');
      if (!/^\d{4,8}$/.test(code)) err('Je code bestaat uit 4 tot 8 cijfers.');
      const id = slug(klas) + '--' + slug(voornaam) + '--' + slug(naam);
      let s = db.students[id];
      if (!s) {
        if (!b.nieuw) return { nieuw: true };
        s = db.students[id] = { id, voornaam, naam, klas, code, created: new Date().toISOString(), lastSeen: new Date().toISOString(), progress: {}, attempts: [] };
      } else if (s.code !== code) err('Deze code klopt niet. Vraag je leerkracht om je code opnieuw in te stellen.', 401);
      s.lastSeen = new Date().toISOString();
      demoSave(db);
      return { token: 'demo:' + id, student: pub(s) };
    }
    if (route.startsWith('teacher/')) {
      if (route === 'teacher/login') { if (b.wachtwoord !== 'demo') err('In de demo is het wachtwoord: demo', 401); return { token: 'demo-teacher' }; }
      if (route === 'teacher/overview') return { students: Object.values(db.students).map(pub) };
      if (route === 'teacher/student') { const s = db.students[qs.get('id')]; if (!s) err('Leerling niet gevonden.', 404); return { student: pub(s) }; }
      if (route === 'teacher/file') { const f = db.files[qs.get('key')]; if (!f) err('In de demo wordt alleen het laatste bestand per module bewaard.', 404); return { b64: f }; }
      if (route === 'teacher/reset') { const s = db.students[b.id]; if (!/^\d{4,8}$/.test(String(b.code || ''))) err('Een code bestaat uit 4 tot 8 cijfers.'); s.code = String(b.code); demoSave(db); return { ok: true }; }
      if (route === 'teacher/module-reset') { delete db.students[b.id].progress[b.moduleId]; demoSave(db); return { ok: true }; }
      if (route === 'teacher/delete') { delete db.students[b.id]; Object.keys(db.files).filter(k => k.startsWith('files/' + b.id + '/')).forEach(k => delete db.files[k]); demoSave(db); return { ok: true }; }
    }
    const s = me();
    if (!s) err('Meld je opnieuw aan.', 401);
    if (route === 'me') return { student: pub(s) };
    if (route === 'progress') { s.progress[b.moduleId] = merge(s.progress[b.moduleId] || {}, b.patch); s.lastSeen = new Date().toISOString(); demoSave(db); return { student: pub(s) }; }
    if (route === 'attempt') {
      const id = Date.now().toString(36);
      let fileKey = null;
      if (b.fileB64) {
        // in de demo enkel het laatste bestand per module bewaren (browseropslag is klein)
        Object.keys(db.files).filter(k => k.startsWith(`files/${s.id}/${b.moduleId}-`)).forEach(k => delete db.files[k]);
        fileKey = `files/${s.id}/${b.moduleId}-${b.mode}-${id}.xlsx`;
        db.files[fileKey] = b.fileB64;
      }
      const attempt = { id, moduleId: b.moduleId, mode: b.mode, seed: b.seed, at: new Date().toISOString(), score: b.score, total: b.total, results: b.results, hints: b.hints, fileName: b.fileName, fileKey, fileModified: b.fileModified };
      s.attempts.push(attempt);
      s.progress[b.moduleId] = merge(s.progress[b.moduleId] || {}, b.patch);
      if (fileKey) s.progress[b.moduleId].lastFile = { key: fileKey, name: b.fileName, at: attempt.at };
      demoSave(db);
      return { student: pub(s), attempt };
    }
    if (route === 'lastfile') {
      const lf = s.progress[qs.get('module')] && s.progress[qs.get('module')].lastFile;
      if (!lf || !db.files[lf.key]) err('Je hebt voor deze module nog geen bestand opgeladen.', 404);
      return { name: lf.name, b64: db.files[lf.key] };
    }
    err('Onbekende vraag.', 404);
  };

  // ---------------- gemeenschappelijk ----------------
  let backend = null, demo = false;
  async function init() {
    if (backend) return { demo };
    try {
      const ctrl = new AbortController(); const to = setTimeout(() => ctrl.abort(), 4000);
      const r = await fetch('/api/ping', { signal: ctrl.signal }); clearTimeout(to);
      const d = await r.json();
      if (d && d.app === 'celvoorcel') { backend = new Remote(); demo = false; return { demo }; }
    } catch (e) { /* geen server: demo */ }
    backend = new Demo(); demo = true;
    return { demo };
  }
  const call = (p, o) => backend.call(p, o);

  const Api = {
    init,
    isDemo: () => demo,
    // leerling
    async login(form) { const d = await call('login', { body: form }); if (d.token) ss.set(TKEY, d.token); return d; },
    logout() { ss.del(TKEY); },
    hasSession: () => !!ss.get(TKEY),
    me: () => call('me'),
    progress: (moduleId, patch) => call('progress', { body: { moduleId, patch } }),
    attempt: body => call('attempt', { body }),
    lastFile: moduleId => call('lastfile?module=' + encodeURIComponent(moduleId)),
    // leerkracht
    async tLogin(wachtwoord) { const d = await call('teacher/login', { body: { wachtwoord } }); ss.set(TKEY + '-t', d.token); return d; },
    tLogout() { ss.del(TKEY + '-t'); },
    tHasSession: () => !!ss.get(TKEY + '-t'),
    tOverview: () => call('teacher/overview', { teacher: true }),
    tStudent: id => call('teacher/student?id=' + encodeURIComponent(id), { teacher: true }),
    tFile: key => call('teacher/file?key=' + encodeURIComponent(key), { teacher: true }),
    tReset: (id, code) => call('teacher/reset', { teacher: true, body: { id, code } }),
    tModuleReset: (id, moduleId) => call('teacher/module-reset', { teacher: true, body: { id, moduleId } }),
    tDelete: id => call('teacher/delete', { teacher: true, body: { id } }),
    // hulp
    async fileToB64(file) {
      const buf = await file.arrayBuffer();
      let s = ''; const bytes = new Uint8Array(buf);
      for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
      return btoa(s);
    },
    b64ToBlob(b64) {
      const bin = atob(b64); const a = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) a[i] = bin.charCodeAt(i);
      return new Blob([a], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    },
    download(blob, name) {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a'); a.href = url; a.download = name; document.body.appendChild(a); a.click();
      setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1500);
    }
  };
  root.CvcApi = Api;
})(window);
