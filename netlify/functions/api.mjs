// Cel voor Cel – opslag van leerlingen, voortgang, pogingen en Excel-bestanden.
// Netlify Function (v2) + Netlify Blobs.
//
// Aanmelden gebeurt via MELK (melk-start.netlify.app):
//  - leerling: MELK geeft een toegangsbewijs; deze functie laat het controleren door het meesterbestand
//    (via de deur, actie llWie) en krijgt leerlingnummer, naam en klas terug. Geen eigen accounts of codes meer.
//  - leerkracht: inloggen met de link per mail (MELK); deze functie vraagt het meesterbestand wie het is
//    (actie lkWie). Een leerkracht ziet enkel de leerlingen van zijn klassen, een beheerder ziet iedereen.
// Omgevingsvariabelen in Netlify:
//  - LEERKRACHT_WACHTWOORD (bestaande waarde laten staan): dient nu enkel nog als geheime sleutel
//    om de toegangsbewijzen van Cel voor Cel te ondertekenen. Niemand hoeft die nog te kennen.
//    (In de plaats mag ook SESSION_SECRET.)
//  - MELK_DEUR (optioneel): ander adres van de deur, standaard https://melk-start.netlify.app/api/deur
import { getStore } from '@netlify/blobs';
import crypto from 'node:crypto';

export const config = { path: '/api/*' };

const MAX_FILE = 3 * 1024 * 1024; // 3 MB
const APP = 'celvoorcel';
const store = () => globalThis.__cvcStore || getStore({ name: 'celvoorcel', consistency: 'strong' });

const json = (data, status = 200) => new Response(JSON.stringify(data), { status, headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' } });
const fail = (msg, status = 400) => json({ error: msg }, status);

// Leest een omgevingsvariabele (werkt met beide manieren van Netlify).
function env(name) {
  let v;
  try { v = globalThis.Netlify && globalThis.Netlify.env && globalThis.Netlify.env.get(name); } catch (e) { v = undefined; }
  return (v || process.env[name] || '').trim();
}
function secret() {
  const s = env('SESSION_SECRET') || env('LEERKRACHT_WACHTWOORD');
  if (!s) throw new Error('De geheime sleutel (LEERKRACHT_WACHTWOORD of SESSION_SECRET) is niet ingesteld in Netlify.');
  return 'cvc:' + s;
}
const b64u = b => Buffer.from(b).toString('base64url');
function sign(payload) {
  const body = b64u(JSON.stringify(payload));
  const mac = crypto.createHmac('sha256', secret()).update(body).digest('base64url');
  return body + '.' + mac;
}
function verify(token) {
  if (!token || !token.includes('.')) return null;
  const [body, mac] = token.split('.');
  const good = crypto.createHmac('sha256', secret()).update(body).digest('base64url');
  if (mac.length !== good.length || !crypto.timingSafeEqual(Buffer.from(mac), Buffer.from(good))) return null;
  const p = JSON.parse(Buffer.from(body, 'base64url').toString());
  if (p.exp && Date.now() > p.exp) return null;
  return p;
}
const hashCode = (salt, code) => crypto.createHash('sha256').update(salt + '|' + String(code).trim()).digest('hex');
const slug = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
const clean = s => String(s || '').replace(/\s+/g, ' ').trim().slice(0, 40);

function publicStudent(s) {
  const { codeHash, salt, ...rest } = s;
  return rest;
}
function deepMerge(target, patch) {
  for (const [k, v] of Object.entries(patch || {})) {
    if (v && typeof v === 'object' && !Array.isArray(v)) target[k] = deepMerge(target[k] && typeof target[k] === 'object' ? target[k] : {}, v);
    else target[k] = v;
  }
  return target;
}

// ---- MELK: het meesterbestand vragen wie iemand is (via de vaste deur) ----
async function deur(actie, token) {
  const adres = env('MELK_DEUR') || 'https://melk-start.netlify.app/api/deur';
  const stop = new AbortController(); const klok = setTimeout(() => stop.abort(), 25000);
  try {
    const r = await fetch(adres, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ actie, token: String(token || ''), app: APP }), signal: stop.signal });
    const d = await r.json().catch(() => null);
    if (!d || typeof d !== 'object') return { ok: false, geenVerbinding: true };
    return d;
  } catch (e) { return { ok: false, geenVerbinding: true }; }
  finally { clearTimeout(klok); }
}
const GEEN_MELK = 'MELK is even niet bereikbaar. Probeer het zo meteen opnieuw.';
// Wat de leerkracht mag zien: een beheerder alles, anders enkel de leerlingen van zijn klassen.
const magZien = (auth, s) => !!s && (auth.b === true || (Array.isArray(auth.kl) && auth.kl.includes(s.klascode)));

async function getStudent(id) { return store().get('students/' + id, { type: 'json' }); }
async function putStudent(s) { await store().setJSON('students/' + s.id, s); }

export default async (req) => {
  const url = new URL(req.url);
  const route = url.pathname.replace(/^\/api\//, '').replace(/\/$/, '');
  const method = req.method;
  const auth = verify((req.headers.get('authorization') || '').replace(/^Bearer\s+/i, ''));
  const body = method === 'POST' ? await req.json().catch(() => ({})) : {};

  try {
    if (route === 'ping') return json({ ok: true, app: 'celvoorcel' });

    // ---------------- leerling (aanmelden via MELK) ----------------
    if (route === 'login') return fail('Je meldt je aan via MELK.', 410);
    if (route === 'melk-login' && method === 'POST') {
      const w = await deur('llWie', body.token);
      if (!w.ok && w.nietVoorJou) return fail(w.fout || 'Cel voor Cel staat niet voor jou open.', 403);
      if (!w.ok) return w.opnieuwInloggen ? fail('Meld je opnieuw aan via MELK.', 401) : fail(w.geenVerbinding ? GEEN_MELK : 'MELK is nog niet klaar voor Cel voor Cel (' + (w.fout || 'onbekende fout') + '). Vraag het aan de beheerder van MELK.', 503);
      const nr = String(w.nr || '').replace(/[^A-Za-z0-9_-]/g, '');
      if (!nr) return fail('Meld je opnieuw aan via MELK.', 401);
      const id = 'melk-' + nr;
      const weergave = clean(w.naam);
      let voornaam = clean(w.voornaam), naam = '';
      if (voornaam && weergave.toLowerCase().startsWith(voornaam.toLowerCase())) naam = weergave.slice(voornaam.length).trim();
      else if (weergave) { voornaam = weergave; naam = ''; }
      const nu = new Date().toISOString();
      let s = await getStudent(id);
      if (!s) s = { id, created: nu, progress: {}, attempts: [] };
      Object.assign(s, { nr, voornaam: voornaam || 'Leerling', naam, klas: clean(w.klasnaam) || String(w.klascode || ''), klascode: String(w.klascode || ''), leerjaar: String(w.leerjaar || ''), schooljaar: String(w.schooljaar || ''), lastSeen: nu });
      await putStudent(s);
      const exp = Math.min(Date.now() + 1000 * 60 * 60 * 12, Number(w.geldigTot) || Date.now() + 1000 * 60 * 60 * 3);
      return json({ token: sign({ role: 's', id, exp }), student: publicStudent(s) });
    }

    if (route.startsWith('teacher/')) {
      // ---------------- leerkracht ----------------
      if (route === 'teacher/login') return fail('Je meldt je aan via MELK (link per mail).', 410);
      if (route === 'teacher/melk' && method === 'POST') {
        const w = await deur('lkWie', body.token);
        if (!w.ok) return w.opnieuwInloggen ? fail('Meld je opnieuw aan als leerkracht.', 401) : fail(w.geenVerbinding ? GEEN_MELK : 'MELK is nog niet klaar voor Cel voor Cel (' + (w.fout || 'onbekende fout') + '). Vraag het aan de beheerder van MELK.', 503);
        const klassen = (w.klassen || []).map(k => ({ klascode: String(k.klascode), klasnaam: String(k.klasnaam || k.klascode) }));
        const token = sign({ role: 't', kl: klassen.map(k => k.klascode), b: w.beheerder === true, naam: clean(w.naam), exp: Date.now() + 1000 * 60 * 60 * 8 });
        return json({ token, naam: clean(w.naam), beheerder: w.beheerder === true, klassen });
      }
      if (!auth || auth.role !== 't' || !Array.isArray(auth.kl)) return fail('Meld je opnieuw aan als leerkracht.', 401);

      if (route === 'teacher/overview') {
        const { blobs } = await store().list({ prefix: 'students/' });
        const all = await Promise.all(blobs.map(b => store().get(b.key, { type: 'json' })));
        return json({ students: all.filter(s => magZien(auth, s)).map(s => { const p = publicStudent(s); p.attempts = (p.attempts || []).map(({ results, ...a }) => a); return p; }) });
      }
      if (route === 'teacher/student') {
        const s = await getStudent(url.searchParams.get('id'));
        return magZien(auth, s) ? json({ student: publicStudent(s) }) : fail('Leerling niet gevonden.', 404);
      }
      if (route === 'teacher/file') {
        const key = url.searchParams.get('key') || '';
        if (!key.startsWith('files/')) return fail('Onbekend bestand.', 404);
        if (!magZien(auth, await getStudent(key.split('/')[1]))) return fail('Bestand niet gevonden.', 404);
        const buf = await store().get(key, { type: 'arrayBuffer' });
        if (!buf) return fail('Bestand niet gevonden.', 404);
        return json({ b64: Buffer.from(buf).toString('base64') });
      }
      if (route === 'teacher/module-reset' && method === 'POST') {
        const s = await getStudent(body.id);
        if (!magZien(auth, s)) return fail('Leerling niet gevonden.', 404);
        delete s.progress[body.moduleId];
        await putStudent(s);
        return json({ ok: true });
      }
      if (route === 'teacher/delete' && method === 'POST') {
        const s = await getStudent(body.id);
        if (!magZien(auth, s)) return fail('Leerling niet gevonden.', 404);
        const { blobs } = await store().list({ prefix: 'files/' + s.id + '/' });
        await Promise.all(blobs.map(b => store().delete(b.key)));
        await store().delete('students/' + s.id);
        return json({ ok: true });
      }
      return fail('Onbekende vraag.', 404);
    }

    // ---------------- vanaf hier: aangemelde leerling ----------------
    if (!auth || auth.role !== 's') return fail('Meld je opnieuw aan.', 401);
    const s = await getStudent(auth.id);
    if (!s) return fail('Je account bestaat niet meer. Meld je opnieuw aan.', 401);

    if (route === 'me') return json({ student: publicStudent(s) });

    if (route === 'progress' && method === 'POST') {
      const mid = String(body.moduleId || '');
      if (!/^m\d+$/.test(mid)) return fail('Onbekende module.');
      s.progress[mid] = deepMerge(s.progress[mid] || {}, body.patch || {});
      s.lastSeen = new Date().toISOString();
      await putStudent(s);
      return json({ student: publicStudent(s) });
    }

    if (route === 'attempt' && method === 'POST') {
      const mid = String(body.moduleId || '');
      if (!/^m\d+$/.test(mid)) return fail('Onbekende module.');
      const id = Date.now().toString(36) + crypto.randomBytes(3).toString('hex');
      let fileKey = null;
      if (body.fileB64) {
        const buf = Buffer.from(body.fileB64, 'base64');
        if (buf.length > MAX_FILE) return fail('Dit bestand is te groot (max. 3 MB).');
        fileKey = `files/${s.id}/${mid}-${body.mode}-${id}.xlsx`;
        await store().set(fileKey, buf);
      }
      const attempt = {
        id, moduleId: mid, mode: body.mode === 'test' ? 'test' : 'practice', seed: String(body.seed || ''),
        at: new Date().toISOString(), score: +body.score || 0, total: +body.total || 0,
        results: Array.isArray(body.results) ? body.results.slice(0, 40).map(r => ({ id: String(r.id), ok: !!r.ok, msg: String(r.msg || '').slice(0, 300) })) : [],
        hints: +body.hints || 0, fileName: String(body.fileName || '').slice(0, 120), fileKey, fileModified: body.fileModified || null
      };
      s.attempts = (s.attempts || []).concat(attempt);
      s.progress[mid] = deepMerge(s.progress[mid] || {}, body.patch || {});
      if (fileKey) s.progress[mid].lastFile = { key: fileKey, name: attempt.fileName, at: attempt.at };
      s.lastSeen = attempt.at;
      await putStudent(s);
      return json({ student: publicStudent(s), attempt });
    }

    if (route === 'lastfile') {
      const mid = url.searchParams.get('module');
      const lf = s.progress[mid] && s.progress[mid].lastFile;
      if (!lf) return fail('Je hebt voor deze module nog geen bestand opgeladen.', 404);
      const buf = await store().get(lf.key, { type: 'arrayBuffer' });
      if (!buf) return fail('Bestand niet gevonden.', 404);
      return json({ name: lf.name, b64: Buffer.from(buf).toString('base64') });
    }

    return fail('Onbekende vraag.', 404);
  } catch (e) {
    return fail(e.message || 'Er ging iets mis.', 500);
  }
};
