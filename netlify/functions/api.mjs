// Cel voor Cel – opslag van leerlingen, voortgang, pogingen en Excel-bestanden.
// Netlify Function (v2) + Netlify Blobs. Omgevingsvariabele: LEERKRACHT_WACHTWOORD (verplicht).
import { getStore } from '@netlify/blobs';
import crypto from 'node:crypto';

export const config = { path: '/api/*' };

const MAX_FILE = 3 * 1024 * 1024; // 3 MB
const store = () => getStore({ name: 'celvoorcel', consistency: 'strong' });

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
  if (!s) throw new Error('LEERKRACHT_WACHTWOORD is niet ingesteld in Netlify.');
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

    // ---------------- leerling ----------------
    if (route === 'login' && method === 'POST') {
      const voornaam = clean(body.voornaam), naam = clean(body.naam), klas = clean(body.klas).toUpperCase(), code = String(body.code || '').trim();
      if (!voornaam || !naam || !klas) return fail('Vul je voornaam, naam en klas in.');
      if (!/^\d{4,8}$/.test(code)) return fail('Je code bestaat uit 4 tot 8 cijfers.');
      const id = slug(klas) + '--' + slug(voornaam) + '--' + slug(naam);
      let s = await getStudent(id);
      if (!s) {
        if (!body.nieuw) return json({ nieuw: true, message: 'Er bestaat nog geen account met deze naam en klas.' });
        const salt = crypto.randomBytes(8).toString('hex');
        s = { id, voornaam, naam, klas, salt, codeHash: hashCode(salt, code), created: new Date().toISOString(), lastSeen: new Date().toISOString(), progress: {}, attempts: [] };
        await putStudent(s);
      } else if (s.codeHash !== hashCode(s.salt, code)) {
        return fail('Deze code klopt niet. Vraag je leerkracht om je code opnieuw in te stellen.', 401);
      } else {
        s.lastSeen = new Date().toISOString();
        await putStudent(s);
      }
      return json({ token: sign({ role: 's', id, exp: Date.now() + 1000 * 60 * 60 * 12 }), student: publicStudent(s) });
    }

    if (route.startsWith('teacher/')) {
      // ---------------- leerkracht ----------------
      if (route === 'teacher/login' && method === 'POST') {
        const pw = String(body.wachtwoord || '');
        const real = env('LEERKRACHT_WACHTWOORD');
        const a = crypto.createHash('sha256').update(pw).digest(), b = crypto.createHash('sha256').update(real).digest();
        if (!real || !crypto.timingSafeEqual(a, b)) return fail('Dit wachtwoord klopt niet.', 401);
        return json({ token: sign({ role: 't', exp: Date.now() + 1000 * 60 * 60 * 12 }) });
      }
      if (!auth || auth.role !== 't') return fail('Meld je opnieuw aan als leerkracht.', 401);

      if (route === 'teacher/overview') {
        const { blobs } = await store().list({ prefix: 'students/' });
        const all = await Promise.all(blobs.map(b => store().get(b.key, { type: 'json' })));
        return json({ students: all.filter(Boolean).map(s => { const p = publicStudent(s); p.attempts = (p.attempts || []).map(({ results, ...a }) => a); return p; }) });
      }
      if (route === 'teacher/student') {
        const s = await getStudent(url.searchParams.get('id'));
        return s ? json({ student: publicStudent(s) }) : fail('Leerling niet gevonden.', 404);
      }
      if (route === 'teacher/file') {
        const key = url.searchParams.get('key') || '';
        if (!key.startsWith('files/')) return fail('Onbekend bestand.', 404);
        const buf = await store().get(key, { type: 'arrayBuffer' });
        if (!buf) return fail('Bestand niet gevonden.', 404);
        return json({ b64: Buffer.from(buf).toString('base64') });
      }
      if (route === 'teacher/reset' && method === 'POST') {
        const s = await getStudent(body.id);
        if (!s) return fail('Leerling niet gevonden.', 404);
        if (!/^\d{4,8}$/.test(String(body.code || ''))) return fail('Een code bestaat uit 4 tot 8 cijfers.');
        s.salt = crypto.randomBytes(8).toString('hex'); s.codeHash = hashCode(s.salt, body.code);
        await putStudent(s);
        return json({ ok: true });
      }
      if (route === 'teacher/module-reset' && method === 'POST') {
        const s = await getStudent(body.id);
        if (!s) return fail('Leerling niet gevonden.', 404);
        delete s.progress[body.moduleId];
        await putStudent(s);
        return json({ ok: true });
      }
      if (route === 'teacher/delete' && method === 'POST') {
        const s = await getStudent(body.id);
        if (!s) return fail('Leerling niet gevonden.', 404);
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
