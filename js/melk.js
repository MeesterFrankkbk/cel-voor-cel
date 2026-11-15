/* =============================================================================
   MELK-KOPPELING voor gewone HTML-apps (zonder bouwstap).
   Zet <script src="melk.js"></script> VÓÓR het script van de app.

   Wat het doet:
   - MELK.start({ app: 'typkampioen' })
       * neemt een nieuwe aanmelding over uit het adres (#melk=...)
       * niemand aangemeld? -> door naar MELK-start (klas, naam, code) en terug
       * toont bovenaan een balk met prentje, naam, klas, "Bewaard" en de knop Klaar
       * na 10 minuten zonder activiteit, na Klaar of bij het sluiten van het tabblad: afgemeld
       * .../leerkracht of #leerkracht: geen aanmelding nodig (leerkrachtenscherm)
       geeft terug: { omgeleid, leerkracht, aanmelding }
   - MELK.stuur({ onderdeel, oefening, opgaven, juist, medianeTijdMs, doeltijdMs, hulp, moeilijk, extra, begin })
       * één pakketje naar het MELK-hart, via een wachtrij op het toestel:
         geen internet = later opnieuw (ook na herladen), nooit dubbel
   - MELK.opStatus(functie(status, aantal))   status: leeg | bezig | bewaard | wacht
   - MELK.actief()   de leerling is bezig (bv. hardop aan het lezen): niet afmelden wegens stilte
   - MELK.leerkracht({ app: 'woordpuzzel', beheerder: true })      (vanaf november 2026)
       * de leerkrachtenzone van een app: inloggen met de link per mail (MELK-start), daarna terug
       * geeft een Promise: { ok, naam, mail, beheerder, klassen: [{ klascode, klasnaam, leerjaar, appAan, instellingen }] }
         met beheerder: true gevraagd en de leerkracht is geen beheerder: { ok: true, geweigerd: true, ... }
   - MELK.slot({ app: 'oefenplein', beheerder: true })  -> Promise met dezelfde gegevens, maar enkel als de toegang mag;
       anders toont het zelf een nette melding ("enkel voor beheerders") en gaat de Promise nooit door.
       Vervangt een wachtwoordscherm in één regel: MELK.slot({...}).then(start)
   - MELK.leerkrachtToken()      toegangsbewijs, voor een achterkant die zelf wil nakijken wie er bewaart
   - MELK.leerkrachtAfmelden()   uit de leerkrachtenzone van deze app
   ============================================================================= */
(function () {
  'use strict';
  var MELK_START = window.MELK_START_ADRES || 'https://melk-start.netlify.app';
  var DEUR = MELK_START + '/api/deur';
  var STIL_MS = 10 * 60 * 1000;
  var K_AANM = 'melk-aanmelding-v1';
  var K_RIJ = 'melk-wachtrij-v1';
  var K_LK = 'melk-leerkracht-v1';
  var APP = '';
  var luisteraars = [], laatste = 'leeg', bezig = false;

  function lees(opslag, k, std) { try { var v = JSON.parse(opslag.getItem(k) || 'null'); return v == null ? std : v; } catch (e) { return std; } }
  function schrijf(opslag, k, v) { try { v == null ? opslag.removeItem(k) : opslag.setItem(k, JSON.stringify(v)); } catch (e) { /* niets */ } }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }

  function tokenInhoud(t) {
    try { var b = t.split('.')[0].replace(/-/g, '+').replace(/_/g, '/'); return JSON.parse(decodeURIComponent(escape(atob(b + '==='.slice((b.length + 3) % 4))))); } catch (e) { return {}; }
  }

  function neemOverUitAdres() {
    var h = new URLSearchParams(location.hash.slice(1));
    var token = h.get('melk');
    if (!token) return;
    var t = tokenInhoud(token);
    schrijf(sessionStorage, K_AANM, {
      token: token, nr: String(t.l || ''), naam: h.get('naam') || '', prentje: h.get('prentje') || '',
      klas: h.get('klas') || '', klasnaam: h.get('klasnaam') || h.get('klas') || '', leerjaar: h.get('leerjaar') || '',
      niveau: h.get('niveau') || '', tot: Number(t.exp) || Date.now() + 3 * 3600 * 1000, laatst: Date.now()
    });
    history.replaceState(null, '', location.pathname + location.search);
  }

  function huidige() {
    var a = lees(sessionStorage, K_AANM, null);
    if (!a) return null;
    if (Date.now() - a.laatst > STIL_MS || Date.now() > a.tot - 60000) { schrijf(sessionStorage, K_AANM, null); return null; }
    return a;
  }
  function activiteit() { var a = lees(sessionStorage, K_AANM, null); if (a) { a.laatst = Date.now(); schrijf(sessionStorage, K_AANM, a); } }

  function naarMelkStart() {
    var terug = location.origin + location.pathname + location.search;
    location.replace(MELK_START + '/?' + (APP ? 'app=' + encodeURIComponent(APP) + '&' : '') + 'terug=' + encodeURIComponent(terug));
  }

  // ---- wachtrij -----------------------------------------------------------
  function meld(s) { laatste = s; var n = lees(localStorage, K_RIJ, []).length; luisteraars.forEach(function (f) { f(s, n); }); }
  function opStatus(f) { luisteraars.push(f); f(laatste, lees(localStorage, K_RIJ, []).length); }

  function verwerk() {
    if (bezig) return Promise.resolve();
    var w = lees(localStorage, K_RIJ, []).filter(function (x) { return x.tot > Date.now(); });
    schrijf(localStorage, K_RIJ, w);
    if (!w.length) { if (laatste === 'bezig' || laatste === 'wacht') meld('bewaard'); return Promise.resolve(); }
    bezig = true; meld('bezig');
    var klaar = w.filter(function (x) { return x.volgende <= Date.now(); });
    return klaar.reduce(function (p, item) {
      return p.then(function () {
        return fetch(DEUR, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ actie: 'inleveren', token: item.token, pakket: item.pakket }) })
          .then(function (r) { return r.json(); }).catch(function () { return null; })
          .then(function (r) {
            var rij = lees(localStorage, K_RIJ, []);
            var i = rij.findIndex(function (x) { return x.pakket.sessienummer === item.pakket.sessienummer; });
            if (i < 0) return;
            if (r && (r.ok || r.opnieuwInloggen)) rij.splice(i, 1);
            else { rij[i].pogingen++; rij[i].volgende = Date.now() + Math.min(300000, 15000 * rij[i].pogingen); }
            schrijf(localStorage, K_RIJ, rij);
          });
      });
    }, Promise.resolve()).then(function () { bezig = false; meld(lees(localStorage, K_RIJ, []).length ? 'wacht' : 'bewaard'); });
  }

  function nieuwNummer() { try { return crypto.randomUUID(); } catch (e) { return Date.now().toString(36) + Math.random().toString(36).slice(2); } }

  function stuur(p) {
    var a = huidige();
    if (!a) return false;
    var nu = new Date().toISOString();
    var pakket = {
      sessienummer: nieuwNummer(), app: APP, onderdeel: String(p.onderdeel || ''), oefening: String(p.oefening || ''),
      begin: p.begin ? new Date(p.begin).toISOString() : nu, einde: nu,
      opgaven: Math.max(0, Math.round(p.opgaven || 0)), juist: Math.max(0, Math.round(p.juist || 0)),
      medianeTijdMs: p.medianeTijdMs ? Math.round(p.medianeTijdMs) : undefined, hulp: p.hulp || 0,
      doeltijdMs: p.doeltijdMs ? Math.round(p.doeltijdMs) : undefined, // optioneel: eigen doeltijd per opgave
      moeilijk: (p.moeilijk || []).slice(0, 10).map(String), extra: p.extra || {}
    };
    var rij = lees(localStorage, K_RIJ, []);
    rij.push({ pakket: pakket, token: a.token, tot: a.tot, pogingen: 0, volgende: 0 });
    schrijf(localStorage, K_RIJ, rij);
    verwerk();
    return true;
  }

  function wachtTotLeeg(ms) {
    var eind = Date.now() + ms;
    function lus() {
      return verwerk().then(function () {
        if (!lees(localStorage, K_RIJ, []).length || Date.now() > eind) return;
        return new Promise(function (r) { setTimeout(r, 300); }).then(lus);
      });
    }
    return lus();
  }

  function klaar() {
    var k = document.getElementById('melk-klaar');
    if (k) { k.disabled = true; k.textContent = 'Even bewaren…'; }
    return wachtTotLeeg(4000).then(function () { schrijf(sessionStorage, K_AANM, null); naarMelkStart(); });
  }

  // ---- naambalk -------------------------------------------------------------
  var DINGEN = { kat: '🐱', hond: '🐶', uil: '🦉', vis: '🐟', raket: '🚀', vos: '🦊', beer: '🐻', konijn: '🐰', leeuw: '🦁', kikker: '🐸', vlinder: '🦋',
    schildpad: '🐢', olifant: '🐘', pinguin: '🐧', bij: '🐝', muis: '🐭', aap: '🐵', paard: '🐴', koe: '🐮', varken: '🐷', eend: '🦆', walvis: '🐳', slak: '🐌',
    egel: '🦔', panda: '🐼', tijger: '🐯', giraf: '🦒', dino: '🦕', auto: '🚗', trein: '🚂', fiets: '🚲', vliegtuig: '✈️', bloem: '🌸', boom: '🌳', zon: '☀️',
    maan: '🌙', ster: '⭐', regenboog: '🌈', appel: '🍎', peer: '🍐', aardbei: '🍓', banaan: '🍌', bal: '⚽', gitaar: '🎸', kroon: '👑', hart: '💖' };
  var KLEUREN = { rood: '#f6c1bd', oranje: '#fbd3a8', geel: '#fbeba0', groen: '#c8e6b0', blauw: '#bcd9f5', paars: '#dccaf2', roze: '#f8c9df', bruin: '#e1cbb3', grijs: '#dcdcdc' };

  function toonNaambalk(a) {
    var d = String(a.prentje || '').toLowerCase().split('-');
    var beeld = DINGEN[d[0]] ? DINGEN[d[0]] : /^[a-z0-9_]+$/.test(d[0] || '') ?
      '<img src="' + MELK_START + '/prentjes/' + d[0] + '.png" alt="" style="width:26px;height:26px;object-fit:contain" onerror="this.replaceWith(document.createTextNode(\'' + esc(a.naam.charAt(0)) + '\'))">' : esc(a.naam.charAt(0));
    var balk = document.createElement('div');
    balk.id = 'melk-balk';
    balk.className = 'no-print';
    balk.setAttribute('style', 'position:sticky;top:0;z-index:9999;display:flex;align-items:center;justify-content:flex-end;gap:12px;padding:6px 14px;' +
      'background:#f3f8ea;border-bottom:1px solid #d5e3c0;font-family:system-ui,"Segoe UI",Arial,sans-serif;color:#23301a');
    balk.innerHTML =
      '<span style="display:inline-flex;align-items:center;justify-content:center;width:34px;height:34px;border-radius:50%;font-size:20px;font-weight:800;background:' + (KLEUREN[d[1]] || '#e6f2d0') + '">' + beeld + '</span>' +
      '<span style="line-height:1.15"><b id="melk-naam" style="font-size:15px">' + esc(a.naam) + '</b><br><small style="color:#5e6b55">Klas ' + esc(a.klasnaam) + '</small></span>' +
      '<span id="melk-status" data-status="leeg" style="font-size:13px;font-weight:700;min-width:110px"></span>' +
      '<button id="melk-klaar" type="button" style="font:inherit;font-weight:800;font-size:14px;border:0;border-radius:10px;background:#5a8f29;color:#fff;padding:8px 16px;cursor:pointer">Klaar</button>';
    document.body.insertBefore(balk, document.body.firstChild);
    document.getElementById('melk-klaar').onclick = klaar;
    opStatus(function (s, n) {
      var el = document.getElementById('melk-status');
      el.setAttribute('data-status', s);
      el.style.color = s === 'wacht' ? '#a15c00' : s === 'bewaard' ? '#2f6b1a' : '#3f4ea0';
      el.textContent = s === 'bezig' ? 'Bewaren…' : s === 'bewaard' ? '✓ Bewaard' : s === 'wacht' ? n + ' nog te bewaren' : '';
      el.title = s === 'wacht' ? 'Geen verbinding: de app probeert het zo opnieuw. Er gaat niets verloren.' : '';
    });
  }

  function start(opts) {
    APP = (opts && opts.app) || '';
    neemOverUitAdres();
    var leerkracht = /^\/leerkracht/i.test(location.pathname) || location.hash === '#leerkracht' || /(^#|&)melklk=/.test(location.hash);
    var a = huidige();
    if (!a && !leerkracht) { naarMelkStart(); return { omgeleid: true, leerkracht: false, aanmelding: null }; }
    if (a) {
      ['pointerdown', 'keydown', 'scroll', 'touchstart'].forEach(function (t) { window.addEventListener(t, activiteit, { passive: true }); });
      setInterval(function () { if (!huidige()) klaar(); }, 15000);
      var toon = function () { toonNaambalk(a); };
      if (document.body) toon(); else document.addEventListener('DOMContentLoaded', toon);
    }
    window.addEventListener('online', verwerk);
    setInterval(verwerk, 20000);
    verwerk();
    return { omgeleid: false, leerkracht: leerkracht, aanmelding: a };
  }

  // ---- leerkrachtenzone (login met de link per mail via MELK-start) -----------
  function neemLkOverUitAdres() {
    var h = new URLSearchParams(location.hash.slice(1));
    var t = h.get('melklk');
    if (!t) return false;
    schrijf(sessionStorage, K_LK, { token: t });
    history.replaceState(null, '', location.pathname + location.search);
    return true;
  }
  function naarLeerkrachtLogin(app) {
    var terug = location.origin + location.pathname + location.search;
    location.replace(MELK_START + '/leerkracht.html?' + (app ? 'app=' + encodeURIComponent(app) + '&' : '') + 'terug=' + encodeURIComponent(terug));
  }
  function leerkrachtToken() { var x = lees(sessionStorage, K_LK, null); return x ? x.token : ''; }
  function leerkrachtAfmelden() { schrijf(sessionStorage, K_LK, null); }
  function leerkracht(opts) {
    opts = opts || {};
    var app = opts.app || APP;
    neemLkOverUitAdres();
    var token = leerkrachtToken();
    if (!token) { naarLeerkrachtLogin(app); return new Promise(function () {}); }
    return fetch(DEUR, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ actie: 'lkWie', token: token, app: app }) })
      .then(function (r) { return r.json(); })
      .catch(function () { return { ok: false, fout: 'Geen verbinding met MELK. Probeer het zo meteen opnieuw.' }; })
      .then(function (r) {
        if (r && r.opnieuwInloggen) { leerkrachtAfmelden(); naarLeerkrachtLogin(app); return new Promise(function () {}); }
        if (!r || !r.ok) return { ok: false, fout: (r && r.fout) || 'Er liep iets mis.' };
        if (opts.beheerder && !r.beheerder) r.geweigerd = true;
        return r;
      });
  }

  function slotScherm(html) {
    var d = document.getElementById('melk-slot');
    if (!d) {
      d = document.createElement('div'); d.id = 'melk-slot';
      d.style.cssText = 'position:fixed;inset:0;z-index:99999;display:flex;align-items:center;justify-content:center;padding:16px;background:#f6f8f1;font:16px/1.5 system-ui,-apple-system,Segoe UI,Arial,sans-serif;color:#223';
      (document.body || document.documentElement).appendChild(d);
    }
    d.innerHTML = '<div style="max-width:440px;background:#fff;border:1px solid #dfe6d3;border-radius:16px;padding:22px 24px;text-align:center">' + html + '</div>';
    return d;
  }
  function slot(opts) {
    opts = opts || {};
    var wacht = function () { slotScherm('<p style="margin:0">Even nakijken wie je bent…</p>'); };
    if (document.body) wacht(); else document.addEventListener('DOMContentLoaded', wacht);
    return leerkracht(opts).then(function (r) {
      var weg = function () { var d = document.getElementById('melk-slot'); if (d) d.remove(); };
      if (!r.ok) { slotScherm('<p style="margin:0 0 8px;font-weight:700">MELK is even niet bereikbaar</p><p style="margin:0">' + esc(r.fout) + '</p>'); return new Promise(function () {}); }
      if (r.geweigerd) {
        slotScherm('<p style="font-size:30px;margin:0">🔒</p><p style="margin:6px 0;font-weight:700">Dit deel is enkel voor beheerders</p>' +
          '<p style="margin:0 0 12px">Je bent aangemeld als ' + esc(r.naam) + '. Vraag het aan de beheerder van MELK als je hier toegang nodig hebt.</p>' +
          '<a href="' + esc(MELK_START) + '/leerkracht.html" style="color:#2f6b1a;font-weight:700">Naar MELK voor leerkrachten</a>');
        return new Promise(function () {});
      }
      weg();
      return r;
    });
  }

  // actief(): de app meldt zelf dat er gewerkt wordt (bv. hardop lezen zonder te klikken)
  window.MELK = { start: start, stuur: stuur, opStatus: opStatus, klaar: klaar, huidige: huidige, actief: activiteit,
    leerkracht: leerkracht, slot: slot, leerkrachtToken: leerkrachtToken, leerkrachtAfmelden: leerkrachtAfmelden };
})();
