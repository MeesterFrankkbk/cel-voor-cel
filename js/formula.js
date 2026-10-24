/* Cel voor Cel – een kleine formule-rekenaar, zoals Excel.
   - locale 'en': zoals formules in een .xlsx-bestand staan (SUM, komma als scheidingsteken, punt als decimaal)
   - locale 'nl': zoals een leerling in een Belgische Excel typt (SOM, puntkomma, komma als decimaal)
   Werkt in de browser (window.CvcFormula) en in Node (module.exports). */
(function (root) {
  'use strict';

  // ---------------- fouten ----------------
  function Err(code) { this.error = code; }
  const isErr = v => v instanceof Err;
  const ERR_NL = { '#DIV/0!': '#DEEL/0!', '#NAME?': '#NAAM?', '#VALUE!': '#WAARDE!', '#REF!': '#VERW!', '#N/A': '#N/B', '#NUM!': '#GETAL!', '#NULL!': '#LEEG!' };
  const ERR_EN = Object.fromEntries(Object.entries(ERR_NL).map(([a, b]) => [b, a]));

  // ---------------- functienamen ----------------
  const NL_EN = {
    'SOM': 'SUM', 'GEMIDDELDE': 'AVERAGE', 'MIN': 'MIN', 'MAX': 'MAX', 'MEDIAAN': 'MEDIAN', 'AANTAL': 'COUNT', 'AANTALARG': 'COUNTA',
    'AANTAL.LEGE.CELLEN': 'COUNTBLANK', 'AFRONDEN': 'ROUND', 'AFRONDEN.NAAR.BOVEN': 'ROUNDUP', 'AFRONDEN.NAAR.BENEDEN': 'ROUNDDOWN', 'INTEGER': 'INT', 'ABS': 'ABS',
    'ALS': 'IF', 'EN': 'AND', 'OF': 'OR', 'NIET': 'NOT', 'AANTAL.ALS': 'COUNTIF', 'SOM.ALS': 'SUMIF', 'ALS.FOUT': 'IFERROR',
    'SPATIES.WISSEN': 'TRIM', 'KLEINE.LETTERS': 'LOWER', 'HOOFDLETTERS': 'UPPER', 'LENGTE': 'LEN', 'TEKST.SAMENVOEGEN': 'CONCATENATE', 'RANG': 'RANK', 'GROOTSTE': 'LARGE', 'KLEINSTE': 'SMALL'
  };
  const EN_NL = Object.fromEntries(Object.entries(NL_EN).map(([a, b]) => [b, a]));
  const TRUE_NAMES = { en: ['TRUE'], nl: ['WAAR'] }, FALSE_NAMES = { en: ['FALSE'], nl: ['ONWAAR'] };

  // ---------------- adressen ----------------
  const colToNum = c => [...c.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  function numToCol(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  const REF = /^\$?([A-Za-z]{1,3})\$?(\d+)$/;
  function parseRef(r) { const m = REF.exec(r); return m ? { c: colToNum(m[1]), r: +m[2] } : null; }
  function expandRange(a, b) {
    const p = parseRef(a), q = parseRef(b), out = [];
    const c1 = Math.min(p.c, q.c), c2 = Math.max(p.c, q.c), r1 = Math.min(p.r, q.r), r2 = Math.max(p.r, q.r);
    if ((c2 - c1 + 1) * (r2 - r1 + 1) > 20000) throw new Err('#REF!');
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) out.push(numToCol(c) + r);
    return { cells: out, rows: r2 - r1 + 1, cols: c2 - c1 + 1 };
  }
  const clean = r => r.replace(/\$/g, '').toUpperCase();

  // ---------------- tokens ----------------
  function tokenize(src, locale) {
    const dec = locale === 'nl' ? ',' : '.', sep = locale === 'nl' ? ';' : ',';
    const s = String(src).replace(/^\s*=/, '');
    const out = [];
    let i = 0;
    while (i < s.length) {
      const ch = s[i];
      if (/\s/.test(ch)) { i++; continue; }
      if (ch === '"') {
        let j = i + 1, str = '';
        while (j < s.length) { if (s[j] === '"') { if (s[j + 1] === '"') { str += '"'; j += 2; continue; } break; } str += s[j++]; }
        if (j >= s.length) throw new Err('#VALUE!');
        out.push({ t: 'str', v: str }); i = j + 1; continue;
      }
      if (ch === '#') { // fout als waarde, bv. #N/A
        const m = /^#[A-Z0-9\/]+[!?]?/i.exec(s.slice(i)); out.push({ t: 'err', v: m ? m[0] : '#' }); i += m ? m[0].length : 1; continue;
      }
      // getal
      const numRe = locale === 'nl' ? /^\d+(,\d+)?/ : /^\d+(\.\d+)?([eE][+-]?\d+)?/;
      const numM = numRe.exec(s.slice(i));
      if (numM && !/^\d+[A-Za-z]/.test(s.slice(i)) && !(/^\d+:/.test(s.slice(i)))) {
        out.push({ t: 'num', v: parseFloat(numM[0].replace(dec, '.')), raw: numM[0] }); i += numM[0].length; continue;
      }
      // bladnaam 'Mijn blad'!A1
      let sheet = null;
      if (ch === "'") {
        const j = s.indexOf("'!", i + 1);
        if (j > 0) { sheet = s.slice(i + 1, j).replace(/''/g, "'"); i = j + 2; }
      }
      const idM = /^[A-Za-z_À-ɏ][A-Za-z0-9_.À-ɏ]*/.exec(s.slice(i)) || /^\$?[A-Za-z]{1,3}\$?\d+/.exec(s.slice(i));
      if (idM) {
        let id = idM[0];
        if (!sheet && s[i + id.length] === '!') { sheet = id; i += id.length + 1; const m2 = /^\$?[A-Za-z]{1,3}\$?\d+/.exec(s.slice(i)); if (!m2) throw new Err('#REF!'); id = m2[0]; }
        else if (sheet) { const m2 = /^\$?[A-Za-z]{1,3}\$?\d+/.exec(s.slice(i)); if (!m2) throw new Err('#REF!'); id = m2[0]; }
        i += id.length;
        if (REF.test(id) && s[i] !== '(') {
          // bereik?
          const rm = /^:(\$?[A-Za-z]{1,3}\$?\d+)/.exec(s.slice(i));
          if (rm) { out.push({ t: 'range', a: id, b: rm[1], sheet }); i += rm[0].length; }
          else out.push({ t: 'ref', v: id, sheet });
          continue;
        }
        if (s[i] === '(') { out.push({ t: 'func', v: id.toUpperCase() }); i++; continue; }
        out.push({ t: 'name', v: id.toUpperCase() });
        continue;
      }
      const two = s.slice(i, i + 2);
      if (['<>', '<=', '>='].includes(two)) { out.push({ t: 'op', v: two }); i += 2; continue; }
      if ('+-*/^&=<>%'.includes(ch)) { out.push({ t: 'op', v: ch }); i++; continue; }
      if (ch === '(' || ch === ')') { out.push({ t: ch }); i++; continue; }
      if (ch === sep) { out.push({ t: 'sep' }); i++; continue; }
      if (ch === '×' || ch === 'x' || ch === 'X' || ch === ':' || ch === ',' || ch === ';' || ch === '.') { out.push({ t: 'bad', v: ch }); i++; continue; }
      out.push({ t: 'bad', v: ch }); i++;
    }
    return out;
  }

  // ---------------- ontleden (AST) ----------------
  function parse(src, locale) {
    const tk = tokenize(src, locale);
    let p = 0;
    const peek = () => tk[p], next = () => tk[p++];
    const isOp = (v) => peek() && peek().t === 'op' && peek().v === v;
    function expr() { return comparison(); }
    function comparison() {
      let l = concat();
      while (peek() && peek().t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(peek().v)) { const op = next().v; l = { k: 'bin', op, l, r: concat() }; }
      return l;
    }
    function concat() { let l = additive(); while (isOp('&')) { next(); l = { k: 'bin', op: '&', l, r: additive() }; } return l; }
    function additive() { let l = term(); while (isOp('+') || isOp('-')) { const op = next().v; l = { k: 'bin', op, l, r: term() }; } return l; }
    function term() { let l = power(); while (isOp('*') || isOp('/')) { const op = next().v; l = { k: 'bin', op, l, r: power() }; } return l; }
    function power() { let l = unary(); while (isOp('^')) { next(); l = { k: 'bin', op: '^', l, r: unary() }; } return l; }
    function unary() { if (isOp('-')) { next(); return { k: 'neg', e: unary() }; } if (isOp('+')) { next(); return unary(); } return postfix(); }
    function postfix() { let e = primary(); while (isOp('%')) { next(); e = { k: 'pct', e }; } return e; }
    function primary() {
      const t = next();
      if (!t) throw new Err('#VALUE!');
      if (t.t === 'num') return { k: 'num', v: t.v };
      if (t.t === 'str') return { k: 'str', v: t.v };
      if (t.t === 'err') return { k: 'err', v: ERR_EN[t.v.toUpperCase()] || t.v.toUpperCase() };
      if (t.t === 'ref') return { k: 'ref', v: t.v, sheet: t.sheet };
      if (t.t === 'range') return { k: 'range', a: t.a, b: t.b, sheet: t.sheet };
      if (t.t === 'name') {
        if (TRUE_NAMES[locale].includes(t.v)) return { k: 'bool', v: true };
        if (FALSE_NAMES[locale].includes(t.v)) return { k: 'bool', v: false };
        return { k: 'badname', v: t.v };
      }
      if (t.t === 'func') {
        const args = [];
        if (peek() && peek().t === ')') { next(); return { k: 'call', name: t.v, args }; }
        while (true) {
          if (peek() && (peek().t === 'sep' || peek().t === ')')) args.push({ k: 'empty' });
          else args.push(expr());
          const n = next();
          if (!n) throw new Err('#VALUE!');
          if (n.t === ')') break;
          if (n.t !== 'sep') throw new Err(n.t === 'bad' && (n.v === ',' || n.v === ';') ? '#SEP' : '#VALUE!');
        }
        return { k: 'call', name: t.v, args };
      }
      if (t.t === '(') { const e = expr(); const c = next(); if (!c || c.t !== ')') throw new Err('#VALUE!'); return e; }
      if (t.t === 'bad') return { k: 'badtok', v: t.v };
      throw new Err('#VALUE!');
    }
    const ast = expr();
    if (p < tk.length) {
      const rest = tk[p];
      return { k: 'trailing', ast, tok: rest };
    }
    return ast;
  }

  // ---------------- rekenen ----------------
  function toNum(v, locale) {
    if (isErr(v)) return v;
    if (v == null || v === '') return v === '' ? new Err('#VALUE!') : 0;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    const t = String(v).trim().replace(locale === 'nl' ? ',' : /^$/, '.');
    if (/^-?\d+(\.\d+)?$/.test(t)) return parseFloat(t);
    return new Err('#VALUE!');
  }
  function toText(v, locale) {
    if (v == null) return '';
    if (typeof v === 'boolean') return v ? (locale === 'nl' ? 'WAAR' : 'TRUE') : (locale === 'nl' ? 'ONWAAR' : 'FALSE');
    if (typeof v === 'number') { const s = String(+v.toPrecision(15)); return locale === 'nl' ? s.replace('.', ',') : s; }
    return String(v);
  }
  function cmp(a, b) {
    const rank = v => typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 0;
    if (a == null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
    if (b == null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
    if (rank(a) !== rank(b)) return rank(a) - rank(b);
    if (typeof a === 'string') { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
    return a < b ? -1 : a > b ? 1 : 0;
  }
  const r15 = x => +x.toPrecision(15);
  function roundHalfAway(x, n) {
    const f = Math.pow(10, n);
    const y = Math.abs(x) * f;
    const r = Math.floor(+(y + 0.5).toPrecision(15));
    return Math.sign(x) * r / f;
  }
  function critFn(crit, locale) {
    if (isErr(crit)) return () => false;
    let op = '=', val = crit;
    if (typeof crit === 'string') {
      const m = /^(<>|<=|>=|=|<|>)?(.*)$/.exec(crit);
      op = m[1] || '='; val = m[2];
      const n = toNum(val, locale);
      if (!isErr(n) && val.trim() !== '') val = n;
    }
    return v => {
      if (op === '=' && typeof val === 'string' && val === '') return v == null || v === '';
      if (typeof val === 'number' && typeof v !== 'number') return op === '<>';
      if (typeof val === 'string' && typeof v !== 'string') return op === '<>';
      const c = cmp(v, val);
      return { '=': c === 0, '<>': c !== 0, '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 }[op];
    };
  }

  function evaluate(src, ctx) {
    const locale = ctx.locale || 'en';
    let ast;
    try { ast = parse(src, locale); } catch (e) { if (isErr(e)) return e.error === '#SEP' ? new Err('#VALUE!') : e; throw e; }
    const get = (a, sheet) => { const v = ctx.get(clean(a), sheet); return v === undefined ? null : v; };
    function rangeVals(n) { const ex = expandRange(n.a, n.b); return { range: true, values: ex.cells.map(c => get(c, n.sheet)), rows: ex.rows, cols: ex.cols, cells: ex.cells }; }
    function ev(n) {
      switch (n.k) {
        case 'num': case 'str': case 'bool': return n.v;
        case 'err': return new Err(n.v);
        case 'empty': return null;
        case 'ref': return get(n.v, n.sheet);
        case 'range': return rangeVals(n);
        case 'badname': case 'badtok': return new Err('#NAME?');
        case 'trailing': return new Err('#NAME?');
        case 'neg': { const v = toNum(scalar(ev(n.e)), locale); return isErr(v) ? v : -v; }
        case 'pct': { const v = toNum(scalar(ev(n.e)), locale); return isErr(v) ? v : v / 100; }
        case 'bin': {
          const a = scalar(ev(n.l)), b = scalar(ev(n.r));
          if (isErr(a)) return a; if (isErr(b)) return b;
          if (n.op === '&') return toText(a, locale) + toText(b, locale);
          if (['=', '<>', '<', '>', '<=', '>='].includes(n.op)) { const c = cmp(a, b); return { '=': c === 0, '<>': c !== 0, '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 }[n.op]; }
          const x = toNum(a, locale), y = toNum(b, locale);
          if (isErr(x)) return x; if (isErr(y)) return y;
          if (n.op === '+') return r15(x + y);
          if (n.op === '-') return r15(x - y);
          if (n.op === '*') return r15(x * y);
          if (n.op === '/') return y === 0 ? new Err('#DIV/0!') : r15(x / y);
          if (n.op === '^') return r15(Math.pow(x, y));
          return new Err('#VALUE!');
        }
        case 'call': return call(n);
      }
      return new Err('#VALUE!');
    }
    function scalar(v) {
      if (v && v.range) { if (v.values.length === 1) return v.values[0]; return new Err('#VALUE!'); }
      return v;
    }
    function nums(args, mode) {
      // mode 'strict': tekst als los argument geeft #WAARDE!
      const out = [];
      for (const a of args) {
        if (a && a.range) { for (const v of a.values) { if (isErr(v)) return v; if (typeof v === 'number') out.push(v); } }
        else if (a == null) { if (mode === 'count') continue; out.push(0); }
        else { if (isErr(a)) return a; const x = toNum(a, locale); if (isErr(x)) { if (mode === 'count') continue; return x; } out.push(x); }
      }
      return out;
    }
    function call(n) {
      let name = n.name;
      if (locale === 'nl') { if (!NL_EN[name]) return new Err('#NAME?'); name = NL_EN[name]; }
      else if (!EN_NL[name]) return new Err('#NAME?');
      if (name === 'IF') {
        if (n.args.length < 2 || n.args.length > 3) return new Err('#VALUE!');
        const c = scalar(ev(n.args[0])); if (isErr(c)) return c;
        const t = typeof c === 'string' ? (c.toUpperCase() === 'TRUE' || c.toUpperCase() === 'WAAR' ? true : c.toUpperCase() === 'FALSE' || c.toUpperCase() === 'ONWAAR' ? false : new Err('#VALUE!')) : !!(typeof c === 'number' ? c !== 0 : c);
        if (isErr(t)) return t;
        const pick = t ? n.args[1] : n.args[2];
        if (!pick) return false;
        const v = scalar(ev(pick)); return v == null ? 0 : v;
      }
      if (name === 'IFERROR') { const v = scalar(ev(n.args[0])); return isErr(v) ? scalar(ev(n.args[1])) : v; }
      const args = n.args.map(ev);
      const firstErr = args.find(a => isErr(a));
      switch (name) {
        case 'SUM': { const v = nums(args); return isErr(v) ? v : r15(v.reduce((a, b) => a + b, 0)); }
        case 'AVERAGE': { const v = nums(args); if (isErr(v)) return v; return v.length ? r15(v.reduce((a, b) => a + b, 0) / v.length) : new Err('#DIV/0!'); }
        case 'MIN': { const v = nums(args); if (isErr(v)) return v; return v.length ? Math.min(...v) : 0; }
        case 'MAX': { const v = nums(args); if (isErr(v)) return v; return v.length ? Math.max(...v) : 0; }
        case 'MEDIAN': { const v = nums(args); if (isErr(v)) return v; if (!v.length) return new Err('#NUM!'); v.sort((a, b) => a - b); const m = Math.floor(v.length / 2); return v.length % 2 ? v[m] : r15((v[m - 1] + v[m]) / 2); }
        case 'COUNT': { const v = nums(args, 'count'); return isErr(v) ? 0 : v.length; }
        case 'COUNTA': { let c = 0; args.forEach(a => { if (a && a.range) a.values.forEach(v => { if (v != null && v !== '') c++; }); else if (a != null) c++; }); return c; }
        case 'COUNTBLANK': { let c = 0; args.forEach(a => { if (a && a.range) a.values.forEach(v => { if (v == null || v === '') c++; }); }); return c; }
        case 'ROUND': case 'ROUNDUP': case 'ROUNDDOWN': {
          if (firstErr) return firstErr;
          const x = toNum(scalar(args[0]), locale), d = args.length > 1 ? toNum(scalar(args[1]), locale) : 0;
          if (isErr(x)) return x; if (isErr(d)) return d;
          if (name === 'ROUND') return roundHalfAway(x, Math.trunc(d));
          const f = Math.pow(10, Math.trunc(d)), y = +(Math.abs(x) * f).toPrecision(15);
          return Math.sign(x) * (name === 'ROUNDUP' ? Math.ceil(y) : Math.floor(y)) / f;
        }
        case 'INT': { const x = toNum(scalar(args[0]), locale); return isErr(x) ? x : Math.floor(x); }
        case 'ABS': { const x = toNum(scalar(args[0]), locale); return isErr(x) ? x : Math.abs(x); }
        case 'AND': case 'OR': {
          const vals = []; args.forEach(a => { if (a && a.range) a.values.forEach(v => { if (v != null && typeof v !== 'string') vals.push(v); }); else vals.push(a); });
          for (const v of vals) if (isErr(v)) return v;
          const bools = vals.map(v => typeof v === 'number' ? v !== 0 : !!v);
          return name === 'AND' ? bools.every(Boolean) : bools.some(Boolean);
        }
        case 'NOT': { const v = scalar(args[0]); return isErr(v) ? v : !(typeof v === 'number' ? v !== 0 : !!v); }
        case 'COUNTIF': { const r = args[0]; if (!r || !r.range) return new Err('#VALUE!'); const f = critFn(scalar(args[1]), locale); return r.values.filter(f).length; }
        case 'SUMIF': {
          const r = args[0]; if (!r || !r.range) return new Err('#VALUE!'); const f = critFn(scalar(args[1]), locale);
          const s = args[2] && args[2].range ? args[2] : r; let t = 0;
          r.values.forEach((v, k) => { if (f(v) && typeof s.values[k] === 'number') t += s.values[k]; }); return r15(t);
        }
        case 'TRIM': { const v = scalar(args[0]); return isErr(v) ? v : toText(v, locale).replace(/\s+/g, ' ').trim(); }
        case 'LOWER': { const v = scalar(args[0]); return isErr(v) ? v : toText(v, locale).toLowerCase(); }
        case 'UPPER': { const v = scalar(args[0]); return isErr(v) ? v : toText(v, locale).toUpperCase(); }
        case 'LEN': { const v = scalar(args[0]); return isErr(v) ? v : toText(v, locale).length; }
        case 'CONCATENATE': { if (firstErr) return firstErr; return args.map(a => toText(scalar(a), locale)).join(''); }
        case 'LARGE': case 'SMALL': {
          const v = nums([args[0]]); if (isErr(v)) return v; const k = toNum(scalar(args[1]), locale); if (isErr(k)) return k;
          v.sort((a, b) => name === 'LARGE' ? b - a : a - b); return k >= 1 && k <= v.length ? v[k - 1] : new Err('#NUM!');
        }
        case 'RANK': {
          const x = toNum(scalar(args[0]), locale), v = nums([args[1]]); if (isErr(x)) return x; if (isErr(v)) return v;
          const asc = args[2] != null && toNum(scalar(args[2]), locale) !== 0;
          if (!v.includes(x)) return new Err('#N/A');
          return 1 + v.filter(y => asc ? y < x : y > x).length;
        }
      }
      return new Err('#NAME?');
    }
    try {
      const v = scalar(ev(ast));
      return v && v.range ? new Err('#VALUE!') : v;
    } catch (e) { if (isErr(e)) return e; throw e; }
  }

  // ---------------- wat staat er in een formule? ----------------
  function info(src, locale) {
    const res = { cells: [], ranges: [], funcs: [], numbers: [], hasRef: false, sheets: [], error: null, badNames: [] };
    let ast;
    try { ast = parse(src, locale); } catch (e) { res.error = isErr(e) ? e.error : 'parse'; return res; }
    (function walk(n) {
      if (!n) return;
      switch (n.k) {
        case 'num': res.numbers.push(n.v); break;
        case 'ref': res.hasRef = true; res.cells.push(clean(n.v)); if (n.sheet) res.sheets.push(n.sheet); break;
        case 'range': res.hasRef = true; res.ranges.push(clean(n.a) + ':' + clean(n.b)); expandRange(n.a, n.b).cells.forEach(c => res.cells.push(c)); if (n.sheet) res.sheets.push(n.sheet); break;
        case 'call': res.funcs.push(locale === 'nl' ? (NL_EN[n.name] || n.name) : n.name); n.args.forEach(walk); break;
        case 'bin': walk(n.l); walk(n.r); break;
        case 'neg': case 'pct': walk(n.e); break;
        case 'badname': res.badNames.push(n.v); break;
        case 'badtok': res.badNames.push(n.v); break;
        case 'trailing': walk(n.ast); res.badNames.push(n.tok.v || n.tok.t); break;
      }
    })(ast);
    res.ops = (String(src).match(/[+\-*/^]/g) || []);
    return res;
  }

  // ---------------- verschuiven (doortrekken) ----------------
  function shift(src, dr, dc) {
    const s = String(src);
    let out = '', i = 0;
    while (i < s.length) {
      const ch = s[i];
      if (ch === '"' || ch === "'") { const j = s.indexOf(ch, i + 1); const end = j < 0 ? s.length : j + 1; out += s.slice(i, end); i = end; continue; }
      const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)/.exec(s.slice(i));
      const prev = i > 0 ? s[i - 1] : '';
      if (m && !/[A-Za-z0-9_.]/.test(prev) && !/[A-Za-z0-9_(]/.test(s[i + m[0].length] || '')) {
        const c = m[1] ? colToNum(m[2]) : colToNum(m[2]) + dc;
        const r = m[3] ? +m[4] : +m[4] + dr;
        if (c < 1 || r < 1) out += '#REF!';
        else out += m[1] + numToCol(c) + m[3] + r;
        i += m[0].length; continue;
      }
      // identifiers die geen adres zijn in één keer overslaan (bv. functienamen)
      const id = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(s.slice(i));
      if (id) { out += id[0]; i += id[0].length; continue; }
      out += ch; i++;
    }
    return out;
  }

  /* Hoe toon je een waarde in een cel (Belgische Excel)? */
  function display(v, locale) {
    locale = locale || 'nl';
    if (isErr(v)) return locale === 'nl' ? (ERR_NL[v.error] || v.error) : v.error;
    if (v == null) return '';
    if (typeof v === 'boolean') return v ? (locale === 'nl' ? 'WAAR' : 'TRUE') : (locale === 'nl' ? 'ONWAAR' : 'FALSE');
    if (typeof v === 'number') { const s = String(+v.toPrecision(10)); return locale === 'nl' ? s.replace('.', ',') : s; }
    return String(v);
  }
  /* Formule uit een bestand (Engels) omzetten naar wat een Belgische leerling ziet. */
  function toDutch(src) {
    const s = String(src);
    let out = '', i = 0;
    while (i < s.length) {
      const ch = s[i];
      if (ch === '"') { const j = s.indexOf('"', i + 1); const end = j < 0 ? s.length : j + 1; out += s.slice(i, end); i = end; continue; }
      const f = /^([A-Za-z][A-Za-z0-9.]*)\(/.exec(s.slice(i));
      if (f && (i === 0 || !/[A-Za-z0-9_.$]/.test(s[i - 1]))) { const up = f[1].toUpperCase(); out += (EN_NL[up] || up) + '('; i += f[0].length; continue; }
      const nm = /^\d+\.\d+/.exec(s.slice(i));
      if (nm && (i === 0 || !/[A-Za-z$]/.test(s[i - 1]))) { out += nm[0].replace('.', ','); i += nm[0].length; continue; }
      const id = /^[A-Za-z_$][A-Za-z0-9_.$]*/.exec(s.slice(i));
      if (id) { out += id[0]; i += id[0].length; continue; }
      out += ch === ',' ? ';' : ch; i++;
    }
    return '=' + out.replace(/^=/, '');
  }

  const api = { Err, isErr, evaluate, info, shift, display, toDutch, tokenize, parse, NL_EN, EN_NL, ERR_NL, expandRange, parseRef, colToNum, numToCol };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.CvcFormula = api;
})(typeof window !== 'undefined' ? window : globalThis);
