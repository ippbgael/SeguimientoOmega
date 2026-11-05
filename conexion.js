/* =====================================================================
   conexion.js · Seguimiento OMEGA
   Todo lo que habla con Microsoft 365 está en este archivo:
     1. Inicio de sesión (MSAL, cuenta de IPPB)
     2. Datos: una lista de SharePoint («OmegaDatos») donde cada fila es
        una cosa de la plataforma. Columna Title = clave, columna Valor = JSON.
          cfg            plazos y ajustes        sprints     fechas de sprints
          u:<correo>     un usuario              p:<id>      un proceso
          e:<id>:<tipo>  un entregable           s:<id>      sesiones de un proceso
          d:<id>         un motivo de desvío     r:<id>      una relación antecesor→sucesor
          l:<id>         una línea del historial (nunca se borra)
     3. Archivos: carpetas del proceso en la biblioteca del sitio.
   La página (index.html) solo usa window.OMEGA_SP; no conoce Graph.
   ===================================================================== */
(function () {
  'use strict';
  const SCOPES = ['User.Read', 'Sites.ReadWrite.All'];
  const G = 'https://graph.microsoft.com/v1.0';
  const COL = 'Valor';                       // columna de texto con el JSON
  const SUBS = ['01 AS IS', '02 TO BE', '03 Funcional', '04 Validación', 'Actas'];
  const NO_BORRAR = /^(cfg|sprints|p:|e:|s:|l:)/; // claves que la página nunca borra

  let C = null, app = null, acc = null, siteId = '', listId = '', driveId = '';
  const ids = new Map();    // clave → id de la fila en SharePoint
  const saved = new Map();  // clave → último texto visto en SharePoint
  let statusCb = () => {}, watchCb = null;
  let timer = null, pending = false, saving = false, again = false, lastD = null, gen = 0;
  const known = new Set();  // carpetas que ya sabemos que existen
  const sleep = ms => new Promise(r => setTimeout(r, ms));

  /* ---------- 1. Inicio de sesión ---------- */
  const redirectUri = () => location.origin + location.pathname.replace(/index\.html$/, '');
  async function init(cfg) {
    C = { lista: 'OmegaDatos', carpetaRaiz: 'OMEGA', biblioteca: '', ...cfg };
    app = new msal.PublicClientApplication({
      auth: { clientId: C.clientId, authority: 'https://login.microsoftonline.com/' + C.tenantId, redirectUri: redirectUri(), postLogoutRedirectUri: redirectUri() },
      cache: { cacheLocation: 'localStorage' }
    });
    if (app.initialize) await app.initialize();
    const r = await app.handleRedirectPromise();
    acc = (r && r.account) || app.getAllAccounts()[0] || null;
    if (acc) app.setActiveAccount(acc);
    return acc;
  }
  const login = () => app.loginRedirect({ scopes: SCOPES, prompt: 'select_account' });
  const logout = () => app.logoutRedirect({ account: acc });
  const email = () => String(acc && acc.username || '').toLowerCase();

  async function token() {
    try { return (await app.acquireTokenSilent({ scopes: SCOPES, account: acc })).accessToken; }
    catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError || /interaction_required|login_required|consent_required/.test(e.errorCode || '')) {
        await app.acquireTokenRedirect({ scopes: SCOPES, account: acc });
        return new Promise(() => {}); // la página se va a Microsoft y vuelve
      }
      throw e;
    }
  }

  // Llamada a Graph con reintento si Microsoft pide esperar (429/503).
  async function g(method, url, body, tries = 0) {
    const t = await token();
    const r = await fetch(url.startsWith('http') ? url : G + url, {
      method,
      headers: { Authorization: 'Bearer ' + t, ...(body !== undefined ? { 'Content-Type': 'application/json' } : {}) },
      body: body !== undefined ? JSON.stringify(body) : undefined
    });
    if ((r.status === 429 || r.status === 503 || r.status === 504) && tries < 5) {
      await sleep(1000 * (+r.headers.get('Retry-After') || 2 ** tries));
      return g(method, url, body, tries + 1);
    }
    if (!r.ok) {
      let m = ''; try { const j = await r.json(); m = j.error ? j.error.code + ': ' + j.error.message : ''; } catch (e) {}
      const err = new Error('HTTP ' + r.status + ' ' + m); err.status = r.status; throw err;
    }
    return r.status === 204 ? null : r.json();
  }

  /* ---------- 2. Sitio, lista y datos ---------- */
  async function connect() {
    const u = new URL(C.sitio); const path = u.pathname.replace(/\/$/, '');
    const s = await g('GET', '/sites/' + u.hostname + (path ? ':' + path : '') + '?$select=id,webUrl');
    siteId = s.id;
    if (C.biblioteca) {
      const ds = await g('GET', '/sites/' + siteId + '/drives?$select=id,name');
      const d = ds.value.find(x => x.name.toLowerCase() === C.biblioteca.toLowerCase());
      if (!d) { const e = new Error('No existe la biblioteca «' + C.biblioteca + '» en el sitio'); e.status = 404; throw e; }
      driveId = d.id;
    } else driveId = (await g('GET', '/sites/' + siteId + '/drive?$select=id')).id;
    const ls = await g('GET', '/sites/' + siteId + "/lists?$select=id,displayName&$filter=displayName eq '" + C.lista.replace(/'/g, "''") + "'");
    listId = ls.value.length ? ls.value[0].id : '';
  }

  async function fetchAll() {
    const I = new Map(), S = new Map();
    let url = '/sites/' + siteId + '/lists/' + listId + '/items?$expand=fields($select=Title,' + COL + ')&$select=id&$top=999';
    while (url) {
      const r = await g('GET', url);
      r.value.forEach(it => { const k = it.fields && it.fields.Title; if (!k) return; I.set(k, it.id); S.set(k, it.fields[COL] || ''); });
      url = r['@odata.nextLink'] || '';
    }
    return { ids: I, saved: S };
  }
  function commit(ex) { ids.clear(); saved.clear(); ex.ids.forEach((v, k) => ids.set(k, v)); ex.saved.forEach((v, k) => saved.set(k, v)); }

  function implode(m) {
    const D = { users: [], P: [], SP: [], E: {}, ses: {}, desv: [], rel: [], log: [], cfg: {} };
    for (const [k, v] of m) {
      let o; try { o = JSON.parse(v); } catch (e) { console.warn('Fila con JSON dañado, se ignora:', k); continue; }
      const i = k.indexOf(':'), t = i < 0 ? k : k.slice(0, i), r = i < 0 ? '' : k.slice(i + 1);
      if (k === 'cfg') D.cfg = o; else if (k === 'sprints') D.SP = o;
      else if (t === 'u') D.users.push(o); else if (t === 'p') D.P.push(o); else if (t === 'e') D.E[r] = o;
      else if (t === 's') D.ses[r] = o; else if (t === 'd') D.desv.push(o); else if (t === 'r') D.rel.push(o); else if (t === 'l') D.log.push(o);
    }
    const byId = (a, b) => String(a.id).localeCompare(String(b.id));
    D.P.sort((a, b) => a.id - b.id); D.desv.sort(byId); D.rel.sort(byId); D.log.sort((a, b) => (a.ts || 0) - (b.ts || 0));
    return D;
  }
  function explode(D) {
    const m = new Map(); const put = (k, o) => m.set(k, JSON.stringify(o));
    put('cfg', D.cfg); put('sprints', D.SP);
    D.users.forEach(u => put('u:' + u.id, u)); D.P.forEach(p => put('p:' + p.id, p));
    Object.entries(D.E).forEach(([k, v]) => put('e:' + k, v)); Object.entries(D.ses).forEach(([k, v]) => put('s:' + k, v));
    D.desv.forEach(d => put('d:' + d.id, d)); D.rel.forEach(r => put('r:' + r.id, r)); D.log.forEach(l => put('l:' + l.id, l));
    return m;
  }

  async function loadAll() {
    if (!listId) return null;
    const ex = await fetchAll(); commit(ex);
    return ex.saved.has('cfg') ? implode(ex.saved) : null;
  }

  // Envía a SharePoint solo las filas que cambiaron, de a 20 por llamada ($batch).
  async function runOps(ops, prog) {
    const base = '/sites/' + siteId + '/lists/' + listId + '/items';
    let done = 0;
    for (let i = 0; i < ops.length; i += 20) {
      let chunk = ops.slice(i, i + 20);
      for (let t = 0; chunk.length && t < 6; t++) {
        const requests = chunk.map((o, j) => ({
          id: String(j),
          method: o.op === 'add' ? 'POST' : o.op === 'set' ? 'PATCH' : 'DELETE',
          url: o.op === 'add' ? base : o.op === 'set' ? base + '/' + ids.get(o.k) + '/fields' : base + '/' + ids.get(o.k),
          headers: o.op === 'del' ? undefined : { 'Content-Type': 'application/json' },
          body: o.op === 'add' ? { fields: { Title: o.k, [COL]: o.v } } : o.op === 'set' ? { [COL]: o.v } : undefined
        }));
        const r = await g('POST', '/$batch', { requests });
        const retry = []; let wait = 0;
        for (const res of r.responses) {
          const o = chunk[+res.id];
          if (res.status >= 200 && res.status < 300) {
            if (o.op === 'add') { ids.set(o.k, res.body.id); saved.set(o.k, o.v); }
            else if (o.op === 'set') saved.set(o.k, o.v);
            else { ids.delete(o.k); saved.delete(o.k); }
            done++;
          } else if (res.status === 429 || res.status >= 500) {
            retry.push(o); wait = Math.max(wait, +((res.headers || {})['Retry-After']) || 2);
          } else if (o.op === 'del' && res.status === 404) { ids.delete(o.k); saved.delete(o.k); done++; }
          else {
            const err = new Error('HTTP ' + res.status + ' ' + (res.body && res.body.error ? res.body.error.code + ': ' + res.body.error.message : '') + ' (' + o.k + ')');
            err.status = res.status; throw err;
          }
        }
        chunk = retry; if (chunk.length) await sleep(wait * 1000);
      }
      if (chunk.length) throw new Error('SharePoint está ocupado; vuelve a intentarlo en un momento.');
      if (prog) prog(done, ops.length);
    }
  }

  function schedule(D, delay = 600) { lastD = D; gen++; pending = true; clearTimeout(timer); timer = setTimeout(flush, delay); }
  async function flush() {
    if (saving) { again = true; return; }
    const D = lastD; if (!D) return;
    saving = true; again = false; pending = false; statusCb({ k: 'saving', msg: '' });
    try {
      const cur = explode(D); const ops = [];
      for (const [k, v] of cur) { if (!ids.has(k)) ops.push({ k, v, op: 'add' }); else if (saved.get(k) !== v) ops.push({ k, v, op: 'set' }); }
      for (const k of ids.keys()) if (!cur.has(k) && !NO_BORRAR.test(k)) ops.push({ k, op: 'del' });
      await runOps(ops);
      statusCb({ k: 'ok', msg: '' });
    } catch (e) { console.error(e); statusCb({ k: 'error', msg: String(e.message || e) }); }
    finally { saving = false; if (again) flush(); }
  }

  // Trae los cambios de los demás. No pisa nada que esté pendiente de guardar.
  async function refresh() {
    if (!listId || pending || saving) return false;
    const g0 = gen; const ex = await fetchAll();
    if (gen !== g0 || pending || saving || !ex.saved.has('cfg')) return false;
    let same = ex.saved.size === saved.size;
    if (same) for (const [k, v] of ex.saved) if (saved.get(k) !== v) { same = false; break; }
    commit(ex);
    if (!same && watchCb) watchCb(implode(ex.saved));
    return !same;
  }
  function watch(cb) {
    watchCb = cb;
    setInterval(() => refresh().catch(e => console.warn('No se pudo actualizar', e)), 120000);
    document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') refresh().catch(() => {}); });
  }

  // Primera carga: crea la lista (si no existe), carga todo y deja «cfg» al final como marca de «carga completa».
  async function setup(D, folders, prog) {
    if (!listId) {
      prog('Creando la lista «' + C.lista + '»…');
      const l = await g('POST', '/sites/' + siteId + '/lists', {
        displayName: C.lista,
        columns: [{ name: COL, text: { allowMultipleLines: true, textType: 'plain', linesForEditing: 6 } }],
        list: { template: 'genericList' }
      });
      listId = l.id; await sleep(2000);
    } else {
      const ex = await fetchAll();
      if (ex.saved.has('cfg')) throw new Error('La lista ya tiene datos de OMEGA; no se volvió a cargar nada.');
      commit(ex); // reintento de una carga a medias: no duplica lo que ya está
    }
    const cur = explode(D);
    const ops = [...cur].filter(([k]) => k !== 'cfg' && !ids.has(k)).map(([k, v]) => ({ k, v, op: 'add' }));
    await runOps(ops, (n, t) => prog('Cargando datos en SharePoint… ' + n + ' de ' + t));
    if (folders) {
      let i = 0;
      for (const p of D.P) { i++; prog('Creando carpetas… ' + i + ' de ' + D.P.length + ' · ' + p.codigo); await ensureProcFolders(p, SUBS); }
    }
    await runOps([{ k: 'cfg', v: cur.get('cfg'), op: 'add' }]);
  }

  /* ---------- 3. Carpetas y archivos ---------- */
  const clean = s => String(s).replace(/["*:<>?\/\\|#%~&{}]/g, '-').replace(/\s+/g, ' ').replace(/[. ]+$/, '').trim().slice(0, 90);
  const procPath = p => [C.carpetaRaiz, 'Sprint ' + p.sprint, clean(p.codigo + ' - ' + p.nombre)].join('/');
  const enc = path => path.split('/').map(encodeURIComponent).join('/');
  async function item(path) {
    try { return await g('GET', '/drives/' + driveId + '/root:/' + enc(path) + '?$select=id,webUrl,folder'); }
    catch (e) { if (e.status === 404) return null; throw e; }
  }
  async function ensurePath(path) {
    let parent = '', fresh = false;
    for (const s of path.split('/')) {
      const p = parent ? parent + '/' + s : s;
      if (!known.has(p)) {
        let it = fresh ? null : await item(p);
        if (!it) {
          try {
            await g('POST', parent ? '/drives/' + driveId + '/root:/' + enc(parent) + ':/children' : '/drives/' + driveId + '/root/children',
              { name: s, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' });
            fresh = true;
          } catch (e) { if (e.status !== 409) throw e; }
        }
        known.add(p);
      }
      parent = p;
    }
  }
  async function ensureProcFolders(p, subs = SUBS) { const base = procPath(p); for (const s of subs) await ensurePath(base + '/' + s); }

  async function folderInfo(p, subs = SUBS) {
    const it = await item(procPath(p));
    if (!it) return { k: 'none' };
    const r = await g('GET', '/drives/' + driveId + '/items/' + it.id + '/children?$select=id,name,folder&$top=200');
    const out = {}; subs.forEach(s => { out[s] = null; });
    for (const f of r.value.filter(x => x.folder && subs.includes(x.name.normalize('NFC')))) {
      const c = await g('GET', '/drives/' + driveId + '/items/' + f.id + '/children?$select=name,webUrl,file,lastModifiedDateTime,lastModifiedBy&$top=200');
      out[f.name.normalize('NFC')] = c.value.filter(x => x.file).map(x => ({
        name: x.name, web: x.webUrl, mod: x.lastModifiedDateTime || '',
        by: x.lastModifiedBy && x.lastModifiedBy.user ? x.lastModifiedBy.user.displayName : ''
      }));
    }
    return { k: 'ok', web: it.webUrl, subs: out };
  }

  // Sube un archivo a la subcarpeta del proceso. Hasta 4 MB de una vez; más grande, por partes.
  async function upload(p, sub, name, file, onPct) {
    const folder = procPath(p) + '/' + sub; await ensurePath(folder);
    const path = '/drives/' + driveId + '/root:/' + enc(folder + '/' + name.replace(/["*:<>?\/\\|#%]/g, '-'));
    if (file.size <= 4 * 1024 * 1024) {
      const t = await token();
      const r = await fetch(G + path + ':/content?@microsoft.graph.conflictBehavior=rename', { method: 'PUT', headers: { Authorization: 'Bearer ' + t, 'Content-Type': file.type || 'application/octet-stream' }, body: file });
      if (!r.ok) { const e = new Error('HTTP ' + r.status + ' al subir'); e.status = r.status; throw e; }
      const it = await r.json(); if (onPct) onPct(100); return { web: it.webUrl, name: it.name };
    }
    const s = await g('POST', path + ':/createUploadSession', { item: { '@microsoft.graph.conflictBehavior': 'rename' } });
    const CH = 320 * 1024 * 16; let pos = 0, it = null;
    while (pos < file.size) {
      const end = Math.min(pos + CH, file.size);
      const r = await fetch(s.uploadUrl, { method: 'PUT', headers: { 'Content-Range': 'bytes ' + pos + '-' + (end - 1) + '/' + file.size }, body: file.slice(pos, end) });
      if (!r.ok) { const e = new Error('HTTP ' + r.status + ' al subir'); e.status = r.status; throw e; }
      if (r.status === 200 || r.status === 201) it = await r.json();
      pos = end; if (onPct) onPct(Math.round(pos / file.size * 100));
    }
    return { web: it.webUrl, name: it.name };
  }

  window.OMEGA_SP = {
    init, login, logout, email, ready: () => !!app, cfg: () => C,
    connect, loadAll, schedule, refresh, watch, setup,
    onStatus: cb => { statusCb = cb; },
    procPath, folderInfo, ensureProcFolders, upload
  };
})();
