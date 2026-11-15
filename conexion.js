/* =====================================================================
   conexion.js · Seguimiento OMEGA · backend Supabase
   Todo lo que habla con el servidor está en este archivo:
     1. Inicio de sesión: correo + contraseña (Supabase Auth).
     2. Datos: tabla «registros». Cada fila es una cosa de la plataforma:
          cfg            plazos y ajustes        sprints     fechas de sprints
          u:<correo>     un usuario y su rol     p:<id>      un proceso
          e:<id>:<tipo>  un entregable           s:<id>      sesiones de un proceso
          d:<id>         un motivo de desvío     r:<id>      una relación antecesor→sucesor
          l:<id>         una línea del historial (nunca se borra)
        Quién puede leer o cambiar cada fila lo decide la base (supabase.sql),
        no esta página.
     3. Documentos: no se tocan. La plataforma solo guarda el enlace.
   La página (index.html) solo usa window.OMEGA_SP.
   ===================================================================== */
(function () {
  'use strict';
  const T = 'registros';
  const BORRABLES = /^(u:|d:|r:)/;      // las demás claves nunca se borran
  let C = null, sb = null, user = null, perfil = null, recovery = false;
  const saved = new Map();              // clave → último texto visto en el servidor
  let statusCb = () => {}, watchCb = null;
  let timer = null, pending = false, saving = false, again = false, lastD = null, gen = 0;
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const fail = (e, code) => { const err = new Error(e && e.message || String(e)); err.code = code || (e && e.code); err.status = e && e.status; return err; };

  /* ---------- 1. Inicio de sesión ---------- */
  function configOk(c) { return !!(c && c.supabaseUrl && c.supabaseKey && /^https:\/\/.+/.test(c.supabaseUrl)); }
  function libOk() { return !!(window.supabase && window.supabase.createClient); }
  async function init(cfg) {
    C = cfg;
    sb = window.supabase.createClient(C.supabaseUrl, C.supabaseKey, { auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true } });
    sb.auth.onAuthStateChange(ev => { if (ev === 'PASSWORD_RECOVERY') recovery = true; });
    const { data, error } = await sb.auth.getSession();
    if (error) throw fail(error);
    user = data.session ? data.session.user : null;
    if (location.hash.includes('type=recovery')) recovery = true;
    return user;
  }
  async function login(email, password) {
    const { data, error } = await sb.auth.signInWithPassword({ email: String(email).trim().toLowerCase(), password });
    if (error) throw fail(error, 'login');
    user = data.user; return user;
  }
  async function logout() { try { await sb.auth.signOut(); } finally { location.hash = ''; location.reload(); } }
  async function resetPassword(email) {
    const { error } = await sb.auth.resetPasswordForEmail(String(email).trim().toLowerCase(), { redirectTo: location.origin + location.pathname });
    if (error) throw fail(error);
  }
  async function changePassword(pw) {
    const { error } = await sb.auth.updateUser({ password: pw });
    if (error) throw fail(error); recovery = false;
  }
  const email = () => String(user && user.email || '').toLowerCase();

  /* ---------- 2. Datos ---------- */
  async function connect() {
    const { data, error } = await sb.rpc('mi_perfil');
    if (error) throw fail(error);
    perfil = data || null;
    if (!perfil) throw fail(new Error('Tu correo no está en la lista de usuarios'), 'noaccess');
  }
  const esGestor = () => !!perfil && (perfil.rol === 'Líder de Procesos' || !!perfil.admin);

  async function fetchAll() {
    const m = new Map(); const page = 1000;
    for (let from = 0; ; from += page) {
      const { data, error } = await sb.from(T).select('clave,valor').order('clave').range(from, from + page - 1);
      if (error) throw fail(error);
      data.forEach(r => m.set(r.clave, JSON.stringify(r.valor)));
      if (data.length < page) break;
    }
    return m;
  }
  function commit(m) { saved.clear(); m.forEach((v, k) => saved.set(k, v)); }

  function implode(m) {
    const D = { users: [], P: [], SP: [], E: {}, ses: {}, desv: [], rel: [], log: [], cfg: {} };
    for (const [k, v] of m) {
      let o; try { o = JSON.parse(v); } catch (e) { continue; }
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
    const m = await fetchAll(); commit(m);
    return m.has('cfg') ? implode(m) : null;
  }

  // Envía solo lo que cambió: filas nuevas en bloques, cambios fila por fila, y borrados.
  async function runOps(ops, prog) {
    const adds = ops.filter(o => o.op === 'add'), sets = ops.filter(o => o.op === 'set'), dels = ops.filter(o => o.op === 'del');
    let done = 0; const tick = () => { if (prog) prog(done, ops.length); };
    // Primero procesos y usuarios: la base los necesita para saber el equipo de lo demás.
    const first = adds.filter(o => /^(u:|p:|sprints)/.test(o.k)), rest = adds.filter(o => !/^(u:|p:|sprints)/.test(o.k));
    for (const group of [first, rest]) for (let i = 0; i < group.length; i += 200) {
      const chunk = group.slice(i, i + 200);
      const { error } = await sb.from(T).insert(chunk.map(o => ({ clave: o.k, valor: JSON.parse(o.v) })));
      if (error) throw fail(error);
      chunk.forEach(o => saved.set(o.k, o.v)); done += chunk.length; tick();
    }
    for (let i = 0; i < sets.length; i += 6) {
      const chunk = sets.slice(i, i + 6);
      const res = await Promise.all(chunk.map(o => sb.from(T).update({ valor: JSON.parse(o.v) }).eq('clave', o.k).select('clave')));
      res.forEach((r, j) => {
        if (r.error) throw fail(r.error);
        if (!r.data || !r.data.length) throw fail(new Error('No tienes permiso para cambiar «' + chunk[j].k + '»'), '42501');
        saved.set(chunk[j].k, chunk[j].v);
      });
      done += chunk.length; tick();
    }
    if (dels.length) {
      const { error } = await sb.from(T).delete().in('clave', dels.map(o => o.k));
      if (error) throw fail(error);
      dels.forEach(o => saved.delete(o.k)); done += dels.length; tick();
    }
  }

  function schedule(D, delay = 600) { lastD = D; gen++; pending = true; clearTimeout(timer); timer = setTimeout(flush, delay); }
  async function flush() {
    if (saving) { again = true; return; }
    const D = lastD; if (!D) return;
    saving = true; again = false; pending = false; statusCb({ k: 'saving', msg: '' });
    try {
      const cur = explode(D); const ops = [];
      for (const [k, v] of cur) { if (!saved.has(k)) ops.push({ k, v, op: 'add' }); else if (saved.get(k) !== v) ops.push({ k, v, op: 'set' }); }
      for (const k of saved.keys()) if (!cur.has(k) && BORRABLES.test(k)) ops.push({ k, op: 'del' });
      await runOps(ops);
      statusCb({ k: 'ok', msg: '' });
    } catch (e) { console.error(e); statusCb({ k: 'error', msg: String(e.message || e) }); }
    finally { saving = false; if (again) flush(); }
  }

  // Trae los cambios de los demás sin pisar lo que está pendiente de guardar.
  async function refresh() {
    if (pending || saving) return false;
    const g0 = gen; const m = await fetchAll();
    if (gen !== g0 || pending || saving || !m.has('cfg')) return false;
    let same = m.size === saved.size;
    if (same) for (const [k, v] of m) if (saved.get(k) !== v) { same = false; break; }
    commit(m);
    if (!same && watchCb) watchCb(implode(m));
    return !same;
  }
  function watch(cb) {
    watchCb = cb;
    setInterval(() => refresh().catch(e => console.warn('No se pudo actualizar', e)), 60000);
    document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') refresh().catch(() => {}); });
  }

  // Primera carga: sube los 55 procesos y deja «cfg» al final como marca de «carga completa».
  async function setup(D, _folders, prog) {
    if (!esGestor()) throw fail(new Error('Solo un administrador puede configurar la plataforma'), '42501');
    const m = await fetchAll(); commit(m);
    if (m.has('cfg')) throw fail(new Error('La plataforma ya tiene datos; no se volvió a cargar nada.'));
    const cur = explode(D);
    const ops = [...cur].filter(([k]) => k !== 'cfg' && !saved.has(k)).map(([k, v]) => ({ k, v, op: 'add' }));
    await runOps(ops, (n, t) => prog('Cargando datos… ' + n + ' de ' + t));
    await runOps([{ k: 'cfg', v: cur.get('cfg'), op: 'add' }]);
  }

  window.OMEGA_SP = {
    kind: 'password', nombre: 'Supabase',
    configOk, libOk, init, login, logout, resetPassword, changePassword, email,
    ready: () => !!sb, cfg: () => C, recovery: () => recovery, esGestor,
    connect, loadAll, schedule, refresh, watch, setup,
    onStatus: cb => { statusCb = cb; }
  };
})();
