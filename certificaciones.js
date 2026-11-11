/*
  SGI Condor Wireline — pestaña Certificaciones.
  Lee en vivo la carpeta de SharePoint Operaciones › Pozos › Certificaciones (Microsoft Graph, con la cuenta
  de CWL de cada usuario), calcula el estado de cada certificado (vencido / por vencer / vigente / sin fecha),
  guarda la clasificación en Firestore (colección "cert") y permite reemplazar un certificado vencido:
  sube el nuevo a la misma carpeta de SharePoint y mueve el anterior a "certificaciones vencidas".
  Configuración: ms-config.js. Usa funciones de index.html ($, esc, S, enqueue, toast, …) en el momento de llamarlas.
*/
(function () {
  'use strict';
  const CFG = Object.assign({ tenant: 'organizations', diasAviso: 30, carpetaVencidas: 'certificaciones vencidas' }, window.SGI_MS_CONFIG || {});
  const SCOPES = ['Files.ReadWrite.All', 'Sites.Read.All'];
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const LSK = 'cwl4.certIdx';
  const MSAL_URLS = ['https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js',
                     'https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.38.3/lib/msal-browser.min.js'];
  const TIPOS = ['Certificado', 'Prueba hidráulica', 'Informe de inspección', 'Calibración', 'Ensayo no destructivo', 'Seguro / póliza', 'Habilitación / registro', 'Manual / ficha técnica', 'Otro'];
  const LINKS = { equipos: 'Equipos', herramientas: 'Herramientas', flota: 'Flota', personal: 'Legajos' };
  const HIST_RE = /vencid|obsolet|histori|anterior|viej|reemplaz|\bbaja\b/i;
  const configured = () => CFG.clientId && !/COMPLETAR/i.test(CFG.clientId);

  const C = {
    msal: null, account: null, connecting: false,
    idx: lsGet(),            // {at, root:{driveId,id,webUrl}, folders:[...], files:[...]}
    scan: null,              // {folders, files} mientras escanea
    f: { q: '', cat: '', st: '', hist: false, sort: 'venc' },
    sel: new Set(), err: '',
  };
  function lsGet() { try { const v = localStorage.getItem(LSK); return v ? JSON.parse(v) : null; } catch (e) { return null; } }
  function lsSet(v) { try { localStorage.setItem(LSK, JSON.stringify(v)); } catch (e) { console.warn('[SGI cert] índice demasiado grande para guardarlo en el equipo'); } }
  const q$ = s => document.querySelector(s);
  const e$ = s => (typeof esc === 'function' ? esc(s) : String(s ?? ''));
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const clsOf = id => ((S.col && S.col.cert) || {})[id] || {};

  /* ================= Microsoft: login y Graph ================= */
  function loadScript(src) { return new Promise((res, rej) => { const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = rej; document.head.appendChild(s); }); }
  async function ensureMsal() {
    if (C.msal) return C.msal;
    if (!window.msal) { let ok = false; for (const u of MSAL_URLS) { try { await loadScript(u); ok = !!window.msal; if (ok) break; } catch (e) { } } if (!ok) throw new Error('No se pudo cargar el inicio de sesión de Microsoft (¿sin internet?)'); }
    C.msal = new msal.PublicClientApplication({
      auth: { clientId: CFG.clientId, authority: 'https://login.microsoftonline.com/' + CFG.tenant, redirectUri: new URL('ms-redirect.html', location.href).href },
      cache: { cacheLocation: 'localStorage' },
    });
    if (C.msal.initialize) await C.msal.initialize();
    const acc = C.msal.getAllAccounts(); if (acc.length) C.account = acc[0];
    return C.msal;
  }
  async function token(interactive) {
    const m = await ensureMsal();
    if (C.account) {
      try { return (await m.acquireTokenSilent({ scopes: SCOPES, account: C.account })).accessToken; }
      catch (e) { if (!interactive) throw e; }
    } else if (!interactive) throw new Error('login');
    const r = C.account ? await m.acquireTokenPopup({ scopes: SCOPES, account: C.account }) : await m.loginPopup({ scopes: SCOPES, prompt: 'select_account' });
    C.account = r.account; try { localStorage.setItem('cwl4.msOn', '1'); } catch (e) { }
    return r.accessToken || (await m.acquireTokenSilent({ scopes: SCOPES, account: C.account })).accessToken;
  }
  async function g(path, opt = {}, tries = 0) {
    const t = await token(false).catch(() => token(true));
    const url = /^https:/.test(path) ? path : GRAPH + path;
    const h = Object.assign({ Authorization: 'Bearer ' + t }, opt.json ? { 'Content-Type': 'application/json' } : {}, opt.headers || {});
    const r = await fetch(url, { method: opt.method || 'GET', headers: h, body: opt.json ? JSON.stringify(opt.json) : opt.body });
    if ((r.status === 429 || r.status === 503 || r.status === 504) && tries < 5) { await sleep((+r.headers.get('Retry-After') || 2 + tries * 2) * 1000); return g(path, opt, tries + 1); }
    if (!r.ok) { let m = r.status + ''; try { const j = await r.json(); m = (j.error && (j.error.message || j.error.code)) || m; } catch (e) { } const er = new Error(m); er.status = r.status; throw er; }
    return r.status === 204 ? null : r.json();
  }
  const SEL = '$select=id,name,folder,file,size,webUrl,lastModifiedDateTime,parentReference&$top=999';
  async function children(driveId, id) {
    let out = [], next = `/drives/${driveId}/items/${id}/children?${SEL}`;
    while (next) { const j = await g(next); out = out.concat(j.value || []); next = j['@odata.nextLink'] || null; }
    return out;
  }
  async function resolveRoot() {
    if (CFG.driveId && CFG.folderId) {
      try { const it = await g(`/drives/${CFG.driveId}/items/${CFG.folderId}?$select=id,name,webUrl`); return { driveId: CFG.driveId, id: it.id, webUrl: it.webUrl }; }
      catch (e) { if (e.status === 401 || e.status === 403) throw e; }
    }
    const site = await g(`/sites/${CFG.site}?$select=id`);
    const drives = (await g(`/sites/${site.id}/drives?$select=id,name,webUrl`)).value || [];
    const d = drives.find(x => x.name === CFG.library || decodeURIComponent(x.webUrl || '').endsWith('/' + CFG.library));
    if (!d) throw new Error(`No encontré la biblioteca "${CFG.library}" en ${CFG.site}`);
    const it = await g(`/drives/${d.id}/root:/${encodeURIComponent(CFG.folder)}?$select=id,name,webUrl`);
    return { driveId: d.id, id: it.id, webUrl: it.webUrl };
  }
  // Recorre toda la carpeta (6 pedidos en paralelo) y arma el índice de carpetas y archivos
  async function scanAll() {
    if (C.scan) return;
    C.err = ''; C.scan = { folders: 0, files: 0 }; render();
    try {
      const root = await resolveRoot();
      const folders = [], files = [];
      const queue = [{ id: root.id, path: [] }];
      let active = 0;
      await new Promise((res, rej) => {
        const pump = () => {
          if (!queue.length && !active) return res();
          while (active < 6 && queue.length) {
            const f = queue.shift(); active++;
            children(root.driveId, f.id).then(list => {
              for (const it of list) {
                if (it.folder) { const p = f.path.concat(it.name); folders.push({ id: it.id, name: it.name, path: p, parentId: f.id, webUrl: it.webUrl, n: it.folder.childCount }); queue.push({ id: it.id, path: p }); C.scan.folders++; }
                else if (it.file) { files.push({ id: it.id, name: it.name, path: f.path, parentId: f.id, webUrl: it.webUrl, size: it.size, mod: it.lastModifiedDateTime, mime: it.file.mimeType }); C.scan.files++; }
              }
              active--; progress(); pump();
            }).catch(rej);
          }
        };
        pump();
      });
      C.idx = { at: new Date().toISOString(), root, folders, files };
      lsSet(C.idx);
      toast(`SharePoint leído: ${files.length} archivos en ${folders.length} carpetas`);
    } catch (e) {
      console.error('[SGI cert]', e);
      C.err = e.status === 403 || e.status === 401 ? 'Tu cuenta de Microsoft no tiene permiso para leer esa carpeta, o falta el consentimiento del administrador para la app.' : (e.message || String(e));
    }
    C.scan = null; render();
  }
  function progress() { const el = q$('#ct-prog'); if (el && C.scan) el.textContent = `Leyendo SharePoint… ${C.scan.folders} carpetas · ${C.scan.files} archivos`; }
  async function connect(interactive) {
    if (!configured()) return;
    C.connecting = true; C.err = ''; render();
    try { await token(interactive); C.connecting = false; await scanAll(); }
    catch (e) { C.connecting = false; if (interactive) C.err = (e && (e.errorMessage || e.message)) || String(e); console.warn('[SGI cert] login', e); render(); }
  }

  /* ================= estado de cada certificado ================= */
  const MES = { ene: 1, feb: 2, mar: 3, abr: 4, may: 5, jun: 6, jul: 7, ago: 8, sep: 9, set: 9, oct: 10, nov: 11, dic: 12 };
  const KW = '(?:vence|vencen|vencimiento|venc|vto|vtos|v[aá]lid[oa]\\s+hasta|hasta)\\.?\\s*[:\\-]?\\s*(?:el\\s+)?';
  function ymd(y, m, d) { y = +y; if (y < 100) y += 2000; m = +m; d = +d; if (!(y > 2000 && y < 2100 && m >= 1 && m <= 12 && d >= 1 && d <= 31)) return ''; const t = new Date(y, m - 1, d); if (t.getMonth() !== m - 1) return ''; return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`; }
  const lastDay = (y, m) => { y = +y; if (y < 100) y += 2000; return ymd(y, m, new Date(y, +m, 0).getDate()); };
  function parseVenc(s) {
    s = String(s || '').replace(/_/g, ' ');
    let m;
    if ((m = s.match(new RegExp(KW + '(\\d{4})[-/.](\\d{1,2})[-/.](\\d{1,2})', 'i')))) return ymd(m[1], m[2], m[3]);
    if ((m = s.match(new RegExp(KW + '(\\d{1,2})[-/.](\\d{1,2})[-/.](\\d{2,4})(?!\\d)', 'i')))) return ymd(m[3], m[2], m[1]);
    if ((m = s.match(new RegExp(KW + '(?:(\\d{1,2})\\s*(?:de\\s+)?)?(ene|feb|mar|abr|may|jun|jul|ago|sep|set|oct|nov|dic)[a-z]*\\.?\\s*(?:de\\s+|del\\s+)?(\\d{2,4})', 'i')))) { const mm = MES[m[2].toLowerCase()]; return m[1] ? ymd(m[3], mm, m[1]) : lastDay(m[3], mm); }
    if ((m = s.match(new RegExp(KW + '(\\d{1,2})[-/.](\\d{4})(?!\\d)', 'i')))) return lastDay(m[2], m[1]);
    return '';
  }
  const cleanCat = n => { const t = String(n || '').replace(/^(cer?tificaci[oó]n(es)?|certificados?)\s+(de\s+)?/i, '').trim(); return t ? t[0].toUpperCase() + t.slice(1) : 'Sin clasificar'; };
  function guessTipo(n) {
    const t = n.toLowerCase();
    return /hidr[aá]ul|hidrost|\bph\b/.test(t) ? 'Prueba hidráulica' : /inspecci/.test(t) ? 'Informe de inspección' : /calibr/.test(t) ? 'Calibración' : /tintas|ultrason|\bend\b|\bndt\b|part[ií]culas/.test(t) ? 'Ensayo no destructivo' :
      /seguro|p[oó]liza/.test(t) ? 'Seguro / póliza' : /manual|ficha/.test(t) ? 'Manual / ficha técnica' : /certific/.test(t) ? 'Certificado' : '';
  }
  // devuelve todo lo que la vista necesita de un archivo
  function info(it) {
    const cl = clsOf(it.id);
    let venc = cl.venc || '', origen = venc ? 'manual' : '';
    if (!venc) { venc = parseVenc(it.name); if (venc) origen = 'nombre'; }
    if (!venc) for (let i = it.path.length - 1; i >= 0; i--) { venc = parseVenc(it.path[i]); if (venc) { origen = 'carpeta'; break; } }
    const hist = !!(cl.reemplazadoPor || cl.estado === 'Reemplazado' || it.path.some(p => HIST_RE.test(p)));
    const d = venc && typeof dayDiff === 'function' ? dayDiff(venc) : null;
    const st = hist ? 'hist' : cl.sinVenc ? 'nc' : d == null ? 'sf' : d < 0 ? 'venc' : d <= CFG.diasAviso ? 'prox' : 'ok';
    return { it, cl, venc, origen, d, st, hist,
      cat: cl.cat || cleanCat(it.path[0]), elem: cl.elem || (it.path.length > 1 ? it.path[1] : ''), tipo: cl.tipo || guessTipo(it.name),
      linkTxt: cl.link && cl.link.rk ? linkLabel(cl.link) : '' };
  }
  const ST = { venc: ['Vencido', 'p-bad'], prox: ['Por vencer', 'p-pend'], ok: ['Vigente', 'p-closed'], sf: ['Sin fecha', 'p-ne'], nc: ['No vence', 'p-ne'], hist: ['Histórico', 'p-ne'] };
  const ORD = { venc: 0, prox: 1, sf: 2, ok: 3, nc: 4, hist: 5 };
  function linkLabel(l) { const R = REG[l.rk]; if (!R) return ''; const d = ((S.col && S.col[R.c]) || {})[l.id]; return d ? `${LINKS[l.rk] || R.t}: ${d.codigo || d.nro || d.nombre || ''}` : `${LINKS[l.rk]}: (eliminado)`; }
  function all() { return C.idx ? C.idx.files.map(info) : []; }

  /* ================= vista ================= */
  function render() { if (S.view === 'cert') window.renderCert(); }
  window.renderCert = function () {
    const v = q$('#v-cert'); if (!v) return;
    if (window.certAuto) window.certAuto();
    const head = `<div class="listhead"><div><h2>Certificaciones</h2><div class="muted small">SharePoint · Operaciones › Pozos › Certificaciones${C.idx ? ` · leído ${e$(fwhen(C.idx.at))}` : ''}</div></div>
      <div class="exp">${C.idx && C.idx.root ? `<a class="btn-g" href="${e$(C.idx.root.webUrl)}" target="_blank" rel="noopener">Abrir en SharePoint ↗</a>` : ''}${configured() ? `<button class="btn-g" id="ct-sync" ${C.scan || C.connecting ? 'disabled' : ''}>${C.account ? '↻ Actualizar desde SharePoint' : 'Conectar con Microsoft'}</button>` : ''}${C.idx ? '<button class="btn-g" id="ct-csv">Exportar CSV</button>' : ''}</div></div>`;
    let top = '';
    if (!configured()) top = setupHTML();
    else if (C.scan) top = `<div class="banner"><span id="ct-prog">Leyendo SharePoint… ${C.scan.folders} carpetas · ${C.scan.files} archivos</span></div>`;
    else if (C.connecting) top = `<div class="banner">Conectando con Microsoft… si no aparece la ventana, habilitá las ventanas emergentes para este sitio.</div>`;
    else if (C.err) top = `<div class="banner bad"><b>No se pudo leer SharePoint.</b> ${e$(C.err)} <button class="btn-g" id="ct-retry">Reintentar</button></div>`;
    else if (!C.account) top = `<div class="banner warn">Para ver los estados en vivo, reemplazar archivos y moverlos, conectate con tu cuenta de CWL. <button class="btn-p" id="ct-login">Conectar con Microsoft</button></div>`;
    if (!C.idx) { v.innerHTML = head + top + (configured() && !C.scan ? `<div class="empty"><b>Todavía no se leyó la carpeta de certificaciones</b>Tocá “Conectar con Microsoft”: se lee toda la carpeta de SharePoint y se arma el semáforo de vencimientos.</div>` : ''); bindHead(); return; }

    const rows = all(); const F = C.f;
    const live = rows.filter(r => !r.hist);
    const n = k => live.filter(r => r.st === k).length;
    const cats = {}; for (const r of live) { const c = cats[r.cat] || (cats[r.cat] = { n: 0, bad: 0, prox: 0 }); c.n++; if (r.st === 'venc') c.bad++; if (r.st === 'prox') c.prox++; }
    const kp = [['venc', 'Vencidos', n('venc'), 'requieren reemplazo'], ['prox', `Vencen en ${CFG.diasAviso} días`, n('prox'), 'renovar'], ['ok', 'Vigentes', n('ok'), 'con fecha válida'], ['sf', 'Sin fecha', n('sf'), 'clasificar vencimiento']];
    v.innerHTML = head + top + `
      <div class="kpis ct-kpis">${kp.map(([k, l, val, s]) => `<button class="kpi ct-k${F.st === k ? ' on' : ''}" data-st="${k}"><div class="l">${l}</div><div class="v" style="color:${k === 'venc' && val ? 'var(--bad)' : k === 'prox' && val ? 'var(--amber)' : 'inherit'}">${val}</div><div class="s">${s}</div></button>`).join('')}</div>
      <div class="colbar ct-bar"><input class="q" id="ct-q" type="search" placeholder="Buscar archivo, equipo, unidad, carpeta…" value="${e$(F.q)}" aria-label="Buscar">
        <select id="ct-st" aria-label="Estado"><option value="">Todos los estados</option>${Object.entries(ST).filter(([k]) => k !== 'hist').map(([k, [l]]) => `<option value="${k}" ${F.st === k ? 'selected' : ''}>${l}</option>`).join('')}</select>
        <select id="ct-sort" aria-label="Orden"><option value="venc" ${F.sort === 'venc' ? 'selected' : ''}>Ordenar por urgencia</option><option value="cat" ${F.sort === 'cat' ? 'selected' : ''}>Ordenar por categoría</option><option value="name" ${F.sort === 'name' ? 'selected' : ''}>Ordenar por nombre</option></select>
        <label class="tog small"><input type="checkbox" id="ct-hist" ${F.hist ? 'checked' : ''}> Ver históricos</label></div>
      <div class="chips" id="ct-cats"><button data-cat="" aria-pressed="${!F.cat}">Todas · ${live.length}</button>${Object.entries(cats).sort((a, b) => b[1].bad - a[1].bad || a[0].localeCompare(b[0], 'es')).map(([c, o]) => `<button data-cat="${e$(c)}" aria-pressed="${F.cat === c}">${e$(c)} · ${o.n}${o.bad ? ` <i class="ct-dot bad" title="${o.bad} vencidos"></i>` : o.prox ? ` <i class="ct-dot warn" title="${o.prox} por vencer"></i>` : ''}</button>`).join('')}</div>
      <div class="ct-selbar" id="ct-selbar" hidden><span id="ct-seln"></span><button class="btn-p" id="ct-bulk">Clasificar selección</button><button class="btn-g" id="ct-selx">Quitar selección</button></div>
      <div class="tbl-wrap"><table class="ct-t"><thead><tr><th style="width:28px"><input type="checkbox" id="ct-all" aria-label="Seleccionar todos"></th><th>Documento</th><th>Categoría</th><th>Elemento</th><th>Tipo</th><th>Vencimiento</th><th>Estado</th><th></th></tr></thead><tbody id="ct-body"></tbody></table></div>
      <p class="hint" style="margin-top:8px">La fecha de vencimiento se toma de lo que cargues al clasificar; si no hay, del nombre del archivo o de la carpeta (por ejemplo “vence 13-01-27”). Los archivos dentro de carpetas “vencidas” o reemplazados se muestran como histórico.</p>`;
    bindHead(); bindList(); body();
  };
  function filtered() {
    const F = C.f, q = F.q.trim().toLowerCase();
    let rows = all().filter(r => (F.hist || !r.hist || F.st === 'hist') && (!F.cat || r.cat === F.cat) && (!F.st || r.st === F.st) &&
      (!q || [r.it.name, r.it.path.join(' '), r.cat, r.elem, r.tipo, r.linkTxt, r.cl.notas].join(' ').toLowerCase().includes(q)));
    const byName = (a, b) => a.it.name.localeCompare(b.it.name, 'es', { numeric: true });
    if (F.sort === 'cat') rows.sort((a, b) => a.cat.localeCompare(b.cat, 'es') || String(a.elem).localeCompare(String(b.elem), 'es', { numeric: true }) || byName(a, b));
    else if (F.sort === 'name') rows.sort(byName);
    else rows.sort((a, b) => ORD[a.st] - ORD[b.st] || (a.d ?? 1e9) - (b.d ?? 1e9) || a.cat.localeCompare(b.cat, 'es') || byName(a, b));
    return rows;
  }
  function body() {
    const rows = filtered(), tb = q$('#ct-body'); if (!tb) return;
    const MAX = 400, show = rows.slice(0, MAX), pend = typeof pendingPaths === 'function' ? pendingPaths() : new Set();
    if (!rows.length) { tb.innerHTML = `<tr><td colspan="8"><div class="empty"><b>Nada coincide con el filtro</b>Cambiá la búsqueda, el estado o la categoría.</div></td></tr>`; selbar(); return; }
    tb.innerHTML = show.map(r => {
      const [sl, sc] = ST[r.st];
      const vh = r.venc ? `<span class="mono small">${e$(fdate(r.venc))}</span>${r.st === 'venc' ? ` <span class="small muted">hace ${-r.d} d</span>` : r.st === 'prox' ? ` <span class="small muted">en ${r.d} d</span>` : ''}${r.origen && r.origen !== 'manual' ? `<div class="small muted">según ${r.origen}</div>` : ''}` : '<span class="muted">—</span>';
      return `<tr data-id="${e$(r.it.id)}"${r.hist ? ' class="ct-hist"' : ''}><td><input type="checkbox" class="ct-ck" ${C.sel.has(r.it.id) ? 'checked' : ''} aria-label="Seleccionar"></td>
        <td class="ct-doc"><a href="${e$(r.it.webUrl)}" target="_blank" rel="noopener">${e$(r.it.name)}</a><div class="small muted">${e$(r.it.path.join(' › ') || 'Certificaciones')}</div>${r.linkTxt ? `<div class="small">🔗 ${e$(r.linkTxt)}</div>` : ''}${pend.has('cert/' + r.it.id) ? ' <span class="pill p-pend">Sin subir</span>' : ''}</td>
        <td>${e$(r.cat)}</td><td>${e$(r.elem || '—')}</td><td class="small">${e$(r.tipo || '—')}</td><td>${vh}</td><td><span class="pill ${sc}">${sl}</span></td>
        <td class="ct-acts"><button class="adj-btn" data-act="cls">Clasificar</button>${r.hist ? '' : `<button class="adj-btn${r.st === 'venc' || r.st === 'prox' ? ' ct-hot' : ''}" data-act="rep">Reemplazar</button>`}</td></tr>`;
    }).join('') + (rows.length > MAX ? `<tr><td colspan="8" class="small muted" style="text-align:center">Se muestran ${MAX} de ${rows.length}. Usá la búsqueda o los filtros para acotar.</td></tr>` : '');
    selbar();
  }
  function selbar() { const b = q$('#ct-selbar'); if (!b) return; b.hidden = !C.sel.size; q$('#ct-seln').textContent = `${C.sel.size} seleccionado${C.sel.size > 1 ? 's' : ''}`; }
  function bindHead() {
    const on = (s, ev, fn) => { const el = q$(s); if (el) el[ev] = fn; };
    on('#ct-sync', 'onclick', () => C.account ? scanAll() : connect(true));
    on('#ct-login', 'onclick', () => connect(true));
    on('#ct-retry', 'onclick', () => connect(true));
    on('#ct-csv', 'onclick', exportCsv);
  }
  function bindList() {
    const F = C.f;
    q$('#ct-q').oninput = e => { F.q = e.target.value; body(); };
    q$('#ct-st').oninput = e => { F.st = e.target.value; window.renderCert(); };
    q$('#ct-sort').oninput = e => { F.sort = e.target.value; body(); };
    q$('#ct-hist').oninput = e => { F.hist = e.target.checked; body(); };
    q$('#ct-cats').onclick = e => { const b = e.target.closest('[data-cat]'); if (!b) return; F.cat = b.dataset.cat; window.renderCert(); };
    document.querySelectorAll('.ct-k').forEach(b => b.onclick = () => { F.st = F.st === b.dataset.st ? '' : b.dataset.st; window.renderCert(); });
    q$('#ct-all').onchange = e => { for (const r of filtered().slice(0, 400)) e.target.checked ? C.sel.add(r.it.id) : C.sel.delete(r.it.id); body(); };
    q$('#ct-selx').onclick = () => { C.sel.clear(); q$('#ct-all').checked = false; body(); };
    q$('#ct-bulk').onclick = () => openCls([...C.sel]);
    q$('#ct-body').onclick = e => {
      const tr = e.target.closest('tr[data-id]'); if (!tr) return; const id = tr.dataset.id;
      if (e.target.classList.contains('ct-ck')) { e.target.checked ? C.sel.add(id) : C.sel.delete(id); selbar(); return; }
      const b = e.target.closest('[data-act]'); if (!b) return;
      if (b.dataset.act === 'cls') openCls([id]); else openRep(id);
    };
  }
  function setupHTML() {
    const red = new URL('ms-redirect.html', location.href).href;
    return `<section class="sec"><header><h3>Falta conectar el SGI con SharePoint</h3></header><div class="body ct-setup">
      <p>Esta pestaña lee la carpeta <b>Operaciones › Pozos › Certificaciones</b> con la cuenta de Microsoft de cada usuario. Se configura una sola vez (lo hace un administrador de Microsoft 365):</p>
      <ol><li>En <b>portal.azure.com</b> → Microsoft Entra ID → <b>Registros de aplicaciones</b> → Nuevo registro. Nombre: <i>SGI Condor</i>. Cuentas: <i>solo este directorio organizativo</i>.</li>
      <li>Plataforma <b>Aplicación de página única (SPA)</b> con el URI de redirección:<br><code>${e$(red)}</code></li>
      <li>Permisos de API → Microsoft Graph → <b>Delegados</b>: <code>Files.ReadWrite.All</code> y <code>Sites.Read.All</code> → <b>Conceder consentimiento de administrador</b>.</li>
      <li>Copiá el <b>Id. de aplicación (cliente)</b> en <code>ms-config.js</code> (campo <code>clientId</code>) y subilo al repositorio.</li></ol></div></section>`;
  }
  function exportCsv() {
    const rows = filtered();
    saveFile(`cwl_certificaciones_${today()}.csv`, csv([['Archivo', 'Carpeta', 'Categoría', 'Elemento', 'Tipo', 'Emisión', 'Vencimiento', 'Origen de la fecha', 'Estado', 'Vínculo SGI', 'Notas', 'Enlace'],
      ...rows.map(r => [r.it.name, r.it.path.join(' / '), r.cat, r.elem, r.tipo, r.cl.emision || '', r.venc, r.origen, ST[r.st][0], r.linkTxt, r.cl.notas || '', r.it.webUrl])]));
  }

  /* ================= modal: clasificar ================= */
  function modal(id, title, html) {
    let m = q$('#' + id);
    if (!m) { document.body.insertAdjacentHTML('beforeend', `<div class="ovl" id="${id}" hidden role="dialog" aria-modal="true"><div class="ovl-box"></div></div>`); m = q$('#' + id); m.onclick = e => { if (e.target === m && !m.dataset.busy) m.hidden = true; }; }
    m.setAttribute('aria-label', title); m.querySelector('.ovl-box').innerHTML = `<h3>${e$(title)}</h3>` + html; m.hidden = false; delete m.dataset.busy; return m;
  }
  const opts = (arr, v) => arr.map(x => `<option ${x === v ? 'selected' : ''}>${e$(x)}</option>`).join('');
  function linkOptions(rk, id) { const R = REG[rk]; if (!R) return ''; return colRows(R.c).map(d => [d._id, d.codigo ? `${d.codigo} · ${d.desc || ''}` : d.nro ? `${d.nro} · ${d.modelo || d.dominio || ''}` : (d.nombre || d._id)]).sort((a, b) => a[1].localeCompare(b[1], 'es', { numeric: true })).map(([i, l]) => `<option value="${e$(i)}" ${i === id ? 'selected' : ''}>${e$(l)}</option>`).join(''); }
  function campoOptions(rk, campo) { const R = REG[rk]; if (!R) return ''; return R.f.filter(f => f.venc).map(f => `<option value="${f.k}" ${f.k === campo ? 'selected' : ''}>${e$(f.l)}</option>`).join(''); }
  function openCls(ids) {
    if (!C.idx) return; ids = ids.filter(Boolean); if (!ids.length) return;
    const one = ids.length === 1, r = one ? info(C.idx.files.find(f => f.id === ids[0])) : null;
    const rows = all(); const cats = [...new Set(rows.map(x => x.cat))].sort((a, b) => a.localeCompare(b, 'es'));
    const elems = [...new Set([...rows.map(x => x.elem), ...colRows('eqp').map(d => d.codigo), ...colRows('hta').map(d => d.codigo), ...colRows('veh').map(d => d.nro)].filter(Boolean))].sort((a, b) => String(a).localeCompare(String(b), 'es', { numeric: true }));
    const cl = one ? r.cl : {}, link = cl.link || {};
    const folders = (C.idx.folders || []).filter(f => !HIST_RE.test(f.name)).sort((a, b) => a.path.join('/').localeCompare(b.path.join('/'), 'es', { numeric: true }));
    const m = modal('ctm', one ? 'Clasificar certificado' : `Clasificar ${ids.length} certificados`, `
      ${one ? `<p class="small" style="margin:0"><a href="${e$(r.it.webUrl)}" target="_blank" rel="noopener">${e$(r.it.name)}</a><br><span class="muted">${e$(r.it.path.join(' › ') || 'Certificaciones')}</span></p>` : '<p class="hint" style="margin:0">Lo que dejes vacío no se cambia.</p>'}
      <div class="grid ct-g">
        <label class="f c2"><span>Categoría</span><input id="cm-cat" list="cm-cats" value="${one ? e$(r.cat) : ''}" placeholder="${one ? '' : '(sin cambios)'}"><datalist id="cm-cats">${cats.map(c => `<option value="${e$(c)}">`).join('')}</datalist></label>
        <label class="f c2"><span>Elemento (equipo, unidad, persona)</span><input id="cm-elem" list="cm-elems" value="${one ? e$(r.elem) : ''}" placeholder="${one ? 'Ej.: WL-105, Hidrogrúa 203' : '(sin cambios)'}"><datalist id="cm-elems">${elems.map(c => `<option value="${e$(c)}">`).join('')}</datalist></label>
        <label class="f c2"><span>Tipo de documento</span><select id="cm-tipo"><option value="">${one ? '' : '(sin cambios)'}</option>${opts(TIPOS, one ? r.tipo : '')}</select></label>
        ${one ? `<label class="f"><span>Fecha de emisión</span><input id="cm-em" type="date" value="${e$(cl.emision || '')}"></label>
        <label class="f"><span>Vencimiento</span><input id="cm-venc" type="date" value="${e$(cl.venc || '')}">${!cl.venc && r.venc ? `<small class="hint">Detectado ${e$(r.origen)}: ${e$(fdate(r.venc))}</small>` : ''}</label>
        <label class="f c4" style="flex-direction:row;align-items:center;gap:8px"><input type="checkbox" id="cm-nv" ${cl.sinVenc ? 'checked' : ''} style="width:auto"> <span style="text-transform:none;letter-spacing:0;font:14px system-ui">No controla vencimiento (manual, ficha técnica, etc.)</span></label>
        <label class="f"><span>Vincular a registro del SGI</span><select id="cm-lrk"><option value="">— sin vínculo —</option>${Object.entries(LINKS).map(([k, l]) => `<option value="${k}" ${k === link.rk ? 'selected' : ''}>${l}</option>`).join('')}</select></label>
        <label class="f c2"><span>Registro</span><select id="cm-lid">${linkOptions(link.rk, link.id)}</select></label>
        <label class="f"><span>Campo de vencimiento</span><select id="cm-lf">${campoOptions(link.rk, link.campo)}</select></label>
        <label class="f c4" style="flex-direction:row;align-items:center;gap:8px"><input type="checkbox" id="cm-lsync" checked style="width:auto"> <span style="text-transform:none;letter-spacing:0;font:14px system-ui">Copiar el vencimiento a ese registro (Equipos, Flota, etc.)</span></label>
        <label class="f c4"><span>Mover a otra carpeta de SharePoint</span><select id="cm-mv"><option value="">— dejar donde está —</option>${folders.map(f => `<option value="${e$(f.id)}" ${f.id === r.it.parentId ? 'disabled' : ''}>${e$(f.path.join(' › '))}</option>`).join('')}</select></label>
        <label class="f c4"><span>Notas</span><textarea id="cm-not" rows="2">${e$(cl.notas || '')}</textarea></label>` : `<label class="f c2"><span>Mover a otra carpeta de SharePoint</span><select id="cm-mv"><option value="">— dejar donde están —</option>${folders.map(f => `<option value="${e$(f.id)}">${e$(f.path.join(' › '))}</option>`).join('')}</select></label>`}
      </div><div class="msg" id="cm-msg"></div>
      <div class="acts"><button type="button" class="btn-g" id="cm-x">Cancelar</button><button type="button" class="btn-p" id="cm-ok">Guardar</button></div>`);
    q$('#cm-x').onclick = () => { m.hidden = true; };
    const lrk = q$('#cm-lrk'); if (lrk) lrk.onchange = () => { q$('#cm-lid').innerHTML = linkOptions(lrk.value); q$('#cm-lf').innerHTML = campoOptions(lrk.value); };
    q$('#cm-ok').onclick = async () => {
      if (S.canWrite === false) return toast('Tu acceso es de solo lectura');
      const val = s => { const el = q$(s); return el ? el.value.trim() : ''; };
      const base = { updatedAt: new Date().toISOString(), editadoPor: S.uid || null };
      const mv = val('#cm-mv');
      const ok = q$('#cm-ok'); ok.disabled = true; m.dataset.busy = 1;
      try {
        if (mv) { q$('#cm-msg').textContent = ''; ok.textContent = 'Moviendo en SharePoint…'; for (const id of ids) await moveItem(id, mv); }
        for (const id of ids) {
          const f = C.idx.files.find(x => x.id === id); const d = Object.assign({ archivo: f ? f.name : '' }, base);
          if (one) {
            Object.assign(d, { cat: val('#cm-cat'), elem: val('#cm-elem'), tipo: val('#cm-tipo'), emision: val('#cm-em'), venc: val('#cm-venc'), sinVenc: q$('#cm-nv').checked, notas: val('#cm-not') });
            const rk = val('#cm-lrk'), lid = val('#cm-lid'), campo = val('#cm-lf');
            d.link = rk && lid ? { rk, id: lid, campo } : null;
            const venc = d.venc || info(f).venc;
            if (d.link && campo && venc && q$('#cm-lsync').checked) syncLink(d.link, venc, d.emision);
          } else {
            const cat = val('#cm-cat'), elem = val('#cm-elem'), tipo = val('#cm-tipo');
            if (cat) d.cat = cat; if (elem) d.elem = elem; if (tipo) d.tipo = tipo;
          }
          enqueue({ op: 'update', path: 'cert/' + id, data: d });
        }
        m.hidden = true; C.sel.clear(); toast(ids.length > 1 ? `${ids.length} certificados clasificados` : 'Clasificación guardada'); render();
      } catch (e) { q$('#cm-msg').textContent = 'No se pudo mover en SharePoint: ' + (e.message || e); }
      ok.disabled = false; ok.textContent = 'Guardar'; delete m.dataset.busy;
    };
  }
  function syncLink(link, venc, emision) {
    const R = REG[link.rk]; if (!R) return;
    const data = { [link.campo]: venc, updatedAt: new Date().toISOString(), editadoPor: S.uid || null };
    if (emision && R.f.some(f => f.k === 'fCert') && link.campo === 'vCert') data.fCert = emision;
    enqueue({ op: 'update', path: `${R.c}/${link.id}`, data });
  }
  async function moveItem(id, folderId) {
    if (!C.account) await token(true);
    const f = C.idx.files.find(x => x.id === id), dest = C.idx.folders.find(x => x.id === folderId);
    const it = await g(`/drives/${C.idx.root.driveId}/items/${id}?@microsoft.graph.conflictBehavior=rename`, { method: 'PATCH', json: { parentReference: { id: folderId } } });
    if (f) { f.parentId = folderId; f.path = dest ? dest.path.slice() : f.path; f.name = it.name; f.webUrl = it.webUrl; lsSet(C.idx); }
    return it;
  }

  /* ================= modal: reemplazar certificado ================= */
  function openRep(id) {
    const f = C.idx && C.idx.files.find(x => x.id === id); if (!f) return;
    const r = info(f), link = r.cl.link;
    const m = modal('ctr', 'Reemplazar certificado', `
      <div class="kvs ct-cur"><div><div class="l">Actual</div><div class="small"><a href="${e$(f.webUrl)}" target="_blank" rel="noopener">${e$(f.name)}</a></div></div><div><div class="l">Vencimiento</div><div class="small">${r.venc ? e$(fdate(r.venc)) : '—'} <span class="pill ${ST[r.st][1]}">${ST[r.st][0]}</span></div></div><div><div class="l">Carpeta</div><div class="small">${e$(f.path.join(' › ') || 'Certificaciones')}</div></div><div><div class="l">Elemento</div><div class="small">${e$(r.elem || '—')}</div></div></div>
      <label class="adj-btn ct-file"><input type="file" id="cr-file" accept=".pdf,image/*,.doc,.docx,.xls,.xlsx">📎 Elegir el certificado nuevo</label><div class="small" id="cr-fn"></div>
      <div class="grid ct-g">
        <label class="f c2"><span>Nombre en SharePoint</span><input id="cr-name" placeholder="Se completa al elegir el archivo"></label>
        <label class="f"><span>Fecha de emisión</span><input id="cr-em" type="date"></label>
        <label class="f"><span>Nuevo vencimiento <em>*</em></span><input id="cr-venc" type="date"></label>
      </div>
      <label style="display:flex;gap:8px;align-items:center;font-size:14px"><input type="checkbox" id="cr-mv" checked style="width:auto"> Mover el certificado anterior a la subcarpeta “${e$(CFG.carpetaVencidas)}”</label>
      <label style="display:flex;gap:8px;align-items:center;font-size:14px"><input type="checkbox" id="cr-tag" checked style="width:auto"> Agregar “vence dd-mm-aa” al nombre del archivo nuevo</label>
      ${link && link.rk ? `<label style="display:flex;gap:8px;align-items:center;font-size:14px"><input type="checkbox" id="cr-link" checked style="width:auto"> Actualizar el vencimiento en ${e$(linkLabel(link))}</label>` : ''}
      <div class="msg" id="cr-msg"></div>
      <div class="acts"><button type="button" class="btn-g" id="cr-x">Cancelar</button><button type="button" class="btn-p" id="cr-ok">Subir y reemplazar</button></div>`);
    let file = null;
    const tagName = () => {
      if (!file) return; const v = q$('#cr-venc').value, dot = file.name.lastIndexOf('.'), base = dot > 0 ? file.name.slice(0, dot) : file.name, ext = dot > 0 ? file.name.slice(dot) : '';
      const clean = base.replace(new RegExp('\\s*' + KW + '[\\d\\-/. a-z]*$', 'i'), '').trim() || base;
      q$('#cr-name').value = q$('#cr-tag').checked && v ? `${clean} vence ${v.slice(8, 10)}-${v.slice(5, 7)}-${v.slice(2, 4)}${ext}` : file.name;
    };
    q$('#cr-file').onchange = e => {
      file = e.target.files[0] || null; if (!file) return;
      q$('#cr-fn').textContent = `${file.name} · ${(file.size / 1048576).toFixed(1)} MB`;
      const pv = parseVenc(file.name); if (pv && !q$('#cr-venc').value) q$('#cr-venc').value = pv;
      tagName();
    };
    q$('#cr-venc').oninput = tagName; q$('#cr-tag').onchange = tagName;
    q$('#cr-x').onclick = () => { if (!m.dataset.busy) m.hidden = true; };
    q$('#cr-ok').onclick = async () => {
      const msg = q$('#cr-msg'); msg.textContent = '';
      if (S.canWrite === false) return toast('Tu acceso es de solo lectura');
      if (!file) return msg.textContent = 'Elegí el archivo del certificado nuevo.';
      const venc = q$('#cr-venc').value; if (!venc) return msg.textContent = 'Cargá la fecha de vencimiento del certificado nuevo.';
      if (typeof dayDiff === 'function' && dayDiff(venc) < 0) return msg.textContent = 'La fecha de vencimiento del certificado nuevo ya pasó. Revisala.';
      const name = (q$('#cr-name').value.trim() || file.name).replace(/[\\/:*?"<>|#%]/g, '-');
      const ok = q$('#cr-ok'); ok.disabled = true; m.dataset.busy = 1;
      const step = t => { ok.textContent = t; };
      try {
        if (!C.account) await token(true);
        step('Subiendo archivo…');
        const nu = await upload(f.parentId, name, file, p => step(`Subiendo… ${Math.round(p * 100)}%`));
        let movedTo = null;
        if (q$('#cr-mv').checked) { step('Moviendo el anterior…'); const vf = await ensureFolder(f.parentId, CFG.carpetaVencidas, f.path); await moveItem(f.id, vf.id); movedTo = vf.id; }
        const now = new Date().toISOString(), em = q$('#cr-em').value;
        // índice local
        C.idx.files.push({ id: nu.id, name: nu.name, path: (C.idx.folders.find(x => x.id === nu.parentReference?.id)?.path || f.path).slice(), parentId: f.parentId, webUrl: nu.webUrl, size: nu.size, mod: nu.lastModifiedDateTime, mime: nu.file && nu.file.mimeType });
        lsSet(C.idx);
        // clasificación: el nuevo hereda la del anterior; el anterior queda como histórico
        const prev = r.cl;
        enqueue({ op: 'set', path: 'cert/' + nu.id, data: { archivo: nu.name, cat: prev.cat || r.cat, elem: prev.elem || r.elem, tipo: prev.tipo || r.tipo || 'Certificado', emision: em, venc, sinVenc: false, notas: prev.notas || '', link: prev.link || null,
          reemplazaA: { id: f.id, archivo: f.name, venc: r.venc || '' }, creadoPor: S.uid || null, createdAt: now, updatedAt: now, editadoPor: S.uid || null } });
        enqueue({ op: 'update', path: 'cert/' + f.id, data: { archivo: f.name, estado: 'Reemplazado', reemplazadoPor: nu.id, reemplazadoEl: now, movidoA: movedTo, venc: prev.venc || r.venc || '', updatedAt: now, editadoPor: S.uid || null } });
        const lk = q$('#cr-link'); if (prev.link && lk && lk.checked && prev.link.campo) syncLink(prev.link, venc, em);
        m.hidden = true; toast('Certificado reemplazado en SharePoint'); render();
      } catch (e) { console.error('[SGI cert] reemplazo', e); msg.textContent = 'No se pudo completar: ' + (e.message || e) + '. Revisá en SharePoint qué quedó hecho antes de reintentar.'; }
      ok.disabled = false; ok.textContent = 'Subir y reemplazar'; delete m.dataset.busy;
    };
  }
  async function upload(parentId, name, file, onP) {
    const d = C.idx.root.driveId, enc = encodeURIComponent(name);
    if (file.size <= 4 * 1048576) {
      const t = await token(false);
      const r = await fetch(`${GRAPH}/drives/${d}/items/${parentId}:/${enc}:/content?@microsoft.graph.conflictBehavior=rename`, { method: 'PUT', headers: { Authorization: 'Bearer ' + t, 'Content-Type': file.type || 'application/octet-stream' }, body: file });
      if (!r.ok) { let m = r.status; try { m = (await r.json()).error.message; } catch (e) { } throw new Error(m); }
      onP && onP(1); return r.json();
    }
    const s = await g(`/drives/${d}/items/${parentId}:/${enc}:/createUploadSession`, { method: 'POST', json: { item: { '@microsoft.graph.conflictBehavior': 'rename' } } });
    const CH = 320 * 1024 * 16; let pos = 0, res = null;
    while (pos < file.size) {
      const end = Math.min(pos + CH, file.size);
      const r = await fetch(s.uploadUrl, { method: 'PUT', headers: { 'Content-Range': `bytes ${pos}-${end - 1}/${file.size}` }, body: file.slice(pos, end) });
      if (!r.ok && r.status !== 202) throw new Error('Error al subir (' + r.status + ')');
      if (r.status === 200 || r.status === 201) res = await r.json();
      pos = end; onP && onP(pos / file.size);
    }
    return res;
  }
  async function ensureFolder(parentId, name, parentPath) {
    const known = C.idx.folders.find(x => x.parentId === parentId && x.name.toLowerCase() === name.toLowerCase()); if (known) return known;
    const d = C.idx.root.driveId; let it;
    try { it = await g(`/drives/${d}/items/${parentId}:/${encodeURIComponent(name)}?$select=id,name,webUrl`); }
    catch (e) { if (e.status !== 404) throw e; it = await g(`/drives/${d}/items/${parentId}/children`, { method: 'POST', json: { name, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' } }); }
    const fo = { id: it.id, name: it.name, path: (parentPath || []).concat(it.name), parentId, webUrl: it.webUrl }; C.idx.folders.push(fo); return fo;
  }

  /* ================= resumen para Inicio ================= */
  window.certResumen = function () {
    if (!C.idx) return null; const live = all().filter(r => !r.hist);
    return { venc: live.filter(r => r.st === 'venc').length, prox: live.filter(r => r.st === 'prox').length, total: live.length };
  };
  window.SGI_CERT = { parseVenc, info, state: C };   // para pruebas y soporte

  // si el usuario ya se conectó antes en este equipo, se reconecta y actualiza en silencio al entrar a la pestaña
  let auto = false;
  window.certAuto = function () {
    if (auto || !configured()) return; auto = true;
    let was = false; try { was = localStorage.getItem('cwl4.msOn') === '1'; } catch (e) { }
    if (!was) return;
    ensureMsal().then(() => { if (!C.account) return; render(); const old = !C.idx || Date.now() - new Date(C.idx.at).getTime() > 6 * 36e5; if (old) scanAll(); }).catch(e => console.warn('[SGI cert]', e));
  };
})();
