/*
  SGI Condor Wireline — conexión con Firebase (login + base de datos Firestore con modo sin conexión).
  La página (index.html) llama a window.SGI_BACKEND.init() y recibe { db, user, downloads }.
  No hace falta tocar este archivo: la configuración va en firebase-config.js.
*/
(function () {
  const cfg = window.SGI_FIREBASE_CONFIG || {};
  const AVATAR = 'data:image/svg+xml;utf8,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><circle cx="12" cy="12" r="12" fill="#0d5b6b"/><circle cx="12" cy="9.5" r="4" fill="#fff"/><path d="M4.5 20c1.5-3.6 4.3-5.2 7.5-5.2s6 1.6 7.5 5.2" fill="#fff"/></svg>');

  /* ---------- pantalla de ingreso ---------- */
  const css = `#sgi-login{position:fixed;inset:0;z-index:100;background:#0e2230;display:flex;align-items:center;justify-content:center;padding:16px;font-family:"IBM Plex Sans",system-ui,sans-serif}
  #sgi-login .box{background:#fff;color:#172126;border-radius:16px;padding:28px 26px;width:100%;max-width:380px;box-shadow:0 20px 60px rgba(0,0,0,.35)}
  #sgi-login img{height:54px;display:block;margin:0 auto 14px}
  #sgi-login h1{font:700 20px "IBM Plex Sans Condensed",system-ui,sans-serif;text-align:center;margin:0 0 4px}
  #sgi-login p{font-size:13.5px;color:#4b5a60;text-align:center;margin:0 0 18px}
  #sgi-login label{display:block;font:600 11px "IBM Plex Sans Condensed",sans-serif;letter-spacing:.08em;text-transform:uppercase;color:#4b5a60;margin:10px 0 4px}
  #sgi-login input{width:100%;box-sizing:border-box;font:15px system-ui,sans-serif;padding:10px;border:1px solid #d5dcda;border-radius:6px;background:#f6f8f7}
  #sgi-login button{width:100%;margin-top:14px;font:600 15px system-ui,sans-serif;padding:11px;border-radius:6px;border:1px solid #0d5b6b;background:#0d5b6b;color:#fff;cursor:pointer}
  #sgi-login button.alt{background:#fff;color:#0d5b6b}
  #sgi-login .lnk{background:none;border:0;color:#0d5b6b;padding:6px;margin-top:6px;font-size:13px;width:auto;display:block;margin-inline:auto}
  #sgi-login .msg{min-height:20px;font-size:13px;color:#b3261e;text-align:center;margin-top:10px}
  #sgi-login .or{display:flex;align-items:center;gap:8px;color:#7a888d;font-size:12px;margin-top:14px}
  #sgi-login .or::before,#sgi-login .or::after{content:"";flex:1;height:1px;background:#e4e9e7}`;
  function loginUI(auth) {
    if (document.getElementById('sgi-login')) return;
    const st = document.createElement('style'); st.textContent = css; document.head.appendChild(st);
    const d = document.createElement('div'); d.id = 'sgi-login';
    d.innerHTML = `<form class="box" novalidate>
      <img src="logo.png" alt="Condor Wireline">
      <h1>SGI Condor Wireline</h1><p>Ingresá con tu usuario para cargar y consultar el sistema.</p>
      <label for="sgi-email">Email</label><input id="sgi-email" type="email" autocomplete="username" required>
      <label for="sgi-pass">Contraseña</label><input id="sgi-pass" type="password" autocomplete="current-password" required>
      <button type="submit">Ingresar</button>
      <button type="button" class="lnk" id="sgi-reset">Olvidé mi contraseña</button>
      <div class="or">o</div>
      <button type="button" class="alt" id="sgi-google">Ingresar con Google</button>
      <div class="msg" id="sgi-msg" role="status"></div></form>`;
    document.body.appendChild(d);
    const msg = t => { d.querySelector('#sgi-msg').textContent = t || ''; };
    const err = e => {
      const c = (e && e.code) || '';
      msg(c.includes('wrong-password') || c.includes('invalid-credential') || c.includes('user-not-found') ? 'Email o contraseña incorrectos.' :
          c.includes('too-many-requests') ? 'Demasiados intentos. Probá de nuevo en unos minutos.' :
          c.includes('network') ? 'Sin conexión. Para ingresar la primera vez necesitás internet.' : 'No se pudo ingresar: ' + (e.message || c));
    };
    d.querySelector('form').onsubmit = ev => { ev.preventDefault(); msg('Ingresando…'); auth.signInWithEmailAndPassword(d.querySelector('#sgi-email').value.trim(), d.querySelector('#sgi-pass').value).catch(err); };
    d.querySelector('#sgi-reset').onclick = () => { const em = d.querySelector('#sgi-email').value.trim(); if (!em) return msg('Escribí tu email y tocá de nuevo.'); auth.sendPasswordResetEmail(em).then(() => msg('Te enviamos un email para cambiar la contraseña.')).catch(err); };
    d.querySelector('#sgi-google').onclick = () => { const p = new firebase.auth.GoogleAuthProvider(); auth.signInWithPopup(p).catch(e => { if (String(e.code).includes('popup')) auth.signInWithRedirect(p); else err(e); }); };
  }
  function blocked(auth, email) {
    loginUI(auth);
    const box = document.querySelector('#sgi-login .box');
    box.innerHTML = `<img src="logo.png" alt=""><h1>Usuario no habilitado</h1><p>La cuenta <b></b> todavía no tiene acceso al sistema. Pedile al administrador que la agregue en la lista de acceso.</p><button type="button" id="sgi-out">Ingresar con otra cuenta</button>`;
    box.querySelector('b').textContent = email;
    box.querySelector('#sgi-out').onclick = () => auth.signOut().then(() => location.reload());
  }

  /* ---------- adaptador con la misma forma que usa la página ---------- */
  function wrapSnap(s) { return { docs: s.docs.map(d => ({ id: d.id, exists: true, data: () => d.data() })), size: s.size, empty: s.empty, metadata: { fromCache: s.metadata.fromCache, hasPendingWrites: s.metadata.hasPendingWrites } }; }
  function mapErr(e) { const c = (e && e.code) || ''; return { code: c === 'permission-denied' ? 'unavailable' : c === 'resource-exhausted' ? 'quota_exceeded' : 'unavailable', message: (e && e.message) || String(e) }; }
  function makeDb(fs) {
    const coll = path => {
      const q = fs.collection(path);
      return {
        path,
        onSnapshot(next, error) { return q.onSnapshot({ includeMetadataChanges: true }, s => next(wrapSnap(s)), e => { console.warn('[SGI]', path, e); error && error(mapErr(e)); }); },
        async get() { return wrapSnap(await q.get()); },
        doc: id => docRef(path + '/' + id),
      };
    };
    const docRef = path => {
      const r = fs.doc(path);
      return {
        path,
        // update() de la página espera "merge profundo": en Firestore eso es set(..., {merge:true})
        set: d => r.set(d).catch(e => { throw mapErr(e); }),
        update: d => r.set(d, { merge: true }).catch(e => { throw mapErr(e); }),
        delete: () => r.delete().catch(e => { throw mapErr(e); }),
        async get() { const s = await r.get(); return { id: s.id, exists: s.exists, data: () => s.data() }; },
        collection: p => coll(path + '/' + p),
      };
    };
    return { collection: coll, doc: docRef };
  }
  const downloads = {
    async save({ filename, data }) {
      const blob = data instanceof Blob ? data : new Blob([data]);
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename;
      document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
      return { status: 'saved' };
    }
  };

  window.SGI_BACKEND = {
    async init() {
      if (!cfg.apiKey || /COMPLETAR/.test(cfg.apiKey)) {
        console.warn('[SGI] Falta completar firebase-config.js: la app funciona solo en este equipo.');
        return { db: null, user: null, downloads };
      }
      firebase.initializeApp(cfg);
      const auth = firebase.auth(), fs = firebase.firestore();
      try { await fs.enablePersistence({ synchronizeTabs: true }); } catch (e) { console.warn('[SGI] Sin caché offline:', e.code); }
      const u = await new Promise(res => auth.onAuthStateChanged(x => { if (x) res(x); else loginUI(auth); }));
      // ¿está habilitado? (lista "acceso" en Firestore, un documento por email)
      const email = String(u.email || '').toLowerCase();
      try {
        const a = await fs.doc('acceso/' + email).get();
        if (!a.exists && !a.metadata.fromCache) { blocked(auth, email); return new Promise(() => {}); }
      } catch (e) { if (e.code === 'permission-denied') { blocked(auth, email); return new Promise(() => {}); } }
      const lg = document.getElementById('sgi-login'); if (lg) lg.remove();
      fs.doc('usuarios/' + u.uid).set({ name: u.displayName || email.split('@')[0], email, photo: u.photoURL || '', ultimoIngreso: new Date().toISOString() }, { merge: true }).catch(() => {});
      const names = {};
      const user = {
        id: async () => u.uid,
        can: async () => true,
        canEdit: async () => true,
        isOwner: async () => false,
        me: async () => ({ id: u.uid, name: u.displayName || email.split('@')[0], avatarUrl: u.photoURL || AVATAR, email, color: '#0d5b6b', isOwner: false, canEdit: true }),
        async profiles(ids) {
          const out = {};
          await Promise.all([].concat(ids).map(async id => {
            if (!(id in names)) { try { const s = await fs.doc('usuarios/' + id).get(); names[id] = s.exists ? (s.data().name || '') : ''; } catch (e) { names[id] = ''; } }
            out[id] = { id, name: names[id], avatarUrl: AVATAR, color: '#0d5b6b', email: null, isMe: id === u.uid, guest: false };
          }));
          return out;
        },
      };
      // botón "Cerrar sesión" al pie del menú
      setTimeout(() => {
        const f = document.querySelector('.sb-foot');
        if (f && !document.getElementById('sgi-logout')) {
          const b = document.createElement('button'); b.id = 'sgi-logout'; b.type = 'button'; b.textContent = 'Cerrar sesión · ' + email;
          b.style.cssText = 'display:block;margin-top:8px;background:none;border:1px solid #2c4a58;color:#dbe6e9;border-radius:6px;padding:6px 8px;font:12px system-ui;cursor:pointer;width:100%;text-align:left;overflow:hidden;text-overflow:ellipsis';
          b.onclick = () => auth.signOut().then(() => location.reload());
          f.appendChild(b);
        }
      }, 300);
      return { db: makeDb(fs), user, downloads };
    }
  };
})();
