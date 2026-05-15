// ── AClímate — Autenticación por PIN ─────────────────────────────────
// ¡IMPORTANTE! Cambiá los PINs antes de entregar a los usuarios.

const USUARIOS = [
  { nombre: 'Ericsen Aguirre',  rol: 'tecnico',    pin: '1111' },
  { nombre: 'Javier Quesada',   rol: 'tecnico',    pin: '2222' },
  { nombre: 'Alexander Alfaro', rol: 'supervisor', pin: '3333' },
  { nombre: 'Andrés Porras',    rol: 'pm',         pin: '4444' },
];

const _SK  = 'aclimate_session';
const _TTL = 24 * 60 * 60 * 1000; // 24 horas

function _getSession() {
  try {
    const s = JSON.parse(localStorage.getItem(_SK) || 'null');
    if (!s || Date.now() > s.expiry) { localStorage.removeItem(_SK); return null; }
    return s;
  } catch(e) { return null; }
}

// Redirige a login si no hay sesión activa
function checkAuth() {
  if (!_getSession()) window.location.href = 'login.html';
}

// Devuelve el objeto de sesión actual o null
function getSession() { return _getSession(); }

// Cierra sesión y va a login
function logout() {
  localStorage.removeItem(_SK);
  window.location.href = 'login.html';
}

// Inyecta badge de usuario en el header de la página actual
function inyectarBadge() {
  const s = _getSession();
  if (!s) return;
  const primerNombre = s.nombre.split(' ')[0];

  const header = document.querySelector('.header');
  if (header) {
    // Header oscuro (azul-dark) → texto blanco
    const b = document.createElement('div');
    b.style.cssText = 'display:flex;flex-direction:column;align-items:flex-end;gap:3px;flex-shrink:0;';
    b.innerHTML = `
      <span style="font-size:.65rem;font-weight:700;color:#fff;opacity:.9;">👤 ${primerNombre}</span>
      <button onclick="logout()" style="background:none;border:1px solid rgba(255,255,255,.4);border-radius:8px;font-size:.58rem;color:rgba(255,255,255,.8);cursor:pointer;padding:1px 6px;line-height:1.5;">Salir</button>
    `;
    header.appendChild(b);
    return;
  }

  // Hero (index.html) → línea debajo con fondo claro
  const hero = document.querySelector('.hero');
  if (hero) {
    const b = document.createElement('div');
    b.style.cssText = 'font-size:.75rem;color:var(--texto-sub);margin-top:6px;';
    b.innerHTML = `👤 ${s.nombre} &nbsp;·&nbsp; <a href="#" onclick="logout();return false;" style="color:var(--azul);font-weight:600;">Salir</a>`;
    hero.appendChild(b);
  }
}

// Auto-selecciona el técnico logueado en el select del formulario
function autoSeleccionarTecnico() {
  const s = _getSession();
  if (!s) return;
  const sel = document.getElementById('select-tecnico')
           || document.querySelector('select[name="tecnico"]');
  if (!sel) return;
  for (const opt of sel.options) {
    if (opt.value === s.nombre) {
      sel.value = s.nombre;
      sel.dispatchEvent(new Event('change'));
      return;
    }
  }
}
