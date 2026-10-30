// Arranque de la aplicación de extracción de datos: mismo diseño que la aplicación de opciones de grado, sin perfiles ni administrador.
import { h, $, icono, getTema, setTema, modoInstalado, descargar } from './util.js';
import { RUTAS } from './rutas.js';

let instalar = null;
function parsear() { const [r, ...p] = (location.hash.replace(/^#\/?/, '') || 'extraccion').split('/'); return { id: r || 'extraccion', params: p.map(decodeURIComponent) }; }

let _turno = 0; // si se navega rápido, solo la última sección pedida se dibuja
async function pintar() {
  const turno = ++_turno;
  const { id, params } = parsear();
  const def = RUTAS.find((r) => r.id === id) || RUTAS[0];
  $('#vista-titulo').textContent = def.nombre;
  document.title = def.nombre + ' · Extracción de datos FUP';
  document.querySelectorAll('[data-ruta]').forEach((a) => a.classList.toggle('on', a.dataset.ruta === def.id));
  const cont = $('#vista'); cont.replaceChildren(h('div', { class: 'vacio' }, 'Cargando…'));
  try {
    const m = await import(def.mod);
    const nodo = await m.render({ params });
    if (turno !== _turno) return;
    cont.replaceChildren(nodo);
    window.scrollTo(0, 0);
  } catch (e) {
    console.error(e);
    cont.replaceChildren(h('div', { class: 'alerta mal' }, 'No se pudo abrir esta sección: ' + e.message));
  }
}

function layout() {
  const nav = h('nav');
  let g = '';
  RUTAS.forEach((r) => {
    if (r.grupo !== g) { g = r.grupo; nav.append(h('div', { class: 'grupo' }, g)); }
    nav.append(h('a', { class: 'nav', href: '#/' + r.id, 'data-ruta': r.id }, icono(r.ico), r.nombre));
  });
  const lado = h('aside', { class: 'lado' },
    h('div', { class: 'marca' }, h('img', { src: 'iconos/logo_fup.png', alt: 'FUP' }), h('div', h('b', 'Extracción de Datos'), h('small', 'Opciones de Grado · Ingeniería Industrial · FUP'))),
    nav,
    h('div', { class: 'pie' }, 'Los archivos se leen en este dispositivo; no se envían a ningún servidor.'));
  const inf = h('nav', { class: 'inf', 'aria-label': 'Navegación principal' },
    ...RUTAS.filter((r) => r.movil).map((r) => h('a', { href: '#/' + r.id, 'data-ruta': r.id }, icono(r.ico), h('span', r.nombre.split(' ')[0]))));
  const barra = h('header', { class: 'barra' }, h('h1', { id: 'vista-titulo' }, ''),
    h('button', { class: 'btn peq sec', id: 'btn-tema', title: 'Cambiar entre tema claro y oscuro', onclick: () => setTema(getTema() === 'oscuro' ? 'claro' : 'oscuro') }),
    h('button', { class: 'btn peq sec', id: 'btn-instalar', style: { display: 'none' }, onclick: async () => { if (instalar) { instalar.prompt(); await instalar.userChoice; instalar = null; $('#btn-instalar').style.display = 'none'; } } }, 'Instalar app'));
  $('#app').replaceChildren(h('div', { class: 'app' }, lado, h('div', { class: 'cuerpo' }, h('div', { id: 'aviso-red' }), barra, h('main', { class: 'principal', id: 'vista' })), inf));
}

async function iniciar() {
  layout();
  const botonTema = () => { const b = $('#btn-tema'); if (b) b.textContent = getTema() === 'oscuro' ? '☀ Claro' : '🌙 Oscuro'; };
  document.addEventListener('tema', botonTema); botonTema();
  window.addEventListener('hashchange', pintar);
  window.addEventListener('beforeinstallprompt', (e) => { e.preventDefault(); instalar = e; const b = $('#btn-instalar'); if (b) b.style.display = ''; });
  const red = () => { const a = $('#aviso-red'); if (a) a.replaceChildren(navigator.onLine ? '' : h('div', { class: 'offline' }, 'Sin conexión: la extracción funciona igual, porque Python va incluido en la aplicación.')); };
  window.addEventListener('online', red); window.addEventListener('offline', red); red();
  if (modoInstalado() === 'android') { // los enlaces a archivos incluidos en la app se guardan con «Compartir»
    document.addEventListener('click', async (ev) => {
      const a = ev.target.closest && ev.target.closest('a[download], a[target=_blank]'); if (!a) return;
      const href = a.getAttribute('href') || ''; if (!href || href.startsWith('#') || /^(https?:|mailto:)/.test(href) && !href.startsWith(location.origin)) return;
      ev.preventDefault(); const r = await fetch(a.href); await descargar(a.getAttribute('download') || decodeURIComponent(href.split('/').pop()), await r.blob());
    });
  }
  if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol) && !modoInstalado() && location.hostname !== 'localhost.invalid') {
    navigator.serviceWorker.register('sw.js').catch((e) => console.warn('Sin modo offline:', e));
  }
  await pintar();
}
iniciar().catch((e) => { document.body.prepend(h('pre', { style: { color: 'red', padding: '12px' } }, 'Error al iniciar: ' + e.message)); console.error(e); });
