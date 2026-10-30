// Utilidades comunes: creación de elementos, fechas, texto, mensajes y ventanas.
export function h(tag, attrs, ...hijos) {
  const el = document.createElement(tag);
  if (attrs && (typeof attrs !== 'object' || attrs instanceof Node || Array.isArray(attrs))) { hijos.unshift(attrs); attrs = null; }
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v == null || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k === 'html') el.innerHTML = v;
    else if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
    else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2).toLowerCase(), v);
    else if (k === 'value') el.value = v;
    else if (v === true) el.setAttribute(k, '');
    else el.setAttribute(k, v);
  }
  const añadir = (x) => {
    if (x == null || x === false) return;
    if (Array.isArray(x)) x.forEach(añadir);
    else el.append(x instanceof Node ? x : document.createTextNode(String(x)));
  };
  hijos.forEach(añadir);
  return el;
}
export const $ = (s, r = document) => r.querySelector(s);
export const $$ = (s, r = document) => [...r.querySelectorAll(s)];
export const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
export const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
export const norm = (s) => String(s ?? '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');

// ---- fechas (siempre en formato ISO AAAA-MM-DD en los datos)
export const iso = (d) => { const x = new Date(d); return x.getFullYear() + '-' + String(x.getMonth() + 1).padStart(2, '0') + '-' + String(x.getDate()).padStart(2, '0'); };
export const hoy = () => iso(new Date());
export const aFecha = (s) => { if (!s) return null; const [a, m, d] = String(s).slice(0, 10).split('-').map(Number); return new Date(a, m - 1, d); };
export const fmtFecha = (s) => { const d = aFecha(s); return d ? d.toLocaleDateString('es-CO', { day: '2-digit', month: 'short', year: 'numeric' }) : '—'; };
export const sumarDias = (s, n) => { const d = aFecha(s); d.setDate(d.getDate() + n); return iso(d); };
export const sumarMeses = (s, n) => { const d = aFecha(s); const dia = d.getDate(); d.setMonth(d.getMonth() + n); if (d.getDate() < dia) d.setDate(0); return iso(d); };
export const diasEntre = (a, b) => Math.round((aFecha(b) - aFecha(a)) / 86400000);
export const mesesEntre = (a, b) => { const x = aFecha(a), y = aFecha(b); let m = (y.getFullYear() - x.getFullYear()) * 12 + y.getMonth() - x.getMonth(); if (y.getDate() < x.getDate()) m--; return m; };
export function diasHabilesEntre(a, b) { // sin sábados ni domingos (los festivos no se descuentan)
  let n = 0; const d = aFecha(a), f = aFecha(b); if (!d || !f) return 0;
  while (d < f) { d.setDate(d.getDate() + 1); const w = d.getDay(); if (w !== 0 && w !== 6) n++; }
  return n;
}
export const textoPlazo = (dias) => dias < 0 ? `vencido hace ${-dias} d` : dias === 0 ? 'vence hoy' : `faltan ${dias} d`;

// ---- avisos y ventanas
let cont;
export function toast(msg, tipo = '') {
  if (!cont) { cont = h('div', { class: 'toasts' }); document.body.append(cont); }
  const t = h('div', { class: 'toast ' + tipo }, msg); cont.append(t); setTimeout(() => t.remove(), 3800);
}
export function modal(titulo, contenido, botones = []) {
  return new Promise((res) => {
    const cerrar = (v) => { f.remove(); res(v); };
    const f = h('div', { class: 'mod-f', onclick: (e) => { if (e.target === f) cerrar(null); } },
      h('div', { class: 'mod', role: 'dialog', 'aria-label': titulo },
        h('h2', titulo), contenido,
        h('div', { class: 'fila', style: { marginTop: '12px', justifyContent: 'flex-end' } },
          ...botones.map((b) => h('button', { class: 'btn ' + (b.clase || ''), onclick: async () => { if (b.valida && !(await b.valida())) return; cerrar(b.valor ?? true); } }, b.texto)))));
    document.body.append(f);
  });
}
// Tema visual: 'claro' (por defecto), 'oscuro' o 'auto' (el del dispositivo). Se guarda solo en este dispositivo.
export const getTema = () => document.documentElement.dataset.tema || 'claro';
export function setTema(t) {
  document.documentElement.dataset.tema = t;
  try { localStorage.setItem('tema', t); } catch { /* sin almacenamiento: vale solo para esta sesión */ }
  document.dispatchEvent(new CustomEvent('tema'));
}
export const confirmar =(msg, si = 'Aceptar') => modal('Confirmar', h('p', msg), [{ texto: 'Cancelar', clase: 'gris', valor: false }, { texto: si, valor: true }]).then((v) => v === true);
// Modo instalado: 'android' (Capacitor), 'windows' (Electron) o null (navegador).
export const modoInstalado = () => (window.Capacitor?.isNativePlatform?.() ? 'android' : window.appEscritorio ? 'windows' : null);
export async function descargar(nombre, datos, tipo = 'application/octet-stream') {
  const blob = datos instanceof Blob ? datos : new Blob([datos], { type: tipo });
  if (modoInstalado() === 'android') { // en Android el WebView no descarga: se guarda en la caché y se abre el menú «Compartir / Guardar»
    try {
      const P = window.Capacitor.Plugins;
      const r = await P.Filesystem.writeFile({ path: 'compartir/' + nombre, data: aBase64(await blob.arrayBuffer()), directory: 'CACHE', recursive: true });
      await P.Share.share({ title: nombre, url: r.uri, dialogTitle: 'Guardar o abrir ' + nombre });
    } catch (e) { if (!/cancel/i.test(String(e?.message))) toast('No se pudo guardar el archivo: ' + (e?.message || e), 'mal'); }
    return;
  }
  const a = h('a', { href: URL.createObjectURL(blob), download: nombre }); document.body.append(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
}
export const leerComoTexto = (f) => new Promise((ok, no) => { const r = new FileReader(); r.onload = () => ok(r.result); r.onerror = no; r.readAsText(f); });
export const leerComoBuffer = (f) => new Promise((ok, no) => { const r = new FileReader(); r.onload = () => ok(r.result); r.onerror = no; r.readAsArrayBuffer(f); });
export const aBase64 = (buf) => { let s = ''; const b = new Uint8Array(buf); for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000)); return btoa(s); };
export const deBase64 = (b64) => { const s = atob(b64); const b = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) b[i] = s.charCodeAt(i); return b.buffer; };
export const tamaño = (n) => n > 1048576 ? (n / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB';
export const pct = (x) => (x == null ? '—' : Math.round(x * 100) + ' %');
export const resaltar = (texto, claves) => {
  let html = esc(texto);
  claves.filter((c) => c.length > 1).forEach((c) => { const re = new RegExp('(' + c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')', 'ig'); html = html.replace(re, '<mark class="hl">$1</mark>'); });
  return html;
};
export function icono(nombre) {
  const P = {
    inicio: 'M3 11l9-8 9 8M5 10v10h5v-6h4v6h5V10', opciones: 'M4 6h16M4 12h16M4 18h10', verificador: 'M9 12l2 2 4-4M12 3a9 9 0 100 18 9 9 0 000-18z',
    expedientes: 'M4 4h6l2 2h8v14H4zM4 9h16', docentes: 'M16 11a4 4 0 10-8 0 4 4 0 008 0zM4 21c1-4 4-6 8-6s7 2 8 6', calendario: 'M4 6h16v14H4zM4 10h16M8 3v4M16 3v4',
    formatos: 'M6 3h9l4 4v14H6zM14 3v5h5M9 13h7M9 17h7', reglamento: 'M5 4h12a2 2 0 012 2v14H7a2 2 0 01-2-2zM9 8h6M9 12h6', diagramas: 'M5 4h5v4H5zM14 4h5v4h-5zM9.5 16h5v4h-5zM7.5 8v4h9V8M12 12v4',
    extraccion: 'M12 3v12M7 10l5 5 5-5M4 20h16', fichas: 'M4 5h16v14H4zM4 15l5-5 4 4 3-3 4 4', ajustes: 'M12 9a3 3 0 100 6 3 3 0 000-6zM19 12l2-1-1-3-2 .5-1.5-1.5.5-2-3-1-1 2h-2l-1-2-3 1 .5 2L6 8.500 4 8 3 11l2 1v2l-2 1 1 3 2-.5L7.500 17.500 7 19.500l3 1 1-2h2l1 2 3-1-.5-2 1.500-1.500 2 .5 1-3-2-1z',
    mas: 'M5 12h.01M12 12h.01M19 12h.01', alerta: 'M12 3l10 18H2zM12 10v5M12 18h.01',
  };
  const s = document.createElementNS('http://www.w3.org/2000/svg', 'svg'); s.setAttribute('viewBox', '0 0 24 24'); s.setAttribute('class', 'ico'); s.setAttribute('aria-hidden', 'true');
  const p = document.createElementNS('http://www.w3.org/2000/svg', 'path'); p.setAttribute('d', P[nombre] || P.opciones); s.append(p); return s;
}
export function cargarScript(src) {
  return new Promise((ok, no) => { if ([...document.scripts].some((s) => s.src.endsWith(src))) return ok(); const s = h('script', { src, onload: ok, onerror: () => no(new Error('No se pudo cargar ' + src)) }); document.head.append(s); });
}
