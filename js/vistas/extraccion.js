// Extracción de datos: ejecuta los scripts de Python del proyecto (extraccion_comun.py + extraer_formato.py) DENTRO de la aplicación con Pyodide.
// Es el mismo código que se usa desde la línea de comandos: un tipo de documento por vez, nunca se mezclan formatos distintos. No se envía nada a ningún servidor.
import { h, cargarScript, descargar, leerComoBuffer, tamaño, toast } from '../util.js';

const RUEDAS = ['et_xmlfile-2.0.0-py3-none-any.whl', 'openpyxl-3.1.5-py2.py3-none-any.whl'];

let _py; const nombres = {}; // promesa del intérprete de Python y nombres de los tipos de documento detectados
async function python(estado) {
  if (_py) return _py;
  _py = (async () => {
    estado('Cargando Python en la aplicación (la primera vez tarda unos segundos)…');
    await cargarScript('vendor/pyodide/pyodide.js');
    const py = await window.loadPyodide({ indexURL: 'vendor/pyodide/' });
    estado('Instalando la biblioteca de lectura de Excel…');
    for (const w of RUEDAS) {
      const r = await fetch('vendor/ruedas/' + w); if (!r.ok) throw new Error('Falta ' + w);
      await py.unpackArchive(await r.arrayBuffer(), 'wheel');
    }
    for (const f of ['extraer_formato.py', 'extraccion_comun.py']) {
      const r = await fetch('py/' + f); if (!r.ok) throw new Error('Falta py/' + f);
      py.FS.writeFile('/home/pyodide/' + f, await r.text());
    }
    await py.runPythonAsync('import sys; sys.path.insert(0, "/home/pyodide"); import openpyxl, extraccion_comun');
    return py;
  })().catch((e) => { _py = null; throw e; });
  return _py;
}
function csv(texto) { // lector sencillo de CSV con separador ; y comillas
  const filas = []; let f = [], c = '', q = false;
  for (let i = 0; i < texto.length; i++) {
    const x = texto[i];
    if (q) { if (x === '"') { if (texto[i + 1] === '"') { c += '"'; i++; } else q = false; } else c += x; }
    else if (x === '"') q = true; else if (x === ';') { f.push(c); c = ''; } else if (x === '\n') { f.push(c); filas.push(f); f = []; c = ''; } else if (x !== '\r') c += x;
  }
  if (c || f.length) { f.push(c); filas.push(f); } return filas;
}
const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

export async function render() {
  let archivos = []; // {nombre, buf, tipo}  tipo: undefined = sin revisar · null = no reconocido · '!…' = no se pudo abrir · clave de tipo
  const estado = h('div', { class: 'nota' }); const panel = h('div'); const lista = h('div');
  const resultados = {}; // un contenedor de resultados por tipo de documento
  const entrada = h('input', { type: 'file', multiple: true, accept: '.xlsx', style: { display: 'none' }, onchange: (e) => { añadir([...e.target.files]); e.target.value = ''; } });
  const zona = h('div', { class: 'tarj', style: { textAlign: 'center', borderStyle: 'dashed', padding: '26px' }, ondragover: (e) => { e.preventDefault(); }, ondrop: (e) => { e.preventDefault(); añadir([...e.dataTransfer.files]); } },
    h('div', { style: { fontSize: '2rem' } }, '📊'), h('p', h('b', 'Arrastra aquí los formatos diligenciados (.xlsx)'), ' o '), h('button', { class: 'btn', onclick: () => entrada.click() }, 'Elegir archivos'), entrada,
    h('p', { class: 'nota' }, 'Carga formatos del mismo tipo (por ejemplo, varios anteproyectos): solo se pueden extraer archivos de un mismo formato; si cargas archivos de formatos distintos, la aplicación te pide quitar los que no correspondan. Los archivos se leen en este dispositivo y no se envían a ningún servidor.'));

  async function añadir(fs) {
    const nuevos = fs.filter((x) => x.name.toLowerCase().endsWith('.xlsx') && !x.name.startsWith('~$'));
    for (const f of nuevos) if (!archivos.some((a) => a.nombre === f.name)) archivos.push({ nombre: f.name, buf: await leerComoBuffer(f) });
    if (fs.length > nuevos.length) toast('Solo se aceptan archivos .xlsx.', 'aviso');
    await revisar();
  }
  // Deja en /datos solo los archivos indicados (Python lee ese directorio completo)
  const preparar = async (py, l) => {
    await py.runPythonAsync('import shutil, os\nshutil.rmtree("/datos", ignore_errors=True)\nos.makedirs("/datos")');
    l.forEach((a) => py.FS.writeFile('/datos/' + a.nombre, new Uint8Array(a.buf)));
  };
  // Averigua de qué formato es cada archivo (clasificar() del script): los de un mismo formato tienen la misma estructura
  async function revisar() {
    pintar();
    const pend = archivos.filter((a) => a.tipo === undefined); if (!pend.length) return;
    try {
      const py = await python((m) => { estado.textContent = m; });
      estado.textContent = 'Revisando de qué formato es cada archivo…';
      await preparar(py, pend);
      const r = JSON.parse(await py.runPythonAsync('import json, os, extraccion_comun as ec\nps = ["/datos/" + n for n in sorted(os.listdir("/datos"))]\nt = ec.clasificar(ps)\nout = {}\nfor p in ps:\n    x = t[p]\n    out[os.path.basename(p)] = [x, ec.nombre_tipo(x)] if (x and not x.startswith("!")) else [x, None]\njson.dumps(out)'));
      pend.forEach((a) => { const x = r[a.nombre]; a.tipo = x === undefined ? null : x[0]; if (x && x[0] && x[1]) nombres[x[0]] = x[1]; });
      estado.textContent = '';
    } catch (e) { console.error(e); estado.textContent = ''; panel.replaceChildren(h('div', { class: 'alerta mal' }, 'No se pudo iniciar Python: ' + e.message)); }
    pintar();
  }

  function limpiarResultados() { Object.keys(resultados).forEach((t) => { if (!archivos.some((a) => a.tipo === t)) { resultados[t].remove(); delete resultados[t]; } }); }
  function quitar(nombres_) { archivos = archivos.filter((a) => !nombres_.includes(a.nombre)); limpiarResultados(); pintar(); }
  function dejarSolo(tipo) { archivos = archivos.filter((a) => a.tipo === tipo); limpiarResultados(); pintar(); }

  function pintar() {
    if (!archivos.length) { lista.replaceChildren(); panel.replaceChildren(); return; }
    const sin = archivos.filter((a) => a.tipo === undefined), nr = archivos.filter((a) => a.tipo === null || (typeof a.tipo === 'string' && a.tipo.startsWith('!')));
    const grupos = {}; archivos.filter((a) => typeof a.tipo === 'string' && !a.tipo.startsWith('!')).forEach((a) => { (grupos[a.tipo] ||= []).push(a); });
    const claves = Object.keys(grupos);
    const problema = claves.length > 1 || nr.length > 0;           // solo se pueden extraer archivos de un mismo formato
    const nombreTipo = (t) => nombres[t] || t;
    const esNuevo = (t) => t.startsWith('nuevo:');
    const fila = (a) => h('li', a.nombre + ' (' + tamaño(a.buf.byteLength) + ') ', h('a', { href: '#', onclick: (e) => { e.preventDefault(); quitar([a.nombre]); } }, 'quitar'));
    lista.replaceChildren(h('div', { class: 'fila', style: { margin: '8px 0' } }, h('b', archivos.length + ' archivo(s) cargados'),
      h('button', { class: 'btn gris peq', onclick: () => { archivos = []; Object.values(resultados).forEach((n) => n.remove()); Object.keys(resultados).forEach((k) => delete resultados[k]); pintar(); } }, 'Limpiar')));
    const alerta = problema && !sin.length ? h('div', { class: 'alerta mal' }, h('b', 'Solo se pueden extraer archivos de un mismo formato.'),
      h('p', { style: { margin: '4px 0 0' } }, claves.length > 1 ? `Cargaste archivos de ${claves.length} formatos distintos. Deja uno solo: quita los demás con los botones de abajo.` : 'Hay archivos que no son formatos FUP o que no se pudieron abrir: quítalos para poder extraer.')) : null;
    panel.replaceChildren(...[
      sin.length ? h('div', { class: 'nota' }, 'Revisando ' + sin.length + ' archivo(s)…') : null, alerta,
      ...claves.map((t) => {
        resultados[t] ||= h('div');
        return h('section', { class: 'tarj', style: { marginBottom: '10px' } },
          h('div', { class: 'fila' }, h('h3', { style: { margin: 0 } }, nombreTipo(t)), h('span', { class: 'chip' }, grupos[t].length + ' archivo(s)'), esNuevo(t) ? h('span', { class: 'chip aviso' }, 'Formato nuevo') : null),
          esNuevo(t) ? h('p', { class: 'nota' }, 'Este formato no está registrado: la aplicación lo lee por su estructura y lo agrupa solo con los archivos de su misma estructura.') : null,
          h('ul', { class: 'l' }, grupos[t].map(fila)),
          problema ? h('div', { class: 'fila' }, sin.length ? null : h('button', { class: 'btn', onclick: () => dejarSolo(t) }, 'Dejar solo este formato'),
              h('button', { class: 'btn gris peq', onclick: () => quitar(grupos[t].map((a) => a.nombre)) }, 'Quitar estos archivos'))
            : h('div', { class: 'fila' }, sin.length ? null : h('button', { class: 'btn', onclick: () => extraer(t) }, 'Extraer los datos'),
              h('button', { class: 'btn gris peq', onclick: () => quitar(grupos[t].map((a) => a.nombre)) }, 'Quitar estos archivos')),
          resultados[t]);
      }),
      nr.length ? h('div', { class: 'alerta aviso' }, h('b', 'No se pueden procesar:'), h('ul', { class: 'l' }, nr.map((a) => h('li', a.nombre + ': ' + (a.tipo === null ? 'no es un formato FUP reconocido (o está vacío).' : 'no se pudo abrir (' + a.tipo.slice(1) + ').') + ' ', h('a', { href: '#', onclick: (e) => { e.preventDefault(); quitar([a.nombre]); } }, 'quitar')))),
        h('button', { class: 'btn gris peq', onclick: () => quitar(nr.map((a) => a.nombre)) }, 'Quitar todos estos archivos')) : null].filter(Boolean));
  }

  async function extraer(tipo) {
    const cont = resultados[tipo]; if (!cont) return;
    cont.replaceChildren(h('p', { class: 'nota' }, 'Extrayendo datos…'));
    const t0 = performance.now();
    try {
      const py = await python((m) => { estado.textContent = m; }); estado.textContent = '';
      await preparar(py, archivos.filter((a) => a.tipo === tipo));
      py.globals.set('TIPO', tipo);
      py.setStdout({ batched: () => {} }); py.setStderr({ batched: () => {} });
      const r = JSON.parse(await py.runPythonAsync('import json, os, shutil, extraccion_comun as ec\nshutil.rmtree("/home/pyodide/salida", ignore_errors=True)\ntry:\n    n, rech = ec.ejecutar(TIPO, ["/datos"])\nexcept SystemExit as e:\n    n, rech = 0, [("(extracción)", str(e))]\nd = "/home/pyodide/salida/" + ec.carpeta_tipo(TIPO)\njson.dumps(dict(n=n, rech=[list(x) for x in rech], archivos=sorted(os.listdir(d)) if os.path.isdir(d) else [], dir=d + "/"))'));
      const dir = r.dir;
      const leer = (ext) => { const n = r.archivos.find((x) => x.endsWith(ext) && !x.startsWith('registro_')); return n ? { nombre: n, datos: py.FS.readFile(dir + n) } : null; };
      const x = leer('.xlsx'), c = leer('.csv'), reg = r.archivos.find((n) => n.startsWith('registro_')); const regTxt = reg ? new TextDecoder('utf-8').decode(py.FS.readFile(dir + reg)) : '';
      const seg = ((performance.now() - t0) / 1000).toFixed(1);
      const rech = r.rech.length ? h('div', { class: 'alerta aviso' }, h('b', 'No se procesaron ' + r.rech.length + ' archivo(s):'), h('ul', { class: 'l' }, r.rech.map(([n, m]) => h('li', n + ': ' + m)))) : null;
      if (!r.n || !x || !c) return void cont.replaceChildren(...[h('div', { class: 'alerta mal' }, 'No se procesó ningún archivo de este tipo.'), rech, regTxt ? h('details', h('summary', 'Ver el registro'), h('pre', { class: 'mer' }, regTxt)) : null].filter(Boolean));
      const filas = csv(new TextDecoder('utf-8').decode(c.datos).replace(/^﻿/, '')); const cab = filas[0] || []; const cuerpo = filas.slice(1);
      cont.replaceChildren(...[
        h('div', { class: 'alerta ok' }, `✓ Se leyeron ${r.n} archivo(s) y ${Math.max(0, cab.length - 4)} campos distintos (${seg} s).`), rech,
        h('div', { class: 'fila' }, h('button', { class: 'btn', onclick: () => descargar(x.nombre, new Blob([x.datos], { type: XLSX })) }, 'Descargar el Excel'),
          h('button', { class: 'btn sec', onclick: () => descargar(c.nombre, new Blob([c.datos], { type: 'text/csv;charset=utf-8' })) }, 'Descargar CSV'),
          reg ? h('button', { class: 'btn gris', onclick: () => descargar(reg, new Blob([regTxt], { type: 'text/plain;charset=utf-8' })) }, 'Descargar registro') : null),
        h('p', { class: 'nota' }, 'El Excel trae: Resumen, Ver proyecto, Base de datos, Datos (lista) y las tablas del formato (presupuesto, cronograma, control de cambios…). Vista previa de las primeras columnas:'),
        h('div', { class: 'tabla-w' }, h('table', { class: 't' }, h('thead', h('tr', cab.slice(0, 8).map((t) => h('th', t)))), h('tbody', cuerpo.map((f) => h('tr', f.slice(0, 8).map((t) => h('td', t.length > 70 ? t.slice(0, 68) + '…' : t)))))))].filter(Boolean));
    } catch (e) { console.error(e); estado.textContent = ''; cont.replaceChildren(h('div', { class: 'alerta mal' }, 'No se pudo extraer: ' + e.message)); }
  }

  return h('div',
    h('p', { class: 'nota' }, 'Reúne en un solo Excel los datos de varios formatos diligenciados del ', h('b', 'mismo tipo'), ' (por ejemplo, 20 anteproyectos o 15 actas de sustentación). Son los mismos scripts de Python del proyecto, ejecutados dentro de esta aplicación: la primera vez carga unos 14 MB y luego queda guardado, y funciona sin internet.'),
    zona, estado, lista, panel);
}
