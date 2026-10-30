// Cómo se usa la aplicación de extracción de datos.
import { h } from '../util.js';

export async function render() {
  const sec = (t, ...c) => h('section', { class: 'tarj', style: { marginBottom: '12px' } }, h('h3', t), ...c);
  const li = (a) => h('ul', { class: 'l' }, a.map((x) => h('li', x)));
  return h('div',
    sec('Para qué sirve', h('p', 'Lee los formatos Excel de opciones de grado que ya fueron diligenciados y reúne sus datos en un solo Excel (y un CSV), con una fila por documento. Sirve para tener en una tabla los datos de muchos anteproyectos, evaluaciones o actas, sin copiarlos a mano.')),
    sec('Pasos',
      li(['Abre la sección «Extracción de datos» y arrastra los archivos .xlsx diligenciados (o pulsa «Elegir archivos»). Todos deben ser del mismo formato.',
        'La aplicación revisa de qué formato es cada archivo: reconoce todos los formatos de las opciones de grado (unificados e individuales: 01, 02, 03, 04, 05, 07 a 11, 14 a 16, 18 a 22, 24, EC01 a EC03, M01, P01, P02, SI01, SI02, SN01, T01, T02, E01, X01 y X02).',
        'Si todos son del mismo formato, pulsa «Extraer los datos». Si hay archivos de formatos distintos, la aplicación muestra una alerta y no deja extraer hasta que quites los que no correspondan (botón «Dejar solo este formato»).',
        'Descarga el Excel, el CSV y el registro. El Excel nuevo no reemplaza ninguno anterior: cada extracción es un archivo aparte con la fecha y la hora en el nombre.'])),
    sec('Formatos nuevos', h('p', 'Si se crea un formato nuevo, la aplicación igual lo extrae: lo reconoce por su diseño (etiquetas, campos y bandas de sección) y lo agrupa con los archivos de su misma estructura. No hace falta registrarlo ni agregar nada.')),
    sec('Qué trae el Excel que se genera', li(['Resumen: lista de los formatos procesados, quién los presenta y cuánto está diligenciado.', 'Ver proyecto: elige un archivo y muestra solo sus datos.',
      'Base de datos: una fila por formato, con todos sus campos agrupados por sección.', 'Datos (lista): todos los campos en formato de lista, ideal para filtrar.', 'Tablas del formato (presupuesto, cronograma, control de cambios, etc.).'])),
    sec('Reglas que aplica la extracción', li(['Cada extracción es de UN solo formato: nunca se mezclan columnas de formatos distintos. Si hay archivos de otro formato, no se extrae nada hasta quitarlos.',
      'Un archivo que no es un formato FUP, que es de otro tipo o que tiene menos del 15 % de sus campos diligenciados se rechaza y se informa el motivo.',
      'Los archivos deben haberse guardado desde Excel (así traen los valores calculados) y con la opción de grado elegida en el filtro del formato.',
      'Si un archivo no es un formato FUP reconocido, la aplicación lo indica y no lo procesa.'])),
    sec('Privacidad', h('p', 'Los archivos se leen dentro de este dispositivo. No se suben a internet ni a ningún servidor, y la aplicación funciona sin conexión.')),
    h('p', { class: 'nota' }, 'Opciones de Grado · Ingeniería Industrial · FUP. Pasantía de Claudia Ximena Franco Velarde. Extracción de datos, versión 1.3.1.'));
}
