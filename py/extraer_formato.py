#!/usr/bin/env python3
"""
extraer_formato.py - Extrae los datos diligenciados de un formato Excel FUP (hoja "Formalizacion").

Funciona por ETIQUETAS y por el estilo de las celdas (no por direcciones fijas), asi que
sigue sirviendo si se agregan o mueven filas. Reconoce:
  - Campos "etiqueta | valor"  (etiqueta = celda azul claro, valor = celda crema)
  - Pares dobles               (ej. Codigo estudiantil | valor | Numero celular | valor)
  - Cuadros de texto largos    (celda crema grande bajo el titulo de la seccion)
  - Tablas                     (fila de encabezados azul claro + filas crema)
  - Subsecciones               (ej. DATOS ESTUDIANTE 1 / 2)

Uso en consola. Al ejecutarlo SIEMPRE guarda, sin pedirlo, resumen_extraccion.csv y
resumen_extraccion.xlsx en la misma carpeta de los formatos:
  python extraer_formato.py                  -> procesa todos los .xlsx de la carpeta del script
                                                (tambien sirve el boton Run/Play de Antigravity)
  python extraer_formato.py formato.xlsx     -> procesa solo ese formato
  python extraer_formato.py carpeta\\          -> procesa todos los .xlsx de esa carpeta
Opciones:
  --sin-xlsx          solo CSV, sin el Excel
  --csv NOMBRE.csv    otro nombre para el CSV        --xlsx NOMBRE.xlsx   otro nombre para el Excel
  -o datos.json       ademas guarda el JSON completo  --ver               muestra el JSON en pantalla
  -c "correo"         solo muestra los campos cuya etiqueta contenga ese texto (no guarda archivos)
  --calculados        incluye las celdas grises (formulas); el archivo debe haberse guardado en Excel

Uso como libreria:
  from extraer_formato import extraer, aplanar
  datos = extraer("formato.xlsx"); plano = aplanar(datos)
Requiere: pip install openpyxl
"""
import argparse, csv, datetime, glob, json, os, re, sys, warnings
import openpyxl

warnings.filterwarnings("ignore")
VERSION = "v11"
LABEL, INPUT, CALC, BAND, SUB = "DEEAF6", "FFF8DC", "F2F2F2", "005790", "003A5D"


def _rgb(c):
    try:
        v = c.fill.fgColor.rgb
        return v[-6:].upper() if isinstance(v, str) else ""
    except Exception:
        return ""


def _val(v):
    if isinstance(v, str):
        v = v.strip()
        return v or None          # " " de las celdas con flecha ▼ = vacio
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.date().isoformat() if isinstance(v, datetime.datetime) else v.isoformat()
    return v


def _maps(ws):
    """celda -> rango combinado (tope izquierdo) y conjunto de celdas ocultas por la combinacion"""
    top, hidden = {}, set()
    for r in ws.merged_cells.ranges:
        top[(r.min_row, r.min_col)] = r
        for rr in range(r.min_row, r.max_row + 1):
            for cc in range(r.min_col, r.max_col + 1):
                if (rr, cc) != (r.min_row, r.min_col):
                    hidden.add((rr, cc))
    return top, hidden


def _clean(t):
    return re.sub(r"\s+", " ", str(t)).strip()


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _peso(valor, formula, wbf):
    """Peso de una seccion: valor guardado o, si el archivo no se guardo en Excel, el de la hoja Listas."""
    v = _num(valor)
    if v is not None:
        return v
    if isinstance(formula, str):
        m = re.fullmatch(r"=Listas!\$?([A-Z]+)\$?(\d+)", formula.strip())
        if m and wbf and "Listas" in wbf.sheetnames:
            return _num(wbf["Listas"][m.group(1) + m.group(2)].value)
    return _num(formula)


def _umbrales(wbf):
    try:
        li = wbf["Listas"]
        return {"aprobado": _num(li["B9"].value), "correcciones": _num(li["B10"].value)}
    except Exception:
        return {}


def _con_puntaje(filas, heads):
    """Convierte las columnas 1..5 en 'Puntaje' y las columnas Si/No en 'Respuesta'."""
    if all(h in heads for h in ("1", "2", "3", "4", "5")):
        nuevas = []
        for f in filas:
            marcas = [h for h in ("1", "2", "3", "4", "5") if f.get(h) not in (None, "") and str(f.get(h)).strip().upper() == "X"]
            p = int(marcas[0]) if len(marcas) == 1 else ("(varias marcas)" if len(marcas) > 1 else None)
            n = {h: f[h] for h in heads if h not in ("1", "2", "3", "4", "5")}
            n["Puntaje"] = p
            nuevas.append(n)
        return nuevas
    if "Sí" in heads and "No" in heads:
        nuevas = []
        for f in filas:
            si = str(f.get("Sí") or "").strip().upper() == "X"; no = str(f.get("No") or "").strip().upper() == "X"
            n = {h: f[h] for h in heads if h not in ("Sí", "No")}
            n["Respuesta"] = "Sí" if si and not no else ("No" if no and not si else ("(varias marcas)" if si and no else None))
            nuevas.append(n)
        # columnas en orden: etiqueta, Respuesta, Observacion
        return [{k: n[k] for k in sorted(n, key=lambda x: (x == "Observación", x == "Respuesta"))} for n in nuevas]
    return filas


def _resultado(out):
    """Promedios por seccion, nota final y concepto calculados por el script (no dependen de que el archivo se guarde en Excel)."""
    secs = out["secciones"]; calc = {}
    # requisitos
    for s, cuerpo in secs.items():
        t = cuerpo.get("tabla") if isinstance(cuerpo, dict) else None
        if isinstance(t, list) and t and "Respuesta" in t[0]:
            calc["Requisitos cumplidos (Sí)"] = f"{sum(1 for x in t if x.get('Respuesta') == 'Sí')} de {len(t)}"
            calc["Requisitos con respuesta"] = f"{sum(1 for x in t if x.get('Respuesta'))} de {len(t)}"
    # promedios
    proms = {}
    for s, cuerpo in secs.items():
        if not isinstance(cuerpo, dict): continue
        for sub, v in cuerpo.items():
            if isinstance(v, dict) and isinstance(v.get("tabla"), list) and v["tabla"] and "Puntaje" in v["tabla"][0]:
                p = [x["Puntaje"] for x in v["tabla"] if isinstance(x.get("Puntaje"), int)]
                letra = sub.strip()[:1]
                proms[letra] = (sub, round(sum(p) / len(p), 2) if p else None, len(p), len(v["tabla"]))
    for letra, (sub, pr, n, tot) in proms.items():
        calc[f"Promedio sección {letra}"] = pr
        calc[f"Criterios calificados sección {letra}"] = f"{n} de {tot}"
    pesos = {lab[:1]: p for lab, p in out.get("_pesos", [])}
    if proms and pesos and all(l in pesos and pesos[l] is not None for l in proms) and all(v[1] is not None for v in proms.values()):
        aportes = [round(proms[l][1] * pesos[l], 3) for l in proms]
        nota = round(sum(aportes), 2)
        calc["Nota final (calculada)"] = nota
        u = out.get("_listas", {})
        if u.get("aprobado") is not None and u.get("correcciones") is not None:
            calc["Concepto (calculado)"] = "Aprobado" if nota >= u["aprobado"] else ("Aprobado con correcciones" if nota >= u["correcciones"] else "Rechazado")
    # la nota final y el concepto tambien se muestran en la tabla de la seccion "NOTA FINAL Y CONCEPTO"
    for s, cuerpo in secs.items():
        if "NOTA FINAL" in s.upper() and isinstance(cuerpo, dict) and isinstance(cuerpo.get("tabla"), list):
            if not cuerpo["tabla"] and calc.get("Nota final (calculada)") is not None:
                cuerpo["tabla"].append({})          # sin evaluador/firma diligenciados: igual se muestra la nota
            for fila_i, fila in enumerate(cuerpo["tabla"]):
                if fila_i: continue
                nuevo = {"Nota final": calc.get("Nota final (calculada)"), "Concepto": calc.get("Concepto (calculado)")}
                nuevo.update({k: v for k, v in fila.items() if k not in nuevo})
                fila.clear(); fila.update(nuevo)
    if calc:
        secs["RESULTADO CALCULADO POR EL SCRIPT"] = calc


def extraer(ruta, hoja="Formalizacion", calculados=False):
    wb = openpyxl.load_workbook(ruta, data_only=True)
    try:
        wbf = openpyxl.load_workbook(ruta)          # con formulas: sirve para leer los pesos de la hoja Listas
    except Exception:
        wbf = None
    if hoja not in wb.sheetnames:
        # formatos unificados: la hoja del formato tiene otro nombre
        hoja = next((n for n in ("Formalizacion", "Anteproyecto", "Evaluacion_Anteproyecto", "Evaluacion_Documento_Final") if n in wb.sheetnames), None)
    if not hoja:
        # formatos nuevos (05, 07, 09, 13 a 24 e individuales): la hoja del formato es la primera que no es auxiliar
        hoja = next((n for n in wb.sheetnames if n not in ("Instrucciones", "Imagenes", "Listas")), None)
    ws = wb[hoja] if hoja else wb.worksheets[0]
    wsf = wbf[ws.title] if wbf else None
    top, hidden = _maps(ws)
    out = {"_archivo": os.path.basename(ruta), "_hoja": ws.title, "titulo": None,
           "secciones": {}, "calculados": {}, "_pesos": []}
    sec = sub = None
    # formatos unificados: la columna A (angosta) solo guarda las etiquetas del filtro; el contenido va de B a G
    tagcol = (ws.column_dimensions["A"].width or 99) < 8
    c0 = 2 if tagcol else 1
    maxc = 8 if tagcol else min(ws.max_column, 6)
    colsocultas = set()
    for _dim in ws.column_dimensions.values():
        if _dim.hidden and _dim.min:
            colsocultas.update(range(_dim.min, (_dim.max or _dim.min) + 1))
    if tagcol:
        t1 = next((_clean(ws.cell(1, c).value) for c in range(2, 8) if ws.cell(1, c).value), None)
        t2 = next((_clean(ws.cell(2, c).value) for c in range(2, 8) if ws.cell(2, c).value), None)
        out["opcion"] = t1

    def celdas(r):
        """celdas visibles de la fila: [(col, celda, rango o None)]"""
        res = []
        for c in range(c0, maxc + 1):
            if (r, c) in hidden or c in colsocultas:
                continue
            res.append((c, ws.cell(r, c), top.get((r, c))))
        return res

    def destino():
        d = out["secciones"].setdefault(sec or "(general)", {})
        return d.setdefault(sub, {}) if sub else d

    r = 1
    while r <= ws.max_row:
        if ws.row_dimensions[r].hidden:      # filas ocultas por el filtro (otras opciones de grado)
            r += 1; continue
        cs = celdas(r)
        fills = [_rgb(c) for _, c, _ in cs]
        textos = [(c, ce, rg) for c, ce, rg in cs if ce.value not in (None, "")]
        # titulo del formato
        if out["titulo"] is None and textos and r <= 3 and (not tagcol or r == 2) and all(_rgb(ce) not in (LABEL, INPUT) for _, ce, _ in textos):
            out["titulo"] = _clean(textos[0][1].value)
        # banda de seccion
        if BAND in fills and textos:
            sec, sub = _clean(textos[0][1].value), None
            r += 1; continue
        if SUB in fills and textos:
            sub = _clean(textos[0][1].value)
            r += 1; continue
        # panel de verificacion (celdas grises con etiqueta)
        if calculados and CALC in fills:
            lab = next((ce for _, ce, _ in cs if _rgb(ce) == LABEL and ce.value), None)
            if lab:
                v = [ _val(ce.value) for _, ce, _ in cs if _rgb(ce) == CALC and ce.value is not None]
                out["calculados"][_clean(lab.value)] = v[0] if len(v) == 1 else v
        # fila de encabezados de tabla: 3 o mas celdas-etiqueta
        labs = [(c, ce, rg) for c, ce, rg in cs if _rgb(ce) == LABEL and (ce.value not in (None, "") or (wsf is not None and isinstance(wsf.cell(r, c).value, str) and wsf.cell(r, c).value.startswith("=")))]
        if len(labs) >= 3:
            heads = [(c, _clean(ce.value)) for c, ce, _ in labs]
            # una celda de encabezado puede abarcar varias columnas (A:C, B:C, D:E)
            nombres_h = [h for _, h in heads]
            if "Peso" in nombres_h and nombres_h[0].lower().startswith("secci"):
                cp = next(c for c, h in heads if h == "Peso")
                rr = r + 1
                while rr <= ws.max_row and _rgb(ws.cell(rr, heads[0][0])) in (LABEL, CALC, "FFFFFF", "") and ws.cell(rr, heads[0][0]).value and re.match(r"^[A-Z]\.", str(ws.cell(rr, heads[0][0]).value).strip()):
                    peso = _peso(ws.cell(rr, cp).value, wsf.cell(rr, cp).value if wsf else None, wbf)
                    out["_pesos"].append((_clean(ws.cell(rr, heads[0][0]).value), peso))
                    rr += 1
                out["_listas"] = _umbrales(wbf)
                r = rr; continue
            if "Marcar" in nombres_h:
                r += 1
                while r <= ws.max_row and ws.cell(r, heads[0][0]).value not in (None, "") and _rgb(ws.cell(r, heads[-1][0])) == CALC: r += 1
                continue
            filas, rr = [], r + 1
            while rr <= ws.max_row:
                cc = celdas(rr)
                if not any(_rgb(ce) == INPUT for _, ce, _ in cc):
                    break
                # una fila con una etiqueta en otra columna que no es la primera de la tabla ya es un campo "etiqueta | valor" (p. ej. duracion de la exposicion bajo una lista de requisitos)
                if any(_rgb(ce) == LABEL and ce.value not in (None, "") and c != heads[0][0] for c, ce, _ in cc):
                    break
                reg = {}
                for c, h in heads:
                    ce = ws.cell(rr, c)
                    reg[h] = _val(ce.value)
                if any(v is not None for v in reg.values()):
                    filas.append(reg)
                rr += 1
            filas = _con_puntaje(filas, nombres_h)
            d = destino()
            k, n = "tabla", 2
            while k in d:
                k = f"tabla_{n}"; n += 1
            d[k] = filas
            r = rr; continue
        # campos etiqueta | valor (hasta dos pares por fila)
        if labs:
            i = 0; primera = None
            while i < len(cs):
                c, ce, rg = cs[i]
                _eti_formula = (_rgb(ce) == LABEL and ce.value in (None, "") and wsf is not None and isinstance(wsf.cell(r, c).value, str) and wsf.cell(r, c).value.startswith("="))
                if _eti_formula and i + 1 < len(cs) and _rgb(cs[i + 1][1]) == INPUT and cs[i + 1][1].value not in (None, ""):
                    destino()["Dato específico de la opción de grado"] = _val(cs[i + 1][1].value)
                    i += 2; continue
                if _rgb(ce) == LABEL and ce.value not in (None, ""):
                    j = i + 1
                    if j < len(cs) and (_rgb(cs[j][1]) == INPUT or (calculados and _rgb(cs[j][1]) == CALC)):
                        if _rgb(cs[j][1]) == INPUT or calculados:
                            clave = _clean(ce.value)
                            if primera and re.match(r"^(Autor|Estudiante)", primera):
                                clave = f"{clave} ({re.sub(chr(92)+'s*'+chr(92)+'(si aplica'+chr(92)+')', '', primera)})"        # p. ej. 'Código estudiantil (Autor 1)'
                            elif primera is None:
                                primera = clave
                            destino()[clave] = _val(cs[j][1].value)
                        if primera is None: primera = _clean(ce.value)
                        i = j
                i += 1
            r += 1; continue
        # cuadro de texto largo: celda crema (sin etiqueta) combinada o suelta
        for c, ce, rg in cs:
            if _rgb(ce) == INPUT and (c == c0):
                nombre = "texto" if not sub else sub
                d = destino()
                k = nombre
                n = 2
                while k in d:
                    k = f"{nombre}_{n}"; n += 1
                d[k] = _val(ce.value)
        r += rg.max_row - rg.min_row + 1 if (cs and cs[0][2] is not None) else 1
    _resultado(out)
    return out



def _recorrer(d, sec, sub=None):
    """Genera (seccion, subseccion, campo, valor) y tablas (seccion, nombre, filas)."""
    campos, tablas = [], []
    for k, v in d.items():
        if isinstance(v, dict):
            c2, t2 = _recorrer(v, sec, k)
            campos += c2; tablas += t2
        elif isinstance(v, list):
            tablas.append((sec, sub, v))
        else:
            campos.append((sec, sub, k, v))
    return campos, tablas


def _wcl(w):  # caracteres aprox. por linea para un ancho de columna
    return max(8, int(w * 1.15))


def _lineas(txt, ancho):
    if txt is None:
        return 1
    return sum(max(1, -(-len(p) // _wcl(ancho))) for p in str(txt).split("\n"))


_CANON = [   # secciones que son lo mismo en las distintas opciones de grado aunque cambie su nombre exacto (patron sin tildes y en mayusculas -> nombre unico)
    (r"^FORMULACION DEL PROBLEMA|^PLANTEAMIENTO DE LA PRACTICA|^OPORTUNIDAD Y PROBLEMA", "FORMULACIÓN DEL PROBLEMA"),
    (r"^FUNDAMENTACION TEORICA|^MARCO DE REFERENCIA|^MARCO TEORICO", "FUNDAMENTACIÓN TEÓRICA O MARCO DE REFERENCIA"),
    (r"^DISENO METODOLOGICO|^ESTRATEGIA DE REVISION|^METODOLOGIA", "METODOLOGÍA"),
    (r"^REVISION PRELIMINAR DE LA BIBLIOGRAFIA|^BIBLIOGRAFIA INICIAL", "REVISIÓN PRELIMINAR DE LA BIBLIOGRAFÍA"),
    (r"^RECURSOS", "RECURSOS DISPONIBLES"),
    (r"^EMPRESA", "EMPRESA U ORGANIZACIÓN"),
]


def canon_sec(s):
    """Nombre unico de una seccion: sin el numero de orden (que cambia de una opcion de grado a otra), sin el limite de palabras o de objetivos,
    y con el mismo nombre las secciones equivalentes (_CANON). Asi la misma seccion de varios formatos cae en la MISMA columna."""
    if not s:
        return s
    t = re.sub(r"^\s*\d+\.\s*", "", s)
    t = re.sub(r"\s*\((\d+ palabras|m[aá]ximo \d+ espec[ií]ficos|\d+ espec[ií]ficos)\)", "", t).strip()
    import unicodedata
    u = "".join(ch for ch in unicodedata.normalize("NFD", t.upper().replace("Ñ", "N")) if unicodedata.category(ch) != "Mn")
    for patron, nombre in _CANON:
        if re.match(patron, u):
            return nombre
    return t


def _limpia_sec(s):
    s = re.sub(r"^\d+\.\s*", "", s or "General")
    s = re.sub(r"\s*\(\d+ palabras\)|\s*\(m[aá]ximo \d+ espec[ií]ficos\)", "", s)
    return s.strip().capitalize()


def _spans(n, total=5):
    """reparte 'total' columnas de la cuadricula entre n columnas de una tabla"""
    n = min(n, total)
    sp = [1] * n
    extra = total - n
    i = 1 if n > 1 else 0
    while extra > 0:
        sp[i % n] += 1; extra -= 1; i += 1 if n > 2 else 0
    return sp


def a_excel(todos, ruta):
    """Excel presentable: Resumen (portada + indice), Base de datos (una fila por formato, encabezados agrupados) y una hoja por tabla.
    No incluye una hoja de ficha por formato: los formatos originales ya existen y no hace falta repetirlos aqui."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter as L
    from openpyxl.formatting.rule import DataBarRule
    AZ, OSC, CL, ZEB, GR = "005790", "003A5D", "DEEAF6", "F4F8FC", "F2F2F2"
    fill = lambda c: PatternFill("solid", fgColor=c)
    fnt = lambda **k: Font(name="Arial", size=k.pop("size", 10), **k)
    lado = Side(style="thin", color="B7C9DA"); borde = Border(lado, lado, lado, lado)
    ahora = datetime.datetime.now()

    def conv(v):
        if isinstance(v, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
            return datetime.date.fromisoformat(v)
        return v

    def fmt(c, v, etiqueta=""):
        if isinstance(v, datetime.date): c.number_format = "dd/mm/yyyy"
        elif isinstance(v, float) and 0 <= v <= 1 and ("%" in etiqueta or "porcentaje" in etiqueta.lower()): c.number_format = "0%"
        elif isinstance(v, float): c.number_format = "#,##0.##"

    def banda(ws, r, c1, c2, texto, color, size=11, alto=22, alin="left"):
        ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        c = ws.cell(r, c1, texto); c.fill = fill(color); c.font = fnt(bold=True, color="FFFFFF", size=size)
        c.alignment = Alignment(horizontal=alin, vertical="center", indent=1)
        for j in range(c1, c2 + 1): ws.cell(r, j).fill = fill(color)
        ws.row_dimensions[r].height = alto

    def pagina(ws, horizontal=False, titulos=None):
        ws.sheet_view.showGridLines = False
        ws.page_setup.orientation = "landscape" if horizontal else "portrait"
        ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
        ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins.left = ws.page_margins.right = 0.5
        ws.oddFooter.left.text = "Extracción de datos de formatos FUP"; ws.oddFooter.right.text = "Página &P de &N"
        if titulos: ws.print_title_rows = titulos

    # ---------- analisis de cada formato
    info = []
    for n, d in enumerate(todos, 1):
        campos, tablas = [], []
        for sec_, cuerpo in d["secciones"].items():
            c_, t_ = _recorrer(cuerpo, sec_); campos += c_; tablas += t_
        llenos = sum(1 for *_, v in campos if v not in (None, "")) + sum(1 for *_, r in tablas if r)
        total = len(campos) + len(tablas)
        nombres = [v for s_, sub, k, v in campos if "nombre" in k.lower() and v]
        director = next((v for s_, sub, k, v in campos if "director" in k.lower() and v), "")
        autores = [v for s_, sub, k, v in campos if k.lower().startswith("autor") and isinstance(v, str) and v]
        if autores:   # formatos de evaluacion: los estudiantes son los autores (no el evaluador)
            nombres = autores
        t_campo = next((v for s_, sub, k, v in campos if k.lower().startswith("título del") and v), None)
        titulo = t_campo or next((v for s_, sub, k, v in campos if "TÍTULO" in (s_ or "").upper() or "TITULO" in (s_ or "").upper()), None) or d.get("titulo") or ""
        info.append(dict(n=n, d=d, campos=campos, tablas=tablas, pct=(llenos / total if total else 0), llenos=llenos,
                         total=total, nombres=[x for x in nombres if isinstance(x, str)][:2], director=director,
                         titulo=titulo))
    wb = Workbook(); ws0 = wb.active; ws0.title = "Resumen"

    # ---------- hoja RESUMEN (portada + indice)
    anchos = [6, 64, 58, 36, 30, 16]
    for j, w in enumerate(anchos, 1): ws0.column_dimensions[L(j)].width = w
    banda(ws0, 1, 1, 6, "RESUMEN DE EXTRACCIÓN DE FORMATOS", AZ, 16, 40)
    ws0.merge_cells("A2:F2")
    c = ws0["A2"]; c.value = f"Fundación Universitaria de Popayán · Generado el {ahora:%d/%m/%Y a las %H:%M}"
    c.font = fnt(italic=True, color="595959"); c.alignment = Alignment(indent=1, vertical="center"); ws0.row_dimensions[2].height = 20
    prom = sum(i["pct"] for i in info) / len(info)
    for col, (tit, val, f) in zip((2, 3), (("Formatos procesados", len(info), "0"), ("Diligenciamiento promedio", prom, "0%"))):
        t = ws0.cell(4, col, tit); t.font = fnt(bold=True, color="FFFFFF"); t.fill = fill(OSC); t.alignment = Alignment(horizontal="center", vertical="center")
        v = ws0.cell(5, col, val); v.font = fnt(bold=True, size=20, color=AZ); v.fill = fill(CL)
        v.alignment = Alignment(horizontal="center", vertical="center"); v.number_format = f
    ws0.row_dimensions[4].height = 20; ws0.row_dimensions[5].height = 38
    for j, t in enumerate(["N°", "Archivo", "Trabajo / título", "Estudiante(s)", "Director", "Diligenciado"], 1):
        c = ws0.cell(7, j, t); c.fill = fill(AZ); c.font = fnt(bold=True, color="FFFFFF"); c.border = borde
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws0.row_dimensions[7].height = 24
    r = 8
    for i in info:
        vals = [i["n"], i["d"]["_archivo"], i["titulo"], "\n".join(i["nombres"]), i["director"], i["pct"]]
        for j, v in enumerate(vals, 1):
            c = ws0.cell(r, j, v); c.border = borde; c.font = fnt()
            c.fill = fill(ZEB if r % 2 else "FFFFFF")
            c.alignment = Alignment(vertical="center", wrap_text=True, horizontal="center" if j in (1, 6) else "left", indent=0 if j in (1, 6) else 1)
        ws0.cell(r, 6).number_format = "0%"
        ws0.row_dimensions[r].height = max(30, 14 * max(_lineas(i["titulo"], anchos[2]), len(i["nombres"]) or 1, _lineas(i["d"]["_archivo"], anchos[1])))
        r += 1
    ws0.conditional_formatting.add(f"F8:F{r-1}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="7FB2D9"))
    ws0.cell(r + 1, 2, "La hoja «Base de datos» reúne todos los formatos (con filtros) y las hojas «BD …» tienen una tabla completa por cada tipo de formato.").font = fnt(italic=True, color="7F7F7F", size=9)
    ws0.auto_filter.ref = f"A7:F{r-1}"
    ws0.freeze_panes = "A8"; ws0.sheet_properties.tabColor = AZ; pagina(ws0, True)

    # ---------- BASE DE DATOS (encabezados agrupados) + una por tipo de formato + datos en lista
    def hoja_base(nombre, items, titulo, alias=True, formato=False):
        wsd = wb.create_sheet(nombre[:31]); wsd.sheet_properties.tabColor = OSC
        cols, filas = [], []
        for it in items:
            f = {}
            if formato:
                f[("FORMATO", "Formato")] = _tipo(it["d"])
                if ("FORMATO", "Formato") not in cols: cols.append(("FORMATO", "Formato"))
            for (s_, nom), v in _columnas(it["d"], alias):
                key = (s_, nom); n = 2
                while key in f:
                    key = (s_, f"{nom} [{n}]"); n += 1
                f[key] = conv(v)
                if key not in cols: cols.append(key)
            filas.append(f)
        nC = len(cols) + 1
        banda(wsd, 1, 1, max(nC, 4), titulo, AZ, 14, 32)
        wsd.cell(3, 1, "Archivo")
        for j, (s_, nom) in enumerate(cols, 2): wsd.cell(3, j, nom)
        j = 2                                                  # bandas de seccion agrupadas
        while j <= nC:
            s_ = cols[j - 2][0]; k = j
            while k + 1 <= nC and cols[k - 1][0] == s_: k += 1
            banda(wsd, 2, j, k, _limpia_sec(s_).upper(), OSC, 9, 22, "center"); j = k + 1
        wsd.cell(2, 1).fill = fill(OSC)
        for j in range(1, nC + 1):
            c = wsd.cell(3, j); c.fill = fill(CL); c.font = fnt(bold=True, size=9); c.border = borde
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            largo = j > 1 and any(isinstance(f.get(cols[j - 2]), str) and len(f.get(cols[j - 2])) > 80 for f in filas)
            wsd.column_dimensions[L(j)].width = 34 if j == 1 else (46 if largo else 24)
        wsd.row_dimensions[3].height = 52
        for n_, (it, f) in enumerate(zip(items, filas)):
            r = 4 + n_
            c = wsd.cell(r, 1, it["d"]["_archivo"]); c.font = fnt(bold=True)
            for j, key in enumerate(cols, 2):
                na = key not in f                      # el campo no existe en el formato de este proyecto
                v = "(No aplica)" if na else f.get(key); c = wsd.cell(r, j, v); c.font = fnt(italic=True, color="8A97A5") if na else fnt(); fmt(c, v, key[1])
                c.alignment = Alignment(vertical="top", wrap_text=False, horizontal="left")
            for j in range(1, nC + 1):
                wsd.cell(r, j).border = borde; wsd.cell(r, j).fill = fill(ZEB if n_ % 2 else "FFFFFF")
        wsd.freeze_panes = "B4"; wsd.auto_filter.ref = f"A3:{L(nC)}{3 + len(items)}"; pagina(wsd, True, "1:3")
        return wsd

    hoja_base("Base de datos", info, "BASE DE DATOS DE FORMATOS", True, True)
    # una hoja por tipo de formato: solo las columnas de ese documento (sin mezclar datos de otros formatos)
    grupos = {}
    for it in info: grupos.setdefault(_tipo(it["d"]), []).append(it)
    if len(grupos) > 1:
        usados = set()
        for tipo, its in grupos.items():
            nom_ = re.sub(r"[\[\]\\/*?:]", "", f"BD {tipo}")[:31]
            while nom_ in usados: nom_ = nom_[:28] + "_2"
            usados.add(nom_)
            hoja_base(nom_, its, f"BASE DE DATOS · {tipo.upper()} ({len(its)} formato(s))", False, False)
    # ---------- DATOS (LISTA): un renglon por dato; filtre por Archivo y vera solo lo de ese proyecto
    wl = wb.create_sheet("Datos (lista)"); wl.sheet_properties.tabColor = "A6C8E3"
    banda(wl, 1, 1, 6, "TODOS LOS DATOS DE CADA PROYECTO (filtre por Archivo, Formato o Sección)", AZ, 14, 32)
    for j, t in enumerate(["Archivo", "Formato", "Sección", "Campo", "Valor", "Fila de tabla"], 1):
        c = wl.cell(2, j, t); c.fill = fill(CL); c.font = fnt(bold=True, size=9); c.border = borde
        c.alignment = Alignment(horizontal="center", vertical="center")
    rr_ = 3
    for it in info:
        for sec_, campo_, val_, fila_ in _registros(it["d"]):
            v_ = conv(val_)
            for j, x in enumerate([it["d"]["_archivo"], _tipo(it["d"]), _limpia_sec(sec_), campo_, v_, fila_], 1):
                c = wl.cell(rr_, j, x); c.font = fnt(size=9); c.border = borde
                c.alignment = Alignment(vertical="top", wrap_text=(j == 5), horizontal="left")
                if j == 5: fmt(c, v_, campo_)
            rr_ += 1
    for j, w in enumerate([36, 30, 36, 58, 72, 12], 1): wl.column_dimensions[L(j)].width = w
    wl.freeze_panes = "A3"; wl.auto_filter.ref = f"A2:F{rr_ - 1}"; pagina(wl, True, "1:2")
    # ---------- columnas auxiliares (ocultas) de "Datos (lista)" para la hoja "Ver proyecto"
    ult = max(rr_ - 1, 3)
    wl.cell(2, 7, "Acum."); wl.cell(2, 8, "Valor (vista)")
    for r_ in range(3, rr_):
        wl.cell(r_, 7, f"=N(G{r_-1})+(A{r_}='Ver proyecto'!$B$3)")
        v_ = wl.cell(r_, 5).value
        if isinstance(v_, datetime.date): v_ = v_.strftime("%d/%m/%Y")
        elif isinstance(v_, float) and 0 <= v_ <= 1 and ("%" in str(wl.cell(r_, 4).value) or "porcentaje" in str(wl.cell(r_, 4).value).lower()): v_ = f"{v_:.0%}"
        wl.cell(r_, 8, v_)
    wl.column_dimensions["G"].hidden = True; wl.column_dimensions["H"].hidden = True
    # ---------- hoja VER PROYECTO: elija un proyecto y vera SOLO sus datos (sin casillas de otros proyectos)
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.formatting.rule import FormulaRule
    wv = wb.create_sheet("Ver proyecto", 1); wv.sheet_properties.tabColor = "00B4A5"
    maxreg = max([len(_registros(it["d"])) for it in info] + [1]) + 10
    for j, w in enumerate([36, 52, 80, 12], 1): wv.column_dimensions[L(j)].width = w
    banda(wv, 1, 1, 4, "VER UN PROYECTO (solo aparecen los datos del proyecto elegido)", AZ, 14, 32)
    a = wv.cell(3, 1, "Elija el proyecto (archivo):"); a.font = fnt(bold=True)
    wv.merge_cells("B3:C3"); sel = wv["B3"]; sel.value = info[0]["d"]["_archivo"]
    sel.fill = fill("FFF2CC"); sel.font = fnt(bold=True, size=11); sel.border = borde
    wv.row_dimensions[3].height = 24
    wv.cell(4, 1, "Datos encontrados:").font = fnt(bold=True)
    wv["B4"] = f"=MAX('Datos (lista)'!$G$3:$G${ult})"; wv["B4"].alignment = Alignment(horizontal="left")
    for j, t in enumerate(["Sección", "Campo", "Valor", "Fila de tabla"], 1):
        c = wv.cell(6, j, t); c.fill = fill(OSC); c.font = fnt(bold=True, color="FFFFFF", size=10); c.border = borde
        c.alignment = Alignment(horizontal="center", vertical="center")
    for k, it in enumerate(info):
        wv.cell(7 + k, 7, it["d"]["_archivo"])
    wv.column_dimensions["E"].hidden = True; wv.column_dimensions["G"].hidden = True
    dv = DataValidation(type="list", formula1=f"=$G$7:$G${6 + len(info)}", allow_blank=False)
    dv.error = "Elija un proyecto de la lista"; wv.add_data_validation(dv); dv.add("B3")
    rng = lambda col: f"'Datos (lista)'!${col}$3:${col}${ult}"
    for r_ in range(7, 7 + maxreg):
        wv.cell(r_, 5, f'=IF(ROW()-6<=$B$4,MATCH(ROW()-6,{rng("G")},0),"")')
        for j, col in ((1, "C"), (2, "D"), (3, "H"), (4, "F")):
            c = wv.cell(r_, j, f'=IF($E{r_}="","",IF(INDEX({rng(col)},$E{r_})="","",INDEX({rng(col)},$E{r_})))')
            c.font = fnt(size=9); c.alignment = Alignment(vertical="top", wrap_text=True, horizontal="left")
    wv.conditional_formatting.add(f"A7:D{6 + maxreg}", FormulaRule(formula=['$E7<>""'], border=borde, fill=fill("F4F8FC")))
    wv.freeze_panes = "A7"; pagina(wv, True, "6:6")
    # ---------- una hoja por tabla
    por_tabla = {}
    for i in info:
        for s_, sub, rows in i["tablas"]:
            por_tabla.setdefault(canon_sec(s_), []).extend({"Archivo": i["d"]["_archivo"], **({"Sección": sub} if sub else {}), **x} for x in rows)
    for s_, rows in por_tabla.items():
        if not rows:
            continue
        nombre = re.sub(r"[\[\]\\/*?:]", "", _limpia_sec(s_))[:22].strip() + " (tabla)"
        ws = wb.create_sheet(nombre); ws.sheet_properties.tabColor = "A6C8E3"
        cT = list(dict.fromkeys(k for x in rows for k in x))
        banda(ws, 1, 1, len(cT), _limpia_sec(s_).upper(), AZ, 12, 28)
        for j, t in enumerate(cT, 1):
            c = ws.cell(2, j, t); c.fill = fill(CL); c.font = fnt(bold=True, size=9); c.border = borde
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            w = max([len(str(conv(x.get(t)))) if not isinstance(conv(x.get(t)), datetime.date) else 11 for x in rows] + [len(t) * 0.8])
            ws.column_dimensions[L(j)].width = max(14, min(55, w + 3))
        ws.row_dimensions[2].height = 32
        for n_, x in enumerate(rows):
            r = 3 + n_
            for j, t in enumerate(cT, 1):
                v = conv(x.get(t)); c = ws.cell(r, j, v); c.font = fnt(); c.border = borde; fmt(c, v, t)
                c.fill = fill(ZEB if n_ % 2 else "FFFFFF"); c.alignment = Alignment(vertical="top", wrap_text=True, horizontal="left")
            ws.row_dimensions[r].height = max(18, 13 * max(_lineas(x.get(t), ws.column_dimensions[L(j)].width) for j, t in enumerate(cT, 1)) + 4)
        ws.freeze_panes = "A3"; ws.auto_filter.ref = f"A2:{L(len(cT))}{2 + len(rows)}"; pagina(ws, True, "1:2")
    wb.save(ruta)


def aplanar(datos):
    """Convierte el resultado en un diccionario plano 'Seccion | Subseccion | Campo' -> valor.
    Las tablas se expanden como 'Seccion | tabla | [n] | Columna'."""
    plano = {"archivo": datos.get("_archivo")}

    def rec(prefijo, v):
        if isinstance(v, dict):
            for k, x in v.items():
                rec(prefijo + [k], x)
        elif isinstance(v, list):
            for n, fila in enumerate(v, 1):
                for k, x in fila.items():
                    plano[" | ".join(prefijo + [f"[{n}]", k])] = x
        else:
            plano[" | ".join(prefijo)] = v
    rec([], datos["secciones"])
    for k, v in datos.get("calculados", {}).items():
        plano["Calculado | " + k] = v
    return plano


def buscar(datos, etiqueta):
    """Todas las coincidencias (sin mayusculas/tildes) de una etiqueta -> lista de (ruta, valor)."""
    import unicodedata
    n = lambda s: "".join(ch for ch in unicodedata.normalize("NFD", s.lower()) if unicodedata.category(ch) != "Mn")
    return [(k, v) for k, v in aplanar(datos).items() if n(etiqueta) in n(k.split(" | ")[-1]) and k != "archivo"]


_G, _T, _D = canon_sec("1. INFORMACIÓN GENERAL"), canon_sec("2. TÍTULO DEL ANTEPROYECTO"), canon_sec("4. DIRECCIÓN DEL TRABAJO")
_ALIAS = {   # campo de una evaluacion -> (seccion, columna) que ya usa el formato de anteproyecto
    "Autor 1": (_G, "Nombre completo (Estudiante 1)"), "Código estudiantil (Autor 1)": (_G, "Código estudiantil (Estudiante 1)"),
    "Autor 2": (_G, "Nombre completo (Estudiante 2)"), "Código estudiantil (Autor 2)": (_G, "Código estudiantil (Estudiante 2)"),
    "Correo institucional del autor 1": (_G, "Correo institucional (Estudiante 1)"), "Correo institucional del autor 2": (_G, "Correo institucional (Estudiante 2)"),
    "Título del trabajo": (_T, "Título del anteproyecto"),
    "Director": (_D, "Director designado"), "Codirector": (_D, "Codirector designado"),
}


def _columnas(d, alias=True):
    """Lista ordenada [((seccion, columna), valor)] con TODO lo que tiene el formato:
    campos, textos, observaciones de cada seccion, puntaje de cada criterio, Si/No de cada requisito y tablas."""
    res = []
    for sec, cuerpo in d["secciones"].items():
        if sec == "RESULTADO CALCULADO POR EL SCRIPT":
            for k, v in cuerpo.items(): res.append(((sec, k), v))
            continue
        campos, tablas = _recorrer(cuerpo, sec)
        if "NOTA FINAL" in sec.upper():          # tabla de nota: sus datos van en el bloque calculado
            campos = [x for x in campos if not x[2].startswith("texto") or x[3]]
            tablas = [x for x in tablas if x[2]]
        for s_, sub, k, v in campos:
            if sub and "ESTUDIANTE" in sub.upper():
                nom = f"{k} ({sub.split(' (')[0].replace('DATOS ', '').title()})"
            elif k.startswith("texto"):
                nom = f"Observaciones {sub.split('.')[0]}. {_limpia_sec(sub).split('.')[0][:40]}" if sub else _limpia_sec(canon_sec(s_))
                if sub: nom = f"Observaciones de la sección: {sub[:45]}"
            else:
                nom = k
            res.append(((canon_sec(sec), nom), v))
        for s_, sub, rows in tablas:
            if rows and ("Puntaje" in rows[0] or "Respuesta" in rows[0]):
                col = "Puntaje" if "Puntaje" in rows[0] else "Respuesta"
                etiqueta = next(iter(rows[0]))
                for x in rows:
                    base = str(x.get(etiqueta) or "").strip()
                    nombre = (base[:90] + ("…" if len(base) > 90 else ""))
                    res.append(((canon_sec(sub or s_), nombre), x.get(col)))
                    if x.get("Observación"):
                        res.append(((canon_sec(sub or s_), nombre + " (observación)"), x.get("Observación")))
            else:
                nom = canon_sec(sub or s_).capitalize()
                txt = "\n".join(" ; ".join(str(v) for v in r.values() if v is not None) for r in rows)
                res.append(((canon_sec(s_), "Tabla: " + nom), txt or None))
    if alias and str(d.get("_hoja", "")).startswith("Evaluacion"):
        res2 = []
        for (sec, nom), v in res:
            base = re.sub(r"\s*\(si aplica\)", "", nom)
            res2.append((_ALIAS[base], v) if sec == _G and base in _ALIAS else ((sec, nom), v))
        res = res2
    return res


def fila_csv(d):
    """Una fila por formato con TODAS las columnas (ver _columnas); si un nombre se repite se agrega el numero de la seccion."""
    fila = {"Archivo": d["_archivo"]}
    for (sec, nom), v in _columnas(d):
        k = nom
        if k in fila:
            k = f"{nom} [{sec[:24]}]"
            n = 2
            while k in fila:
                k = f"{nom} [{sec[:24]}-{n}]"; n += 1
        fila[k] = v
    return fila


def _registros(d):
    """[(seccion, campo, valor, fila_de_tabla)]: todos los datos de un formato; las tablas se expanden celda por celda."""
    res = []
    for sec, cuerpo in d["secciones"].items():
        if sec == "RESULTADO CALCULADO POR EL SCRIPT":
            res += [(sec, k, v, None) for k, v in cuerpo.items()]
            continue
        campos, tablas = _recorrer(cuerpo, sec)
        for s_, sub, k, v in campos:
            if sub and "ESTUDIANTE" in sub.upper():
                nom = f"{k} ({sub.split(' (')[0].replace('DATOS ', '').title()})"
            elif k.startswith("texto"):
                nom = f"Observaciones de la sección: {sub[:45]}" if sub else _limpia_sec(canon_sec(s_))
            else:
                nom = k
            res.append((canon_sec(s_), nom, v, None))
        for s_, sub, rows in tablas:
            nombre = canon_sec(sub or s_).capitalize()
            if rows and ("Puntaje" in rows[0] or "Respuesta" in rows[0]):
                col = "Puntaje" if "Puntaje" in rows[0] else "Respuesta"; etiqueta = next(iter(rows[0]))
                for n, x in enumerate(rows, 1):
                    res.append((canon_sec(sub or s_), str(x.get(etiqueta) or "").strip(), x.get(col), n))
                    if x.get("Observación"): res.append((canon_sec(sub or s_), str(x.get(etiqueta) or "").strip() + " (observación)", x["Observación"], n))
            else:
                for n, x in enumerate(rows, 1):
                    for col, v in x.items():
                        if v not in (None, ""): res.append((canon_sec(s_), f"{nombre}: {col}", v, n))
    return res


def _tipo(d):
    """Tipo de formato: 'Anteproyecto', 'Evaluación del anteproyecto', 'Evaluación del documento final', etc."""
    if d.get("_hoja") in ("Anteproyecto", "Formalizacion"):
        return "Anteproyecto"
    t = d.get("titulo") or d.get("_hoja") or "Formato"
    return re.sub(r"\s+", " ", t).strip().capitalize()


def _guardar_seguro(funcion, ruta):
    """Guarda; si el archivo esta abierto en Excel, agrega fecha y hora al nombre."""
    try:
        funcion(ruta); return ruta
    except PermissionError:
        base, ext = os.path.splitext(ruta)
        nueva = f"{base}_{datetime.datetime.now():%Y%m%d_%H%M%S}{ext}"
        funcion(nueva); return nueva


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("ruta", nargs="?", default=None,
                    help="archivo .xlsx o carpeta con varios (si se omite: la carpeta donde esta este script)")
    ap.add_argument("-o", "--json", help="ademas, guardar un JSON completo en este archivo")
    ap.add_argument("--csv", help="nombre del CSV (por defecto: resumen_extraccion.csv junto a los formatos)")
    ap.add_argument("--xlsx", help="nombre del Excel (por defecto: resumen_extraccion.xlsx junto a los formatos)")
    ap.add_argument("--sin-xlsx", action="store_true", help="no crear el Excel, solo el CSV")
    ap.add_argument("--ver", action="store_true", help="ademas, mostrar el JSON completo en la terminal")
    ap.add_argument("-c", "--campo", help="mostrar solo los campos cuya etiqueta contenga este texto (no guarda archivos)")
    ap.add_argument("--calculados", action="store_true", help="incluir celdas grises (formulas)")
    a = ap.parse_args()
    print(f"extraer_formato {VERSION}  (guarda CSV y Excel automaticamente)")
    ruta = a.ruta or os.path.dirname(os.path.abspath(__file__))
    carpeta = ruta if os.path.isdir(ruta) else (os.path.dirname(os.path.abspath(ruta)))
    archivos = sorted(glob.glob(os.path.join(ruta, "*.xlsx"))) if os.path.isdir(ruta) else [ruta]
    # se ignoran temporales de Excel y los resumenes que genera este mismo script
    archivos = [f for f in archivos if not os.path.basename(f).startswith(("~$", "resumen_extraccion"))]
    if not archivos:
        sys.exit(f"No encontre archivos .xlsx en: {ruta}")
    todos = []
    for f in archivos:
        try:
            todos.append(extraer(f, calculados=a.calculados))
        except Exception as e:
            print(f"  (omitido {os.path.basename(f)}: {e})")
    if a.campo:
        for d in todos:
            for k, v in buscar(d, a.campo):
                print(f"{d['_archivo']}: {k} = {v!r}")
        return
    # ---- se guarda siempre, sin pedirlo
    ruta_csv = a.csv or os.path.join(carpeta, "resumen_extraccion.csv")
    filas = [fila_csv(d) for d in todos]
    cols = list(dict.fromkeys(k for f in filas for k in f))
    def _csv(r):
        with open(r, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=cols, delimiter=";")
            w.writeheader(); w.writerows(filas)
    print("CSV guardado:  ", os.path.abspath(_guardar_seguro(_csv, ruta_csv)), f"({len(filas)} formato(s), {len(cols)} columnas)")
    if not a.sin_xlsx:
        ruta_x = a.xlsx or os.path.join(carpeta, "resumen_extraccion.xlsx")
        print("Excel guardado:", os.path.abspath(_guardar_seguro(lambda r: a_excel(todos, r), ruta_x)), "(Resumen, Base de datos, Bases por tipo de formato, Datos (lista) y tablas)")
    if a.json:
        txt = json.dumps(todos[0] if len(todos) == 1 else todos, ensure_ascii=False, indent=2, default=str)
        print("JSON guardado: ", _guardar_seguro(lambda r: open(r, "w", encoding="utf-8").write(txt), a.json))
    if a.ver:
        print(json.dumps(todos[0] if len(todos) == 1 else todos, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    main()
