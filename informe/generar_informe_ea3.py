"""Genera el informe Word de la Evidencia de aprendizaje 3.

Lee informe/datos_ea3.json (recolectar_datos_ea3.py), data/resumen.json y las
capturas de informe/capturas_ea3/ (capturar_pantallas.py). Los ayudantes de
formato son los mismos de los informes de la EA1 y la EA2.

Uso (desde la raíz del repositorio):
    python informe/recolectar_datos_ea3.py
    python informe/capturar_pantallas.py
    python informe/generar_informe_ea3.py
"""
import json
import tempfile
from datetime import datetime
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch, Patch
from PIL import Image

RAIZ = Path(__file__).resolve().parent.parent
INFORME = RAIZ / "informe"
CAPTURAS = INFORME / "capturas_ea3"
FIG_DIR = RAIZ / "figuras"
TMP = Path(tempfile.mkdtemp(prefix="ea3_informe_"))

# --- Datos de portada --------------------------------------------------------
AUTOR = "John Neider Cuervo Quintero"   # evidencia individual
INSTITUCION = "Institución Universitaria Digital de Antioquia (IU Digital)"
PROGRAMA = "Ingeniería de Software y Datos"
ASIGNATURA = "Programación para Análisis de Datos (PREICA2602B020048)"
ASIGNATURA_APA = "PREICA2602B020048: Programación para Análisis de Datos"
ACTIVIDAD = "Evidencia de aprendizaje 3"
DOCENTE = "Andrés Felipe Palacio"
CIUDAD = "Medellín"

SALIDA = RAIZ / "entrega" / "Cuervo_Quintero_John_Actividad_EA3.docx"

TINTA = RGBColor(0x1B, 0x26, 0x3B)
ACENTO = RGBColor(0x0E, 0x7C, 0x7B)
GRIS = RGBColor(0x5B, 0x63, 0x70)
BLANCO = RGBColor(0xFF, 0xFF, 0xFF)
HEX_TINTA, HEX_ACENTO, HEX_SUAVE, HEX_LINEA = "1B263B", "0E7C7B", "EAF4F3", "D5DBE3"

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]


def cop(x):
    return f"${x:,.0f}".replace(",", ".") if x is not None else "—"


def pct(x, dec=1):
    return f"{x:.{dec}f} %".replace(".", ",") if x is not None else "—"


def fecha_larga(iso):
    d = datetime.fromisoformat(iso.replace("Z", "+00:00"))
    return f"{d.day} de {MESES[d.month - 1]} de {d.year}"


def duracion(s):
    return f"{s // 60} min {s % 60:02d} s" if s >= 60 else f"{s} s"


# --- Utilidades de formato (las mismas del informe de la EA1) ----------------
def _xml(tag, **attrs):
    el = OxmlElement(tag)
    for k, v in attrs.items():
        el.set(qn(f"w:{k}"), str(v))
    return el


def estilos(doc):
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    normal.font.color.rgb = RGBColor(0x26, 0x2B, 0x33)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.2
    for nivel, tam, color, antes in ((1, 17, TINTA, 20), (2, 12.5, ACENTO, 12), (3, 11, TINTA, 8)):
        h = doc.styles[f"Heading {nivel}"]
        h.font.name = "Calibri"
        h.font.size = Pt(tam)
        h.font.bold = True
        h.font.italic = False
        h.font.color.rgb = color
        h.paragraph_format.space_before = Pt(antes)
        h.paragraph_format.space_after = Pt(6)
        for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
            h.element.rPr.rFonts.attrib.pop(qn(attr), None)
    ppr = doc.styles["Heading 1"].element.get_or_add_pPr()
    borde = _xml("w:pBdr")
    borde.append(_xml("w:bottom", val="single", sz=6, space=4, color=HEX_ACENTO))
    ppr.append(borde)
    for s in doc.sections:
        s.top_margin = Cm(2.3)
        s.bottom_margin = Cm(2.2)
        s.left_margin = s.right_margin = Cm(2.4)


def campo_pagina(run):
    for tipo, texto in (("begin", None), (None, "PAGE"), ("end", None)):
        if tipo:
            run._r.append(_xml("w:fldChar", fldCharType=tipo))
        else:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = texto
            run._r.append(el)


def encabezado_y_pie(seccion):
    seccion.different_first_page_header_footer = True
    p = seccion.first_page_header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run()
    campo_pagina(r)
    r.font.size = Pt(11)
    p = seccion.header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("CI/CD y despliegue · App del Galaxy A55 en Mercado Libre")
    r.font.size, r.font.color.rgb = Pt(8.5), GRIS
    p = seccion.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run(f"{ACTIVIDAD} · IU Digital   |   ")
    r.font.size, r.font.color.rgb = Pt(8.5), GRIS
    r = p.add_run()
    campo_pagina(r)
    r.font.size, r.font.color.rgb, r.bold = Pt(8.5), ACENTO, True


def parrafo(doc, texto, negrita_inicial=None, cursiva=False, alinear=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = doc.add_paragraph()
    p.alignment = alinear
    if negrita_inicial:
        p.add_run(negrita_inicial).bold = True
    r = p.add_run(texto)
    r.italic = cursiva
    return p


def vinetas(doc, items, estilo="List Bullet"):
    for it in items:
        p = doc.add_paragraph(style=estilo)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(3)
        if isinstance(it, tuple):
            p.add_run(it[0]).bold = True
            p.add_run(it[1])
        else:
            p.add_run(it)


def sombrear(celda, color):
    celda._tc.get_or_add_tcPr().append(_xml("w:shd", val="clear", color="auto", fill=color))


def bordes_celda(celda, **lados):
    tc_bordes = _xml("w:tcBorders")
    for lado in ("top", "left", "bottom", "right"):
        valor = lados.get(lado)
        if valor:
            tc_bordes.append(_xml(f"w:{lado}", val="single", sz=valor[0], space=0, color=valor[1]))
        else:
            tc_bordes.append(_xml(f"w:{lado}", val="nil"))
    celda._tc.get_or_add_tcPr().append(tc_bordes)


def bordes_tabla(t, visibles=True):
    borders = _xml("w:tblBorders")
    for borde in ("top", "left", "bottom", "right", "insideH", "insideV"):
        if visibles and borde in ("top", "bottom"):
            borders.append(_xml(f"w:{borde}", val="single", sz=8, space=0, color=HEX_TINTA))
        elif visibles and borde == "insideH":
            borders.append(_xml(f"w:{borde}", val="single", sz=4, space=0, color=HEX_LINEA))
        else:
            borders.append(_xml(f"w:{borde}", val="nil"))
    t._tbl.tblPr.append(borders)
    margenes = _xml("w:tblCellMar")
    for lado, valor in (("top", 45), ("bottom", 45), ("left", 100), ("right", 100)):
        margenes.append(_xml(f"w:{lado}", w=valor, type="dxa"))
    t._tbl.tblPr.append(margenes)


def anchos_tabla(t, anchos):
    t.autofit = False
    for col, w in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), anchos, strict=False):
        col.set(qn("w:w"), str(int(Cm(w).twips)))
    for fila in t.rows:
        for i, w in enumerate(anchos):
            fila.cells[i].width = Cm(w)


def mantener_junta(t):
    t.rows[0]._tr.get_or_add_trPr().append(_xml("w:tblHeader"))
    for fila in t.rows:
        fila._tr.get_or_add_trPr().append(_xml("w:cantSplit"))
    if len(t.rows) <= 14:
        for fila in t.rows[:-1]:
            for celda in fila.cells:
                for par in celda.paragraphs:
                    par.paragraph_format.keep_with_next = True


def titulo_apa(doc, etiqueta, titulo):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.keep_with_next = True
    p.add_run(etiqueta).bold = True
    p2 = doc.add_paragraph()
    p2.paragraph_format.keep_with_next = True
    p2.paragraph_format.space_after = Pt(4)
    p2.add_run(titulo).italic = True


def nota(doc, texto):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    r = p.add_run("Nota. ")
    r.italic, r.font.size, r.font.color.rgb = True, Pt(9), GRIS
    r = p.add_run(texto)
    r.font.size, r.font.color.rgb = Pt(9), GRIS


class Numerador:
    """Lleva la cuenta de tablas y figuras para los rótulos APA."""

    def __init__(self):
        self.tablas = 0
        self.figuras = 0

    def tabla(self):
        self.tablas += 1
        return f"Tabla {self.tablas}"

    def figura(self):
        self.figuras += 1
        return f"Figura {self.figuras}"


N = Numerador()


def tabla(doc, titulo, encabezado, filas, anchos=None, nota_txt=None, tam=9.5):
    titulo_apa(doc, N.tabla(), titulo)
    t = doc.add_table(rows=1, cols=len(encabezado))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    bordes_tabla(t)
    for i, h in enumerate(encabezado):
        c = t.rows[0].cells[i]
        c.text = ""
        r = c.paragraphs[0].add_run(h)
        r.bold, r.font.size, r.font.color.rgb = True, Pt(tam), BLANCO
        sombrear(c, HEX_TINTA)
    for k, fila in enumerate(filas):
        celdas = t.add_row().cells
        for i, v in enumerate(fila):
            celdas[i].text = ""
            p = celdas[i].paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.05
            r = p.add_run(str(v))
            r.font.size = Pt(tam)
            if k % 2:
                sombrear(celdas[i], "F4F6F9")
    mantener_junta(t)
    if anchos:
        anchos_tabla(t, anchos)
    if nota_txt:
        nota(doc, nota_txt)
    else:
        doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def destacado(doc, titulo, texto):
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    bordes_tabla(t, visibles=False)
    c = t.cell(0, 0)
    sombrear(c, HEX_SUAVE)
    bordes_celda(c, left=(36, HEX_ACENTO))
    p = c.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(titulo.upper())
    r.bold, r.font.size, r.font.color.rgb = True, Pt(8.5), ACENTO
    p = c.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    p.add_run(texto).font.size = Pt(10.5)
    anchos_tabla(t, [16.2])
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def indicadores(doc, datos):
    t = doc.add_table(rows=2, cols=len(datos))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    bordes_tabla(t, visibles=False)
    for i, (valor, etiqueta) in enumerate(datos):
        for fila in (0, 1):
            sombrear(t.cell(fila, i), "F4F6F9")
        bordes_celda(t.cell(0, i), top=(18, HEX_ACENTO))
        p = t.cell(0, i).paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(str(valor))
        r.bold, r.font.size, r.font.color.rgb = True, Pt(13), TINTA
        p = t.cell(1, i).paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(etiqueta)
        r.font.size, r.font.color.rgb = Pt(8.5), GRIS
    anchos_tabla(t, [16.2 / len(datos)] * len(datos))
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def figura(doc, titulo, ruta, texto_nota, ancho=15.5):
    ruta = Path(ruta)
    if not ruta.exists():
        raise FileNotFoundError(ruta)
    titulo_apa(doc, N.figura(), titulo)
    doc.add_picture(str(ruta), width=Cm(ancho))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.paragraphs[-1].paragraph_format.keep_with_next = True
    nota(doc, texto_nota)


def recorte(nombre, alto=None, borde=True):
    """Recorta una captura (alto en píxeles desde arriba) y le agrega un borde fino."""
    img = Image.open(CAPTURAS / nombre).convert("RGB")
    if alto and img.height > alto:
        img = img.crop((0, 0, img.width, alto))
    if borde:
        marco = Image.new("RGB", (img.width + 4, img.height + 4), (0xD5, 0xDB, 0xE3))
        marco.paste(img, (2, 2))
        img = marco
    destino = TMP / nombre
    img.save(destino)
    return destino


def codigo(doc, lineas, tam=7.5):
    for linea in lineas:
        cp = doc.add_paragraph()
        cp.paragraph_format.space_after = Pt(0)
        cp.paragraph_format.line_spacing = 1.0
        r = cp.add_run(linea if linea else " ")
        r.font.name = "Consolas"
        r._element.rPr.rFonts.set(qn("w:eastAsia"), "Consolas")
        r.font.size = Pt(tam)


def bloque_codigo(doc, rel):
    ruta = RAIZ / rel
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(rel)
    r.bold, r.font.color.rgb = True, ACENTO
    codigo(doc, ruta.read_text(encoding="utf-8").splitlines())


# --- Figuras propias del informe ---------------------------------------------
def figura_flujo():
    """Diagrama del flujo GitHub Flow + CI/CD con despliegue en Render."""
    fig, ax = plt.subplots(figsize=(8.6, 4.3))
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 46)
    ax.axis("off")
    azul, verde, rojo, gris, tinta, morado = "#2a78d6", "#0e7c7b", "#c0392b", "#8a919c", "#1b263b", "#6250d6"

    def caja(x, y, w, h, titulo, sub, color):
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.4,rounding_size=1.2",
                                    fc="white", ec=color, lw=1.6))
        ax.text(x + w / 2, y + h - 2.6, titulo, ha="center", va="center", fontsize=9.5, weight="bold",
                color=tinta)
        ax.text(x + w / 2, y + h / 2 - 1.6, sub, ha="center", va="center", fontsize=7.2, color="#52514e",
                linespacing=1.25)

    def flecha(a, b, color=gris, texto=None, rad=0.0, dy=1.6):
        ax.add_patch(FancyArrowPatch(a, b, arrowstyle="-|>", mutation_scale=12, lw=1.3, color=color,
                                     connectionstyle=f"arc3,rad={rad}"))
        if texto:
            ax.text((a[0] + b[0]) / 2, (a[1] + b[1]) / 2 + dy, texto, ha="center", fontsize=7.5, color=color,
                    weight="bold")

    ax.text(1, 44, "DESARROLLO (GitHub Flow)", fontsize=8, color=azul, weight="bold")
    caja(1, 28, 15, 13, "Issue", "objetivo y\ntareas", azul)
    caja(21, 28, 15, 13, "Rama", "feature/, ci/,\ndocs/, demo/", azul)
    caja(41, 28, 15, 13, "Commits", "Conventional\nCommits", azul)
    caja(61, 28, 16, 13, "Pull request", "plantilla +\nCloses #n", azul)
    caja(82, 28, 17, 13, "CI en el PR", "ruff · pytest 3.12/3.13\npipeline de datos\nimagen Docker", verde)
    for x in (16.8, 36.8, 56.8):
        flecha((x, 34.5), (x + 3.4, 34.5))
    flecha((77.8, 34.5), (81.2, 34.5))
    flecha((88, 27.4), (49, 27.4), color=rojo, texto="falla: merge bloqueado", rad=-0.12, dy=-4.2)

    ax.text(1, 19, "INTEGRACIÓN, ENTREGA Y DESPLIEGUE", fontsize=8, color=verde, weight="bold")
    caja(82, 3, 17, 13, "Merge a main", "rama protegida:\ncheck «ci»", verde)
    caja(61, 3, 16, 13, "CI en main", "mismas pruebas\nsobre el merge", verde)
    caja(40, 3, 16, 13, "Render", "checksPass:\nconstruye la imagen\ny despliega", morado)
    caja(19, 3, 16, 13, "Deploy Render", "/salud = commit\n+ smoke test", verde)
    flecha((92, 27.4), (92, 16.8), color=verde)
    ax.text(93.2, 21.5, "pasa", fontsize=7.5, color=verde, weight="bold")
    flecha((81.2, 9.5), (77.8, 9.5))
    flecha((60.2, 9.5), (56.8, 9.5), color=morado, texto="checks ✓", dy=1.4)
    flecha((39.2, 9.5), (35.8, 9.5))
    ax.text(1, 9.5, "usuarios\n(app en la\nnube)", fontsize=8, color=tinta, va="center")
    flecha((18.2, 9.5), (10.5, 9.5))
    destino = TMP / "fig_flujo_ea3.png"
    fig.savefig(destino, dpi=220, bbox_inches="tight")
    plt.close(fig)
    return destino


def figura_ejecuciones(ejec):
    """Duración de cada ejecución de GitHub Actions, en orden."""
    ejec = [e for e in ejec if e["conclusion"] in ("success", "failure")]
    colores = {"CI": "#2a78d6", "Deploy": "#8a919c", "Deploy Render": "#0e7c7b"}
    fig, ax = plt.subplots(figsize=(8, 3.4))
    x = range(1, len(ejec) + 1)
    ax.bar(x, [e["duracion_s"] for e in ejec], width=0.7,
           color=["#c0392b" if e["conclusion"] == "failure" else colores.get(e["workflow"], "#8a919c") for e in ejec])
    ax.set_xticks(list(x), [str(i) for i in x], fontsize=8)
    ax.set_ylabel("Segundos", color="#52514e")
    ax.set_xlabel("Ejecución (en orden cronológico)", color="#52514e")
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.grid(axis="y", color="#e4e3df", linewidth=0.6)
    ax.set_axisbelow(True)
    ax.legend(handles=[Patch(color="#2a78d6", label="CI"), Patch(color="#0e7c7b", label="Deploy Render"),
                       Patch(color="#8a919c", label="Deploy Pages"), Patch(color="#c0392b", label="Fallo")],
              frameon=False, fontsize=8, loc="upper center", bbox_to_anchor=(0.5, 1.16), ncol=4)
    destino = TMP / "fig_ejecuciones_ea3.png"
    fig.tight_layout()
    fig.savefig(destino, dpi=220)
    plt.close(fig)
    return destino


# --- Secciones -----------------------------------------------------------------
def portada(doc, fecha):
    """Portada de trabajo de estudiante según APA 7: centrada, a doble espacio y sin adornos."""

    def linea(texto="", negrita=False, antes=0):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 2.0
        p.paragraph_format.space_before = Pt(antes)
        p.paragraph_format.space_after = Pt(0)
        if texto:
            r = p.add_run(texto)
            r.bold = negrita
            r.font.size = Pt(12)
            r.font.color.rgb = RGBColor(0, 0, 0)

    linea("Integración y despliegue continuos de una aplicación web:", negrita=True, antes=110)
    linea("CI/CD con GitHub Actions y despliegue en Render del análisis", negrita=True)
    linea("de precios del Samsung Galaxy A55 en Mercado Libre", negrita=True)
    linea()
    linea(AUTOR)
    linea(PROGRAMA)
    linea(INSTITUCION.replace(" (IU Digital)", ""))
    linea(ASIGNATURA_APA)
    linea(DOCENTE)
    linea(fecha)
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def resumen(doc, D, R, M):
    doc.add_heading("Resumen", 1)
    parrafo(doc, "En esta evidencia convertí el análisis de precios del Samsung Galaxy A55, que construí en la EA1 con "
            "Selenium, BeautifulSoup y Scrapy y que en la EA2 puse bajo un flujo DevOps, en una aplicación web "
            "desplegada en la nube. La app, hecha con FastAPI, sirve un tablero y una API REST con filtros. La "
            "empaqueté en una imagen Docker y monté en GitHub Actions un pipeline de CI/CD: en cada pull request se "
            f"revisa el estilo, se ejecutan {M['pruebas']} pruebas automatizadas en Python 3.12 y 3.13 "
            f"(cobertura de {pct(M['cobertura'], 0)}, 100 % en la app) y se construye y prueba la imagen. Cuando un "
            "cambio llega a main con los checks en verde, Render lo despliega y un workflow verifica que producción "
            "sirva exactamente ese commit.")
    parrafo(doc, f"Trabajé con {M['issues']} issues y {M['prs']} pull requests, y el pipeline se ejecutó "
            f"{M['ejecuciones']} veces. El último cambio pasó de merge a producción verificada en "
            f"{duracion(M['merge_a_prod_s'])}, sin pasos manuales. Una prueba de fallo controlada demostró que el "
            "pipeline detecta errores que las pruebas unitarias no ven: una imagen sin datos pasó pytest, pero el "
            "smoke test del contenedor la bloqueó y Render no la desplegó.")
    p = doc.add_paragraph()
    p.add_run("Palabras clave: ").bold = True
    p.add_run("DevOps, CI/CD, GitHub Actions, Docker, FastAPI, Render, despliegue continuo, pruebas "
              "automatizadas, web scraping.").italic = True


def introduccion(doc):
    doc.add_heading("1. Introducción", 1)
    parrafo(doc, "Al terminar la EA2 el proyecto ya tenía control de versiones, pruebas automatizadas y un tablero "
            "estático que GitHub Pages publicaba en cada merge. Pero un sitio estático no es un servicio: no recibe "
            "consultas, no tiene un proceso corriendo que pueda fallar y no exige empaquetar el entorno de "
            "ejecución. La Unidad 3 propone el siguiente paso, que es desarrollar una aplicación, automatizar su "
            "integración y su entrega con GitHub Actions y desplegarla en un servicio real en la nube.")
    parrafo(doc, "Eso fue lo que hice. Escribí una aplicación web con FastAPI que expone los resultados del scraper "
            "como tablero y como API. La empaqueté en un contenedor Docker para que el mismo artefacto corra igual "
            "en mi equipo, en GitHub Actions y en Render. Y extendí el pipeline para que nada llegue a producción "
            "sin pasar por las pruebas. Me apoyé en las prácticas de entrega continua de Humble y Farley (2010), en "
            "las de DevOps de Kim et al. (2021) y en las métricas de Forsgren et al. (2018): cambios pequeños, "
            "frecuentes, verificados de forma automática y con el tiempo de entrega como indicador.")
    parrafo(doc, "El informe mantiene la estructura de las evidencias anteriores. Primero retoma la página y el "
            "producto analizados. Luego presenta el tema, los objetivos y la metodología (el scraping, la app, Docker "
            "y el pipeline de CI/CD) y cierra con los resultados medidos en el propio repositorio y las "
            "conclusiones. El código está en el repositorio público, en el Anexo A y en el notebook que acompaña "
            "este documento.")


def descripcion(doc, R):
    doc.add_heading("2. La página y el artículo analizados", 1)
    doc.add_heading("2.1 Mercado Libre Colombia", 2)
    parrafo(doc, "Mercado Libre (MercadoLibre, Inc., NASDAQ: MELI) nació en 1999 en Buenos Aires y hoy opera el "
            "ecosistema de comercio electrónico y pagos digitales más grande de América Latina, con presencia en 18 "
            "países (MercadoLibre, Inc., 2025). En Colombia funciona bajo el dominio mercadolibre.com.co. Es un "
            "marketplace: miles de vendedores, desde personas que venden su celular usado hasta tiendas oficiales, "
            "compiten en la misma búsqueda. Cada publicación muestra precio, descuento, envío y condición del "
            "producto.")
    parrafo(doc, "El scraping trabaja sobre dos tipos de página: el listado de búsqueda "
            "(listado.mercadolibre.com.co/samsung-galaxy-a55), con hasta 48 tarjetas por página, y la ficha de "
            "producto (rutas /p/MCO…). El sitio combina HTML estático, contenido que se carga con JavaScript y una "
            "verificación anti-bot que redirige a /gz/account-verification las peticiones que no vienen de un "
            "navegador real.")
    doc.add_heading("2.2 Samsung Galaxy A55 5G", 2)
    parrafo(doc, "El Galaxy A55 5G es un teléfono de gama media-alta lanzado en marzo de 2024. Lo elegí porque es "
            "un modelo muy buscado, sus unidades son comparables entre sí y lo venden tanto la tienda oficial como "
            "revendedores y personas naturales, así que la oferta es variada (Samsung Electronics, 2024).")
    tabla(doc, "Especificaciones principales del Samsung Galaxy A55 5G", ["Característica", "Especificación"], [
        ("Pantalla", "6,6\" Super AMOLED FHD+, 120 Hz"),
        ("Procesador", "Samsung Exynos 1480 (4 nm)"),
        ("Memoria RAM / almacenamiento", "8 GB / 128 GB o 256 GB, ampliable con microSD"),
        ("Cámaras", "Trasera de 50 MP con OIS, 12 MP ultra gran angular y 5 MP macro; frontal de 32 MP"),
        ("Batería", "5.000 mAh con carga rápida de 25 W"),
        ("Construcción", "Marco de aluminio, Gorilla Glass Victus+, resistencia IP67"),
    ], anchos=[5.2, 11], nota_txt="Tomado de Samsung Electronics (2024).")
    doc.add_heading("2.3 Lo que dejaron la EA1 y la EA2", 2)
    cond = {c["condicion"]: c for c in R["por_condicion"]}
    parrafo(doc, "La EA1 extrajo el 25 de octubre de 2026 las tres primeras páginas de la búsqueda «samsung galaxy "
            "a55». Estas cifras son las que ahora sirve la app web y las que las pruebas de regresión deben "
            "reproducir en cada cambio:")
    indicadores(doc, [(R["publicaciones_totales"], "publicaciones extraídas"),
                      (R["publicaciones_a55"], "son un Galaxy A55"),
                      (cop(cond["Nuevo"]["median"]), f"mediana nuevo (n={cond['Nuevo']['count']})"),
                      (cop(cond["Usado"]["median"]), f"mediana usado (n={cond['Usado']['count']})"),
                      (pct(R["pct_envio_gratis"], 0), "con envío gratis")])
    parrafo(doc, "La EA2 dejó el proyecto en un repositorio de GitHub con issues, ramas, pull requests y una rama "
            "main protegida. También dejó 37 pruebas, un workflow de integración continua y otro que publica el "
            "tablero en GitHub Pages. Para la EA3 creé un repositorio nuevo a partir de ese, de modo que la "
            "evidencia anterior quedara intacta, y seguí trabajando con el mismo flujo.")


def tema(doc):
    doc.add_heading("3. Tema de interés", 1)
    destacado(doc, "Pregunta de esta evidencia",
              "¿Cómo llevar el análisis del Galaxy A55 a una aplicación web en la nube que se pruebe, se empaquete "
              "y se despliegue sola cada vez que cambia el código, sin que un error llegue a producción?")
    parrafo(doc, "El tema es la automatización del ciclo completo de entrega de un servicio: desde el commit hasta "
            "la aplicación corriendo en un proveedor de nube. Pasar de un sitio estático a un servicio cambia lo "
            "que hay que verificar. Ya no basta con que el código pase las pruebas: también tiene que construirse "
            "la imagen, arrancar el proceso, responder en el puerto correcto y servir los datos correctos. Las "
            "preguntas que guiaron el trabajo fueron:")
    vinetas(doc, [
        "¿Qué debe ofrecer una aplicación web simple sobre estos datos y cómo reutilizar el código de las "
        "evidencias anteriores en lugar de reescribirlo?",
        "¿Cómo empaquetarla para que corra igual en mi equipo, en el CI y en la nube?",
        "¿Cómo encadenar pruebas, construcción de la imagen y despliegue para que solo llegue a producción lo que "
        "pasó todas las verificaciones?",
        "¿Cómo saber, sin entrar a ningún panel, que producción está sirviendo la versión correcta?",
    ])


def objetivos(doc):
    doc.add_heading("4. Objetivos", 1)
    doc.add_heading("Escenario", 2)
    parrafo(doc, "Imaginé que formo parte del equipo de desarrollo de una empresa que vende celulares en línea. El "
            "área comercial necesita consultar en cualquier momento los precios de la competencia en Mercado "
            "Libre, desde el navegador o desde otros sistemas. El proyecto de software del equipo es el scraper y "
            "su análisis. Mi tarea es versionarlo con Git y GitHub, convertirlo en un servicio web y establecer un "
            "pipeline de CI/CD con GitHub Actions que integre pruebas automatizadas y despliegue continuo, "
            "documentado en un repositorio de GitHub.")
    doc.add_heading("4.1 Objetivo general", 2)
    parrafo(doc, "Desarrollar y desplegar en la nube una aplicación web que publique el análisis de precios del "
            "Samsung Galaxy A55, con control de versiones en Git y GitHub y un pipeline de CI/CD en GitHub Actions "
            "que integre pruebas automatizadas, construcción de la imagen y despliegue continuo verificado, "
            "documentado en un repositorio público con la explicación del flujo DevOps implementado.")
    doc.add_heading("4.2 Objetivos específicos", 2)
    vinetas(doc, [
        "Desarrollar con FastAPI una aplicación web con tablero, API REST con filtros, chequeo de salud y "
        "documentación automática, reutilizando el código de la EA1 y la EA2.",
        "Ampliar la suite de pruebas automatizadas para cubrir todas las rutas de la app, incluidos los casos de "
        "error (404, 422 y 503).",
        "Empaquetar la app en una imagen Docker liviana, segura (sin privilegios de root) y con chequeo de salud.",
        "Extender el pipeline de integración continua para construir y probar la imagen en cada pull request.",
        "Configurar el despliegue continuo en Render, condicionado a los checks de GitHub Actions, y verificarlo "
        "automáticamente en producción.",
        "Documentar el flujo DevOps en el repositorio y evaluarlo con datos reales y con una prueba de fallo "
        "controlada.",
    ], estilo="List Number")


def metodologia(doc, D, R, M, fig_flujo):
    doc.add_heading("5. Metodología", 1)
    parrafo(doc, "La metodología tiene dos capas. La de datos es el scraping de la EA1, que sigue siendo la fuente "
            "de todo lo que muestra la app. La de desarrollo es GitHub Flow con un pipeline de CI/CD que ahora "
            "termina en un servicio en la nube.")

    doc.add_heading("5.1 Metodología de scraping", 2)
    parrafo(doc, "La extracción es la de la EA1 y está organizada en fases. La configuración común (URL, selectores "
            "CSS, pausas y rutas) vive en scraping/config.py, así que un cambio en el HTML del sitio se corrige en "
            "un solo lugar.")
    tabla(doc, "Fases del scraping y dónde se ejecutan", ["Fase", "Qué se hace", "Herramienta", "Dónde"], [
        ("1. Reconocimiento", "robots.txt e inspección del HTML para ubicar los selectores", "Navegador, config.py",
         "Manual"),
        ("2. Captura", "Listado y ficha en Firefox; espera del JavaScript; guarda el HTML en data/raw/", "Selenium",
         "Equipo local"),
        ("3. Parsing", "Título, precio, descuento, envío y condición de cada tarjeta", "BeautifulSoup", "Local y CI"),
        ("4. Rastreo", "Spider con paginación, limpieza y deduplicación", "Scrapy", "Local y CI"),
        ("5. Análisis", "Filtrado del A55, estadísticos y figuras", "pandas, matplotlib", "Local y CI"),
        ("6. Publicación", "Tablero y API sobre los datos procesados", "FastAPI, 06_sitio.py", "Render y Pages"),
    ], anchos=[3.2, 6.6, 3.2, 3.2])
    parrafo(doc, "El pipeline automatizado no hace scraping en vivo. Mercado Libre redirige a una verificación "
            "anti-bot las peticiones que no vienen de un navegador, los servidores de GitHub no tienen sesión "
            "gráfica para Selenium y consultar el sitio en cada push sería poco respetuoso con su infraestructura "
            "(Krotov et al., 2020). Las páginas capturadas por Selenium se versionan en data/raw/ y funcionan como "
            "datos de prueba. La app, por su parte, sirve los datos ya procesados (resumen.json y a55_limpio.csv), "
            "así que en producción no hace ninguna petición a Mercado Libre.")

    doc.add_heading("5.2 Flujo de trabajo con Git y GitHub", 2)
    parrafo(doc, "Seguí GitHub Flow, igual que en la EA2: main siempre está lista para desplegarse y todo cambio "
            "nace en una rama corta, se propone en un pull request, pasa el CI y se fusiona (GitHub, 2026c). Antes de "
            "programar registré un issue por cada pieza de la evidencia. Cada pull request explica qué cambia y cómo "
            "se probó, y cierra su issue con «Closes #n». La rama main está protegida: solo cambia por pull request "
            "y exige el check «" + ", ".join(D["proteccion_main"]["checks"]) + "» en verde.")
    issues = {i["numero"]: i["titulo"] for i in D["issues"]}
    filas = []
    for p in D["pull_requests"]:
        estado = "Fusionado" if p["fusionado"] else "Cerrado (demo)"
        filas.append((f"#{p['numero']}", p["rama"], p["titulo"].split(": ", 1)[-1],
                      ", ".join(f"#{n}" for n in p["issues"]) or "—", f"{p['checks_ok']}/{p['checks_total']}", estado))
    tabla(doc, "Pull requests de la EA3", ["PR", "Rama", "Cambio", "Issue", "Checks", "Resultado"], filas,
          anchos=[1.1, 4.0, 5.6, 1.3, 1.5, 2.7], tam=8.5,
          nota_txt="Datos de la API de GitHub. El PR #5 se abrió antes de existir el job docker, por eso tiene 3 "
                   "checks. El PR #11 es la prueba de fallo controlada (sección 6.4).")
    parrafo(doc, "Issues: " + "; ".join(f"#{n} {t}" for n, t in sorted(issues.items())) + ".")

    doc.add_heading("5.3 La aplicación web", 2)
    parrafo(doc, "Elegí FastAPI porque genera la documentación OpenAPI y valida los parámetros a partir de las "
            "anotaciones de tipo, lo que reduce el código propio y deja la API autodescrita en /docs (Ramírez, "
            "2026). La decisión de diseño más importante fue no reescribir nada: app/main.py importa "
            "scraping/config.py para ubicar los datos y reutiliza la función construir() de 06_sitio.py, la misma "
            "que genera el tablero de GitHub Pages. Solo le agregué un parámetro opcional para que el pie del "
            "tablero muestre el commit desplegado en Render.")
    tabla(doc, "Rutas de la aplicación", ["Ruta", "Respuesta", "Uso"], [
        ("GET /", "Tablero HTML con indicadores, gráficas y las 30 publicaciones", "Personas"),
        ("GET /salud", "Estado, commit desplegado (RENDER_GIT_COMMIT) y entorno", "Render y el CD"),
        ("GET /api/resumen", "Indicadores del análisis en JSON; 503 si faltan los datos", "Otros sistemas"),
        ("GET /api/publicaciones", "Publicaciones del A55 con filtros condicion, capacidad, precio_max y limite; "
                                   "422 si un parámetro no es válido", "Otros sistemas"),
        ("GET /api/publicaciones/{id}", "Una publicación por su id MCO; 404 si no existe", "Otros sistemas"),
        ("GET /docs", "Documentación OpenAPI interactiva generada por FastAPI", "Desarrolladores"),
    ], anchos=[4.4, 8.6, 3.2])
    parrafo(doc, "La configuración sale de variables de entorno, como recomienda la metodología Twelve-Factor "
            "(Wiggins, 2017): el puerto lo asigna Render en PORT, el entorno se marca con APP_ENV y el commit lo "
            "inyecta Render en RENDER_GIT_COMMIT. Esta última es la que permite verificar el despliegue desde "
            "afuera.")

    doc.add_heading("5.4 Pruebas automatizadas", 2)
    parrafo(doc, "Agregué tests/test_app.py, que usa el TestClient de FastAPI para ejecutar la app en memoria. Cubre "
            "todas las rutas, los filtros solos y combinados, los parámetros inválidos (422), una publicación "
            "inexistente (404), los datos faltantes (503) y que el tablero muestre el commit de Render. Las pruebas "
            "de la EA2 siguen igual: ninguna usa Internet y todas escriben en una carpeta temporal.")
    descr = {
        "tests/test_parsers.py": ("Unitarias", "Limpieza de precios, id MCO, capacidad y parsing de tarjetas"),
        "tests/test_listado_html.py": ("Integración", "BeautifulSoup sobre las 3 páginas reales y vigencia de los "
                                                      "selectores"),
        "tests/test_scrapy.py": ("Integración", "Pipeline de limpieza y spider con origen=local"),
        "tests/test_analisis.py": ("Regresión", "El análisis reproduce las cifras de la EA1"),
        "tests/test_sitio.py": ("Integración", "Generación del tablero y escape de HTML"),
        "tests/test_app.py": ("API (nuevas)", "Rutas, filtros, códigos 404/422/503 y commit en el tablero"),
    }
    ppa = D["pruebas_por_archivo"]
    filas = [(Path(a).name, descr[a][0], descr[a][1], n) for a, n in ppa.items() if a in descr]
    filas.append(("Total", "", f"Cobertura total {pct(M['cobertura'])}; app/main.py "
                  f"{pct(D['cobertura']['archivos'].get('app/main.py', 0), 0)}", sum(ppa.values())))
    tabla(doc, "Suite de pruebas automatizadas", ["Archivo", "Tipo", "Qué valida", "Pruebas"], filas,
          anchos=[3.6, 2.6, 8.0, 2.0], tam=9)

    doc.add_heading("5.5 Contenerización con Docker", 2)
    parrafo(doc, "La imagen se construye con el Dockerfile del repositorio (Docker Inc., 2026). Cada decisión "
            "responde a un riesgo concreto:")
    tabla(doc, "Decisiones del Dockerfile", ["Decisión", "Por qué"], [
        ("Base python:3.12-slim", "Imagen oficial y reducida; la misma versión que prueba el CI"),
        ("requirements-app.txt aparte", "La imagen solo lleva FastAPI, uvicorn y pandas: ni Selenium ni Scrapy"),
        (".dockerignore en lista blanca", "Solo viajan la app, dos módulos del scraper, los datos procesados y las "
                                          "figuras; el HTML de data/raw/ no entra"),
        ("Usuario app (uid 10001)", "El proceso no corre como root dentro del contenedor"),
        ("HEALTHCHECK a /salud", "Docker y Render saben si el proceso responde, no solo si existe"),
        ("Puerto desde $PORT", "Render asigna el puerto; en local se usa 8000"),
    ], anchos=[5.0, 11.2])
    parrafo(doc, f"La imagen final pesa unos {D['imagen']['tamano_mb']} MB, casi todo por pandas y NumPy. Es un "
            "tamaño aceptable para el plan gratuito y el CI la construye en menos de un minuto.")

    doc.add_heading("5.6 Integración continua con GitHub Actions", 2)
    parrafo(doc, "El workflow ci.yml se ejecuta en cada pull request hacia main y en cada push a main (GitHub, "
            "2026a). A lo que ya tenía en la EA2 le agregué un job docker que corre en paralelo con las pruebas:")
    tabla(doc, "Jobs del workflow de integración continua (ci.yml)", ["Job", "Pasos", "Propósito"], [
        ("Pruebas (3.12 y 3.13)", "checkout · setup-python con caché · ruff · pytest --cov · BeautifulSoup y Scrapy "
                                  "--local · análisis · resumen · artefactos", "Calidad del código y del pipeline "
                                                                               "de datos"),
        ("Imagen Docker", "docker build · docker run · curl a /salud (con el commit), /api/resumen, filtros, 404, / y "
                          "una figura · espera del HEALTHCHECK", "Que el artefacto desplegable funcione"),
        ("ci", "needs: [pruebas, docker]", "Check único exigido por la protección de main"),
    ], anchos=[3.4, 8.6, 4.2], tam=9)

    doc.add_heading("5.7 Despliegue continuo en Render", 2)
    parrafo(doc, "Elegí Render porque ofrece un plan gratuito para servicios web con Docker y se configura como "
            "código con un archivo Blueprint, render.yaml (Render, 2026a). Ese archivo declara el servicio "
            "ea3-galaxy-a55: runtime Docker, plan free, healthCheckPath /salud y autoDeployTrigger: checksPass. "
            "Esta última opción hace que Render despliegue un commit de main solo cuando sus checks de GitHub "
            "Actions terminan en verde (Render, 2026b).")
    parrafo(doc, "Mi primer diseño usaba un Deploy Hook: una URL secreta de Render que el workflow llamaría después "
            "del CI. Funciona, pero obliga a guardar ese token en los secretos de GitHub. Al revisar la "
            "documentación encontré checksPass, que logra la misma garantía sin tokens, porque la condición la "
            "evalúa Render directamente sobre los checks de GitHub. Dejé el Deploy Hook como opción: si alguien "
            "configura el secreto, el workflow también lo usa.")
    parrafo(doc, "El workflow deploy-render.yml se dispara con workflow_run cuando CI termina con éxito en main, el "
            "mismo patrón que el despliegue a Pages de la EA2. No despliega: verifica. Consulta /salud en producción "
            "cada 10 segundos, hasta 15 minutos, hasta que el commit reportado sea el que CI validó. Después hace "
            "un smoke test de la API y del tablero y publica en la ejecución un resumen con la URL, el commit y el "
            "tiempo que tardó en quedar en línea. Usa el environment produccion de GitHub, que enlaza cada "
            "ejecución con la URL pública. La Figura 1 resume el flujo completo.")
    figura(doc, "Flujo DevOps y pipeline de CI/CD de la EA3", fig_flujo,
           "Elaboración propia. Arriba, el trabajo en ramas; abajo, la integración y el despliegue. Si el CI falla, "
           "el merge queda bloqueado y Render no despliega. El tablero estático de la EA2 se sigue publicando en "
           "GitHub Pages en paralelo.")
    parrafo(doc, "La documentación del flujo está en el repositorio: el README tiene insignias de estado de los tres "
            "workflows, los enlaces a la app, la tabla de rutas, el diagrama del pipeline en Mermaid y las "
            "instrucciones para correr la app con uvicorn o con Docker.")


def resultados(doc, D, R, M, fig_ejec):
    doc.add_heading("6. Resultados", 1).paragraph_format.page_break_before = True
    prod = D["produccion"]
    parrafo(doc, f"Los datos de esta sección salen de la API de GitHub, de Docker y de consultas a la app en "
            f"producción hechas el {fecha_larga(D['fecha'])}, con el script informe/recolectar_datos_ea3.py.")
    indicadores(doc, [(M["prs_fusionados"], "pull requests fusionados"), (M["ejecuciones"], "ejecuciones de Actions"),
                      (M["pruebas"], "pruebas automatizadas"), (pct(M["cobertura"], 0), "de cobertura"),
                      (f"{M['ci_s']} s", "duración del CI en main")])

    doc.add_heading("6.1 La aplicación en producción", 2)
    parrafo(doc, f"La app está publicada en {prod['url']}. Al consultar /salud respondió HTTP {prod['estado_http']} "
            f"con el commit {prod['salud']['commit'][:7]} y el entorno «{prod['salud']['entorno']}». Ese commit es "
            f"exactamente el último de main ({D['main']}). El notebook comprueba además que el resumen que publica "
            "producción es idéntico, campo por campo, al del repositorio.")
    figura(doc, "Tablero de la aplicación desplegada en Render", recorte("10_app_tablero.png", alto=1150),
           "Captura de https://ea3-galaxy-a55.onrender.com. El pie del tablero indica el commit desplegado.",
           ancho=14)
    figura(doc, "Documentación OpenAPI generada por FastAPI", recorte("11_app_docs.png", alto=760),
           "Captura de /docs en producción. Cada ruta se puede probar desde el navegador.", ancho=14)
    lat = prod["latencia_ms"]
    tabla(doc, "Latencia de la app en producción", ["Ruta", "Latencia mediana"],
          [(r, f"{v} ms") for r, v in lat.items()], anchos=[9, 4],
          nota_txt="Mediana de 5 peticiones desde Medellín, con el servicio ya activo. El plan gratuito apaga el "
                   "servicio tras unos minutos sin tráfico y la primera petición puede tardar unos 50 segundos.")

    doc.add_heading("6.2 Ejecuciones del pipeline", 2)
    ejec = [e for e in D["ejecuciones"] if e["conclusion"] in ("success", "failure")]
    por_wf = {}
    for e in ejec:
        por_wf.setdefault((e["workflow"], e["conclusion"]), []).append(e["duracion_s"])
    nombres = {"CI": "CI", "Deploy": "Deploy (GitHub Pages)", "Deploy Render": "Deploy Render"}
    filas = [(nombres.get(w, w), "Éxito" if c == "success" else "Fallo", len(v), duracion(sorted(v)[len(v) // 2]),
              duracion(max(v))) for (w, c), v in sorted(por_wf.items())]
    tabla(doc, "Ejecuciones de GitHub Actions por workflow", ["Workflow", "Resultado", "Ejecuciones", "Mediana",
                                                              "Máximo"], filas, anchos=[5.2, 2.4, 2.6, 3.0, 3.0])
    figura(doc, "Duración de cada ejecución de GitHub Actions", fig_ejec,
           "Elaboración propia con datos de la API de GitHub. Las dos primeras corresponden a la creación del "
           "repositorio de la EA3 a partir del de la EA2.")
    jobs = D["jobs_ci"]["jobs"]
    parrafo(doc, "Gracias a que los jobs corren en paralelo, agregar la imagen Docker no alargó el CI. En la última "
            "ejecución sobre main los jobs tardaron " + ", ".join(f"{j['nombre']} {j['duracion_s']} s" for j in jobs
                                                                  if j["nombre"] != "ci")
            + ", y el total quedó alrededor de un minuto.")
    figura(doc, "Ejecución del workflow CI sobre main", recorte("04_ci_main.png", alto=720),
           "Captura de GitHub Actions: los jobs de pruebas y de la imagen Docker corren en paralelo y el job «ci» "
           "los agrupa.", ancho=14.5)

    doc.add_heading("6.3 Despliegue continuo verificado", 2)
    parrafo(doc, f"El ciclo completo se probó con el merge del PR #9. El CI pasó en main, Render registró el evento "
            f"«New commit via Auto-Deploy», construyó la imagen y la puso en línea en alrededor de un minuto, y el "
            f"workflow Deploy Render confirmó en su tercer intento que /salud ya reportaba el commit "
            f"{prod['salud']['commit'][:7]}. Desde el merge hasta la producción verificada pasaron "
            f"{duracion(M['merge_a_prod_s'])}, sin ninguna intervención manual.")
    tabla(doc, "Eventos de despliegue en Render", ["Commit", "Evento", "Inicio", "En línea"],
          [(e["commit"], e["evento"], e["inicio"], e["en_linea"]) for e in D["eventos_render"]],
          anchos=[2.2, 9.0, 2.4, 2.4],
          nota_txt="Copiado del panel Events del servicio en Render (hora de Colombia). El PR #11 no generó ningún "
                   "evento: Render no despliega commits con checks fallidos.")
    figura(doc, "Panel de eventos del servicio ea3-galaxy-a55 en Render", recorte("13_render_eventos.png"),
           "Captura del panel de Render: servicio Docker, plan Free, gestionado por Blueprint y en estado Live.",
           ancho=11)
    figura(doc, "Ejecución del workflow Deploy Render", recorte("05_deploy_render.png", alto=760),
           "Captura de GitHub Actions. El job se ejecuta en el environment «produccion», que enlaza a la URL "
           "pública.", ancho=14.5)

    doc.add_heading("6.4 Prueba de fallo controlada", 2)
    demo = next(p for p in D["pull_requests"] if p["numero"] == 11)
    parrafo(doc, "Para comprobar que el pipeline protege producción abrí el PR #11, que saca a propósito "
            "a55_limpio.csv de la imagen Docker. Es un error realista: olvidar un archivo en el Dockerfile o en el "
            f".dockerignore. El resultado fue {demo['checks_ok']} de {demo['checks_total']} checks en verde. Las "
            "pruebas de pytest pasaron en Python 3.12 y 3.13, porque usan el repositorio completo y no ven el "
            "error. En cambio, el job Imagen Docker falló: el contenedor arrancó y /salud respondió, pero "
            "/api/publicaciones devolvió 503. El check «ci» quedó en rojo, GitHub marcó el PR como bloqueado y "
            "Render no registró ningún despliegue. Lo cerré sin fusionar.")
    figura(doc, "Ejecución fallida del CI en el PR de demostración", recorte("07_ci_fallido.png", alto=900),
           "Captura de GitHub Actions: el job «Imagen Docker» y el job «ci» fallan; las pruebas de Python pasan.",
           ancho=14.5)
    destacado(doc, "Lección", "Probar la imagen, y no solo el código, fue lo que detectó el error. Sin el job docker, "
              "este cambio habría pasado el CI y la app en producción habría respondido 503 a todas las consultas "
              "de publicaciones.")

    doc.add_heading("6.5 Dificultades y decisiones", 2)
    vinetas(doc, [
        ("Tokens de despliegue. ", "El Deploy Hook de Render es una URL secreta. Para no depender de guardarla en "
                                   "GitHub cambié a autoDeployTrigger: checksPass (PR #8), sin perder la garantía "
                                   "de que solo se despliega lo que pasó el CI."),
        ("Arranque en frío. ", "El plan gratuito apaga el servicio sin tráfico. Por eso la verificación del "
                               "despliegue espera hasta 15 minutos y reintenta cada 10 segundos."),
        ("Permisos en el contenedor. ", "config.py crea la carpeta data/raw al importarse. Con el usuario sin "
                                        "privilegios fallaba, así que el Dockerfile la crea y le asigna el dueño "
                                        "antes de cambiar de usuario."),
        ("Mismo tablero, dos destinos. ", "En lugar de duplicar el HTML, agregué un parámetro a construir() para "
                                          "que la app indique el commit de Render, sin cambiar el sitio de Pages."),
    ])


def conclusiones(doc, M):
    doc.add_heading("7. Conclusiones", 1)
    parrafo(doc, "La evidencia cumplió su objetivo: el análisis del Galaxy A55 dejó de ser un archivo o un sitio "
            "estático y pasó a ser un servicio web en la nube. Ese servicio se prueba, se empaqueta y se despliega "
            "solo, y cada cambio deja rastro desde el issue hasta producción. Con base en lo que medí, saco cinco "
            "conclusiones.")
    vinetas(doc, [
        ("Reutilizar antes que reescribir. ", "La app es un archivo de unas 120 líneas porque reutiliza config.py y "
                                              "06_sitio.py. Una sola fuente de verdad para el tablero evita que la "
                                              "versión web y la de Pages se desalineen."),
        ("Hay que probar el artefacto, no solo el código. ", "La prueba de fallo controlada pasó todas las pruebas "
                                                             "unitarias y solo la detectó el smoke test del "
                                                             "contenedor. En un pipeline de despliegue, construir y "
                                                             "ejecutar la imagen es una prueba más."),
        ("Un despliegue se verifica desde afuera. ", "Exponer el commit en /salud permite que el pipeline confirme "
                                                     "qué versión sirve producción. «El workflow terminó» no es lo "
                                                     "mismo que «producción sirve lo que validé»."),
        ("La seguridad también es diseño. ", "Con checksPass no hay tokens de Render en GitHub, el contenedor no "
                                             "corre como root y la imagen no lleva el HTML capturado ni "
                                             "herramientas que no necesita."),
        ("La velocidad no se sacrificó. ", f"El CI tarda alrededor de {M['ci_s']} segundos y el ciclo completo, "
                                           f"de merge a producción verificada, {duracion(M['merge_a_prod_s'])}. "
                                           "Es la entrega pequeña y frecuente que proponen Forsgren et al. (2018)."),
    ])
    parrafo(doc, "Como limitaciones, los datos siguen siendo la captura del 25 de octubre: el CI no puede hacer "
            "scraping en vivo sin evadir la protección del sitio. El plan gratuito de Render tiene arranque en frío. "
            "Y no hay un entorno de staging previo a producción. Para el proyecto de cierre propongo programar la "
            "captura con Selenium en un equipo propio, versionar los datos con fecha para construir una serie de "
            "tiempo de precios y agregar un entorno de pruebas con las vistas previas de pull request que ofrece "
            "Render.")


def bibliografia(doc):
    doc.add_heading("Referencias", 1)
    refs = [
        ("Chacon, S., y Straub, B. (2014). ", "Pro Git", " (2.ª ed.). Apress. https://git-scm.com/book/es/v2"),
        ("Cuervo Quintero, J. N. (2026). ", "Flujo de trabajo DevOps para un proyecto de web scraping: Git, GitHub y "
         "CI/CD con GitHub Actions sobre el scraper del Samsung Galaxy A55 en Mercado Libre",
         " [Evidencia de aprendizaje 2, Programación para Análisis de Datos]. Institución Universitaria Digital de "
         "Antioquia."),
        ("Cuervo Quintero, J. N., y Cuervo Quintero, Y. M. (2026). ", "Web scraping en el comercio electrónico: "
         "precio y oferta del Samsung Galaxy A55 en Mercado Libre Colombia",
         " [Evidencia de aprendizaje 1, Programación para Análisis de Datos]. Institución Universitaria Digital "
         "de Antioquia."),
        ("Docker Inc. (2026). ", "Dockerfile reference", ". https://docs.docker.com/reference/dockerfile/"),
        ("Forsgren, N., Humble, J., y Kim, G. (2018). ", "Accelerate: The science of lean software and DevOps",
         ". IT Revolution Press."),
        ("GitHub. (2026a). ", "Documentación de GitHub Actions", ". https://docs.github.com/es/actions"),
        ("GitHub. (2026b). ", "Eventos que desencadenan flujos de trabajo",
         ". https://docs.github.com/es/actions/reference/workflows-and-actions/events-that-trigger-workflows"),
        ("GitHub. (2026c). ", "Flujo de GitHub", ". https://docs.github.com/es/get-started/using-github/github-flow"),
        ("Humble, J., y Farley, D. (2010). ", "Continuous delivery: Reliable software releases through build, test, "
         "and deployment automation", ". Addison-Wesley."),
        ("Kim, G., Humble, J., Debois, P., Willis, J., y Forsgren, N. (2021). ", "The DevOps handbook: How to "
         "create world-class agility, reliability, and security in technology organizations",
         " (2.ª ed.). IT Revolution Press."),
        ("Krotov, V., Johnson, L., y Silva, L. (2020). Tutorial: Legality and ethics of web scraping. ",
         "Communications of the Association for Information Systems, 47", ", 539–563. "
         "https://doi.org/10.17705/1CAIS.04724"),
        ("MercadoLibre, Inc. (2025). ", "Annual report (Form 10-K) for the fiscal year ended December 31, 2024",
         ". U.S. Securities and Exchange Commission. https://investor.mercadolibre.com/"),
        ("Mitchell, R. (2024). ", "Web scraping with Python: Data extraction from the modern web",
         " (3.ª ed.). O'Reilly Media."),
        ("pytest-dev. (2026). ", "pytest documentation", ". https://docs.pytest.org/"),
        ("Ramírez, S. (2026). ", "FastAPI documentation", ". https://fastapi.tiangolo.com/"),
        ("Render. (2026a). ", "Blueprint YAML reference", ". https://render.com/docs/blueprint-spec"),
        ("Render. (2026b). ", "Deploying on Render", ". https://render.com/docs/deploys"),
        ("Samsung Electronics. (2024). ", "Galaxy A55 5G: Especificaciones",
         ". https://www.samsung.com/co/smartphones/galaxy-a/galaxy-a55-5g/"),
        ("Wiggins, A. (2017). ", "The twelve-factor app", ". https://12factor.net/"),
    ]
    for antes, cursiva, despues in refs:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(1.27)
        p.paragraph_format.first_line_indent = Cm(-1.27)
        p.add_run(antes)
        p.add_run(cursiva).italic = True
        p.add_run(despues)


def anexo(doc, D):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    doc.add_heading("Anexo A. Código fuente e implementación", 1)
    parrafo(doc, f"Todo el código está en {D['url_repo']} (rama main, commit {D['main']}), en el archivo ZIP y en el "
            "notebook que acompañan este informe. Estructura del repositorio:")
    codigo(doc, [
        "EA3-DevOps-Scraping-GalaxyA55/",
        "├── app/main.py                    app web FastAPI (tablero + API)",
        "├── Dockerfile, .dockerignore      imagen de producción",
        "├── render.yaml                    Blueprint de Render (checksPass)",
        "├── .github/workflows/",
        "│   ├── ci.yml                     pruebas, pipeline de datos e imagen Docker",
        "│   ├── deploy-render.yml          verificación del despliegue en Render",
        "│   └── deploy.yml                 tablero estático en GitHub Pages (EA2)",
        "├── scraping/                      scraper de la EA1 (Selenium, BeautifulSoup, Scrapy, análisis)",
        "├── tests/                         56 pruebas (pytest), incluido test_app.py",
        "├── data/, figuras/                datos procesados y gráficas",
        "├── informe/                       recolección de datos, capturas, notebook e informe",
        "├── requirements.txt, requirements-app.txt, requirements-dev.txt",
        "└── README.md, CONTRIBUTING.md, pyproject.toml",
    ], tam=8.5)
    doc.add_paragraph()
    doc.add_heading("Cómo reproducir el proyecto", 2)
    codigo(doc, [
        f"git clone {D['url_repo']}.git && cd EA3-DevOps-Scraping-GalaxyA55",
        "python3 -m venv .venv && source .venv/bin/activate",
        "pip install -r requirements-dev.txt",
        "ruff check . && pytest --cov",
        "uvicorn app.main:app --reload                  # http://127.0.0.1:8000",
        "docker build -t ea3-galaxy-a55 . && docker run --rm -p 8000:8000 ea3-galaxy-a55",
    ], tam=8.5)
    doc.add_paragraph()
    parrafo(doc, "A continuación va el código de los archivos creados en esta evidencia. Los scripts de scraping "
            "heredados de la EA1 y la EA2 están completos en el repositorio y en el ZIP.")
    for rel in ["app/main.py", "Dockerfile", ".dockerignore", "requirements-app.txt", "render.yaml",
                ".github/workflows/ci.yml", ".github/workflows/deploy-render.yml", "tests/test_app.py"]:
        bloque_codigo(doc, rel)

    doc.add_heading("Anexo B. Enlaces", 1)
    vinetas(doc, [
        ("Repositorio: ", D["url_repo"]),
        ("App en producción (Render): ", D["produccion"]["url"]),
        ("Documentación de la API: ", D["produccion"]["url"] + "/docs"),
        ("Tablero estático (GitHub Pages): ", D["url_pages"]),
        ("Ejecuciones de GitHub Actions: ", D["url_repo"] + "/actions"),
        ("Pull requests: ", D["url_repo"] + "/pulls?q=is%3Apr"),
    ])


def main():
    D = json.loads((INFORME / "datos_ea3.json").read_text(encoding="utf-8"))
    R = json.loads((RAIZ / "data" / "resumen.json").read_text(encoding="utf-8"))
    ejec = [e for e in D["ejecuciones"] if e["conclusion"] in ("success", "failure")]
    # Tiempo de merge a producción verificada: inicio del CI en main hasta el fin de Deploy Render del mismo commit
    render = [e for e in ejec if e["workflow"] == "Deploy Render" and e["rama"] == "main" and e["sha"] == D["main"]]
    ci_main = [e for e in ejec if e["workflow"] == "CI" and e["evento"] == "push" and e["sha"] == D["main"]]
    merge_a_prod = round(segundos_entre(ci_main[0]["inicio"], render[-1]["fin"])) if render and ci_main else 0
    M = {
        "prs": len(D["pull_requests"]),
        "prs_fusionados": sum(p["fusionado"] for p in D["pull_requests"]),
        "issues": len(D["issues"]),
        "ejecuciones": len(ejec),
        "pruebas": sum(D["pruebas_por_archivo"].values()),
        "cobertura": D["cobertura"]["total"],
        "ci_s": ci_main[0]["duracion_s"] if ci_main else "—",
        "merge_a_prod_s": merge_a_prod,
    }
    hoy = datetime.now()
    fecha = f"{hoy.day} de {MESES[hoy.month - 1]} de {hoy.year}"

    doc = Document()
    estilos(doc)
    encabezado_y_pie(doc.sections[0])
    portada(doc, fecha)
    resumen(doc, D, R, M)
    introduccion(doc)
    descripcion(doc, R)
    tema(doc)
    objetivos(doc)
    metodologia(doc, D, R, M, figura_flujo())
    resultados(doc, D, R, M, figura_ejecuciones(D["ejecuciones"]))
    conclusiones(doc, M)
    bibliografia(doc)
    anexo(doc, D)
    SALIDA.parent.mkdir(exist_ok=True)
    doc.save(SALIDA)
    print(f"Informe generado: {SALIDA}")


def segundos_entre(inicio, fin):
    a = datetime.fromisoformat(inicio.replace("Z", "+00:00"))
    b = datetime.fromisoformat(fin.replace("Z", "+00:00"))
    return (b - a).total_seconds()


if __name__ == "__main__":
    main()
