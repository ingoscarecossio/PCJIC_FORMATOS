from __future__ import annotations

import io
import math
import os
import base64
import sqlite3
import hashlib
import hmac
import secrets
import json
import re
import shutil
import subprocess
import tempfile
from copy import deepcopy
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple


# Dependencia opcional para despliegue productivo con PostgreSQL/Supabase.
# La app sigue funcionando en modo local con SQLite si no se configura DATABASE_URL.
try:
    import psycopg2
    from psycopg2.extras import RealDictCursor
except Exception:  # pragma: no cover - opcional en modo local
    psycopg2 = None
    RealDictCursor = None

import pandas as pd
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

APP_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = APP_DIR / "plantilla"
TEMPLATE_GC72 = TEMPLATE_DIR / "FD-GC72-Informe_Academico.docx"
TEMPLATE_GC71 = TEMPLATE_DIR / "FD-GC71.docx"
LOGO_POLI = TEMPLATE_DIR / "logo_poli.png"
LOGO_ICONTEC = TEMPLATE_DIR / "logo_icontec.png"
DATA_DIR = APP_DIR / "app_data"
DB_PATH = DATA_DIR / "fdgc_app.sqlite3"

# -----------------------------------------------------------------------------
# Configuración productiva / Streamlit Cloud
# -----------------------------------------------------------------------------
# Modo recomendado en producción:
#   APP_ENV="production"
#   DATABASE_URL="postgresql://usuario:clave@host:5432/base"
#   INITIAL_ADMIN_USER / INITIAL_ADMIN_PASSWORD en st.secrets o variables de entorno.
# Si DATABASE_URL no existe, la app funciona con SQLite local para desarrollo o demo.

APP_VERSION = "4.0.0-cloud-production"
DEFAULT_MAX_EVIDENCE_MB = 15


def _get_secret_value(name: str, default: Optional[str] = None) -> Optional[str]:
    """Lee configuración desde Streamlit secrets o variables de entorno sin romper ejecución local."""
    try:
        import streamlit as _st
        if hasattr(_st, "secrets") and name in _st.secrets:
            return _st.secrets.get(name)
    except Exception:
        pass
    return os.getenv(name, default)


def get_app_env() -> str:
    return str(_get_secret_value("APP_ENV", os.getenv("APP_ENV", "local")) or "local").lower().strip()


def get_database_url() -> str:
    return str(_get_secret_value("DATABASE_URL", os.getenv("DATABASE_URL", "")) or "").strip()


def usar_postgres() -> bool:
    url = get_database_url().lower()
    return url.startswith("postgres://") or url.startswith("postgresql://")


def postgres_url_normalizada() -> str:
    # Algunos proveedores entregan postgres:// y psycopg2 espera postgresql:// en ciertos entornos.
    url = get_database_url()
    if url.startswith("postgres://"):
        return "postgresql://" + url[len("postgres://"):]
    return url


def _traducir_sql(sql: str) -> str:
    """Traduce placeholders SQLite (?) a PostgreSQL (%s) para consultas parametrizadas simples."""
    return sql.replace("?", "%s") if usar_postgres() else sql


class PgConnectionAdapter:
    """Adaptador mínimo para conservar la API usada por la app con SQLite."""

    def __init__(self):
        if psycopg2 is None:
            raise RuntimeError("psycopg2-binary no está instalado. Revise requirements.txt.")
        self._conn = psycopg2.connect(postgres_url_normalizada())

    def execute(self, sql: str, params: Tuple[Any, ...] = tuple()):
        cur = self._conn.cursor(cursor_factory=RealDictCursor)
        cur.execute(_traducir_sql(sql), params or tuple())
        return cur

    def cursor(self):
        return self._conn.cursor(cursor_factory=RealDictCursor)

    def commit(self):
        self._conn.commit()

    def rollback(self):
        self._conn.rollback()

    def close(self):
        self._conn.close()


def conexion_db():
    """Conexión unificada. PostgreSQL si DATABASE_URL está configurado; SQLite en local/demo."""
    if usar_postgres():
        return PgConnectionAdapter()
    DATA_DIR.mkdir(exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def read_sql_df(sql: str, params: Tuple[Any, ...] = tuple()) -> pd.DataFrame:
    """Lectura robusta a DataFrame para SQLite/PostgreSQL."""
    conn = conexion_db()
    try:
        if usar_postgres():
            cur = conn.execute(sql, params)
            rows = cur.fetchall()
            return pd.DataFrame([dict(r) for r in rows])
        return pd.read_sql_query(sql, conn, params=params)
    finally:
        conn.close()


def db_execute(sql: str, params: Tuple[Any, ...] = tuple(), fetchone: bool = False, fetchall: bool = False):
    conn = conexion_db()
    try:
        cur = conn.execute(sql, params)
        result = None
        if fetchone:
            row = cur.fetchone()
            result = dict(row) if row is not None and usar_postgres() else row
        elif fetchall:
            rows = cur.fetchall()
            result = [dict(r) for r in rows] if usar_postgres() else rows
        conn.commit()
        return result
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        raise
    finally:
        conn.close()


def _safe_int_secret(name: str, default: int) -> int:
    try:
        return int(_get_secret_value(name, str(default)) or default)
    except Exception:
        return default


def max_evidence_bytes() -> int:
    return _safe_int_secret("MAX_EVIDENCE_MB", DEFAULT_MAX_EVIDENCE_MB) * 1024 * 1024


def initial_admin_config() -> Tuple[str, str, str, str]:
    """Credenciales de arranque parametrizables por secretos. No las deje quemadas en producción."""
    return (
        str(_get_secret_value("INITIAL_ADMIN_USER", "admin") or "admin"),
        str(_get_secret_value("INITIAL_ADMIN_PASSWORD", "Admin123*") or "Admin123*"),
        str(_get_secret_value("INITIAL_ADMIN_NAME", "Administrador del sistema") or "Administrador del sistema"),
        str(_get_secret_value("INITIAL_ADMIN_EMAIL", "") or ""),
    )


def add_column_if_missing(table: str, column: str, ddl_type: str):
    """Migración defensiva compatible con SQLite/PostgreSQL."""
    conn = conexion_db()
    try:
        if usar_postgres():
            cur = conn.execute(
                "SELECT column_name FROM information_schema.columns WHERE table_name=%s AND column_name=%s",
                (table, column),
            )
            exists = cur.fetchone() is not None
            if not exists:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {ddl_type}")
        else:
            cur = conn.execute(f"PRAGMA table_info({table})")
            cols = [r[1] for r in cur.fetchall()]
            if column not in cols:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {ddl_type}")
        conn.commit()
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        # No se detiene la app por una migración idempotente fallida; se mostrará en diagnóstico.
    finally:
        conn.close()


def health_status() -> Dict[str, Any]:
    """Diagnóstico simple para saber si el despliegue está sano."""
    status = {
        "version": APP_VERSION,
        "app_env": get_app_env(),
        "database": "PostgreSQL" if usar_postgres() else "SQLite local",
        "database_ok": False,
        "templates_ok": TEMPLATE_GC71.exists() and TEMPLATE_GC72.exists(),
        "storage": "DB + local cache" if usar_postgres() else "SQLite/local",
        "max_evidence_mb": _safe_int_secret("MAX_EVIDENCE_MB", DEFAULT_MAX_EVIDENCE_MB),
    }
    try:
        row = db_execute("SELECT COUNT(*) AS n FROM usuarios", fetchone=True)
        status["database_ok"] = row is not None
    except Exception as exc:
        status["database_error"] = str(exc)[:300]
    return status


DIAS_MAP = {
    "Lunes": 0,
    "Martes": 1,
    "Miércoles": 2,
    "Miercoles": 2,
    "Jueves": 3,
    "Viernes": 4,
    "Sábado": 5,
    "Sabado": 5,
    "Domingo": 6,
}
DIAS_INV = {v: k for k, v in DIAS_MAP.items() if "é" in k or k not in ("Miercoles", "Sabado")}

# -----------------------------------------------------------------------------
# FD-GC72 - Informe académico
# -----------------------------------------------------------------------------
COLUMNAS_GC72 = [
    "Código",
    "Grupo",
    "Asignatura",
    "% Avance en contenido",
    "% Evaluado",
    "Estudiantes matriculados",
    "Desertaron N°",
    "Desertaron %",
    "Aprueban evaluación parcial N°",
    "Aprueban evaluación parcial %",
    "Reprueban evaluación parcial N°",
    "Reprueban evaluación parcial %",
    "Aprueban a la fecha N°",
    "Aprueban a la fecha %",
    "Reprueban a la fecha N°",
    "Reprueban a la fecha %",
]
NUMERICAS_ENTERAS_GC72 = [
    "Estudiantes matriculados",
    "Desertaron N°",
    "Aprueban evaluación parcial N°",
    "Reprueban evaluación parcial N°",
    "Aprueban a la fecha N°",
    "Reprueban a la fecha N°",
]
NUMERICAS_PORCENTAJE_GC72 = [
    "% Avance en contenido",
    "% Evaluado",
    "Desertaron %",
    "Aprueban evaluación parcial %",
    "Reprueban evaluación parcial %",
    "Aprueban a la fecha %",
    "Reprueban a la fecha %",
]
MAPA_TABLA_WORD_GC72 = COLUMNAS_GC72[:]

PREFORMAS_GC72: Dict[str, Dict[str, str]] = {
    "aspectos_positivos": {
        "Participación activa": "Se evidenció participación activa y disposición del grupo para el desarrollo de las actividades académicas propuestas.",
        "Avance adecuado": "El curso presentó un avance adecuado frente a los contenidos programados, manteniendo coherencia entre las actividades de clase y los resultados esperados.",
        "Aplicación práctica": "Los estudiantes lograron relacionar los conceptos abordados con situaciones prácticas, favoreciendo la comprensión aplicada de la asignatura.",
        "Mejora progresiva": "Se observó una mejora progresiva en la apropiación de los temas, especialmente en las actividades de seguimiento y retroalimentación.",
        "Buen cumplimiento": "La mayoría de los estudiantes cumplió con las actividades evaluativas y entregables definidos para el periodo reportado.",
        "Trabajo colaborativo": "El trabajo colaborativo fortaleció la discusión académica y permitió resolver dudas de manera más efectiva durante el curso.",
    },
    "inconvenientes": {
        "Inasistencia intermitente": "Se presentaron dificultades asociadas a inasistencia intermitente de algunos estudiantes, lo que afectó la continuidad del proceso formativo.",
        "Entregas tardías": "Algunos estudiantes realizaron entregas fuera de los tiempos establecidos, situación que limitó la retroalimentación oportuna.",
        "Brechas conceptuales": "Se identificaron brechas conceptuales en temas base, por lo cual fue necesario reforzar contenidos previos para avanzar con mayor solidez.",
        "Baja participación puntual": "En algunas sesiones se presentó baja participación de parte del grupo, especialmente en actividades que requerían preparación previa.",
        "Dificultades técnicas": "Se presentaron dificultades técnicas o de acceso a herramientas requeridas para el desarrollo de algunas actividades académicas.",
        "Carga académica acumulada": "La acumulación de actividades de otros espacios académicos incidió en el ritmo de trabajo y en la oportunidad de algunas entregas.",
    },
    "propuestas": {
        "Seguimiento formativo": "Fortalecer el seguimiento formativo mediante actividades cortas de verificación, retroalimentación temprana y acompañamiento focalizado.",
        "Talleres aplicados": "Incorporar talleres aplicados por unidades temáticas para consolidar la relación entre teoría, práctica y evaluación.",
        "Nivelación inicial": "Implementar una actividad diagnóstica y espacios de nivelación para atender brechas conceptuales antes de abordar contenidos de mayor complejidad.",
        "Rúbricas claras": "Socializar rúbricas y criterios de evaluación desde el inicio de cada actividad, con el fin de mejorar la calidad de las entregas.",
        "Aprendizaje basado en problemas": "Utilizar ejercicios basados en problemas reales del contexto profesional para incrementar la pertinencia y motivación del proceso formativo.",
        "Alertas tempranas": "Aplicar alertas tempranas frente a inasistencia, bajo desempeño o entregas pendientes, articulando acciones de mejora con los estudiantes.",
    },
}

# -----------------------------------------------------------------------------
# FD-GC71 - Guía didáctica
# -----------------------------------------------------------------------------
COLUMNAS_MODULOS = [
    "Unidad",
    "Contenido / tema central",
    "Horas presenciales",
    "Sesiones",
    "Trabajo presencial",
    "Trabajo independiente",
]
COLUMNAS_HORARIOS = ["Día", "Hora inicio", "Hora fin", "Lugar / ambiente"]
COLUMNAS_EVALUACIONES = [
    "Tipo de evaluación",
    "Procedimiento de evaluación",
    "Valor (%)",
    "Fecha de realización",
    "Unidad relacionada",
    "Corte",
]
COLUMNAS_SESIONES = [
    "Unidad",
    "N° sesión",
    "Fecha",
    "Horario",
    "Contenido por desarrollar",
    "Descripción del trabajo presencial",
    "Descripción trabajo independiente",
    "Lugar / ambiente",
]

TEXTOS_PREDEFINIDOS_GC71 = {
    "justificacion": "La asignatura aporta a la formación académica y profesional mediante la articulación entre fundamentos conceptuales, aplicación práctica y análisis de situaciones propias del campo disciplinar. Su desarrollo favorece la comprensión de problemas del contexto, el uso de herramientas técnicas y la toma de decisiones sustentada en criterios académicos, éticos y profesionales.",
    "competencias": "La asignatura tributa al desarrollo de competencias asociadas con el análisis crítico, la resolución de problemas, la comunicación técnica, el trabajo colaborativo y la aplicación de conocimientos disciplinares en escenarios reales o simulados.",
    "resultados": "Al finalizar la asignatura, el estudiante estará en capacidad de reconocer, interpretar y aplicar los conceptos y procedimientos centrales del curso, integrando evidencias, criterios técnicos y estrategias de solución acordes con los resultados de aprendizaje del programa.",
    "objetivo_general": "Desarrollar en el estudiante capacidades conceptuales, metodológicas y prácticas para comprender y aplicar los contenidos de la asignatura en situaciones propias de su formación académica y profesional.",
    "objetivos_especificos": "1. Reconocer los fundamentos conceptuales de la asignatura.\n2. Aplicar procedimientos y herramientas propias del área de formación.\n3. Analizar casos, ejercicios o problemas relacionados con el contexto profesional.\n4. Comunicar resultados de manera clara, ordenada y técnicamente sustentada.",
    "metodologias": "La asignatura se desarrollará mediante clases orientadoras, talleres aplicados, análisis de casos, ejercicios prácticos, aprendizaje basado en problemas, socialización de avances y retroalimentación permanente. Se promoverá la participación activa del estudiante y la integración entre trabajo presencial e independiente.",
    "ambientes": "Aula de clase, plataforma institucional, recursos digitales de apoyo y, cuando aplique, laboratorios, salidas pedagógicas o ambientes especializados requeridos para el logro de los resultados de aprendizaje.",
    "medios": "Presentaciones, guías de clase, material bibliográfico, bases de datos académicas, plataforma virtual, software especializado cuando aplique, tablero, equipos audiovisuales y recursos institucionales necesarios para el desarrollo de las actividades.",
    "referencias": "Bibliografía básica y complementaria definida por el docente, documentos institucionales del programa, artículos académicos recientes, recursos digitales especializados y fuentes en segunda lengua cuando sean pertinentes para la asignatura.",
}

# -----------------------------------------------------------------------------
# Utilidades generales
# -----------------------------------------------------------------------------
def df_vacio(columnas: List[str], filas: int = 5) -> pd.DataFrame:
    return pd.DataFrame([{c: "" for c in columnas} for _ in range(filas)])


def limpiar_numero(valor) -> Optional[float]:
    if valor is None or pd.isna(valor):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip().replace("%", "").replace(",", ".")
    if texto == "":
        return None
    try:
        return float(texto)
    except ValueError:
        return None


def formato_entero(valor) -> str:
    numero = limpiar_numero(valor)
    if numero is None:
        return ""
    return str(int(round(numero)))


def formato_porcentaje(valor) -> str:
    numero = limpiar_numero(valor)
    if numero is None:
        return ""
    return f"{int(round(numero))}%"


def porcentaje(numerador, denominador) -> str:
    n = limpiar_numero(numerador)
    d = limpiar_numero(denominador)
    if n is None or d in (None, 0):
        return ""
    return formato_porcentaje((n / d) * 100)


def limpiar_df(df: pd.DataFrame, columnas: List[str]) -> pd.DataFrame:
    df = df.copy()
    for col in columnas:
        if col not in df.columns:
            df[col] = ""
    df = df[columnas]
    mask = df.apply(lambda row: any(str(v).strip() for v in row.values if not pd.isna(v)), axis=1)
    return df[mask].reset_index(drop=True)


def normalizar_texto(texto: str) -> str:
    texto = str(texto or "").strip().upper()
    reemplazos = {
        "Á": "A",
        "É": "E",
        "Í": "I",
        "Ó": "O",
        "Ú": "U",
        "Ü": "U",
        "Ñ": "N",
    }
    for a, b in reemplazos.items():
        texto = texto.replace(a, b)
    texto = re.sub(r"\s+", " ", texto)
    return texto


def nombre_archivo_seguro(nombre: str, fecha: date | str, prefijo: str) -> str:
    base = re.sub(r"[^A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9]+", "_", str(nombre).strip()).strip("_") or "docente"
    f = fecha.strftime("%Y%m%d") if hasattr(fecha, "strftime") else re.sub(r"\D+", "", str(fecha))
    return f"{prefijo}_{base}_{f}"


def convertir_doc_si_es_posible(docx_bytes: bytes, nombre_base: str) -> Optional[bytes]:
    ejecutable = shutil.which("soffice") or shutil.which("libreoffice")
    if not ejecutable:
        return None
    with tempfile.TemporaryDirectory() as td:
        td_path = Path(td)
        docx_path = td_path / f"{nombre_base}.docx"
        doc_path = td_path / f"{nombre_base}.doc"
        docx_path.write_bytes(docx_bytes)
        subprocess.run(
            [ejecutable, "--headless", "--convert-to", "doc", "--outdir", str(td_path), str(docx_path)],
            check=False,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=60,
        )
        if doc_path.exists():
            return doc_path.read_bytes()
    return None

# -----------------------------------------------------------------------------
# Utilidades Word
# -----------------------------------------------------------------------------
def set_cell_width(cell, width_inches: float):
    cell.width = Inches(width_inches)
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.find(qn("w:tcW"))
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)
    tc_w.set(qn("w:w"), str(int(width_inches * 1440)))
    tc_w.set(qn("w:type"), "dxa")


def fijar_anchos_tabla(tabla, anchos: List[float]):
    tabla.autofit = False
    for row in tabla.rows:
        for idx, width in enumerate(anchos):
            if idx < len(row.cells):
                set_cell_width(row.cells[idx], width)


def shade_cell(cell, fill: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill.replace("#", ""))


def set_cell_text(
    cell,
    texto: str,
    font_size: float = 8.0,
    bold: bool = False,
    align=WD_ALIGN_PARAGRAPH.CENTER,
    color: Optional[str] = None,
):
    cell.text = ""
    parrafo = cell.paragraphs[0]
    parrafo.alignment = align
    parrafo.paragraph_format.space_after = Pt(0)
    run = parrafo.add_run(str(texto) if texto is not None else "")
    run.bold = bold
    run.font.size = Pt(font_size)
    if color:
        color = color.replace("#", "")
        run.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def add_paragraph_in_cell(cell, texto: str, font_size: float = 8.5, bold: bool = False, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = cell.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.05
    r = p.add_run(texto)
    r.bold = bold
    r.font.size = Pt(font_size)
    return p


def set_section_text(cell, titulo: str, texto: str, font_size: float = 8.5):
    cell.text = ""
    p1 = cell.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p1.add_run(titulo)
    r1.bold = True
    r1.font.size = Pt(9)
    if texto:
        for parte in str(texto).split("\n"):
            if parte.strip():
                add_paragraph_in_cell(cell, parte.strip(), font_size=font_size)


def set_label_value(cell, etiqueta: str, valor: str):
    cell.text = ""
    parrafo = cell.paragraphs[0]
    parrafo.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r1 = parrafo.add_run(etiqueta)
    r1.bold = True
    r1.font.size = Pt(10)
    r2 = parrafo.add_run(f" {valor}" if valor else "")
    r2.font.size = Pt(10)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def aplicar_bordes_tabla(tabla, color="000000", size="4"):
    tbl = tabla._tbl
    tbl_pr = tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = f"w:{edge}"
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)


def remover_parrafo(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None

# -----------------------------------------------------------------------------
# FD-GC72
# -----------------------------------------------------------------------------
def normalizar_dataframe_gc72(df: pd.DataFrame, calcular_porcentajes: bool = True) -> pd.DataFrame:
    df = limpiar_df(df, COLUMNAS_GC72)
    if calcular_porcentajes:
        for idx, row in df.iterrows():
            matriculados = row.get("Estudiantes matriculados", "")
            df.at[idx, "Desertaron %"] = porcentaje(row.get("Desertaron N°", ""), matriculados)
            df.at[idx, "Aprueban evaluación parcial %"] = porcentaje(row.get("Aprueban evaluación parcial N°", ""), matriculados)
            df.at[idx, "Reprueban evaluación parcial %"] = porcentaje(row.get("Reprueban evaluación parcial N°", ""), matriculados)
            df.at[idx, "Aprueban a la fecha %"] = porcentaje(row.get("Aprueban a la fecha N°", ""), matriculados)
            df.at[idx, "Reprueban a la fecha %"] = porcentaje(row.get("Reprueban a la fecha N°", ""), matriculados)
    return df


def texto_preformas(seleccionadas: Iterable[str], categoria: str, adicional: str = "") -> str:
    partes = [PREFORMAS_GC72[categoria][k] for k in seleccionadas if k in PREFORMAS_GC72[categoria]]
    adicional = (adicional or "").strip()
    if adicional:
        partes.append(adicional)
    return " ".join(partes).strip()


def curso_key(row, idx: int) -> str:
    codigo = re.sub(r"\W+", "_", str(row.get("Código", "")).strip())
    grupo = re.sub(r"\W+", "_", str(row.get("Grupo", "")).strip())
    asignatura = re.sub(r"\W+", "_", str(row.get("Asignatura", "")).strip())[:40]
    return f"curso_{idx}_{codigo}_{grupo}_{asignatura}"


def agregar_parrafo_antes(parrafo_referencia, texto: str = "", bold_prefix: Optional[str] = None, space_after: int = 3):
    nuevo = parrafo_referencia.insert_paragraph_before()
    nuevo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    nuevo.paragraph_format.space_after = Pt(space_after)
    nuevo.paragraph_format.line_spacing = 1.05
    if bold_prefix and texto.startswith(bold_prefix):
        r1 = nuevo.add_run(bold_prefix)
        r1.bold = True
        r1.font.size = Pt(10)
        r2 = nuevo.add_run(texto[len(bold_prefix):])
        r2.font.size = Pt(10)
    else:
        r = nuevo.add_run(texto)
        r.font.size = Pt(10)
    return nuevo


def construir_analisis(cursos: pd.DataFrame, analisis_por_curso: Dict[str, Dict[str, str]], modo: str) -> List[Dict[str, str]]:
    bloques = []
    if modo == "Consolidado institucional":
        positivos, inconvenientes, propuestas = [], [], []
        for data in analisis_por_curso.values():
            if data.get("positivos"):
                positivos.append(data["positivos"])
            if data.get("inconvenientes"):
                inconvenientes.append(data["inconvenientes"])
            if data.get("propuestas"):
                propuestas.append(data["propuestas"])
        bloques.append({
            "titulo": "Análisis consolidado de los cursos reportados",
            "positivos": " ".join(dict.fromkeys(positivos)),
            "inconvenientes": " ".join(dict.fromkeys(inconvenientes)),
            "propuestas": " ".join(dict.fromkeys(propuestas)),
        })
        return bloques

    for idx, row in cursos.iterrows():
        key = curso_key(row, idx)
        nombre = str(row.get("Asignatura", "")).strip() or "Curso sin nombre"
        codigo = str(row.get("Código", "")).strip()
        grupo = str(row.get("Grupo", "")).strip()
        detalles = []
        if codigo:
            detalles.append(f"Código {codigo}")
        if grupo:
            detalles.append(f"Grupo {grupo}")
        encabezado = f"{nombre} ({' - '.join(detalles)})" if detalles else nombre
        data = analisis_por_curso.get(key, {})
        bloques.append({
            "titulo": encabezado,
            "positivos": data.get("positivos", ""),
            "inconvenientes": data.get("inconvenientes", ""),
            "propuestas": data.get("propuestas", ""),
        })
    return bloques


def crear_informe_gc72_docx(docente: str, periodo: str, fecha_entrega: date | str, cursos: pd.DataFrame, bloques_analisis: List[Dict[str, str]]) -> bytes:
    if not TEMPLATE_GC72.exists():
        raise FileNotFoundError(f"No se encontró la plantilla: {TEMPLATE_GC72}")
    doc = Document(str(TEMPLATE_GC72))
    cursos = normalizar_dataframe_gc72(cursos, calcular_porcentajes=False)

    datos = doc.tables[0]
    set_label_value(datos.rows[0].cells[0], "DOCENTE:", docente)
    set_label_value(datos.rows[1].cells[0], "PERÍODO ACADÉMICO:", periodo)
    fecha_texto = fecha_entrega.strftime("%d/%m/%Y") if hasattr(fecha_entrega, "strftime") else str(fecha_entrega)
    set_label_value(datos.rows[2].cells[0], "FECHA DE ENTREGA:", fecha_texto)

    tabla = doc.tables[1]
    fijar_anchos_tabla(tabla, [0.62, 0.55, 1.38, 0.72, 0.68, 0.90, 0.43, 0.38, 0.55, 0.38, 0.55, 0.40, 0.50, 0.38, 0.50, 0.38])
    fila_inicio = 3
    filas_necesarias = max(8, len(cursos))
    while len(tabla.rows) < fila_inicio + filas_necesarias:
        tabla.add_row()

    for i in range(filas_necesarias):
        row_cells = tabla.rows[fila_inicio + i].cells
        if i < len(cursos):
            registro = cursos.iloc[i]
            for j, col in enumerate(MAPA_TABLA_WORD_GC72):
                valor = registro.get(col, "")
                if col in NUMERICAS_ENTERAS_GC72:
                    texto = formato_entero(valor)
                elif col in NUMERICAS_PORCENTAJE_GC72:
                    texto = formato_porcentaje(valor)
                    if j >= 7:
                        texto = texto.replace("%", "")
                else:
                    texto = "" if pd.isna(valor) else str(valor).strip()
                set_cell_text(row_cells[j], texto, font_size=6.8 if j >= 6 else 7.2)
        else:
            for j in range(min(len(row_cells), len(MAPA_TABLA_WORD_GC72))):
                set_cell_text(row_cells[j], "", font_size=7.5)

    firma = None
    for p in doc.paragraphs:
        if p.text.strip().startswith("Firma:"):
            firma = p
            break
    if firma is None:
        firma = doc.add_paragraph("Firma: ___________________")

    agregar_parrafo_antes(firma, "")
    for bloque in bloques_analisis:
        titulo = bloque.get("titulo", "Curso")
        p_titulo = firma.insert_paragraph_before()
        p_titulo.paragraph_format.space_before = Pt(4)
        p_titulo.paragraph_format.space_after = Pt(2)
        run = p_titulo.add_run(titulo)
        run.bold = True
        run.font.size = Pt(10)

        positivos = bloque.get("positivos", "").strip() or "Sin observaciones específicas para este apartado."
        inconvenientes = bloque.get("inconvenientes", "").strip() or "No se reportan inconvenientes relevantes para el periodo informado."
        propuestas = bloque.get("propuestas", "").strip() or "Mantener seguimiento y retroalimentación permanente para sostener el avance del curso."

        agregar_parrafo_antes(firma, f"1. Aspectos positivos: {positivos}", bold_prefix="1. Aspectos positivos:")
        agregar_parrafo_antes(firma, f"2. Inconvenientes presentados: {inconvenientes}", bold_prefix="2. Inconvenientes presentados:")
        agregar_parrafo_antes(firma, f"3. Propuestas metodológicas: {propuestas}", bold_prefix="3. Propuestas metodológicas:")
        agregar_parrafo_antes(firma, "")

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()

# -----------------------------------------------------------------------------
# Listados, calificaciones y resumen académico
# -----------------------------------------------------------------------------
def leer_tabla_excel(uploaded_file) -> pd.DataFrame:
    if uploaded_file is None:
        return pd.DataFrame()
    nombre = uploaded_file.name.lower()
    if nombre.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    if nombre.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded_file, sheet_name=0, header=None)
    raise ValueError("Formato no soportado. Use .xls, .xlsx o .csv.")


def detectar_fila_encabezado(df: pd.DataFrame, palabras_clave: List[str]) -> int:
    claves = [normalizar_texto(p) for p in palabras_clave]
    for idx, row in df.iterrows():
        texto = " | ".join(normalizar_texto(v) for v in row.values if not pd.isna(v))
        coincidencias = sum(1 for clave in claves if clave in texto)
        if coincidencias >= max(1, min(2, len(claves))):
            return int(idx)
    return 0


def normalizar_columnas_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def leer_listado_estudiantes(uploaded_file) -> pd.DataFrame:
    if uploaded_file is None:
        return pd.DataFrame(columns=["Nombre completo", "Documento", "Correo", "Plan", "Observación", "Estado"])
    nombre = uploaded_file.name.lower()
    if nombre.endswith(".csv"):
        raw = pd.read_csv(uploaded_file, header=None)
    else:
        raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)
    header_row = detectar_fila_encabezado(raw, ["NOMBRE COMPLETO", "DOCUMENTO", "CORREO"])
    df = raw.iloc[header_row + 1:].copy()
    df.columns = [str(c).strip() for c in raw.iloc[header_row].tolist()]
    df = normalizar_columnas_df(df)
    rename = {}
    for col in df.columns:
        n = normalizar_texto(col)
        if "NOMBRE" in n:
            rename[col] = "Nombre completo"
        elif "DOCUMENTO" in n or "CEDULA" in n or "CARN" in n:
            rename[col] = "Documento"
        elif "CORREO" in n or "EMAIL" in n:
            rename[col] = "Correo"
        elif n == "PLAN" or "PROGRAMA" in n:
            rename[col] = "Plan"
        elif "OBSERV" in n:
            rename[col] = "Observación"
    df = df.rename(columns=rename)
    for col in ["Nombre completo", "Documento", "Correo", "Plan", "Observación"]:
        if col not in df.columns:
            df[col] = ""
    df = df[["Nombre completo", "Documento", "Correo", "Plan", "Observación"]]
    df = df.dropna(how="all").copy()
    df = df[df["Nombre completo"].astype(str).str.strip().ne("")].reset_index(drop=True)
    df["Documento"] = df["Documento"].apply(lambda x: "" if pd.isna(x) else str(x).replace(".0", "").strip())
    df["Estado"] = df["Observación"].apply(lambda x: "Desertó" if re.search(r"DESERT|RETI|CANCEL", normalizar_texto(x)) else "Activo")
    return df


def leer_calificaciones(uploaded_file) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if uploaded_file is None:
        return pd.DataFrame(), pd.DataFrame()
    nombre = uploaded_file.name.lower()
    if not nombre.endswith((".xlsx", ".xls", ".csv")):
        raise ValueError("Formato de calificaciones no soportado.")
    if nombre.endswith(".csv"):
        return pd.read_csv(uploaded_file), pd.DataFrame()
    xls = pd.ExcelFile(uploaded_file)
    hoja_cal = None
    hoja_eval = None
    for sheet in xls.sheet_names:
        n = normalizar_texto(sheet)
        if "CALIFIC" in n:
            hoja_cal = sheet
        if "EVALU" in n:
            hoja_eval = sheet
    if hoja_cal is None:
        hoja_cal = xls.sheet_names[0]
    cal = pd.read_excel(xls, sheet_name=hoja_cal)
    ev = pd.read_excel(xls, sheet_name=hoja_eval) if hoja_eval else pd.DataFrame()
    return normalizar_columnas_df(cal), normalizar_columnas_df(ev)


def encontrar_columna(df: pd.DataFrame, patrones: List[str]) -> Optional[str]:
    patrones_norm = [normalizar_texto(p) for p in patrones]
    for col in df.columns:
        n = normalizar_texto(col)
        if any(p in n for p in patrones_norm):
            return col
    return None


def contar_por_estado_y_nota(df: pd.DataFrame, columna_nota: Optional[str], corte_aprobacion: float, estado_col: Optional[str]) -> Tuple[int, int]:
    if df.empty or not columna_nota or columna_nota not in df.columns:
        return 0, 0
    notas = pd.to_numeric(df[columna_nota], errors="coerce")
    activo = pd.Series([True] * len(df), index=df.index)
    if estado_col and estado_col in df.columns:
        activo = ~df[estado_col].astype(str).apply(lambda x: bool(re.search(r"DESERT|RETI|CANCEL", normalizar_texto(x))))
    validas = notas.notna() & activo
    aprueban = int((notas[validas] >= corte_aprobacion).sum())
    reprueban = int((notas[validas] < corte_aprobacion).sum())
    return aprueban, reprueban


def resumen_gc72_desde_archivos(
    listado_df: pd.DataFrame,
    calificaciones_df: pd.DataFrame,
    evaluaciones_df: pd.DataFrame,
    codigo: str,
    grupo: str,
    asignatura: str,
    avance_contenido: float,
    porcentaje_evaluado_manual: float,
    corte_aprobacion: float = 3.0,
) -> pd.DataFrame:
    matriculados = int(len(listado_df)) if not listado_df.empty else int(len(calificaciones_df))
    estado_col_listado = "Estado" if "Estado" in listado_df.columns else None
    desertaron = 0
    if not listado_df.empty and estado_col_listado:
        desertaron = int(listado_df[estado_col_listado].astype(str).apply(lambda x: bool(re.search(r"DESERT|RETI|CANCEL", normalizar_texto(x)))).sum())
    estado_col = encontrar_columna(calificaciones_df, ["Estado", "Observación", "Observacion"])
    parcial_col = encontrar_columna(calificaciones_df, ["Nota parcial", "Parcial", "Corte 1", "Primer corte"])
    acumulada_col = encontrar_columna(calificaciones_df, ["Nota acumulada", "Acumulada", "Definitiva", "Nota final", "Final"])
    if acumulada_col is None:
        acumulada_col = parcial_col
    apr_par, rep_par = contar_por_estado_y_nota(calificaciones_df, parcial_col, corte_aprobacion, estado_col)
    apr_fecha, rep_fecha = contar_por_estado_y_nota(calificaciones_df, acumulada_col, corte_aprobacion, estado_col)

    # Si la plantilla trae hoja de evaluaciones, se usa para sugerir el % evaluado.
    porcentaje_evaluado = porcentaje_evaluado_manual
    if not evaluaciones_df.empty:
        col_valor = encontrar_columna(evaluaciones_df, ["Valor", "%"])
        if col_valor:
            valores = pd.to_numeric(evaluaciones_df[col_valor], errors="coerce").fillna(0)
            # Si aún no hay fecha o nota, el docente puede ajustar manualmente en la interfaz.
            porcentaje_evaluado = float(min(100, valores.sum())) if valores.sum() > 0 else porcentaje_evaluado_manual

    data = [{
        "Código": codigo,
        "Grupo": grupo,
        "Asignatura": asignatura,
        "% Avance en contenido": avance_contenido,
        "% Evaluado": porcentaje_evaluado,
        "Estudiantes matriculados": matriculados,
        "Desertaron N°": desertaron,
        "Desertaron %": porcentaje(desertaron, matriculados),
        "Aprueban evaluación parcial N°": apr_par,
        "Aprueban evaluación parcial %": porcentaje(apr_par, matriculados),
        "Reprueban evaluación parcial N°": rep_par,
        "Reprueban evaluación parcial %": porcentaje(rep_par, matriculados),
        "Aprueban a la fecha N°": apr_fecha,
        "Aprueban a la fecha %": porcentaje(apr_fecha, matriculados),
        "Reprueban a la fecha N°": rep_fecha,
        "Reprueban a la fecha %": porcentaje(rep_fecha, matriculados),
    }]
    return pd.DataFrame(data, columns=COLUMNAS_GC72)

# -----------------------------------------------------------------------------
# Planificación de clases y FD-GC71
# -----------------------------------------------------------------------------
def parse_time_value(value) -> Optional[time]:
    if isinstance(value, time):
        return value
    if isinstance(value, datetime):
        return value.time()
    if value is None or pd.isna(value):
        return None
    texto = str(value).strip()
    if not texto:
        return None
    for fmt in ["%H:%M", "%H:%M:%S", "%I:%M %p", "%I:%M%p"]:
        try:
            return datetime.strptime(texto.upper(), fmt).time()
        except ValueError:
            pass
    return None


def horas_entre(inicio: time, fin: time) -> float:
    di = datetime.combine(date.today(), inicio)
    df = datetime.combine(date.today(), fin)
    if df <= di:
        df += timedelta(days=1)
    return round((df - di).total_seconds() / 3600, 2)


def generar_fechas_clase(inicio: date, fin: date, horarios_df: pd.DataFrame, fechas_excluidas: Iterable[date] | None = None) -> pd.DataFrame:
    horarios = limpiar_df(horarios_df, COLUMNAS_HORARIOS)
    fechas_excluidas = set(fechas_excluidas or [])
    registros = []
    actual = inicio
    while actual <= fin:
        dia_nombre = DIAS_INV.get(actual.weekday(), "")
        for _, h in horarios.iterrows():
            dia = str(h.get("Día", "")).strip()
            if DIAS_MAP.get(dia, -1) == actual.weekday() and actual not in fechas_excluidas:
                hi = parse_time_value(h.get("Hora inicio")) or time(0, 0)
                hf = parse_time_value(h.get("Hora fin")) or time(0, 0)
                registros.append({
                    "Fecha": actual,
                    "Día": dia_nombre,
                    "Hora inicio": hi.strftime("%H:%M") if hi else "",
                    "Hora fin": hf.strftime("%H:%M") if hf else "",
                    "Horas": horas_entre(hi, hf) if hi and hf else 0,
                    "Lugar / ambiente": h.get("Lugar / ambiente", ""),
                })
        actual += timedelta(days=1)
    df = pd.DataFrame(registros)
    if df.empty:
        return pd.DataFrame(columns=["Fecha", "Día", "Hora inicio", "Hora fin", "Horas", "Lugar / ambiente"])
    df["orden"] = df["Fecha"].astype(str) + " " + df["Hora inicio"].astype(str)
    df = df.sort_values("orden").drop(columns=["orden"]).reset_index(drop=True)
    return df


def expandir_plan_sesiones(modulos_df: pd.DataFrame, fechas_clase_df: pd.DataFrame, criterio: str = "Horas presenciales") -> pd.DataFrame:
    modulos = limpiar_df(modulos_df, COLUMNAS_MODULOS)
    registros = []
    idx_fecha = 0
    n_sesion = 1
    for _, mod in modulos.iterrows():
        unidad = str(mod.get("Unidad", "")).strip() or f"Unidad {len(registros)+1}"
        contenido = str(mod.get("Contenido / tema central", "")).strip()
        trabajo_p = str(mod.get("Trabajo presencial", "")).strip()
        trabajo_i = str(mod.get("Trabajo independiente", "")).strip()
        horas_objetivo = limpiar_numero(mod.get("Horas presenciales")) or 0
        sesiones_objetivo = limpiar_numero(mod.get("Sesiones"))
        if criterio == "Sesiones" and sesiones_objetivo:
            n_bloques = int(max(1, round(sesiones_objetivo)))
        else:
            if fechas_clase_df.empty:
                n_bloques = int(max(1, math.ceil(horas_objetivo / 2))) if horas_objetivo else 1
            else:
                acumuladas = 0.0
                n_bloques = 0
                tmp_idx = idx_fecha
                while acumuladas < max(0.01, horas_objetivo) and tmp_idx < len(fechas_clase_df):
                    acumuladas += limpiar_numero(fechas_clase_df.iloc[tmp_idx].get("Horas")) or 0
                    n_bloques += 1
                    tmp_idx += 1
                n_bloques = max(1, n_bloques)
        for parte in range(1, n_bloques + 1):
            if idx_fecha < len(fechas_clase_df):
                f = fechas_clase_df.iloc[idx_fecha]
                fecha = f.get("Fecha")
                horario = f"{f.get('Hora inicio', '')} - {f.get('Hora fin', '')}".strip(" -")
                lugar = f.get("Lugar / ambiente", "")
            else:
                fecha = "Por programar"
                horario = "Por programar"
                lugar = ""
            sufijo = f" (parte {parte} de {n_bloques})" if n_bloques > 1 else ""
            registros.append({
                "Unidad": unidad,
                "N° sesión": n_sesion,
                "Fecha": fecha,
                "Horario": horario,
                "Contenido por desarrollar": f"{contenido}{sufijo}",
                "Descripción del trabajo presencial": trabajo_p,
                "Descripción trabajo independiente": trabajo_i,
                "Lugar / ambiente": lugar,
            })
            idx_fecha += 1
            n_sesion += 1
    return pd.DataFrame(registros, columns=COLUMNAS_SESIONES)


def configurar_documento_gc71(doc: Document):
    section = doc.sections[0]
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(0.35)
    section.bottom_margin = Inches(0.35)
    section.left_margin = Inches(0.35)
    section.right_margin = Inches(0.35)
    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(8.5)


def agregar_tabla_header_gc71(doc: Document):
    table = doc.add_table(rows=2, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    aplicar_bordes_tabla(table)
    widths = [2.55, 3.9, 1.75]
    for row in table.rows:
        for i, w in enumerate(widths):
            set_cell_width(row.cells[i], w)
    logo_cell = table.cell(0, 0).merge(table.cell(1, 0))
    title_cell = table.cell(0, 1).merge(table.cell(1, 1))
    if LOGO_POLI.exists():
        logo_cell.text = ""
        p = logo_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(LOGO_POLI), width=Inches(1.55))
    else:
        set_cell_text(logo_cell, "POLITÉCNICO COLOMBIANO\nJAIME ISAZA CADAVID", 8, True)
    set_cell_text(title_cell, "GUÍA DIDÁCTICA DE ASIGNATURA Y\nCONCERTACIÓN DE EVALUACIÓN", 12, True)
    set_cell_text(table.cell(0, 2), "Código: FD-GC71", 10, False)
    set_cell_text(table.cell(1, 2), "Versión: 09", 10, True)
    return table


def agregar_fila_seccion(table, titulo: str, cols: int, fill="FFF2CC"):
    row = table.add_row()
    cell = row.cells[0].merge(row.cells[cols - 1])
    shade_cell(cell, fill)
    set_cell_text(cell, titulo, 8.5, True)
    return row


def agregar_tabla_identificacion_gc71(doc: Document, datos: Dict[str, str]):
    doc.add_paragraph()
    t = doc.add_table(rows=0, cols=4)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    aplicar_bordes_tabla(t)
    widths = [2.2, 2.05, 1.85, 2.0]
    for _ in range(1):
        pass
    agregar_fila_seccion(t, "1. IDENTIFICACIÓN DE LA ASIGNATURA", 4)
    campos = [
        ("PROGRAMA ACADÉMICO", datos.get("programa", "")),
        ("ASIGNATURA", datos.get("asignatura", "")),
        ("CÓDIGO", datos.get("codigo", "")),
        ("ÁREA DE FORMACIÓN", datos.get("area", "")),
        ("PRERREQUISITO(S)", datos.get("prerrequisitos", "")),
        ("CORREQUISITO(S)", datos.get("correquisitos", "")),
    ]
    for etiqueta, valor in campos:
        row = t.add_row()
        c0 = row.cells[0].merge(row.cells[1])
        c1 = row.cells[2].merge(row.cells[3])
        set_cell_text(c0, etiqueta, 8.3, True, align=WD_ALIGN_PARAGRAPH.LEFT)
        set_cell_text(c1, valor, 8.3, False, align=WD_ALIGN_PARAGRAPH.LEFT)
    row = t.add_row()
    set_cell_text(row.cells[0], "TIPO DE ASIGNATURA", 8.3, True, align=WD_ALIGN_PARAGRAPH.LEFT)
    tipo = datos.get("tipo_asignatura", "Teórico-práctica")
    set_cell_text(row.cells[1], f"{'☒' if tipo == 'Teórica' else '☐'} Teórica", 8.3)
    set_cell_text(row.cells[2], f"{'☒' if tipo == 'Teórico-práctica' else '☐'} Teórico-práctica", 8.3)
    set_cell_text(row.cells[3], f"{'☒' if tipo == 'Práctica' else '☐'} Práctica", 8.3)
    row = t.add_row()
    set_cell_text(row.cells[0].merge(row.cells[1]), "NÚMERO DE CRÉDITOS", 8.3, True, align=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(row.cells[2].merge(row.cells[3]), datos.get("creditos", ""), 8.3, align=WD_ALIGN_PARAGRAPH.LEFT)
    row = t.add_row()
    set_cell_text(row.cells[0], "DISTRIBUCIÓN HORARIA SEMANAL", 8.3, True, align=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(row.cells[1], f"HTP: {datos.get('htp', '')}", 8.3)
    set_cell_text(row.cells[2], f"HTI: {datos.get('hti', '')}", 8.3)
    set_cell_text(row.cells[3], f"Total: {datos.get('ht_total', '')}", 8.3)
    for etiqueta, valor in [
        ("PROFESOR", datos.get("profesor", "")),
        ("CORREO ELECTRÓNICO", datos.get("correo", "")),
        ("GRUPO", datos.get("grupo", "")),
        ("PERÍODO ACADÉMICO", datos.get("periodo", "")),
    ]:
        row = t.add_row()
        c0 = row.cells[0].merge(row.cells[1])
        c1 = row.cells[2].merge(row.cells[3])
        set_cell_text(c0, etiqueta, 8.3, True, align=WD_ALIGN_PARAGRAPH.LEFT)
        set_cell_text(c1, valor, 8.3, align=WD_ALIGN_PARAGRAPH.LEFT)
    for row in t.rows:
        for i, w in enumerate(widths):
            if i < len(row.cells):
                set_cell_width(row.cells[i], w)
    return t


def agregar_seccion_texto_gc71(doc: Document, numero_titulo: str, texto: str):
    t = doc.add_table(rows=2, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    aplicar_bordes_tabla(t)
    shade_cell(t.rows[0].cells[0], "FFF2CC")
    set_cell_text(t.rows[0].cells[0], numero_titulo, 8.5, True)
    set_cell_text(t.rows[1].cells[0], texto, 8.2, align=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_width(t.rows[0].cells[0], 8.1)
    set_cell_width(t.rows[1].cells[0], 8.1)
    return t


def agregar_contenidos_gc71(doc: Document, sesiones_df: pd.DataFrame):
    t = doc.add_table(rows=0, cols=5)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    aplicar_bordes_tabla(t)
    widths = [0.75, 1.0, 2.45, 1.95, 1.95]
    agregar_fila_seccion(t, "7. CONTENIDOS TEMÁTICOS DE LA ASIGNATURA", 5)
    sesiones = limpiar_df(sesiones_df, COLUMNAS_SESIONES)
    for unidad, dfu in sesiones.groupby("Unidad", sort=False):
        agregar_fila_seccion(t, str(unidad).upper(), 5, fill="FFF2CC")
        row = t.add_row()
        headers = ["N° sesión", "Fecha", "Contenido por desarrollar", "Descripción del trabajo presencial", "Descripción trabajo independiente"]
        for i, h in enumerate(headers):
            shade_cell(row.cells[i], "F2F2F2")
            set_cell_text(row.cells[i], h, 7.5, True)
        for _, s in dfu.iterrows():
            row = t.add_row()
            fecha = s.get("Fecha")
            if hasattr(fecha, "strftime"):
                fecha_txt = fecha.strftime("%d/%m/%Y")
            else:
                fecha_txt = str(fecha)
            values = [
                s.get("N° sesión", ""),
                fecha_txt,
                s.get("Contenido por desarrollar", ""),
                s.get("Descripción del trabajo presencial", ""),
                s.get("Descripción trabajo independiente", ""),
            ]
            for i, v in enumerate(values):
                set_cell_text(row.cells[i], v, 7.2 if i >= 2 else 7.0, align=WD_ALIGN_PARAGRAPH.LEFT if i >= 2 else WD_ALIGN_PARAGRAPH.CENTER)
    for row in t.rows:
        for i, w in enumerate(widths):
            if i < len(row.cells):
                set_cell_width(row.cells[i], w)
    return t


def agregar_evaluaciones_gc71(doc: Document, asignatura: str, grupo: str, evaluaciones_df: pd.DataFrame):
    t = doc.add_table(rows=0, cols=4)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    aplicar_bordes_tabla(t)
    widths = [1.45, 4.15, 0.9, 1.6]
    agregar_fila_seccion(t, "11. EVALUACIÓN DE LA ASIGNATURA", 4)
    row = t.add_row()
    c0 = row.cells[0].merge(row.cells[1])
    c1 = row.cells[2].merge(row.cells[3])
    set_cell_text(c0, f"ASIGNATURA: {asignatura}", 7.6, True, align=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(c1, f"GRUPO: {grupo}", 7.6, True, align=WD_ALIGN_PARAGRAPH.LEFT)
    row = t.add_row()
    headers = ["TIPO DE EVALUACIÓN°", "PROCEDIMIENTO DE EVALUACIÓN\n(Descripción de la actividad evaluativa)", "VALOR (%)", "FECHA DE\nREALIZACIÓN"]
    for i, h in enumerate(headers):
        shade_cell(row.cells[i], "FFF2CC")
        set_cell_text(row.cells[i], h, 7.2, True)
    evaluaciones = limpiar_df(evaluaciones_df, COLUMNAS_EVALUACIONES)
    for _, ev in evaluaciones.iterrows():
        row = t.add_row()
        fecha = ev.get("Fecha de realización")
        if hasattr(fecha, "strftime"):
            fecha_txt = fecha.strftime("%d/%m/%Y")
        else:
            fecha_txt = str(fecha)
        vals = [ev.get("Tipo de evaluación", ""), ev.get("Procedimiento de evaluación", ""), ev.get("Valor (%)", ""), fecha_txt]
        for i, v in enumerate(vals):
            set_cell_text(row.cells[i], v, 7.2, align=WD_ALIGN_PARAGRAPH.LEFT if i in [0, 1] else WD_ALIGN_PARAGRAPH.CENTER)
    for row in t.rows:
        for i, w in enumerate(widths):
            if i < len(row.cells):
                set_cell_width(row.cells[i], w)
    return t


def agregar_evidencia_y_control_gc71(doc: Document, datos: Dict[str, str], representantes_df: pd.DataFrame):
    t = doc.add_table(rows=0, cols=3)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    aplicar_bordes_tabla(t)
    agregar_fila_seccion(t, "12. EVIDENCIA DE PRESENTACIÓN DE LA GUÍA Y CONCERTACIÓN DE EVALUACIÓN AL GRUPO DE ESTUDIANTES", 3)
    row = t.add_row()
    cell = row.cells[0].merge(row.cells[2])
    set_cell_text(cell, "Se deja constancia de socialización de la Guía Didáctica de Asignatura y aprobación de la concertación de evaluación según el reglamento estudiantil; para ello firman tres estudiantes en representación del grupo:", 7.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    row = t.add_row()
    for i, h in enumerate(["Nombre de los estudiantes", "N° de cédula o carné estudiantil", "Firma"]):
        shade_cell(row.cells[i], "FFF2CC")
        set_cell_text(row.cells[i], h, 7.5, True)
    reps = representantes_df.copy() if representantes_df is not None else pd.DataFrame()
    for i in range(3):
        row = t.add_row()
        nombre = reps.iloc[i].get("Nombre", "") if i < len(reps) else ""
        docu = reps.iloc[i].get("Documento", "") if i < len(reps) else ""
        set_cell_text(row.cells[0], nombre, 7.5, align=WD_ALIGN_PARAGRAPH.LEFT)
        set_cell_text(row.cells[1], docu, 7.5)
        set_cell_text(row.cells[2], "", 7.5)
    row = t.add_row()
    for i, h in enumerate(["Nombre del docente del curso", "Cédula", "Firma"]):
        shade_cell(row.cells[i], "FFF2CC")
        set_cell_text(row.cells[i], h, 7.5, True)
    row = t.add_row()
    set_cell_text(row.cells[0], datos.get("profesor", ""), 7.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(row.cells[1], datos.get("cedula_docente", ""), 7.5)
    set_cell_text(row.cells[2], "", 7.5)
    row = t.add_row()
    cell = row.cells[0].merge(row.cells[2])
    set_cell_text(cell, f"Fecha de socialización de la Guía Didáctica: {datos.get('fecha_socializacion', '')}", 7.5, True, align=WD_ALIGN_PARAGRAPH.LEFT)
    row = t.add_row()
    cell = row.cells[0].merge(row.cells[2])
    set_cell_text(cell, "Nota: El docente se compromete a devolver las evaluaciones, socializar la calificación con los estudiantes y a ingresar dicha calificación al sistema académico, correcta y oportunamente.", 7.2, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    for row in t.rows:
        set_cell_width(row.cells[0], 3.4)
        set_cell_width(row.cells[1], 2.0)
        set_cell_width(row.cells[2], 2.7)

    doc.add_paragraph()
    c = doc.add_table(rows=3, cols=2)
    c.alignment = WD_TABLE_ALIGNMENT.CENTER
    aplicar_bordes_tabla(c)
    cell = c.rows[0].cells[0].merge(c.rows[0].cells[1])
    shade_cell(cell, "FFF2CC")
    set_cell_text(cell, "CONTROL DE CAMBIOS Y VIGENCIA (DILIGENCIAR LOS DATOS ESPECÍFICOS)", 8, True)
    set_cell_text(c.rows[1].cells[0], "Fecha de Revisión por parte del Coordinador de Área:", 7.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(c.rows[1].cells[1], datos.get("fecha_revision", ""), 7.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(c.rows[2].cells[0], "Fecha de aprobación y acta de sesión del Comité de currículo del programa:", 7.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(c.rows[2].cells[1], datos.get("fecha_aprobacion", ""), 7.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    for row in c.rows:
        set_cell_width(row.cells[0], 4.0)
        set_cell_width(row.cells[1], 4.1)


def crear_gc71_docx(
    datos: Dict[str, str],
    sesiones_df: pd.DataFrame,
    evaluaciones_df: pd.DataFrame,
    representantes_df: Optional[pd.DataFrame] = None,
) -> bytes:
    doc = Document()
    configurar_documento_gc71(doc)
    agregar_tabla_header_gc71(doc)
    agregar_tabla_identificacion_gc71(doc, datos)
    agregar_seccion_texto_gc71(doc, "2. JUSTIFICACIÓN", datos.get("justificacion", ""))
    agregar_seccion_texto_gc71(doc, "3. COMPETENCIAS A LAS QUE LE TRIBUTA LA ASIGNATURA", datos.get("competencias", ""))
    agregar_seccion_texto_gc71(doc, "4. RESULTADOS DE APRENDIZAJE A LOS QUE LE TRIBUTA LA ASIGNATURA", datos.get("resultados", ""))
    agregar_seccion_texto_gc71(doc, "5. OBJETIVOS DE APRENDIZAJE DE LA ASIGNATURA", f"OBJETIVO(S) GENERAL(ES)\n{datos.get('objetivo_general', '')}\n\nOBJETIVOS ESPECÍFICOS\n{datos.get('objetivos_especificos', '')}")
    agregar_seccion_texto_gc71(doc, "6. METODOLOGÍAS Y ESTRATEGIAS DIDÁCTICAS DE LA ASIGNATURA", datos.get("metodologias", ""))
    agregar_contenidos_gc71(doc, sesiones_df)
    agregar_seccion_texto_gc71(doc, "8. AMBIENTES DE APRENDIZAJE DE LA ASIGNATURA", datos.get("ambientes", ""))
    agregar_seccion_texto_gc71(doc, "9. MEDIOS EDUCATIVOS PARA LA ASIGNATURA", datos.get("medios", ""))
    agregar_seccion_texto_gc71(doc, "10. REFERENCIAS BIBLIOGRÁFICAS", datos.get("referencias", ""))
    agregar_evaluaciones_gc71(doc, datos.get("asignatura", ""), datos.get("grupo", ""), evaluaciones_df)
    agregar_evidencia_y_control_gc71(doc, datos, representantes_df if representantes_df is not None else pd.DataFrame())
    # Certificación en pie visual al final.
    if LOGO_ICONTEC.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run().add_picture(str(LOGO_ICONTEC), width=Inches(1.0))
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()

# -----------------------------------------------------------------------------
# Excel de evaluación
# -----------------------------------------------------------------------------
def crear_plantilla_evaluacion_xlsx(
    estudiantes_df: pd.DataFrame,
    evaluaciones_df: pd.DataFrame,
    datos: Dict[str, str],
    corte_aprobacion: float = 3.0,
) -> bytes:
    import xlsxwriter

    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    fmt_title = workbook.add_format({"bold": True, "font_size": 14, "align": "center", "valign": "vcenter", "bg_color": "#D9EAF7", "border": 1})
    fmt_header = workbook.add_format({"bold": True, "bg_color": "#FFF2CC", "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
    fmt_cell = workbook.add_format({"border": 1, "valign": "vcenter"})
    fmt_num = workbook.add_format({"border": 1, "num_format": "0.00", "valign": "vcenter"})
    fmt_pct = workbook.add_format({"border": 1, "num_format": "0%", "valign": "vcenter"})
    fmt_note = workbook.add_format({"border": 1, "font_color": "#666666", "italic": True, "text_wrap": True})

    # Configuración
    ws = workbook.add_worksheet("Configuración")
    ws.merge_range("A1:D1", "Plantilla de evaluación generada desde FD-GC71", fmt_title)
    config_rows = [
        ("Programa", datos.get("programa", "")),
        ("Asignatura", datos.get("asignatura", "")),
        ("Código", datos.get("codigo", "")),
        ("Grupo", datos.get("grupo", "")),
        ("Periodo académico", datos.get("periodo", "")),
        ("Docente", datos.get("profesor", "")),
        ("Nota mínima aprobatoria", corte_aprobacion),
    ]
    ws.write_row("A3", ["Campo", "Valor"], fmt_header)
    for r, (k, v) in enumerate(config_rows, start=3):
        ws.write(r, 0, k, fmt_cell)
        ws.write(r, 1, v, fmt_num if isinstance(v, (int, float)) else fmt_cell)
    ws.set_column("A:A", 26)
    ws.set_column("B:D", 32)

    # Estudiantes
    est = estudiantes_df.copy()
    if est.empty:
        est = pd.DataFrame(columns=["Nombre completo", "Documento", "Correo", "Plan", "Observación", "Estado"])
    for col in ["Nombre completo", "Documento", "Correo", "Plan", "Observación", "Estado"]:
        if col not in est.columns:
            est[col] = ""
    est = est[["Nombre completo", "Documento", "Correo", "Plan", "Observación", "Estado"]]
    ws = workbook.add_worksheet("Estudiantes")
    ws.write_row(0, 0, est.columns.tolist(), fmt_header)
    for r, row in est.iterrows():
        for c, col in enumerate(est.columns):
            ws.write(r + 1, c, "" if pd.isna(row[col]) else row[col], fmt_cell)
    ws.autofilter(0, 0, max(1, len(est)), len(est.columns) - 1)
    ws.freeze_panes(1, 0)
    ws.set_column("A:A", 34)
    ws.set_column("B:B", 16)
    ws.set_column("C:C", 34)
    ws.set_column("D:D", 42)
    ws.set_column("E:F", 18)

    # Evaluaciones
    ev = limpiar_df(evaluaciones_df, COLUMNAS_EVALUACIONES)
    ws = workbook.add_worksheet("Evaluaciones")
    ws.write_row(0, 0, COLUMNAS_EVALUACIONES, fmt_header)
    total_valor = 0.0
    for r, row in ev.iterrows():
        for c, col in enumerate(COLUMNAS_EVALUACIONES):
            value = row.get(col, "")
            if col == "Fecha de realización" and hasattr(value, "strftime"):
                value = value.strftime("%d/%m/%Y")
            if col == "Valor (%)":
                val = limpiar_numero(value) or 0
                total_valor += val
                ws.write_number(r + 1, c, val, fmt_num)
            else:
                ws.write(r + 1, c, "" if pd.isna(value) else value, fmt_cell)
    ws.write(len(ev) + 2, 1, "Total concertado", fmt_header)
    ws.write_formula(len(ev) + 2, 2, f"=SUM(C2:C{len(ev)+1})", fmt_num)
    ws.write(len(ev) + 4, 0, "Nota", fmt_note)
    ws.merge_range(len(ev) + 4, 1, len(ev) + 4, 5, "El total de Valor (%) debe sumar 100. Las columnas de notas se generan automáticamente en la hoja Calificaciones.", fmt_note)
    ws.freeze_panes(1, 0)
    ws.set_column("A:A", 20)
    ws.set_column("B:B", 52)
    ws.set_column("C:C", 12)
    ws.set_column("D:F", 18)

    # Calificaciones
    ws = workbook.add_worksheet("Calificaciones")
    fixed_headers = ["Nombre completo", "Documento", "Correo", "Estado"]
    eval_headers = [f"E{i+1} - {str(row.get('Tipo de evaluación', 'Evaluación')).strip()[:25]}" for i, (_, row) in enumerate(ev.iterrows())]
    headers = fixed_headers + eval_headers + ["Nota parcial", "Nota acumulada", "Aprueba a la fecha", "Observación docente"]
    ws.write_row(0, 0, headers, fmt_header)
    n = max(len(est), 1)
    for r in range(n):
        if r < len(est):
            ws.write(r + 1, 0, est.iloc[r].get("Nombre completo", ""), fmt_cell)
            ws.write(r + 1, 1, est.iloc[r].get("Documento", ""), fmt_cell)
            ws.write(r + 1, 2, est.iloc[r].get("Correo", ""), fmt_cell)
            ws.write(r + 1, 3, est.iloc[r].get("Estado", "Activo"), fmt_cell)
        else:
            for c in range(4):
                ws.write(r + 1, c, "", fmt_cell)
        # notas por evaluación
        for c in range(len(eval_headers)):
            ws.write_blank(r + 1, 4 + c, None, fmt_num)
        first_eval_col = 4
        last_eval_col = 4 + len(eval_headers) - 1
        # Weighted formulas. C in Evaluaciones is Valor (%)
        excel_row = r + 2
        if eval_headers:
            weighted_terms = []
            for i in range(len(eval_headers)):
                col_letter = xlsxwriter.utility.xl_col_to_name(first_eval_col + i)
                peso_cell = f"Evaluaciones!$C${i+2}"
                weighted_terms.append(f"IF(ISNUMBER({col_letter}{excel_row}),{col_letter}{excel_row}*{peso_cell}/100,0)")
            formula = "=" + "+".join(weighted_terms)
            ws.write_formula(r + 1, last_eval_col + 1, formula, fmt_num)  # Nota parcial = acumulada sugerida
            ws.write_formula(r + 1, last_eval_col + 2, formula, fmt_num)
            ws.write_formula(r + 1, last_eval_col + 3, f'=IF({xlsxwriter.utility.xl_col_to_name(last_eval_col + 2)}{excel_row}>=Configuración!$B$10,"Sí","No")', fmt_cell)
        else:
            ws.write_blank(r + 1, 4, None, fmt_num)
    ws.freeze_panes(1, 4)
    ws.autofilter(0, 0, n, len(headers) - 1)
    ws.set_column("A:A", 34)
    ws.set_column("B:B", 16)
    ws.set_column("C:C", 32)
    ws.set_column("D:D", 14)
    if eval_headers:
        ws.set_column(4, 4 + len(eval_headers) - 1, 14)
        ws.set_column(4 + len(eval_headers), len(headers) - 1, 18)
    else:
        ws.set_column("E:H", 16)

    # Resumen para FD-GC72
    ws = workbook.add_worksheet("Resumen FD-GC72")
    resumen_headers = ["Indicador", "Valor", "Observación"]
    ws.write_row(0, 0, resumen_headers, fmt_header)
    cal_sheet = "Calificaciones"
    nota_parcial_col = xlsxwriter.utility.xl_col_to_name(4 + len(eval_headers)) if eval_headers else "E"
    nota_acum_col = xlsxwriter.utility.xl_col_to_name(5 + len(eval_headers)) if eval_headers else "F"
    estado_col = "D"
    filas = [
        ("Código", datos.get("codigo", ""), ""),
        ("Grupo", datos.get("grupo", ""), ""),
        ("Asignatura", datos.get("asignatura", ""), ""),
        ("Estudiantes matriculados", f"=COUNTA({cal_sheet}!A2:A{n+1})", "Desde listado de clase"),
        ("Desertaron N°", f'=COUNTIF({cal_sheet}!{estado_col}2:{estado_col}{n+1},"*Desert*")+COUNTIF({cal_sheet}!{estado_col}2:{estado_col}{n+1},"*Retir*")+COUNTIF({cal_sheet}!{estado_col}2:{estado_col}{n+1},"*Cancel*")', "Según columna Estado"),
        ("Aprueban evaluación parcial N°", f'=COUNTIF({cal_sheet}!{nota_parcial_col}2:{nota_parcial_col}{n+1},">="&Configuración!$B$10)', "Nota parcial >= mínima"),
        ("Reprueban evaluación parcial N°", f'=COUNTIF({cal_sheet}!{nota_parcial_col}2:{nota_parcial_col}{n+1},"<"&Configuración!$B$10)', "Nota parcial < mínima"),
        ("Aprueban a la fecha N°", f'=COUNTIF({cal_sheet}!{nota_acum_col}2:{nota_acum_col}{n+1},">="&Configuración!$B$10)', "Nota acumulada >= mínima"),
        ("Reprueban a la fecha N°", f'=COUNTIF({cal_sheet}!{nota_acum_col}2:{nota_acum_col}{n+1},"<"&Configuración!$B$10)', "Nota acumulada < mínima"),
        ("% Evaluado sugerido", "=Evaluaciones!C" + str(len(ev) + 3), "Puede ajustarse en la app"),
        ("% Avance en contenido", "", "Diligenciar en la app con base en sesiones realizadas / sesiones planificadas"),
    ]
    for r, (ind, val, obs) in enumerate(filas, start=1):
        ws.write(r, 0, ind, fmt_cell)
        if isinstance(val, str) and val.startswith("="):
            ws.write_formula(r, 1, val, fmt_num)
        else:
            ws.write(r, 1, val, fmt_cell)
        ws.write(r, 2, obs, fmt_cell)
    ws.set_column("A:A", 34)
    ws.set_column("B:B", 18)
    ws.set_column("C:C", 62)

    workbook.close()
    output.seek(0)
    return output.getvalue()

# -----------------------------------------------------------------------------
# Interfaz Streamlit
# -----------------------------------------------------------------------------
def ui_gc71(st):
    st.header("FD-GC71 - Guía didáctica y concertación de evaluación")
    st.caption("Planifica el curso desde módulos, intensidad y horario real. La app arma las fechas y genera el formato FD-GC71.")
    usuario_actual = st.session_state.get("auth_user", {})

    with st.expander("1. Identificación de la asignatura", expanded=True):
        c1, c2, c3 = st.columns([1.2, 1.2, 0.8])
        with c1:
            programa = st.text_input("Programa académico", value="")
            asignatura = st.text_input("Asignatura", value="")
            codigo = st.text_input("Código", value="")
            area = st.text_input("Área de formación", value="")
        with c2:
            profesor = st.text_input("Profesor", value=usuario_actual.get("nombre_completo", ""))
            cedula_docente = st.text_input("Cédula docente", value="")
            correo = st.text_input("Correo electrónico", value=usuario_actual.get("email", ""))
            grupo = st.text_input("Grupo", value="")
        with c3:
            periodo = st.text_input("Periodo académico", value="2026-1")
            tipo = st.selectbox("Tipo de asignatura", ["Teórica", "Teórico-práctica", "Práctica"], index=1)
            creditos = st.number_input("Número de créditos", min_value=0, step=1, value=3)
            htp = st.number_input("HTP semanal", min_value=0.0, step=0.5, value=2.0)
            hti = st.number_input("HTI semanal", min_value=0.0, step=0.5, value=4.0)
        prerrequisitos = st.text_input("Prerrequisito(s)", value="Ninguno")
        correquisitos = st.text_input("Correquisito(s)", value="Ninguno")

    with st.expander("2. Textos académicos base", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            justificacion = st.text_area("Justificación", TEXTOS_PREDEFINIDOS_GC71["justificacion"], height=140)
            competencias = st.text_area("Competencias a las que tributa", TEXTOS_PREDEFINIDOS_GC71["competencias"], height=130)
            resultados = st.text_area("Resultados de aprendizaje", TEXTOS_PREDEFINIDOS_GC71["resultados"], height=130)
            objetivo_general = st.text_area("Objetivo general", TEXTOS_PREDEFINIDOS_GC71["objetivo_general"], height=110)
        with col2:
            objetivos_especificos = st.text_area("Objetivos específicos", TEXTOS_PREDEFINIDOS_GC71["objetivos_especificos"], height=140)
            metodologias = st.text_area("Metodologías y estrategias didácticas", TEXTOS_PREDEFINIDOS_GC71["metodologias"], height=130)
            ambientes = st.text_area("Ambientes de aprendizaje", TEXTOS_PREDEFINIDOS_GC71["ambientes"], height=100)
            medios = st.text_area("Medios educativos", TEXTOS_PREDEFINIDOS_GC71["medios"], height=100)
            referencias = st.text_area("Referencias bibliográficas", TEXTOS_PREDEFINIDOS_GC71["referencias"], height=100)

    st.subheader("3. Módulos / unidades e intensidad")
    st.write("Define la intensidad por unidad. Puedes trabajar por horas presenciales o por número exacto de sesiones.")
    criterio = st.radio("Criterio de expansión del cronograma", ["Horas presenciales", "Sesiones"], horizontal=True)
    default_modulos = pd.DataFrame([
        {"Unidad": "UNIDAD 1. Fundamentos", "Contenido / tema central": "Introducción, conceptos base y alcance de la asignatura", "Horas presenciales": 4, "Sesiones": 2, "Trabajo presencial": "Clase orientadora, explicación conceptual y taller diagnóstico.", "Trabajo independiente": "Lectura de apoyo y preparación de preguntas orientadoras."},
        {"Unidad": "UNIDAD 2. Aplicación", "Contenido / tema central": "Desarrollo de procedimientos, ejercicios y análisis de casos", "Horas presenciales": 8, "Sesiones": 4, "Trabajo presencial": "Taller aplicado, solución de ejercicios y discusión guiada.", "Trabajo independiente": "Desarrollo de actividad práctica y revisión bibliográfica."},
        {"Unidad": "UNIDAD 3. Integración", "Contenido / tema central": "Proyecto, socialización y retroalimentación", "Horas presenciales": 4, "Sesiones": 2, "Trabajo presencial": "Acompañamiento al proyecto y socialización de resultados.", "Trabajo independiente": "Ajuste de entregables y preparación de sustentación."},
    ])
    modulos_df = st.data_editor(
        default_modulos,
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        column_config={
            "Horas presenciales": st.column_config.NumberColumn("Horas presenciales", min_value=0.0, step=0.5),
            "Sesiones": st.column_config.NumberColumn("Sesiones", min_value=1, step=1),
        },
        key="modulos_gc71",
    )

    st.subheader("4. Horario de clase")
    c1, c2, c3 = st.columns(3)
    with c1:
        fecha_inicio = st.date_input("Fecha de inicio", value=date.today(), format="DD/MM/YYYY", key="gc71_inicio")
    with c2:
        fecha_fin = st.date_input("Fecha de finalización", value=date.today() + timedelta(days=112), format="DD/MM/YYYY", key="gc71_fin")
    with c3:
        fechas_no_clase_txt = st.text_area("Fechas sin clase (dd/mm/aaaa, una por línea)", value="", height=100)
    default_horarios = pd.DataFrame([
        {"Día": "Lunes", "Hora inicio": "18:00", "Hora fin": "20:00", "Lugar / ambiente": "Aula de clase"},
    ])
    horarios_df = st.data_editor(
        default_horarios,
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        column_config={"Día": st.column_config.SelectboxColumn("Día", options=["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]),},
        key="horarios_gc71",
    )
    fechas_excluidas = []
    for linea in fechas_no_clase_txt.splitlines():
        linea = linea.strip()
        if not linea:
            continue
        try:
            fechas_excluidas.append(datetime.strptime(linea, "%d/%m/%Y").date())
        except ValueError:
            st.warning(f"No pude interpretar esta fecha sin clase: {linea}. Use dd/mm/aaaa.")
    fechas_clase = generar_fechas_clase(fecha_inicio, fecha_fin, horarios_df, fechas_excluidas)
    sesiones_df = expandir_plan_sesiones(modulos_df, fechas_clase, criterio=criterio)

    with st.expander("Vista previa del cronograma generado", expanded=True):
        st.dataframe(sesiones_df, use_container_width=True)
        if len(sesiones_df) > len(fechas_clase):
            st.warning("Hay más sesiones requeridas que fechas disponibles. Las últimas sesiones quedan como 'Por programar'.")

    st.subheader("5. Evaluación concertada")
    default_evals = pd.DataFrame([
        {"Tipo de evaluación": "Parcial", "Procedimiento de evaluación": "Prueba escrita o práctica sobre los contenidos de las unidades desarrolladas.", "Valor (%)": 30, "Fecha de realización": sesiones_df.iloc[min(2, len(sesiones_df)-1)]["Fecha"] if not sesiones_df.empty else "", "Unidad relacionada": "UNIDAD 1", "Corte": "Parcial"},
        {"Tipo de evaluación": "Taller / Proyecto", "Procedimiento de evaluación": "Actividad aplicada con entrega de evidencias y criterios definidos en rúbrica.", "Valor (%)": 40, "Fecha de realización": sesiones_df.iloc[min(5, len(sesiones_df)-1)]["Fecha"] if not sesiones_df.empty else "", "Unidad relacionada": "UNIDAD 2", "Corte": "Seguimiento"},
        {"Tipo de evaluación": "Final", "Procedimiento de evaluación": "Socialización, sustentación o prueba final integradora.", "Valor (%)": 30, "Fecha de realización": sesiones_df.iloc[-1]["Fecha"] if not sesiones_df.empty else "", "Unidad relacionada": "UNIDAD 3", "Corte": "Final"},
    ])
    evaluaciones_df = st.data_editor(
        default_evals,
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        column_config={"Valor (%)": st.column_config.NumberColumn("Valor (%)", min_value=0, max_value=100, step=1)},
        key="evaluaciones_gc71",
    )
    total_eval = sum(limpiar_numero(v) or 0 for v in evaluaciones_df.get("Valor (%)", []))
    if round(total_eval, 2) != 100:
        st.warning(f"La evaluación suma {total_eval:.1f}%. El formato puede generarse, pero lo sano es que sume 100%. Aquí no maquillamos cadáveres.")
    else:
        st.success("La evaluación concertada suma 100%.")

    st.subheader("6. Listado tradicional de clase y plantilla de evaluación")
    listado_file = st.file_uploader("Cargar listado tradicional de clase (.xls, .xlsx o .csv)", type=["xls", "xlsx", "csv"], key="listado_gc71")
    estudiantes_df = pd.DataFrame(columns=["Nombre completo", "Documento", "Correo", "Plan", "Observación", "Estado"])
    if listado_file is not None:
        try:
            estudiantes_df = leer_listado_estudiantes(listado_file)
            st.success(f"Listado leído: {len(estudiantes_df)} estudiantes.")
            st.dataframe(estudiantes_df.head(20), use_container_width=True)
        except Exception as e:
            st.error(f"No se pudo leer el listado: {e}")
    representantes_df = pd.DataFrame(columns=["Nombre", "Documento"])
    if not estudiantes_df.empty:
        reps = estudiantes_df.head(3).copy()
        representantes_df = pd.DataFrame({"Nombre": reps["Nombre completo"], "Documento": reps["Documento"]})

    st.subheader("7. Fechas administrativas")
    c1, c2, c3 = st.columns(3)
    with c1:
        fecha_socializacion = st.date_input("Fecha de socialización", value=fecha_inicio, format="DD/MM/YYYY")
    with c2:
        fecha_revision = st.text_input("Fecha revisión coordinador", value="")
    with c3:
        fecha_aprobacion = st.text_input("Fecha aprobación comité / acta", value="")

    datos = {
        "programa": programa,
        "asignatura": asignatura,
        "codigo": codigo,
        "area": area,
        "prerrequisitos": prerrequisitos,
        "correquisitos": correquisitos,
        "tipo_asignatura": tipo,
        "creditos": str(creditos),
        "htp": str(htp),
        "hti": str(hti),
        "ht_total": str(htp + hti),
        "profesor": profesor,
        "cedula_docente": cedula_docente,
        "correo": correo,
        "grupo": grupo,
        "periodo": periodo,
        "justificacion": justificacion,
        "competencias": competencias,
        "resultados": resultados,
        "objetivo_general": objetivo_general,
        "objetivos_especificos": objetivos_especificos,
        "metodologias": metodologias,
        "ambientes": ambientes,
        "medios": medios,
        "referencias": referencias,
        "fecha_socializacion": fecha_socializacion.strftime("%d/%m/%Y"),
        "fecha_revision": fecha_revision,
        "fecha_aprobacion": fecha_aprobacion,
    }

    st.subheader("8. Descargar")
    col_doc, col_xlsx = st.columns(2)
    with col_doc:
        if not profesor.strip() or not asignatura.strip():
            st.info("Para descargar FD-GC71 ingrese como mínimo profesor y asignatura.")
        else:
            try:
                docx_bytes = crear_gc71_docx(datos, sesiones_df, evaluaciones_df, representantes_df)
                nombre_base = nombre_archivo_seguro(profesor, fecha_socializacion, "FD_GC71_Guia_Didactica")
                st.download_button(
                    "Descargar FD-GC71 en Word (.docx)",
                    data=docx_bytes,
                    file_name=f"{nombre_base}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"No se pudo generar FD-GC71: {e}")
    with col_xlsx:
        try:
            plantilla_bytes = crear_plantilla_evaluacion_xlsx(estudiantes_df, evaluaciones_df, datos)
            nombre_xlsx = nombre_archivo_seguro(profesor or "docente", fecha_socializacion, "Plantilla_Evaluacion")
            st.download_button(
                "Descargar plantilla Excel de evaluación",
                data=plantilla_bytes,
                file_name=f"{nombre_xlsx}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"No se pudo generar la plantilla de evaluación: {e}")

    st.session_state["ultimo_datos_gc71"] = datos
    st.session_state["ultimas_sesiones_gc71"] = sesiones_df
    st.session_state["ultimas_evaluaciones_gc71"] = evaluaciones_df
    st.session_state["ultimo_listado_estudiantes"] = estudiantes_df


def ui_gc72(st):
    st.header("FD-GC72 - Informe académico")
    st.caption("Puede diligenciarlo manualmente o alimentarlo con el listado de clase y las plantillas de evaluación cargadas a mitad/final del curso.")

    with st.sidebar:
        st.markdown("### Opciones FD-GC72")
        calcular = st.checkbox("Calcular porcentajes automáticamente", value=True)
        modo_analisis = st.radio("Tipo de análisis descriptivo", ["Por cada curso", "Consolidado institucional"], index=0)
        exportar_doc = st.checkbox("Intentar generar .doc además de .docx", value=False)

    st.subheader("1. Datos generales")
    c1, c2, c3 = st.columns([1.4, 1, 1])
    with c1:
        docente = st.text_input("Docente", value=st.session_state.get("ultimo_datos_gc71", {}).get("profesor", ""), key="gc72_docente")
    with c2:
        periodo = st.text_input("Período académico", value=st.session_state.get("ultimo_datos_gc71", {}).get("periodo", "2026-1"), key="gc72_periodo")
    with c3:
        fecha_entrega = st.date_input("Fecha de entrega", value=date.today(), format="DD/MM/YYYY", key="gc72_fecha")

    st.subheader("2. Alimentación automática desde listado y calificaciones")
    c1, c2 = st.columns(2)
    with c1:
        listado_file = st.file_uploader("Listado tradicional de clase", type=["xls", "xlsx", "csv"], key="listado_gc72")
    with c2:
        calificaciones_file = st.file_uploader("Plantilla/calificaciones de mitad o final del curso", type=["xlsx", "xls", "csv"], key="calificaciones_gc72")
    datos_gc71 = st.session_state.get("ultimo_datos_gc71", {})
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        codigo_auto = st.text_input("Código asignatura", value=datos_gc71.get("codigo", ""), key="gc72_codigo_auto")
    with c2:
        grupo_auto = st.text_input("Grupo", value=datos_gc71.get("grupo", ""), key="gc72_grupo_auto")
    with c3:
        asignatura_auto = st.text_input("Asignatura", value=datos_gc71.get("asignatura", ""), key="gc72_asignatura_auto")
    with c4:
        corte_aprobacion = st.number_input("Nota mínima aprobatoria", min_value=0.0, max_value=5.0, value=3.0, step=0.1)
    sesiones_previas = st.session_state.get("ultimas_sesiones_gc71", pd.DataFrame())
    avance_sugerido = 0.0
    if isinstance(sesiones_previas, pd.DataFrame) and not sesiones_previas.empty:
        sesiones_realizadas = st.number_input("Sesiones realizadas para calcular avance", min_value=0, max_value=len(sesiones_previas), value=min(len(sesiones_previas), 0), step=1)
        avance_sugerido = round((sesiones_realizadas / len(sesiones_previas)) * 100, 1)
    else:
        avance_sugerido = st.number_input("% Avance en contenido", min_value=0.0, max_value=100.0, value=0.0, step=1.0)
    porcentaje_evaluado_manual = st.number_input("% Evaluado sugerido/manual", min_value=0.0, max_value=100.0, value=0.0, step=1.0)

    df_auto = pd.DataFrame(columns=COLUMNAS_GC72)
    if listado_file is not None or calificaciones_file is not None:
        try:
            listado_df = leer_listado_estudiantes(listado_file) if listado_file is not None else st.session_state.get("ultimo_listado_estudiantes", pd.DataFrame())
            cal_df, ev_df = leer_calificaciones(calificaciones_file) if calificaciones_file is not None else (pd.DataFrame(), pd.DataFrame())
            df_auto = resumen_gc72_desde_archivos(listado_df, cal_df, ev_df, codigo_auto, grupo_auto, asignatura_auto, avance_sugerido, porcentaje_evaluado_manual, corte_aprobacion)
            st.success("Resumen generado desde los archivos cargados. Revísalo antes de descargar: Excel no reemplaza criterio docente, pero sí evita trabajo de mula.")
            st.dataframe(df_auto, use_container_width=True)
        except Exception as e:
            st.error(f"No se pudo generar el resumen automático: {e}")

    st.subheader("3. Cursos")
    archivo = st.file_uploader("Opcional: cargar cursos desde Excel o CSV con columnas FD-GC72", type=["xlsx", "xls", "csv"], key="cursos_gc72")
    base = df_auto if not df_auto.empty else df_vacio(COLUMNAS_GC72, filas=3)
    if archivo is not None:
        try:
            if archivo.name.lower().endswith(".csv"):
                base = pd.read_csv(archivo)
            else:
                base = pd.read_excel(archivo)
        except Exception as e:
            st.error(f"No se pudo leer archivo de cursos: {e}")
    for col in COLUMNAS_GC72:
        if col not in base.columns:
            base[col] = ""
    base = base[COLUMNAS_GC72]

    config_columnas = {
        "Código": st.column_config.TextColumn("Código", width="small"),
        "Grupo": st.column_config.TextColumn("Grupo", width="small"),
        "Asignatura": st.column_config.TextColumn("Asignatura", width="medium"),
        "% Avance en contenido": st.column_config.NumberColumn("% Avance", min_value=0, max_value=100, step=1),
        "% Evaluado": st.column_config.NumberColumn("% Evaluado", min_value=0, max_value=100, step=1),
        "Estudiantes matriculados": st.column_config.NumberColumn("Matriculados", min_value=0, step=1),
        "Desertaron N°": st.column_config.NumberColumn("Desertaron N°", min_value=0, step=1),
        "Desertaron %": st.column_config.TextColumn("Desertaron %", disabled=calcular),
        "Aprueban evaluación parcial N°": st.column_config.NumberColumn("Aprueban parcial N°", min_value=0, step=1),
        "Aprueban evaluación parcial %": st.column_config.TextColumn("Aprueban parcial %", disabled=calcular),
        "Reprueban evaluación parcial N°": st.column_config.NumberColumn("Reprueban parcial N°", min_value=0, step=1),
        "Reprueban evaluación parcial %": st.column_config.TextColumn("Reprueban parcial %", disabled=calcular),
        "Aprueban a la fecha N°": st.column_config.NumberColumn("Aprueban fecha N°", min_value=0, step=1),
        "Aprueban a la fecha %": st.column_config.TextColumn("Aprueban fecha %", disabled=calcular),
        "Reprueban a la fecha N°": st.column_config.NumberColumn("Reprueban fecha N°", min_value=0, step=1),
        "Reprueban a la fecha %": st.column_config.TextColumn("Reprueban fecha %", disabled=calcular),
    }
    cursos_editados = st.data_editor(base, column_config=config_columnas, hide_index=True, num_rows="dynamic", use_container_width=True, key="tabla_gc72")
    cursos = normalizar_dataframe_gc72(cursos_editados, calcular_porcentajes=calcular)
    if cursos.empty:
        st.info("Agregue al menos un curso para activar el análisis y la descarga.")
        return

    st.subheader("4. Análisis descriptivo con preformas")
    analisis_por_curso: Dict[str, Dict[str, str]] = {}
    cursos_iteracion = cursos if modo_analisis == "Por cada curso" else cursos.head(1).copy()
    if modo_analisis == "Consolidado institucional":
        cursos_iteracion.loc[:, "Asignatura"] = "Consolidado institucional"
        cursos_iteracion.loc[:, "Código"] = ""
        cursos_iteracion.loc[:, "Grupo"] = ""
    for idx, row in cursos_iteracion.iterrows():
        key = curso_key(row, idx)
        titulo = str(row.get("Asignatura", "")).strip() or f"Curso {idx + 1}"
        subtitulo = f"{titulo} | Código: {row.get('Código', '')} | Grupo: {row.get('Grupo', '')}" if modo_analisis == "Por cada curso" else "Análisis consolidado para todos los cursos"
        with st.expander(subtitulo, expanded=(idx == 0)):
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                sel_pos = st.multiselect("Aspectos positivos", list(PREFORMAS_GC72["aspectos_positivos"].keys()), default=["Avance adecuado", "Aplicación práctica"], key=f"pos_{key}")
                add_pos = st.text_area("Texto adicional positivo", key=f"pos_add_{key}", height=90)
            with col_b:
                sel_inc = st.multiselect("Inconvenientes", list(PREFORMAS_GC72["inconvenientes"].keys()), default=[], key=f"inc_{key}")
                add_inc = st.text_area("Texto adicional de inconvenientes", key=f"inc_add_{key}", height=90)
            with col_c:
                sel_prop = st.multiselect("Propuestas metodológicas", list(PREFORMAS_GC72["propuestas"].keys()), default=["Seguimiento formativo", "Talleres aplicados"], key=f"prop_{key}")
                add_prop = st.text_area("Texto adicional de propuestas", key=f"prop_add_{key}", height=90)
            analisis_por_curso[key] = {
                "positivos": texto_preformas(sel_pos, "aspectos_positivos", add_pos),
                "inconvenientes": texto_preformas(sel_inc, "inconvenientes", add_inc),
                "propuestas": texto_preformas(sel_prop, "propuestas", add_prop),
            }
    bloques = construir_analisis(cursos if modo_analisis == "Por cada curso" else cursos_iteracion, analisis_por_curso, modo_analisis)

    st.subheader("5. Descargar")
    if not docente.strip():
        st.warning("Ingrese el nombre del docente antes de descargar.")
        return
    try:
        docx_bytes = crear_informe_gc72_docx(docente, periodo, fecha_entrega, cursos, bloques)
        nombre_base = nombre_archivo_seguro(docente, fecha_entrega, "FD_GC72_Informe_Academico")
        st.download_button("Descargar FD-GC72 en Word (.docx)", data=docx_bytes, file_name=f"{nombre_base}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
        if exportar_doc:
            doc_bytes = convertir_doc_si_es_posible(docx_bytes, nombre_base)
            if doc_bytes:
                st.download_button("Descargar FD-GC72 en Word 97-2003 (.doc)", data=doc_bytes, file_name=f"{nombre_base}.doc", mime="application/msword", use_container_width=True)
            else:
                st.info("No se pudo generar .doc. Instale LibreOffice o use .docx.")
    except Exception as e:
        st.error(f"No se pudo generar FD-GC72: {e}")


def ui_ayuda(st):
    st.header("Flujo recomendado")
    st.markdown(
        """
1. **Inicio del curso:** diligencie FD-GC71, defina módulos, intensidad, horario y evaluación concertada. Descargue la guía y la plantilla de evaluación.
2. **Durante el curso:** cargue notas en la plantilla Excel. Mantenga actualizada la columna Estado si hay retiros o deserciones.
3. **Mitad/final del curso:** suba el listado tradicional y la plantilla de evaluación al módulo FD-GC72. La app calcula matriculados, desertores, aprobados, reprobados y porcentajes.
4. **Cierre:** revise la tabla, ajuste el análisis descriptivo con preformas y descargue el informe académico.

Columnas esperadas del listado tradicional: **NOMBRE COMPLETO, DOCUMENTO, PLAN, OBSERVACIÓN, CORREO**. Si el archivo trae más columnas, no pasa nada; la app toma lo que necesita.
        """
    )


# -----------------------------------------------------------------------------
# Seguridad, login, perfiles y auditoría local
# -----------------------------------------------------------------------------
ROLES_PERMISOS = {
    "Administrador": {
        "descripcion": "Control total: usuarios, perfiles, auditoría y generación de formatos.",
        "modulos": ["Inicio", "FD-GC71 - Planeación", "FD-GC72 - Informe académico", "Usuarios y perfiles", "Auditoría", "Mi cuenta", "Ayuda / flujo recomendado"],
    },
    "Coordinador": {
        "descripcion": "Revisión académica: puede usar formatos, ver auditoría y acompañar cierres.",
        "modulos": ["Inicio", "FD-GC71 - Planeación", "FD-GC72 - Informe académico", "Auditoría", "Mi cuenta", "Ayuda / flujo recomendado"],
    },
    "Docente": {
        "descripcion": "Operación docente: planeación, evaluación, informes y descarga de soportes.",
        "modulos": ["Inicio", "FD-GC71 - Planeación", "FD-GC72 - Informe académico", "Mi cuenta", "Ayuda / flujo recomendado"],
    },
    "Consulta": {
        "descripcion": "Acceso de lectura al flujo y orientación operativa.",
        "modulos": ["Inicio", "Mi cuenta", "Ayuda / flujo recomendado"],
    },
}


def ahora_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# La conexión productiva fue definida al inicio del archivo.
# Se conserva este punto solo por compatibilidad con la versión anterior.
# No redefinir conexion_db aquí.


def hash_password(password: str, salt: Optional[str] = None) -> Tuple[str, str]:
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 240_000)
    return salt, digest.hex()


def verificar_password(password: str, salt: str, password_hash: str) -> bool:
    _, digest = hash_password(password, salt)
    return hmac.compare_digest(digest, password_hash)


def validar_password_seguro(password: str):
    """Política de contraseña. En producción exige mayor rigor; en local conserva usabilidad."""
    min_len = 12 if get_app_env() == "production" else 8
    if len(password or "") < min_len:
        raise ValueError(f"La contraseña debe tener mínimo {min_len} caracteres.")
    if get_app_env() == "production":
        reglas = [
            (re.search(r"[A-ZÁÉÍÓÚÑ]", password or ""), "una mayúscula"),
            (re.search(r"[a-záéíóúñ]", password or ""), "una minúscula"),
            (re.search(r"\d", password or ""), "un número"),
            (re.search(r"[^A-Za-z0-9ÁÉÍÓÚáéíóúÑñ]", password or ""), "un carácter especial"),
        ]
        faltantes = [nombre for ok, nombre in reglas if not ok]
        if faltantes:
            raise ValueError("La contraseña debe incluir " + ", ".join(faltantes) + ".")



def init_db():
    conn = conexion_db()
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT NOT NULL UNIQUE,
            nombre_completo TEXT NOT NULL,
            email TEXT,
            rol TEXT NOT NULL,
            salt TEXT NOT NULL,
            password_hash TEXT NOT NULL,
            activo INTEGER NOT NULL DEFAULT 1,
            debe_cambiar_clave INTEGER NOT NULL DEFAULT 0,
            creado_en TEXT NOT NULL,
            actualizado_en TEXT NOT NULL,
            ultimo_login TEXT
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS auditoria (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT NOT NULL,
            usuario TEXT,
            rol TEXT,
            accion TEXT NOT NULL,
            detalle TEXT
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS proyectos_guardados (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario_id INTEGER NOT NULL,
            tipo TEXT NOT NULL,
            titulo TEXT NOT NULL,
            payload_json TEXT NOT NULL,
            creado_en TEXT NOT NULL,
            actualizado_en TEXT NOT NULL,
            FOREIGN KEY(usuario_id) REFERENCES usuarios(id)
        )
        """
    )
    conn.commit()

    total = cur.execute("SELECT COUNT(*) AS n FROM usuarios").fetchone()["n"]
    if total == 0:
        salt, digest = hash_password("Admin123*")
        cur.execute(
            """
            INSERT INTO usuarios(usuario, nombre_completo, email, rol, salt, password_hash, activo, debe_cambiar_clave, creado_en, actualizado_en)
            VALUES (?, ?, ?, ?, ?, ?, 1, 1, ?, ?)
            """,
            ("admin", "Administrador del sistema", "", "Administrador", salt, digest, ahora_iso(), ahora_iso()),
        )
        conn.commit()
    conn.close()


def registrar_auditoria(accion: str, detalle: str = ""):
    try:
        usuario = st.session_state.get("auth_user", {}) if "st" in globals() else {}
    except Exception:
        usuario = {}
    conn = conexion_db()
    conn.execute(
        "INSERT INTO auditoria(fecha, usuario, rol, accion, detalle) VALUES (?, ?, ?, ?, ?)",
        (ahora_iso(), usuario.get("usuario", "sistema"), usuario.get("rol", ""), accion, detalle[:1500] if detalle else ""),
    )
    conn.commit()
    conn.close()


def autenticar_usuario(usuario: str, password: str) -> Optional[Dict[str, str]]:
    """Autenticación con bloqueo temporal por intentos fallidos."""
    usuario_clean = usuario.strip()
    conn = conexion_db()
    row = conn.execute("SELECT * FROM usuarios WHERE lower(usuario)=lower(?)", (usuario_clean,)).fetchone()
    if not row:
        conn.close()
        return None
    if not row["activo"]:
        conn.close()
        return {"error": "Usuario inactivo. Solicite habilitación al administrador."}

    bloqueado_hasta = None
    try:
        bloqueado_hasta = row["bloqueado_hasta"]
    except Exception:
        bloqueado_hasta = None
    if bloqueado_hasta:
        try:
            until = datetime.fromisoformat(str(bloqueado_hasta))
            if datetime.now() < until:
                conn.close()
                return {"error": f"Usuario bloqueado temporalmente hasta {until.strftime('%H:%M:%S')}."}
        except Exception:
            pass

    if not verificar_password(password, row["salt"], row["password_hash"]):
        try:
            intentos = int(row["intentos_fallidos"] or 0) + 1
        except Exception:
            intentos = 1
        bloqueado = None
        if intentos >= 5:
            bloqueado = (datetime.now() + timedelta(minutes=15)).isoformat(timespec="seconds")
            intentos = 0
        try:
            conn.execute("UPDATE usuarios SET intentos_fallidos=?, bloqueado_hasta=?, actualizado_en=? WHERE id=?", (intentos, bloqueado, ahora_iso(), row["id"]))
            conn.commit()
        except Exception:
            pass
        conn.close()
        return None

    conn.execute("UPDATE usuarios SET ultimo_login=?, intentos_fallidos=0, bloqueado_hasta=NULL WHERE id=?", (ahora_iso(), row["id"]))
    conn.commit()
    conn.close()
    return {
        "id": row["id"],
        "usuario": row["usuario"],
        "nombre_completo": row["nombre_completo"],
        "email": row["email"] or "",
        "rol": row["rol"],
        "debe_cambiar_clave": bool(row["debe_cambiar_clave"]),
    }


def crear_usuario(usuario: str, nombre: str, email: str, rol: str, password: str, activo: bool = True, debe_cambiar: bool = True):
    if rol not in ROLES_PERMISOS:
        raise ValueError("Rol no válido.")
    validar_password_seguro(password)
    salt, digest = hash_password(password)
    conn = conexion_db()
    conn.execute(
        """
        INSERT INTO usuarios(usuario, nombre_completo, email, rol, salt, password_hash, activo, debe_cambiar_clave, creado_en, actualizado_en)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (usuario.strip(), nombre.strip(), email.strip(), rol, salt, digest, int(activo), int(debe_cambiar), ahora_iso(), ahora_iso()),
    )
    conn.commit()
    conn.close()
    registrar_auditoria("Crear usuario", f"Usuario={usuario.strip()} | Rol={rol}")


def actualizar_usuario(user_id: int, nombre: str, email: str, rol: str, activo: bool):
    if rol not in ROLES_PERMISOS:
        raise ValueError("Rol no válido.")
    conn = conexion_db()
    conn.execute(
        "UPDATE usuarios SET nombre_completo=?, email=?, rol=?, activo=?, actualizado_en=? WHERE id=?",
        (nombre.strip(), email.strip(), rol, int(activo), ahora_iso(), user_id),
    )
    conn.commit()
    conn.close()
    registrar_auditoria("Actualizar usuario", f"ID={user_id} | Rol={rol} | Activo={activo}")


def cambiar_password(user_id: int, password: str, forzar_cambio: bool = False):
    validar_password_seguro(password)
    salt, digest = hash_password(password)
    conn = conexion_db()
    conn.execute(
        "UPDATE usuarios SET salt=?, password_hash=?, debe_cambiar_clave=?, actualizado_en=? WHERE id=?",
        (salt, digest, int(forzar_cambio), ahora_iso(), user_id),
    )
    conn.commit()
    conn.close()
    registrar_auditoria("Cambiar contraseña", f"ID={user_id}")


def listar_usuarios() -> pd.DataFrame:
    df = read_sql_df(
        "SELECT id, usuario, nombre_completo, email, rol, activo, debe_cambiar_clave, creado_en, actualizado_en, ultimo_login FROM usuarios ORDER BY rol, usuario"
    )
    if not df.empty:
        df["activo"] = df["activo"].map({1: "Sí", 0: "No"})
        df["debe_cambiar_clave"] = df["debe_cambiar_clave"].map({1: "Sí", 0: "No"})
    return df


def tiene_permiso(modulo: str) -> bool:
    user = st.session_state.get("auth_user")
    if not user:
        return False
    return modulo in ROLES_PERMISOS.get(user.get("rol"), {}).get("modulos", [])


def pantalla_login(st):
    st.set_page_config(page_title="Login | Gestor FD-GC71 / FD-GC72", layout="centered")
    st.title("Ingreso al gestor académico")
    st.caption("FD-GC71 / FD-GC72 con control de acceso por perfiles.")
    with st.form("login_form", clear_on_submit=False):
        usuario = st.text_input("Usuario", value="admin")
        password = st.text_input("Contraseña", type="password")
        entrar = st.form_submit_button("Entrar", use_container_width=True)
    if entrar:
        user = autenticar_usuario(usuario, password)
        if user and not user.get("error"):
            st.session_state["auth_user"] = user
            registrar_auditoria("Login", "Ingreso correcto")
            st.rerun()
        elif user and user.get("error"):
            st.error(user["error"])
        else:
            st.error("Usuario o contraseña incorrectos.")
    with st.expander("Primera instalación"):
        if get_app_env() == "production":
            st.info("Las credenciales iniciales se leen desde los secretos INITIAL_ADMIN_USER e INITIAL_ADMIN_PASSWORD.")
        else:
            admin_user, _, _, _ = initial_admin_config()
            st.write(f"Usuario inicial: `{admin_user}`")
            st.write("Contraseña inicial: definida en secrets o, en local, `Admin123*`.")
        st.warning("Cambie la contraseña apenas ingrese y cree usuarios nominales. La puerta abierta también sirve para que entre el incendio.")


def ui_inicio(st):
    user = st.session_state.get("auth_user", {})
    st.header("Panel de inicio")
    c1, c2, c3 = st.columns(3)
    c1.metric("Usuario", user.get("usuario", ""))
    c2.metric("Perfil", user.get("rol", ""))
    c3.metric("Estado", "Activo")
    st.info(ROLES_PERMISOS.get(user.get("rol"), {}).get("descripcion", ""))
    if get_app_env() == "production" and not usar_postgres():
        st.error("APP_ENV está en production pero no hay DATABASE_URL. Configure PostgreSQL antes de operar con datos reales.")
    elif usar_postgres():
        st.success("Persistencia productiva activa: PostgreSQL externo configurado.")
    st.subheader("Permisos habilitados")
    permisos = ROLES_PERMISOS.get(user.get("rol"), {}).get("modulos", [])
    st.write(", ".join([p for p in permisos if p not in ["Inicio", "Mi cuenta"]]))
    if user.get("debe_cambiar_clave"):
        st.warning("Debe cambiar la contraseña inicial desde el módulo Mi cuenta antes de operar formalmente.")

    st.subheader("Ruta operativa recomendada")
    st.markdown(
        """
- **Inicio del curso:** planee la asignatura en FD-GC71, defina unidades, intensidad, horario y evaluación concertada.
- **Durante el curso:** use la plantilla de evaluación descargada para registrar notas y estados.
- **Mitad y final:** cargue listado y calificaciones en FD-GC72 para calcular avance, evaluación, aprobados, reprobados y desertores.
- **Cierre:** descargue el informe académico con análisis descriptivo por curso.
        """
    )


def ui_mi_cuenta(st):
    user = st.session_state.get("auth_user", {})
    st.header("Mi cuenta")
    st.write(f"**Nombre:** {user.get('nombre_completo', '')}")
    st.write(f"**Usuario:** {user.get('usuario', '')}")
    st.write(f"**Correo:** {user.get('email', '')}")
    st.write(f"**Perfil:** {user.get('rol', '')}")
    with st.form("form_cambiar_clave"):
        nueva = st.text_input("Nueva contraseña", type="password")
        confirmar = st.text_input("Confirmar contraseña", type="password")
        enviar = st.form_submit_button("Actualizar contraseña", use_container_width=True)
    if enviar:
        if nueva != confirmar:
            st.error("Las contraseñas no coinciden.")
        else:
            try:
                cambiar_password(int(user["id"]), nueva, forzar_cambio=False)
                st.session_state["auth_user"]["debe_cambiar_clave"] = False
                st.success("Contraseña actualizada.")
            except Exception as e:
                st.error(str(e))


def ui_admin_usuarios(st):
    st.header("Usuarios y perfiles")
    st.caption("Administración local de usuarios. La base queda en app_data/fdgc_app.sqlite3.")
    usuarios = listar_usuarios()
    st.dataframe(usuarios, use_container_width=True, hide_index=True)

    tab_crear, tab_editar, tab_clave = st.tabs(["Crear usuario", "Editar perfil", "Resetear contraseña"])
    with tab_crear:
        with st.form("crear_usuario"):
            c1, c2 = st.columns(2)
            with c1:
                usuario = st.text_input("Usuario nuevo")
                nombre = st.text_input("Nombre completo")
                email = st.text_input("Correo")
            with c2:
                rol = st.selectbox("Perfil", list(ROLES_PERMISOS.keys()), index=2)
                password = st.text_input("Contraseña inicial", type="password", value="Cambio123*")
                activo = st.checkbox("Activo", value=True)
                debe_cambiar = st.checkbox("Forzar cambio de contraseña", value=True)
            enviar = st.form_submit_button("Crear usuario", use_container_width=True)
        if enviar:
            try:
                crear_usuario(usuario, nombre, email, rol, password, activo, debe_cambiar)
                st.success("Usuario creado.")
                st.rerun()
            except Exception as e:
                st.error(f"No se pudo crear: {e}")

    with tab_editar:
        if usuarios.empty:
            st.info("No hay usuarios.")
        else:
            ids = usuarios["id"].tolist()
            seleccionado = st.selectbox("Usuario a editar", ids, format_func=lambda x: usuarios.loc[usuarios["id"] == x, "usuario"].iloc[0], key="editar_usuario_id")
            fila = usuarios.loc[usuarios["id"] == seleccionado].iloc[0]
            with st.form("editar_usuario"):
                nombre = st.text_input("Nombre completo", value=str(fila["nombre_completo"]))
                email = st.text_input("Correo", value=str(fila["email"]))
                rol = st.selectbox("Perfil", list(ROLES_PERMISOS.keys()), index=list(ROLES_PERMISOS.keys()).index(str(fila["rol"])))
                activo = st.checkbox("Activo", value=(str(fila["activo"]) == "Sí"))
                enviar = st.form_submit_button("Guardar cambios", use_container_width=True)
            if enviar:
                try:
                    actualizar_usuario(int(seleccionado), nombre, email, rol, activo)
                    st.success("Usuario actualizado.")
                    st.rerun()
                except Exception as e:
                    st.error(f"No se pudo actualizar: {e}")

    with tab_clave:
        if usuarios.empty:
            st.info("No hay usuarios.")
        else:
            seleccionado = st.selectbox("Usuario", usuarios["id"].tolist(), format_func=lambda x: usuarios.loc[usuarios["id"] == x, "usuario"].iloc[0], key="reset_usuario_id")
            with st.form("reset_clave"):
                nueva = st.text_input("Nueva contraseña", type="password", value="Cambio123*")
                forzar = st.checkbox("Forzar cambio al iniciar", value=True)
                enviar = st.form_submit_button("Resetear contraseña", use_container_width=True)
            if enviar:
                try:
                    cambiar_password(int(seleccionado), nueva, forzar_cambio=forzar)
                    st.success("Contraseña reseteada.")
                except Exception as e:
                    st.error(str(e))

    st.subheader("Matriz de perfiles")
    matriz = []
    for rol, cfg in ROLES_PERMISOS.items():
        matriz.append({"Perfil": rol, "Descripción": cfg["descripcion"], "Módulos": ", ".join(cfg["modulos"])})
    st.dataframe(pd.DataFrame(matriz), use_container_width=True, hide_index=True)


def ui_auditoria(st):
    st.header("Auditoría")
    st.caption("Registro básico de ingresos y cambios de seguridad.")
    df = read_sql_df("SELECT fecha, usuario, rol, accion, detalle FROM auditoria ORDER BY id DESC LIMIT 500")
    st.dataframe(df, use_container_width=True, hide_index=True)
    csv = df.to_csv(index=False).encode("utf-8-sig")
    st.download_button("Descargar auditoría CSV", data=csv, file_name="auditoria_fdgc.csv", mime="text/csv", use_container_width=True)


def main():
    import streamlit as st
    globals()["st"] = st
    init_db()

    if "auth_user" not in st.session_state:
        pantalla_login(st)
        return

    user = st.session_state.get("auth_user", {})
    st.set_page_config(page_title="Gestor FD-GC71 / FD-GC72", layout="wide")
    st.title("Gestor académico FD-GC71 / FD-GC72")
    st.caption("Planeación de clases, concertación de evaluación, plantilla de notas, informe académico y control de acceso por perfiles.")

    with st.sidebar:
        st.markdown(f"**{user.get('nombre_completo', '')}**")
        st.caption(f"Perfil: {user.get('rol', '')}")
        if st.button("Cerrar sesión", use_container_width=True):
            registrar_auditoria("Logout", "Cierre de sesión")
            st.session_state.pop("auth_user", None)
            st.rerun()
        st.divider()
        modulos = ROLES_PERMISOS.get(user.get("rol"), {}).get("modulos", ["Inicio", "Mi cuenta", "Ayuda / flujo recomendado"])
        pagina = st.radio("Módulo", modulos)

    if not tiene_permiso(pagina):
        st.error("Este perfil no tiene permisos para abrir este módulo.")
        return

    if pagina == "Inicio":
        ui_inicio(st)
    elif pagina.startswith("FD-GC71"):
        ui_gc71(st)
    elif pagina.startswith("FD-GC72"):
        ui_gc72(st)
    elif pagina == "Usuarios y perfiles":
        ui_admin_usuarios(st)
    elif pagina == "Auditoría":
        ui_auditoria(st)
    elif pagina == "Mi cuenta":
        ui_mi_cuenta(st)
    else:
        ui_ayuda(st)


# =============================================================================
# VERSIÓN ENTERPRISE / ESTADIO: expediente, persistencia, control, evidencias
# =============================================================================
# Esta sección se define al final para sobreescribir/elevar algunas funciones
# anteriores sin romper la compatibilidad del prototipo inicial.

import zipfile
import uuid
from html import escape

EVIDENCE_DIR = DATA_DIR / "evidencias"
EXPORT_DIR = DATA_DIR / "exportaciones"
BACKUP_DIR = DATA_DIR / "backups"

MODULO_CENTRO = "Centro de control"
MODULO_EXPEDIENTE = "Expediente académico"
MODULO_PLANEADOR = "Planeador superior"
MODULO_EVIDENCIAS = "Evidencias y soportes"
MODULO_VALIDADOR = "Validador institucional"
MODULO_BACKUP = "Copias y restauración"
MODULO_DIAGNOSTICO = "Diagnóstico productivo"

ROLES_PERMISOS = {
    "Administrador": {
        "descripcion": "Gobierno total del sistema: usuarios, cursos, formatos, evidencias, auditoría, respaldos y parametrización.",
        "modulos": ["Inicio", MODULO_CENTRO, MODULO_EXPEDIENTE, MODULO_PLANEADOR, "FD-GC71 - Planeación", "FD-GC72 - Informe académico", MODULO_EVIDENCIAS, MODULO_VALIDADOR, MODULO_BACKUP, MODULO_DIAGNOSTICO, "Usuarios y perfiles", "Auditoría", "Mi cuenta", "Ayuda / flujo recomendado"],
    },
    "Coordinador": {
        "descripcion": "Revisión y cierre académico: puede controlar cursos, validar planeación, revisar evidencias, generar informes y consultar auditoría.",
        "modulos": ["Inicio", MODULO_CENTRO, MODULO_EXPEDIENTE, MODULO_PLANEADOR, "FD-GC71 - Planeación", "FD-GC72 - Informe académico", MODULO_EVIDENCIAS, MODULO_VALIDADOR, MODULO_DIAGNOSTICO, "Auditoría", "Mi cuenta", "Ayuda / flujo recomendado"],
    },
    "Docente": {
        "descripcion": "Operación docente: planeación, evaluación, evidencias, informes y descarga de soportes de sus cursos.",
        "modulos": ["Inicio", MODULO_CENTRO, MODULO_EXPEDIENTE, MODULO_PLANEADOR, "FD-GC71 - Planeación", "FD-GC72 - Informe académico", MODULO_EVIDENCIAS, MODULO_VALIDADOR, "Mi cuenta", "Ayuda / flujo recomendado"],
    },
    "Consulta": {
        "descripcion": "Consulta controlada: lectura de ruta operativa, cuenta propia y ayuda.",
        "modulos": ["Inicio", MODULO_CENTRO, "Mi cuenta", "Ayuda / flujo recomendado"],
    },
}


def safe_json_loads(value, default=None):
    if default is None:
        default = {}
    if value in (None, "", float("nan")):
        return default
    try:
        return json.loads(value)
    except Exception:
        return default


def df_to_payload(df: pd.DataFrame) -> List[Dict[str, str]]:
    if df is None or df.empty:
        return []
    cleaned = df.copy().fillna("")
    return cleaned.astype(str).to_dict(orient="records")


def payload_to_df(payload, columnas: Optional[List[str]] = None) -> pd.DataFrame:
    if not payload:
        return pd.DataFrame(columns=columnas or [])
    df = pd.DataFrame(payload)
    if columnas:
        for c in columnas:
            if c not in df.columns:
                df[c] = ""
        df = df[columnas]
    return df


def init_db():
    """Inicializa la base en modo local o productivo.

    - SQLite: desarrollo / demo / uso local.
    - PostgreSQL: Streamlit Cloud o servidor institucional con persistencia real.
    """
    DATA_DIR.mkdir(exist_ok=True)
    EVIDENCE_DIR.mkdir(exist_ok=True)
    EXPORT_DIR.mkdir(exist_ok=True)
    BACKUP_DIR.mkdir(exist_ok=True)
    conn = conexion_db()

    if usar_postgres():
        ddl = {
            "usuarios": """
                CREATE TABLE IF NOT EXISTS usuarios (
                    id SERIAL PRIMARY KEY,
                    usuario TEXT UNIQUE NOT NULL,
                    nombre_completo TEXT NOT NULL,
                    email TEXT,
                    rol TEXT NOT NULL,
                    salt TEXT NOT NULL,
                    password_hash TEXT NOT NULL,
                    activo INTEGER NOT NULL DEFAULT 1,
                    debe_cambiar_clave INTEGER NOT NULL DEFAULT 1,
                    creado_en TEXT NOT NULL,
                    actualizado_en TEXT,
                    ultimo_login TEXT
                )
            """,
            "auditoria": """
                CREATE TABLE IF NOT EXISTS auditoria (
                    id SERIAL PRIMARY KEY,
                    fecha TEXT NOT NULL,
                    usuario TEXT,
                    rol TEXT,
                    accion TEXT NOT NULL,
                    detalle TEXT
                )
            """,
            "cursos": """
                CREATE TABLE IF NOT EXISTS cursos (
                    id SERIAL PRIMARY KEY,
                    codigo TEXT,
                    grupo TEXT,
                    asignatura TEXT NOT NULL,
                    programa TEXT,
                    periodo TEXT,
                    profesor TEXT,
                    email_profesor TEXT,
                    creditos TEXT,
                    htp DOUBLE PRECISION DEFAULT 0,
                    hti DOUBLE PRECISION DEFAULT 0,
                    fecha_inicio TEXT,
                    fecha_fin TEXT,
                    estado TEXT DEFAULT 'Planeación',
                    avance_contenido DOUBLE PRECISION DEFAULT 0,
                    avance_evaluado DOUBLE PRECISION DEFAULT 0,
                    propietario_usuario TEXT,
                    creado_por TEXT,
                    creado_en TEXT NOT NULL,
                    actualizado_en TEXT,
                    payload_json TEXT DEFAULT '{}'
                )
            """,
            "evidencias": """
                CREATE TABLE IF NOT EXISTS evidencias (
                    id SERIAL PRIMARY KEY,
                    curso_id INTEGER REFERENCES cursos(id) ON DELETE SET NULL,
                    tipo TEXT,
                    nombre_original TEXT NOT NULL,
                    nombre_archivo TEXT NOT NULL,
                    mime TEXT,
                    tamano INTEGER DEFAULT 0,
                    descripcion TEXT,
                    subido_por TEXT,
                    subido_en TEXT NOT NULL,
                    contenido_b64 TEXT
                )
            """,
            "artefactos": """
                CREATE TABLE IF NOT EXISTS artefactos (
                    id SERIAL PRIMARY KEY,
                    curso_id INTEGER REFERENCES cursos(id) ON DELETE SET NULL,
                    tipo TEXT NOT NULL,
                    nombre_archivo TEXT NOT NULL,
                    descripcion TEXT,
                    generado_por TEXT,
                    generado_en TEXT NOT NULL
                )
            """,
        }
    else:
        ddl = {
            "usuarios": """
                CREATE TABLE IF NOT EXISTS usuarios (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    usuario TEXT UNIQUE NOT NULL,
                    nombre_completo TEXT NOT NULL,
                    email TEXT,
                    rol TEXT NOT NULL,
                    salt TEXT NOT NULL,
                    password_hash TEXT NOT NULL,
                    activo INTEGER NOT NULL DEFAULT 1,
                    debe_cambiar_clave INTEGER NOT NULL DEFAULT 1,
                    creado_en TEXT NOT NULL,
                    actualizado_en TEXT,
                    ultimo_login TEXT
                )
            """,
            "auditoria": """
                CREATE TABLE IF NOT EXISTS auditoria (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    fecha TEXT NOT NULL,
                    usuario TEXT,
                    rol TEXT,
                    accion TEXT NOT NULL,
                    detalle TEXT
                )
            """,
            "cursos": """
                CREATE TABLE IF NOT EXISTS cursos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo TEXT,
                    grupo TEXT,
                    asignatura TEXT NOT NULL,
                    programa TEXT,
                    periodo TEXT,
                    profesor TEXT,
                    email_profesor TEXT,
                    creditos TEXT,
                    htp REAL DEFAULT 0,
                    hti REAL DEFAULT 0,
                    fecha_inicio TEXT,
                    fecha_fin TEXT,
                    estado TEXT DEFAULT 'Planeación',
                    avance_contenido REAL DEFAULT 0,
                    avance_evaluado REAL DEFAULT 0,
                    propietario_usuario TEXT,
                    creado_por TEXT,
                    creado_en TEXT NOT NULL,
                    actualizado_en TEXT,
                    payload_json TEXT DEFAULT '{}'
                )
            """,
            "evidencias": """
                CREATE TABLE IF NOT EXISTS evidencias (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    curso_id INTEGER,
                    tipo TEXT,
                    nombre_original TEXT NOT NULL,
                    nombre_archivo TEXT NOT NULL,
                    mime TEXT,
                    tamano INTEGER DEFAULT 0,
                    descripcion TEXT,
                    subido_por TEXT,
                    subido_en TEXT NOT NULL,
                    contenido_b64 TEXT,
                    FOREIGN KEY(curso_id) REFERENCES cursos(id)
                )
            """,
            "artefactos": """
                CREATE TABLE IF NOT EXISTS artefactos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    curso_id INTEGER,
                    tipo TEXT NOT NULL,
                    nombre_archivo TEXT NOT NULL,
                    descripcion TEXT,
                    generado_por TEXT,
                    generado_en TEXT NOT NULL,
                    FOREIGN KEY(curso_id) REFERENCES cursos(id)
                )
            """,
        }

    for sql in ddl.values():
        conn.execute(sql)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_cursos_estado ON cursos(estado)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_cursos_owner ON cursos(propietario_usuario)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_evidencias_curso ON evidencias(curso_id)")
    conn.commit()
    conn.close()

    # Migraciones defensivas para instalaciones existentes.
    add_column_if_missing("evidencias", "contenido_b64", "TEXT")
    add_column_if_missing("usuarios", "intentos_fallidos", "INTEGER DEFAULT 0")
    add_column_if_missing("usuarios", "bloqueado_hasta", "TEXT")

    row = db_execute("SELECT COUNT(*) AS n FROM usuarios", fetchone=True)
    count = row["n"] if row is not None else 0
    if count == 0:
        admin_user, admin_password, admin_name, admin_email = initial_admin_config()
        salt, digest = hash_password(admin_password)
        db_execute(
            """
            INSERT INTO usuarios(usuario, nombre_completo, email, rol, salt, password_hash, activo, debe_cambiar_clave, creado_en, actualizado_en)
            VALUES (?, ?, ?, ?, ?, ?, 1, 1, ?, ?)
            """,
            (admin_user, admin_name, admin_email, "Administrador", salt, digest, ahora_iso(), ahora_iso()),
        )


def cursos_visibles_query(user: Dict[str, str]) -> Tuple[str, Tuple]:
    rol = user.get("rol", "")
    if rol in ("Administrador", "Coordinador"):
        return "SELECT * FROM cursos ORDER BY actualizado_en DESC, creado_en DESC", tuple()
    if rol == "Docente":
        return "SELECT * FROM cursos WHERE propietario_usuario=? OR creado_por=? ORDER BY actualizado_en DESC, creado_en DESC", (user.get("usuario", ""), user.get("usuario", ""))
    return "SELECT * FROM cursos ORDER BY actualizado_en DESC, creado_en DESC LIMIT 25", tuple()


def listar_cursos_visibles() -> pd.DataFrame:
    user = st.session_state.get("auth_user", {})
    q, params = cursos_visibles_query(user)
    df = read_sql_df(q, params=params)
    return df


def get_curso(curso_id: int) -> Optional[Dict]:
    conn = conexion_db()
    row = conn.execute("SELECT * FROM cursos WHERE id=?", (curso_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def upsert_curso(curso_id: Optional[int], datos: Dict[str, str], payload: Optional[Dict] = None) -> int:
    user = st.session_state.get("auth_user", {})
    now = ahora_iso()
    payload_json = json.dumps(payload or datos.get("payload", {}), ensure_ascii=False)
    values = {
        "codigo": datos.get("codigo", "").strip(),
        "grupo": datos.get("grupo", "").strip(),
        "asignatura": datos.get("asignatura", "").strip() or "Asignatura sin nombre",
        "programa": datos.get("programa", "").strip(),
        "periodo": datos.get("periodo", "").strip(),
        "profesor": datos.get("profesor", "").strip(),
        "email_profesor": datos.get("correo", datos.get("email_profesor", "")).strip(),
        "creditos": str(datos.get("creditos", "")).strip(),
        "htp": limpiar_numero(datos.get("htp", datos.get("horas_presenciales", 0))) or 0,
        "hti": limpiar_numero(datos.get("hti", datos.get("horas_independientes", 0))) or 0,
        "fecha_inicio": str(datos.get("fecha_inicio", "")),
        "fecha_fin": str(datos.get("fecha_fin", "")),
        "estado": datos.get("estado", "Planeación"),
        "avance_contenido": limpiar_numero(datos.get("avance_contenido", 0)) or 0,
        "avance_evaluado": limpiar_numero(datos.get("avance_evaluado", 0)) or 0,
        "propietario_usuario": datos.get("propietario_usuario", user.get("usuario", "")),
        "payload_json": payload_json,
    }
    conn = conexion_db()
    if curso_id:
        conn.execute(
            """
            UPDATE cursos SET codigo=?, grupo=?, asignatura=?, programa=?, periodo=?, profesor=?, email_profesor=?, creditos=?, htp=?, hti=?, fecha_inicio=?, fecha_fin=?, estado=?, avance_contenido=?, avance_evaluado=?, propietario_usuario=?, actualizado_en=?, payload_json=?
            WHERE id=?
            """,
            (values["codigo"], values["grupo"], values["asignatura"], values["programa"], values["periodo"], values["profesor"], values["email_profesor"], values["creditos"], values["htp"], values["hti"], values["fecha_inicio"], values["fecha_fin"], values["estado"], values["avance_contenido"], values["avance_evaluado"], values["propietario_usuario"], now, values["payload_json"], int(curso_id)),
        )
        new_id = int(curso_id)
        accion = "Actualizar expediente"
    else:
        params_insert = (values["codigo"], values["grupo"], values["asignatura"], values["programa"], values["periodo"], values["profesor"], values["email_profesor"], values["creditos"], values["htp"], values["hti"], values["fecha_inicio"], values["fecha_fin"], values["estado"], values["avance_contenido"], values["avance_evaluado"], values["propietario_usuario"], user.get("usuario", ""), now, now, values["payload_json"])
        if usar_postgres():
            cur = conn.execute(
                """
                INSERT INTO cursos(codigo, grupo, asignatura, programa, periodo, profesor, email_profesor, creditos, htp, hti, fecha_inicio, fecha_fin, estado, avance_contenido, avance_evaluado, propietario_usuario, creado_por, creado_en, actualizado_en, payload_json)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                RETURNING id
                """,
                params_insert,
            )
            new_id = int(cur.fetchone()["id"])
        else:
            cur = conn.execute(
                """
                INSERT INTO cursos(codigo, grupo, asignatura, programa, periodo, profesor, email_profesor, creditos, htp, hti, fecha_inicio, fecha_fin, estado, avance_contenido, avance_evaluado, propietario_usuario, creado_por, creado_en, actualizado_en, payload_json)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                params_insert,
            )
            new_id = int(cur.lastrowid)
        accion = "Crear expediente"
    conn.commit()
    conn.close()
    registrar_auditoria(accion, f"Curso ID={new_id} | {values['asignatura']} | Grupo {values['grupo']}")
    return new_id


def eliminar_curso(curso_id: int):
    conn = conexion_db()
    curso = conn.execute("SELECT asignatura, grupo FROM cursos WHERE id=?", (curso_id,)).fetchone()
    conn.execute("DELETE FROM cursos WHERE id=?", (curso_id,))
    conn.commit()
    conn.close()
    registrar_auditoria("Eliminar expediente", f"Curso ID={curso_id} | {dict(curso) if curso else ''}")


def registrar_artefacto(curso_id: Optional[int], tipo: str, nombre_archivo: str, descripcion: str = ""):
    conn = conexion_db()
    user = st.session_state.get("auth_user", {})
    conn.execute(
        "INSERT INTO artefactos(curso_id, tipo, nombre_archivo, descripcion, generado_por, generado_en) VALUES (?, ?, ?, ?, ?, ?)",
        (curso_id, tipo, nombre_archivo, descripcion, user.get("usuario", ""), ahora_iso()),
    )
    conn.commit()
    conn.close()


def seleccionar_curso_widget(label="Curso / expediente", key="curso_sel", incluir_nuevo=False) -> Optional[int]:
    df = listar_cursos_visibles()
    opciones = []
    if incluir_nuevo:
        opciones.append(None)
    opciones.extend(df["id"].tolist() if not df.empty else [])
    if not opciones:
        st.info("Aún no hay cursos registrados. Cree uno desde Expediente académico o desde Planeador superior.")
        return None
    def fmt(x):
        if x is None:
            return "➕ Crear nuevo expediente"
        fila = df.loc[df["id"] == x].iloc[0]
        return f"#{x} | {fila.get('asignatura','')} | Grupo {fila.get('grupo','')} | {fila.get('periodo','')} | {fila.get('estado','')}"
    return st.selectbox(label, opciones, format_func=fmt, key=key)


def build_ics_calendar(sesiones_df: pd.DataFrame, datos: Dict[str, str]) -> bytes:
    asignatura = datos.get("asignatura", "Clase")
    grupo = datos.get("grupo", "")
    lugar_default = datos.get("ambiente", "")
    stamp = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Gestor FDGC//Calendario Academico//ES",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
    ]
    for i, row in sesiones_df.reset_index(drop=True).iterrows():
        fecha = pd.to_datetime(row.get("Fecha", ""), errors="coerce")
        if pd.isna(fecha):
            continue
        horario = str(row.get("Horario", "")).strip()
        m = re.search(r"(\d{1,2}:\d{2})\s*[-–]\s*(\d{1,2}:\d{2})", horario)
        inicio = "08:00" if not m else m.group(1)
        fin = "10:00" if not m else m.group(2)
        dtstart = datetime.combine(fecha.date(), parse_time_value(inicio) or time(8,0)).strftime("%Y%m%dT%H%M%S")
        dtend = datetime.combine(fecha.date(), parse_time_value(fin) or time(10,0)).strftime("%Y%m%dT%H%M%S")
        uid = f"fdgc-{datos.get('codigo','curso')}-{datos.get('grupo','grupo')}-{i+1}@local"
        summary = f"{asignatura} {grupo} - Sesión {row.get('N° sesión', i+1)}"
        desc = f"Unidad: {row.get('Unidad','')}\\nContenido: {row.get('Contenido por desarrollar','')}\\nTrabajo presencial: {row.get('Descripción del trabajo presencial','')}\\nTrabajo independiente: {row.get('Descripción trabajo independiente','')}"
        location = str(row.get("Lugar / ambiente", "") or lugar_default)
        lines += [
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"DTSTAMP:{stamp}",
            f"DTSTART:{dtstart}",
            f"DTEND:{dtend}",
            f"SUMMARY:{summary}",
            f"DESCRIPTION:{desc}",
            f"LOCATION:{location}",
            "END:VEVENT",
        ]
    lines.append("END:VCALENDAR")
    return "\r\n".join(lines).encode("utf-8")


def validar_plan(modulos_df: pd.DataFrame, horarios_df: pd.DataFrame, sesiones_df: pd.DataFrame, evaluaciones_df: pd.DataFrame, datos: Dict[str, str]) -> Tuple[List[str], List[str]]:
    errores, alertas = [], []
    if not datos.get("asignatura"):
        errores.append("Falta la asignatura.")
    if not datos.get("grupo"):
        alertas.append("No se indicó grupo.")
    if limpiar_df(modulos_df, COLUMNAS_MODULOS).empty:
        errores.append("Debe registrar al menos una unidad/módulo.")
    if limpiar_df(horarios_df, COLUMNAS_HORARIOS).empty:
        errores.append("Debe registrar al menos un día y horario de clase.")
    ev = limpiar_df(evaluaciones_df, COLUMNAS_EVALUACIONES)
    total_eval = sum((limpiar_numero(v) or 0) for v in ev.get("Valor (%)", [])) if not ev.empty else 0
    if ev.empty:
        alertas.append("No hay evaluación concertada. El FD-GC71 saldrá sin plan evaluativo.")
    elif round(total_eval, 2) != 100:
        errores.append(f"La evaluación concertada suma {total_eval:.2f}%, debe sumar 100%.")
    if sesiones_df is not None and not sesiones_df.empty:
        fechas = pd.to_datetime(sesiones_df["Fecha"], errors="coerce")
        if fechas.isna().any():
            alertas.append("Hay sesiones sin fecha válida.")
        repetidas = sesiones_df.groupby(["Fecha", "Horario"]).size()
        if (repetidas > 1).any():
            alertas.append("Existen sesiones con la misma fecha y horario; revise si es intencional.")
    htp = limpiar_numero(datos.get("htp", 0)) or 0
    total_horas_mod = sum((limpiar_numero(v) or 0) for v in modulos_df.get("Horas presenciales", [])) if modulos_df is not None and not modulos_df.empty else 0
    if htp and total_horas_mod and abs(total_horas_mod - htp) > 2:
        alertas.append(f"Las horas presenciales por módulo suman {total_horas_mod:g}, pero la identificación reporta HTP={htp:g}.")
    return errores, alertas


def crear_paquete_curso_zip(datos: Dict[str, str], sesiones_df: pd.DataFrame, evaluaciones_df: pd.DataFrame, estudiantes_df: pd.DataFrame, representantes_df: pd.DataFrame, curso_id: Optional[int] = None) -> bytes:
    gc71 = crear_gc71_docx(datos, sesiones_df, evaluaciones_df, representantes_df)
    excel = crear_plantilla_evaluacion_xlsx(estudiantes_df, evaluaciones_df, datos)
    ics = build_ics_calendar(sesiones_df, datos)
    ses_csv = sesiones_df.to_csv(index=False).encode("utf-8-sig")
    ev_csv = evaluaciones_df.to_csv(index=False).encode("utf-8-sig")
    payload = {
        "datos": datos,
        "sesiones": df_to_payload(sesiones_df),
        "evaluaciones": df_to_payload(evaluaciones_df),
        "estudiantes": df_to_payload(estudiantes_df),
        "representantes": df_to_payload(representantes_df),
        "generado_en": ahora_iso(),
        "curso_id": curso_id,
    }
    buf = io.BytesIO()
    nombre_base = nombre_archivo_seguro(datos.get("asignatura", "curso"), date.today(), "FDGC_Paquete")
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(f"{nombre_base}/01_FD-GC71_Guia_Didactica.docx", gc71)
        z.writestr(f"{nombre_base}/02_Plantilla_Evaluacion.xlsx", excel)
        z.writestr(f"{nombre_base}/03_Calendario_Clases.ics", ics)
        z.writestr(f"{nombre_base}/04_Sesiones.csv", ses_csv)
        z.writestr(f"{nombre_base}/05_Evaluaciones.csv", ev_csv)
        z.writestr(f"{nombre_base}/06_Payload_Reutilizable.json", json.dumps(payload, ensure_ascii=False, indent=2, default=str).encode("utf-8"))
        z.writestr(f"{nombre_base}/LEAME.txt", "Paquete generado desde Gestor FD-GC71/FD-GC72 Enterprise. Incluye guía, plantilla de notas, calendario, sesiones, evaluación y payload para trazabilidad.\n")
    return buf.getvalue()


def ui_inicio(st):
    user = st.session_state.get("auth_user", {})
    st.header("Panel de inicio")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Usuario", user.get("usuario", ""))
    c2.metric("Perfil", user.get("rol", ""))
    c3.metric("Estado", "Activo")
    cursos = listar_cursos_visibles()
    c4.metric("Cursos visibles", 0 if cursos.empty else len(cursos))
    st.info(ROLES_PERMISOS.get(user.get("rol"), {}).get("descripcion", ""))
    if get_app_env() == "production" and not usar_postgres():
        st.error("APP_ENV está en production pero no hay DATABASE_URL. Configure PostgreSQL antes de operar con datos reales.")
    elif usar_postgres():
        st.success("Persistencia productiva activa: PostgreSQL externo configurado.")
    if user.get("debe_cambiar_clave"):
        st.warning("Debe cambiar la contraseña inicial desde Mi cuenta antes de operar formalmente.")
    st.subheader("Ruta institucional recomendada")
    st.markdown(
        """
1. **Expediente académico:** registre el curso y deje trazabilidad de programa, grupo, periodo, docente y estado.
2. **Planeador superior:** defina unidades, intensidad, horario y evaluación; genere FD-GC71, calendario `.ics` y plantilla de notas en un solo paquete.
3. **Durante el curso:** cargue evidencias, use la plantilla de evaluación y conserve soportes por corte.
4. **FD-GC72:** al corte parcial y final, cargue listado/notas y genere el informe académico con análisis descriptivo.
5. **Validador institucional:** revise que evaluación sume 100%, que el cronograma sea consistente y que los soportes estén completos.
        """
    )
    st.caption("La idea es simple: menos copiar y pegar, más control. El sufrimiento académico no debe ser un requisito de grado.")


def ui_centro_control(st):
    st.header("Centro de control académico")
    df = listar_cursos_visibles()
    evid = read_sql_df("SELECT * FROM evidencias")
    art = read_sql_df("SELECT * FROM artefactos")
    aud = read_sql_df("SELECT * FROM auditoria ORDER BY id DESC LIMIT 200")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Cursos", len(df))
    c2.metric("Evidencias", len(evid))
    c3.metric("Artefactos", len(art))
    c4.metric("Auditoría", len(aud))
    if df.empty:
        st.info("Aún no hay cursos registrados.")
        return
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Cursos por estado")
        estado = df.groupby("estado", dropna=False).size().reset_index(name="Cursos")
        st.bar_chart(estado.set_index("estado"))
    with col2:
        st.subheader("Avance promedio")
        avance = df[["asignatura", "avance_contenido", "avance_evaluado"]].copy()
        avance["Curso"] = avance["asignatura"].str.slice(0, 28)
        st.dataframe(avance[["Curso", "avance_contenido", "avance_evaluado"]], use_container_width=True, hide_index=True)
    st.subheader("Semáforo académico")
    sem = df.copy()
    sem["Riesgo"] = sem.apply(lambda r: "Alto" if (float(r.get("avance_contenido") or 0) < 50 and str(r.get("estado")) not in ["Planeación"]) else ("Medio" if float(r.get("avance_evaluado") or 0) < 30 and str(r.get("estado")) in ["En ejecución", "Corte parcial"] else "Bajo"), axis=1)
    st.dataframe(sem[["id", "codigo", "grupo", "asignatura", "periodo", "profesor", "estado", "avance_contenido", "avance_evaluado", "Riesgo"]], use_container_width=True, hide_index=True)


def ui_expediente_academico(st):
    st.header("Expediente académico del curso")
    st.caption("Registro maestro para no depender de archivos sueltos con nombres tipo FINAL_final_ahora_si_v8.docx. Todos hemos estado ahí.")
    df = listar_cursos_visibles()
    tab_lista, tab_crear, tab_editar = st.tabs(["Listado", "Crear expediente", "Editar / cerrar"])
    with tab_lista:
        if df.empty:
            st.info("No hay expedientes registrados.")
        else:
            filtro = st.text_input("Buscar por asignatura, código, grupo, docente o periodo")
            view = df.copy()
            if filtro:
                f = normalizar_texto(filtro)
                mask = view.apply(lambda row: f in normalizar_texto(" ".join(str(x) for x in row.values)), axis=1)
                view = view[mask]
            st.dataframe(view[["id", "codigo", "grupo", "asignatura", "programa", "periodo", "profesor", "estado", "avance_contenido", "avance_evaluado", "actualizado_en"]], use_container_width=True, hide_index=True)
            st.download_button("Descargar inventario CSV", view.to_csv(index=False).encode("utf-8-sig"), "inventario_expedientes.csv", "text/csv", use_container_width=True)
    with tab_crear:
        with st.form("crear_expediente"):
            c1, c2, c3 = st.columns(3)
            with c1:
                codigo = st.text_input("Código")
                grupo = st.text_input("Grupo")
                asignatura = st.text_input("Asignatura")
                programa = st.text_input("Programa académico")
            with c2:
                periodo = st.text_input("Período académico")
                profesor = st.text_input("Profesor", value=st.session_state.get("auth_user", {}).get("nombre_completo", ""))
                correo = st.text_input("Correo docente", value=st.session_state.get("auth_user", {}).get("email", ""))
                creditos = st.text_input("Número de créditos")
            with c3:
                htp = st.number_input("HTP semanal / referencia", min_value=0.0, step=0.5)
                hti = st.number_input("HTI semanal / referencia", min_value=0.0, step=0.5)
                fecha_inicio = st.date_input("Fecha inicio", value=date.today())
                fecha_fin = st.date_input("Fecha fin", value=date.today() + timedelta(days=110))
                estado = st.selectbox("Estado", ["Planeación", "En ejecución", "Corte parcial", "Cierre final", "Archivado"])
            enviar = st.form_submit_button("Crear expediente", use_container_width=True)
        if enviar:
            datos = dict(codigo=codigo, grupo=grupo, asignatura=asignatura, programa=programa, periodo=periodo, profesor=profesor, correo=correo, creditos=creditos, htp=htp, hti=hti, fecha_inicio=fecha_inicio, fecha_fin=fecha_fin, estado=estado)
            new_id = upsert_curso(None, datos, payload={"origen": "Expediente académico"})
            st.success(f"Expediente creado: #{new_id}")
            st.rerun()
    with tab_editar:
        curso_id = seleccionar_curso_widget("Seleccione expediente", key="edit_expediente")
        if curso_id:
            curso = get_curso(int(curso_id))
            with st.form("editar_expediente"):
                c1, c2, c3 = st.columns(3)
                with c1:
                    codigo = st.text_input("Código", value=curso.get("codigo") or "")
                    grupo = st.text_input("Grupo", value=curso.get("grupo") or "")
                    asignatura = st.text_input("Asignatura", value=curso.get("asignatura") or "")
                    programa = st.text_input("Programa", value=curso.get("programa") or "")
                with c2:
                    periodo = st.text_input("Periodo", value=curso.get("periodo") or "")
                    profesor = st.text_input("Profesor", value=curso.get("profesor") or "")
                    correo = st.text_input("Correo", value=curso.get("email_profesor") or "")
                    creditos = st.text_input("Créditos", value=curso.get("creditos") or "")
                with c3:
                    estado = st.selectbox("Estado", ["Planeación", "En ejecución", "Corte parcial", "Cierre final", "Archivado"], index=max(0, ["Planeación", "En ejecución", "Corte parcial", "Cierre final", "Archivado"].index(curso.get("estado") or "Planeación") if (curso.get("estado") or "Planeación") in ["Planeación", "En ejecución", "Corte parcial", "Cierre final", "Archivado"] else 0))
                    avance_contenido = st.slider("% avance contenido", 0, 100, int(float(curso.get("avance_contenido") or 0)))
                    avance_evaluado = st.slider("% evaluado", 0, 100, int(float(curso.get("avance_evaluado") or 0)))
                    fecha_inicio = st.text_input("Fecha inicio", value=curso.get("fecha_inicio") or "")
                    fecha_fin = st.text_input("Fecha fin", value=curso.get("fecha_fin") or "")
                guardar = st.form_submit_button("Guardar cambios", use_container_width=True)
            if guardar:
                payload = safe_json_loads(curso.get("payload_json"), {})
                datos = dict(codigo=codigo, grupo=grupo, asignatura=asignatura, programa=programa, periodo=periodo, profesor=profesor, correo=correo, creditos=creditos, fecha_inicio=fecha_inicio, fecha_fin=fecha_fin, estado=estado, avance_contenido=avance_contenido, avance_evaluado=avance_evaluado)
                upsert_curso(int(curso_id), datos, payload=payload)
                st.success("Expediente actualizado.")
                st.rerun()
            if st.session_state.get("auth_user", {}).get("rol") == "Administrador":
                with st.expander("Zona peligrosa"):
                    st.warning("Eliminar borra el expediente. Las evidencias físicas quedan en carpeta, pero sin vínculo visible.")
                    if st.button("Eliminar expediente", type="secondary"):
                        eliminar_curso(int(curso_id))
                        st.success("Expediente eliminado.")
                        st.rerun()


def ui_planeador_superior(st):
    st.header("Planeador superior FD-GC71")
    st.caption("Genera cronograma por horario real, valida evaluación, produce FD-GC71, Excel de notas, calendario ICS, CSV y paquete trazable.")
    curso_id = seleccionar_curso_widget("Usar expediente existente o crear desde cero", key="curso_planeador", incluir_nuevo=True)
    curso = get_curso(int(curso_id)) if curso_id else {}
    saved_payload = safe_json_loads(curso.get("payload_json") if curso else None, {})

    tab_datos, tab_plan, tab_generar = st.tabs(["1. Datos base", "2. Planeación y evaluación", "3. Validar y descargar"])
    with tab_datos:
        c1, c2, c3 = st.columns(3)
        with c1:
            programa = st.text_input("Programa académico", value=curso.get("programa", saved_payload.get("datos", {}).get("programa", "")) if curso else "")
            asignatura = st.text_input("Asignatura", value=curso.get("asignatura", saved_payload.get("datos", {}).get("asignatura", "")) if curso else "")
            codigo = st.text_input("Código", value=curso.get("codigo", saved_payload.get("datos", {}).get("codigo", "")) if curso else "")
            grupo = st.text_input("Grupo", value=curso.get("grupo", saved_payload.get("datos", {}).get("grupo", "")) if curso else "")
        with c2:
            profesor = st.text_input("Profesor", value=(curso.get("profesor") if curso else st.session_state.get("auth_user", {}).get("nombre_completo", "")) or "")
            correo = st.text_input("Correo", value=(curso.get("email_profesor") if curso else st.session_state.get("auth_user", {}).get("email", "")) or "")
            periodo = st.text_input("Período académico", value=curso.get("periodo", saved_payload.get("datos", {}).get("periodo", "")) if curso else "")
            creditos = st.text_input("Número de créditos", value=str(curso.get("creditos", "")) if curso else "")
        with c3:
            fecha_inicio = st.date_input("Inicio de clases", value=pd.to_datetime(curso.get("fecha_inicio") or date.today()).date() if curso and curso.get("fecha_inicio") else date.today())
            fecha_fin = st.date_input("Fin de clases", value=pd.to_datetime(curso.get("fecha_fin") or (date.today()+timedelta(days=110))).date() if curso and curso.get("fecha_fin") else date.today()+timedelta(days=110))
            htp = st.number_input("HTP total/referencia", min_value=0.0, step=0.5, value=float(curso.get("htp") or 0) if curso else 0.0)
            hti = st.number_input("HTI total/referencia", min_value=0.0, step=0.5, value=float(curso.get("hti") or 0) if curso else 0.0)
            estado = st.selectbox("Estado", ["Planeación", "En ejecución", "Corte parcial", "Cierre final", "Archivado"], index=0)
        st.subheader("Textos académicos con preforma editable")
        textos = {}
        for k, label in [
            ("justificacion", "Justificación"), ("competencias", "Competencias"), ("resultados", "Resultados de aprendizaje"),
            ("objetivo_general", "Objetivo general"), ("objetivos_especificos", "Objetivos específicos"), ("metodologias", "Metodologías y estrategias didácticas"),
            ("ambientes", "Ambientes de aprendizaje"), ("medios", "Medios educativos"), ("referencias", "Referencias bibliográficas"),
        ]:
            base = saved_payload.get("datos", {}).get(k, TEXTOS_PREDEFINIDOS_GC71.get(k, ""))
            textos[k] = st.text_area(label, value=base, height=90 if k not in ("objetivos_especificos", "referencias") else 120)

    with tab_plan:
        st.subheader("Unidades / módulos e intensidad")
        mod_default = payload_to_df(saved_payload.get("modulos"), COLUMNAS_MODULOS)
        if mod_default.empty:
            mod_default = pd.DataFrame([
                {"Unidad": "UNIDAD 1", "Contenido / tema central": "Fundamentos y contexto de la asignatura", "Horas presenciales": 8, "Sesiones": 4, "Trabajo presencial": "Clase orientadora, discusión guiada y taller aplicado.", "Trabajo independiente": "Lectura previa y desarrollo de actividad de preparación."},
                {"Unidad": "UNIDAD 2", "Contenido / tema central": "Métodos, herramientas y procedimientos", "Horas presenciales": 12, "Sesiones": 6, "Trabajo presencial": "Ejercicios prácticos y análisis de casos.", "Trabajo independiente": "Desarrollo de ejercicios y consulta bibliográfica."},
                {"Unidad": "UNIDAD 3", "Contenido / tema central": "Aplicación, integración y cierre", "Horas presenciales": 12, "Sesiones": 6, "Trabajo presencial": "Proyecto integrador y socialización de resultados.", "Trabajo independiente": "Preparación de entregables y retroalimentación final."},
            ])
        modulos_df = st.data_editor(mod_default, num_rows="dynamic", use_container_width=True, key="super_modulos")
        st.subheader("Horario de clase")
        hor_default = payload_to_df(saved_payload.get("horarios"), COLUMNAS_HORARIOS)
        if hor_default.empty:
            hor_default = pd.DataFrame([{ "Día": "Lunes", "Hora inicio": "18:00", "Hora fin": "20:00", "Lugar / ambiente": "Aula / plataforma institucional" }])
        horarios_df = st.data_editor(hor_default, num_rows="dynamic", use_container_width=True, key="super_horarios")
        criterio = st.radio("Criterio para repartir sesiones", ["Horas presenciales", "Sesiones"], horizontal=True)
        excluir = st.text_area("Fechas sin clase a excluir (una por línea, formato AAAA-MM-DD)", value="")
        fechas_excluidas = []
        for line in excluir.splitlines():
            try:
                fechas_excluidas.append(pd.to_datetime(line.strip()).date())
            except Exception:
                pass
        fechas_clase_df = generar_fechas_clase(fecha_inicio, fecha_fin, horarios_df, fechas_excluidas)
        sesiones_df = expandir_plan_sesiones(modulos_df, fechas_clase_df, criterio=criterio)
        st.subheader("Sesiones generadas")
        sesiones_edit = st.data_editor(sesiones_df, num_rows="dynamic", use_container_width=True, key="super_sesiones")
        st.subheader("Evaluación concertada")
        ev_default = payload_to_df(saved_payload.get("evaluaciones"), COLUMNAS_EVALUACIONES)
        if ev_default.empty:
            ev_default = pd.DataFrame([
                {"Tipo de evaluación": "Seguimiento", "Procedimiento de evaluación": "Talleres, actividades de clase y ejercicios aplicados", "Valor (%)": 30, "Fecha de realización": str(fecha_inicio + timedelta(days=30)), "Unidad relacionada": "UNIDAD 1", "Corte": "Primer corte"},
                {"Tipo de evaluación": "Parcial", "Procedimiento de evaluación": "Evaluación individual teórico-práctica", "Valor (%)": 30, "Fecha de realización": str(fecha_inicio + timedelta(days=60)), "Unidad relacionada": "UNIDAD 2", "Corte": "Segundo corte"},
                {"Tipo de evaluación": "Proyecto final", "Procedimiento de evaluación": "Entrega y sustentación de proyecto integrador", "Valor (%)": 40, "Fecha de realización": str(fecha_fin), "Unidad relacionada": "UNIDAD 3", "Corte": "Final"},
            ])
        evaluaciones_df = st.data_editor(ev_default, num_rows="dynamic", use_container_width=True, key="super_evaluaciones")
        st.subheader("Estudiantes / listado tradicional")
        uploaded_list = st.file_uploader("Cargar listado tradicional (.xls/.xlsx) opcional", type=["xls", "xlsx"], key="super_listado")
        if uploaded_list:
            estudiantes_df = leer_listado_estudiantes(uploaded_list)
            st.success(f"Listado leído: {len(estudiantes_df)} estudiantes.")
        else:
            estudiantes_df = payload_to_df(saved_payload.get("estudiantes"), ["Nombre completo", "Documento", "Correo", "Plan", "Observación", "Estado"])
            if estudiantes_df.empty:
                estudiantes_df = df_vacio(["Nombre completo", "Documento", "Correo", "Plan", "Observación", "Estado"], 5)
            estudiantes_df = st.data_editor(estudiantes_df, num_rows="dynamic", use_container_width=True, key="super_estudiantes")
        st.session_state["super_planeacion"] = dict(modulos=modulos_df, horarios=horarios_df, sesiones=sesiones_edit, evaluaciones=evaluaciones_df, estudiantes=estudiantes_df)

    with tab_generar:
        planeacion = st.session_state.get("super_planeacion", {})
        modulos_df = planeacion.get("modulos", pd.DataFrame(columns=COLUMNAS_MODULOS))
        horarios_df = planeacion.get("horarios", pd.DataFrame(columns=COLUMNAS_HORARIOS))
        sesiones_df = planeacion.get("sesiones", pd.DataFrame(columns=COLUMNAS_SESIONES))
        evaluaciones_df = planeacion.get("evaluaciones", pd.DataFrame(columns=COLUMNAS_EVALUACIONES))
        estudiantes_df = planeacion.get("estudiantes", pd.DataFrame())
        representantes_df = pd.DataFrame(columns=["Nombre de los estudiantes", "N° de cédula o carné estudiantil", "Firma"])
        datos = dict(programa=programa, asignatura=asignatura, codigo=codigo, grupo=grupo, profesor=profesor, correo=correo, periodo=periodo, creditos=creditos, htp=htp, hti=hti, fecha_inicio=fecha_inicio, fecha_fin=fecha_fin, estado=estado, **textos)
        errores, alertas = validar_plan(modulos_df, horarios_df, sesiones_df, evaluaciones_df, datos)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Sesiones", 0 if sesiones_df is None else len(limpiar_df(sesiones_df, COLUMNAS_SESIONES)))
        c2.metric("Evaluación", f"{sum((limpiar_numero(v) or 0) for v in evaluaciones_df.get('Valor (%)', [])) if evaluaciones_df is not None and not evaluaciones_df.empty else 0:.0f}%")
        c3.metric("Estudiantes", 0 if estudiantes_df is None else len(limpiar_df(estudiantes_df, list(estudiantes_df.columns))))
        c4.metric("Errores", len(errores))
        if errores:
            st.error("Corrija antes de generar: " + " | ".join(errores))
        if alertas:
            st.warning("Alertas: " + " | ".join(alertas))
        if not errores:
            st.success("Plan listo para generar y guardar.")
            col_a, col_b, col_c, col_d = st.columns(4)
            gc71_bytes = crear_gc71_docx(datos, sesiones_df, evaluaciones_df, representantes_df)
            excel_bytes = crear_plantilla_evaluacion_xlsx(estudiantes_df, evaluaciones_df, datos)
            ics_bytes = build_ics_calendar(sesiones_df, datos)
            paquete_bytes = crear_paquete_curso_zip(datos, sesiones_df, evaluaciones_df, estudiantes_df, representantes_df, curso_id)
            base = nombre_archivo_seguro(asignatura, date.today(), "FDGC")
            col_a.download_button("FD-GC71 Word", gc71_bytes, f"{base}_FD-GC71.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
            col_b.download_button("Plantilla Excel", excel_bytes, f"{base}_Evaluacion.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            col_c.download_button("Calendario ICS", ics_bytes, f"{base}_Calendario.ics", "text/calendar", use_container_width=True)
            col_d.download_button("Paquete completo ZIP", paquete_bytes, f"{base}_Paquete.zip", "application/zip", use_container_width=True)
            if st.button("Guardar/actualizar expediente con esta planeación", use_container_width=True, type="primary"):
                payload = {
                    "datos": {k: str(v) for k, v in datos.items()},
                    "modulos": df_to_payload(modulos_df),
                    "horarios": df_to_payload(horarios_df),
                    "sesiones": df_to_payload(sesiones_df),
                    "evaluaciones": df_to_payload(evaluaciones_df),
                    "estudiantes": df_to_payload(estudiantes_df),
                    "guardado_en": ahora_iso(),
                }
                new_id = upsert_curso(int(curso_id) if curso_id else None, datos, payload=payload)
                registrar_artefacto(new_id, "FD-GC71", f"{base}_FD-GC71.docx", "Generado desde Planeador superior")
                registrar_artefacto(new_id, "Excel evaluación", f"{base}_Evaluacion.xlsx", "Generado desde Planeador superior")
                registrar_auditoria("Guardar planeación superior", f"Curso ID={new_id}")
                st.success(f"Planeación guardada en expediente #{new_id}.")
                st.rerun()


def ui_evidencias(st):
    st.header("Evidencias y soportes")
    curso_id = seleccionar_curso_widget("Curso", key="curso_evidencias")
    if not curso_id:
        return
    curso = get_curso(int(curso_id))
    st.write(f"**Curso:** {curso.get('asignatura')} | **Grupo:** {curso.get('grupo')} | **Periodo:** {curso.get('periodo')}")
    tab_subir, tab_listar = st.tabs(["Subir soporte", "Repositorio"])
    with tab_subir:
        tipo = st.selectbox("Tipo de evidencia", ["Listado de clase", "Evaluación parcial", "Evaluación final", "Acta de concertación", "Asistencia", "Soporte metodológico", "Otro"])
        descripcion = st.text_area("Descripción / observación")
        files = st.file_uploader("Archivos", accept_multiple_files=True, type=["pdf", "doc", "docx", "xls", "xlsx", "csv", "png", "jpg", "jpeg", "zip"])
        if st.button("Guardar evidencias", type="primary", use_container_width=True):
            if not files:
                st.error("Seleccione al menos un archivo.")
            else:
                conn = conexion_db()
                user = st.session_state.get("auth_user", {})
                curso_dir = EVIDENCE_DIR / str(curso_id)
                curso_dir.mkdir(parents=True, exist_ok=True)
                for f in files:
                    raw = f.getvalue()
                    if len(raw) > max_evidence_bytes():
                        st.error(f"El archivo {f.name} pesa {len(raw)/1024/1024:.1f} MB y supera el límite configurado de {max_evidence_bytes()/1024/1024:.0f} MB.")
                        continue
                    ext = Path(f.name).suffix.lower()
                    internal = f"{uuid.uuid4().hex}{ext}"
                    # Caché local para descarga rápida. En Streamlit Cloud no se asume persistente.
                    (curso_dir / internal).write_bytes(raw)
                    contenido_b64 = base64.b64encode(raw).decode("ascii")
                    conn.execute(
                        "INSERT INTO evidencias(curso_id, tipo, nombre_original, nombre_archivo, mime, tamano, descripcion, subido_por, subido_en, contenido_b64) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        (int(curso_id), tipo, f.name, internal, getattr(f, "type", ""), len(raw), descripcion, user.get("usuario", ""), ahora_iso(), contenido_b64),
                    )
                conn.commit(); conn.close()
                registrar_auditoria("Subir evidencia", f"Curso ID={curso_id} | {len(files)} archivo(s) | Tipo={tipo}")
                st.success("Evidencias guardadas.")
                st.rerun()
    with tab_listar:
        df = read_sql_df("SELECT id, tipo, nombre_original, nombre_archivo, tamano, descripcion, subido_por, subido_en, contenido_b64 FROM evidencias WHERE curso_id=? ORDER BY id DESC", params=(int(curso_id),))
        if df.empty:
            st.info("Este curso aún no tiene evidencias.")
        else:
            cols_hide = [c for c in ["nombre_archivo", "contenido_b64"] if c in df.columns]
            st.dataframe(df.drop(columns=cols_hide), use_container_width=True, hide_index=True)
            for _, r in df.iterrows():
                path = EVIDENCE_DIR / str(curso_id) / str(r["nombre_archivo"])
                data = None
                if path.exists():
                    data = path.read_bytes()
                elif str(r.get("contenido_b64", "") or "").strip():
                    try:
                        data = base64.b64decode(str(r.get("contenido_b64")))
                    except Exception:
                        data = None
                if data is not None:
                    with st.expander(f"Descargar #{r['id']} - {r['nombre_original']}"):
                        st.write(r.get("descripcion", ""))
                        st.download_button("Descargar archivo", data, r["nombre_original"], use_container_width=True, key=f"download_evid_{r['id']}")
                else:
                    st.warning(f"No se encontró el contenido de la evidencia #{r['id']}: {r['nombre_original']}")


def ui_validador_institucional(st):
    st.header("Validador institucional")
    st.caption("Revisa coherencia básica antes de entregar: evaluación, listado, calificaciones y estructura. No reemplaza al comité; evita que el comité te mire feo.")
    tab_excel, tab_expediente = st.tabs(["Validar Excel de evaluación", "Validar expediente"])
    with tab_excel:
        f = st.file_uploader("Cargue plantilla de evaluación diligenciada", type=["xlsx", "xls"], key="validar_excel")
        corte = st.number_input("Nota mínima aprobatoria", min_value=0.0, max_value=5.0, value=3.0, step=0.1)
        if f:
            try:
                sheets = pd.read_excel(f, sheet_name=None)
                hallazgos = []
                if "Evaluaciones" not in sheets:
                    hallazgos.append(("Error", "No existe hoja Evaluaciones."))
                else:
                    ev = sheets["Evaluaciones"]
                    col_valor = encontrar_columna(ev, ["valor", "%"])
                    if col_valor:
                        total = pd.to_numeric(ev[col_valor], errors="coerce").fillna(0).sum()
                        if abs(total - 100) > 0.01:
                            hallazgos.append(("Error", f"La evaluación suma {total:.2f}%, debe sumar 100%."))
                        else:
                            hallazgos.append(("OK", "La evaluación suma 100%."))
                    else:
                        hallazgos.append(("Alerta", "No se identificó columna de valor porcentual."))
                if "Calificaciones" not in sheets:
                    hallazgos.append(("Error", "No existe hoja Calificaciones."))
                else:
                    cal = sheets["Calificaciones"]
                    if len(cal) == 0:
                        hallazgos.append(("Alerta", "La hoja Calificaciones no tiene estudiantes."))
                    notas_cols = [c for c in cal.columns if normalizar_texto(c).startswith("E") or "NOTA" in normalizar_texto(c)]
                    hallazgos.append(("Info", f"Filas de calificaciones: {len(cal)} | columnas de notas detectadas: {len(notas_cols)}."))
                    # Conteos aproximados
                    notas_num = pd.to_numeric(cal[notas_cols[-1]], errors="coerce") if notas_cols else pd.Series(dtype=float)
                    if not notas_num.empty:
                        hallazgos.append(("Info", f"Aprobados aproximados: {(notas_num >= corte).sum()} | Reprobados aproximados: {(notas_num < corte).sum()}."))
                st.dataframe(pd.DataFrame(hallazgos, columns=["Nivel", "Hallazgo"]), use_container_width=True, hide_index=True)
            except Exception as e:
                st.error(f"No se pudo validar: {e}")
    with tab_expediente:
        curso_id = seleccionar_curso_widget("Curso a validar", key="validar_curso")
        if curso_id:
            curso = get_curso(int(curso_id))
            payload = safe_json_loads(curso.get("payload_json"), {})
            modulos = payload_to_df(payload.get("modulos"), COLUMNAS_MODULOS)
            horarios = payload_to_df(payload.get("horarios"), COLUMNAS_HORARIOS)
            sesiones = payload_to_df(payload.get("sesiones"), COLUMNAS_SESIONES)
            evaluaciones = payload_to_df(payload.get("evaluaciones"), COLUMNAS_EVALUACIONES)
            datos = payload.get("datos", dict(curso))
            errores, alertas = validar_plan(modulos, horarios, sesiones, evaluaciones, datos)
            evid_row = db_execute("SELECT COUNT(*) AS n FROM evidencias WHERE curso_id=?", (int(curso_id),), fetchone=True)
            evid_count = evid_row["n"] if evid_row is not None else 0
            hallazgos = [("Error", e) for e in errores] + [("Alerta", a) for a in alertas]
            hallazgos.append(("Info", f"Evidencias asociadas: {evid_count}."))
            if not errores:
                hallazgos.append(("OK", "El expediente pasa las validaciones estructurales básicas."))
            st.dataframe(pd.DataFrame(hallazgos, columns=["Nivel", "Hallazgo"]), use_container_width=True, hide_index=True)


def ui_backup(st):
    st.header("Copias y restauración")
    st.caption("Respaldo portable: exporta tablas, evidencias y manifiesto. En producción esto no se improvisa el viernes a las 5:59 p. m.")
    if st.session_state.get("auth_user", {}).get("rol") != "Administrador":
        st.warning("Solo el perfil Administrador puede generar respaldos completos.")
        return

    def build_backup() -> bytes:
        buf = io.BytesIO()
        manifest = {
            "generado_en": ahora_iso(),
            "database": "PostgreSQL" if usar_postgres() else "SQLite",
            "version": APP_VERSION,
            "incluye": ["tablas_csv", "manifiesto", "evidencias_db", "cache_local_si_existe"],
        }
        tablas = ["usuarios", "auditoria", "cursos", "evidencias", "artefactos"]
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for tabla in tablas:
                try:
                    df = read_sql_df(f"SELECT * FROM {tabla}")
                    z.writestr(f"tablas/{tabla}.csv", df.to_csv(index=False).encode("utf-8-sig"))
                except Exception as exc:
                    z.writestr(f"tablas/{tabla}_ERROR.txt", str(exc))
            # Archivo SQLite completo si aplica.
            if DB_PATH.exists() and not usar_postgres():
                z.write(DB_PATH, f"app_data/{DB_PATH.name}")
            # Evidencias desde DB para que funcione en Streamlit Cloud.
            try:
                evid = read_sql_df("SELECT id, curso_id, nombre_original, nombre_archivo, contenido_b64 FROM evidencias")
                for _, r in evid.iterrows():
                    data = None
                    if str(r.get("contenido_b64", "") or "").strip():
                        data = base64.b64decode(str(r.get("contenido_b64")))
                    else:
                        path = EVIDENCE_DIR / str(r.get("curso_id")) / str(r.get("nombre_archivo"))
                        if path.exists():
                            data = path.read_bytes()
                    if data is not None:
                        safe_name = re.sub(r"[^A-Za-z0-9_. -]", "_", str(r.get("nombre_original", "evidencia")))
                        z.writestr(f"evidencias/curso_{r.get('curso_id')}/id_{r.get('id')}_{safe_name}", data)
            except Exception as exc:
                z.writestr("evidencias_ERROR.txt", str(exc))
            # Cache local de exportaciones si existe.
            for folder in [EXPORT_DIR]:
                if folder.exists():
                    for file in folder.rglob("*"):
                        if file.is_file():
                            z.write(file, str(file.relative_to(APP_DIR)))
            z.writestr("manifest_backup.json", json.dumps(manifest, ensure_ascii=False, indent=2))
        return buf.getvalue()

    backup_bytes = build_backup()
    st.download_button("Descargar respaldo completo", backup_bytes, f"backup_fdgc_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip", "application/zip", use_container_width=True)
    st.divider()
    st.subheader("Estado técnico")
    st.json(health_status())
    st.write(f"Tamaño del respaldo actual: {len(backup_bytes)/1024:.1f} KB")


def ui_ayuda(st):
    st.header("Ayuda / flujo recomendado")
    st.markdown(
        """
### Flujo de operación recomendado

**Inicio de semestre**
- Crear expediente del curso.
- Definir unidades, intensidad, horario y evaluación concertada en Planeador superior.
- Descargar paquete completo: FD-GC71, Excel de evaluación, calendario ICS, CSV y JSON de trazabilidad.

**Mitad de curso**
- Cargar la plantilla de evaluación diligenciada.
- Usar FD-GC72 para calcular aprobados, reprobados, avance y análisis descriptivo.
- Guardar soportes en Evidencias.

**Final de curso**
- Actualizar evaluación final.
- Generar FD-GC72 definitivo.
- Validar expediente y descargar respaldo.

### Reglas de calidad
- La evaluación debe sumar exactamente 100%.
- Cada sesión debe tener fecha, contenido, trabajo presencial e independiente.
- La plantilla de evaluación debe conservar las hojas generadas por la app.
- El listado tradicional puede cargarse al inicio, mitad o final para alimentar los cálculos.
        """
    )


def ui_diagnostico_productivo(st):
    st.header("Diagnóstico productivo")
    st.caption("Lectura rápida del estado técnico del despliegue.")
    status = health_status()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Versión", status.get("version", ""))
    c2.metric("Ambiente", status.get("app_env", ""))
    c3.metric("Base", status.get("database", ""))
    c4.metric("Plantillas", "OK" if status.get("templates_ok") else "Revisar")
    if get_app_env() == "production" and not usar_postgres():
        st.error("Modo producción sin PostgreSQL. Esto sirve para demo, no para operación institucional.")
    elif usar_postgres() and status.get("database_ok"):
        st.success("Base de datos externa operativa.")
    elif not status.get("database_ok"):
        st.error("No hay conexión sana a la base de datos.")
    st.json(status)
    st.subheader("Validación de repositorio")
    checks = [
        ("requirements.txt", (APP_DIR / "requirements.txt").exists()),
        ("runtime.txt", (APP_DIR / "runtime.txt").exists()),
        (".gitignore", (APP_DIR / ".gitignore").exists()),
        ("secrets.example.toml", (APP_DIR / ".streamlit" / "secrets.example.toml").exists()),
        ("FD-GC71.docx", TEMPLATE_GC71.exists()),
        ("FD-GC72.docx", TEMPLATE_GC72.exists()),
    ]
    st.dataframe(pd.DataFrame(checks, columns=["Elemento", "OK"]), use_container_width=True, hide_index=True)


def main():
    import streamlit as st
    globals()["st"] = st
    init_db()

    if "auth_user" not in st.session_state:
        pantalla_login(st)
        return

    user = st.session_state.get("auth_user", {})
    st.set_page_config(page_title="Gestor Académico Enterprise FD-GC71 / FD-GC72", layout="wide")
    st.title("Gestor académico Enterprise FD-GC71 / FD-GC72")
    st.caption("Planeación inteligente, expediente por curso, evaluación, informes, evidencias, auditoría, perfiles y respaldo.")

    with st.sidebar:
        st.markdown(f"**{user.get('nombre_completo', '')}**")
        st.caption(f"Perfil: {user.get('rol', '')}")
        if st.button("Cerrar sesión", use_container_width=True):
            registrar_auditoria("Logout", "Cierre de sesión")
            st.session_state.pop("auth_user", None)
            st.rerun()
        st.divider()
        modulos = ROLES_PERMISOS.get(user.get("rol"), {}).get("modulos", ["Inicio", "Mi cuenta", "Ayuda / flujo recomendado"])
        pagina = st.radio("Módulo", modulos)

    if not tiene_permiso(pagina):
        st.error("Este perfil no tiene permisos para abrir este módulo.")
        return

    if pagina == "Inicio":
        ui_inicio(st)
    elif pagina == MODULO_CENTRO:
        ui_centro_control(st)
    elif pagina == MODULO_EXPEDIENTE:
        ui_expediente_academico(st)
    elif pagina == MODULO_PLANEADOR:
        ui_planeador_superior(st)
    elif pagina.startswith("FD-GC71"):
        ui_gc71(st)
    elif pagina.startswith("FD-GC72"):
        ui_gc72(st)
    elif pagina == MODULO_EVIDENCIAS:
        ui_evidencias(st)
    elif pagina == MODULO_VALIDADOR:
        ui_validador_institucional(st)
    elif pagina == MODULO_BACKUP:
        ui_backup(st)
    elif pagina == MODULO_DIAGNOSTICO:
        ui_diagnostico_productivo(st)
    elif pagina == "Usuarios y perfiles":
        ui_admin_usuarios(st)
    elif pagina == "Auditoría":
        ui_auditoria(st)
    elif pagina == "Mi cuenta":
        ui_mi_cuenta(st)
    else:
        ui_ayuda(st)


if __name__ == "__main__":
    main()
