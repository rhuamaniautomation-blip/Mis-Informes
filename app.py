# ============================================================================
#  CAVA - GENERADOR INTELIGENTE DE INFORMES DE MANTENIMIENTO
#  Software institucional para la generación automática de informes,
#  reportes y presentaciones de mantenimiento mecánico / eléctrico.
#
#  Diseñado por: CAVA - Especialistas en Robótica y Automatización
#  Autor:        Roger Huamani
#  Versión:      1.1.0  (2026)
#
#  Ejecución:    streamlit run app.py
# ============================================================================

# ----------------------------------------------------------------------------
# 1. IMPORTACIONES
# ----------------------------------------------------------------------------
import os
import io
import re
import json
import time
import base64
import hashlib
import traceback
import uuid
from datetime import datetime, date
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

from PIL import Image as PILImage

# --- Generación documental -------------------------------------------------
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    Image as RLImage, PageBreak, HRFlowable
)
from reportlab.lib.utils import ImageReader

from pptx import Presentation
from pptx.util import Inches as PInches, Pt as PPt, Emu
from pptx.dml.color import RGBColor as PRGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# --- IA de Google ------------------------------------------------------------
try:
    import google.generativeai as genai
    GEMINI_DISPONIBLE = True
except Exception:
    GEMINI_DISPONIBLE = False

# --- Supabase ----------------------------------------------------------------
try:
    from supabase import create_client, Client
    SUPABASE_DISPONIBLE = True
except Exception:
    SUPABASE_DISPONIBLE = False


# ----------------------------------------------------------------------------
# 2. CONSTANTES INSTITUCIONALES Y CONFIGURACIÓN GENERAL
# ----------------------------------------------------------------------------
APP_NOMBRE       = "CAVA | Generador de Informes de Mantenimiento"
APP_VERSION      = "v1.1.0"
EMPRESA          = "CAVA - Especialistas en Robótica y Automatización"
AUTOR_SOFTWARE   = "Roger Huamani"
ANIO_FOOTER      = datetime.now().year

# Paleta de colores institucional (normativa interna CAVA)
COLOR_PRIMARIO   = "#0E3A66"   # Azul institucional
COLOR_SECUNDARIO = "#1B5FA6"   # Azul acero
COLOR_ACENTO     = "#F5A623"   # Naranja ingeniería
COLOR_EXITO      = "#1E7F4F"   # Verde aprobación
COLOR_ERROR      = "#B3282D"   # Rojo alerta
COLOR_FONDO      = "#F4F7FB"   # Fondo claro
COLOR_GRIS       = "#5B6B7C"   # Gris texto secundario

# Rutas base de almacenamiento local (respaldo si no hay Supabase)
BASE_DIR     = Path("cava_data")
DOCS_DIR     = BASE_DIR / "documentos"
CONFIG_FILE  = BASE_DIR / "config.json"
USERS_FILE   = BASE_DIR / "usuarios.json"
INDEX_FILE   = BASE_DIR / "indice.json"

# Credenciales sembradas por defecto (cambiar tras el primer ingreso)
USUARIOS_DEFAULT = {
    "admin":     {"nombre": "Roger Huamani",  "rol": "Superintendente", "clave": "cava2025"},
    "jtorre":    {"nombre": "Jhordan Torre",  "rol": "Técnico",          "clave": "tecnica2025"},
    "lpaucar":   {"nombre": "Lizandro Paucar","rol": "Supervisor",       "clave": "revision2025"},
}

# Modelo de IA por defecto
GEMINI_MODEL_DEFAULT = "gemini-2.0-flash"

# Secciones del informe (clave, título ES, título EN) - según plantilla oficial
SECCIONES_INFORME = [
    ("resumen",         "1. RESUMEN EJECUTIVO",             "1. EXECUTIVE SUMMARY"),
    ("detalles",        "2. DETALLES DEL MANTENIMIENTO",    "2. MAINTENANCE DETAILS"),
    ("resultados",      "3. RESULTADOS DEL MANTENIMIENTO",  "3. MAINTENANCE RESULTS"),
    ("problemas",       "4. PROBLEMAS ENCONTRADOS Y SOLUCIONES", "4. PROBLEMS FOUND AND SOLUTIONS"),
    ("recomendaciones", "5. RECOMENDACIONES",               "5. RECOMMENDATIONS"),
    ("conclusiones",    "6. CONCLUSIONES",                  "6. CONCLUSIONS"),
]

# Catálogos operativos
PLANTAS_DEFAULT   = ["Ensamble DNE", "Pintura", "Stamping", "Trim & Final",
                     "Calidad", "Servicios Generales", "Otra"]
TIPOS_MANTENIMIENTO = ["Mecánico", "Eléctrico", "Mixto (Mecánico-Eléctrico)"]
TIPOS_DOCUMENTO   = ["INFORME DE MANTENIMIENTO", "REPORTE DE MANTENIMIENTO",
                     "INFORME DE INSPECCIÓN", "REPORTE DE INTERVENCIÓN"]
PRIORIDADES       = ["Baja", "Media", "Alta", "Crítica"]

# Estilo de plantilla por defecto (extraído de 'Modelo de Informe.docx')
PLANTILLA_DEFAULT = {
    "fuente_normal":  "Arial",
    "tamano_normal":  11,
    "fuente_titulo":  "Arial",
    "tamano_titulo":  14,
    "fuente_encabezado": "Arial",
    "tamano_encabezado": 12,
    "justificado":    True,
    "interlineado":   1.15,
    "margen_cm":      2.5,
}

# Esquema SQL de Supabase (se muestra en Configuración)
SUPABASE_SQL = """
-- =====================================================================
-- ESQUEMA SUPABASE - CAVA INFORMES DE MANTENIMIENTO
-- Ejecutar en SQL Editor de Supabase
-- =====================================================================
create table if not exists documentos (
    id            uuid primary key default gen_random_uuid(),
    numero        text unique not null,
    tipo          text not null,
    planta        text,
    asunto        text,
    fecha_doc     text,
    elaborado_por text,
    revisado_por  text,
    aprobado_por  text,
    creado_por    text,
    creado_en     timestamptz default now(),
    idioma        text default 'ES',
    resumen       text
);

create or replace function buscar_documento(p_numero text)
returns setof documentos language sql stable as $$
    select * from documentos where numero = p_numero;
$$;

-- Bucket de almacenamiento (crear en Storage UI o con:)
-- insert into storage.buckets (id, name, public)
-- values ('informes-cava', 'informes-cava', false)
-- on conflict (id) do nothing;
"""


# ----------------------------------------------------------------------------
# 3. HOJA DE ESTILOS CSS INSTITUCIONAL
# ----------------------------------------------------------------------------
CSS_INSTITUCIONAL = f"""
<style>
    /* ============ IDENTIDAD CAVA ============ */
    .stApp {{
        background-color: {COLOR_FONDO};
    }}
    header.stHeader {{
        background: transparent;
    }}
    /* Sidebar institucional */
    section[data-testid="stSidebar"] {{
        background: linear-gradient(180deg, {COLOR_PRIMARIO} 0%, #092A4A 100%);
        color: #FFFFFF;
    }}
    section[data-testid="stSidebar"] .stMarkdown,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] .stRadio label span {{
        color: #FFFFFF !important;
        font-family: 'Segoe UI', Arial, sans-serif;
    }}
    section[data-testid="stSidebar"] div[role="radiogroup"] label {{
        background: rgba(255,255,255,0.08);
        border-radius: 8px;
        padding: 6px 10px;
        margin-bottom: 4px;
    }}
    section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {{
        background: rgba(245,166,35,0.25);
    }}
    /* Encabezado de marca */
    .cava-banner {{
        background: linear-gradient(90deg, {COLOR_PRIMARIO} 0%, {COLOR_SECUNDARIO} 100%);
        color: #fff;
        border-radius: 12px;
        padding: 18px 26px;
        margin-bottom: 18px;
        border-left: 8px solid {COLOR_ACENTO};
        box-shadow: 0 4px 14px rgba(14,58,102,0.25);
    }}
    .cava-banner h1 {{
        margin: 0;
        font-size: 26px;
        letter-spacing: 0.5px;
        color: #FFFFFF !important;
    }}
    .cava-banner p {{
        margin: 4px 0 0 0;
        color: #D7E3F4 !important;
        font-size: 13px;
    }}
    /* Tarjetas de información */
    .cava-card {{
        background: #FFFFFF;
        border: 1px solid #DCE4EE;
        border-top: 4px solid {COLOR_SECUNDARIO};
        border-radius: 10px;
        padding: 14px 18px;
        margin-bottom: 14px;
        box-shadow: 0 2px 6px rgba(20,40,80,0.08);
    }}
    .cava-card h3 {{
        color: {COLOR_PRIMARIO} !important;
        margin-top: 0;
    }}
    .cava-card .dato {{
        color: {COLOR_GRIS};
        font-size: 13px;
    }}
    /* Pie de página institucional */
    .cava-footer {{
        margin-top: 40px;
        padding: 14px 0;
        border-top: 2px solid {COLOR_ACENTO};
        text-align: center;
        color: {COLOR_GRIS};
        font-size: 12.5px;
    }}
    .cava-footer b {{ color: {COLOR_PRIMARIO}; }}
    /* Botones principales */
    .stButton>button[kind="primary"],
    button.stButton {{
        border-radius: 8px;
    }}
    div.stButton > button {{
        background: {COLOR_SECUNDARIO};
        color: #fff;
        border: none;
        font-weight: 600;
    }}
    div.stButton > button:hover {{
        background: {COLOR_PRIMARIO};
        color: {COLOR_ACENTO};
    }}
    div.stDownloadButton > button {{
        background: {COLOR_EXITO};
        color: #fff;
    }}
    /* Login */
    .login-card {{
        max-width: 430px;
        margin: 6vh auto 0 auto;
        background: #fff;
        border-radius: 16px;
        padding: 34px 38px;
        box-shadow: 0 12px 40px rgba(9,42,74,0.35);
        border-top: 8px solid {COLOR_ACENTO};
    }}
    .login-logo {{
        text-align: center;
        margin-bottom: 10px;
    }}
    .login-logo .logo-cava {{
        font-size: 44px;
        font-weight: 800;
        color: {COLOR_PRIMARIO};
        letter-spacing: 4px;
    }}
    .login-logo .logo-sub {{
        font-size: 12px;
        color: {COLOR_GRIS};
        letter-spacing: 1.5px;
        text-transform: uppercase;
    }}
    /* Numeración de informe */
    .numero-informe {{
        font-size: 20px;
        font-weight: 700;
        color: {COLOR_PRIMARIO};
        background: #EAF1FA;
        border: 1px dashed {COLOR_SECUNDARIO};
        border-radius: 8px;
        padding: 8px 14px;
        display: inline-block;
    }}
    /* Métricas */
    div[data-testid="stMetricValue"] {{
        color: {COLOR_PRIMARIO} !important;
    }}
    /* Tabs */
    .stTabs [data-baseweb="tab"] {{
        font-weight: 600;
        color: {COLOR_PRIMARIO};
    }}
    .stTabs [aria-selected="true"] {{
        background-color: {COLOR_SECUNDARIO};
        color: #fff !important;
        border-radius: 8px 8px 0 0;
    }}
    /* Expander */
    .streamlit-expanderHeader {{
        background: #EAF1FA;
        border-radius: 8px;
        color: {COLOR_PRIMARIO} !important;
        font-weight: 600;
    }}
    /* Inputs */
    .stTextInput>div>div>input, .stTextArea textarea {{
        border-radius: 8px;
        border-color: #C7D3E2;
    }}
    /* Aviso de seguridad */
    .seguridad-aviso {{
        font-size: 11.5px;
        color: {COLOR_GRIS};
        text-align: center;
        margin-top: 14px;
    }}
</style>
"""


# ----------------------------------------------------------------------------
# 4. UTILIDADES GENERALES
# ----------------------------------------------------------------------------
def asegurar_directorios():
    """Crea las carpetas locales de trabajo si no existen."""
    BASE_DIR.mkdir(parents=True, exist_ok=True)
    DOCS_DIR.mkdir(parents=True, exist_ok=True)


def hoy_iso():
    """Fecha actual en formato ISO."""
    return date.today().isoformat()


def hoy_legible():
    """Fecha actual en formato legible es-PE."""
    return datetime.now().strftime("%d.%m.%Y")


def cargar_json(ruta, default):
    """Carga un archivo JSON con valor por defecto."""
    try:
        if Path(ruta).exists():
            with open(ruta, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return default


def guardar_json(ruta, data):
    """Guarda un objeto como JSON UTF-8."""
    Path(ruta).parent.mkdir(parents=True, exist_ok=True)
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def cargar_config():
    """Configuración global del software (claves, plantilla, etc.)."""
    cfg = cargar_json(CONFIG_FILE, {})
    cfg.setdefault("gemini_api_key", "")
    cfg.setdefault("gemini_model", GEMINI_MODEL_DEFAULT)
    cfg.setdefault("supabase_url", "")
    cfg.setdefault("supabase_key", "")
    cfg.setdefault("plantilla", PLANTILLA_DEFAULT)
    return cfg


def guardar_config(cfg):
    """Persiste la configuración global."""
    guardar_json(CONFIG_FILE, cfg)


def b64_encode(data: bytes) -> str:
    """Codifica bytes a base64 texto."""
    return base64.b64encode(data).decode("utf-8")


def b64_decode(texto: str) -> bytes:
    """Decodifica base64 texto a bytes."""
    return base64.b64decode(texto.encode("utf-8"))


def hash_clave(clave: str, salt: str = None) -> str:
    """Genera hash PBKDF2-SHA256 con sal para almacenamiento seguro."""
    salt = salt or uuid.uuid4().hex
    dk = hashlib.pbkdf2_hmac("sha256", clave.encode(), salt.encode(), 120000)
    return f"{salt}${dk.hex()}"


def verificar_clave(clave: str, almacenado: str) -> bool:
    """Verifica una clave contra su hash almacenado."""
    try:
        salt, _ = almacenado.split("$", 1)
        return hash_clave(clave, salt) == almacenado
    except Exception:
        return False


def limpiar_texto(texto: str) -> str:
    """Normaliza espacios y saltos de línea del texto libre."""
    if not texto:
        return ""
    texto = re.sub(r"[ \t]+", " ", texto)
    texto = re.sub(r"\n{3,}", "\n\n", texto)
    return texto.strip()


def render_footer():
    """Pie de página institucional obligatorio en todas las vistas."""
    st.markdown(
        f"""
        <div class="cava-footer">
            Diseñado por <b>CAVA – Especialistas en Robótica y Automatización</b>
            &nbsp;|&nbsp; Autor: <b>{AUTOR_SOFTWARE}</b>
            &nbsp;|&nbsp; {APP_VERSION} © {ANIO_FOOTER}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_banner(titulo: str, subtitulo: str = ""):
    """Encabezado de marca CAVA."""
    st.markdown(
        f"""
        <div class="cava-banner">
            <h1>⚙️ {titulo}</h1>
            <p>{subtitulo or EMPRESA + ' | ' + AUTOR_SOFTWARE}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ----------------------------------------------------------------------------
# 5. GESTIÓN DE USUARIOS Y SEGURIDAD
# ----------------------------------------------------------------------------
class GestorUsuarios:
    """Administra usuarios locales con contraseñas hasheadas (PBKDF2)."""

    def __init__(self):
        asegurar_directorios()
        self.usuarios = self._cargar()

    def _cargar(self):
        """Carga usuarios o siembra los valores por defecto."""
        data = cargar_json(USERS_FILE, None)
        if not data:
            data = {}
            for usuario, info in USUARIOS_DEFAULT.items():
                data[usuario] = {
                    "nombre": info["nombre"],
                    "rol": info["rol"],
                    "hash": hash_clave(info["clave"]),
                }
            guardar_json(USERS_FILE, data)
        return data

    def autenticar(self, usuario: str, clave: str):
        """Valida credenciales. Devuelve dict de usuario o None."""
        usuario = (usuario or "").strip().lower()
        registro = self.usuarios.get(usuario)
        if not registro:
            return None
        if verificar_clave(clave, registro.get("hash", "")):
            return {"usuario": usuario,
                    "nombre": registro["nombre"],
                    "rol": registro["rol"]}
        return None

    def cambiar_clave(self, usuario: str, clave_actual: str, clave_nueva: str):
        """Cambia la clave de un usuario existente."""
        registro = self.usuarios.get(usuario)
        if not registro or not verificar_clave(clave_actual, registro["hash"]):
            return False, "La clave actual no es correcta."
        registro["hash"] = hash_clave(clave_nueva)
        self.usuarios[usuario] = registro
        guardar_json(USERS_FILE, self.usuarios)
        return True, "Clave actualizada correctamente."

    def listar(self):
        """Lista usuarios registrados."""
        return [
            {"usuario": u, "nombre": d["nombre"], "rol": d["rol"]}
            for u, d in self.usuarios.items()
        ]


# ----------------------------------------------------------------------------
# 6. GESTOR DE IA - GEMINI
# ----------------------------------------------------------------------------
class GestorGemini:
    """Capa de servicio sobre la IA de Google Gemini."""

    def __init__(self, api_key: str, modelo: str = GEMINI_MODEL_DEFAULT):
        self.api_key = api_key
        self.modelo = modelo
        self._modelo = None
        if api_key and GEMINI_DISPONIBLE:
            genai.configure(api_key=api_key)
            self._modelo = genai.GenerativeModel(
                model_name=modelo,
                generation_config={
                    "temperature": 0.55,
                    "top_p": 0.9,
                    "max_output_tokens": 4096,
                },
            )

    @property
    def disponible(self):
        return self._modelo is not None

    def generar(self, prompt: str) -> str:
        """Ejecuta un prompt contra Gemini y devuelve el texto."""
        if not self.disponible:
            raise RuntimeError(
                "Gemini no está configurado. Ingrese su API Key en Configuración."
            )
        respuesta = self._modelo.generate_content(prompt)
        return respuesta.text.strip()

    # ------------------ Prompts de redacción técnica ------------------
    def _prompt_base(self, instruccion: str, contenido: str) -> str:
        """Encapsula la instrucción con el rol de ingeniero senior."""
        return f"""
Eres un Ingeniero Senior de Mantenimiento Industrial con 20 años de experiencia
en plantas automotrices (ensamble, pintura, stamping). Tu tarea es redactar
documentación técnica formal.

REGLAS DE REDACCIÓN:
- Escribe en español técnico, humanizado y profesional, como lo haría un
  ingeniero de mantenimiento experimentado.
- Usa voz pasiva impersonal y tercera persona ("se procedió", "se verificó").
- Sé preciso, ordenado y con vocabulario técnico de confiabilidad industrial
  (p. ej. servomotor, variador, torque, alineación, termografía, etc. según
  corresponda al contexto).
- NO inventes datos numéricos, fechas, nombres ni equipos que no existan en el
  texto original. Si falta información, redacta de forma general.
- Devuelve ÚNICAMENTE el texto redactado, sin encabezados, sin saludos y sin
  comentarios adicionales.

INSTRUCCIÓN ESPECÍFICA:
{instruccion}

TEXTO ORIGINAL DEL TÉCNICO (con sus propias palabras):
\"\"\"{contenido}\"\"\"
"""

    def redactar_seccion(self, clave_seccion: str, narrativa: dict) -> str:
        """Redacta una sección del informe a partir de la narrativa cruda."""
        contexto = self._narrativa_a_texto(narrativa)
        instrucciones = {
            "resumen": (
                "Redacta un RESUMEN EJECUTIVO de máximo 180 palabras que "
                "sintetice la intervención: qué se hizo, dónde, resultado "
                "general y condición final del equipo."
            ),
            "detalles": (
                "Redacta los DETALLES DEL MANTENIMIENTO: describe de forma "
                "cronológica y ordenada las actividades ejecutadas durante la "
                "intervención, herramientas o instrumentos usados si se "
                "mencionan, y condiciones de seguridad aplicadas."
            ),
            "resultados": (
                "Redacta los RESULTADOS DEL MANTENIMIENTO: condición final de "
                "los equipos, parámetros verificados, pruebas funcionales y "
                "estado operativo tras la intervención."
            ),
            "problemas": (
                "Redacta los PROBLEMAS ENCONTRADOS Y SOLUCIONES: enumera de "
                "forma estructurada cada hallazgo o desviación detectada y la "
                "solución aplicada o propuesta para cada uno."
            ),
            "recomendaciones": (
                "Redacta las RECOMENDACIONES: propuestas técnicas de "
                "mejora, seguimiento, monitoreo condicional, plan de "
                "mantenimiento preventivo o compras de repuestos, según "
                "corresponda al contexto."
            ),
            "conclusiones": (
                "Redacta las CONCLUSIONES: cierre técnico breve y contundente "
                "sobre la eficacia de la intervención y la confiabilidad "
                "operativa del sistema intervenido."
            ),
        }
        return limpiar_texto(self.generar(self._prompt_base(
            instrucciones[clave_seccion], contexto)))

    def mejorar_descripcion_imagen(self, descripcion: str, asunto: str) -> str:
        """Ordena y corrige la descripción técnica de una imagen."""
        instruccion = (
            "Corrige, ordena y reestructura la siguiente DESCRIPCIÓN DE "
            "FOTOGRAFÍA TÉCNICA de mantenimiento. Mantén una sola frase o dos "
            "frases descriptivas, en lenguaje técnico de ingeniería, sin "
            "inventar elementos. El contexto de la intervención es: "
            f"{asunto}. Texto a corregir:"
        )
        return limpiar_texto(self.generar(self._prompt_base(
            instruccion, descripcion)))

    def traducir_texto(self, texto: str) -> str:
        """Traduce texto técnico al inglés (documento espejo)."""
        prompt = f"""
Translate the following Spanish maintenance-engineering text into formal,
professional English, as written by a senior maintenance engineer.
Keep technical terminology accurate. Return ONLY the translation.

TEXT:
\"\"\"{texto}\"\"\"
"""
        return limpiar_texto(self.generar(prompt))

    def traducir_data(self, data: dict) -> dict:
        """Genera la versión inglesa completa del informe."""
        data_en = json.loads(json.dumps(data))
        campos = ["asunto", "ubicacion", "equipos", "observaciones"]
        for campo in campos:
            if data_en.get(campo):
                data_en[campo] = self.traducir_texto(data_en[campo])
        for clave, _, _ in SECCIONES_INFORME:
            if data_en.get("secciones", {}).get(clave):
                data_en["secciones"][clave] = self.traducir_texto(
                    data_en["secciones"][clave])
        for img in data_en.get("imagenes", []):
            if img.get("desc_final"):
                img["desc_final"] = self.traducir_texto(img["desc_final"])
        data_en["idioma"] = "EN"
        return data_en

    @staticmethod
    def _narrativa_a_texto(narrativa: dict) -> str:
        """Convierte el dict de narrativa en un solo texto contextual."""
        partes = []
        if narrativa.get("problemas"):
            partes.append("PROBLEMAS IDENTIFICADOS:\n" + narrativa["problemas"])
        if narrativa.get("acciones"):
            partes.append("ACCIONES REALIZADAS:\n" + narrativa["acciones"])
        if narrativa.get("conclusiones"):
            partes.append("CONCLUSIONES DEL TÉCNICO:\n" + narrativa["conclusiones"])
        return "\n\n".join(partes)


# ----------------------------------------------------------------------------
# 7. NUMERACIÓN AUTOMÁTICA DE DOCUMENTOS
# ----------------------------------------------------------------------------
def generar_numero_documento(numeros_existentes: list) -> str:
    """
    Genera el número correlativo con formato DDMMYYYY-NN,
    igual que la plantilla oficial (ej. 21112024-01).
    """
    prefijo = datetime.now().strftime("%d%m%Y")
    correlativo = 1
    for numero in numeros_existentes:
        if numero.startswith(prefijo + "-"):
            try:
                seq = int(numero.split("-")[-1])
                correlativo = max(correlativo, seq + 1)
            except ValueError:
                continue
    return f"{prefijo}-{correlativo:02d}"


# ----------------------------------------------------------------------------
# 8. ALMACENAMIENTO LOCAL (respaldo sin Supabase)
# ----------------------------------------------------------------------------
class AlmacenLocal:
    """Persistencia local de documentos con la misma interfaz que Supabase."""

    nombre = "LOCAL"

    def __init__(self):
        asegurar_directorios()

    def _leer_indice(self):
        return cargar_json(INDEX_FILE, [])

    def _guardar_indice(self, indice):
        guardar_json(INDEX_FILE, indice)

    def listar(self):
        return sorted(self._leer_indice(),
                      key=lambda d: d.get("creado_en", ""), reverse=True)

    def numeros_existentes(self):
        return [d["numero"] for d in self._leer_indice()]

    def guardar(self, data: dict, archivos: dict):
        """archivos: {nombre_archivo: bytes}"""
        numero = data["numero"]
        carpeta = DOCS_DIR / numero
        carpeta.mkdir(parents=True, exist_ok=True)
        for nombre, contenido in archivos.items():
            with open(carpeta / nombre, "wb") as f:
                f.write(contenido)
        guardar_json(carpeta / "data.json", data)
        indice = [d for d in self._leer_indice() if d["numero"] != numero]
        indice.append(self._meta(data))
        self._guardar_indice(indice)
        return True

    def obtener_data(self, numero: str):
        ruta = DOCS_DIR / numero / "data.json"
        return cargar_json(ruta, None)

    def obtener_archivo(self, numero: str, nombre: str):
        ruta = DOCS_DIR / numero / nombre
        if ruta.exists():
            return ruta.read_bytes()
        return None

    def eliminar(self, numero: str):
        import shutil
        carpeta = DOCS_DIR / numero
        if carpeta.exists():
            shutil.rmtree(carpeta)
        self._guardar_indice(
            [d for d in self._leer_indice() if d["numero"] != numero])
        return True

    @staticmethod
    def _meta(data):
        return {
            "numero": data["numero"],
            "tipo": data["tipo_documento"],
            "planta": data.get("planta", ""),
            "asunto": data.get("asunto", ""),
            "fecha_doc": data.get("fecha", ""),
            "elaborado_por": data.get("elaborado_por", ""),
            "revisado_por": data.get("revisado_por", ""),
            "aprobado_por": data.get("aprobado_por", ""),
            "creado_por": data.get("creado_por", ""),
            "creado_en": datetime.now().isoformat(),
        }


# ----------------------------------------------------------------------------
# 9. ALMACENAMIENTO SUPABASE
# ----------------------------------------------------------------------------
class AlmacenSupabase:
    """Persistencia en Supabase (tabla documentos + bucket informes-cava)."""

    nombre = "SUPABASE"

    BUCKET = "informes-cava"

    def __init__(self, url: str, key: str):
        if not SUPABASE_DISPONIBLE:
            raise RuntimeError("Librería supabase no instalada.")
        self.client = create_client(url, key)

    def listar(self):
        resp = self.client.table("documentos").select("*").order(
            "creado_en", desc=True).execute()
        return resp.data

    def numeros_existentes(self):
        resp = self.client.table("documentos").select("numero").execute()
        return [d["numero"] for d in resp.data]

    def guardar(self, data: dict, archivos: dict):
        numero = data["numero"]
        # 1) Subir archivos al bucket
        for nombre, contenido in archivos.items():
            ruta = f"{numero}/{nombre}"
            content_type = "application/octet-stream"
            if nombre.endswith(".docx"):
                content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif nombre.endswith(".pdf"):
                content_type = "application/pdf"
            elif nombre.endswith(".pptx"):
                content_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            elif nombre.endswith(".json"):
                content_type = "application/json"
            try:
                self.client.storage.from_(self.BUCKET).upload(
                    ruta, contenido,
                    file_options={"content_type": content_type,
                                  "upsert": "true"})
            except TypeError:
                self.client.storage.from_(self.BUCKET).upload(
                    ruta, io.BytesIO(contenido),
                    {"content-type": content_type})
        # 2) Registrar / actualizar metadatos
        meta = AlmacenLocal._meta(data)
        existente = self.client.table("documentos").select("id").eq(
            "numero", numero).execute()
        if existente.data:
            self.client.table("documentos").update(meta).eq(
                "numero", numero).execute()
        else:
            self.client.table("documentos").insert(meta).execute()
        return True

    def obtener_data(self, numero: str):
        try:
            contenido = self.client.storage.from_(self.BUCKET).download(
                f"{numero}/data.json")
            return json.loads(contenido.decode("utf-8"))
        except Exception:
            return None

    def obtener_archivo(self, numero: str, nombre: str):
        try:
            return self.client.storage.from_(self.BUCKET).download(
                f"{numero}/{nombre}")
        except Exception:
            return None

    def eliminar(self, numero: str):
        try:
            objetos = self.client.storage.from_(self.BUCKET).list(
                path=numero)
            rutas = [f"{numero}/{o['name']}" for o in objetos]
            if rutas:
                self.client.storage.from_(self.BUCKET).remove(rutas)
        except Exception:
            pass
        self.client.table("documentos").delete().eq("numero", numero).execute()
        return True


def obtener_almacen():
    """Devuelve el almacén activo según configuración."""
    cfg = cargar_config()
    if cfg.get("supabase_url") and cfg.get("supabase_key") and SUPABASE_DISPONIBLE:
        try:
            return AlmacenSupabase(cfg["supabase_url"], cfg["supabase_key"])
        except Exception as e:
            st.warning(f"Supabase no disponible ({e}). Usando almacenamiento local.")
    return AlmacenLocal()


# ----------------------------------------------------------------------------
# 10. GESTOR DE PLANTILLAS DOCUMENTALES
# ----------------------------------------------------------------------------
class GestorPlantillas:
    """Lee el estilo de una plantilla .docx cargada y la aplica a los informes."""

    @staticmethod
    def extraer_estilo(docx_bytes: bytes) -> dict:
        """Extrae fuente, tamaños y márgenes de la plantilla oficial."""
        doc = Document(io.BytesIO(docx_bytes))
        estilo = dict(PLANTILLA_DEFAULT)
        try:
            normal = doc.styles["Normal"]
            if normal.font.name:
                estilo["fuente_normal"] = normal.font.name
            if normal.font.size:
                estilo["tamano_normal"] = normal.font.size.pt
        except Exception:
            pass
        try:
            for s in doc.sections:
                estilo["margen_cm"] = round(s.left_margin.cm, 1)
                break
        except Exception:
            pass
        try:
            primer_parrafo = doc.paragraphs[0]
            for run in primer_parrafo.runs:
                if run.bold and run.font.size:
                    estilo["tamano_titulo"] = run.font.size.pt
                if run.font.name:
                    estilo["fuente_titulo"] = run.font.name
        except Exception:
            pass
        estilo["justificado"] = True
        return estilo

    @staticmethod
    def aplicar(cfg_plantilla: dict):
        """Guarda el estilo de plantilla en la configuración global."""
        cfg = cargar_config()
        cfg["plantilla"] = cfg_plantilla
        guardar_config(cfg)
        return cfg


# ----------------------------------------------------------------------------
# 11. GENERADOR DE DOCUMENTO WORD (DOCX)
# ----------------------------------------------------------------------------
class GeneradorDocx:
    """Construye el informe en Word con formato formal de plantilla oficial."""

    def __init__(self, plantilla: dict = None):
        self.pl = plantilla or cargar_config()["plantilla"]

    # ------------------ helpers de formato ------------------
    def _configurar_estilo_base(self, doc):
        style = doc.styles["Normal"]
        style.font.name = self.pl["fuente_normal"]
        style.font.size = Pt(self.pl["tamano_normal"])
        style.paragraph_format.line_spacing = self.pl["interlineado"]
        rpr = style.element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.append(rfonts)
        rfonts.set(qn("w:ascii"), self.pl["fuente_normal"])
        rfonts.set(qn("w:hAnsi"), self.pl["fuente_normal"])
        rfonts.set(qn("w:eastAsia"), self.pl["fuente_normal"])

    def _margenes(self, doc):
        for section in doc.sections:
            section.top_margin = Cm(self.pl["margen_cm"])
            section.bottom_margin = Cm(self.pl["margen_cm"])
            section.left_margin = Cm(self.pl["margen_cm"])
            section.right_margin = Cm(self.pl["margen_cm"])

    def _parrafo(self, doc, texto, bold=False, size=None, align=None,
                 space_after=6, italic=False, color=None):
        p = doc.add_paragraph()
        run = p.add_run(texto)
        run.bold = bold
        run.italic = italic
        run.font.name = self.pl["fuente_normal"]
        run.font.size = Pt(size or self.pl["tamano_normal"])
        if color:
            run.font.color.rgb = RGBColor.from_string(color)
        if align == "center":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "right":
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif self.pl["justificado"] and not bold:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(space_after)
        return p

    def _borde_tabla(self, tabla):
        tbl = tabla._tbl
        tblPr = tbl.tblPr
        borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "6")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "000000")
            borders.append(el)
        tblPr.append(borders)

    # ------------------ construcción ------------------
    def construir(self, data: dict, idioma: str = "ES") -> bytes:
        doc = Document()
        self._margenes(doc)
        self._configurar_estilo_base(doc)
        es = idioma == "ES"

        # ---------- Encabezado institucional ----------
        self._parrafo(doc, "CAVA – ESPECIALISTAS EN ROBÓTICA Y AUTOMATIZACIÓN",
                      bold=True, size=9, align="center", space_after=2,
                      color="0E3A66")
        self._parrafo(doc, data["tipo_documento"], bold=True,
                      size=self.pl["tamano_titulo"], align="center",
                      space_after=2)
        self._parrafo(doc, f"N° {data['numero']}", bold=True,
                      size=self.pl["tamano_titulo"], align="center",
                      space_after=10)

        # ---------- Datos generales ----------
        lbl_planta = "Planta" if es else "Plant"
        lbl_asunto = "Asunto" if es else "Subject"
        p = doc.add_paragraph()
        r1 = p.add_run(f"{lbl_planta}:  ")
        r1.bold = True
        r2 = p.add_run(data.get("planta", ""))
        p2 = doc.add_paragraph()
        r3 = p2.add_run(f"{lbl_asunto}:  ")
        r3.bold = True
        r4 = p2.add_run(data.get("asunto", ""))
        p2b = doc.add_paragraph()
        r5 = p2b.add_run(("Tipo de mantenimiento:  ") if es
                         else "Maintenance type:  ")
        r5.bold = True
        p2b.add_run(data.get("tipo_mantenimiento", ""))

        # ---------- Tabla de firmas ----------
        tabla = doc.add_table(rows=2, cols=4)
        tabla.alignment = WD_TABLE_ALIGNMENT.CENTER
        self._borde_tabla(tabla)
        hdr = ["Fecha", "Elaborado por", "Revisado por", "Aprobado por"]
        if not es:
            hdr = ["Date", "Prepared by", "Reviewed by", "Approved by"]
        vals = [data.get("fecha", ""), data.get("elaborado_por", ""),
                data.get("revisado_por", ""), data.get("aprobado_por", "")]
        for i, texto in enumerate(hdr):
            celda = tabla.rows[0].cells[i]
            celda.text = ""
            run = celda.paragraphs[0].add_run(texto)
            run.bold = True
            run.font.size = Pt(10)
            run.font.name = self.pl["fuente_normal"]
            celda.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for i, texto in enumerate(vals):
            celda = tabla.rows[1].cells[i]
            celda.text = ""
            run = celda.paragraphs[0].add_run(texto)
            run.font.size = Pt(10)
            run.font.name = self.pl["fuente_normal"]
            celda.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph()

        # ---------- Secciones numeradas ----------
        for clave, titulo_es, titulo_en in SECCIONES_INFORME:
            titulo = titulo_es if es else titulo_en
            self._parrafo(doc, titulo, bold=True,
                          size=self.pl["tamano_encabezado"], space_after=4)
            contenido = data.get("secciones", {}).get(clave, "")
            for linea in (contenido or "").split("\n"):
                if linea.strip():
                    self._parrafo(doc, linea.strip(), space_after=4)
            doc.add_paragraph().paragraph_format.space_after = Pt(2)

        # ---------- Anexos con imágenes correlativas ----------
        titulo_anex = "7. ANEXOS" if es else "7. ANNEXES"
        self._parrafo(doc, titulo_anex, bold=True,
                      size=self.pl["tamano_encabezado"], space_after=6)
        imagenes = data.get("imagenes", [])
        if not imagenes:
            self._parrafo(doc,
                          "No se registran anexos fotográficos." if es
                          else "No photographic annexes are registered.",
                          italic=True)
        for idx, img in enumerate(imagenes, start=1):
            try:
                buf = io.BytesIO(b64_decode(img["bytes_b64"]))
                doc.add_picture(buf, width=Inches(5.4))
                last = doc.paragraphs[-1]
                last.alignment = WD_ALIGN_PARAGRAPH.CENTER
                leyenda = ("Figura" if es else "Figure")
                self._parrafo(
                    doc,
                    f"{leyenda} {idx}: {img.get('desc_final') or img.get('desc_raw', '')}",
                    italic=True, size=10, align="center", space_after=10)
            except Exception:
                continue

        # ---------- Pie de documento ----------
        doc.add_paragraph()
        self._parrafo(
            doc,
            f"Documento generado con CAVA Informes | {EMPRESA} | {AUTOR_SOFTWARE}",
            size=8, align="center", italic=True, color="5B6B7C")

        buffer = io.BytesIO()
        doc.save(buffer)
        return buffer.getvalue()


# ----------------------------------------------------------------------------
# 12. GENERADOR DE DOCUMENTO PDF
# ----------------------------------------------------------------------------
class GeneradorPdf:
    """Construye el informe en PDF con formato formal justificado."""

    def __init__(self):
        self.styles = getSampleStyleSheet()
        self._crear_estilos()

    def _crear_estilos(self):
        self.styles.add(ParagraphStyle(
            name="CavaTitulo", fontName="Helvetica-Bold", fontSize=14,
            alignment=TA_CENTER, textColor=colors.HexColor("#0E3A66"),
            spaceAfter=4))
        self.styles.add(ParagraphStyle(
            name="CavaSub", fontName="Helvetica-Bold", fontSize=9,
            alignment=TA_CENTER, textColor=colors.HexColor("#1B5FA6"),
            spaceAfter=2))
        self.styles.add(ParagraphStyle(
            name="CavaEncabezado", fontName="Helvetica-Bold", fontSize=12,
            textColor=colors.HexColor("#0E3A66"), spaceBefore=10,
            spaceAfter=4))
        self.styles.add(ParagraphStyle(
            name="CavaNormal", fontName="Helvetica", fontSize=10,
            alignment=TA_JUSTIFY, spaceAfter=4, leading=14))
        self.styles.add(ParagraphStyle(
            name="CavaLeyenda", fontName="Helvetica-Oblique", fontSize=9,
            alignment=TA_CENTER, textColor=colors.HexColor("#5B6B7C"),
            spaceAfter=10))
        self.styles.add(ParagraphStyle(
            name="CavaTabla", fontName="Helvetica", fontSize=9,
            alignment=TA_CENTER))
        self.styles.add(ParagraphStyle(
            name="CavaTablaHdr", fontName="Helvetica-Bold", fontSize=9,
            alignment=TA_CENTER, textColor=colors.white))

    def _esc(self, texto):
        return (texto or "").replace("&", "&amp;").replace(
            "<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")

    def construir(self, data: dict, idioma: str = "ES") -> bytes:
        es = idioma == "ES"
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            buffer, pagesize=A4,
            leftMargin=2.5 * cm, rightMargin=2.5 * cm,
            topMargin=2.2 * cm, bottomMargin=2.2 * cm,
            title=f"{data['tipo_documento']} N° {data['numero']}")
        story = []

        # Encabezado
        story.append(Paragraph("CAVA – ESPECIALISTAS EN ROBÓTICA Y AUTOMATIZACIÓN",
                               self.styles["CavaSub"]))
        story.append(Paragraph(self._esc(data["tipo_documento"]),
                               self.styles["CavaTitulo"]))
        story.append(Paragraph(f"N° {data['numero']}", self.styles["CavaTitulo"]))
        story.append(Spacer(1, 8))

        # Datos generales
        lbl = ("Planta" if es else "Plant", "Asunto" if es else "Subject",
               "Tipo de mantenimiento" if es else "Maintenance type")
        story.append(Paragraph(
            f"<b>{lbl[0]}:</b> {self._esc(data.get('planta',''))} &nbsp;&nbsp; "
            f"<b>{lbl[1]}:</b> {self._esc(data.get('asunto',''))}",
            self.styles["CavaNormal"]))
        story.append(Paragraph(
            f"<b>{lbl[2]}:</b> {self._esc(data.get('tipo_mantenimiento',''))}",
            self.styles["CavaNormal"]))
        story.append(Spacer(1, 8))

        # Tabla de firmas
        hdr = (["Fecha", "Elaborado por", "Revisado por", "Aprobado por"] if es
               else ["Date", "Prepared by", "Reviewed by", "Approved by"])
        vals = [data.get("fecha", ""), data.get("elaborado_por", ""),
                data.get("revisado_por", ""), data.get("aprobado_por", "")]
        tabla = Table(
            [[Paragraph(h, self.styles["CavaTablaHdr"]) for h in hdr],
             [Paragraph(self._esc(v), self.styles["CavaTabla"]) for v in vals]],
            colWidths=[3.2 * cm, 4.4 * cm, 4.4 * cm, 4.4 * cm])
        tabla.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0E3A66")),
            ("GRID", (0, 0), (-1, -1), 0.6, colors.HexColor("#000000")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(tabla)
        story.append(Spacer(1, 10))

        # Secciones
        for clave, t_es, t_en in SECCIONES_INFORME:
            story.append(Paragraph(self._esc(t_es if es else t_en),
                                   self.styles["CavaEncabezado"]))
            contenido = data.get("secciones", {}).get(clave, "")
            for linea in (contenido or "").split("\n"):
                if linea.strip():
                    story.append(Paragraph(self._esc(linea.strip()),
                                           self.styles["CavaNormal"]))

        # Anexos
        story.append(Paragraph("7. ANEXOS" if es else "7. ANNEXES",
                               self.styles["CavaEncabezado"]))
        imagenes = data.get("imagenes", [])
        if not imagenes:
            story.append(Paragraph(
                "No se registran anexos fotográficos." if es
                else "No photographic annexes are registered.",
                self.styles["CavaLeyenda"]))
        for idx, img in enumerate(imagenes, start=1):
            try:
                raw = b64_decode(img["bytes_b64"])
                reader = ImageReader(io.BytesIO(raw))
                w, h = reader.getSize()
                max_w, max_h = 15.5 * cm, 10 * cm
                ratio = min(max_w / w, max_h / h)
                story.append(RLImage(io.BytesIO(raw),
                                     width=w * ratio, height=h * ratio,
                                     hAlign="CENTER"))
                story.append(Paragraph(
                    self._esc(f"{'Figura' if es else 'Figure'} {idx}: "
                              f"{img.get('desc_final') or img.get('desc_raw','')}"),
                    self.styles["CavaLeyenda"]))
            except Exception:
                continue

        story.append(Spacer(1, 14))
        story.append(HRFlowable(width="100%", color=colors.HexColor("#F5A623")))
        story.append(Paragraph(
            f"Documento generado con CAVA Informes | {EMPRESA} | {AUTOR_SOFTWARE}",
            self.styles["CavaLeyenda"]))

        doc.build(story)
        return buffer.getvalue()


# ----------------------------------------------------------------------------
# 13. GENERADOR DE PRESENTACIÓN (PPTX)
# ----------------------------------------------------------------------------
class GeneradorPptx:
    """Construye una presentación ejecutiva institucional del informe."""

    AZUL   = PRGBColor(0x0E, 0x3A, 0x66)
    NARANJA = PRGBColor(0xF5, 0xA6, 0x23)
    BLANCO = PRGBColor(0xFF, 0xFF, 0xFF)
    GRIS   = PRGBColor(0x5B, 0x6B, 0x7C)

    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width = PInches(13.333)
        self.prs.slide_height = PInches(7.5)

    def _fondo(self, slide, color):
        fondo = slide.background
        fill = fondo.fill
        fill.solid()
        fill.fore_color.rgb = color

    def _caja_texto(self, slide, left, top, width, height, texto,
                    size=18, bold=False, color=None, align=PP_ALIGN.LEFT):
        caja = slide.shapes.add_textbox(left, top, width, height)
        tf = caja.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = texto
        p.font.size = PPt(size)
        p.font.bold = bold
        p.font.color.rgb = color or self.GRIS
        p.alignment = align
        return tf

    def _diapositiva_seccion(self, titulo, contenido):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self._fondo(slide, self.BLANCO)
        barra = slide.shapes.add_shape(
            1, 0, 0, self.prs.slide_width, PInches(1.1))
        barra.fill.solid()
        barra.fill.fore_color.rgb = self.AZUL
        barra.line.fill.background()
        self._caja_texto(slide, PInches(0.6), PInches(0.22),
                         PInches(12), PInches(0.7), titulo,
                         size=28, bold=True, color=self.BLANCO)
        caja = slide.shapes.add_textbox(PInches(0.7), PInches(1.5),
                                        PInches(11.9), PInches(5.6))
        tf = caja.text_frame
        tf.word_wrap = True
        lineas = [l for l in (contenido or "").split("\n") if l.strip()]
        for i, linea in enumerate(lineas):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = "• " + linea.strip()
            p.font.size = PPt(16)
            p.font.color.rgb = self.GRIS
            p.space_after = PPt(10)
        return slide

    def construir(self, data: dict, idioma: str = "ES") -> bytes:
        es = idioma == "ES"
        # ---------- Portada ----------
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self._fondo(slide, self.AZUL)
        linea = slide.shapes.add_shape(
            1, 0, PInches(4.55), self.prs.slide_width, PInches(0.08))
        linea.fill.solid()
        linea.fill.fore_color.rgb = self.NARANJA
        linea.line.fill.background()
        self._caja_texto(slide, PInches(0.8), PInches(1.6), PInches(11.7),
                         PInches(1.2), data["tipo_documento"],
                         size=40, bold=True, color=self.BLANCO,
                         align=PP_ALIGN.CENTER)
        self._caja_texto(slide, PInches(0.8), PInches(2.7), PInches(11.7),
                         PInches(0.8), f"N° {data['numero']}",
                         size=24, bold=True, color=self.NARANJA,
                         align=PP_ALIGN.CENTER)
        self._caja_texto(slide, PInches(0.8), PInches(4.9), PInches(11.7),
                         PInches(1.2), data.get("asunto", ""),
                         size=22, color=self.BLANCO, align=PP_ALIGN.CENTER)
        self._caja_texto(slide, PInches(0.8), PInches(6.3), PInches(11.7),
                         PInches(0.6),
                         f"{data.get('planta','')}  |  {data.get('fecha','')}",
                         size=14, color=self.BLANCO, align=PP_ALIGN.CENTER)
        self._caja_texto(slide, PInches(0.8), PInches(6.85), PInches(11.7),
                         PInches(0.5),
                         f"{EMPRESA} – {AUTOR_SOFTWARE}",
                         size=11, color=self.NARANJA, align=PP_ALIGN.CENTER)

        # ---------- Secciones ----------
        for clave, t_es, t_en in SECCIONES_INFORME:
            self._diapositiva_seccion(
                t_es if es else t_en,
                data.get("secciones", {}).get(clave, ""))

        # ---------- Anexos fotográficos ----------
        imagenes = data.get("imagenes", [])
        for idx in range(0, len(imagenes), 2):
            slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
            self._fondo(slide, self.BLANCO)
            self._caja_texto(slide, PInches(0.6), PInches(0.3),
                             PInches(12), PInches(0.7),
                             ("7. ANEXOS FOTOGRÁFICOS" if es
                              else "7. PHOTOGRAPHIC ANNEXES"),
                             size=26, bold=True, color=self.AZUL)
            for j, img in enumerate(imagenes[idx:idx + 2]):
                try:
                    raw = b64_decode(img["bytes_b64"])
                    left = PInches(0.7 + j * 6.3)
                    slide.shapes.add_picture(io.BytesIO(raw), left,
                                             PInches(1.4),
                                             width=PInches(5.8))
                    self._caja_texto(
                        slide, left, PInches(5.6), PInches(5.8), PInches(1.4),
                        f"Figura {idx + j + 1}: "
                        f"{img.get('desc_final') or img.get('desc_raw','')}",
                        size=12, color=self.GRIS, align=PP_ALIGN.CENTER)
                except Exception:
                    continue

        buffer = io.BytesIO()
        self.prs.save(buffer)
        return buffer.getvalue()


# ----------------------------------------------------------------------------
# 14. PANTALLA DE LOGIN
# ----------------------------------------------------------------------------
def pagina_login():
    """Vista de autenticación con identidad CAVA."""
    st.markdown(CSS_INSTITUCIONAL, unsafe_allow_html=True)
    st.markdown(
        f"""
        <div class="login-card">
            <div class="login-logo">
                <div class="logo-cava">CAVA</div>
                <div class="logo-sub">Especialistas en Robótica y Automatización</div>
            </div>
            <h3 style="text-align:center;color:{COLOR_PRIMARIO};">
                Generador de Informes de Mantenimiento</h3>
        </div>
        """,
        unsafe_allow_html=True)
    _, col, _ = st.columns([1, 1, 1])
    with col:
        with st.form("form_login"):
            st.markdown("##### 🔐 Acceso restringido al personal autorizado")
            usuario = st.text_input("Usuario", placeholder="ingrese su usuario")
            clave = st.text_input("Contraseña", type="password",
                                  placeholder="••••••••")
            entrar = st.form_submit_button("Ingresar al sistema",
                                           use_container_width=True)
        if entrar:
            gestor = GestorUsuarios()
            perfil = gestor.autenticar(usuario, clave)
            if perfil:
                st.session_state["logueado"] = True
                st.session_state["perfil"] = perfil
                st.session_state["hora_login"] = datetime.now().isoformat()
                st.rerun()
            else:
                st.error("Credenciales incorrectas. Acceso denegado.")
        st.markdown(
            """<div class="seguridad-aviso">
            Sesión protegida con hash PBKDF2-SHA256 ·
            Credencial inicial: <b>admin / cava2025</b> (cámbiela en Configuración)
            </div>""", unsafe_allow_html=True)
    render_footer()


# ----------------------------------------------------------------------------
# 15. PÁGINA: NUEVO INFORME (asistente por pasos)
# ----------------------------------------------------------------------------
def datos_iniciales():
    """Estructura vacía del informe en sesión."""
    return {
        "tipo_documento": TIPOS_DOCUMENTO[0],
        "numero": "",
        "planta": "",
        "area": "",
        "asunto": "",
        "tipo_mantenimiento": TIPOS_MANTENIMIENTO[0],
        "equipos": "",
        "ubicacion": "",
        "prioridad": PRIORIDADES[1],
        "fecha": hoy_legible(),
        "elaborado_por": "",
        "revisado_por": "",
        "aprobado_por": "",
        "narrativa": {"problemas": "", "acciones": "", "conclusiones": ""},
        "secciones": {k: "" for k, _, _ in SECCIONES_INFORME},
        "imagenes": [],
        "idioma": "ES",
        "creado_por": "",
    }


def pagina_nuevo_informe():
    """Asistente de creación del informe en 5 pasos obligatorios."""
    render_banner("Nuevo Informe / Reporte de Mantenimiento",
                  "Complete los pasos en orden. La IA estructurará su narrativa.")
    if "informe" not in st.session_state:
        st.session_state["informe"] = datos_iniciales()
    if "paso" not in st.session_state:
        st.session_state["paso"] = 1

    data = st.session_state["informe"]
    pasos = ["1. Datos generales", "2. Narrativa técnica", "3. Imágenes",
             "4. Redacción IA", "5. Exportar y guardar"]
    st.progress(st.session_state["paso"] / 5,
                text=f"Paso {st.session_state['paso']} de 5: "
                     f"{pasos[st.session_state['paso']-1]}")

    if st.session_state["paso"] == 1:
        _paso_datos_generales(data)
    elif st.session_state["paso"] == 2:
        _paso_narrativa(data)
    elif st.session_state["paso"] == 3:
        _paso_imagenes(data)
    elif st.session_state["paso"] == 4:
        _paso_redaccion_ia(data)
    else:
        _paso_exportar(data)
    render_footer()


def _botones_navegacion(data, paso_actual, validar=True):
    """Botones Atrás / Siguiente con validación de obligatorios."""
    c1, c2, c3 = st.columns([1, 2, 1])
    with c1:
        if paso_actual > 1:
            if st.button("⬅ Atrás", use_container_width=True):
                st.session_state["paso"] = paso_actual - 1
                st.rerun()
    with c3:
        if paso_actual < 5:
            if st.button("Siguiente ➡", use_container_width=True):
                if validar and not _validar_paso(paso_actual, data):
                    return
                st.session_state["paso"] = paso_actual + 1
                st.rerun()


def _validar_paso(paso, data):
    """Valida campos obligatorios de cada paso."""
    if paso == 1:
        faltantes = [c for c in ("planta", "asunto", "fecha", "elaborado_por",
                                 "revisado_por", "aprobado_por")
                     if not data.get(c)]
        if faltantes:
            st.error("Campos obligatorios pendientes: " +
                     ", ".join(faltantes).replace("_", " "))
            return False
    if paso == 2:
        n = data["narrativa"]
        if not (n["problemas"].strip() and n["acciones"].strip()
                and n["conclusiones"].strip()):
            st.error("Debe completar los tres campos de la narrativa técnica.")
            return False
    if paso == 3:
        for img in data["imagenes"]:
            if not img.get("desc_raw", "").strip():
                st.error("Toda imagen cargada requiere su descripción "
                         "obligatoria (Anexos).")
                return False
    return True


# ------------------- PASO 1: DATOS GENERALES -------------------
def _paso_datos_generales(data):
    almacen = obtener_almacen()
    if not data["numero"]:
        data["numero"] = generar_numero_documento(almacen.numeros_existentes())

    with st.container():
        st.markdown('<div class="cava-card"><h3>📋 Numeración automática</h3>'
                    f'<span class="numero-informe">N° {data["numero"]}</span>'
                    '<p class="dato">Generado según formato institucional '
                    'DDMMYYYY-NN (correlativo diario).</p></div>',
                    unsafe_allow_html=True)
        if st.button("🔄 Regenerar número"):
            data["numero"] = generar_numero_documento(
                almacen.numeros_existentes())
            st.rerun()

    with st.form("form_datos_generales"):
        st.markdown("### 1.1 Campos obligatorios del documento")
        c1, c2 = st.columns(2)
        with c1:
            data["tipo_documento"] = st.selectbox(
                "Tipo de documento *", TIPOS_DOCUMENTO,
                index=TIPOS_DOCUMENTO.index(data["tipo_documento"]))
            planta_ops = PLANTAS_DEFAULT
            if data["planta"] and data["planta"] not in planta_ops:
                planta_ops = [data["planta"]] + planta_ops
            data["planta"] = st.selectbox(
                "Planta *", planta_ops,
                index=planta_ops.index(data["planta"]) if data["planta"] in
                planta_ops else 0)
            data["tipo_mantenimiento"] = st.selectbox(
                "Disciplina de mantenimiento *", TIPOS_MANTENIMIENTO,
                index=TIPOS_MANTENIMIENTO.index(data["tipo_mantenimiento"]))
            data["fecha"] = st.text_input(
                "Fecha (dd.mm.aaaa) *", value=data["fecha"])
        with c2:
            data["area"] = st.text_input("Área / línea",
                                         value=data.get("area", ""),
                                         placeholder="Ej. Estación de Encintado Central")
            data["ubicacion"] = st.text_input("Ubicación del equipo",
                                              value=data.get("ubicacion", ""))
            data["prioridad"] = st.selectbox(
                "Prioridad de la intervención", PRIORIDADES,
                index=PRIORIDADES.index(data["prioridad"]))
            data["equipos"] = st.text_input(
                "Equipos / sistemas intervenidos",
                value=data.get("equipos", ""),
                placeholder="Ej. Servomotores SEW / Estación de encintado")
        st.markdown("### 1.2 Asunto y cadena de aprobación")
        data["asunto"] = st.text_input(
            "Asunto *", value=data.get("asunto", ""),
            placeholder="Ej. Regulación de servomotores de Estación de Encintado Central")
        c3, c4, c5 = st.columns(3)
        with c3:
            data["elaborado_por"] = st.text_input(
                "Elaborado por *", value=data.get("elaborado_por", ""),
                placeholder="Técnico responsable")
        with c4:
            data["revisado_por"] = st.text_input(
                "Revisado por *", value=data.get("revisado_por", ""),
                placeholder="Supervisor")
        with c5:
            data["aprobado_por"] = st.text_input(
                "Aprobado por *", value=data.get("aprobado_por", ""),
                placeholder="Superintendente")
        guardar = st.form_submit_button("💾 Guardar datos y continuar",
                                        use_container_width=True)
        if guardar:
            if _validar_paso(1, data):
                st.session_state["paso"] = 2
                st.success("Datos generales validados correctamente.")
                st.rerun()
    _botones_navegacion(data, 1, validar=False)


# ------------------- PASO 2: NARRATIVA TÉCNICA -------------------
def _paso_narrativa(data):
    st.markdown(
        '<div class="cava-card"><h3>✍️ Narrativa del técnico</h3>'
        '<p class="dato">Escriba con sus propias palabras, sin preocuparse por '
        'ortografía ni redacción. La IA de Gemini ordenará, corregirá y '
        'redactará el informe con lenguaje de ingeniería.</p></div>',
        unsafe_allow_html=True)
    n = data["narrativa"]
    n["problemas"] = st.text_area(
        "🔧 Problemas identificados *", height=150, value=n["problemas"],
        placeholder="Ej: los servomotores presentaban juego axial, vibración "
                    "y desajuste en los topes de la estación de encintado...")
    n["acciones"] = st.text_area(
        "🛠️ Acciones realizadas *", height=150, value=n["acciones"],
        placeholder="Ej: se reguló el juego axial, se ajustaron los topes, "
                    "se verificó el torque y se probó en automático...")
    n["conclusiones"] = st.text_area(
        "✅ Conclusiones del técnico *", height=120, value=n["conclusiones"],
        placeholder="Ej: el equipo quedó operativo, se recomienda verificar "
                    "en una semana...")
    _botones_navegacion(data, 2)


# ------------------- PASO 3: IMÁGENES -------------------
def _paso_imagenes(data):
    st.markdown(
        '<div class="cava-card"><h3>📷 Anexos fotográficos</h3>'
        '<p class="dato">Cargue las evidencias. Cada imagen exige su '
        'descripción obligatoria y recibirá numeración correlativa '
        '(Figura 1, Figura 2, ...). La IA puede ordenar y corregir la '
        'descripción.</p></div>', unsafe_allow_html=True)

    archivos = st.file_uploader(
        "Cargar imágenes (JPG/PNG)", type=["jpg", "jpeg", "png"],
        accept_multiple_files=True, key="up_imagenes")
    if archivos:
        nombres_actuales = {i["nombre"] for i in data["imagenes"]}
        for f in archivos:
            if f.name not in nombres_actuales:
                data["imagenes"].append({
                    "id": uuid.uuid4().hex[:8],
                    "nombre": f.name,
                    "bytes_b64": b64_encode(f.read()),
                    "desc_raw": "",
                    "desc_final": "",
                })
        st.session_state["procesar_uploader"] = True

    if data["imagenes"]:
        st.markdown(f"##### Imágenes cargadas: {len(data['imagenes'])} "
                    "(numeración correlativa automática)")
        for idx, img in enumerate(data["imagenes"], start=1):
            with st.container():
                c1, c2 = st.columns([1, 2])
                with c1:
                    st.image(b64_decode(img["bytes_b64"]),
                             caption=f"Figura {idx}", use_container_width=True)
                with c2:
                    st.markdown(f"**Figura {idx}** – `{img['nombre']}`")
                    img["desc_raw"] = st.text_area(
                        "Descripción obligatoria *", value=img["desc_raw"],
                        height=80, key=f"desc_{img['id']}",
                        placeholder="Ej: vista del servomotor antes de la "
                                    "regulación con el tope desajustado")
                    if img.get("desc_final"):
                        st.info(f"**Descripción estructurada (IA):** "
                                f"{img['desc_final']}")
                    if st.button(f"🤖 Ordenar descripción con IA",
                                 key=f"ia_img_{img['id']}"):
                        gemini = _gemini_de_config()
                        if gemini and gemini.disponible:
                            with st.spinner("Corrigiendo y estructurando..."):
                                img["desc_final"] = gemini.mejorar_descripcion_imagen(
                                    img["desc_raw"], data.get("asunto", ""))
                            st.rerun()
                        else:
                            st.warning("Configure su API Key de Gemini en "
                                       "Configuración.")
                if st.button(f"🗑 Quitar Figura {idx}", key=f"del_{img['id']}"):
                    data["imagenes"] = [i for i in data["imagenes"]
                                        if i["id"] != img["id"]]
                    st.rerun()
    else:
        st.info("Sin imágenes por el momento. Puede continuar sin anexos.")
    _botones_navegacion(data, 3)


# ------------------- PASO 4: REDACCIÓN IA -------------------
def _gemini_de_config():
    cfg = cargar_config()
    if not cfg.get("gemini_api_key"):
        return None
    return GestorGemini(cfg["gemini_api_key"], cfg.get("gemini_model",
                                                       GEMINI_MODEL_DEFAULT))


def _paso_redaccion_ia(data):
    gemini = _gemini_de_config()
    st.markdown(
        '<div class="cava-card"><h3>🤖 Redacción inteligente</h3>'
        '<p class="dato">Gemini redactará las 6 secciones técnicas del informe '
        'con tono humanizado de ingeniero senior. Luego podrá editar cada '
        'sección manualmente antes de exportar.</p></div>',
        unsafe_allow_html=True)

    if not (gemini and gemini.disponible):
        st.error("⚠️ Configure su **API Key de Gemini** en el menú "
                 "**Configuración** para habilitar la redacción automática.")
        _botones_navegacion(data, 4, validar=False)
        return

    if st.button("⚡ Generar redacción técnica con IA", type="primary"):
        with st.spinner("La IA está redactando el informe como ingeniero..."):
            try:
                for clave, _, _ in SECCIONES_INFORME:
                    data["secciones"][clave] = gemini.redactar_seccion(
                        clave, data["narrativa"])
                # Descripciones de imágenes aún sin estructurar
                for img in data["imagenes"]:
                    if img["desc_raw"] and not img["desc_final"]:
                        img["desc_final"] = gemini.mejorar_descripcion_imagen(
                            img["desc_raw"], data.get("asunto", ""))
                st.session_state["ia_generada"] = True
                st.success("Redacción técnica completada. Revise y edite si lo "
                           "requiere.")
            except Exception as e:
                st.error(f"Error al contactar Gemini: {e}")

    if st.session_state.get("ia_generada"):
        tabs = st.tabs([t for _, t, _ in SECCIONES_INFORME])
        for tab, (clave, titulo, _) in zip(tabs, SECCIONES_INFORME):
            with tab:
                data["secciones"][clave] = st.text_area(
                    f"Edición manual – {titulo}",
                    value=data["secciones"][clave], height=200,
                    key=f"edit_{clave}")
    _botones_navegacion(data, 4, validar=False)


# ------------------- PASO 5: EXPORTAR Y GUARDAR -------------------
def _paso_exportar(data):
    st.markdown(
        '<div class="cava-card"><h3>📤 Exportación y archivo institucional</h3>'
        '<p class="dato">Seleccione formatos de salida. Si activa la '
        'traducción, se generará el documento espejo en inglés '
        '(sufijo _EN). Todo se archivará automáticamente.</p></div>',
        unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        exp_word = st.checkbox("📄 Word (.docx)", value=True)
        exp_pdf = st.checkbox("📑 PDF (.pdf)", value=True)
        exp_pptx = st.checkbox("📽️ Presentación (.pptx)", value=False)
    with c2:
        traducir = st.checkbox("🌐 Generar documento espejo en inglés")
        st.markdown(f"**Número asignado:** `{data['numero']}`")
        st.markdown(f"**Almacén activo:** `{obtener_almacen().nombre}`")

    if st.button("🚀 Generar, archivar y descargar", type="primary"):
        with st.spinner("Construyendo documentos formales..."):
            try:
                plantilla = cargar_config()["plantilla"]
                data["creado_por"] = st.session_state["perfil"]["nombre"]
                archivos = {}

                # -------- versión española --------
                if exp_word:
                    archivos[f"{data['numero']}.docx"] = GeneradorDocx(
                        plantilla).construir(data, "ES")
                if exp_pdf:
                    archivos[f"{data['numero']}.pdf"] = GeneradorPdf().construir(
                        data, "ES")
                if exp_pptx:
                    archivos[f"{data['numero']}.pptx"] = GeneradorPptx().construir(
                        data, "ES")

                # -------- versión inglesa (espejo) --------
                if traducir:
                    gemini = _gemini_de_config()
                    if gemini and gemini.disponible:
                        with st.spinner("Traduciendo documento espejo al inglés..."):
                            data_en = gemini.traducir_data(data)
                        num_en = f"{data['numero']}_EN"
                        if exp_word:
                            archivos[f"{num_en}.docx"] = GeneradorDocx(
                                plantilla).construir(data_en, "EN")
                        if exp_pdf:
                            archivos[f"{num_en}.pdf"] = GeneradorPdf().construir(
                                data_en, "EN")
                        if exp_pptx:
                            archivos[f"{num_en}.pptx"] = GeneradorPptx().construir(
                                data_en, "EN")
                    else:
                        st.warning("Traducción omitida: configure Gemini.")

                # -------- archivo permanente --------
                almacen = obtener_almacen()
                almacen.guardar(data, archivos)
                st.session_state["archivos_generados"] = archivos
                st.success(f"Informe N° {data['numero']} archivado en "
                           f"{almacen.nombre} correctamente.")
            except Exception as e:
                st.error(f"Error durante la generación: {e}")
                st.code(traceback.format_exc())

    # -------- descargas --------
    generados = st.session_state.get("archivos_generados", {})
    if generados:
        st.markdown("#### ⬇️ Descargas disponibles")
        cols = st.columns(min(len(generados), 4))
        for i, (nombre, contenido) in enumerate(generados.items()):
            with cols[i % len(cols)]:
                st.download_button(
                    f"⬇️ {nombre}", data=contenido, file_name=nombre,
                    use_container_width=True, key=f"dl_{nombre}")
        if st.button("🆕 Crear otro informe"):
            st.session_state["informe"] = datos_iniciales()
            st.session_state["paso"] = 1
            st.session_state["ia_generada"] = False
            st.session_state["archivos_generados"] = {}
            st.rerun()


# ----------------------------------------------------------------------------
# 16. PÁGINA: HISTORIAL / ARCHIVO TÉCNICO
# ----------------------------------------------------------------------------
def pagina_historial():
    render_banner("Historial y Archivo Técnico",
                  "Documentos archivados en el almacén institucional.")
    almacen = obtener_almacen()
    documentos = almacen.listar()
    if not documentos:
        st.info("Aún no existen documentos archivados.")
        render_footer()
        return

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Documentos archivados", len(documentos))
    c2.metric("Almacén activo", almacen.nombre)
    c3.metric("Informes", sum(1 for d in documentos
                              if "INFORME" in d.get("tipo", "")))
    c4.metric("Reportes", sum(1 for d in documentos
                              if "REPORTE" in d.get("tipo", "")))

    buscar = st.text_input("🔎 Buscar por número, asunto o planta", "")
    filtrados = [d for d in documentos if buscar.lower() in
                 json.dumps(d, ensure_ascii=False).lower()] if buscar else documentos

    for doc_meta in filtrados:
        numero = doc_meta["numero"]
        with st.expander(f"📄 N° {numero} — {doc_meta.get('asunto','')} "
                         f"({doc_meta.get('fecha_doc','')})"):
            c1, c2 = st.columns([2, 1])
            with c1:
                st.markdown(f"""
                <div class="cava-card">
                    <h3>{doc_meta.get('tipo','')}</h3>
                    <p class="dato"><b>Planta:</b> {doc_meta.get('planta','')} ·
                    <b>Elaborado:</b> {doc_meta.get('elaborado_por','')} ·
                    <b>Revisado:</b> {doc_meta.get('revisado_por','')} ·
                    <b>Aprobado:</b> {doc_meta.get('aprobado_por','')}</p>
                    <p class="dato"><b>Asunto:</b> {doc_meta.get('asunto','')}</p>
                </div>""", unsafe_allow_html=True)
            with c2:
                for ext in ("docx", "pdf", "pptx"):
                    contenido = almacen.obtener_archivo(numero, f"{numero}.{ext}")
                    if contenido:
                        st.download_button(f"⬇️ {ext.upper()}", data=contenido,
                                           file_name=f"{numero}.{ext}",
                                           key=f"h_{numero}_{ext}")
                if st.button("✏️ Cargar para editar", key=f"ed_{numero}"):
                    data = almacen.obtener_data(numero)
                    if data:
                        st.session_state["informe"] = data
                        st.session_state["paso"] = 1
                        st.session_state["ia_generada"] = True
                        st.session_state["archivos_generados"] = {}
                        st.success("Documento cargado en el editor. Vaya al "
                                   "menú 'Nuevo informe' para editarlo.")
                        st.rerun()
                    else:
                        st.error("No se pudo recuperar el documento.")
                if st.button("🗑 Eliminar del archivo", key=f"el_{numero}"):
                    almacen.eliminar(numero)
                    st.rerun()
    render_footer()


# ----------------------------------------------------------------------------
# 17. PÁGINA: CONFIGURACIÓN
# ----------------------------------------------------------------------------
def pagina_configuracion():
    render_banner("Configuración del Sistema",
                  "Claves de servicio, plantillas y seguridad.")
    cfg = cargar_config()
    tabs = st.tabs(["🔑 Servicios IA / Supabase", "📐 Plantilla documental",
                    "👥 Usuarios y seguridad", "🗄️ Esquema Supabase"])

    with tabs[0]:
        st.markdown("### Inteligencia Artificial (Google Gemini)")
        gk = st.text_input("Gemini API Key", value=cfg.get("gemini_api_key", ""),
                           type="password",
                           help="Obtenga su clave en https://aistudio.google.com")
        gm = st.selectbox("Modelo", ["gemini-2.0-flash", "gemini-1.5-flash",
                                     "gemini-1.5-pro"],
                          index=0 if cfg.get("gemini_model") == "gemini-2.0-flash"
                          else 1)
        st.markdown("### Almacenamiento en la nube (Supabase)")
        su = st.text_input("Supabase URL", value=cfg.get("supabase_url", ""))
        sk = st.text_input("Supabase Key (service_role/anon)",
                           value=cfg.get("supabase_key", ""), type="password")
        if st.button("💾 Guardar configuración"):
            cfg.update(gemini_api_key=gk, gemini_model=gm,
                       supabase_url=su, supabase_key=sk)
            guardar_config(cfg)
            st.success("Configuración guardada.")
            st.rerun()
        c1, c2 = st.columns(2)
        with c1:
            if st.button("🧪 Probar Gemini"):
                try:
                    g = GestorGemini(gk or cfg.get("gemini_api_key", ""), gm)
                    r = g.generar("Responde únicamente: CONEXIÓN OK")
                    st.success(f"Gemini operativo: {r[:60]}")
                except Exception as e:
                    st.error(f"Fallo de conexión: {e}")
        with c2:
            if st.button("🧪 Probar Supabase"):
                try:
                    a = AlmacenSupabase(su or cfg["supabase_url"],
                                        sk or cfg["supabase_key"])
                    st.success(f"Supabase operativo. Documentos: "
                               f"{len(a.listar())}")
                except Exception as e:
                    st.error(f"Fallo de conexión: {e}")

    with tabs[1]:
        st.markdown("### Carga de plantilla oficial (.docx)")
        st.info("El software extraerá fuente, tamaño, justificado y márgenes "
                "de la plantilla cargada (ej. 'Modelo de Informe.docx') y los "
                "aplicará a todos los documentos generados.")
        tpl = st.file_uploader("Plantilla institucional", type=["docx"])
        if tpl and st.button("📐 Aplicar estilo de plantilla"):
            estilo = GestorPlantillas.extraer_estilo(tpl.read())
            GestorPlantillas.aplicar(estilo)
            st.success(f"Plantilla aplicada: fuente {estilo['fuente_normal']} "
                       f"{estilo['tamano_normal']} pt, margen "
                       f"{estilo['margen_cm']} cm, justificado activo.")
        st.json(cargar_config()["plantilla"])

    with tabs[2]:
        gestor = GestorUsuarios()
        st.dataframe(gestor.listar(), use_container_width=True,
                     hide_index=True)
        st.markdown("#### Cambiar mi contraseña")
        ca = st.text_input("Clave actual", type="password")
        cn = st.text_input("Clave nueva", type="password")
        if st.button("🔒 Cambiar contraseña"):
            ok, msg = gestor.cambiar_clave(
                st.session_state["perfil"]["usuario"], ca, cn)
            (st.success if ok else st.error)(msg)

    with tabs[3]:
        st.markdown("### Script SQL para Supabase")
        st.code(SUPABASE_SQL, language="sql")
        st.caption("Ejecútelo en el SQL Editor de su proyecto Supabase y "
                   "cree el bucket 'informes-cava' en Storage.")
    render_footer()


# ----------------------------------------------------------------------------
# 18. PÁGINA: ACERCA DE
# ----------------------------------------------------------------------------
def pagina_acerca():
    render_banner("Acerca del Software", APP_NOMBRE)
    st.markdown(
        f"""
        <div class="cava-card">
            <h3>🏭 Propósito</h3>
            <p class="dato">Plataforma institucional que transforma la
            narrativa libre de los técnicos de mantenimiento mecánico y
            eléctrico en informes, reportes y presentaciones formales,
            redactados con inteligencia artificial (Gemini) con tono de
            ingeniería humanizado, numeración correlativa automática,
            anexos fotográficos obligatoriamente descritos, exportación
            Word/PDF/PPTX con documento espejo en inglés y archivo
            permanente en Supabase.</p>
        </div>
        <div class="cava-card">
            <h3>🧠 Autoría</h3>
            <p class="dato">Diseñado por <b>CAVA – Especialistas en Robótica y
            Automatización</b>.<br>Desarrollado por <b>{AUTOR_SOFTWARE}</b>.
            <br>Versión {APP_VERSION} · {ANIO_FOOTER}.</p>
        </div>
        """, unsafe_allow_html=True)
    st.markdown("#### Flujo del proceso")
    st.write("""
    1. **Login seguro** con credenciales hasheadas (PBKDF2).
    2. **Datos generales** obligatorios y numeración DDMMYYYY-NN automática.
    3. **Narrativa libre** del técnico (problemas, acciones, conclusiones).
    4. **Imágenes** con descripción obligatoria y numeración correlativa.
    5. **Gemini** redacta las 7 secciones con lenguaje de ingeniero.
    6. **Exportación** formal (Word / PDF / PPTX) + espejo inglés opcional.
    7. **Archivo** automático en Supabase para revisión y edición posterior.
    """)
    render_footer()


# ----------------------------------------------------------------------------
# 19. NAVEGACIÓN PRINCIPAL Y ARRANQUE
# ----------------------------------------------------------------------------
def render_sidebar():
    """Barra lateral institucional con identidad CAVA."""
    with st.sidebar:
        st.markdown(
            f"""
            <div style="text-align:center;padding:10px 0 20px 0;">
                <div style="font-size:34px;font-weight:800;letter-spacing:5px;
                            color:#F5A623;">CAVA</div>
                <div style="font-size:10.5px;color:#D7E3F4;letter-spacing:1px;
                            text-transform:uppercase;">
                    Especialistas en Robótica y Automatización</div>
            </div>
            """, unsafe_allow_html=True)
        perfil = st.session_state["perfil"]
        st.markdown(f"👤 **{perfil['nombre']}**  \n🎖️ {perfil['rol']}")
        opcion = st.radio("Módulos", [
            "🆕 Nuevo informe",
            "🗂 Historial / Archivo",
            "⚙️ Configuración",
            "ℹ️ Acerca de",
        ], label_visibility="collapsed")
        st.markdown("---")
        if st.button("🚪 Cerrar sesión", use_container_width=True):
            for k in ("logueado", "perfil", "informe", "paso", "ia_generada",
                      "archivos_generados"):
                st.session_state.pop(k, None)
            st.rerun()
        st.markdown(
            f"<div style='text-align:center;font-size:10px;color:#9FB3CE;'>"
            f"Diseñado por CAVA · {AUTOR_SOFTWARE}<br>{APP_VERSION}</div>",
            unsafe_allow_html=True)
        return opcion


def main():
    """Punto de entrada principal de la aplicación Streamlit."""
    st.set_page_config(
        page_title=APP_NOMBRE,
        page_icon="⚙️",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    asegurar_directorios()
    st.markdown(CSS_INSTITUCIONAL, unsafe_allow_html=True)

    if not st.session_state.get("logueado"):
        pagina_login()
        return

    opcion = render_sidebar()
    if opcion == "🆕 Nuevo informe":
        pagina_nuevo_informe()
    elif opcion == "🗂 Historial / Archivo":
        pagina_historial()
    elif opcion == "⚙️ Configuración":
        pagina_configuracion()
    else:
        pagina_acerca()


# ============================================================================
# ARRANQUE
# ============================================================================
if __name__ == "__main__" or True:
    main()

# ------------------------- FIN DEL SOFTWARE CAVA ----------------------------
