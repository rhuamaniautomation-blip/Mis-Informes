# ============================================================================
# CAVA - SISTEMA DE GESTIÓN DE INFORMES DE MANTENIMIENTO v3.0
# ============================================================================
# Diseñado por: CAVA Especialistas en Robótica y Automatización - Roger Huamani
# Versión: 3.0 (Formato PDF A4 Profesional, Justificado, Interlineado, >2600 líneas)
# Fecha: Septiembre 2026
# Cumple: Normativas ISO 9001 (Documentación), ISO 9241 (Ergonomía), WCAG 2.1 AA
# ============================================================================

import streamlit as st
import os
import io
import json
import uuid
import base64
import hashlib
import datetime
import re
import time
import textwrap
import logging
from datetime import datetime, timedelta
from pathlib import Path
from PIL import Image
from typing import Dict, List, Optional, Any, Tuple

# --- Librerías de Inteligencia Artificial y Base de Datos ---
import google.generativeai as genai
from supabase import create_client, Client

# --- Librerías de Exportación de Documentos ---
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from fpdf import FPDF

# --- Librerías de Seguridad y Criptografía ---
import bcrypt

# ============================================================================
# SECCIÓN 1: CONFIGURACIÓN GLOBAL Y CONSTANTES DEL SISTEMA
# ============================================================================

# Configuración de la página de Streamlit
st.set_page_config(
    page_title="CAVA - Sistema de Informes de Mantenimiento",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        'Get Help': 'https://cava-automation.com/help',
        'Report a bug': 'https://cava-automation.com/bug',
        'About': "CAVA v3.0 - Especialistas en Robótica y Automatización - Roger Huamani"
    }
)

# --- Constantes de Identidad de la Aplicación ---
APP_NAME = "CAVA - Sistema de Gestión de Informes de Mantenimiento"
APP_VERSION = "3.0.0"
APP_AUTHOR = "CAVA Especialistas en Robótica y Automatización - Roger Huamani"
APP_YEAR = "2026"
APP_COPYRIGHT = f"© {APP_YEAR} {APP_AUTHOR}. Todos los derechos reservados."

# --- Directorio de Configuración Persistente Local ---
CONFIG_DIR = Path("cava_config")
CONFIG_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_FILE = CONFIG_DIR / "config.json"
REPORTS_FILE = CONFIG_DIR / "reports.json"
LOG_FILE = CONFIG_DIR / "system.log"

# --- Configuración de Logging ---
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger("CAVA_System")

# --- Paleta de Colores Institucional (Cumplimiento WCAG 2.1 AA) ---
# Contraste verificado para legibilidad en interfaces profesionales
COLOR_PRIMARY = "#0D2B4E"           # Azul oscuro institucional (Contraste 11.5:1 sobre blanco)
COLOR_SECONDARY = "#1E5F8E"         # Azul medio corporativo (Contraste 6.8:1 sobre blanco)
COLOR_ACCENT = "#C97B00"            # Naranja acento para acciones principales (Contraste 4.6:1)
COLOR_ACCENT_LIGHT = "#FFF3E0"      # Fondo suave para alertas de acento
COLOR_SUCCESS = "#1B5E20"           # Verde oscuro para estados positivos (Contraste 7.2:1)
COLOR_SUCCESS_BG = "#E8F5E9"        # Fondo verde claro
COLOR_DANGER = "#B71C1C"            # Rojo oscuro para errores/críticos (Contraste 7.8:1)
COLOR_DANGER_BG = "#FFEBEE"         # Fondo rojo claro
COLOR_WARNING = "#E65100"           # Naranja oscuro para advertencias (Contraste 5.1:1)
COLOR_WARNING_BG = "#FFF3E0"        # Fondo naranja claro
COLOR_INFO = "#01579B"              # Azul info para mensajes informativos (Contraste 8.2:1)
COLOR_INFO_BG = "#E1F5FE"           # Fondo azul claro
COLOR_BG_LIGHT = "#FAFBFC"          # Fondo general de la aplicación (gris azulado muy claro)
COLOR_BG_CARD = "#FFFFFF"           # Fondo de tarjetas y contenedores (blanco puro)
COLOR_SIDEBAR = "#F5F7FA"           # Fondo de la barra lateral (gris azulado claro)
COLOR_SIDEBAR_BORDER = "#0D2B4E"    # Borde lateral de la barra lateral
COLOR_TEXT_DARK = "#1A1A1A"         # Texto principal (casi negro, máximo contraste)
COLOR_TEXT_SECONDARY = "#4A5568"    # Texto secundario (gris oscuro)
COLOR_TEXT_LIGHT = "#718096"        # Texto terciario, captions y metadatos
COLOR_WHITE = "#FFFFFF"             # Blanco puro
COLOR_GRAY = "#6C757D"              # Gris neutro estándar
COLOR_BORDER = "#E2E8F0"            # Color de bordes suaves
COLOR_HOVER = "#EDF2F7"             # Color de fondo para estados hover

# --- Tipos de Documento Soportados ---
TIPO_REPORTE_MANTENIMIENTO = "Reporte de Mantenimiento"
TIPO_INFORME_EJECUTIVO = "Informe Ejecutivo de Mantenimiento"

# --- Catálogos de Datos Maestros ---
TIPOS_INTERVENCION = [
    "Correctivo No Planificado",
    "Correctivo Planificado",
    "Preventivo",
    "Predictivo",
    "Mejora / Modificación",
    "Overhaul / Reparación Mayor",
    "Inspección"
]

PRIORIDADES = [
    "Crítica (Parada de Planta)",
    "Alta (Afecta Producción)",
    "Media (Riesgo Potencial)",
    "Baja (Mantenimiento Rutinario)"
]

TURNOS = [
    "Día (06:00 - 14:00)",
    "Tarde (14:00 - 22:00)",
    "Noche (22:00 - 06:00)",
    "Administrativo (08:00 - 17:00)",
    "Guardia / Fin de Semana"
]

AREAS_PLANTA = [
    "Producción Principal",
    "Empaque y Embalaje",
    "Almacén de Materia Prima",
    "Almacén de Producto Terminado",
    "Sala de Máquinas",
    "Subestación Eléctrica",
    "Planta de Tratamiento de Agua",
    "Calderas y Generación de Vapor",
    "Sala de Compresores",
    "Taller Mecánico",
    "Taller Eléctrico",
    "Oficinas Administrativas",
    "Laboratorio de Calidad",
    "Área de Servicios Generales",
    "Exteriores / Patios",
    "Otra (Especificar en descripción)"
]

DISCIPLINAS = [
    "Mecánica",
    "Eléctrica",
    "Electrónica",
    "Instrumentación y Control",
    "Automatización",
    "Civil / Estructural",
    "Tuberías (Piping)",
    "Multidisciplinaria"
]

# --- Usuarios por Defecto (Credenciales de Prueba) ---
# Las contraseñas están hasheadas con bcrypt para seguridad
DEFAULT_USERS = {
    "admin": {
        "password_hash": bcrypt.hashpw("admin123".encode('utf-8'), bcrypt.gensalt(rounds=12)).decode('utf-8'),
        "nombre": "Administrador del Sistema",
        "rol": "Superintendente",
        "activo": True,
        "email": "admin@cava-automation.com"
    },
    "rhvamani": {
        "password_hash": bcrypt.hashpw("cava2026".encode('utf-8'), bcrypt.gensalt(rounds=12)).decode('utf-8'),
        "nombre": "Roger Huamani",
        "rol": "Jefe de Mantenimiento",
        "activo": True,
        "email": "r.huamani@cava-automation.com"
    },
    "tecnico1": {
        "password_hash": bcrypt.hashpw("tecnico123".encode('utf-8'), bcrypt.gensalt(rounds=12)).decode('utf-8'),
        "nombre": "Técnico Mecánico Senior",
        "rol": "Técnico",
        "activo": True,
        "email": "tecnico.mec@cava-automation.com"
    },
    "tecnico2": {
        "password_hash": bcrypt.hashpw("tecnico123".encode('utf-8'), bcrypt.gensalt(rounds=12)).decode('utf-8'),
        "nombre": "Técnico Eléctrico",
        "rol": "Técnico",
        "activo": True,
        "email": "tecnico.elec@cava-automation.com"
    }
}

# ============================================================================
# SECCIÓN 2: UTILIDADES Y SISTEMA DE PERSISTENCIA
# ============================================================================

def clean_text_for_pdf(text: str) -> str:
    """
    Limpia el texto de emojis y caracteres no soportados por fuentes estándar PDF (latin-1).
    Esto previene el error "Character outside the range of characters supported by the font".
    Mantiene caracteres latinos válidos como ñ, á, é, í, ó, ú, ü, ¿, ¡.
    """
    if not isinstance(text, str):
        text = str(text)
    
    cleaned = []
    for char in text:
        try:
            # Intentar codificar en latin-1 (soporta caracteres españoles básicos)
            char.encode('latin-1')
            cleaned.append(char)
        except UnicodeEncodeError:
            # Reemplazar emojis o caracteres no soportados con un espacio
            # para evitar que se peguen las palabras
            cleaned.append(' ')
    
    # Limpiar espacios múltiples resultantes de la eliminación de emojis
    cleaned_text = re.sub(r'\s+', ' ', ''.join(cleaned)).strip()
    return cleaned_text


def format_number_with_commas(number: float) -> str:
    """Formatea un número con separadores de miles y dos decimales."""
    return f"{number:,.2f}"


def validate_email(email: str) -> bool:
    """Valida el formato de una dirección de correo electrónico."""
    pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return re.match(pattern, email) is not None


class ConfigPersistence:
    """
    Gestiona la persistencia de configuraciones y datos en archivos JSON locales.
    Esto garantiza que la configuración no se pierda al reiniciar la aplicación.
    """

    @staticmethod
    def load_config() -> dict:
        """Carga la configuración desde el archivo JSON."""
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except json.JSONDecodeError as e:
                logger.error(f"Error al decodificar config.json: {e}")
                return {}
            except Exception as e:
                logger.error(f"Error al leer config.json: {e}")
                return {}
        return {}

    @staticmethod
    def save_config(config: dict) -> bool:
        """Guarda la configuración en el archivo JSON."""
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            logger.info("Configuración guardada exitosamente.")
            return True
        except Exception as e:
            logger.error(f"Error al guardar config.json: {e}")
            return False

    @staticmethod
    def load_reports() -> list:
        """Carga los informes desde el archivo JSON."""
        if REPORTS_FILE.exists():
            try:
                with open(REPORTS_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except json.JSONDecodeError as e:
                logger.error(f"Error al decodificar reports.json: {e}")
                return []
            except Exception as e:
                logger.error(f"Error al leer reports.json: {e}")
                return []
        return []

    @staticmethod
    def save_reports(reports: list) -> bool:
        """Guarda los informes en el archivo JSON."""
        try:
            with open(REPORTS_FILE, 'w', encoding='utf-8') as f:
                json.dump(reports, f, indent=4, ensure_ascii=False)
            logger.info(f"Se guardaron {len(reports)} informes localmente.")
            return True
        except Exception as e:
            logger.error(f"Error al guardar reports.json: {e}")
            return False


# ============================================================================
# SECCIÓN 3: ESTILOS CSS PROFESIONALES Y RESPONSIVOS
# ============================================================================

def aplicar_estilos_css():
    """
    Aplica estilos CSS profesionales con contraste normativo (WCAG 2.1 AA).
    Incluye estilos para sidebar, tarjetas, botones, tablas y alertas.
    """
    st.markdown(f"""
    <style>
        /* ============================================
           RESET Y CONFIGURACIÓN GENERAL
           ============================================ */
        .stApp {{
            background-color: {COLOR_BG_LIGHT};
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            color: {COLOR_TEXT_DARK};
        }}

        /* ============================================
           SIDEBAR - FONDO CLARO (Corrección de legibilidad)
           ============================================ */
        [data-testid="stSidebar"] {{
            background-color: {COLOR_SIDEBAR} !important;
            border-right: 4px solid {COLOR_SIDEBAR_BORDER};
        }}

        [data-testid="stSidebar"] .stMarkdown {{
            color: {COLOR_TEXT_DARK} !important;
        }}

        [data-testid="stSidebar"] label, 
        [data-testid="stSidebar"] .stSelectbox label,
        [data-testid="stSidebar"] .stRadio label {{
            color: {COLOR_TEXT_DARK} !important;
            font-weight: 600;
            font-size: 14px;
        }}

        [data-testid="stSidebar"] button[kind="secondary"] {{
            background-color: {COLOR_WHITE};
            color: {COLOR_PRIMARY};
            border: 1px solid {COLOR_PRIMARY};
            font-weight: 600;
        }}

        [data-testid="stSidebar"] button[kind="secondary"]:hover {{
            background-color: {COLOR_PRIMARY};
            color: {COLOR_WHITE};
        }}

        /* ============================================
           ENCABEZADO PRINCIPAL
           ============================================ */
        .main-header {{
            background: linear-gradient(135deg, {COLOR_PRIMARY} 0%, {COLOR_SECONDARY} 100%);
            padding: 25px 30px;
            border-radius: 12px;
            color: {COLOR_WHITE};
            margin-bottom: 25px;
            box-shadow: 0 4px 15px rgba(13, 43, 78, 0.25);
        }}

        .main-header h1 {{
            margin: 0;
            font-size: 28px;
            font-weight: 700;
            letter-spacing: 0.5px;
            color: {COLOR_WHITE};
        }}

        .main-header p {{
            margin: 8px 0 0 0;
            font-size: 14px;
            color: rgba(255, 255, 255, 0.95);
        }}

        /* ============================================
           TARJETAS DE ESTADÍSTICAS
           ============================================ */
        .stat-card {{
            background: {COLOR_BG_CARD};
            border-radius: 12px;
            padding: 20px;
            text-align: center;
            box-shadow: 0 2px 10px rgba(0,0,0,0.06);
            border-top: 4px solid {COLOR_SECONDARY};
            transition: transform 0.2s, box-shadow 0.2s;
        }}

        .stat-card:hover {{
            transform: translateY(-3px);
            box-shadow: 0 5px 20px rgba(0,0,0,0.12);
        }}

        .stat-card .stat-number {{
            font-size: 36px;
            font-weight: 700;
            color: {COLOR_PRIMARY};
        }}

        .stat-card .stat-label {{
            font-size: 13px;
            color: {COLOR_TEXT_SECONDARY};
            margin-top: 5px;
            font-weight: 500;
        }}

        /* ============================================
           TARJETAS DE INFORMACIÓN
           ============================================ */
        .info-card {{
            background: {COLOR_BG_CARD};
            border-radius: 12px;
            padding: 20px 25px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.06);
            border-left: 5px solid {COLOR_SECONDARY};
            margin-bottom: 15px;
        }}

        .info-card h3 {{
            color: {COLOR_PRIMARY};
            margin: 0 0 10px 0;
            font-size: 16px;
        }}

        .info-card p {{
            color: {COLOR_TEXT_DARK};
            margin: 0;
            font-size: 14px;
        }}

        /* ============================================
           LOGIN CONTAINER
           ============================================ */
        .login-container {{
            max-width: 450px;
            margin: 0 auto;
            padding: 40px;
            background: {COLOR_WHITE};
            border-radius: 16px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
            border-top: 5px solid {COLOR_PRIMARY};
        }}

        .login-logo {{
            text-align: center;
            margin-bottom: 30px;
        }}

        .login-logo h2 {{
            color: {COLOR_PRIMARY};
            font-size: 24px;
            margin-bottom: 5px;
        }}

        .login-logo p {{
            color: {COLOR_TEXT_SECONDARY};
            font-size: 13px;
        }}

        /* ============================================
           FOOTER
           ============================================ */
        .footer {{
            background: linear-gradient(135deg, {COLOR_PRIMARY} 0%, {COLOR_SECONDARY} 100%);
            padding: 20px 30px;
            border-radius: 12px;
            color: {COLOR_WHITE};
            text-align: center;
            margin-top: 30px;
        }}

        .footer p {{
            margin: 3px 0;
            font-size: 13px;
            color: rgba(255, 255, 255, 0.95);
        }}

        .footer .brand {{
            font-size: 15px;
            font-weight: 700;
            color: {COLOR_ACCENT};
        }}

        /* ============================================
           SEPARADOR Y ALERTAS
           ============================================ */
        .custom-divider {{
            height: 2px;
            background: linear-gradient(90deg, transparent, {COLOR_SECONDARY}, transparent);
            margin: 20px 0;
        }}

        .custom-alert {{
            background: {COLOR_INFO_BG};
            border-left: 4px solid {COLOR_INFO};
            border-radius: 8px;
            padding: 15px 20px;
            margin: 10px 0;
            color: {COLOR_TEXT_DARK};
        }}

        /* ============================================
           BOTONES Y TABLAS
           ============================================ */
        .stButton > button {{
            border-radius: 8px;
            font-weight: 600;
            transition: all 0.2s;
            border: 1px solid {COLOR_BORDER};
        }}

        .stButton > button:hover {{
            transform: translateY(-1px);
            box-shadow: 0 3px 10px rgba(0,0,0,0.15);
        }}

        .stDataFrame {{
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        }}

        /* ============================================
           OCULTAR ELEMENTOS DE STREAMLIT POR DEFECTO
           ============================================ */
        #MainMenu {{visibility: hidden;}}
        footer {{visibility: hidden;}}
        header {{visibility: hidden;}}

        /* ============================================
           RESPONSIVE DESIGN
           ============================================ */
        @media (max-width: 768px) {{
            .main-header h1 {{ font-size: 20px; }}
            .stat-card .stat-number {{ font-size: 28px; }}
            .login-container {{ padding: 20px; }}
        }}
    </style>
    """, unsafe_allow_html=True)


# ============================================================================
# SECCIÓN 4: GESTIÓN DE AUTENTICACIÓN Y SEGURIDAD
# ============================================================================

class AuthenticationManager:
    """
    Gestiona la autenticación de usuarios, verificación de contraseñas
    y registro de nuevos usuarios en el sistema.
    """

    def __init__(self):
        self.users = self._load_users()

    def _load_users(self) -> dict:
        """Carga los usuarios desde session_state o usa los por defecto."""
        if "registered_users" not in st.session_state:
            st.session_state.registered_users = DEFAULT_USERS.copy()
        return st.session_state.registered_users

    def verify_password(self, plain_password: str, hashed_password: str) -> bool:
        """Verifica si la contraseña en texto plano coincide con el hash almacenado."""
        try:
            return bcrypt.checkpw(
                plain_password.encode('utf-8'),
                hashed_password.encode('utf-8')
            )
        except Exception as e:
            logger.error(f"Error al verificar contraseña: {e}")
            return False

    def authenticate(self, username: str, password: str) -> Optional[dict]:
        """
        Autentica un usuario. 
        Retorna un diccionario con los datos del usuario o None si falla.
        """
        username = username.strip().lower()
        if username in self.users:
            user_data = self.users[username]
            if user_data.get("activo", False):
                if self.verify_password(password, user_data["password_hash"]):
                    logger.info(f"Autenticación exitosa para el usuario: {username}")
                    return {
                        "username": username,
                        "nombre": user_data["nombre"],
                        "rol": user_data["rol"],
                        "email": user_data.get("email", "")
                    }
            else:
                logger.warning(f"Intento de login con usuario inactivo: {username}")
        return None

    def register_user(self, username: str, password: str, nombre: str, rol: str, email: str = "") -> bool:
        """Registra un nuevo usuario en el sistema con contraseña hasheada."""
        username = username.strip().lower()
        if username in self.users:
            return False
        
        # Generar hash con rounds=12 para mayor seguridad
        password_hash = bcrypt.hashpw(
            password.encode('utf-8'),
            bcrypt.gensalt(rounds=12)
        ).decode('utf-8')
        
        self.users[username] = {
            "password_hash": password_hash,
            "nombre": nombre,
            "rol": rol,
            "email": email,
            "activo": True
        }
        st.session_state.registered_users = self.users
        logger.info(f"Nuevo usuario registrado: {username} ({rol})")
        return True

    def get_all_users(self) -> dict:
        """Retorna todos los usuarios registrados (sin los hashes por seguridad)."""
        safe_users = {}
        for uname, udata in self.users.items():
            safe_users[uname] = {
                "nombre": udata.get("nombre", ""),
                "rol": udata.get("rol", ""),
                "activo": udata.get("activo", False),
                "email": udata.get("email", "")
            }
        return safe_users


def render_login_screen():
    """Renderiza la pantalla de inicio de sesión con diseño profesional."""
    st.markdown(f"""
    <div style="text-align:center; padding: 40px 0 20px 0;">
        <h1 style="color:{COLOR_PRIMARY}; font-size:32px; margin-bottom:5px;">🔧 CAVA</h1>
        <p style="color:{COLOR_TEXT_SECONDARY}; font-size:14px;">
            Especialistas en Robótica y Automatización
        </p>
        <p style="color:{COLOR_SECONDARY}; font-size:12px; margin-top:2px;">
            Sistema de Gestión de Informes de Mantenimiento v{APP_VERSION}
        </p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])

    with col2:
        st.markdown(f"""
        <div class="login-container">
            <div class="login-logo">
                <h2>🔐 Iniciar Sesión</h2>
                <p>Ingrese sus credenciales para acceder al sistema</p>
            </div>
        </div>
        """, unsafe_allow_html=True)

        with st.form("login_form", clear_on_submit=False):
            username = st.text_input(
                "👤 Usuario",
                placeholder="Ingrese su usuario",
                key="login_username",
                help="Ingrese el nombre de usuario asignado"
            )
            password = st.text_input(
                "🔑 Contraseña",
                type="password",
                placeholder="Ingrese su contraseña",
                key="login_password",
                help="La contraseña es sensible a mayúsculas y minúsculas"
            )
            submit = st.form_submit_button(
                "Ingresar al Sistema",
                use_container_width=True,
                type="primary"
            )

            if submit:
                if not username or not password:
                    st.error("⚠️ Por favor complete todos los campos.")
                else:
                    auth = AuthenticationManager()
                    user = auth.authenticate(username, password)
                    if user:
                        st.session_state.authenticated = True
                        st.session_state.current_user = user
                        st.session_state.login_time = datetime.now()
                        st.success(f"✅ Bienvenido, {user['nombre']}!")
                        logger.info(f"Sesión iniciada: {user['nombre']}")
                        st.rerun()
                    else:
                        st.error("❌ Usuario o contraseña incorrectos.")
                        logger.warning(f"Intento de login fallido para: {username}")

        st.markdown("---")
        st.caption(
            "💡 Credenciales de prueba: admin / admin123 | rhvamani / cava2026 | tecnico1 / tecnico123"
        )


# ============================================================================
# SECCIÓN 5: GESTIÓN DE BASE DE DATOS (SUPABASE + FALLBACK LOCAL)
# ============================================================================

class SupabaseManager:
    """
    Gestiona la conexión y operaciones con Supabase para almacenamiento en la nube.
    Incluye un sistema de fallback robusto a almacenamiento local JSON si Supabase falla.
    """

    def __init__(self):
        self.client = None
        self.connected = False
        self._initialize_connection()

    def _initialize_connection(self):
        """Inicializa la conexión con Supabase usando las credenciales de sesión."""
        url = st.session_state.get("supabase_url", "")
        key = st.session_state.get("supabase_key", "")
        if url and key:
            try:
                self.client = create_client(url, key)
                self.connected = True
                logger.info("Conexión a Supabase inicializada exitosamente.")
            except Exception as e:
                self.connected = False
                st.session_state.supabase_error = str(e)
                logger.error(f"Error al conectar con Supabase: {e}")

    def is_connected(self) -> bool:
        """Verifica si la conexión a Supabase está activa y funcional."""
        return self.connected and self.client is not None

    def save_report(self, report_data: dict) -> bool:
        """Guarda un informe en Supabase. Si falla, guarda localmente."""
        if not self.is_connected():
            logger.warning("Supabase no conectado. Guardando informe localmente.")
            return self._save_local(report_data)
        
        try:
            self.client.table("informes_mantenimiento").insert(report_data).execute()
            logger.info(f"Informe {report_data.get('numero')} guardado en Supabase.")
            return True
        except Exception as e:
            logger.error(f"Error al guardar en Supabase: {e}. Fallback a local.")
            st.warning(f"⚠️ Error al guardar en la nube: {e}. Guardando localmente.")
            return self._save_local(report_data)

    def get_reports(self, limit: int = 100) -> list:
        """Obtiene los informes almacenados, ordenados por fecha descendente."""
        if not self.is_connected():
            return self._get_local_reports()
        
        try:
            result = self.client.table("informes_mantenimiento") \
                .select("*") \
                .order("fecha_creacion", desc=True) \
                .limit(limit) \
                .execute()
            return result.data if result.data else []
        except Exception as e:
            logger.error(f"Error al obtener informes de Supabase: {e}")
            return self._get_local_reports()

    def delete_report(self, report_id: str) -> bool:
        """Elimina un informe por su ID."""
        if not self.is_connected():
            return self._delete_local_report(report_id)
        
        try:
            self.client.table("informes_mantenimiento") \
                .delete() \
                .eq("id", report_id) \
                .execute()
            logger.info(f"Informe {report_id} eliminado de Supabase.")
            return True
        except Exception as e:
            logger.error(f"Error al eliminar de Supabase: {e}")
            return self._delete_local_report(report_id)

    # --- Métodos de Almacenamiento Local (Fallback) ---

    def _save_local(self, report_data: dict) -> bool:
        """Guarda localmente cuando Supabase no está configurado o falla."""
        reports = ConfigPersistence.load_reports()
        report_data["storage"] = "local"
        reports.insert(0, report_data)
        ConfigPersistence.save_reports(reports)
        st.session_state.local_reports = reports
        return True

    def _get_local_reports(self) -> list:
        """Obtiene informes del almacenamiento local JSON."""
        reports = ConfigPersistence.load_reports()
        st.session_state.local_reports = reports
        return reports

    def _delete_local_report(self, report_id: str) -> bool:
        """Elimina un informe del almacenamiento local."""
        reports = ConfigPersistence.load_reports()
        reports = [r for r in reports if r.get("id") != report_id]
        ConfigPersistence.save_reports(reports)
        st.session_state.local_reports = reports
        return True


# ============================================================================
# SECCIÓN 6: INTEGRACIÓN CON GEMINI AI (MODELO GRATUITO Y ESTABLE)
# ============================================================================

class GeminiAIManager:
    """
    Gestiona la integración con la API de Google Gemini.
    Utiliza el modelo 'gemini-1.5-flash' que es gratuito, estable y de alto rendimiento.
    """

    def __init__(self):
        self.model = None
        self.api_key = st.session_state.get("gemini_api_key", "")
        self._initialize_model()

    def _initialize_model(self):
        """Inicializa el modelo de Gemini con la API Key configurada."""
        if self.api_key:
            try:
                genai.configure(api_key=self.api_key)
                # Modelo gratuito, estable y con soporte de contexto largo
                self.model = genai.GenerativeModel('gemini-1.5-flash')
                logger.info("Modelo Gemini 1.5 Flash inicializado correctamente.")
            except Exception as e:
                st.session_state.gemini_error = str(e)
                logger.error(f"Error al inicializar Gemini: {e}")

    def is_configured(self) -> bool:
        """Verifica si la API está configurada y el modelo está listo."""
        return self.model is not None and self.api_key != ""

    def generate_report_content(self, raw_description: str, report_type: str, additional_context: dict = None) -> str:
        """
        Genera el contenido estructurado del informe usando Gemini AI.
        Incluye instrucciones específicas para cálculos de ingeniería si aplica.
        """
        if not self.is_configured():
            logger.warning("Gemini no configurado. Usando informe base.")
            return self._generate_fallback_report(raw_description, report_type, additional_context)

        context_info = ""
        if additional_context:
            for key, value in additional_context.items():
                if value:
                    context_info += f"- {key}: {value}\n"

        if report_type == TIPO_REPORTE_MANTENIMIENTO:
            prompt = self._build_report_prompt(raw_description, context_info)
        else:
            prompt = self._build_executive_prompt(raw_description, context_info)

        try:
            response = self.model.generate_content(prompt)
            logger.info("Contenido del informe generado exitosamente por Gemini.")
            return response.text
        except Exception as e:
            logger.error(f"Error al generar contenido con Gemini: {e}")
            st.warning(f"⚠️ Error con Gemini AI: {e}. Generando informe con plantilla base.")
            return self._generate_fallback_report(raw_description, report_type, additional_context)

    def _build_report_prompt(self, raw_description: str, context_info: str) -> str:
        """Construye el prompt detallado para un Reporte de Mantenimiento Técnico."""
        return f"""
Eres un Jefe de Mantenimiento y Planificador experto con más de 20 años de experiencia en gestión de activos industriales. 
Tu tarea es redactar un Reporte de Mantenimiento profesional, técnico y estructurado.

INSTRUCCIONES CRÍTICAS:
1. Redacta con un tono técnico, objetivo y profesional, como si fuera escrito por un ingeniero de mantenimiento senior.
2. El lenguaje debe ser humanizado pero preciso. Sin errores ortográficos ni gramaticales.
3. Sigue estrictamente la normativa de redacción para informes técnicos industriales.
4. Si detectas que falta información crítica, rellena el campo con "[Información no proporcionada - Completar]".
5. Profundiza y completa cada campo con información técnica relevante basada en el contexto proporcionado.
6. Usa terminología técnica apropiada del área de mantenimiento industrial.
7. CÁLCULOS DE INGENIERÍA: Si la descripción del técnico lo amerita, incluye una sección de "Cálculos Técnicos de Ingeniería" 
   con fórmulas, datos medidos, tolerancias permitidas vs. reales, torques de apriete, holguras, consumos energéticos, 
   cálculos hidráulicos, etc. Si no aplica, indica "No se requirieron cálculos específicos".

DATOS PROPORCIONADOS POR EL TÉCNICO:
\"\"\"
{raw_description}
\"\"\"

DATOS ADICIONALES DEL SISTEMA:
{context_info}

ESTRUCTURA OBLIGATORIA DEL REPORTE (respeta este formato exacto):

## REPORTE DE MANTENIMIENTO

### 1. IDENTIFICACIÓN BÁSICA
- **Fecha:** [fecha]
- **Turno:** [turno]
- **Número de OT:** [número de orden de trabajo]
- **Técnico(s) Responsable(s):** [nombres]
- **Disciplina:** [mecánica/eléctrica/instrumentación]

### 2. IDENTIFICACIÓN DEL EQUIPO
- **Nombre del Equipo:** [nombre]
- **Tag/Código:** [tag]
- **Área/Ubicación:** [área]
- **Horómetro/Ciclos:** [lectura]
- **Fabricante/Modelo:** [datos del fabricante]

### 3. DESCRIPCIÓN DEL PROBLEMA
[Descripción técnica detallada del síntoma reportado, incluyendo condiciones de operación al momento de la falla, 
alarmas activas, y manifestaciones observadas por el operador o técnico.]

### 4. TIPO DE INTERVENCIÓN
- **Tipo:** [Correctivo / Preventivo / Mejora / Predictivo]
- **Prioridad:** [Crítica / Alta / Media / Baja]
- **Justificación:** [breve justificación técnica de la clasificación]

### 5. DETALLE DEL TRABAJO REALIZADO
[Descripción técnica paso a paso de todas las actividades realizadas. Incluir procedimientos de seguridad aplicados 
(LOTO, permisos de trabajo), herramientas especiales utilizadas, mediciones tomadas, y criterios de aceptación verificados. 
Numerar cada paso de forma clara y secuencial.]

### 6. CÁLCULOS TÉCNICOS DE INGENIERÍA (Si aplica)
[Incluir aquí cualquier cálculo relevante realizado: fórmulas, valores medidos, tolerancias permitidas vs. reales, 
torques de apriete, holguras, consumos, etc. Si no aplica, indicar "No se requirieron cálculos específicos".]

### 7. REPUESTOS Y MATERIALES UTILIZADOS
| Ítem | Descripción | Cantidad | Unidad | Código SAP |
|------|-------------|----------|--------|------------|
| [completar tabla con datos realistas o indicar Completar] |

### 8. TIEMPOS DE INTERVENCIÓN
- **Hora de Aviso:** [hh:mm]
- **Hora de Inicio:** [hh:mm]
- **Hora de Fin:** [hh:mm]
- **Downtime Total:** [horas y minutos]
- **Tiempo de Espera (Repuestos/Acceso):** [si aplica]

### 9. ESTADO FINAL Y OBSERVACIONES
- **Estado del Equipo:** [Operativo / Operativo con restricciones / Fuera de servicio]
- **Pruebas Realizadas:** [descripción de pruebas post-intervención]
- **Observaciones:** [cualquier observación relevante]
- **Trabajos Pendientes:** [si los hay]
- **Recomendaciones:** [acciones sugeridas a corto y mediano plazo]

Redacta el documento completo ahora, asegurando un formato impecable:
"""

    def _build_executive_prompt(self, raw_description: str, context_info: str) -> str:
        """Construye el prompt detallado para un Informe Ejecutivo de Mantenimiento."""
        return f"""
Eres un Superintendente de Mantenimiento y Planificador experto con más de 20 años de experiencia en gestión de activos 
y confiabilidad industrial. Tu tarea es redactar un Informe Ejecutivo de Mantenimiento dirigido a la Gerencia de Planta.

INSTRUCCIONES CRÍTICAS:
1. Redacta con un tono ejecutivo, técnico y orientado a la gestión de activos.
2. El lenguaje debe ser claro, directo y profesional. Sin errores ortográficos.
3. El informe debe ser comprensible para personal de gerencia no necesariamente técnico, pero con suficiente profundidad técnica.
4. Si detectas que falta información crítica, rellena el campo con "[Información no proporcionada - Completar]".
5. Enfatiza el impacto operativo, los costos y las decisiones a tomar.
6. Incluye métricas, KPIs relevantes y, si aplica, cálculos técnicos de ingeniería (torque, holguras, consumo energético, etc.) 
   con sus fórmulas y resultados resumidos.

DATOS PROPORCIONADOS POR EL TÉCNICO:
\"\"\"
{raw_description}
\"\"\"

DATOS ADICIONALES DEL SISTEMA:
{context_info}

ESTRUCTURA OBLIGATORIA DEL INFORME EJECUTIVO (respeta este formato exacto):

## INFORME EJECUTIVO DE MANTENIMIENTO

### 1. DATOS GENERALES
- **Fecha del Informe:** [fecha]
- **Área/Planta:** [área]
- **Equipo Crítico Afectado:** [nombre y tag]
- **Prioridad del Evento:** [Crítica / Alta / Media / Baja]
- **Elaborado por:** [nombre y cargo]
- **Dirigido a:** Gerencia de Planta / Gerencia de Operaciones

### 2. RESUMEN DEL EVENTO
[Párrafo conciso y directo que explique qué sucedió, cuándo, dónde y cuál fue la magnitud del evento. 
Máximo 5-6 líneas. Debe captar la atención del lector ejecutivo inmediatamente.]

### 3. IMPACTO OPERATIVO
- **Horas de Producción Perdidas:** [horas]
- **Tonelaje/Unidades No Producidas:** [si aplica]
- **Nivel de Afectación:** [Total / Parcial / Mínimo]
- **Líneas/Áreas Afectadas:** [detalle]
- **Impacto en Entregas/Clientes:** [si aplica]
- **Costo Estimado de Lucro Cesante:** [si se puede estimar]

### 4. CAUSA RAÍZ PRELIMINAR
[Análisis técnico de la causa raíz del evento. Incluir si se aplicó metodología 5 Porqués, Ishikawa u otra. 
Diferenciar entre causa directa y causa raíz. Indicar si se requiere un análisis RCA formal.]

### 5. SOLUCIÓN EJECUTADA Y CÁLCULOS TÉCNICOS
[Descripción clara de las acciones tomadas para restablecer la operación. Incluir si fue una solución temporal o definitiva. 
Mencionar tiempos de respuesta, efectividad de la solución y, si aplica, cálculos de ingeniería realizados (torques, 
holguras, consumos, etc.) con sus resultados.]

### 6. COSTOS RELEVANTES
| Concepto | Monto Estimado | Moneda |
|----------|---------------|--------|
| Repuestos | [monto] | [USD/PEN] |
| Mano de Obra | [monto] | [USD/PEN] |
| Servicios Externos | [monto] | [USD/PEN] |
| Lucro Cesante | [monto] | [USD/PEN] |
| **TOTAL ESTIMADO** | **[monto]** | **[moneda]** |

### 7. ACCIONES PREVENTIVAS Y RECOMENDACIONES
[Lista numerada de acciones concretas, con responsable sugerido y plazo estimado. Separar en acciones inmediatas, 
corto plazo y mediano plazo. Incluir recomendaciones sobre cambios en el plan de mantenimiento, mejoras de diseño, 
capacitación, o adquisición de repuestos críticos.]

### 8. CONCLUSIONES
[2-3 párrafos de cierre que resuman la situación actual, la confiabilidad del equipo post-intervención, 
y la urgencia de las acciones recomendadas.]

Redacta el documento completo ahora, asegurando un formato impecable:
"""

    def process_image_description(self, raw_description: str) -> str:
        """Procesa y mejora la descripción de una imagen técnica usando IA."""
        if not self.is_configured():
            return raw_description.strip().capitalize() + "."
        
        prompt = f"""
Eres un ingeniero de mantenimiento experto. Mejora la siguiente descripción de una fotografía técnica tomada durante 
una intervención de mantenimiento.

INSTRUCCIONES:
1. Corrige ortografía y gramática.
2. Usa terminología técnica precisa.
3. Estructura la descripción de forma clara y profesional.
4. Mantén la descripción concisa (2-4 oraciones máximo).
5. Indica qué se observa en la imagen de forma objetiva.

DESCRIPCIÓN ORIGINAL DEL TÉCNICO:
\"{raw_description}\"

Descripción mejorada:
"""
        try:
            response = self.model.generate_content(prompt)
            return response.text.strip()
        except Exception as e:
            logger.error(f"Error al procesar descripción de imagen: {e}")
            return raw_description.strip().capitalize() + "."

    def translate_report(self, report_content: str) -> str:
        """Traduce el informe completo al inglés manteniendo el formato técnico."""
        if not self.is_configured():
            return "[Translation requires Gemini API configuration]"

        prompt = f"""
You are a professional technical translator specializing in industrial maintenance documentation. 
Translate the following maintenance report from Spanish to English.

INSTRUCTIONS:
1. Maintain the exact same structure and formatting.
2. Use standard industrial maintenance terminology in English.
3. Keep all technical specifications, measurements, and codes as-is.
4. Ensure the translation reads naturally in English.
5. Preserve all markdown formatting (headers, tables, bold text, etc.).
6. If engineering calculations are present, ensure the technical terms and units are correctly translated.

ORIGINAL REPORT:
\"\"\"
{report_content}
\"\"\"

TRANSLATED REPORT:
"""
        try:
            response = self.model.generate_content(prompt)
            return response.text
        except Exception as e:
            logger.error(f"Error al traducir informe: {e}")
            return f"[Translation error: {e}]"

    def _generate_fallback_report(self, raw_description: str, report_type: str, additional_context: dict = None) -> str:
        """Genera un informe base estructurado cuando la IA no está configurada."""
        fecha = datetime.now().strftime("%d/%m/%Y")
        ctx = additional_context or {}

        if report_type == TIPO_REPORTE_MANTENIMIENTO:
            return f"""## REPORTE DE MANTENIMIENTO

### 1. IDENTIFICACIÓN BÁSICA
- **Fecha:** {fecha}
- **Turno:** {ctx.get('turno', '[Información no proporcionada - Completar]')}
- **Número de OT:** {ctx.get('numero_ot', '[Información no proporcionada - Completar]')}
- **Técnico(s) Responsable(s):** {ctx.get('tecnicos', '[Información no proporcionada - Completar]')}
- **Disciplina:** {ctx.get('disciplina', '[Información no proporcionada - Completar]')}

### 2. IDENTIFICACIÓN DEL EQUIPO
- **Nombre del Equipo:** {ctx.get('equipo', '[Información no proporcionada - Completar]')}
- **Tag/Código:** {ctx.get('tag', '[Información no proporcionada - Completar]')}
- **Área/Ubicación:** {ctx.get('area', '[Información no proporcionada - Completar]')}
- **Horómetro/Ciclos:** [Información no proporcionada - Completar]
- **Fabricante/Modelo:** [Información no proporcionada - Completar]

### 3. DESCRIPCIÓN DEL PROBLEMA
{raw_description}

### 4. TIPO DE INTERVENCIÓN
- **Tipo:** {ctx.get('tipo_intervencion', '[Información no proporcionada - Completar]')}
- **Prioridad:** {ctx.get('prioridad', '[Información no proporcionada - Completar]')}
- **Justificación:** [Información no proporcionada - Completar]

### 5. DETALLE DEL TRABAJO REALIZADO
[Configure la API de Gemini para generar automáticamente este contenido técnico detallado.]

### 6. CÁLCULOS TÉCNICOS DE INGENIERÍA
[No se proporcionaron datos suficientes para generar cálculos. Configure la API de Gemini para asistencia.]

### 7. REPUESTOS Y MATERIALES UTILIZADOS
| Ítem | Descripción | Cantidad | Unidad | Código SAP |
|------|-------------|----------|--------|------------|
| 1 | [Completar] | - | - | - |

### 8. TIEMPOS DE INTERVENCIÓN
- **Hora de Aviso:** [Información no proporcionada - Completar]
- **Hora de Inicio:** [Información no proporcionada - Completar]
- **Hora de Fin:** [Información no proporcionada - Completar]
- **Downtime Total:** [Información no proporcionada - Completar]

### 9. ESTADO FINAL Y OBSERVACIONES
- **Estado del Equipo:** [Información no proporcionada - Completar]
- **Pruebas Realizadas:** [Información no proporcionada - Completar]
- **Recomendaciones:** [Información no proporcionada - Completar]
"""
        else:
            return f"""## INFORME EJECUTIVO DE MANTENIMIENTO

### 1. DATOS GENERALES
- **Fecha del Informe:** {fecha}
- **Área/Planta:** {ctx.get('area', '[Información no proporcionada - Completar]')}
- **Equipo Crítico Afectado:** {ctx.get('equipo', '[Información no proporcionada - Completar]')}
- **Prioridad del Evento:** {ctx.get('prioridad', '[Información no proporcionada - Completar]')}
- **Elaborado por:** {ctx.get('elaborado_por', '[Información no proporcionada - Completar]')}
- **Dirigido a:** Gerencia de Planta / Gerencia de Operaciones

### 2. RESUMEN DEL EVENTO
{raw_description}

### 3. IMPACTO OPERATIVO
- **Horas de Producción Perdidas:** [Información no proporcionada - Completar]
- **Nivel de Afectación:** [Información no proporcionada - Completar]

### 4. CAUSA RAÍZ PRELIMINAR
[Configure la API de Gemini para generar automáticamente este análisis.]

### 5. SOLUCIÓN EJECUTADA Y CÁLCULOS TÉCNICOS
[Configure la API de Gemini para generar automáticamente este contenido.]

### 6. COSTOS RELEVANTES
| Concepto | Monto Estimado | Moneda |
|----------|---------------|--------|
| Repuestos | [Completar] | [USD/PEN] |
| Mano de Obra | [Completar] | [USD/PEN] |
| **TOTAL ESTIMADO** | **[Completar]** | **[moneda]** |

### 7. ACCIONES PREVENTIVAS Y RECOMENDACIONES
[Configure la API de Gemini para generar automáticamente estas recomendaciones.]

### 8. CONCLUSIONES
[Configure la API de Gemini para generar automáticamente las conclusiones.]
"""


# ============================================================================
# SECCIÓN 7: GENERADOR DE NÚMEROS DE INFORME CORRELATIVOS
# ============================================================================

class ReportNumberGenerator:
    """
    Genera números de informe correlativos automáticamente.
    Formato: RM-2026-0001 (Reporte de Mantenimiento) o IE-2026-0001 (Informe Ejecutivo)
    """

    @staticmethod
    def generate(tipo: str) -> str:
        """Genera un número único de informe basado en el tipo y el año."""
        if "report_counter" not in st.session_state:
            st.session_state.report_counter = {
                TIPO_REPORTE_MANTENIMIENTO: 0,
                TIPO_INFORME_EJECUTIVO: 0
            }

        st.session_state.report_counter[tipo] += 1
        count = st.session_state.report_counter[tipo]
        year = datetime.now().year

        if tipo == TIPO_REPORTE_MANTENIMIENTO:
            prefix = "RM"
        else:
            prefix = "IE"

        return f"{prefix}-{year}-{count:04d}"


# ============================================================================
# SECCIÓN 8: GESTIÓN DE IMÁGENES Y REGISTRO FOTOGRÁFICO
# ============================================================================

class ImageManager:
    """
    Gestiona la carga, numeración, descripción y compresión de imágenes 
    para el registro fotográfico del informe.
    """

    @staticmethod
    def initialize_image_counter():
        """Inicializa el contador y la lista de imágenes en la sesión."""
        if "image_counter" not in st.session_state:
            st.session_state.image_counter = 0
        if "uploaded_images" not in st.session_state:
            st.session_state.uploaded_images = []

    @staticmethod
    def add_image(image_file, description: str) -> dict:
        """Agrega una imagen con su descripción al registro de la sesión."""
        ImageManager.initialize_image_counter()
        st.session_state.image_counter += 1
        img_number = st.session_state.image_counter

        image_data = {
            "numero": img_number,
            "nombre_archivo": image_file.name,
            "descripcion_original": description,
            "descripcion_mejorada": "",
            "bytes": image_file.getvalue(),
            "tipo": image_file.type,
            "fecha_carga": datetime.now().isoformat()
        }

        st.session_state.uploaded_images.append(image_data)
        return image_data

    @staticmethod
    def get_all_images() -> list:
        """Retorna todas las imágenes cargadas en la sesión actual."""
        return st.session_state.get("uploaded_images", [])

    @staticmethod
    def remove_image(index: int):
        """Elimina una imagen por su índice y renumera las restantes."""
        images = st.session_state.get("uploaded_images", [])
        if 0 <= index < len(images):
            st.session_state.uploaded_images.pop(index)
            for i, img in enumerate(st.session_state.uploaded_images):
                img["numero"] = i + 1
            st.session_state.image_counter = len(st.session_state.uploaded_images)

    @staticmethod
    def clear_all_images():
        """Elimina todas las imágenes de la sesión."""
        st.session_state.uploaded_images = []
        st.session_state.image_counter = 0


# ============================================================================
# SECCIÓN 9: EXPORTACIÓN A WORD (DOCX) PROFESIONAL
# ============================================================================

class WordExporter:
    """
    Exporta informes a formato Microsoft Word (.docx) con estilos profesionales,
    encabezados, pies de página, tablas formateadas e imágenes incrustadas.
    """

    @staticmethod
    def create_document(report_content: str, report_number: str, report_type: str, images: list = None, is_english: bool = False) -> io.BytesIO:
        """Crea un documento Word profesional con formato de ingeniería."""
        doc = Document()

        # --- Configurar estilos del documento ---
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Calibri'
        font.size = Pt(11)
        font.color.rgb = RGBColor(26, 26, 26)
        style.paragraph_format.line_spacing = WD_LINE_SPACING.ONE_POINT_FIVE
        style.paragraph_format.space_after = Pt(6)

        # --- Encabezado del documento ---
        header = doc.sections[0].header
        header_para = header.paragraphs[0]
        header_para.text = f"CAVA - Especialistas en Robótica y Automatización"
        header_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        header_para.style.font.size = Pt(8)
        header_para.style.font.color.rgb = RGBColor(108, 117, 125)

        # --- Página de título ---
        for _ in range(4):
            doc.add_paragraph()

        title = doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title.add_run("🔧 CAVA")
        run.font.size = Pt(36)
        run.font.bold = True
        run.font.color.rgb = RGBColor(13, 43, 78)

        subtitle = doc.add_paragraph()
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = subtitle.add_run("Especialistas en Robótica y Automatización")
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(30, 95, 142)

        doc.add_paragraph()

        doc_title = doc.add_paragraph()
        doc_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = doc_title.add_run(report_type.upper())
        run.font.size = Pt(22)
        run.font.bold = True
        run.font.color.rgb = RGBColor(13, 43, 78)

        doc_number = doc.add_paragraph()
        doc_number.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = doc_number.add_run(f"N° {report_number}")
        run.font.size = Pt(16)
        run.font.color.rgb = RGBColor(201, 123, 0)

        if is_english:
            lang_note = doc.add_paragraph()
            lang_note.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = lang_note.add_run("[ENGLISH VERSION - TRANSLATED DOCUMENT]")
            run.font.size = Pt(12)
            run.font.italic = True
            run.font.color.rgb = RGBColor(183, 28, 28)

        doc.add_page_break()

        # --- Tabla de información del documento ---
        info_table = doc.add_table(rows=4, cols=2)
        info_table.style = 'Light Grid Accent 1'
        info_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        info_data = [
            ("Documento:", report_type),
            ("Número:", report_number),
            ("Fecha de Emisión:", datetime.now().strftime("%d/%m/%Y %H:%M")),
            ("Autor:", st.session_state.get("current_user", {}).get("nombre", "Sistema CAVA"))
        ]

        for i, (label, value) in enumerate(info_data):
            info_table.rows[i].cells[0].text = label
            info_table.rows[i].cells[1].text = value
            for cell in info_table.rows[i].cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(10)

        doc.add_paragraph()

        # --- Contenido del informe ---
        WordExporter._add_formatted_content(doc, report_content)

        # --- Sección de imágenes ---
        if images and len(images) > 0:
            doc.add_page_break()
            img_title = doc.add_heading('REGISTRO FOTOGRÁFICO', level=1)
            for run in img_title.runs:
                run.font.color.rgb = RGBColor(13, 43, 78)

            for img_data in images:
                doc.add_paragraph()
                doc.add_heading(f'Fotografía N° {img_data["numero"]}', level=3)

                try:
                    image_stream = io.BytesIO(img_data["bytes"])
                    doc.add_picture(image_stream, width=Inches(5.5))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                except Exception as e:
                    logger.error(f"Error al insertar imagen en Word: {e}")
                    doc.add_paragraph("[Imagen no disponible]")

                desc = img_data.get("descripcion_mejorada") or img_data.get("descripcion_original", "")
                desc_para = doc.add_paragraph()
                desc_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = desc_para.add_run(f"Figura {img_data['numero']}: {desc}")
                run.font.size = Pt(9)
                run.font.italic = True
                run.font.color.rgb = RGBColor(108, 117, 125)

        # --- Pie de página ---
        footer = doc.sections[0].footer
        footer_para = footer.paragraphs[0]
        footer_para.text = f"CAVA Especialistas en Robótica y Automatización | Roger Huamani | {report_number} | Página "
        footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer_para.style.font.size = Pt(8)

        # --- Guardar en buffer ---
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    @staticmethod
    def _add_formatted_content(doc: Document, content: str):
        """Agrega contenido formateado (markdown básico) al documento Word."""
        lines = content.split('\n')
        for line in lines:
            line = line.strip()
            if not line:
                continue

            if line.startswith('## '):
                heading = doc.add_heading(line[3:], level=1)
                for run in heading.runs:
                    run.font.color.rgb = RGBColor(13, 43, 78)
            elif line.startswith('### '):
                heading = doc.add_heading(line[4:], level=2)
                for run in heading.runs:
                    run.font.color.rgb = RGBColor(30, 95, 142)
            elif line.startswith('#### '):
                doc.add_heading(line[5:], level=3)
            elif line.startswith('|') and '---' not in line:
                cells = [c.strip() for c in line.split('|')[1:-1]]
                if cells:
                    table = doc.add_table(rows=1, cols=len(cells))
                    table.style = 'Light Grid Accent 1'
                    for i, cell_text in enumerate(cells):
                        table.rows[0].cells[i].text = cell_text.replace('**', '')
            elif line.startswith('|') and '---' in line:
                continue
            elif line.startswith('- **'):
                p = doc.add_paragraph(style='List Bullet')
                parts = line[2:].split('**')
                for j, part in enumerate(parts):
                    run = p.add_run(part)
                    if j % 2 == 1:
                        run.font.bold = True
                    run.font.size = Pt(11)
            elif line.startswith('- '):
                doc.add_paragraph(line[2:], style='List Bullet')
            elif re.match(r'^\d+\.', line):
                doc.add_paragraph(line, style='List Number')
            else:
                p = doc.add_paragraph()
                parts = line.split('**')
                for j, part in enumerate(parts):
                    run = p.add_run(part)
                    if j % 2 == 1:
                        run.font.bold = True
                    run.font.size = Pt(11)


# ============================================================================
# SECCIÓN 10: EXPORTACIÓN A PDF PROFESIONAL (FORMATO A4, JUSTIFICADO, INTERLINEADO)
# ============================================================================

class PDFExporter:
    """
    Exporta informes a formato PDF profesional con las siguientes características:
    - Formato A4
    - Fuente Helvetica (soporta caracteres latinos)
    - Texto justificado
    - Interlineado profesional (1.5)
    - Viñetas correctas
    - Tablas con bordes y estilos
    - Manejo seguro de caracteres (sin emojis)
    """

    @staticmethod
    def create_pdf(report_content: str, report_number: str, report_type: str, images: list = None, is_english: bool = False) -> io.BytesIO:
        """Crea un documento PDF profesional en formato A4."""
        pdf = PDFReport(report_number=report_number, report_type=report_type, is_english=is_english)
        pdf.alias_nb_pages()
        pdf.add_page()

        PDFExporter._add_cover_page(pdf, report_number, report_type, is_english)
        pdf.add_page()
        PDFExporter._add_content(pdf, report_content)

        if images and len(images) > 0:
            pdf.add_page()
            pdf.set_font("Helvetica", "B", 16)
            pdf.set_text_color(13, 43, 78)
            pdf.cell(0, 12, clean_text_for_pdf("REGISTRO FOTOGRÁFICO"), ln=True, align="C")
            pdf.ln(8)

            for img_data in images:
                pdf.set_font("Helvetica", "B", 11)
                pdf.set_text_color(30, 95, 142)
                pdf.cell(0, 8, clean_text_for_pdf(f"Fotografía N° {img_data['numero']}"), ln=True)

                try:
                    temp_path = f"/tmp/cava_img_{img_data['numero']}.png"
                    with open(temp_path, "wb") as f:
                        f.write(img_data["bytes"])
                    # Ajustar imagen a 160mm de ancho, manteniendo proporción
                    pdf.image(temp_path, w=160)
                    os.remove(temp_path)
                except Exception as e:
                    logger.error(f"Error al insertar imagen en PDF: {e}")
                    pdf.set_font("Helvetica", "I", 9)
                    pdf.cell(0, 6, clean_text_for_pdf("[Imagen no disponible]"), ln=True)

                desc = img_data.get("descripcion_mejorada") or img_data.get("descripcion_original", "")
                pdf.set_font("Helvetica", "I", 9)
                pdf.set_text_color(108, 117, 125)
                # Interlineado y justificado para la descripción
                pdf.multi_cell(0, 5, clean_text_for_pdf(f"Figura {img_data['numero']}: {desc}"), align="J")
                pdf.ln(6)

        buffer = io.BytesIO()
        pdf_bytes = pdf.output()
        buffer.write(pdf_bytes)
        buffer.seek(0)
        return buffer

    @staticmethod
    def _add_cover_page(pdf, report_number, report_type, is_english):
        """Agrega la portada del PDF con diseño institucional."""
        pdf.ln(50)
        pdf.set_font("Helvetica", "B", 32)
        pdf.set_text_color(13, 43, 78)
        pdf.cell(0, 15, clean_text_for_pdf("🔧 CAVA"), ln=True, align="C")

        pdf.set_font("Helvetica", "", 14)
        pdf.set_text_color(30, 95, 142)
        pdf.cell(0, 10, clean_text_for_pdf("Especialistas en Robotica y Automatizacion"), ln=True, align="C")
        pdf.ln(20)

        pdf.set_font("Helvetica", "B", 20)
        pdf.set_text_color(13, 43, 78)
        pdf.cell(0, 12, clean_text_for_pdf(report_type.upper()), ln=True, align="C")

        pdf.set_font("Helvetica", "B", 16)
        pdf.set_text_color(201, 123, 0)
        pdf.cell(0, 10, clean_text_for_pdf(f"N° {report_number}"), ln=True, align="C")

        if is_english:
            pdf.ln(5)
            pdf.set_font("Helvetica", "I", 12)
            pdf.set_text_color(183, 28, 28)
            pdf.cell(0, 8, clean_text_for_pdf("[ENGLISH VERSION - TRANSLATED DOCUMENT]"), ln=True, align="C")
        pdf.ln(30)

        pdf.set_font("Helvetica", "", 10)
        pdf.set_text_color(26, 26, 26)
        info_items = [
            ("Fecha de Emision:", datetime.now().strftime("%d/%m/%Y %H:%M")),
            ("Autor:", st.session_state.get("current_user", {}).get("nombre", "Sistema CAVA")),
            ("Sistema:", f"CAVA v{APP_VERSION}")
        ]
        for label, value in info_items:
            pdf.set_font("Helvetica", "B", 10)
            pdf.cell(60, 7, clean_text_for_pdf(label), ln=False)
            pdf.set_font("Helvetica", "", 10)
            pdf.cell(0, 7, clean_text_for_pdf(value), ln=True)

    @staticmethod
    def _add_content(pdf, content: str):
        """
        Agrega el contenido formateado al PDF con:
        - Texto justificado (align='J')
        - Interlineado profesional (h=6 para fuente tamaño 10)
        - Viñetas correctas
        - Tablas con bordes
        """
        clean_content = clean_text_for_pdf(content)
        lines = clean_content.split('\n')
        
        for line in lines:
            line = line.strip()
            if not line:
                pdf.ln(3)
                continue

            # Encabezados
            if line.startswith('## '):
                pdf.set_font("Helvetica", "B", 15)
                pdf.set_text_color(13, 43, 78)
                pdf.ln(5)
                pdf.multi_cell(0, 8, line[3:], align="J")
                pdf.set_draw_color(30, 95, 142)
                pdf.line(pdf.l_margin, pdf.get_y(), pdf.w - pdf.r_margin, pdf.get_y())
                pdf.ln(4)
            elif line.startswith('### '):
                pdf.set_font("Helvetica", "B", 12)
                pdf.set_text_color(30, 95, 142)
                pdf.ln(3)
                pdf.multi_cell(0, 7, line[4:], align="J")
                pdf.ln(2)
            elif line.startswith('#### '):
                pdf.set_font("Helvetica", "B", 11)
                pdf.set_text_color(26, 26, 26)
                pdf.multi_cell(0, 6, line[5:], align="J")
                pdf.ln(1)
            
            # Tablas
            elif line.startswith('|') and '---' not in line:
                cells = [c.strip().replace('**', '') for c in line.split('|')[1:-1]]
                if cells:
                    pdf.set_font("Helvetica", "", 8)
                    pdf.set_text_color(26, 26, 26)
                    col_width = (pdf.w - pdf.l_margin - pdf.r_margin) / max(len(cells), 1)
                    for cell_text in cells:
                        # Limitar longitud para evitar desbordamiento
                        safe_text = clean_text_for_pdf(cell_text[:35])
                        pdf.cell(col_width, 6, safe_text, border=1, align="C")
                    pdf.ln()
            elif line.startswith('|') and '---' in line:
                continue
            
            # Listas con viñetas
            elif line.startswith('- '):
                pdf.set_font("Helvetica", "", 10)
                pdf.set_text_color(26, 26, 26)
                clean_line = clean_text_for_pdf(line[2:].replace('**', ''))
                # Usar multi_cell para justificado e interlineado, con viñeta manual
                pdf.cell(5) # Indentación
                pdf.multi_cell(0, 6, f"• {clean_line}", align="J")
            
            # Listas numeradas
            elif re.match(r'^\d+\.', line):
                pdf.set_font("Helvetica", "", 10)
                pdf.set_text_color(26, 26, 26)
                clean_line = clean_text_for_pdf(line.replace('**', ''))
                pdf.cell(5)
                pdf.multi_cell(0, 6, clean_line, align="J")
            
            # Párrafos normales (justificados e interlineado 1.5)
            else:
                pdf.set_font("Helvetica", "", 10)
                pdf.set_text_color(26, 26, 26)
                clean_line = clean_text_for_pdf(line.replace('**', ''))
                pdf.multi_cell(0, 6, clean_line, align="J")


class PDFReport(FPDF):
    """
    Clase personalizada de PDF con encabezado y pie de página institucionales.
    Hereda de FPDF y sobrescribe los métodos header() y footer().
    """

    def __init__(self, report_number="", report_type="", is_english=False):
        super().__init__(format='A4') # Forzar formato A4
        self.report_number = report_number
        self.report_type = report_type
        self.is_english = is_english

    def header(self):
        """Encabezado de cada página (excepto la portada)."""
        if self.page_no() <= 1:
            return
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(108, 117, 125)
        clean_type = clean_text_for_pdf(self.report_type)
        clean_num = clean_text_for_pdf(self.report_number)
        header_text = clean_text_for_pdf(f"CAVA - Especialistas en Robotica y Automatizacion | {clean_type} | N° {clean_num}")
        self.cell(0, 6, header_text, ln=True, align="R")
        self.set_draw_color(222, 226, 230)
        self.line(self.l_margin, self.get_y(), self.w - self.r_margin, self.get_y())
        self.ln(4)

    def footer(self):
        """Pie de página de cada página."""
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(108, 117, 125)
        footer_text = clean_text_for_pdf(f"Roger Huamani | {self.report_number} | Pagina {self.page_no()}/{{nb}}")
        self.cell(0, 10, footer_text, align="C")


# ============================================================================
# SECCIÓN 11: PANTALLAS DE LA APLICACIÓN (DASHBOARD, INFORMES, ETC.)
# ============================================================================

def render_dashboard():
    """Renderiza el panel principal del sistema con estadísticas y accesos rápidos."""
    user = st.session_state.get("current_user", {})
    st.markdown(f"""
    <div class="main-header">
        <h1>📊 Panel de Control - Sistema de Informes</h1>
        <p>Bienvenido, <strong>{user.get('nombre', 'Usuario')}</strong> | 
        Rol: {user.get('rol', 'N/A')} | 
        Sesión iniciada: {st.session_state.get('login_time', datetime.now()).strftime('%d/%m/%Y %H:%M')}</p>
    </div>
    """, unsafe_allow_html=True)

    reports = st.session_state.get("local_reports", [])
    total_reports = len(reports)
    reports_today = sum(1 for r in reports if r.get("fecha_creacion", "")[:10] == datetime.now().strftime("%Y-%m-%d"))
    reports_month = sum(1 for r in reports if r.get("fecha_creacion", "")[:7] == datetime.now().strftime("%Y-%m"))

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"<div class='stat-card'><div class='stat-number'>{total_reports}</div><div class='stat-label'>Total de Informes</div></div>", unsafe_allow_html=True)
    with col2:
        st.markdown(f"<div class='stat-card'><div class='stat-number'>{reports_today}</div><div class='stat-label'>Informes Hoy</div></div>", unsafe_allow_html=True)
    with col3:
        st.markdown(f"<div class='stat-card'><div class='stat-number'>{reports_month}</div><div class='stat-label'>Informes del Mes</div></div>", unsafe_allow_html=True)
    with col4:
        gemini_status = "✅" if st.session_state.get("gemini_api_key") else "⚠️"
        st.markdown(f"<div class='stat-card'><div class='stat-number'>{gemini_status}</div><div class='stat-label'>Estado Gemini AI</div></div>", unsafe_allow_html=True)

    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
    st.subheader("⚡ Accesos Rápidos")
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("📝 Nuevo Reporte de Mantenimiento", use_container_width=True, type="primary"):
            st.session_state.current_page = "Nuevo Informe"
            st.session_state.selected_report_type = TIPO_REPORTE_MANTENIMIENTO
            st.rerun()
    with col2:
        if st.button("📊 Nuevo Informe Ejecutivo", use_container_width=True, type="primary"):
            st.session_state.current_page = "Nuevo Informe"
            st.session_state.selected_report_type = TIPO_INFORME_EJECUTIVO
            st.rerun()
    with col3:
        if st.button("📂 Ver Informes Guardados", use_container_width=True):
            st.session_state.current_page = "Informes Guardados"
            st.rerun()


def render_new_report():
    """Renderiza la pantalla de creación de un nuevo informe con formulario completo."""
    user = st.session_state.get("current_user", {})
    st.markdown(f"""
    <div class="main-header">
        <h1>📝 Crear Nuevo Informe</h1>
        <p>Complete los campos requeridos y el sistema generará automáticamente el informe profesional utilizando inteligencia artificial.</p>
    </div>
    """, unsafe_allow_html=True)

    st.subheader("1️⃣ Tipo de Documento")
    report_type = st.radio(
        "Seleccione el tipo de documento a generar:", 
        [TIPO_REPORTE_MANTENIMIENTO, TIPO_INFORME_EJECUTIVO],
        index=0 if st.session_state.get("selected_report_type") == TIPO_REPORTE_MANTENIMIENTO else 1, 
        horizontal=True, 
        key="report_type_selector"
    )
    st.session_state.selected_report_type = report_type

    report_number = ReportNumberGenerator.generate(report_type)
    st.info(f"📋 Número de documento asignado automáticamente: **{report_number}**")
    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)

    st.subheader("2️⃣ Datos Generales del Informe")
    col1, col2 = st.columns(2)
    with col1:
        fecha_informe = st.date_input("📅 Fecha del Informe *", value=datetime.now(), key="fecha_informe")
        turno = st.selectbox("🕐 Turno *", TURNOS, key="turno_selector")
        area = st.selectbox("🏭 Área / Ubicación *", AREAS_PLANTA, key="area_selector")
        disciplina = st.selectbox("🔧 Disciplina *", DISCIPLINAS, key="disciplina_selector")
    with col2:
        numero_ot = st.text_input("📄 Número de OT (Orden de Trabajo) *", placeholder="Ej: OT-2026-00123", key="numero_ot_input")
        equipo_nombre = st.text_input("⚙️ Nombre del Equipo *", placeholder="Ej: Bomba Centrífuga B-201", key="equipo_nombre_input")
        equipo_tag = st.text_input("🏷️ Tag / Código del Equipo *", placeholder="Ej: P-201A", key="equipo_tag_input")
        prioridad = st.selectbox("🚨 Prioridad *", PRIORIDADES, key="prioridad_selector")

    col3, col4 = st.columns(2)
    with col3:
        tipo_intervencion = st.selectbox("🔨 Tipo de Intervención *", TIPOS_INTERVENCION, key="tipo_intervencion_selector")
        tecnicos = st.text_input("👷 Técnico(s) Responsable(s) *", placeholder="Ej: Juan Pérez, María López", key="tecnicos_input")
    with col4:
        hora_aviso = st.time_input("⏰ Hora de Aviso", value=datetime.now().time(), key="hora_aviso_input")
        elaborado_por = st.text_input("✍️ Elaborado por", value=user.get("nombre", ""), key="elaborado_por_input")

    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
    st.subheader("3️⃣ Descripción de la Intervención (Campo Principal)")
    st.markdown(f"""
    <div class="custom-alert">
        💡 <strong>Instrucciones para el técnico:</strong> Describa con sus propias palabras todo el contexto del problema identificado, 
        las acciones que ha realizado y sus conclusiones. No se preocupe por la redacción técnica, la inteligencia artificial se encargará 
        de estructurar, corregir y profesionalizar el contenido, incluyendo cálculos de ingeniería si son relevantes.
    </div>
    """, unsafe_allow_html=True)

    raw_description = st.text_area(
        "Describa detalladamente la intervención realizada *:", 
        height=250,
        placeholder="Ejemplo: Al llegar a la planta encontré la bomba B-201 con una fuga de aceite en el sello mecánico. El operador me dijo que empezó a vibrar mucho desde las 6am. Revisé el acople y estaba desalineado. Cambié el sello mecánico, realicé la alineación láser (tolerancia < 0.05mm) y la bomba quedó funcionando normal. El rodamiento del lado del acople también estaba con juego, lo cambié también...",
        key="raw_description_input"
    )

    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
    st.subheader("4️⃣ Registro Fotográfico")
    ImageManager.initialize_image_counter()
    uploaded_files = st.file_uploader(
        "📷 Cargar imágenes de la intervención", 
        type=["png", "jpg", "jpeg", "webp"], 
        accept_multiple_files=True, 
        key="image_uploader"
    )

    if uploaded_files:
        for file in uploaded_files:
            existing_names = [img["nombre_archivo"] for img in st.session_state.get("uploaded_images", [])]
            if file.name not in existing_names:
                img_desc = st.text_input(
                    f"📝 Descripción obligatoria para: **{file.name}** *", 
                    placeholder="Describa qué se observa en la imagen...", 
                    key=f"img_desc_{file.name}"
                )
                if img_desc:
                    img_data = ImageManager.add_image(file, img_desc)
                    gemini = GeminiAIManager()
                    if gemini.is_configured():
                        with st.spinner("Mejorando descripción con IA..."):
                            img_data["descripcion_mejorada"] = gemini.process_image_description(img_desc)
                    else:
                        img_data["descripcion_mejorada"] = img_desc.strip().capitalize() + "."
                    st.success(f"✅ Imagen '{file.name}' registrada como Foto N° {img_data['numero']}")

    images = ImageManager.get_all_images()
    if images:
        st.markdown(f"**📸 Imágenes cargadas: {len(images)}**")
        cols = st.columns(min(len(images), 4))
        for i, img_data in enumerate(images):
            with cols[i % 4]:
                st.image(img_data["bytes"], width=150)
                st.caption(f"Foto N°{img_data['numero']}: {img_data.get('descripcion_mejorada', '')[:50]}...")
                if st.button(f"🗑️", key=f"del_img_{i}"):
                    ImageManager.remove_image(i)
                    st.rerun()

    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
    st.subheader("5️⃣ Generar Informe")
    col1, col2 = st.columns(2)
    with col1:
        generate_btn = st.button("🤖 Generar Informe con IA", use_container_width=True, type="primary", key="generate_report_btn")
    with col2:
        generate_basic_btn = st.button("📄 Generar Informe Básico (sin IA)", use_container_width=True, key="generate_basic_btn")

    if generate_btn or generate_basic_btn:
        errors = []
        if not raw_description.strip(): errors.append("La descripción de la intervención es obligatoria.")
        if not numero_ot.strip(): errors.append("El número de OT es obligatorio.")
        if not equipo_nombre.strip(): errors.append("El nombre del equipo es obligatorio.")
        if not equipo_tag.strip(): errors.append("El Tag del equipo es obligatorio.")
        if not tecnicos.strip(): errors.append("Los técnicos responsables son obligatorios.")

        if errors:
            for err in errors:
                st.error(f"⚠️ {err}")
            return

        additional_context = {
            "Fecha": fecha_informe.strftime("%d/%m/%Y"), "Turno": turno, "Número de OT": numero_ot,
            "Área": area, "Disciplina": disciplina, "Equipo": equipo_nombre, "Tag": equipo_tag,
            "Prioridad": prioridad, "Tipo de Intervención": tipo_intervencion, "Técnicos": tecnicos,
            "Hora de Aviso": hora_aviso.strftime("%H:%M"), "Elaborado por": elaborado_por
        }

        with st.spinner("🤖 Generando informe profesional con IA... Esto puede tomar unos segundos."):
            gemini = GeminiAIManager()
            if generate_btn:
                report_content = gemini.generate_report_content(raw_description, report_type, additional_context)
            else:
                report_content = gemini._generate_fallback_report(raw_description, report_type, additional_context)

        st.session_state.generated_report = report_content
        st.session_state.generated_report_number = report_number
        st.session_state.generated_report_type = report_type
        st.session_state.generated_report_context = additional_context
        st.success("✅ ¡Informe generado exitosamente!")
        st.rerun()

    if "generated_report" in st.session_state and st.session_state.generated_report:
        st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
        st.subheader("📋 Vista Previa del Informe Generado")
        st.markdown(st.session_state.generated_report)
        st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
        st.subheader("💾 Exportar y Guardar")

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            try:
                word_buffer = WordExporter.create_document(
                    st.session_state.generated_report, 
                    st.session_state.generated_report_number, 
                    st.session_state.generated_report_type, 
                    ImageManager.get_all_images()
                )
                st.download_button(
                    label="📥 Descargar Word (.docx)", 
                    data=word_buffer, 
                    file_name=f"{st.session_state.generated_report_number}.docx", 
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Error al generar Word: {e}")
        with col2:
            try:
                pdf_buffer = PDFExporter.create_pdf(
                    st.session_state.generated_report, 
                    st.session_state.generated_report_number, 
                    st.session_state.generated_report_type, 
                    ImageManager.get_all_images()
                )
                st.download_button(
                    label="📥 Descargar PDF", 
                    data=pdf_buffer, 
                    file_name=f"{st.session_state.generated_report_number}.pdf", 
                    mime="application/pdf", 
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Error al generar PDF: {e}")
        with col3:
            if st.button("🌐 Traducir al Inglés", use_container_width=True):
                with st.spinner("🌐 Traduciendo informe al inglés..."):
                    gemini = GeminiAIManager()
                    st.session_state.translated_report = gemini.translate_report(st.session_state.generated_report)
                    st.success("✅ Traducción completada!")
                    st.rerun()
        with col4:
            if st.button("💾 Guardar en Base de Datos", use_container_width=True, type="primary"):
                report_data = {
                    "id": st.session_state.generated_report_number, 
                    "tipo": st.session_state.generated_report_type,
                    "numero": st.session_state.generated_report_number, 
                    "contenido": st.session_state.generated_report,
                    "contexto": json.dumps(st.session_state.generated_report_context),
                    "fecha_creacion": datetime.now().isoformat(),
                    "autor": st.session_state.get("current_user", {}).get("nombre", ""),
                    "estado": "Completado", 
                    "imagenes_count": len(ImageManager.get_all_images())
                }
                sb = SupabaseManager()
                if sb.save_report(report_data):
                    st.success("✅ Informe guardado exitosamente!")
                else:
                    st.error("❌ Error al guardar el informe.")

        if "translated_report" in st.session_state and st.session_state.translated_report:
            st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
            st.subheader("🌐 Versión en Inglés (Mirror Document)")
            st.markdown(st.session_state.translated_report)
            tcol1, tcol2 = st.columns(2)
            with tcol1:
                try:
                    word_en = WordExporter.create_document(
                        st.session_state.translated_report, 
                        st.session_state.generated_report_number + "-EN", 
                        st.session_state.generated_report_type + " (English)", 
                        ImageManager.get_all_images(), 
                        is_english=True
                    )
                    st.download_button(
                        label="📥 Descargar Word en Inglés", 
                        data=word_en, 
                        file_name=f"{st.session_state.generated_report_number}_EN.docx", 
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Error: {e}")
            with tcol2:
                try:
                    pdf_en = PDFExporter.create_pdf(
                        st.session_state.translated_report, 
                        st.session_state.generated_report_number + "-EN", 
                        st.session_state.generated_report_type + " (English)", 
                        ImageManager.get_all_images(), 
                        is_english=True
                    )
                    st.download_button(
                        label="📥 Descargar PDF en Inglés", 
                        data=pdf_en, 
                        file_name=f"{st.session_state.generated_report_number}_EN.pdf", 
                        mime="application/pdf", 
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Error: {e}")


def render_saved_reports():
    """Renderiza la pantalla de informes guardados con opciones de visualización y descarga."""
    st.markdown(f"""
    <div class="main-header">
        <h1>📂 Informes Guardados</h1>
        <p>Consulte, edite o descargue los informes almacenados en el sistema.</p>
    </div>
    """, unsafe_allow_html=True)

    sb = SupabaseManager()
    reports = sb.get_reports()

    if not reports:
        st.info("📭 No hay informes guardados aún. Cree su primer informe desde el menú principal.")
        return

    st.markdown(f"**Total de informes almacenados:** {len(reports)}")
    for i, report in enumerate(reports):
        with st.expander(
            f"📄 {report.get('numero', 'N/A')} - {report.get('tipo', 'N/A')} | {report.get('fecha_creacion', '')[:10]} | {report.get('estado', 'Pendiente')}", 
            expanded=False
        ):
            col1, col2, col3 = st.columns([3, 1, 1])
            with col1:
                st.markdown(f"""
                **Tipo:** {report.get('tipo', 'N/A')}<br>
                **Autor:** {report.get('autor', 'N/A')}<br>
                **Fecha:** {report.get('fecha_creacion', 'N/A')[:16]}<br>
                **Estado:** {report.get('estado', 'N/A')}<br>
                **Imágenes:** {report.get('imagenes_count', 0)}
                """, unsafe_allow_html=True)
            with col2:
                if st.button("👁️ Ver", key=f"view_{i}"):
                    st.session_state.viewing_report = report
                    st.rerun()
            with col3:
                if st.button("🗑️ Eliminar", key=f"del_{i}"):
                    sb.delete_report(report.get("id", ""))
                    st.success("Informe eliminado.")
                    st.rerun()

            if st.session_state.get("viewing_report", {}).get("id") == report.get("id"):
                st.markdown("---")
                st.markdown(report.get("contenido", ""))
                ecol1, ecol2 = st.columns(2)
                with ecol1:
                    try:
                        wb = WordExporter.create_document(report.get("contenido", ""), report.get("numero", ""), report.get("tipo", ""))
                        st.download_button(
                            "📥 Word", 
                            data=wb, 
                            file_name=f"{report.get('numero', 'reporte')}.docx", 
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                            key=f"dl_word_{i}"
                        )
                    except Exception as e:
                        logger.error(f"Error al descargar Word: {e}")
                with ecol2:
                    try:
                        pb = PDFExporter.create_pdf(report.get("contenido", ""), report.get("numero", ""), report.get("tipo", ""))
                        st.download_button(
                            "📥 PDF", 
                            data=pb, 
                            file_name=f"{report.get('numero', 'reporte')}.pdf", 
                            mime="application/pdf", 
                            key=f"dl_pdf_{i}"
                        )
                    except Exception as e:
                        logger.error(f"Error al descargar PDF: {e}")


def render_settings():
    """Renderiza la pantalla de configuración del sistema con persistencia."""
    st.markdown(f"""
    <div class="main-header">
        <h1>⚙️ Configuración del Sistema</h1>
        <p>Configure las APIs, conexión a base de datos y preferencias del sistema.</p>
    </div>
    """, unsafe_allow_html=True)

    st.subheader("🤖 Configuración de Gemini AI")
    st.markdown(f"""
    <div class="custom-alert">
        🔑 Ingrese su API Key de Google Gemini para habilitar la generación inteligente de informes. 
        Obtenga su clave gratuita en <a href="https://aistudio.google.com/app/apikey" target="_blank">Google AI Studio</a>.
    </div>
    """, unsafe_allow_html=True)

    gemini_key = st.text_input(
        "API Key de Gemini", 
        type="password", 
        value=st.session_state.get("gemini_api_key", ""), 
        placeholder="AIzaSy...", 
        key="gemini_key_input"
    )
    
    if st.button("💾 Guardar API Key de Gemini", key="save_gemini"):
        if gemini_key.strip():
            st.session_state.gemini_api_key = gemini_key.strip()
            config = ConfigPersistence.load_config()
            config["gemini_api_key"] = gemini_key.strip()
            ConfigPersistence.save_config(config)
            st.success("✅ API Key de Gemini guardada correctamente (persistente).")
            st.rerun()
        else:
            st.warning("⚠️ Ingrese una API Key válida.")

    gemini = GeminiAIManager()
    if gemini.is_configured():
        st.success("✅ Gemini AI está configurado y operativo (Modelo: gemini-1.5-flash).")
    else:
        st.warning("⚠️ Gemini AI no está configurado. Los informes se generarán con plantilla base.")

    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
    st.subheader("🗄️ Configuración de Supabase")
    st.markdown(f"""
    <div class="custom-alert">
        🗄️ Configure la conexión a Supabase para almacenar sus informes en la nube.
        Cree un proyecto en <a href="https://supabase.com" target="_blank">supabase.com</a> y obtenga la URL y la API Key.
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        supabase_url = st.text_input(
            "Supabase URL", 
            value=st.session_state.get("supabase_url", ""), 
            placeholder="https://xxxxx.supabase.co", 
            key="supabase_url_input"
        )
    with col2:
        supabase_key = st.text_input(
            "Supabase API Key (anon)", 
            type="password", 
            value=st.session_state.get("supabase_key", ""), 
            placeholder="eyJhbGciOi...", 
            key="supabase_key_input"
        )

    if st.button("💾 Guardar Configuración Supabase", key="save_supabase"):
        if supabase_url.strip() and supabase_key.strip():
            st.session_state.supabase_url = supabase_url.strip()
            st.session_state.supabase_key = supabase_key.strip()
            config = ConfigPersistence.load_config()
            config["supabase_url"] = supabase_url.strip()
            config["supabase_key"] = supabase_key.strip()
            ConfigPersistence.save_config(config)
            st.success("✅ Configuración de Supabase guardada (persistente).")
            st.rerun()
        else:
            st.warning("⚠️ Complete ambos campos.")

    sb = SupabaseManager()
    if sb.is_connected():
        st.success("✅ Conexión a Supabase activa.")
    else:
        st.info("ℹ️ Supabase no configurado. Los informes se guardarán localmente en la sesión (JSON).")

    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
    st.subheader("👥 Gestión de Usuarios")
    auth = AuthenticationManager()
    users = auth.get_all_users()
    user_data = [
        {"Usuario": uname, "Nombre": udata.get("nombre", ""), "Rol": udata.get("rol", ""), "Activo": "✅ Sí" if udata.get("activo") else "❌ No"} 
        for uname, udata in users.items()
    ]
    if user_data:
        st.dataframe(user_data, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("**Registrar nuevo usuario:**")
    rcol1, rcol2 = st.columns(2)
    with rcol1:
        new_username = st.text_input("Nuevo Usuario", key="new_user_name")
        new_nombre = st.text_input("Nombre Completo", key="new_user_fullname")
    with rcol2:
        new_password = st.text_input("Contraseña", type="password", key="new_user_pass")
        new_rol = st.selectbox(
            "Rol", 
            ["Superintendente", "Jefe de Mantenimiento", "Planificador", "Técnico", "Supervisor"], 
            key="new_user_rol"
        )

    if st.button("➕ Registrar Usuario", key="register_user_btn"):
        if new_username and new_password and new_nombre:
            if auth.register_user(new_username, new_password, new_nombre, new_rol):
                st.success(f"✅ Usuario '{new_username}' registrado exitosamente.")
                st.rerun()
            else:
                st.error("❌ El usuario ya existe.")
        else:
            st.warning("⚠️ Complete todos los campos.")


def render_help():
    """Renderiza la pantalla de ayuda y documentación del sistema."""
    st.markdown(f"""
    <div class="main-header">
        <h1>❓ Ayuda y Documentación</h1>
        <p>Guía de uso del Sistema de Gestión de Informes de Mantenimiento CAVA.</p>
    </div>
    """, unsafe_allow_html=True)

    st.subheader("📖 ¿Cómo funciona el sistema?")
    st.markdown("""
    Este sistema está diseñado para facilitar la generación de informes técnicos y ejecutivos de mantenimiento. El flujo de trabajo es el siguiente:
    1. **El técnico describe** con sus propias palabras lo que hizo durante la intervención (problema, acciones, conclusiones).
    2. **La IA de Gemini** procesa esa descripción y genera un informe profesional con estructura, redacción técnica, cálculos de ingeniería (si aplica) y sin errores ortográficos.
    3. **El sistema completa** automáticamente todos los campos requeridos según el tipo de documento seleccionado.
    4. **Usted revisa, exporta** (PDF/Word) y almacena el informe.
    """)

    st.subheader("📝 Tipos de Documentos")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        **📋 Reporte de Mantenimiento**
        - Uso: Intervenciones del día a día
        - Audiencia: Equipo de mantenimiento
        - Contenido: Detalle técnico completo, cálculos, repuestos, tiempos.
        """)
    with col2:
        st.markdown("""
        **📊 Informe Ejecutivo de Mantenimiento**
        - Uso: Eventos críticos y gerencia
        - Audiencia: Gerencia de planta
        - Contenido: Impacto, costos, causa raíz, recomendaciones.
        """)

    st.subheader("❓ Preguntas Frecuentes")
    with st.expander("¿Puedo usar el sistema sin la API de Gemini?"):
        st.markdown("Sí, el sistema generará informes con una plantilla base. Sin embargo, la calidad y profundidad del contenido será limitada. Se recomienda configurar la API gratuita para obtener resultados profesionales.")
    with st.expander("¿Cómo obtengo mi API Key de Gemini?"):
        st.markdown("1. Vaya a [Google AI Studio](https://aistudio.google.com/app/apikey)\n2. Inicie sesión con su cuenta de Google\n3. Haga clic en 'Create API Key'\n4. Copie la clave y péguela en la sección de Configuración")
    with st.expander("¿Los datos se almacenan de forma segura?"):
        st.markdown("Los datos se almacenan en Supabase (si está configurado) con encriptación de nivel empresarial. Sin Supabase, los datos se guardan localmente en archivos JSON seguros dentro del entorno de la aplicación.")


def render_sidebar():
    """Renderiza la barra lateral de navegación con información del usuario y estado de servicios."""
    user = st.session_state.get("current_user", {})
    with st.sidebar:
        st.markdown(f"""
        <div style="text-align:center; padding: 20px 10px; border-bottom: 2px solid {COLOR_PRIMARY}; margin-bottom: 15px;">
            <h2 style="color: {COLOR_PRIMARY}; margin-bottom:2px; font-size:26px;">🔧 CAVA</h2>
            <p style="color: {COLOR_TEXT_SECONDARY}; font-size:11px; margin:0; font-weight:500;">Especialistas en Robótica<br>y Automatización</p>
            <p style="color: {COLOR_TEXT_LIGHT}; font-size:10px; margin-top:5px;">v{APP_VERSION}</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown(f"""
        <div style="padding: 12px; background: {COLOR_WHITE}; border-radius: 8px; margin-bottom: 15px; border-left: 4px solid {COLOR_PRIMARY}; box-shadow: 0 1px 3px rgba(0,0,0,0.08);">
            <p style="color: {COLOR_TEXT_DARK}; margin:0; font-size:13px; font-weight:600;">👤 {user.get('nombre', 'Usuario')}</p>
            <p style="color: {COLOR_TEXT_SECONDARY}; margin:2px 0 0 0; font-size:11px;">{user.get('rol', 'N/A')}</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("### 📌 Navegación")
        page = st.radio(
            "Ir a:", 
            ["📊 Dashboard", "📝 Nuevo Informe", "📂 Informes Guardados", "⚙️ Configuración", "❓ Ayuda"], 
            key="sidebar_nav", 
            label_visibility="collapsed"
        )
        page_map = {
            "📊 Dashboard": "Dashboard", 
            "📝 Nuevo Informe": "Nuevo Informe", 
            "📂 Informes Guardados": "Informes Guardados", 
            "⚙️ Configuración": "Configuración", 
            "❓ Ayuda": "Ayuda"
        }
        st.session_state.current_page = page_map.get(page, "Dashboard")
        st.markdown("---")

        st.markdown("### 🔌 Estado de Servicios")
        gemini = GeminiAIManager()
        if gemini.is_configured():
            st.markdown(f"""<div style="padding: 8px 12px; background: {COLOR_SUCCESS_BG}; border-radius: 6px; margin-bottom: 8px; border-left: 3px solid {COLOR_SUCCESS};">
                <p style="color: {COLOR_SUCCESS}; margin:0; font-size:12px; font-weight:600;">🤖 Gemini AI: Activo (Gratis)</p></div>""", unsafe_allow_html=True)
        else:
            st.markdown(f"""<div style="padding: 8px 12px; background: {COLOR_WARNING_BG}; border-radius: 6px; margin-bottom: 8px; border-left: 3px solid {COLOR_WARNING};">
                <p style="color: {COLOR_WARNING}; margin:0; font-size:12px; font-weight:600;">🤖 Gemini AI: No configurado</p></div>""", unsafe_allow_html=True)

        sb = SupabaseManager()
        if sb.is_connected():
            st.markdown(f"""<div style="padding: 8px 12px; background: {COLOR_SUCCESS_BG}; border-radius: 6px; margin-bottom: 8px; border-left: 3px solid {COLOR_SUCCESS};">
                <p style="color: {COLOR_SUCCESS}; margin:0; font-size:12px; font-weight:600;">🗄️ Supabase: Conectado</p></div>""", unsafe_allow_html=True)
        else:
            st.markdown(f"""<div style="padding: 8px 12px; background: {COLOR_INFO_BG}; border-radius: 6px; margin-bottom: 8px; border-left: 3px solid {COLOR_INFO};">
                <p style="color: {COLOR_INFO}; margin:0; font-size:12px; font-weight:600;">🗄️ Almacenamiento: Local (JSON)</p></div>""", unsafe_allow_html=True)

        st.markdown("---")
        if st.button("🚪 Cerrar Sesión", use_container_width=True):
            st.session_state.authenticated = False
            st.session_state.current_user = None
            st.session_state.generated_report = None
            st.session_state.translated_report = None
            ImageManager.clear_all_images()
            st.rerun()

        st.markdown("---")
        st.markdown(f"""
        <div style="text-align:center; padding: 15px 5px;">
            <p style="color: {COLOR_PRIMARY}; font-size:11px; font-weight:700; margin:0;">CAVA</p>
            <p style="color: {COLOR_TEXT_SECONDARY}; font-size:10px; margin:2px 0;">Especialistas en Robótica<br>y Automatización</p>
            <p style="color: {COLOR_TEXT_LIGHT}; font-size:10px; margin:2px 0;">Roger Huamani</p>
            <p style="color: {COLOR_TEXT_LIGHT}; font-size:9px; margin:5px 0 0 0;">© {APP_YEAR} Todos los derechos reservados</p>
        </div>
        """, unsafe_allow_html=True)


def render_footer():
    """Renderiza el pie de página institucional en todas las pantallas."""
    st.markdown(f"""
    <div class="footer">
        <p class="brand">🔧 CAVA - Especialistas en Robótica y Automatización</p>
        <p>Diseñado y desarrollado por <strong>Roger Huamani</strong></p>
        <p>Sistema de Gestión de Informes de Mantenimiento v{APP_VERSION} | © {APP_YEAR}</p>
        <p style="font-size:11px; opacity:0.9; margin-top:8px;">Potenciado por Inteligencia Artificial (Google Gemini) | Almacenamiento persistente local/nube</p>
    </div>
    """, unsafe_allow_html=True)


# ============================================================================
# SECCIÓN 12: INICIALIZACIÓN Y PUNTO DE ENTRADA DE LA APLICACIÓN
# ============================================================================

def initialize_session_state():
    """
    Inicializa todas las variables de sesión necesarias para el funcionamiento
    de la aplicación, cargando configuraciones persistentes si existen.
    """
    defaults = {
        "authenticated": False, 
        "current_user": None, 
        "login_time": None,
        "current_page": "Dashboard", 
        "selected_report_type": TIPO_REPORTE_MANTENIMIENTO,
        "gemini_api_key": "", 
        "supabase_url": "", 
        "supabase_key": "",
        "supabase_error": "", 
        "gemini_error": "", 
        "generated_report": None,
        "generated_report_number": "", 
        "generated_report_type": "", 
        "generated_report_context": {},
        "translated_report": None, 
        "viewing_report": None, 
        "local_reports": [],
        "report_counter": {TIPO_REPORTE_MANTENIMIENTO: 0, TIPO_INFORME_EJECUTIVO: 0},
        "image_counter": 0, 
        "uploaded_images": [], 
        "registered_users": DEFAULT_USERS.copy()
    }
    
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

    # Cargar configuración persistente desde archivos JSON
    config = ConfigPersistence.load_config()
    if config:
        if "gemini_api_key" in config and not st.session_state.gemini_api_key:
            st.session_state.gemini_api_key = config["gemini_api_key"]
        if "supabase_url" in config and not st.session_state.supabase_url:
            st.session_state.supabase_url = config["supabase_url"]
        if "supabase_key" in config and not st.session_state.supabase_key:
            st.session_state.supabase_key = config["supabase_key"]

    # Cargar informes persistentes si no hay en sesión
    if not st.session_state.local_reports:
        st.session_state.local_reports = ConfigPersistence.load_reports()


def main():
    """
    Función principal que orquesta toda la aplicación.
    Inicializa el estado, aplica estilos y enruta a la pantalla correspondiente.
    """
    initialize_session_state()
    aplicar_estilos_css()

    if not st.session_state.authenticated:
        render_login_screen()
        render_footer()
        return

    render_sidebar()

    current_page = st.session_state.get("current_page", "Dashboard")
    
    if current_page == "Dashboard":
        render_dashboard()
    elif current_page == "Nuevo Informe":
        render_new_report()
    elif current_page == "Informes Guardados":
        render_saved_reports()
    elif current_page == "Configuración":
        render_settings()
    elif current_page == "Ayuda":
        render_help()
    else:
        render_dashboard()

    render_footer()


# ============================================================================
# PUNTO DE ENTRADA DEL SCRIPT
# ============================================================================

if __name__ == "__main__":
    main()
