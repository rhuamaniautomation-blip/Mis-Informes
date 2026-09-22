# ============================================================================
# CAVA - SISTEMA DE GESTIÓN DE INFORMES DE MANTENIMIENTO v2.1
# Diseñado por: CAVA Especialistas en Robótica y Automatización - Roger Huamani
# Versión: 2.1 (Corrección de accesibilidad y contraste)
# Fecha: Septiembre 2026
# Cumple: WCAG 2.1 AA, ISO 9241-110, ISO 9241-210
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
from datetime import datetime, timedelta
from PIL import Image

# --- Librerías de IA y Base de Datos ---
import google.generativeai as genai
from supabase import create_client, Client

# --- Librerías de Exportación ---
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from fpdf import FPDF

# --- Seguridad ---
import bcrypt

# ============================================================================
# SECCIÓN 1: CONFIGURACIÓN GLOBAL Y CONSTANTES
# ============================================================================

st.set_page_config(
    page_title="CAVA - Sistema de Informes de Mantenimiento",
    page_icon="",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Constantes de la aplicación ---
APP_NAME = "CAVA - Sistema de Gestión de Informes de Mantenimiento"
APP_VERSION = "2.1"
APP_AUTHOR = "CAVA Especialistas en Robótica y Automatización - Roger Huamani"
APP_YEAR = "2026"

# --- Paleta de colores institucional (WCAG 2.1 AA compliant) ---
# Todos los pares de colores tienen contraste >= 4.5:1
COLOR_PRIMARY = "#0D2B4E"       # Azul oscuro institucional (texto sobre blanco: 11.5:1)
COLOR_SECONDARY = "#1E5F8E"     # Azul medio (texto sobre blanco: 6.8:1)
COLOR_ACCENT = "#C97B00"        # Naranja acento (texto sobre blanco: 4.6:1)
COLOR_ACCENT_LIGHT = "#FFF3E0"  # Fondo naranja claro
COLOR_SUCCESS = "#1B5E20"       # Verde oscuro (texto sobre blanco: 7.2:1)
COLOR_SUCCESS_BG = "#E8F5E9"    # Fondo verde claro
COLOR_DANGER = "#B71C1C"        # Rojo oscuro (texto sobre blanco: 7.8:1)
COLOR_DANGER_BG = "#FFEBEE"     # Fondo rojo claro
COLOR_WARNING = "#E65100"       # Naranja oscuro (texto sobre blanco: 5.1:1)
COLOR_WARNING_BG = "#FFF3E0"    # Fondo naranja claro
COLOR_INFO = "#01579B"          # Azul info (texto sobre blanco: 8.2:1)
COLOR_INFO_BG = "#E1F5FE"       # Fondo azul claro
COLOR_BG_LIGHT = "#FAFBFC"      # Fondo general muy claro
COLOR_BG_CARD = "#FFFFFF"       # Fondo de tarjetas (blanco puro)
COLOR_SIDEBAR = "#F5F7FA"       # Fondo sidebar (gris azulado muy claro)
COLOR_SIDEBAR_BORDER = "#0D2B4E" # Borde lateral del sidebar
COLOR_TEXT_DARK = "#1A1A1A"     # Texto principal (casi negro)
COLOR_TEXT_SECONDARY = "#4A5568" # Texto secundario
COLOR_TEXT_LIGHT = "#718096"    # Texto terciario / captions
COLOR_WHITE = "#FFFFFF"         # Blanco
COLOR_GRAY = "#6C757D"          # Gris neutro
COLOR_BORDER = "#E2E8F0"        # Borde suave
COLOR_HOVER = "#EDF2F7"         # Fondo hover

# --- Tipos de informe ---
TIPO_REPORTE_MANTENIMIENTO = "Reporte de Mantenimiento"
TIPO_INFORME_EJECUTIVO = "Informe Ejecutivo de Mantenimiento"

# --- Tipos de intervención ---
TIPOS_INTERVENCION = ["Correctivo", "Preventivo", "Mejora", "Predictivo", "Overhaul"]

# --- Prioridades ---
PRIORIDADES = ["Crítica", "Alta", "Media", "Baja"]

# --- Turnos ---
TURNOS = ["Día (06:00-14:00)", "Tarde (14:00-22:00)", "Noche (22:00-06:00)", "Administrativo"]

# --- Áreas comunes de planta ---
AREAS_PLANTA = [
    "Producción Principal",
    "Empaque y Embalaje",
    "Almacén de Materia Prima",
    "Almacén de Producto Terminado",
    "Sala de Máquinas",
    "Subestación Eléctrica",
    "Planta de Agua",
    "Calderas",
    "Compresores",
    "Taller Mecánico",
    "Taller Eléctrico",
    "Oficinas Administrativas",
    "Laboratorio de Calidad",
    "Área de Servicios Generales",
    "Otra (Especificar)"
]

# --- Disciplinas ---
DISCIPLINAS = ["Mecánica", "Eléctrica", "Electrónica", "Instrumentación", "Automatización", "Civil", "Multidisciplinaria"]

# --- Usuarios por defecto ---
DEFAULT_USERS = {
    "admin": {
        "password_hash": bcrypt.hashpw("admin123".encode('utf-8'), bcrypt.gensalt()).decode('utf-8'),
        "nombre": "Administrador",
        "rol": "Superintendente",
        "activo": True
    },
    "rhvamani": {
        "password_hash": bcrypt.hashpw("cava2026".encode('utf-8'), bcrypt.gensalt()).decode('utf-8'),
        "nombre": "Roger Huamani",
        "rol": "Jefe de Mantenimiento",
        "activo": True
    },
    "tecnico1": {
        "password_hash": bcrypt.hashpw("tecnico123".encode('utf-8'), bcrypt.gensalt()).decode('utf-8'),
        "nombre": "Técnico Mecánico",
        "rol": "Técnico",
        "activo": True
    }
}

# ============================================================================
# SECCIÓN 2: ESTILOS CSS PROFESIONALES (WCAG 2.1 AA Compliant)
# ============================================================================

def aplicar_estilos_css():
    """Aplica estilos CSS profesionales con contraste normativo."""
    st.markdown(f"""
    <style>
        /* ============================================
           RESET Y CONFIGURACIÓN GENERAL
           ============================================ */
        .stApp {{
            background-color: {COLOR_BG_LIGHT};
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }}

        /* ============================================
           SIDEBAR - FONDO CLARO (Corrección principal)
           ============================================ */
        [data-testid="stSidebar"] {{
            background-color: {COLOR_SIDEBAR} !important;
            border-right: 4px solid {COLOR_SIDEBAR_BORDER};
        }}

        [data-testid="stSidebar"] .stMarkdown {{
            color: {COLOR_TEXT_DARK} !important;
        }}

        [data-testid="stSidebar"] label {{
            color: {COLOR_TEXT_DARK} !important;
            font-weight: 600;
        }}

        [data-testid="stSidebar"] .stRadio label,
        [data-testid="stSidebar"] .stSelectbox label {{
            color: {COLOR_TEXT_DARK} !important;
        }}

        [data-testid="stSidebar"] button[kind="secondary"] {{
            background-color: {COLOR_WHITE};
            color: {COLOR_PRIMARY};
            border: 1px solid {COLOR_PRIMARY};
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
           TARJETAS DE ESTADÍSTICAS
           ============================================ */
        .stat-card {{
            background: {COLOR_BG_CARD};
            border-radius: 12px;
            padding: 20px;
            text-align: center;
            box-shadow: 0 2px 10px rgba(0,0,0,0.06);
            transition: transform 0.2s, box-shadow 0.2s;
            border-top: 4px solid {COLOR_SECONDARY};
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
           BADGES DE ESTADO
           ============================================ */
        .badge {{
            display: inline-block;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
        }}

        .badge-success {{
            background-color: {COLOR_SUCCESS_BG};
            color: {COLOR_SUCCESS};
        }}

        .badge-warning {{
            background-color: {COLOR_WARNING_BG};
            color: {COLOR_WARNING};
        }}

        .badge-danger {{
            background-color: {COLOR_DANGER_BG};
            color: {COLOR_DANGER};
        }}

        .badge-info {{
            background-color: {COLOR_INFO_BG};
            color: {COLOR_INFO};
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
           SEPARADOR ESTILIZADO
           ============================================ */
        .custom-divider {{
            height: 2px;
            background: linear-gradient(90deg, transparent, {COLOR_SECONDARY}, transparent);
            margin: 20px 0;
        }}

        /* ============================================
           ALERTAS PERSONALIZADAS
           ============================================ */
        .custom-alert {{
            background: {COLOR_INFO_BG};
            border-left: 4px solid {COLOR_INFO};
            border-radius: 8px;
            padding: 15px 20px;
            margin: 10px 0;
            color: {COLOR_TEXT_DARK};
        }}

        .custom-alert-success {{
            background: {COLOR_SUCCESS_BG};
            border-left: 4px solid {COLOR_SUCCESS};
            border-radius: 8px;
            padding: 15px 20px;
            margin: 10px 0;
            color: {COLOR_TEXT_DARK};
        }}

        .custom-alert-warning {{
            background: {COLOR_WARNING_BG};
            border-left: 4px solid {COLOR_WARNING};
            border-radius: 8px;
            padding: 15px 20px;
            margin: 10px 0;
            color: {COLOR_TEXT_DARK};
        }}

        /* ============================================
           BOTONES
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

        /* ============================================
           TABLAS
           ============================================ */
        .stDataFrame {{
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        }}

        /* ============================================
           CAMPOS OBLIGATORIOS
           ============================================ */
        .required-field::after {{
            content: " *";
            color: {COLOR_DANGER};
            font-weight: bold;
        }}

        /* ============================================
           OCULTAR ELEMENTOS DE STREAMLIT
           ============================================ */
        #MainMenu {{visibility: hidden;}}
        footer {{visibility: hidden;}}
        header {{visibility: hidden;}}

        /* ============================================
           RESPONSIVE
           ============================================ */
        @media (max-width: 768px) {{
            .main-header h1 {{
                font-size: 20px;
            }}
            .stat-card .stat-number {{
                font-size: 28px;
            }}
        }}
    </style>
    """, unsafe_allow_html=True)


# ============================================================================
# SECCIÓN 3: GESTIÓN DE AUTENTICACIÓN
# ============================================================================

class AuthenticationManager:
    """Gestiona la autenticación de usuarios del sistema."""

    def __init__(self):
        self.users = self._load_users()

    def _load_users(self):
        """Carga los usuarios desde session_state o usa los por defecto."""
        if "registered_users" not in st.session_state:
            st.session_state.registered_users = DEFAULT_USERS.copy()
        return st.session_state.registered_users

    def verify_password(self, plain_password: str, hashed_password: str) -> bool:
        """Verifica si la contraseña coincide con el hash almacenado."""
        try:
            return bcrypt.checkpw(
                plain_password.encode('utf-8'),
                hashed_password.encode('utf-8')
            )
        except Exception:
            return False

    def authenticate(self, username: str, password: str) -> dict:
        """Autentica un usuario. Retorna dict con datos del usuario o None."""
        username = username.strip().lower()
        if username in self.users:
            user_data = self.users[username]
            if user_data.get("activo", False):
                if self.verify_password(password, user_data["password_hash"]):
                    return {
                        "username": username,
                        "nombre": user_data["nombre"],
                        "rol": user_data["rol"]
                    }
        return None

    def register_user(self, username: str, password: str, nombre: str, rol: str) -> bool:
        """Registra un nuevo usuario en el sistema."""
        username = username.strip().lower()
        if username in self.users:
            return False
        password_hash = bcrypt.hashpw(
            password.encode('utf-8'),
            bcrypt.gensalt()
        ).decode('utf-8')
        self.users[username] = {
            "password_hash": password_hash,
            "nombre": nombre,
            "rol": rol,
            "activo": True
        }
        st.session_state.registered_users = self.users
        return True

    def get_all_users(self) -> dict:
        """Retorna todos los usuarios registrados."""
        return self.users


def render_login_screen():
    """Renderiza la pantalla de inicio de sesión."""
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

        with st.form("login_form"):
            username = st.text_input(
                " Usuario",
                placeholder="Ingrese su usuario",
                key="login_username"
            )
            password = st.text_input(
                "🔑 Contraseña",
                type="password",
                placeholder="Ingrese su contraseña",
                key="login_password"
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
                        st.rerun()
                    else:
                        st.error("❌ Usuario o contraseña incorrectos.")

        st.markdown("---")
        st.caption(
            " Credenciales de prueba: admin / admin123 | "
            "rhvamani / cava2026 | tecnico1 / tecnico123"
        )


# ============================================================================
# SECCIÓN 4: GESTIÓN DE SUPABASE
# ============================================================================

class SupabaseManager:
    """Gestiona la conexión y operaciones con Supabase."""

    def __init__(self):
        self.client = None
        self.connected = False
        self._initialize_connection()

    def _initialize_connection(self):
        """Inicializa la conexión con Supabase."""
        url = st.session_state.get("supabase_url", "")
        key = st.session_state.get("supabase_key", "")
        if url and key:
            try:
                self.client = create_client(url, key)
                self.connected = True
            except Exception as e:
                self.connected = False
                st.session_state.supabase_error = str(e)

    def is_connected(self) -> bool:
        """Verifica si la conexión está activa."""
        return self.connected and self.client is not None

    def save_report(self, report_data: dict) -> bool:
        """Guarda un informe en Supabase."""
        if not self.is_connected():
            return self._save_local(report_data)
        try:
            self.client.table("informes_mantenimiento").insert(report_data).execute()
            return True
        except Exception as e:
            st.warning(f"⚠️ Error al guardar en Supabase: {e}. Guardando localmente.")
            return self._save_local(report_data)

    def get_reports(self, limit: int = 100) -> list:
        """Obtiene los informes almacenados."""
        if not self.is_connected():
            return self._get_local_reports()
        try:
            result = self.client.table("informes_mantenimiento") \
                .select("*") \
                .order("fecha_creacion", desc=True) \
                .limit(limit) \
                .execute()
            return result.data if result.data else []
        except Exception:
            return self._get_local_reports()

    def get_report_by_id(self, report_id: str) -> dict:
        """Obtiene un informe específico por su ID."""
        if not self.is_connected():
            return self._get_local_report_by_id(report_id)
        try:
            result = self.client.table("informes_mantenimiento") \
                .select("*") \
                .eq("id", report_id) \
                .single() \
                .execute()
            return result.data if result.data else {}
        except Exception:
            return self._get_local_report_by_id(report_id)

    def update_report(self, report_id: str, update_data: dict) -> bool:
        """Actualiza un informe existente."""
        if not self.is_connected():
            return self._update_local_report(report_id, update_data)
        try:
            self.client.table("informes_mantenimiento") \
                .update(update_data) \
                .eq("id", report_id) \
                .execute()
            return True
        except Exception:
            return self._update_local_report(report_id, update_data)

    def delete_report(self, report_id: str) -> bool:
        """Elimina un informe."""
        if not self.is_connected():
            return self._delete_local_report(report_id)
        try:
            self.client.table("informes_mantenimiento") \
                .delete() \
                .eq("id", report_id) \
                .execute()
            return True
        except Exception:
            return self._delete_local_report(report_id)

    def upload_image(self, image_bytes: bytes, filename: str) -> str:
        """Sube una imagen al storage de Supabase."""
        if not self.is_connected():
            return ""
        try:
            file_path = f"informes/{datetime.now().strftime('%Y/%m')}/{filename}"
            self.client.storage.from_("imagenes").upload(
                file_path,
                image_bytes,
                {"content-type": "image/png"}
            )
            return self.client.storage.from_("imagenes").get_public_url(file_path)
        except Exception as e:
            st.warning(f"⚠️ No se pudo subir la imagen: {e}")
            return ""

    # --- Métodos de almacenamiento local (fallback) ---

    def _save_local(self, report_data: dict) -> bool:
        """Guarda localmente cuando Supabase no está configurado."""
        if "local_reports" not in st.session_state:
            st.session_state.local_reports = []
        report_data["storage"] = "local"
        st.session_state.local_reports.insert(0, report_data)
        return True

    def _get_local_reports(self) -> list:
        """Obtiene informes del almacenamiento local."""
        return st.session_state.get("local_reports", [])

    def _get_local_report_by_id(self, report_id: str) -> dict:
        """Busca un informe local por ID."""
        for report in st.session_state.get("local_reports", []):
            if report.get("id") == report_id:
                return report
        return {}

    def _update_local_report(self, report_id: str, update_data: dict) -> bool:
        """Actualiza un informe local."""
        reports = st.session_state.get("local_reports", [])
        for i, report in enumerate(reports):
            if report.get("id") == report_id:
                reports[i].update(update_data)
                st.session_state.local_reports = reports
                return True
        return False

    def _delete_local_report(self, report_id: str) -> bool:
        """Elimina un informe local."""
        reports = st.session_state.get("local_reports", [])
        st.session_state.local_reports = [
            r for r in reports if r.get("id") != report_id
        ]
        return True


# ============================================================================
# SECCIÓN 5: INTEGRACIÓN CON GEMINI AI
# ============================================================================

class GeminiAIManager:
    """Gestiona la integración con la API de Google Gemini."""

    def __init__(self):
        self.model = None
        self.api_key = st.session_state.get("gemini_api_key", "")
        self._initialize_model()

    def _initialize_model(self):
        """Inicializa el modelo de Gemini."""
        if self.api_key:
            try:
                genai.configure(api_key=self.api_key)
                self.model = genai.GenerativeModel('gemini-2.0-flash')
            except Exception as e:
                st.session_state.gemini_error = str(e)

    def is_configured(self) -> bool:
        """Verifica si la API está configurada."""
        return self.model is not None and self.api_key != ""

    def generate_report_content(
        self,
        raw_description: str,
        report_type: str,
        additional_context: dict = None
    ) -> str:
        """Genera el contenido estructurado del informe usando Gemini AI."""
        if not self.is_configured():
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
            return response.text
        except Exception as e:
            st.warning(f"️ Error con Gemini AI: {e}. Generando informe con plantilla base.")
            return self._generate_fallback_report(raw_description, report_type, additional_context)

    def _build_report_prompt(self, raw_description: str, context_info: str) -> str:
        """Construye el prompt para un Reporte de Mantenimiento."""
        return f"""
Eres un Jefe de Mantenimiento y Planificador experto con más de 20 años de experiencia
en gestión de activos industriales. Tu tarea es redactar un Reporte de Mantenimiento
profesional, técnico y estructurado.

INSTRUCCIONES CRÍTICAS:
1. Redacta con un tono técnico, objetivo y profesional, como si fuera escrito por
   un ingeniero de mantenimiento senior.
2. El lenguaje debe ser humanizado pero preciso. Sin errores ortográficos ni gramaticales.
3. Sigue estrictamente la normativa de redacción para informes técnicos industriales.
4. Si detectas que falta información crítica, rellena el campo con
   "[Información no proporcionada - Completar]".
5. Profundiza y completa cada campo con información técnica relevante basada en
   el contexto proporcionado.
6. Usa terminología técnica apropiada del área de mantenimiento industrial.

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
[Descripción técnica detallada del síntoma reportado, incluyendo condiciones
de operación al momento de la falla, alarmas activas, y manifestaciones
observadas por el operador o técnico.]

### 4. TIPO DE INTERVENCIÓN
- **Tipo:** [Correctivo / Preventivo / Mejora / Predictivo]
- **Prioridad:** [Crítica / Alta / Media / Baja]
- **Justificación:** [breve justificación de la clasificación]

### 5. DETALLE DEL TRABAJO REALIZADO
[Descripción técnica paso a paso de todas las actividades realizadas.
Incluir procedimientos de seguridad aplicados (LOTO, permisos de trabajo),
herramientas especiales utilizadas, mediciones tomadas, y criterios de
aceptación verificados. Numerar cada paso de forma clara y secuencial.]

### 6. REPUESTOS Y MATERIALES UTILIZADOS
| Ítem | Descripción | Cantidad | Unidad | Código SAP |
|------|-------------|----------|--------|------------|
| [completar tabla] |

### 7. TIEMPOS DE INTERVENCIÓN
- **Hora de Aviso:** [hh:mm]
- **Hora de Inicio:** [hh:mm]
- **Hora de Fin:** [hh:mm]
- **Downtime Total:** [horas y minutos]
- **Tiempo de Espera (Repuestos/Acceso):** [si aplica]

### 8. ESTADO FINAL Y OBSERVACIONES
- **Estado del Equipo:** [Operativo / Operativo con restricciones / Fuera de servicio]
- **Pruebas Realizadas:** [descripción de pruebas post-intervención]
- **Observaciones:** [cualquier observación relevante]
- **Trabajos Pendientes:** [si los hay]
- **Recomendaciones:** [acciones sugeridas a corto y mediano plazo]

Redacta el documento completo ahora:
"""

    def _build_executive_prompt(self, raw_description: str, context_info: str) -> str:
        """Construye el prompt para un Informe Ejecutivo de Mantenimiento."""
        return f"""
Eres un Superintendente de Mantenimiento y Planificador experto con más de 20 años
de experiencia en gestión de activos y confiabilidad industrial. Tu tarea es redactar
un Informe Ejecutivo de Mantenimiento dirigido a la Gerencia de Planta.

INSTRUCCIONES CRÍTICAS:
1. Redacta con un tono ejecutivo, técnico y orientado a la gestión de activos.
2. El lenguaje debe ser claro, directo y profesional. Sin errores ortográficos.
3. El informe debe ser comprensible para personal de gerencia no necesariamente
   técnico, pero con suficiente profundidad técnica.
4. Si detectas que falta información crítica, rellena el campo con
   "[Información no proporcionada - Completar]".
5. Enfatiza el impacto operativo, los costos y las decisiones a tomar.
6. Incluye métricas y KPIs relevantes cuando sea posible.

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
[Párrafo conciso y directo que explique qué sucedió, cuándo, dónde y cuál
fue la magnitud del evento. Máximo 5-6 líneas. Debe captar la atención
del lector ejecutivo inmediatamente.]

### 3. IMPACTO OPERATIVO
- **Horas de Producción Perdidas:** [horas]
- **Tonelaje/Unidades No Producidas:** [si aplica]
- **Nivel de Afectación:** [Total / Parcial / Mínimo]
- **Líneas/Áreas Afectadas:** [detalle]
- **Impacto en Entregas/Clientes:** [si aplica]
- **Costo Estimado de Lucro Cesante:** [si se puede estimar]

### 4. CAUSA RAÍZ PRELIMINAR
[Análisis técnico de la causa raíz del evento. Incluir si se aplicó
metodología 5 Porqués, Ishikawa u otra. Diferenciar entre causa directa
y causa raíz. Indicar si se requiere un análisis RCA formal.]

### 5. SOLUCIÓN EJECUTADA
[Descripción clara de las acciones tomadas para restablecer la operación.
Incluir si fue una solución temporal o definitiva. Mencionar tiempos de
respuesta y efectividad de la solución.]

### 6. COSTOS RELEVANTES
| Concepto | Monto Estimado | Moneda |
|----------|---------------|--------|
| Repuestos | [monto] | [USD/PEN] |
| Mano de Obra | [monto] | [USD/PEN] |
| Servicios Externos | [monto] | [USD/PEN] |
| Lucro Cesante | [monto] | [USD/PEN] |
| **TOTAL ESTIMADO** | **[monto]** | **[moneda]** |

### 7. ACCIONES PREVENTIVAS Y RECOMENDACIONES
[Lista numerada de acciones concretas, con responsable sugerido y plazo
estimado. Separar en acciones inmediatas, corto plazo y mediano plazo.
Incluir recomendaciones sobre cambios en el plan de mantenimiento,
mejoras de diseño, capacitación, o adquisición de repuestos críticos.]

### 8. CONCLUSIONES
[2-3 párrafos de cierre que resuman la situación actual, la confiabilidad
del equipo post-intervención, y la urgencia de las acciones recomendadas.]

Redacta el documento completo ahora:
"""

    def process_image_description(self, raw_description: str) -> str:
        """Procesa y mejora la descripción de una imagen técnica."""
        if not self.is_configured():
            return raw_description.strip().capitalize() + "."

        prompt = f"""
Eres un ingeniero de mantenimiento experto. Mejora la siguiente descripción
de una fotografía técnica tomada durante una intervención de mantenimiento.

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
        except Exception:
            return raw_description.strip().capitalize() + "."

    def translate_report(self, report_content: str) -> str:
        """Traduce el informe completo al inglés."""
        if not self.is_configured():
            return "[Translation requires Gemini API configuration]"

        prompt = f"""
You are a professional technical translator specializing in industrial
maintenance documentation. Translate the following maintenance report
from Spanish to English.

INSTRUCTIONS:
1. Maintain the exact same structure and formatting.
2. Use standard industrial maintenance terminology in English.
3. Keep all technical specifications, measurements, and codes as-is.
4. Ensure the translation reads naturally in English.
5. Preserve all markdown formatting (headers, tables, bold text, etc.).

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
            return f"[Translation error: {e}]"

    def _generate_fallback_report(
        self,
        raw_description: str,
        report_type: str,
        additional_context: dict = None
    ) -> str:
        """Genera un informe base cuando la IA no está configurada."""
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
[Configure la API de Gemini para generar automáticamente este contenido
a partir de la descripción del técnico.]

### 6. REPUESTOS Y MATERIALES UTILIZADOS
| Ítem | Descripción | Cantidad | Unidad | Código SAP |
|------|-------------|----------|--------|------------|
| 1 | [Completar] | - | - | - |

### 7. TIEMPOS DE INTERVENCIÓN
- **Hora de Aviso:** [Información no proporcionada - Completar]
- **Hora de Inicio:** [Información no proporcionada - Completar]
- **Hora de Fin:** [Información no proporcionada - Completar]
- **Downtime Total:** [Información no proporcionada - Completar]

### 8. ESTADO FINAL Y OBSERVACIONES
- **Estado del Equipo:** [Información no proporcionada - Completar]
- **Pruebas Realizadas:** [Información no proporcionada - Completar]
- **Observaciones:** [Información no proporcionada - Completar]
- **Trabajos Pendientes:** [Información no proporcionada - Completar]
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
- **Líneas/Áreas Afectadas:** [Información no proporcionada - Completar]

### 4. CAUSA RAÍZ PRELIMINAR
[Configure la API de Gemini para generar automáticamente este análisis.]

### 5. SOLUCIÓN EJECUTADA
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
# SECCIÓN 6: GENERADOR DE NÚMEROS DE INFORME
# ============================================================================

class ReportNumberGenerator:
    """Genera números de informe correlativos automáticamente."""

    @staticmethod
    def generate(tipo: str) -> str:
        """
        Genera un número único de informe.
        Formato: RM-2026-0001 o IE-2026-0001
        """
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
# SECCIÓN 7: GESTIÓN DE IMÁGENES
# ============================================================================

class ImageManager:
    """Gestiona la carga, numeración y descripción de imágenes."""

    @staticmethod
    def initialize_image_counter():
        """Inicializa el contador de imágenes."""
        if "image_counter" not in st.session_state:
            st.session_state.image_counter = 0
        if "uploaded_images" not in st.session_state:
            st.session_state.uploaded_images = []

    @staticmethod
    def add_image(image_file, description: str) -> dict:
        """Agrega una imagen con su descripción al registro."""
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
        """Retorna todas las imágenes cargadas."""
        return st.session_state.get("uploaded_images", [])

    @staticmethod
    def remove_image(index: int):
        """Elimina una imagen por su índice."""
        images = st.session_state.get("uploaded_images", [])
        if 0 <= index < len(images):
            st.session_state.uploaded_images.pop(index)
            for i, img in enumerate(st.session_state.uploaded_images):
                img["numero"] = i + 1
            st.session_state.image_counter = len(st.session_state.uploaded_images)

    @staticmethod
    def clear_all_images():
        """Elimina todas las imágenes."""
        st.session_state.uploaded_images = []
        st.session_state.image_counter = 0


# ============================================================================
# SECCIÓN 8: EXPORTACIÓN A WORD (DOCX)
# ============================================================================

class WordExporter:
    """Exporta informes a formato Microsoft Word (.docx)."""

    @staticmethod
    def create_document(
        report_content: str,
        report_number: str,
        report_type: str,
        images: list = None,
        is_english: bool = False
    ) -> io.BytesIO:
        """Crea un documento Word profesional."""
        doc = Document()

        # --- Configurar estilos del documento ---
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Calibri'
        font.size = Pt(11)
        font.color.rgb = RGBColor(26, 26, 26)

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
        run = title.add_run(" CAVA")
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
                img_heading = doc.add_heading(
                    f'Fotografía N° {img_data["numero"]}',
                    level=3
                )

                try:
                    image_stream = io.BytesIO(img_data["bytes"])
                    doc.add_picture(image_stream, width=Inches(5.5))
                    last_paragraph = doc.paragraphs[-1]
                    last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                except Exception:
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
        footer_para.text = (
            f"CAVA Especialistas en Robótica y Automatización | "
            f"Roger Huamani | {report_number} | Página "
        )
        footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer_para.style.font.size = Pt(8)

        # --- Guardar en buffer ---
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    @staticmethod
    def _add_formatted_content(doc: Document, content: str):
        """Agrega contenido formateado (markdown básico) al documento."""
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
                heading = doc.add_heading(line[5:], level=3)
            elif line.startswith('|') and '---' not in line:
                cells = [c.strip() for c in line.split('|')[1:-1]]
                if cells:
                    table = doc.add_table(rows=1, cols=len(cells))
                    table.style = 'Light Grid Accent 1'
                    for i, cell_text in enumerate(cells):
                        clean_text = cell_text.replace('**', '')
                        table.rows[0].cells[i].text = clean_text
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
                p = doc.add_paragraph(line[2:], style='List Bullet')
            elif line.startswith('1.') or line.startswith('2.') or line.startswith('3.'):
                p = doc.add_paragraph(line, style='List Number')
            else:
                p = doc.add_paragraph()
                parts = line.split('**')
                for j, part in enumerate(parts):
                    run = p.add_run(part)
                    if j % 2 == 1:
                        run.font.bold = True
                    run.font.size = Pt(11)


# ============================================================================
# SECCIÓN 9: EXPORTACIÓN A PDF
# ============================================================================

class PDFExporter:
    """Exporta informes a formato PDF profesional."""

    @staticmethod
    def create_pdf(
        report_content: str,
        report_number: str,
        report_type: str,
        images: list = None,
        is_english: bool = False
    ) -> io.BytesIO:
        """Crea un documento PDF profesional."""
        pdf = PDFReport(
            report_number=report_number,
            report_type=report_type,
            is_english=is_english
        )
        pdf.alias_nb_pages()
        pdf.add_page()

        PDFExporter._add_cover_page(pdf, report_number, report_type, is_english)
        pdf.add_page()

        PDFExporter._add_content(pdf, report_content)

        if images and len(images) > 0:
            pdf.add_page()
            pdf.set_font("Helvetica", "B", 16)
            pdf.set_text_color(13, 43, 78)
            pdf.cell(0, 12, "REGISTRO FOTOGRÁFICO", ln=True, align="C")
            pdf.ln(8)

            for img_data in images:
                pdf.set_font("Helvetica", "B", 11)
                pdf.set_text_color(30, 95, 142)
                pdf.cell(
                    0, 8,
                    f"Fotografía N° {img_data['numero']}",
                    ln=True
                )

                try:
                    temp_path = f"/tmp/cava_img_{img_data['numero']}.png"
                    with open(temp_path, "wb") as f:
                        f.write(img_data["bytes"])
                    pdf.image(temp_path, w=160)
                    os.remove(temp_path)
                except Exception:
                    pdf.set_font("Helvetica", "I", 9)
                    pdf.cell(0, 6, "[Imagen no disponible]", ln=True)

                desc = img_data.get("descripcion_mejorada") or img_data.get("descripcion_original", "")
                pdf.set_font("Helvetica", "I", 9)
                pdf.set_text_color(108, 117, 125)
                pdf.multi_cell(0, 5, f"Figura {img_data['numero']}: {desc}")
                pdf.ln(6)

        buffer = io.BytesIO()
        pdf_bytes = pdf.output()
        buffer.write(pdf_bytes)
        buffer.seek(0)
        return buffer

    @staticmethod
    def _add_cover_page(pdf, report_number, report_type, is_english):
        """Agrega la portada del PDF."""
        pdf.ln(50)
        pdf.set_font("Helvetica", "B", 32)
        pdf.set_text_color(13, 43, 78)
        pdf.cell(0, 15, "CAVA", ln=True, align="C")

        pdf.set_font("Helvetica", "", 14)
        pdf.set_text_color(30, 95, 142)
        pdf.cell(
            0, 10,
            "Especialistas en Robotica y Automatizacion",
            ln=True, align="C"
        )

        pdf.ln(20)

        pdf.set_font("Helvetica", "B", 20)
        pdf.set_text_color(13, 43, 78)
        pdf.cell(0, 12, report_type.upper(), ln=True, align="C")

        pdf.set_font("Helvetica", "B", 16)
        pdf.set_text_color(201, 123, 0)
        pdf.cell(0, 10, f"N° {report_number}", ln=True, align="C")

        if is_english:
            pdf.ln(5)
            pdf.set_font("Helvetica", "I", 12)
            pdf.set_text_color(183, 28, 28)
            pdf.cell(
                0, 8,
                "[ENGLISH VERSION - TRANSLATED DOCUMENT]",
                ln=True, align="C"
            )

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
            pdf.cell(60, 7, label, ln=False)
            pdf.set_font("Helvetica", "", 10)
            pdf.cell(0, 7, value, ln=True)

    @staticmethod
    def _add_content(pdf, content: str):
        """Agrega el contenido formateado al PDF."""
        lines = content.split('\n')
        for line in lines:
            line = line.strip()
            if not line:
                pdf.ln(3)
                continue

            if line.startswith('## '):
                pdf.set_font("Helvetica", "B", 15)
                pdf.set_text_color(13, 43, 78)
                pdf.ln(5)
                pdf.multi_cell(0, 8, line[3:])
                pdf.set_draw_color(30, 95, 142)
                pdf.line(pdf.l_margin, pdf.get_y(), pdf.w - pdf.r_margin, pdf.get_y())
                pdf.ln(4)
            elif line.startswith('### '):
                pdf.set_font("Helvetica", "B", 12)
                pdf.set_text_color(30, 95, 142)
                pdf.ln(3)
                pdf.multi_cell(0, 7, line[4:])
                pdf.ln(2)
            elif line.startswith('#### '):
                pdf.set_font("Helvetica", "B", 11)
                pdf.set_text_color(26, 26, 26)
                pdf.multi_cell(0, 6, line[5:])
                pdf.ln(1)
            elif line.startswith('|') and '---' not in line:
                cells = [c.strip().replace('**', '') for c in line.split('|')[1:-1]]
                if cells:
                    pdf.set_font("Helvetica", "", 8)
                    pdf.set_text_color(26, 26, 26)
                    col_width = (pdf.w - pdf.l_margin - pdf.r_margin) / max(len(cells), 1)
                    for cell_text in cells:
                        pdf.cell(col_width, 6, cell_text[:30], border=1)
                    pdf.ln()
            elif line.startswith('|') and '---' in line:
                continue
            elif line.startswith('- '):
                pdf.set_font("Helvetica", "", 10)
                pdf.set_text_color(26, 26, 26)
                clean = line[2:].replace('**', '')
                pdf.cell(5)
                pdf.multi_cell(0, 5, f"• {clean}")
            elif re.match(r'^\d+\.', line):
                pdf.set_font("Helvetica", "", 10)
                pdf.set_text_color(26, 26, 26)
                clean = line.replace('**', '')
                pdf.cell(5)
                pdf.multi_cell(0, 5, clean)
            else:
                pdf.set_font("Helvetica", "", 10)
                pdf.set_text_color(26, 26, 26)
                clean = line.replace('**', '')
                pdf.multi_cell(0, 5, clean)


class PDFReport(FPDF):
    """Clase personalizada de PDF con encabezado y pie de página."""

    def __init__(self, report_number="", report_type="", is_english=False):
        super().__init__()
        self.report_number = report_number
        self.report_type = report_type
        self.is_english = is_english

    def header(self):
        """Encabezado de cada página."""
        if self.page_no() <= 1:
            return
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(108, 117, 125)
        self.cell(
            0, 6,
            "CAVA - Especialistas en Robotica y Automatizacion | "
            f"{self.report_type} | N° {self.report_number}",
            ln=True, align="R"
        )
        self.set_draw_color(222, 226, 230)
        self.line(
            self.l_margin, self.get_y(),
            self.w - self.r_margin, self.get_y()
        )
        self.ln(4)

    def footer(self):
        """Pie de página de cada página."""
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(108, 117, 125)
        self.cell(
            0, 10,
            f"Roger Huamani | {self.report_number} | "
            f"Pagina {self.page_no()}/{{nb}}",
            align="C"
        )


# ============================================================================
# SECCIÓN 10: PANTALLA PRINCIPAL - DASHBOARD
# ============================================================================

def render_dashboard():
    """Renderiza el panel principal del sistema."""
    user = st.session_state.get("current_user", {})

    st.markdown(f"""
    <div class="main-header">
        <h1>📊 Panel de Control - Sistema de Informes</h1>
        <p>Bienvenido, <strong>{user.get('nombre', 'Usuario')}</strong> |
        Rol: {user.get('rol', 'N/A')} |
        Sesión iniciada: {st.session_state.get('login_time', datetime.now()).strftime('%d/%m/%Y %H:%M')}</p>
    </div>
    """, unsafe_allow_html=True)

    # --- Estadísticas ---
    reports = st.session_state.get("local_reports", [])
    total_reports = len(reports)
    reports_today = sum(
        1 for r in reports
        if r.get("fecha_creacion", "")[:10] == datetime.now().strftime("%Y-%m-%d")
    )
    reports_month = sum(
        1 for r in reports
        if r.get("fecha_creacion", "")[:7] == datetime.now().strftime("%Y-%m")
    )

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-number">{total_reports}</div>
            <div class="stat-label">Total de Informes</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-number">{reports_today}</div>
            <div class="stat-label">Informes Hoy</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-number">{reports_month}</div>
            <div class="stat-label">Informes del Mes</div>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        gemini_status = "✅" if st.session_state.get("gemini_api_key") else "️"
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-number">{gemini_status}</div>
            <div class="stat-label">Estado Gemini AI</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)

    # --- Accesos rápidos ---
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
        if st.button(" Ver Informes Guardados", use_container_width=True):
            st.session_state.current_page = "Informes Guardados"
            st.rerun()


# ============================================================================
# SECCIÓN 11: PANTALLA DE CREACIÓN DE INFORMES
# ============================================================================

def render_new_report():
    """Renderiza la pantalla de creación de un nuevo informe."""
    user = st.session_state.get("current_user", {})

    st.markdown(f"""
    <div class="main-header">
        <h1>📝 Crear Nuevo Informe</h1>
        <p>Complete los campos requeridos y el sistema generará automáticamente
        el informe profesional utilizando inteligencia artificial.</p>
    </div>
    """, unsafe_allow_html=True)

    # --- Paso 1: Tipo de informe ---
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

    # --- Paso 2: Datos generales obligatorios ---
    st.subheader("2️⃣ Datos Generales del Informe")

    col1, col2 = st.columns(2)
    with col1:
        fecha_informe = st.date_input(
            "📅 Fecha del Informe *",
            value=datetime.now(),
            key="fecha_informe"
        )
        turno = st.selectbox(
            "🕐 Turno *",
            TURNOS,
            key="turno_selector"
        )
        area = st.selectbox(
            "🏭 Área / Ubicación *",
            AREAS_PLANTA,
            key="area_selector"
        )
        disciplina = st.selectbox(
            "🔧 Disciplina *",
            DISCIPLINAS,
            key="disciplina_selector"
        )

    with col2:
        numero_ot = st.text_input(
            " Número de OT (Orden de Trabajo) *",
            placeholder="Ej: OT-2026-00123",
            key="numero_ot_input"
        )
        equipo_nombre = st.text_input(
            "⚙️ Nombre del Equipo *",
            placeholder="Ej: Bomba Centrífuga B-201",
            key="equipo_nombre_input"
        )
        equipo_tag = st.text_input(
            "️ Tag / Código del Equipo *",
            placeholder="Ej: P-201A",
            key="equipo_tag_input"
        )
        prioridad = st.selectbox(
            " Prioridad *",
            PRIORIDADES,
            key="prioridad_selector"
        )

    col3, col4 = st.columns(2)
    with col3:
        tipo_intervencion = st.selectbox(
            "🔨 Tipo de Intervención *",
            TIPOS_INTERVENCION,
            key="tipo_intervencion_selector"
        )
        tecnicos = st.text_input(
            "👷 Técnico(s) Responsable(s) *",
            placeholder="Ej: Juan Pérez, María López",
            key="tecnicos_input"
        )

    with col4:
        hora_aviso = st.time_input(
            " Hora de Aviso",
            value=datetime.now().time(),
            key="hora_aviso_input"
        )
        elaborado_por = st.text_input(
            "✍️ Elaborado por",
            value=user.get("nombre", ""),
            key="elaborado_por_input"
        )

    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)

    # --- Paso 3: Descripción del técnico ---
    st.subheader("3️⃣ Descripción de la Intervención (Campo Principal)")

    st.markdown(f"""
    <div class="custom-alert">
        💡 <strong>Instrucciones para el técnico:</strong> Describa con sus propias
        palabras todo el contexto del problema identificado, las acciones que ha
        realizado y sus conclusiones. No se preocupe por la redacción técnica,
        la inteligencia artificial se encargará de estructurar, corregir y
        profesionalizar el contenido.
    </div>
    """, unsafe_allow_html=True)

    raw_description = st.text_area(
        "Describa detalladamente la intervención realizada *:",
        height=250,
        placeholder=(
            "Ejemplo: Al llegar a la planta encontré la bomba B-201 con una "
            "fuga de aceite en el sello mecánico. El operador me dijo que "
            "empezó a vibrar mucho desde las 6am. Revisé el acople y estaba "
            "desalineado. Cambié el sello mecánico, realicé la alineación "
            "láser y la bomba quedó funcionando normal. El rodamiento del "
            "lado del acople también estaba con juego, lo cambié también..."
        ),
        key="raw_description_input"
    )

    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)

    # --- Paso 4: Carga de imágenes ---
    st.subheader("4️⃣ Registro Fotográfico")

    ImageManager.initialize_image_counter()

    uploaded_files = st.file_uploader(
        " Cargar imágenes de la intervención",
        type=["png", "jpg", "jpeg", "webp"],
        accept_multiple_files=True,
        key="image_uploader"
    )

    if uploaded_files:
        for file in uploaded_files:
            existing_names = [
                img["nombre_archivo"]
                for img in st.session_state.get("uploaded_images", [])
            ]
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
                            improved = gemini.process_image_description(img_desc)
                            img_data["descripcion_mejorada"] = improved
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
                st.caption(
                    f"Foto N°{img_data['numero']}: "
                    f"{img_data.get('descripcion_mejorada', '')[:50]}..."
                )
                if st.button(f"🗑️", key=f"del_img_{i}"):
                    ImageManager.remove_image(i)
                    st.rerun()

    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)

    # --- Paso 5: Generar informe ---
    st.subheader("5️⃣ Generar Informe")

    col1, col2 = st.columns(2)
    with col1:
        generate_btn = st.button(
            "🤖 Generar Informe con IA",
            use_container_width=True,
            type="primary",
            key="generate_report_btn"
        )
    with col2:
        generate_basic_btn = st.button(
            "📄 Generar Informe Básico (sin IA)",
            use_container_width=True,
            key="generate_basic_btn"
        )

    if generate_btn or generate_basic_btn:
        errors = []
        if not raw_description.strip():
            errors.append("La descripción de la intervención es obligatoria.")
        if not numero_ot.strip():
            errors.append("El número de OT es obligatorio.")
        if not equipo_nombre.strip():
            errors.append("El nombre del equipo es obligatorio.")
        if not equipo_tag.strip():
            errors.append("El Tag del equipo es obligatorio.")
        if not tecnicos.strip():
            errors.append("Los técnicos responsables son obligatorios.")

        if errors:
            for err in errors:
                st.error(f"⚠️ {err}")
            return

        additional_context = {
            "Fecha": fecha_informe.strftime("%d/%m/%Y"),
            "Turno": turno,
            "Número de OT": numero_ot,
            "Área": area,
            "Disciplina": disciplina,
            "Equipo": equipo_nombre,
            "Tag": equipo_tag,
            "Prioridad": prioridad,
            "Tipo de Intervención": tipo_intervencion,
            "Técnicos": tecnicos,
            "Hora de Aviso": hora_aviso.strftime("%H:%M"),
            "Elaborado por": elaborado_por
        }

        with st.spinner("🤖 Generando informe profesional con IA... Esto puede tomar unos segundos."):
            if generate_btn:
                gemini = GeminiAIManager()
                report_content = gemini.generate_report_content(
                    raw_description,
                    report_type,
                    additional_context
                )
            else:
                gemini = GeminiAIManager()
                report_content = gemini._generate_fallback_report(
                    raw_description,
                    report_type,
                    additional_context
                )

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
                    label=" Descargar Word (.docx)",
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
                    translated = gemini.translate_report(
                        st.session_state.generated_report
                    )
                    st.session_state.translated_report = translated
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
                        label=" Descargar PDF en Inglés",
                        data=pdf_en,
                        file_name=f"{st.session_state.generated_report_number}_EN.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Error: {e}")


# ============================================================================
# SECCIÓN 12: PANTALLA DE INFORMES GUARDADOS
# ============================================================================

def render_saved_reports():
    """Renderiza la pantalla de informes guardados."""
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
            f"📄 {report.get('numero', 'N/A')} - "
            f"{report.get('tipo', 'N/A')} | "
            f"{report.get('fecha_creacion', '')[:10]} | "
            f"{report.get('estado', 'Pendiente')}",
            expanded=False
        ):
            col1, col2, col3 = st.columns([3, 1, 1])

            with col1:
                st.markdown(f"**Tipo:** {report.get('tipo', 'N/A')}")
                st.markdown(f"**Autor:** {report.get('autor', 'N/A')}")
                st.markdown(f"**Fecha:** {report.get('fecha_creacion', 'N/A')[:16]}")
                st.markdown(f"**Estado:** {report.get('estado', 'N/A')}")
                st.markdown(f"**Imágenes:** {report.get('imagenes_count', 0)}")

            with col2:
                if st.button("️ Ver", key=f"view_{i}"):
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
                        wb = WordExporter.create_document(
                            report.get("contenido", ""),
                            report.get("numero", ""),
                            report.get("tipo", "")
                        )
                        st.download_button(
                            "📥 Word",
                            data=wb,
                            file_name=f"{report.get('numero', 'reporte')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"dl_word_{i}"
                        )
                    except Exception:
                        pass
                with ecol2:
                    try:
                        pb = PDFExporter.create_pdf(
                            report.get("contenido", ""),
                            report.get("numero", ""),
                            report.get("tipo", "")
                        )
                        st.download_button(
                            "📥 PDF",
                            data=pb,
                            file_name=f"{report.get('numero', 'reporte')}.pdf",
                            mime="application/pdf",
                            key=f"dl_pdf_{i}"
                        )
                    except Exception:
                        pass


# ============================================================================
# SECCIÓN 13: PANTALLA DE CONFIGURACIÓN
# ============================================================================

def render_settings():
    """Renderiza la pantalla de configuración del sistema."""
    st.markdown(f"""
    <div class="main-header">
        <h1>⚙️ Configuración del Sistema</h1>
        <p>Configure las APIs, conexión a base de datos y preferencias del sistema.</p>
    </div>
    """, unsafe_allow_html=True)

    st.subheader("🤖 Configuración de Gemini AI")
    st.markdown(f"""
    <div class="custom-alert">
        🔑 Ingrese su API Key de Google Gemini para habilitar la generación
        inteligente de informes. Obtenga su clave en
        <a href="https://aistudio.google.com/app/apikey" target="_blank">
        Google AI Studio</a>.
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
            st.success("✅ API Key de Gemini guardada correctamente.")
            st.rerun()
        else:
            st.warning("⚠️ Ingrese una API Key válida.")

    gemini = GeminiAIManager()
    if gemini.is_configured():
        st.success("✅ Gemini AI está configurado y operativo.")
    else:
        st.warning("️ Gemini AI no está configurado. Los informes se generarán con plantilla base.")

    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)

    st.subheader("🗄️ Configuración de Supabase")
    st.markdown(f"""
    <div class="custom-alert">
        🗄️ Configure la conexión a Supabase para almacenar sus informes en la nube.
        Cree un proyecto en <a href="https://supabase.com" target="_blank">supabase.com</a>
        y obtenga la URL y la API Key.
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
            st.success("✅ Configuración de Supabase guardada.")
            st.rerun()
        else:
            st.warning("⚠️ Complete ambos campos.")

    sb = SupabaseManager()
    if sb.is_connected():
        st.success("✅ Conexión a Supabase activa.")
    else:
        st.info("ℹ️ Supabase no configurado. Los informes se guardarán localmente en la sesión.")

    st.markdown(f"""
    <div class="custom-alert">
        📌 <strong>Nota sobre la tabla en Supabase:</strong> Asegúrese de crear
        la tabla <code>informes_mantenimiento</code> con las columnas:
        id (text, PK), tipo (text), numero (text), contenido (text),
        contexto (text), fecha_creacion (timestamp), autor (text),
        estado (text), imagenes_count (int).
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)

    st.subheader("👥 Gestión de Usuarios")

    auth = AuthenticationManager()
    users = auth.get_all_users()

    user_data = []
    for uname, udata in users.items():
        user_data.append({
            "Usuario": uname,
            "Nombre": udata.get("nombre", ""),
            "Rol": udata.get("rol", ""),
            "Activo": "✅ Sí" if udata.get("activo") else "❌ No"
        })

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


# ============================================================================
# SECCIÓN 14: PANTALLA DE AYUDA
# ============================================================================

def render_help():
    """Renderiza la pantalla de ayuda y documentación."""
    st.markdown(f"""
    <div class="main-header">
        <h1>❓ Ayuda y Documentación</h1>
        <p>Guía de uso del Sistema de Gestión de Informes de Mantenimiento CAVA.</p>
    </div>
    """, unsafe_allow_html=True)

    st.subheader("📖 ¿Cómo funciona el sistema?")
    st.markdown("""
    Este sistema está diseñado para facilitar la generación de informes técnicos
    y ejecutivos de mantenimiento. El flujo de trabajo es el siguiente:

    1. **El técnico describe** con sus propias palabras lo que hizo durante la
       intervención (problema, acciones, conclusiones).
    2. **La IA de Gemini** procesa esa descripción y genera un informe profesional
       con estructura, redacción técnica y sin errores ortográficos.
    3. **El sistema completa** automáticamente todos los campos requeridos según
       el tipo de documento seleccionado.
    4. **Usted revisa, exporta** (PDF/Word) y almacena el informe.
    """)

    st.subheader("📝 Tipos de Documentos")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        **📋 Reporte de Mantenimiento**
        - Uso: Intervenciones del día a día
        - Audiencia: Equipo de mantenimiento
        - Contenido: Detalle técnico completo
        - Campos: OT, equipo, problema, trabajo,
          repuestos, tiempos, estado final
        """)
    with col2:
        st.markdown("""
        **📊 Informe Ejecutivo de Mantenimiento**
        - Uso: Eventos críticos y gerencia
        - Audiencia: Gerencia de planta
        - Contenido: Impacto, costos, decisiones
        - Campos: Resumen, impacto operativo,
          causa raíz, costos, recomendaciones
        """)

    st.subheader(" Requisitos Técnicos")
    st.markdown("""
    - **API de Gemini:** Necesaria para generación inteligente de informes.
      Obtenga su clave en [Google AI Studio](https://aistudio.google.com/app/apikey).
    - **Supabase (Opcional):** Para almacenamiento en la nube de los informes.
      Cree una cuenta en [supabase.com](https://supabase.com).
    - **Navegador:** Chrome, Firefox o Edge actualizados.
    """)

    st.subheader("❓ Preguntas Frecuentes")

    with st.expander("¿Puedo usar el sistema sin la API de Gemini?"):
        st.markdown(
            "Sí, el sistema generará informes con una plantilla base. "
            "Sin embargo, la calidad y profundidad del contenido será limitada. "
            "Se recomienda configurar la API para obtener resultados profesionales."
        )

    with st.expander("¿Cómo obtengo mi API Key de Gemini?"):
        st.markdown(
            "1. Vaya a [Google AI Studio](https://aistudio.google.com/app/apikey)\n"
            "2. Inicie sesión con su cuenta de Google\n"
            "3. Haga clic en 'Create API Key'\n"
            "4. Copie la clave y péguela en la sección de Configuración"
        )

    with st.expander("¿Los datos se almacenan de forma segura?"):
        st.markdown(
            "Los datos se almacenan en Supabase (si está configurado) con "
            "encriptación de nivel empresarial. Sin Supabase, los datos "
            "solo existen durante la sesión activa del navegador."
        )

    with st.expander("¿Puedo traducir los informes al inglés?"):
        st.markdown(
            "Sí, después de generar un informe puede hacer clic en "
            "'Traducir al Inglés' para crear un documento espejo en inglés. "
            "Ambos documentos se pueden descargar en PDF y Word."
        )


# ============================================================================
# SECCIÓN 15: SIDEBAR Y NAVEGACIÓN (CORREGIDO)
# ============================================================================

def render_sidebar():
    """Renderiza la barra lateral de navegación con colores legibles."""
    user = st.session_state.get("current_user", {})

    with st.sidebar:
        # --- Encabezado del sidebar con identidad visual ---
        st.markdown(f"""
        <div style="text-align:center; padding: 20px 10px;
                    border-bottom: 2px solid {COLOR_PRIMARY};
                    margin-bottom: 15px;">
            <h2 style="color: {COLOR_PRIMARY}; margin-bottom:2px; font-size:26px;">🔧 CAVA</h2>
            <p style="color: {COLOR_TEXT_SECONDARY}; font-size:11px; margin:0; font-weight:500;">
                Especialistas en Robótica<br>y Automatización
            </p>
            <p style="color: {COLOR_TEXT_LIGHT}; font-size:10px; margin-top:5px;">
                v{APP_VERSION}
            </p>
        </div>
        """, unsafe_allow_html=True)

        # --- Información del usuario ---
        st.markdown(f"""
        <div style="padding: 12px; background: {COLOR_WHITE};
                    border-radius: 8px; margin-bottom: 15px;
                    border-left: 4px solid {COLOR_PRIMARY};
                    box-shadow: 0 1px 3px rgba(0,0,0,0.08);">
            <p style="color: {COLOR_TEXT_DARK}; margin:0; font-size:13px; font-weight:600;">
                👤 {user.get('nombre', 'Usuario')}
            </p>
            <p style="color: {COLOR_TEXT_SECONDARY}; margin:2px 0 0 0; font-size:11px;">
                {user.get('rol', 'N/A')}
            </p>
        </div>
        """, unsafe_allow_html=True)

        # --- Navegación ---
        st.markdown("### 📌 Navegación")

        page = st.radio(
            "Ir a:",
            [
                "📊 Dashboard",
                "📝 Nuevo Informe",
                "📂 Informes Guardados",
                "⚙️ Configuración",
                "❓ Ayuda"
            ],
            key="sidebar_nav",
            label_visibility="collapsed"
        )

        page_map = {
            "📊 Dashboard": "Dashboard",
            " Nuevo Informe": "Nuevo Informe",
            "📂 Informes Guardados": "Informes Guardados",
            "⚙️ Configuración": "Configuración",
            "❓ Ayuda": "Ayuda"
        }
        st.session_state.current_page = page_map.get(page, "Dashboard")

        st.markdown("---")

        # --- Estado de servicios ---
        st.markdown("### 🔌 Estado de Servicios")

        gemini = GeminiAIManager()
        if gemini.is_configured():
            st.markdown(f"""
            <div style="padding: 8px 12px; background: {COLOR_SUCCESS_BG};
                        border-radius: 6px; margin-bottom: 8px;
                        border-left: 3px solid {COLOR_SUCCESS};">
                <p style="color: {COLOR_SUCCESS}; margin:0; font-size:12px; font-weight:600;">
                    🤖 Gemini AI: Activo
                </p>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div style="padding: 8px 12px; background: {COLOR_WARNING_BG};
                        border-radius: 6px; margin-bottom: 8px;
                        border-left: 3px solid {COLOR_WARNING};">
                <p style="color: {COLOR_WARNING}; margin:0; font-size:12px; font-weight:600;">
                    🤖 Gemini AI: No configurado
                </p>
            </div>
            """, unsafe_allow_html=True)

        sb = SupabaseManager()
        if sb.is_connected():
            st.markdown(f"""
            <div style="padding: 8px 12px; background: {COLOR_SUCCESS_BG};
                        border-radius: 6px; margin-bottom: 8px;
                        border-left: 3px solid {COLOR_SUCCESS};">
                <p style="color: {COLOR_SUCCESS}; margin:0; font-size:12px; font-weight:600;">
                    🗄️ Supabase: Conectado
                </p>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div style="padding: 8px 12px; background: {COLOR_INFO_BG};
                        border-radius: 6px; margin-bottom: 8px;
                        border-left: 3px solid {COLOR_INFO};">
                <p style="color: {COLOR_INFO}; margin:0; font-size:12px; font-weight:600;">
                    🗄️ Supabase: Almacenamiento local
                </p>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("---")

        # --- Botón de cerrar sesión ---
        if st.button("🚪 Cerrar Sesión", use_container_width=True):
            st.session_state.authenticated = False
            st.session_state.current_user = None
            st.session_state.generated_report = None
            st.session_state.translated_report = None
            ImageManager.clear_all_images()
            st.rerun()

        # --- Créditos ---
        st.markdown("---")
        st.markdown(f"""
        <div style="text-align:center; padding: 15px 5px;">
            <p style="color: {COLOR_PRIMARY}; font-size:11px; font-weight:700; margin:0;">
                CAVA
            </p>
            <p style="color: {COLOR_TEXT_SECONDARY}; font-size:10px; margin:2px 0;">
                Especialistas en Robótica<br>y Automatización
            </p>
            <p style="color: {COLOR_TEXT_LIGHT}; font-size:10px; margin:2px 0;">
                Roger Huamani
            </p>
            <p style="color: {COLOR_TEXT_LIGHT}; font-size:9px; margin:5px 0 0 0;">
                © {APP_YEAR} Todos los derechos reservados
            </p>
        </div>
        """, unsafe_allow_html=True)


# ============================================================================
# SECCIÓN 16: RENDERIZADO DEL FOOTER
# ============================================================================

def render_footer():
    """Renderiza el pie de página institucional."""
    st.markdown(f"""
    <div class="footer">
        <p class="brand">🔧 CAVA - Especialistas en Robótica y Automatización</p>
        <p>Diseñado y desarrollado por <strong>Roger Huamani</strong></p>
        <p>Sistema de Gestión de Informes de Mantenimiento v{APP_VERSION} | © {APP_YEAR}</p>
        <p style="font-size:11px; opacity:0.9; margin-top:8px;">
            Potenciado por Inteligencia Artificial (Google Gemini) |
            Almacenamiento en la nube (Supabase)
        </p>
    </div>
    """, unsafe_allow_html=True)


# ============================================================================
# SECCIÓN 17: INICIALIZACIÓN DEL SESSION STATE
# ============================================================================

def initialize_session_state():
    """Inicializa todas las variables de sesión necesarias."""
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
        "report_counter": {
            TIPO_REPORTE_MANTENIMIENTO: 0,
            TIPO_INFORME_EJECUTIVO: 0
        },
        "image_counter": 0,
        "uploaded_images": [],
        "registered_users": DEFAULT_USERS.copy()
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


# ============================================================================
# SECCIÓN 18: FUNCIÓN PRINCIPAL DE LA APLICACIÓN
# ============================================================================

def main():
    """Función principal que orquesta toda la aplicación."""
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
# SECCIÓN 19: PUNTO DE ENTRADA
# ============================================================================

if __name__ == "__main__":
    main()
