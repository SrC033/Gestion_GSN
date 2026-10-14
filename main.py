#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Sistema de Sincronización Excel → Google Sheets → Dashboard GSN
Grupo GSN - Gestión de Distribución
Archivo principal: main.py
"""

import os
import sys
import json
import time
import threading
import logging
import webbrowser
import subprocess
import pandas as pd
from datetime import datetime
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor

# Flask y SocketIO
from flask import Flask, render_template, request, jsonify
from flask_socketio import SocketIO, emit

# Google API
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError

# Watchdog
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Intentar importar webview (opcional, para ventana nativa)
try:
    import webview
    WEBVIEW_AVAILABLE = True
except ImportError:
    WEBVIEW_AVAILABLE = False
    print('⚠️ PyWebView no instalado. Usando navegador predeterminado.')

# ============================================================
# CONFIGURACIÓN
# ============================================================

BASE_DIR = Path(r'C:\Users\PC\Documents\Gestion_GSN')
EXCEL_DIR = BASE_DIR / 'tabla_excel'
LOGS_DIR = BASE_DIR / 'logs'
TEMPLATES_DIR = BASE_DIR / 'templates'

# Archivos
CREDENTIALS_FILE = BASE_DIR / 'credentials.json'
TOKEN_FILE = BASE_DIR / 'token.json'

# Google Sheets (se actualizará desde la interfaz)
SPREADSHEET_ID = None
SHEET_NAME = 'BASE DATOS'

# Scopes OAuth
SCOPES = [
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/spreadsheets'
]

# Variables globales
EXCEL_PATH = None
MONITORING_ACTIVE = False

# ============================================================
# CREAR CARPETAS
# ============================================================

for dir_path in [EXCEL_DIR, LOGS_DIR, TEMPLATES_DIR]:
    dir_path.mkdir(parents=True, exist_ok=True)

# ============================================================
# CONFIGURACIÓN DE LOGS
# ============================================================

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOGS_DIR / f'sync_{datetime.now().strftime("%Y%m%d")}.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ============================================================
# APLICACIÓN FLASK
# ============================================================

app = Flask(__name__)
app.config['SECRET_KEY'] = 'gsn_dashboard_sync_2026'
socketio = SocketIO(app, cors_allowed_origins="*", async_mode='threading')

# Estado global
estado = {
    'conectado': False,
    'monitoreando': False,
    'ultima_sincronizacion': None,
    'total_registros': 0,
    'clientes': [],
    'archivo_excel': None,
    'spreadsheet_id': None,
    'mensaje': '💡 Configura la URL y KEY, luego presiona "Conectar"'
}

# Servicios de Google
drive_service = None
sheets_service = None
observer = None
executor = ThreadPoolExecutor(max_workers=2)

# ============================================================
# AUTENTICACIÓN GOOGLE
# ============================================================

def autenticar_google():
    """Autentica con Google usando OAuth 2.0"""
    global drive_service, sheets_service
    
    creds = None
    
    if TOKEN_FILE.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
        except Exception as e:
            logger.warning(f'Error al cargar token: {e}')
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                logger.info('✅ Token refrescado')
            except Exception as e:
                logger.warning(f'Error al refrescar token: {e}')
                creds = None
        
        if not creds:
            if not CREDENTIALS_FILE.exists():
                estado['mensaje'] = '❌ No se encontró credentials.json'
                logger.error('❌ No se encontró el archivo de credenciales')
                return False
            
            try:
                flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
                creds = flow.run_local_server(port=0)
                logger.info('✅ Autenticación exitosa')
                
                with open(TOKEN_FILE, 'w') as token:
                    token.write(creds.to_json())
                    
            except Exception as e:
                logger.error(f'❌ Error en autenticación: {e}')
                estado['mensaje'] = f'❌ Error de autenticación: {str(e)}'
                return False
    
    try:
        drive_service = build('drive', 'v3', credentials=creds)
        sheets_service = build('sheets', 'v4', credentials=creds)
        estado['conectado'] = True
        estado['mensaje'] = '✅ Conectado a Google'
        return True
    except Exception as e:
        logger.error(f'❌ Error al crear servicios: {e}')
        estado['mensaje'] = f'❌ Error al conectar: {str(e)}'
        return False

# ============================================================
# FUNCIONES DE SINCRONIZACIÓN
# ============================================================

def leer_excel():
    """Lee el archivo Excel local - detecta automáticamente la pestaña"""
    global EXCEL_PATH
    
    if not EXCEL_PATH or not EXCEL_PATH.exists():
        estado['mensaje'] = f'❌ No se encuentra el archivo'
        return None
    
    try:
        # Leer todas las hojas para identificar cuál tiene datos
        excel_data = pd.read_excel(EXCEL_PATH, sheet_name=None, engine='openpyxl')
        
        # Buscar la hoja que contenga datos (no vacía)
        df = None
        for sheet_name, sheet_df in excel_data.items():
            # Eliminar filas vacías
            sheet_df = sheet_df.dropna(how='all')
            if len(sheet_df) > 0:
                df = sheet_df
                logger.info(f'📋 Usando pestaña: {sheet_name}')
                break
        
        if df is None:
            estado['mensaje'] = '❌ No se encontraron datos en el Excel'
            return None
        
        df = df.fillna('')
        
        # Identificar columna de cliente
        cliente_col = None
        for col in df.columns:
            col_upper = str(col).upper().strip()
            if 'SUB CLIENTE' in col_upper or 'CLIENTE' in col_upper:
                cliente_col = col
                break
        
        clientes = []
        if cliente_col:
            clientes = df[cliente_col].dropna().unique().tolist()
            clientes = [str(c).strip() for c in clientes if str(c).strip()]
        
        estado['total_registros'] = len(df)
        estado['clientes'] = clientes[:30]
        
        logger.info(f'✅ Excel leído: {len(df)} filas, {len(clientes)} clientes')
        return df
        
    except Exception as e:
        logger.error(f'❌ Error al leer Excel: {e}')
        estado['mensaje'] = f'❌ Error al leer: {str(e)}'
        return None

def actualizar_google_sheets(df):
    """Actualiza Google Sheets con los datos del DataFrame"""
    global sheets_service, SPREADSHEET_ID
    
    if sheets_service is None:
        logger.error('❌ No hay conexión con Google')
        return False
    
    if not SPREADSHEET_ID:
        logger.error('❌ No hay SPREADSHEET_ID configurado')
        estado['mensaje'] = '❌ No hay ID del Google Sheet configurado'
        return False
    
    try:
        # Convertir DataFrame a lista de listas
        values = [df.columns.tolist()] + df.values.tolist()
        
        # Obtener información de la hoja para ver si existe la pestaña
        try:
            sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
            sheets = sheet_metadata.get('sheets', [])
            sheet_names = [s['properties']['title'] for s in sheets]
            
            # Si la pestaña no existe, crearla
            if SHEET_NAME not in sheet_names:
                logger.info(f'📝 Creando pestaña "{SHEET_NAME}"...')
                add_sheet_request = {
                    'addSheet': {
                        'properties': {
                            'title': SHEET_NAME
                        }
                    }
                }
                sheets_service.spreadsheets().batchUpdate(
                    spreadsheetId=SPREADSHEET_ID,
                    body={'requests': [add_sheet_request]}
                ).execute()
                logger.info(f'✅ Pestaña "{SHEET_NAME}" creada')
        except Exception as e:
            logger.warning(f'⚠️ No se pudo verificar pestaña: {e}')
        
        # Limpiar hoja existente
        try:
            sheets_service.spreadsheets().values().clear(
                spreadsheetId=SPREADSHEET_ID,
                range=f"'{SHEET_NAME}'!A:ZZ"
            ).execute()
        except Exception as e:
            logger.warning(f'⚠️ No se pudo limpiar: {e}')
        
        # Escribir nuevos datos
        body = {
            'values': values,
            'majorDimension': 'ROWS'
        }
        
        result = sheets_service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{SHEET_NAME}'!A1",
            valueInputOption='USER_ENTERED',
            body=body
        ).execute()
        
        estado['ultima_sincronizacion'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        estado['mensaje'] = f'✅ Actualizado: {len(df)} registros'
        
        logger.info(f'✅ Google Sheets actualizado: {result.get("updatedCells")} celdas')
        return True
        
    except HttpError as error:
        logger.error(f'❌ Error HTTP: {error}')
        estado['mensaje'] = f'❌ Error HTTP: {error}'
        return False
    except Exception as e:
        logger.error(f'❌ Error al actualizar Sheets: {e}')
        estado['mensaje'] = f'❌ Error: {str(e)}'
        return False

def sincronizar():
    """Ejecuta la sincronización completa - versión mejorada"""
    logger.info('🔄 Iniciando sincronización...')
    estado['mensaje'] = '🔄 Sincronizando...'
    socketio.emit('estado_actualizado', estado)
    
    # Verificar que el archivo existe y está accesible
    if not EXCEL_PATH or not EXCEL_PATH.exists():
        estado['mensaje'] = f'❌ No se encuentra el archivo'
        socketio.emit('estado_actualizado', estado)
        return False
    
    # Verificar que el archivo no está siendo usado por otra aplicación
    try:
        with open(EXCEL_PATH, 'rb') as f:
            pass
    except PermissionError:
        estado['mensaje'] = '⚠️ El archivo está siendo usado. Cierra Excel y prueba de nuevo.'
        socketio.emit('estado_actualizado', estado)
        return False
    
    # Leer Excel
    df = leer_excel()
    if df is None:
        socketio.emit('estado_actualizado', estado)
        return False
    
    # Verificar autenticación
    if sheets_service is None:
        if not autenticar_google():
            estado['mensaje'] = '❌ Error de autenticación'
            socketio.emit('estado_actualizado', estado)
            return False
    
    # Actualizar Google Sheets
    success = actualizar_google_sheets(df)
    socketio.emit('estado_actualizado', estado)
    
    if success:
        socketio.emit('sincronizacion_completa', {
            'total': estado['total_registros'],
            'clientes': estado['clientes'],
            'timestamp': estado['ultima_sincronizacion']
        })
        logger.info('✅ Sincronización completada exitosamente')
    else:
        logger.error('❌ Falló la sincronización')
    
    return success

# ============================================================
# MONITOREO DE ARCHIVOS (WATCHDOG - CORREGIDO)
# ============================================================

class ExcelEventHandler(FileSystemEventHandler):
    """Manejador de eventos para cambios en el Excel"""
    
    def __init__(self):
        self.last_modified = 0
        self.cooldown = 5  # Segundos
        self.pending = False
        self.last_size = -1
    
    def on_modified(self, event):
        if event.src_path == str(EXCEL_PATH):
            current_time = time.time()
            
            # Evitar múltiples eventos seguidos
            if (current_time - self.last_modified) < self.cooldown:
                return
            
            # Verificar que el archivo no esté siendo escrito
            try:
                current_size = os.path.getsize(event.src_path)
                if current_size == self.last_size:
                    return
                self.last_size = current_size
            except:
                return
            
            self.last_modified = current_time
            logger.info(f'📝 Detectado cambio en: {EXCEL_PATH.name}')
            estado['mensaje'] = f'📝 Detectado cambio en Excel'
            socketio.emit('estado_actualizado', estado)
            
            # Esperar a que el archivo esté completamente guardado
            time.sleep(3)
            
            # Ejecutar sincronización
            if estado['monitoreando']:
                logger.info('🔄 Ejecutando sincronización automática...')
                socketio.start_background_task(sincronizar)
    
    def on_created(self, event):
        """Cuando se crea un archivo nuevo"""
        if event.src_path == str(EXCEL_PATH):
            logger.info(f'📄 Archivo creado: {EXCEL_PATH.name}')
            time.sleep(2)
            if estado['monitoreando']:
                socketio.start_background_task(sincronizar)

def iniciar_monitoreo():
    """Inicia el monitoreo del archivo Excel"""
    global observer, estado
    
    if not EXCEL_PATH or not EXCEL_PATH.exists():
        estado['mensaje'] = f'❌ No se encuentra el archivo'
        logger.error(f'❌ El archivo Excel no existe: {EXCEL_PATH}')
        return False
    
    if observer and observer.is_alive():
        estado['mensaje'] = '⚠️ Ya está monitoreando'
        return True
    
    # Autenticar si es necesario
    if sheets_service is None:
        if not autenticar_google():
            estado['mensaje'] = '❌ Error de autenticación'
            return False
    
    try:
        event_handler = ExcelEventHandler()
        observer = Observer()
        observer.schedule(event_handler, path=str(EXCEL_PATH.parent), recursive=False)
        observer.start()
        
        estado['monitoreando'] = True
        estado['mensaje'] = '👀 Monitoreando Excel...'
        logger.info('👀 Monitoreo iniciado')
        
        # Sincronización inicial después de 1 segundo
        def sync_inicial():
            time.sleep(1)
            if estado['monitoreando']:
                sincronizar()
        
        threading.Thread(target=sync_inicial, daemon=True).start()
        
        return True
        
    except Exception as e:
        logger.error(f'❌ Error al iniciar monitoreo: {e}')
        estado['mensaje'] = f'❌ Error: {str(e)}'
        return False

def detener_monitoreo():
    """Detiene el monitoreo"""
    global observer, estado
    
    if observer and observer.is_alive():
        observer.stop()
        observer.join()
        observer = None
        
    estado['monitoreando'] = False
    logger.info('⏹️ Monitoreo detenido')

# ============================================================
# RUTAS FLASK
# ============================================================

@app.route('/')
def index():
    """Página principal"""
    return render_template('index.html')

@app.route('/api/estado')
def api_estado():
    """API: obtener estado actual"""
    return jsonify(estado)

@app.route('/api/conectar', methods=['POST'])
def api_conectar():
    """API: conectar con Google Sheets"""
    global SPREADSHEET_ID
    
    data = request.json
    url = data.get('url', '')
    key = data.get('key', '')
    
    if key:
        SPREADSHEET_ID = key
        estado['spreadsheet_id'] = key
        logger.info(f'✅ SPREADSHEET_ID configurado: {SPREADSHEET_ID}')
    
    # Autenticar con Google
    if sheets_service is None:
        if not autenticar_google():
            return jsonify({'success': False, 'estado': estado})
    
    # Iniciar monitoreo
    iniciar_monitoreo()
    
    estado['mensaje'] = f'🔗 Conectado al Sheet: {SPREADSHEET_ID}'
    socketio.emit('estado_actualizado', estado)
    
    return jsonify({'success': True, 'estado': estado})

@app.route('/api/detener')
def api_detener():
    """API: detener monitoreo"""
    detener_monitoreo()
    estado['conectado'] = False
    estado['monitoreando'] = False
    estado['mensaje'] = '⏹️ Conexión detenida'
    socketio.emit('estado_actualizado', estado)
    return jsonify({'success': True, 'estado': estado})

@app.route('/api/sincronizar')
def api_sincronizar():
    """API: sincronizar manualmente"""
    if not SPREADSHEET_ID:
        estado['mensaje'] = '❌ Primero conecta con un Google Sheet'
        socketio.emit('estado_actualizado', estado)
        return jsonify({'success': False, 'estado': estado})
    
    socketio.start_background_task(sincronizar)
    return jsonify({'success': True, 'estado': estado})

# ============================================================
# EVENTOS SOCKET.IO
# ============================================================

@socketio.on('connect')
def handle_connect():
    """Cliente conectado"""
    logger.info('🟢 Cliente conectado')
    emit('estado_actualizado', estado)

@socketio.on('conectar')
def handle_conectar(data):
    """Conectar con Google Sheets desde el cliente"""
    global SPREADSHEET_ID
    
    url = data.get('url', '')
    key = data.get('key', '')
    
    if key:
        SPREADSHEET_ID = key
        estado['spreadsheet_id'] = key
        logger.info(f'✅ SPREADSHEET_ID configurado: {SPREADSHEET_ID}')
    
    # Autenticar con Google
    if sheets_service is None:
        if not autenticar_google():
            emit('estado_actualizado', estado)
            return
    
    # Iniciar monitoreo
    iniciar_monitoreo()
    
    estado['conectado'] = True
    estado['mensaje'] = f'🔗 Conectado al Sheet: {SPREADSHEET_ID}'
    emit('estado_actualizado', estado)

@socketio.on('detener')
def handle_detener():
    """Detener monitoreo desde el cliente"""
    detener_monitoreo()
    estado['conectado'] = False
    estado['monitoreando'] = False
    estado['mensaje'] = '⏹️ Conexión detenida'
    emit('estado_actualizado', estado)

@socketio.on('sincronizar')
def handle_sincronizar():
    """Sincronizar desde el cliente"""
    if not SPREADSHEET_ID:
        estado['mensaje'] = '❌ Primero conecta con un Google Sheet'
        emit('estado_actualizado', estado)
        return
    
    socketio.start_background_task(sincronizar)

# ============================================================
# FUNCIÓN PARA INICIAR EL SERVIDOR EN UN HILO
# ============================================================

def iniciar_servidor():
    """Inicia el servidor Flask en un hilo separado"""
    socketio.run(app, host='127.0.0.1', port=5000, debug=False, allow_unsafe_werkzeug=True)

# ============================================================
# FUNCIÓN PRINCIPAL
# ============================================================

def main():
    """Función principal - Abre ventana independiente"""
    global EXCEL_PATH
    
    print("""
    ╔═══════════════════════════════════════════════════════════════════╗
    ║     SISTEMA DE SINCRONIZACIÓN EXCEL → GOOGLE SHEETS              ║
    ║                    Grupo GSN - Gestión de Distribución           ║
    ║                                                                  ║
    ║     📊 Excel local → ☁️ Google Sheets → 🌐 Dashboard            ║
    ╚═══════════════════════════════════════════════════════════════════╝
    """)
    
    # Verificar credenciales
    if not CREDENTIALS_FILE.exists():
        logger.warning('⚠️ No se encontró credentials.json')
        logger.info('📋 Descarga el archivo de Google Cloud Console')
        logger.info('   https://console.cloud.google.com/')
    
    # Buscar el archivo Excel en la carpeta tabla_excel
    excel_files = list(EXCEL_DIR.glob('*.xlsx')) + list(EXCEL_DIR.glob('*.xls'))
    
    if excel_files:
        # Tomar el primer archivo encontrado
        EXCEL_PATH = excel_files[0]
        estado['archivo_excel'] = str(EXCEL_PATH)
        estado['mensaje'] = f'📁 Archivo encontrado: {EXCEL_PATH.name}'
        logger.info(f'✅ Archivo cargado: {EXCEL_PATH.name}')
        
        # Mostrar información del archivo
        try:
            df = pd.read_excel(EXCEL_PATH, engine='openpyxl')
            logger.info(f'📊 Registros: {len(df)} filas, {len(df.columns)} columnas')
        except Exception as e:
            logger.warning(f'⚠️ No se pudo leer el archivo: {e}')
    else:
        logger.warning('⚠️ No se encontraron archivos Excel en la carpeta tabla_excel')
        estado['mensaje'] = '⚠️ No se encontró archivo Excel en tabla_excel'
        estado['archivo_excel'] = 'No encontrado'
    
    # Iniciar servidor en hilo separado
    logger.info('🚀 Iniciando servidor web...')
    server_thread = threading.Thread(target=iniciar_servidor, daemon=True)
    server_thread.start()
    
    time.sleep(2)
    
    # Abrir en navegador o PyWebView
    if WEBVIEW_AVAILABLE:
        logger.info('🪟 Abriendo ventana de la aplicación con PyWebView...')
        window = webview.create_window(
            title='GSN - Sincronizador Excel → Google Sheets',
            url='http://127.0.0.1:5000',
            width=900,
            height=650,
            min_size=(700, 550),
            resizable=True,
            fullscreen=False,
            confirm_close=True,
        )
        webview.start(window, debug=False)
    else:
        logger.info('🌐 Abriendo navegador en: http://127.0.0.1:5000')
        webbrowser.open('http://127.0.0.1:5000')
        
        logger.info('💡 Presiona Ctrl+C en la terminal para detener el servidor')
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            pass
    
    logger.info('🛑 Cerrando aplicación...')
    detener_monitoreo()
    executor.shutdown(wait=False)
    sys.exit(0)

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print('\n⏹️ Programa detenido por el usuario')
    except Exception as e:
        logger.error(f'❌ Error inesperado: {e}')
        raise