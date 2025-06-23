# main.py - Backend corregido con nuevas funcionalidades
from flask import Flask, jsonify, send_from_directory
from flask_cors import CORS
import requests
import openpyxl
from io import BytesIO
import schedule
import time
import threading
from datetime import datetime
import os
import logging

# Configurar logging para debug
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

# Variables globales
dashboard_data = {
    "last_update": None,
    "florida_data": {},
    "texas_data": {},
    "global_data": {},
    "status": "waiting"
}

workbook = None  # Variable global para el workbook

def download_and_process_excel():
    """Descarga el Excel de SharePoint y procesa los datos"""
    global dashboard_data, workbook
    
    try:
        logger.info("🔄 Iniciando descarga de SharePoint...")
        
        # URL del SharePoint
        url = "https://916foods-my.sharepoint.com/personal/it_support_916foods_com/_layouts/15/download.aspx?share=EZEBqKqQF9pFitMhSuZPwj4B4xV5tW0qtHLdceNN5-I9Ug"
        
        # Headers para evitar bloqueos
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        
        response = requests.get(url, headers=headers, timeout=30)
        logger.info(f"📥 Respuesta SharePoint: Status {response.status_code}, Tamaño: {len(response.content)} bytes")
        
        response.raise_for_status()
        
        # Verificar que el archivo no esté vacío
        if len(response.content) < 1000:
            raise Exception(f"Archivo muy pequeño o vacío: {len(response.content)} bytes")
        
        # Cargar el Excel en memoria
        logger.info("📊 Cargando archivo Excel...")
        workbook = openpyxl.load_workbook(BytesIO(response.content), data_only=True)
        
        logger.info(f"📋 Hojas encontradas en Excel: {workbook.sheetnames}")
        
        # Verificar que existen las hojas necesarias
        required_sheets = ['FLO', 'TEX']
        for sheet_name in required_sheets:
            if sheet_name not in workbook.sheetnames:
                raise Exception(f"Hoja '{sheet_name}' no encontrada. Disponibles: {workbook.sheetnames}")
        
        logger.info("✅ Archivo Excel cargado correctamente")
        
        # Procesar datos de Florida
        logger.info("🏖️ Procesando datos de Florida...")
        florida_data = process_sheet_data(workbook, 'FLO')
        
        # Procesar datos de Texas  
        logger.info("🤠 Procesando datos de Texas...")
        texas_data = process_sheet_data(workbook, 'TEX')
        
        # Combinar datos globales
        logger.info("🌍 Combinando datos globales...")
        global_data = combine_regional_data(florida_data, texas_data)
        
        # Actualizar datos globales
        dashboard_data = {
            "last_update": datetime.now().isoformat(),
            "florida_data": florida_data,
            "texas_data": texas_data,
            "global_data": global_data,
            "status": "success"
        }
        
        logger.info("✅ Datos procesados correctamente")
        logger.info(f"📊 Resumen - FL: {florida_data.get('aloha19', {}).get('total', 0)} tiendas, TX: {texas_data.get('aloha19', {}).get('total', 0)} tiendas")
        
    except Exception as e:
        error_msg = f"Error procesando datos: {str(e)}"
        logger.error(f"❌ {error_msg}")
        dashboard_data["status"] = f"error: {str(e)}"

def read_excel_cell(sheet, cell):
    """Lee una celda del Excel de forma segura con DEBUG"""
    try:
        cell_obj = sheet[cell]
        value = cell_obj.value
        
        # DEBUG: Mostrar valor exacto de cada celda
        logger.info(f"🔍 Celda {cell}: '{value}' (tipo: {type(value)})")
        
        if value is None:
            logger.warning(f"⚠️ Celda {cell} está vacía")
            return 0
        
        # Convertir a número
        if isinstance(value, (int, float)):
            num_value = float(value)
        else:
            # Si es texto, intentar convertir
            try:
                num_value = float(str(value).strip())
            except:
                logger.warning(f"⚠️ No se pudo convertir '{value}' a número en celda {cell}")
                return 0
        
        # Validar que no sea negativo
        if num_value < 0:
            logger.warning(f"⚠️ Valor negativo en celda {cell}: {num_value}")
            return 0
        
        logger.info(f"✅ Celda {cell} = {num_value}")
        return num_value
        
    except Exception as e:
        logger.error(f"❌ Error leyendo celda {cell}: {str(e)}")
        return 0

def process_sheet_data(workbook, sheet_name):
    """Procesa los datos de una hoja específica (FLO o TEX) con DEBUG completo"""
    try:
        sheet = workbook[sheet_name]
        logger.info(f"📋 Procesando hoja: {sheet_name}")
        
        # DEBUG: Mostrar información básica de la hoja
        logger.info(f"📏 Dimensiones de la hoja: {sheet.max_row} filas x {sheet.max_column} columnas")
        
        if sheet_name == 'FLO':
            logger.info("🏖️ === PROCESANDO FLORIDA (FLO) ===")
            
            # DEBUG: Leer y mostrar cada celda individual
            logger.info("📊 Leyendo datos de Aloha 19...")
            stage1 = read_excel_cell(sheet, 'B3')
            stage2 = read_excel_cell(sheet, 'B4')
            finished = read_excel_cell(sheet, 'B5')
            total = read_excel_cell(sheet, 'B6')
            
            logger.info("🔌 Leyendo datos de Wiring... (CORREGIDO)")
            # CORREGIDO: Florida wiring debe ser B11 (finished) y B12 (pending)
            wiring_finished = read_excel_cell(sheet, 'B11')  # Cambiado de B10 a B11
            wiring_pending = read_excel_cell(sheet, 'B12')   # Cambiado de B11 a B12
            
            logger.info("🤖 Leyendo datos de Tecnologías (Columna C - YES)...")
            fresh_ai = read_excel_cell(sheet, 'C15')
            edmb = read_excel_cell(sheet, 'C16')
            idmb = read_excel_cell(sheet, 'C17')
            qb = read_excel_cell(sheet, 'C18')
            kiosk = read_excel_cell(sheet, 'C19')
            
            logger.info("📋 Leyendo datos de Proyectos...")
            signed = read_excel_cell(sheet, 'B24')
            quote = read_excel_cell(sheet, 'B25')
            paid = read_excel_cell(sheet, 'B26')
            
            logger.info("🏗️ Leyendo tipos de proyectos...")
            project_edmb_idmb_qb = read_excel_cell(sheet, 'B30')
            project_fai_edmb_idmb_qb = read_excel_cell(sheet, 'B31')
            
            # Datos de Florida
            data = {
                "aloha19": {
                    "stage1": stage1,
                    "stage2": stage2, 
                    "finished": finished,
                    "total": total
                },
                "wiring": {
                    "pending": wiring_pending,    # CORREGIDO: B12
                    "finished": wiring_finished  # CORREGIDO: B11
                },
                "technologies": {
                    "fresh_ai": fresh_ai,
                    "edmb": edmb,
                    "idmb": idmb,
                    "qb": qb,
                    "kiosk": kiosk
                },
                "projects": {
                    "signed": signed,
                    "quote": quote,
                    "paid": paid
                },
                "project_types": {
                    "edmb_idmb_qb": project_edmb_idmb_qb,
                    "fai_edmb_idmb_qb": project_fai_edmb_idmb_qb
                }
            }
            
        else:  # TEX
            logger.info("🤠 === PROCESANDO TEXAS (TEX) ===")
            
            logger.info("📊 Leyendo datos de Aloha 19...")
            stage1 = read_excel_cell(sheet, 'B3')
            stage2 = read_excel_cell(sheet, 'B4')
            close = read_excel_cell(sheet, 'B5')
            finished = read_excel_cell(sheet, 'B6')
            total = read_excel_cell(sheet, 'B7')
            
            logger.info("🔌 Leyendo datos de Wiring...")
            wiring_pending = read_excel_cell(sheet, 'B12')
            wiring_finished = read_excel_cell(sheet, 'B13')
            
            logger.info("🤖 Leyendo datos de Tecnologías (Columna B)...")
            fresh_ai = read_excel_cell(sheet, 'B18')
            edmb = read_excel_cell(sheet, 'B19')
            idmb = read_excel_cell(sheet, 'B20')
            qb = read_excel_cell(sheet, 'B21')
            kiosk = read_excel_cell(sheet, 'B22')
            
            logger.info("📋 Leyendo datos de Proyectos...")
            # NOTA: El usuario especificó Quote en B7, pero B7 ya se usa para total de Aloha19
            # Mantenemos B27 para Quote y agregamos log para verificar
            quote = read_excel_cell(sheet, 'B27')  # Verificar si debe ser B7 según usuario
            pending = read_excel_cell(sheet, 'B28')  # Texas Pending
            
            logger.info(f"⚠️ VERIFICAR: Usuario dice Quote=5 en B7, pero B7={total} (Aloha19 total)")
            logger.info(f"📊 Actualmente leyendo Quote desde B27={quote}")
            
            logger.info("🏗️ Leyendo tipos de proyectos...")
            project_edmb = read_excel_cell(sheet, 'B33')
            
            # Datos de Texas
            data = {
                "aloha19": {
                    "stage1": stage1,
                    "stage2": stage2,
                    "close": close,
                    "finished": finished,
                    "total": total
                },
                "wiring": {
                    "pending": wiring_pending,
                    "finished": wiring_finished
                },
                "technologies": {
                    "fresh_ai": fresh_ai,
                    "edmb": edmb,
                    "idmb": idmb,
                    "qb": qb,
                    "kiosk": kiosk
                },
                "projects": {
                    "quote": quote,
                    "pending": pending
                },
                "project_types": {
                    "edmb": project_edmb
                }
            }
        
        logger.info(f"✅ Datos procesados para {sheet_name}:")
        logger.info(f"   📊 Aloha19 Total: {data['aloha19']['total']}")
        logger.info(f"   📊 Aloha19 Finished: {data['aloha19']['finished']}")
        logger.info(f"   🔌 Wiring Finished: {data['wiring']['finished']}")
        logger.info(f"   🔌 Wiring Pending: {data['wiring']['pending']}")
        logger.info(f"   🤖 Fresh AI: {data['technologies']['fresh_ai']}")
        
        return data
        
    except Exception as e:
        logger.error(f"❌ Error procesando hoja {sheet_name}: {str(e)}")
        return {}

def combine_regional_data(florida_data, texas_data):
    """Combina los datos de Florida y Texas para vista global con DEBUG"""
    try:
        logger.info("🌍 === COMBINANDO DATOS GLOBALES ===")
        
        # Helper para obtener valores seguros
        def safe_get(data, path, default=0):
            try:
                result = data
                for key in path.split('.'):
                    result = result[key]
                return result or default
            except:
                return default
        
        fl_total = safe_get(florida_data, 'aloha19.total')
        tx_total = safe_get(texas_data, 'aloha19.total')
        
        fl_finished = safe_get(florida_data, 'aloha19.finished')
        tx_finished = safe_get(texas_data, 'aloha19.finished')
        
        logger.info(f"🏖️ Florida - Total: {fl_total}, Finished: {fl_finished}")
        logger.info(f"🤠 Texas - Total: {tx_total}, Finished: {tx_finished}")
        
        # CORREGIDO: Combinar proyectos de ambas regiones
        fl_signed = safe_get(florida_data, 'projects.signed')
        fl_quote = safe_get(florida_data, 'projects.quote')
        fl_paid = safe_get(florida_data, 'projects.paid')
        
        tx_quote = safe_get(texas_data, 'projects.quote')
        tx_pending = safe_get(texas_data, 'projects.pending')
        
        logger.info(f"🏖️ Florida Projects - Signed: {fl_signed}, Quote: {fl_quote}, Paid: {fl_paid}")
        logger.info(f"🤠 Texas Projects - Quote: {tx_quote}, Pending: {tx_pending}")
        
        global_data = {
            "aloha19": {
                "stage1": safe_get(florida_data, 'aloha19.stage1') + safe_get(texas_data, 'aloha19.stage1'),
                "stage2": safe_get(florida_data, 'aloha19.stage2') + safe_get(texas_data, 'aloha19.stage2'),
                "close": safe_get(texas_data, 'aloha19.close'),  # Solo Texas tiene "close"
                "finished": fl_finished + tx_finished,
                "total": fl_total + tx_total
            },
            "wiring": {
                "pending": safe_get(florida_data, 'wiring.pending') + safe_get(texas_data, 'wiring.pending'),
                "finished": safe_get(florida_data, 'wiring.finished') + safe_get(texas_data, 'wiring.finished')
            },
            "technologies": {
                "fresh_ai": safe_get(florida_data, 'technologies.fresh_ai') + safe_get(texas_data, 'technologies.fresh_ai'),
                "edmb": safe_get(florida_data, 'technologies.edmb') + safe_get(texas_data, 'technologies.edmb'),
                "idmb": safe_get(florida_data, 'technologies.idmb') + safe_get(texas_data, 'technologies.idmb'),
                "qb": safe_get(florida_data, 'technologies.qb') + safe_get(texas_data, 'technologies.qb'),
                "kiosk": safe_get(florida_data, 'technologies.kiosk') + safe_get(texas_data, 'technologies.kiosk')
            },
            # CORREGIDO: Combinar proyectos globales
            "projects": {
                "signed": fl_signed,  # Solo Florida tiene signed
                "quote": fl_quote + tx_quote,  # Florida + Texas quotes
                "paid": fl_paid,  # Solo Florida tiene paid
                "pending": tx_pending  # Solo Texas tiene pending
            },
            # AGREGADO: Datos separados para gráficas individuales
            "project_types_florida": {
                "edmb_idmb_qb": safe_get(florida_data, 'project_types.edmb_idmb_qb'),
                "fai_edmb_idmb_qb": safe_get(florida_data, 'project_types.fai_edmb_idmb_qb')
            },
            "project_types_texas": {
                "edmb": safe_get(texas_data, 'project_types.edmb')
            }
        }
        
        logger.info(f"🌍 Global combinado - Total: {global_data['aloha19']['total']}, Finished: {global_data['aloha19']['finished']}")
        logger.info(f"🌍 Global Fresh AI: {global_data['technologies']['fresh_ai']}")
        
        return global_data
        
    except Exception as e:
        logger.error(f"❌ Error combinando datos: {str(e)}")
        return {}

# NUEVAS FUNCIONES PARA TABLAS DETALLADAS

def get_table_data(sheet_name, columns=None, filter_rows=True, max_row=None):
    """Obtiene datos de una hoja para tabla con filtros opcionales"""
    try:
        global workbook
        if not workbook or sheet_name not in workbook.sheetnames:
            return {"error": f"Hoja {sheet_name} no encontrada"}
        
        sheet = workbook[sheet_name]
        logger.info(f"📋 Leyendo tabla de hoja: {sheet_name}")
        
        # Definir rangos específicos por hoja
        if sheet_name == 'TEX-COM':
            actual_max_row = min(59, sheet.max_row) if max_row is None else max_row
        elif sheet_name == 'FLO-COM':
            actual_max_row = min(28, sheet.max_row) if max_row is None else max_row
        else:
            actual_max_row = sheet.max_row if max_row is None else max_row
        
        # Si no se especifican columnas, leer de A hasta T (20)
        if not columns:
            columns = [chr(65 + i) for i in range(20)]  # A-T
        
        table_data = []
        
        # Mapeo de columnas a nombres legibles (basado en la fila 1)
        column_names = {
            'A': 'STORE', 'B': 'ADDRESS', 'C': 'DM', 'D': 'A19', 'E': 'WIRING',
            'F': 'FRESH AI', 'G': 'EDMB', 'H': 'IDMB', 'I': 'QB', 'J': 'KIOSK',
            'K': 'A19 UP', 'L': 'START REMOD', 'M': 'END REMOD', 'N': 'PROJECT',
            'O': 'AUV', 'P': 'COST', 'Q': 'STATUS', 'R': 'INSTALLATION', 'S': 'Col_S', 'T': 'Col_T'
        }
        
        # Leer datos fila por fila (empezar desde fila 2 para evitar headers)
        for row_num in range(2, actual_max_row + 1):
            row_data = {}
            valid_row = False
            
            for col in columns:
                try:
                    cell_value = sheet[f"{col}{row_num}"].value
                    
                    # Convertir valores None a "---"
                    if cell_value is None:
                        cell_value = "---"
                    elif isinstance(cell_value, (int, float)) and cell_value == 0:
                        cell_value = "---"  # Cambiar 0 por "---"
                    else:
                        cell_value = str(cell_value).strip()
                        if cell_value in ["", "0", "0.0"]:
                            cell_value = "---"
                    
                    row_data[col] = cell_value
                    row_data[f"{col}_name"] = column_names.get(col, f"Col_{col}")
                    
                    # Marcar como fila válida si tiene contenido real
                    if cell_value not in ["---", "", " "]:
                        valid_row = True
                        
                except Exception as e:
                    row_data[col] = "---"
                    row_data[f"{col}_name"] = column_names.get(col, f"Col_{col}")
            
            # Agregar fila solo si es válida o si no estamos filtrando
            if valid_row or not filter_rows:
                table_data.append({"row": row_num, "data": row_data})
        
        logger.info(f"✅ Tabla {sheet_name} leída: {len(table_data)} filas (rango hasta fila {actual_max_row})")
        return {"data": table_data, "columns": columns, "column_names": column_names}
        
    except Exception as e:
        logger.error(f"❌ Error leyendo tabla {sheet_name}: {str(e)}")
        return {"error": str(e)}

# Rutas de la API
@app.route('/')
def home():
    return jsonify({
        "message": "916 Foods Dashboard API",
        "status": dashboard_data["status"],
        "last_update": dashboard_data["last_update"]
    })

@app.route('/api/data')
def get_dashboard_data():
    """Endpoint principal que devuelve todos los datos"""
    logger.info(f"📡 API request - Status: {dashboard_data['status']}")
    return jsonify(dashboard_data)

@app.route('/api/florida')
def get_florida_data():
    """Endpoint para datos solo de Florida"""
    return jsonify({
        "data": dashboard_data["florida_data"],
        "last_update": dashboard_data["last_update"],
        "status": dashboard_data["status"]
    })

@app.route('/api/texas') 
def get_texas_data():
    """Endpoint para datos solo de Texas"""
    return jsonify({
        "data": dashboard_data["texas_data"],
        "last_update": dashboard_data["last_update"],
        "status": dashboard_data["status"]
    })

@app.route('/api/refresh')
def manual_refresh():
    """Endpoint para forzar actualización manual"""
    logger.info("🔄 Refresh manual solicitado")
    threading.Thread(target=download_and_process_excel).start()
    return jsonify({"message": "Actualización iniciada"})

# NUEVOS ENDPOINTS PARA TABLAS DETALLADAS

@app.route('/api/table/<region>/detailed')
def get_detailed_regional_table(region):
    """Obtiene tabla detallada regional de hojas FLO-COM o TEX-COM"""
    try:
        if region.lower() == 'florida':
            sheet_name = 'FLO-COM'
        elif region.lower() == 'texas':
            sheet_name = 'TEX-COM'
        else:
            return jsonify({"error": "Región debe ser 'florida' o 'texas'"})
        
        # Leer toda la tabla de la hoja COM
        result = get_table_data(sheet_name, filter_rows=True)
        
        if "error" in result:
            return jsonify(result)
        
        return jsonify({
            "status": "success",
            "region": region,
            "sheet": sheet_name,
            "data": result["data"],
            "columns": result["columns"],
            "total_rows": len(result["data"])
        })
        
    except Exception as e:
        logger.error(f"❌ Error en tabla detallada {region}: {str(e)}")
        return jsonify({"error": str(e)})

@app.route('/api/table/projects')
def get_project_details_table():
    """Obtiene tabla de detalles de proyectos con columnas específicas y filtros"""
    try:
        # Columnas específicas requeridas: STORE, ADDRESS, PROJECT, AUV, COST, STATUS, INSTALLATION
        # Mapeo: A=STORE, B=ADDRESS, N=PROJECT, O=AUV, P=COST, Q=STATUS, R=INSTALLATION
        required_columns = ['A', 'B', 'N', 'O', 'P', 'Q', 'R']
        
        # Filtros válidos para la columna PROJECT
        valid_projects = ['FAI,EDMB,IDMB,QB', 'EDMB,IDMB,QB', 'EDMB']
        
        # Intentar con ambas hojas
        project_data = []
        
        for sheet_name in ['FLO-COM', 'TEX-COM']:
            result = get_table_data(sheet_name, required_columns, filter_rows=False)
            
            if "error" not in result:
                # Filtrar filas según PROJECT
                for row_info in result["data"]:
                    row_data = row_info["data"]
                    
                    # Verificar si PROJECT tiene alguno de los valores válidos
                    project_value = row_data.get('N', '---').strip()
                    
                    # Verificar si el proyecto coincide con alguno de los filtros válidos
                    project_matches = False
                    for valid_project in valid_projects:
                        if valid_project.upper() in project_value.upper():
                            project_matches = True
                            break
                    
                    # Solo incluir si tiene un proyecto válido y datos completos
                    if project_matches:
                        # Verificar que tenga al menos STORE o ADDRESS válidos
                        store = row_data.get('A', '---')
                        address = row_data.get('B', '---')
                        
                        if store not in ['---', '', ' '] or address not in ['---', '', ' ']:
                            # Agregar información de la hoja de origen
                            row_data['_source_sheet'] = sheet_name
                            row_data['_region'] = 'Florida' if sheet_name == 'FLO-COM' else 'Texas'
                            project_data.append(row_info)
        
        # Mapeo de nombres de columnas para display
        column_display_names = {
            'A': 'STORE',
            'B': 'ADDRESS', 
            'N': 'PROJECT',
            'O': 'AUV',
            'P': 'COST',
            'Q': 'STATUS',
            'R': 'INSTALLATION'
        }
        
        return jsonify({
            "status": "success",
            "data": project_data,
            "columns": required_columns,
            "column_names": column_display_names,
            "total_rows": len(project_data),
            "filters_applied": {
                "project_types": valid_projects,
                "note": "Solo se muestran filas con PROJECT que contenga: FAI,EDMB,IDMB,QB | EDMB,IDMB,QB | EDMB"
            }
        })
        
    except Exception as e:
        logger.error(f"❌ Error en tabla de proyectos: {str(e)}")
        return jsonify({"error": str(e)})

@app.route('/api/debug')
def debug_info():
    """Endpoint para información de debug"""
    return jsonify({
        "status": dashboard_data["status"],
        "last_update": dashboard_data["last_update"],
        "data_summary": {
            "florida_total": dashboard_data.get("florida_data", {}).get("aloha19", {}).get("total", 0),
            "texas_total": dashboard_data.get("texas_data", {}).get("aloha19", {}).get("total", 0),
            "global_total": dashboard_data.get("global_data", {}).get("aloha19", {}).get("total", 0),
            "florida_wiring_pending": dashboard_data.get("florida_data", {}).get("wiring", {}).get("pending", 0),
            "florida_wiring_finished": dashboard_data.get("florida_data", {}).get("wiring", {}).get("finished", 0)
        }
    })

@app.route('/api/sheets-available')
def list_available_sheets():
    """Lista todas las hojas disponibles en el Excel"""
    try:
        global workbook
        if not workbook:
            return jsonify({"error": "No workbook loaded"})
        
        return jsonify({
            "status": "success",
            "sheets": workbook.sheetnames,
            "required_for_dashboard": ["FLO", "TEX"],
            "required_for_tables": ["FLO-COM", "TEX-COM"],
            "com_sheets_available": {
                "FLO-COM": "FLO-COM" in workbook.sheetnames,
                "TEX-COM": "TEX-COM" in workbook.sheetnames
            }
        })
        
    except Exception as e:
        return jsonify({"error": str(e)})

def run_scheduler():
    """Ejecuta el scheduler en un hilo separado"""
    while True:
        schedule.run_pending()
        time.sleep(60)

if __name__ == '__main__':
    logger.info("🚀 Iniciando 916 Foods Dashboard API...")
    
    # Configurar actualizaciones automáticas cada 30 minutos
    schedule.every(30).minutes.do(download_and_process_excel)
    
    # Ejecutar una vez al inicio
    logger.info("🔄 Ejecutando carga inicial de datos...")
    download_and_process_excel()
    
    # Iniciar scheduler en hilo separado
    scheduler_thread = threading.Thread(target=run_scheduler)
    scheduler_thread.daemon = True
    scheduler_thread.start()
    
    # Iniciar servidor Flask
    port = int(os.environ.get('PORT', 5000))
    logger.info(f"🌐 Servidor iniciando en puerto {port}")
    app.run(host='0.0.0.0', port=port, debug=False)