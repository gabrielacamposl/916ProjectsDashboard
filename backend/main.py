# main.py - Backend completo con soporte para TEX-COM y FLO-COM
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
    "florida_com_data": {},  # Nueva sección para FLO-COM
    "texas_com_data": {},    # Nueva sección para TEX-COM
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
        required_sheets = ['FLO', 'TEX', 'FLO-COM', 'TEX-COM']
        missing_sheets = []
        for sheet_name in required_sheets:
            if sheet_name not in workbook.sheetnames:
                missing_sheets.append(sheet_name)
        
        if missing_sheets:
            logger.warning(f"⚠️ Hojas faltantes: {missing_sheets}. Disponibles: {workbook.sheetnames}")
        
        logger.info("✅ Archivo Excel cargado correctamente")
        
        # Procesar datos existentes
        florida_data = {}
        texas_data = {}
        florida_com_data = {}
        texas_com_data = {}
        
        # Procesar datos de Florida
        if 'FLO' in workbook.sheetnames:
            logger.info("🏖️ Procesando datos de Florida...")
            florida_data = process_sheet_data(workbook, 'FLO')
        
        # Procesar datos de Texas  
        if 'TEX' in workbook.sheetnames:
            logger.info("🤠 Procesando datos de Texas...")
            texas_data = process_sheet_data(workbook, 'TEX')
        
        # Procesar datos de FLO-COM
        if 'FLO-COM' in workbook.sheetnames:
            logger.info("🏖️📊 Procesando datos de Florida COM...")
            florida_com_data = process_com_sheet_data(workbook, 'FLO-COM', 'A1:T28')
        
        # Procesar datos de TEX-COM
        if 'TEX-COM' in workbook.sheetnames:
            logger.info("🤠📊 Procesando datos de Texas COM...") 
            texas_com_data = process_com_sheet_data(workbook, 'TEX-COM', 'A1:T59')
        
        # Combinar datos globales
        logger.info("🌍 Combinando datos globales...")
        global_data = combine_regional_data(florida_data, texas_data)
        
        # Actualizar datos globales
        dashboard_data = {
            "last_update": datetime.now().isoformat(),
            "florida_data": florida_data,
            "texas_data": texas_data,
            "florida_com_data": florida_com_data,
            "texas_com_data": texas_com_data,
            "global_data": global_data,
            "status": "success"
        }
        
        logger.info("✅ Datos procesados correctamente")
        logger.info(f"📊 Resumen - FL: {florida_data.get('aloha19', {}).get('total', 0)} tiendas, TX: {texas_data.get('aloha19', {}).get('total', 0)} tiendas")
        logger.info(f"📊 COM - FL: {len(florida_com_data.get('data', []))} filas, TX: {len(texas_com_data.get('data', []))} filas")
        
    except Exception as e:
        error_msg = f"Error procesando datos: {str(e)}"
        logger.error(f"❌ {error_msg}")
        dashboard_data["status"] = f"error: {str(e)}"

def process_com_sheet_data(workbook, sheet_name, cell_range):
    """Procesa los datos de las hojas COM con rangos específicos"""
    try:
        sheet = workbook[sheet_name]
        logger.info(f"📋 Procesando hoja COM: {sheet_name} - Rango: {cell_range}")
        
        # Parsear el rango (ej: "A1:T28")
        start_cell, end_cell = cell_range.split(':')
        
        # Convertir coordenadas de letra a número
        def col_letter_to_num(letter):
            num = 0
            for char in letter:
                num = num * 26 + (ord(char) - ord('A') + 1)
            return num
        
        def parse_cell(cell):
            col_letters = ''.join([c for c in cell if c.isalpha()])
            row_num = int(''.join([c for c in cell if c.isdigit()]))
            col_num = col_letter_to_num(col_letters)
            return row_num, col_num
        
        start_row, start_col = parse_cell(start_cell)
        end_row, end_col = parse_cell(end_cell)
        
        logger.info(f"📏 Procesando desde fila {start_row} col {start_col} hasta fila {end_row} col {end_col}")
        
        # Extraer datos del rango especificado
        data = []
        headers = []
        
        # Primera fila como headers
        first_row = True
        for row in sheet.iter_rows(min_row=start_row, max_row=end_row, 
                                  min_col=start_col, max_col=end_col, 
                                  values_only=True):
            if first_row:
                headers = [str(cell) if cell is not None else f"Col_{i}" for i, cell in enumerate(row)]
                first_row = False
                logger.info(f"📝 Headers encontrados: {headers[:5]}...") # Mostrar solo primeros 5
            else:
                # Convertir fila a diccionario
                row_dict = {}
                for i, cell_value in enumerate(row):
                    if i < len(headers):
                        row_dict[headers[i]] = cell_value
                
                # Solo agregar filas que no estén completamente vacías
                if any(value is not None and str(value).strip() != '' for value in row):
                    data.append(row_dict)
        
        result = {
            "sheet_name": sheet_name,
            "range_processed": cell_range,
            "total_rows": len(data),
            "headers": headers,
            "data": data,
            "summary": {
                "non_empty_rows": len([row for row in data if any(v for v in row.values() if v is not None)]),
                "total_columns": len(headers)
            }
        }
        
        logger.info(f"✅ Procesados {len(data)} filas de {sheet_name}")
        logger.info(f"   📊 Columnas: {len(headers)}")
        logger.info(f"   📝 Filas no vacías: {result['summary']['non_empty_rows']}")
        
        return result
        
    except Exception as e:
        logger.error(f"❌ Error procesando hoja COM {sheet_name}: {str(e)}")
        return {
            "sheet_name": sheet_name,
            "error": str(e),
            "data": [],
            "headers": []
        }

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
            
            logger.info("🔌 Leyendo datos de Wiring...")
            wiring_pending = read_excel_cell(sheet, 'B10')
            wiring_finished = read_excel_cell(sheet, 'B11')
            
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
            quote = read_excel_cell(sheet, 'B27')
            pending = read_excel_cell(sheet, 'B28')
            
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
            }
        }
        
        logger.info(f"🌍 Global combinado - Total: {global_data['aloha19']['total']}, Finished: {global_data['aloha19']['finished']}")
        logger.info(f"🌍 Global Fresh AI: {global_data['technologies']['fresh_ai']}")
        
        return global_data
        
    except Exception as e:
        logger.error(f"❌ Error combinando datos: {str(e)}")
        return {}

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

# NUEVOS ENDPOINTS PARA DATOS COM
@app.route('/api/florida-com')
def get_florida_com_data():
    """Endpoint para datos COM de Florida"""
    return jsonify({
        "data": dashboard_data["florida_com_data"],
        "last_update": dashboard_data["last_update"],
        "status": dashboard_data["status"]
    })

@app.route('/api/texas-com') 
def get_texas_com_data():
    """Endpoint para datos COM de Texas"""
    return jsonify({
        "data": dashboard_data["texas_com_data"],
        "last_update": dashboard_data["last_update"],
        "status": dashboard_data["status"]
    })

@app.route('/api/com-summary')
def get_com_summary():
    """Endpoint para resumen de datos COM"""
    florida_com = dashboard_data.get("florida_com_data", {})
    texas_com = dashboard_data.get("texas_com_data", {})
    
    return jsonify({
        "florida_com": {
            "total_rows": len(florida_com.get("data", [])),
            "columns": len(florida_com.get("headers", [])),
            "range": florida_com.get("range_processed", "N/A"),
            "non_empty_rows": florida_com.get("summary", {}).get("non_empty_rows", 0)
        },
        "texas_com": {
            "total_rows": len(texas_com.get("data", [])),
            "columns": len(texas_com.get("headers", [])),
            "range": texas_com.get("range_processed", "N/A"),
            "non_empty_rows": texas_com.get("summary", {}).get("non_empty_rows", 0)
        },
        "last_update": dashboard_data["last_update"],
        "status": dashboard_data["status"]
    })

@app.route('/api/refresh')
def manual_refresh():
    """Endpoint para forzar actualización manual"""
    logger.info("🔄 Refresh manual solicitado")
    threading.Thread(target=download_and_process_excel).start()
    return jsonify({"message": "Actualización iniciada"})

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
            "florida_com_rows": len(dashboard_data.get("florida_com_data", {}).get("data", [])),
            "texas_com_rows": len(dashboard_data.get("texas_com_data", {}).get("data", []))
        }
    })

@app.route('/api/inspect')
def inspect_cells():
    """Ver exactamente qué hay en cada celda"""
    try:
        global workbook
        if not workbook:
            return jsonify({"error": "No workbook loaded"})
        
        inspection = {
            "sheets": workbook.sheetnames,
            "florida": {},
            "texas": {}
        }
        
        # Florida cells
        if 'FLO' in workbook.sheetnames:
            flo = workbook['FLO']
            fl_cells = ['B3','B4','B5','B6','B10','B11','C15','C16','C17','C18','C19','B24','B25','B26','B30','B31']
            for cell in fl_cells:
                try:
                    inspection["florida"][cell] = {"value": flo[cell].value, "type": str(type(flo[cell].value))}
                except: 
                    inspection["florida"][cell] = {"error": "no existe"}
        
        # Texas cells  
        if 'TEX' in workbook.sheetnames:
            tex = workbook['TEX']
            tx_cells = ['B3','B4','B5','B6','B7','B12','B13','B18','B19','B20','B21','B22','B27','B28','B33']
            for cell in tx_cells:
                try:
                    inspection["texas"][cell] = {"value": tex[cell].value, "type": str(type(tex[cell].value))}
                except:
                    inspection["texas"][cell] = {"error": "no existe"}
        
        return jsonify(inspection)
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route('/api/preview/<sheet>')
def preview_sheet(sheet):
    """Ver primeras filas de una hoja"""
    try:
        global workbook
        if not workbook or sheet not in workbook.sheetnames:
            return jsonify({"error": f"Sheet {sheet} not found"})
        
        ws = workbook[sheet]
        preview = {}
        
        # Primeras 15 filas, columnas A-T para hojas COM, A-F para otras
        max_col = 'T' if 'COM' in sheet else 'F'
        cols = [chr(ord('A') + i) for i in range(ord(max_col) - ord('A') + 1)]
        
        for row in range(1, 16):
            for col in cols:
                cell_addr = f"{col}{row}"
                try:
                    val = ws[cell_addr].value
                    if val is not None:
                        preview[cell_addr] = val
                except:
                    pass
        
        return jsonify(preview)
    except Exception as e:
        return jsonify({"error": str(e)})

# NUEVAS FUNCIONES DE INSPECCIÓN DIRECTA
@app.route('/api/sheets')
def list_all_sheets():
    """Ver todas las hojas del Excel - descarga fresh"""
    try:
        logger.info("🔍 Descargando Excel para inspección...")
        
        # URL del SharePoint
        url = "https://916foods-my.sharepoint.com/personal/it_support_916foods_com/_layouts/15/download.aspx?share=EZEBqKqQF9pFitMhSuZPwj4B4xV5tW0qtHLdceNN5-I9Ug"
        
        # Headers para evitar bloqueos
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        
        response = requests.get(url, headers=headers, timeout=30)
        logger.info(f"📥 Respuesta: {response.status_code}, Tamaño: {len(response.content)} bytes")
        
        response.raise_for_status()
        
        # Verificar tamaño del archivo
        if len(response.content) < 1000:
            return jsonify({
                "status": "error",
                "error": f"Archivo muy pequeño: {len(response.content)} bytes"
            })
        
        # Cargar Excel
        workbook_temp = openpyxl.load_workbook(BytesIO(response.content), data_only=True)
        
        logger.info(f"📋 Hojas encontradas: {workbook_temp.sheetnames}")
        
        return jsonify({
            "status": "success",
            "file_size_bytes": len(response.content),
            "total_sheets": len(workbook_temp.sheetnames),
            "sheet_names": workbook_temp.sheetnames,
            "looking_for": ["FLO", "TEX", "FLO-COM", "TEX-COM"],
            "flo_exists": "FLO" in workbook_temp.sheetnames,
            "tex_exists": "TEX" in workbook_temp.sheetnames,
            "flo_com_exists": "FLO-COM" in workbook_temp.sheetnames,
            "tex_com_exists": "TEX-COM" in workbook_temp.sheetnames
        })
        
    except Exception as e:
        logger.error(f"❌ Error: {str(e)}")
        return jsonify({
            "status": "error", 
            "error": str(e)
        })

@app.route('/api/inspect-cells')
def inspect_specific_cells():
    """Inspeccionar celdas específicas"""
    try:
        logger.info("🔍 Inspeccionando celdas específicas...")
        
        # Descargar Excel fresh
        url = "https://916foods-my.sharepoint.com/personal/it_support_916foods_com/_layouts/15/download.aspx?share=EZEBqKqQF9pFitMhSuZPwj4B4xV5tW0qtHLdceNN5-I9Ug"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        
        workbook_temp = openpyxl.load_workbook(BytesIO(response.content), data_only=True)
        
        result = {
            "status": "success",
            "available_sheets": workbook_temp.sheetnames,
            "inspection": {}
        }
        
        # Inspeccionar cada hoja disponible
        for sheet_name in workbook_temp.sheetnames:
            try:
                sheet = workbook_temp[sheet_name]
                sheet_data = {
                    "max_row": sheet.max_row,
                    "max_column": sheet.max_column,
                    "sample_cells": {}
                }
                
                # Para hojas COM, mostrar más columnas
                if 'COM' in sheet_name:
                    # Leer muestra de celdas de las primeras 10 filas, columnas A-T
                    cols = [chr(ord('A') + i) for i in range(20)]  # A-T
                    max_rows = min(11, sheet.max_row + 1)
                else:
                    # Para otras hojas, columnas A-F
                    cols = ['A', 'B', 'C', 'D', 'E', 'F']
                    max_rows = min(11, sheet.max_row + 1)
                
                for row in range(1, max_rows):
                    for col_letter in cols:
                        cell_addr = f"{col_letter}{row}"
                        try:
                            cell_value = sheet[cell_addr].value
                            if cell_value is not None:
                                sheet_data["sample_cells"][cell_addr] = {
                                    "value": cell_value,
                                    "type": str(type(cell_value))
                                }
                        except:
                            pass
                
                result["inspection"][sheet_name] = sheet_data
                
            except Exception as e:
                result["inspection"][sheet_name] = {"error": str(e)}
        
        return jsonify(result)
        
    except Exception as e:
        logger.error(f"❌ Error en inspección: {str(e)}")
        return jsonify({
            "status": "error",
            "error": str(e)
        })

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