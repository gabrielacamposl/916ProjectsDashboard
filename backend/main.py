# main.py - Backend para Railway
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

app = Flask(__name__)
CORS(app)  # Permitir requests desde Vercel

# Variable global para almacenar los datos procesados
dashboard_data = {
    "last_update": None,
    "florida_data": {},
    "texas_data": {},
    "global_data": {},
    "status": "waiting"
}

def download_and_process_excel():
    """Descarga el Excel de SharePoint y procesa los datos"""
    global dashboard_data
    
    try:
        print(f"Descargando archivo... {datetime.now()}")
        
        # URL del SharePoint
        url = "https://916foods-my.sharepoint.com/personal/it_support_916foods_com/_layouts/15/download.aspx?share=EZEBqKqQF9pFitMhSuZPwj4B4xV5tW0qtHLdceNN5-I9Ug"
        
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        
        # Cargar el Excel en memoria
        workbook = openpyxl.load_workbook(BytesIO(response.content))
        
        print("Archivo descargado exitosamente")
        
        # Procesar datos de Florida
        florida_data = process_sheet_data(workbook, 'FLO')
        
        # Procesar datos de Texas  
        texas_data = process_sheet_data(workbook, 'TEX')
        
        # Combinar datos globales
        global_data = combine_regional_data(florida_data, texas_data)
        
        # Actualizar datos globales
        dashboard_data = {
            "last_update": datetime.now().isoformat(),
            "florida_data": florida_data,
            "texas_data": texas_data,
            "global_data": global_data,
            "status": "success"
        }
        
        print("Datos procesados correctamente")
        
    except Exception as e:
        print(f"Error procesando datos: {str(e)}")
        dashboard_data["status"] = f"error: {str(e)}"

def read_excel_cell(sheet, cell):
    """Lee una celda del Excel de forma segura"""
    try:
        value = sheet[cell].value
        if value is None:
            return 0
        num_value = float(value)
        return max(0, num_value)  # No permitir negativos
    except:
        return 0

def process_sheet_data(workbook, sheet_name):
    """Procesa los datos de una hoja específica (FLO o TEX)"""
    try:
        sheet = workbook[sheet_name]
        
        if sheet_name == 'FLO':
            # Datos de Florida
            data = {
                # Aloha 19
                "aloha19": {
                    "stage1": read_excel_cell(sheet, 'B3'),
                    "stage2": read_excel_cell(sheet, 'B4'), 
                    "finished": read_excel_cell(sheet, 'B5'),
                    "total": read_excel_cell(sheet, 'B6')
                },
                # Wiring
                "wiring": {
                    "pending": read_excel_cell(sheet, 'B10'),
                    "finished": read_excel_cell(sheet, 'B11')
                },
                # Tecnologías (Columna C - YES)
                "technologies": {
                    "fresh_ai": read_excel_cell(sheet, 'C15'),
                    "edmb": read_excel_cell(sheet, 'C16'),
                    "idmb": read_excel_cell(sheet, 'C17'),
                    "qb": read_excel_cell(sheet, 'C18'),
                    "kiosk": read_excel_cell(sheet, 'C19')
                },
                # Proyectos
                "projects": {
                    "signed": read_excel_cell(sheet, 'B24'),
                    "quote": read_excel_cell(sheet, 'B25'),
                    "paid": read_excel_cell(sheet, 'B26')
                },
                # Tipos de proyectos
                "project_types": {
                    "edmb_idmb_qb": read_excel_cell(sheet, 'B30'),
                    "fai_edmb_idmb_qb": read_excel_cell(sheet, 'B31')
                }
            }
            
        else:  # TEX
            # Datos de Texas
            data = {
                # Aloha 19
                "aloha19": {
                    "stage1": read_excel_cell(sheet, 'B3'),
                    "stage2": read_excel_cell(sheet, 'B4'),
                    "close": read_excel_cell(sheet, 'B5'),
                    "finished": read_excel_cell(sheet, 'B6'),
                    "total": read_excel_cell(sheet, 'B7')
                },
                # Wiring
                "wiring": {
                    "pending": read_excel_cell(sheet, 'B12'),
                    "finished": read_excel_cell(sheet, 'B13')
                },
                # Tecnologías (Columna B)
                "technologies": {
                    "fresh_ai": read_excel_cell(sheet, 'B18'),
                    "edmb": read_excel_cell(sheet, 'B19'),
                    "idmb": read_excel_cell(sheet, 'B20'),
                    "qb": read_excel_cell(sheet, 'B21'),
                    "kiosk": read_excel_cell(sheet, 'B22')
                },
                # Proyectos
                "projects": {
                    "quote": read_excel_cell(sheet, 'B27'),
                    "pending": read_excel_cell(sheet, 'B28')
                },
                # Tipos de proyectos
                "project_types": {
                    "edmb": read_excel_cell(sheet, 'B33')
                }
            }
        
        return data
        
    except Exception as e:
        print(f"Error procesando hoja {sheet_name}: {str(e)}")
        return {}

def combine_regional_data(florida_data, texas_data):
    """Combina los datos de Florida y Texas para vista global"""
    try:
        global_data = {
            "aloha19": {
                "stage1": florida_data.get("aloha19", {}).get("stage1", 0) + texas_data.get("aloha19", {}).get("stage1", 0),
                "stage2": florida_data.get("aloha19", {}).get("stage2", 0) + texas_data.get("aloha19", {}).get("stage2", 0),
                "close": texas_data.get("aloha19", {}).get("close", 0),  # Solo Texas tiene "close"
                "finished": florida_data.get("aloha19", {}).get("finished", 0) + texas_data.get("aloha19", {}).get("finished", 0),
                "total": florida_data.get("aloha19", {}).get("total", 0) + texas_data.get("aloha19", {}).get("total", 0)
            },
            "wiring": {
                "pending": florida_data.get("wiring", {}).get("pending", 0) + texas_data.get("wiring", {}).get("pending", 0),
                "finished": florida_data.get("wiring", {}).get("finished", 0) + texas_data.get("wiring", {}).get("finished", 0)
            },
            "technologies": {
                "fresh_ai": florida_data.get("technologies", {}).get("fresh_ai", 0) + texas_data.get("technologies", {}).get("fresh_ai", 0),
                "edmb": florida_data.get("technologies", {}).get("edmb", 0) + texas_data.get("technologies", {}).get("edmb", 0),
                "idmb": florida_data.get("technologies", {}).get("idmb", 0) + texas_data.get("technologies", {}).get("idmb", 0),
                "qb": florida_data.get("technologies", {}).get("qb", 0) + texas_data.get("technologies", {}).get("qb", 0),
                "kiosk": florida_data.get("technologies", {}).get("kiosk", 0) + texas_data.get("technologies", {}).get("kiosk", 0)
            }
        }
        return global_data
    except Exception as e:
        print(f"Error combinando datos: {str(e)}")
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
    threading.Thread(target=download_and_process_excel).start()
    return jsonify({"message": "Actualización iniciada"})

def run_scheduler():
    """Ejecuta el scheduler en un hilo separado"""
    while True:
        schedule.run_pending()
        time.sleep(60)  # Revisar cada minuto

if __name__ == '__main__':
    # Configurar actualizaciones automáticas cada 30 minutos
    schedule.every(30).minutes.do(download_and_process_excel)
    
    # Ejecutar una vez al inicio
    download_and_process_excel()
    
    # Iniciar scheduler en hilo separado
    scheduler_thread = threading.Thread(target=run_scheduler)
    scheduler_thread.daemon = True
    scheduler_thread.start()
    
    # Iniciar servidor Flask
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)