from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import subprocess
import os
import sys
from datetime import datetime

app = Flask(__name__)
CORS(app)

# ==========================================
# CONFIGURACIÓN DE RUTAS (MAPA DEL TESORO)
# ==========================================
# Ajustado para coincidir con tu estructura de carpetas 'WORD_TEAM4'

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

SCRIPTS_CONFIG = {
    # 1. PROYECTO DE TESIS
    'proyecto': {
        'folder': 'Formato_ProyectoDeTesis',
        'script': 'generador_proyecto_tesis_unac.py',
        'jsons': {
            'cuant': os.path.join('formats', 'unac_proyecto_cuant.json'),
            'cual':  os.path.join('formats', 'unac_proyecto_cual.json')
        }
    },

    # 2. INFORME DE TESIS (PREGRADO)
    'pregrado': {
        'folder': 'Formato_InformeDeTesis',
        'script': 'generador_informe_tesis_unac.py',
        'jsons': {
            'cuant': os.path.join('formats', 'unac_informe_cuant.json'),
            'cual':  os.path.join('formats', 'unac_informe_cual.json')
        }
    },

    # 3. MAESTRÍA
    'maestria': {
        'folder': 'FormatoMaestria',
        'script': 'generate_from_json.py', # Asumo este nombre basado en tu historial
        'jsons': {
            'cuant': os.path.join('formats', 'unac_maestria_cuant.json'),
            'cual':  os.path.join('formats', 'unac_maestria_cual.json')
        }
    }
}

@app.route('/')
def index():
    # Intenta servir desde view/ o desde la raiz
    if os.path.exists(os.path.join('view', 'index.html')):
        return send_file(os.path.join('view', 'index.html'))
    return send_file('index.html')

@app.route('/generate', methods=['POST'])
def generate_document():
    try:
        data = request.json
        fmt_type = data.get('format')      # proyecto, pregrado, maestria
        sub_type = data.get('sub_type')    # cuant, cual
        
        print(f"\n[SOLICITUD] Generando: {fmt_type.upper()} - {sub_type.upper()}")

        if fmt_type not in SCRIPTS_CONFIG:
            return jsonify({'error': 'Tipo de formato no válido'}), 400

        config = SCRIPTS_CONFIG[fmt_type]
        
        # Rutas absolutas para evitar confusiones
        work_dir = os.path.join(BASE_DIR, "..", config['folder']) # Subimos un nivel si server.py está en 'view', sino ajustar
        
        # CORRECCIÓN: Si server.py está en la raíz (WORD_TEAM4/view/server.py o WORD_TEAM4/server.py)
        # Asumiremos que server.py está en la RAÍZ (junto a las carpetas de formatos)
        # Si server.py está dentro de 'view', usa: os.path.join(BASE_DIR, "..", config['folder'])
        # Si server.py está en la raíz, usa: os.path.join(BASE_DIR, config['folder'])
        
        # Detectamos dónde estamos para ser flexibles:
        if os.path.exists(os.path.join(BASE_DIR, config['folder'])):
             work_dir = os.path.join(BASE_DIR, config['folder'])
        elif os.path.exists(os.path.join(BASE_DIR, "..", config['folder'])):
             work_dir = os.path.join(BASE_DIR, "..", config['folder'])
        else:
             return jsonify({'error': f"No encuentro la carpeta: {config['folder']}"}), 500

        work_dir = os.path.abspath(work_dir)
        script_path = os.path.join(work_dir, config['script'])
        
        # Validar Script
        if not os.path.exists(script_path):
            return jsonify({'error': f"Script no encontrado: {config['script']}"}), 500

        # Validar JSON
        json_rel = config['jsons'].get(sub_type)
        json_path = os.path.join(work_dir, json_rel)
        
        if not os.path.exists(json_path):
            return jsonify({'error': f"JSON no encontrado: {json_rel}"}), 500

        # Preparar Salida
        output_folder = os.path.join(BASE_DIR, 'descargas')
        if not os.path.exists(output_folder): os.makedirs(output_folder)
        
        filename = f"UNAC_{fmt_type.upper()}_{sub_type.upper()}_{datetime.now().strftime('%H%M%S')}.docx"
        output_path = os.path.join(output_folder, filename)

        # EJECUCIÓN DEL SUBPROCESO
        # Llamamos a python pasando: [script, ruta_json, ruta_salida]
        print(f"   -> Ejecutando en: {work_dir}")
        print(f"   -> Script: {config['script']}")
        
        cmd = [sys.executable, script_path, json_path, output_path]
        
        result = subprocess.run(
            cmd,
            cwd=work_dir, # Importante: el script corre "dentro" de su carpeta
            capture_output=True,
            text=True
        )

        if result.returncode != 0:
            print(f"[ERROR PYTHON]: {result.stderr}")
            return jsonify({'error': 'Falló la generación interna. Ver consola.'}), 500

        if os.path.exists(output_path):
            print(f"   -> Éxito. Enviando archivo.")
            return send_file(output_path, as_attachment=True, download_name=filename)
        else:
            return jsonify({'error': 'El script corrió pero no generó el archivo .docx'}), 500

    except Exception as e:
        print(f"[ERROR SERVER]: {e}")
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    print("🚀 Servidor UNAC iniciado en http://localhost:5000")
    app.run(debug=True, port=5000)