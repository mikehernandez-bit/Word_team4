from flask import Flask, request, send_file, jsonify
try:
    from flask_cors import CORS
except ImportError:
    CORS = None
import os
import subprocess
import sys
import platform

app = Flask(__name__)
if CORS:
    CORS(app)
else:
    print("[WARN] flask_cors no instalado; CORS desactivado.")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOCS_DIR = os.path.join(BASE_DIR, "docs")

SCRIPTS_CONFIG = {
    "proyecto": {
        "script": "generador_proyecto_tesis.py",
        "jsons": {
            "cuant": os.path.join("formats", "proyecto", "unac_proyecto_cuant.json"),
            "cual": os.path.join("formats", "proyecto", "unac_proyecto_cual.json"),
        },
    },
    "informe": {
        "script": "generador_informe_tesis.py",
        "jsons": {
            "cuant": os.path.join("formats", "informe", "unac_informe_cuant.json"),
            "cual": os.path.join("formats", "informe", "unac_informe_cual.json"),
        },
    },
    "maestria": {
        "script": "generador_maestria.py",
        "jsons": {
            "cuant": os.path.join("formats", "maestria", "unac_maestria_cuant.json"),
            "cual": os.path.join("formats", "maestria", "unac_maestria_cual.json"),
        },
    },
}

ALIASES = {
    "pregrado": "informe",
}


def open_document(path: str) -> None:
    try:
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin":
            subprocess.run(["open", path], check=False)
        else:
            subprocess.run(["xdg-open", path], check=False)
    except Exception as exc:
        print(f"[WARN] No se pudo abrir el documento: {exc}")


@app.route("/")
def index():
    view_path = os.path.join(BASE_DIR, "view", "index.html")
    if os.path.exists(view_path):
        return send_file(view_path)
    return send_file(os.path.join(BASE_DIR, "index.html"))


@app.route("/generate", methods=["POST"])
def generate_document():
    try:
        data = request.json or {}
        fmt_type = (data.get("format") or "").strip().lower()
        sub_type = (data.get("sub_type") or "").strip().lower()

        if fmt_type in ALIASES:
            fmt_type = ALIASES[fmt_type]

        if fmt_type not in SCRIPTS_CONFIG:
            return jsonify({"error": "Formato no valido"}), 400

        config = SCRIPTS_CONFIG[fmt_type]
        if sub_type not in config["jsons"]:
            return jsonify({"error": "Subtipo no valido"}), 400

        script_path = os.path.join(BASE_DIR, config["script"])
        if not os.path.exists(script_path):
            return jsonify({"error": f"Script no encontrado: {config['script']}"}), 500

        json_rel = config["jsons"][sub_type]
        json_path = os.path.join(BASE_DIR, json_rel)
        if not os.path.exists(json_path):
            return jsonify({"error": f"JSON no encontrado: {json_rel}"}), 500

        os.makedirs(DOCS_DIR, exist_ok=True)
        filename = f"UNAC_{fmt_type.upper()}_{sub_type.upper()}.docx"
        output_path = os.path.join(DOCS_DIR, filename)

        cmd = [sys.executable, script_path, json_path, output_path]
        result = subprocess.run(
            cmd,
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
        )

        if result.returncode != 0:
            print("[ERROR PYTHON]", result.stderr)
            return jsonify({"error": "Fallo la generacion interna. Revisa consola."}), 500

        if not os.path.exists(output_path):
            return jsonify({"error": "El script corrio pero no genero el DOCX"}), 500

        open_document(output_path)
        return jsonify({"ok": True, "filename": filename, "path": output_path})

    except Exception as exc:
        print("[ERROR SERVER]", exc)
        return jsonify({"error": str(exc)}), 500


if __name__ == "__main__":
    print("Servidor CentroFormatosUNAC listo en http://localhost:5000")
    app.run(debug=True, port=5000)
