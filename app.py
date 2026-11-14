#!/usr/bin/env python3
"""
Backend Flask per il Gestore Script Automazione (con supporto ad Upload Dati)
"""

import os
import sys
import subprocess
import ast
import shlex
from flask import Flask, render_template, jsonify, Response, request, send_from_directory
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Configurazione cartelle
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
SCRIPTS_DIR = os.path.join(BASE_DIR, "scripts")
UPLOADS_DIR = os.path.join(BASE_DIR, "uploads")

# Assicura che le cartelle esistano
os.makedirs(SCRIPTS_DIR, exist_ok=True)
os.makedirs(UPLOADS_DIR, exist_ok=True)

# Tipi di file di dati consentiti
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv', 'txt', 'json', 'md'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_script_metadata(filename):
    """
    Legge il file python ed estrae i metadati (Titolo, Descrizione, Categoria, Istruzioni)
    dalla docstring iniziale se presente.
    """
    file_path = os.path.join(SCRIPTS_DIR, filename)
    title = filename.replace(".py", "").replace("_", " ").title()
    description = "Nessuna descrizione fornita."
    instructions = "Nessuna istruzione fornita."
    category = "Utility"
    
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
        
        # Effettua il parsing dell'AST del file per ricavare la docstring in modo sicuro
        tree = ast.parse(content)
        docstring = ast.get_docstring(tree)
        
        if docstring:
            lines = docstring.strip().split("\n")
            desc_parts = []
            inst_parts = []
            in_instructions = False
            
            for line in lines:
                line_str = line.strip()
                
                # Rileva inizio blocco istruzioni
                if line_str.lower().startswith("istruzioni:") or line_str.lower().startswith("instructions:"):
                    in_instructions = True
                    parts = line_str.split(":", 1)
                    if len(parts) > 1 and parts[1].strip():
                        inst_parts.append(parts[1].strip())
                    continue
                
                # Se all'interno del blocco istruzioni, accumula le righe
                if in_instructions:
                    # Se si incontra un altro tag noto, esce dal blocco istruzioni
                    if (line_str.startswith("Nome:") or line_str.lower().startswith("title:") or
                        line_str.startswith("Categoria:") or line_str.lower().startswith("category:") or
                        line_str.startswith("Descrizione:") or line_str.lower().startswith("description:")):
                        in_instructions = False
                    else:
                        inst_parts.append(line_str)
                        continue
                
                # Ignora righe vuote per la descrizione
                if not line_str:
                    continue
                
                # Cerca tag specifici come Nome, Descrizione, Categoria
                if line_str.startswith("Nome:") or line_str.lower().startswith("title:"):
                    title = line_str.split(":", 1)[1].strip()
                elif line_str.startswith("Categoria:") or line_str.lower().startswith("category:"):
                    category = line_str.split(":", 1)[1].strip()
                elif line_str.startswith("Descrizione:") or line_str.lower().startswith("description:"):
                    desc_parts.append(line_str.split(":", 1)[1].strip())
                else:
                    desc_parts.append(line_str)
            
            if desc_parts:
                description = " ".join(desc_parts)
            if inst_parts:
                instructions = "\n".join(inst_parts).strip()
    except Exception:
        # Se c'è un errore di parsing (es. file non valido), lasciamo i valori di default
        pass
        
    return {
        "filename": filename,
        "title": title,
        "description": description,
        "instructions": instructions,
        "category": category,
        "size": os.path.getsize(file_path)
    }

@app.route('/')
def index():
    return render_template("index.html")

@app.route('/api/scripts')
def list_scripts():
    """Ritorna la lista di tutti gli script Python nella cartella scripts/ con metadati."""
    scripts = []
    if os.path.exists(SCRIPTS_DIR):
        for f in os.listdir(SCRIPTS_DIR):
            if f.endswith('.py'):
                scripts.append(get_script_metadata(f))
    # Ordina per categoria e poi per titolo
    scripts.sort(key=lambda x: (x['category'], x['title']))
    return jsonify(scripts)

@app.route('/api/uploads', methods=['GET'])
def list_uploads():
    """Ritorna la lista dei file caricati nella cartella uploads/."""
    files = []
    if os.path.exists(UPLOADS_DIR):
        for f in os.listdir(UPLOADS_DIR):
            if not f.startswith('.'):
                path = os.path.join(UPLOADS_DIR, f)
                files.append({
                    "name": f,
                    "size": os.path.getsize(path),
                    "modified": os.path.getmtime(path)
                })
    # Ordina per data di modifica decrescente (più recenti prima)
    files.sort(key=lambda x: x['modified'], reverse=True)
    return jsonify(files)

@app.route('/api/upload', methods=['POST'])
def upload_file():
    """Gestisce il caricamento di file Excel/CSV/TXT/JSON."""
    if 'file' not in request.files:
        return jsonify({"error": "Nessun file inviato"}), 400
        
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "Nome file vuoto"}), 400
        
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        dest_path = os.path.join(UPLOADS_DIR, filename)
        file.save(dest_path)
        return jsonify({
            "success": True,
            "filename": filename,
            "path": f"uploads/{filename}"
        })
    else:
        return jsonify({"error": "Estensione file non consentita (usa Excel, CSV, TXT o JSON)"}), 400

@app.route('/api/uploads/<filename>', methods=['DELETE'])
def delete_upload(filename):
    """Elimina un file caricato."""
    safe_name = os.path.basename(filename)
    file_path = os.path.join(UPLOADS_DIR, safe_name)
    if os.path.exists(file_path):
        try:
            os.remove(file_path)
            return jsonify({"success": True})
        except Exception as e:
            return jsonify({"error": str(e)}), 500
    return jsonify({"error": "File non trovato"}), 404

@app.route('/api/download/<filename>')
def download_file(filename):
    """Permette di scaricare un file dalla cartella uploads/."""
    # Evita directory traversal
    safe_name = os.path.basename(filename)
    return send_from_directory(UPLOADS_DIR, safe_name, as_attachment=True)

@app.route('/api/run/<filename>')
def run_script(filename):
    """Esegue lo script in tempo reale e invia l'output in streaming (SSE)."""
    # Rimuovi eventuali percorsi per evitare directory traversal, ma mantieni gli spazi
    safe_name = os.path.basename(filename)
    script_path = os.path.join(SCRIPTS_DIR, safe_name)
    
    if not os.path.exists(script_path):
        return Response("data: ❌ Errore: Script non trovato.\n\n", mimetype='text/event-stream')
        
    # Leggi gli argomenti e i file inviati
    args_str = request.args.get('args', '')
    files_str = request.args.get('files', '')
    
    # Costruisci il comando di esecuzione
    cmd = [sys.executable, script_path]
    if args_str:
        cmd.extend(shlex.split(args_str))
        
    # Copia l'ambiente corrente e aggiungi variabili utili
    env = os.environ.copy()
    env['PYTHONUNBUFFERED'] = '1'
    env['PYTHONPATH'] = BASE_DIR # Permette di caricare il mock di google.colab
    
    if files_str:
        env['INPUT_FILES'] = files_str
        
    # Estrai parametri da inviare come variabili d'ambiente per supportare gli script Colab riconvertiti
    if args_str:
        try:
            parts = shlex.split(args_str)
            if "--ref-date" in parts:
                idx = parts.index("--ref-date")
                if idx + 1 < len(parts):
                    env['REF_DATE'] = parts[idx+1]
            if "--out-excel" in parts:
                idx = parts.index("--out-excel")
                if idx + 1 < len(parts):
                    env['OUT_EXCEL'] = parts[idx+1]
            if "--out-dash" in parts:
                idx = parts.index("--out-dash")
                if idx + 1 < len(parts):
                    env['OUT_DASH'] = parts[idx+1]
        except Exception:
            pass
        
    print(f"🚀 Esecuzione comando: {' '.join(cmd)}")
    print(f"   INPUT_FILES: {files_str}")
    
    def generate():
        yield "data: 🚀 Avvio dello script in corso...\n\n"
        try:
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                env=env,
                cwd=BASE_DIR # Esegui dalla root del progetto
            )
            
            # Legge l'output in tempo reale riga per riga
            for line in process.stdout:
                # Invia la riga formattata come evento SSE terminando con \n\n
                yield f"data: {line.rstrip()}\n\n"
                
            process.wait()
            
            # Controlla l'esito dello script
            if process.returncode == 0:
                yield "data: [SUCCESS] Script completato con successo (codice 0).\n\n"
            else:
                yield f"data: [ERROR] Lo script ha riscontrato un errore (codice {process.returncode}).\n\n"
        except Exception as e:
            yield f"data: [ERROR] Impossibile avviare il processo: {e}\n\n"
            
    return Response(generate(), mimetype='text/event-stream')

if __name__ == '__main__':
    # Avvia l'applicazione locale sulla porta 5001
    print("🚀 Avvio del server locale in corso...")
    print("Visualizza la dashboard su: http://127.0.0.1:5001")
    app.run(host='127.0.0.1', port=5001, debug=True)