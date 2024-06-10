import os
import webbrowser
from flask import Flask, send_from_directory, jsonify, request
import subprocess
import app  # Assurez-vous que app.py est dans le même répertoire ou dans le PYTHONPATH

server = Flask(__name__, static_folder='static')  # Le dossier static doit exister

# Configurer le dossier img pour être servi
@server.route('/img/<path:filename>')
def serve_img(filename):
    return send_from_directory('img', filename)

# Intégrer les routes de app.py dans server.py
app.init_app(server)

@server.route('/')
def start_page():
    return send_from_directory('web', 'start.html')

@server.route('/execute/<script_name>', methods=['POST'])
def execute_script(script_name):
    script_map = {
        'import': 'import.py',
        'app': None,  # On ne l'exécute pas comme un script séparé
        'calculate': 'calculate.py',
        'NF': 'NF.py',
        'allnotes': 'allnotes.py'
    }
    
    if script_name in script_map:
        if script_name == 'app':
            return jsonify({'redirect': '/index'})
        else:
            try:
                result = subprocess.run(['python', f'scripts/{script_map[script_name]}'], 
                                     capture_output=True, text=True, check=True)
                return jsonify({'message': result.stdout or 'Script exécuté avec succès'})
            except subprocess.CalledProcessError as e:
                error_msg = e.stderr or f'Erreur lors de l\'exécution de {script_name}'
                return jsonify({'message': error_msg}), 500
    else:
        return jsonify({'message': 'Script non trouvé'}), 404

if __name__ == '__main__':
    # Ouvrir automatiquement start.html dans le navigateur par défaut
    webbrowser.open('http://127.0.0.1:5000/')
    server.run(host='127.0.0.1', port=5000, debug=True)
