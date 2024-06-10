from flask import Flask, request, jsonify, send_from_directory
import sqlite3
from flask_cors import CORS

DATABASE = 'data/saisie.db'

def get_db_connection():
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    return conn

def init_app(app):
    CORS(app)

    @app.route('/api/data', methods=['GET'])
    def get_data():
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM saisie")
        rows = cursor.fetchall()
        conn.close()

        data = [{key: row[key] if row[key] is not None else '' for key in row.keys()} for row in rows]

        return jsonify(data)

    @app.route('/api/data', methods=['PUT'])
    def update_data():
        updates = request.json
        conn = get_db_connection()
        cursor = conn.cursor()
        for update in updates:
            cursor.execute('''
                UPDATE saisie
                SET gym = ?, course = ?, poid = ?, saut = ?
                WHERE numero = ?
            ''', (update['gym'], update['course'], update['poid'], update['saut'], update['numero']))
        conn.commit()
        conn.close()
        return jsonify({"ممتاز": "تم الحفظ بنجاح"})

    @app.route('/index')
    def index():
        return send_from_directory('static', 'index.html')

    return app