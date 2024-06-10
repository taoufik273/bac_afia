import os
import sys
from tkinter import filedialog, messagebox
import pandas as pd
import sqlite3

def resource_path(relative_path):
    """ Obtenir le chemin absolu vers la ressource, fonctionne pour dev et pour PyInstaller """
    try:
        # PyInstaller crée une variable temporaire pour stocker le chemin d'accès aux fichiers.
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def importer_donnees():
    # Ouvrir une boîte de dialogue pour sélectionner le fichier Excel
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    
    # Si aucun fichier n'est sélectionné, retourner
    if not filename:
        return
    
    try:
        # Lire le fichier Excel
        df = pd.read_excel(filename)
        
        # Connexion à la base de données SQLite
        db_path = resource_path('data/saisie.db')
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        # Créer la table si elle n'existe pas déjà, avec une contrainte d'unicité
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS saisie (
            numero INTEGER,
            jour INTEGER,
            heure INTEGER,
            serie INTEGER,
            ordre INTEGER,
            massar TEXT UNIQUE,
            cin TEXT,
            nom TEXT,
            sexe TEXT,
            gym TEXT,
            course TEXT,
            poid TEXT,
            saut TEXT
        )
        ''')

        # Ajouter les colonnes manquantes avec des valeurs vides
        df['gym'] = ""
        df['course'] = ""
        df['poid'] = ""
        df['saut'] = ""

        # Préparation de la requête d'insertion
        insert_query = '''
        INSERT INTO saisie (numero, jour, heure, serie, ordre, massar, cin, nom, sexe, gym, course, poid, saut)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        '''

        # Conversion des données du DataFrame en liste de tuples
        data_to_insert = df.values.tolist()

        # Vérifier et insérer les données une par une pour éviter les doublons
        for record in data_to_insert:
            cursor.execute('SELECT COUNT(*) FROM saisie WHERE massar = ?', (record[5],))
            if cursor.fetchone()[0] == 0:
                cursor.execute(insert_query, record)

        # Commit the transaction
        conn.commit()

        # Fermeture de la connexion
        cursor.close()
        conn.close()

        # Afficher un message de succès
        messagebox.showinfo("ممتاز", "تم الاستيراد بنجاح")

    except Exception as e:
        # Afficher un message d'erreur en cas de problème
        messagebox.showerror("Erreur", "Une erreur s'est produite lors de l'importation: " + str(e))

# Appel de la fonction importer_donnees() au démarrage du script
if __name__ == "__main__":
    importer_donnees()
