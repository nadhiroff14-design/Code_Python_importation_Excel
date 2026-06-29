import pandas as pd
import mysql.connector
from mysql.connector import Error
import sys
import os

def load_excel_data(file_path, sheet_name, columns):
    """Charge les données Excel avec gestion robuste des erreurs"""
    try:
        print(f"Chargement du fichier {file_path}...")
        df = pd.read_excel(
            file_path,
            sheet_name=sheet_name,
            usecols=columns,
            dtype={'TELEPHONE': str, 'IFU': str, 'QIP': str, 'RCCM': str}
        ).replace({pd.NA: None})
        print(f"{len(df)} lignes chargées avec succès")
        return df
    except Exception as e:
        print(f"ERREUR: Impossible de lire le fichier Excel - {str(e)}")
        sys.exit(1)

def db_connection(config):
    """Établit une connexion MySQL avec reconnexion automatique"""
    try:
        conn = mysql.connector.connect(**config)
        print("Connexion MySQL établie")
        return conn
    except Error as e:
        print(f"ERREUR MySQL: {str(e)}")
        sys.exit(1)

def main():
    # Configuration centrale
    CONFIG = {
        'excel': {
            'path': 'BASE_DE_PROSPECTION_CELTIIS_ABDALLAH-SURFACE-PRO8-GZ.xlsx',
            'sheet': '1 CONSEILS COM',
            'columns': [
                'NOM ENTREPRISES',
                'NOM DU POINT FOCALE',
                'FONCTION',
                'TELEPHONE',
                'EMAIL',
                'VILLE',
                'ACTIVITE PRINCIPALE',
                'IFU',
                'QIP',
                'RCCM'
            ]
        },
        'database': {
            'host': 'localhost',
            'user': 'root',
            'password': '',
            'database': 'planning',
            'charset': 'utf8mb4'
        },
        'mapping': {
            'numero_portefeuille': 1,
            'nom_portefeuille': 'Conseils et Communication_Isnelle HOUEKPONHOUNDE'
        }
    }

    # Vérification du fichier
    if not os.path.exists(CONFIG['excel']['path']):
        print(f"ERREUR: Fichier {CONFIG['excel']['path']} introuvable")
        sys.exit(1)

    # 1. Chargement des données
    df = load_excel_data(
        CONFIG['excel']['path'],
        CONFIG['excel']['sheet'],
        CONFIG['excel']['columns']
    )

    # 2. Connexion à la base de données
    conn = db_connection(CONFIG['database'])
    cursor = conn.cursor()

    # 3. Préparation de la requête SQL
    sql = """
    INSERT INTO portefeuilles (
        numero_portefeuille, nom_portefeuille, nom_Entreprise,
        nom_Responsable, fonction, telephone, email, ville,
        activite_Principale, IFU, QIP, RCCM, portefeuille
    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    """

    # 4. Importation des données
    success = 0
    errors = 0

    for index, row in df.iterrows():
        try:
            data = (
                CONFIG['mapping']['numero_portefeuille'],
                CONFIG['mapping']['nom_portefeuille'],
                row['NOM ENTREPRISES'],
                row['NOM DU POINT FOCALE'],
                row['FONCTION'],
                row['TELEPHONE'],
                row['EMAIL'],
                row['VILLE'],
                row['ACTIVITE PRINCIPALE'],
                row['IFU'],
                row['QIP'],
                row['RCCM'],
                CONFIG['mapping']['nom_portefeuille']
            )
            cursor.execute(sql, data)
            success += 1
        except Error as e:
            print(f"Erreur ligne {index+1}: {str(e)}")
            errors += 1
            continue

    conn.commit()
    print(f"\nRésultat final: {success} importations réussies, {errors} erreurs")

    # Fermeture propre
    cursor.close()
    conn.close()

if __name__ == "__main__":
    main()