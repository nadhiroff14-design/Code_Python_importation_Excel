import pandas as pd
import mysql.connector
from mysql.connector import Error

# 1. Configuration
EXCEL_FILE = "BASE_DE_PROSPECTION_CELTIIS_ABDALLAH-SURFACE-PRO8-GZ.xlsx"
SHEET_NAME = "1 CONSEILS COM"
DB_CONFIG = {
    'host': 'localhost',
    'user': 'root',
    'password': 'votre_mt_de_passe',
    'database': 'planning'
}

# 2. Mapping des colonnes
COLUMN_MAPPING = {
    'NOM ENTREPRISES': 'nom_Entreprise',
    'NOM DU POINT FOCALE': 'nom_Responsable',
    'FONCTION': 'fonction',
    'TELEPHONE': 'telephone',
    'EMAIL': 'email',
    'VILLE': 'ville',
    'ACTIVITE PRINCIPALE': 'activite_Principale',
    'IFU': 'IFU',
    'QIP': 'QIP',
    'RCCM': 'RCCM'
}

# 3. Lecture du fichier Excel
try:
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, dtype=str)
    print(f"Lecture réussie : {len(df)} lignes trouvées")
    
    # Nettoyage des colonnes
    df.rename(columns=lambda x: x.strip(), inplace=True)  # Supprimer les espaces autour des noms de colonnes
    
    # 4. Préparation des données
    # Ajout des colonnes fixes
    df['numero_portefeuille'] = '1'
    df['portefeuille'] = 'Conseils et Communication_Isnelle HOUEKPONHOUNDE'
    
    # Renommage des colonnes
    df.rename(columns=COLUMN_MAPPING, inplace=True)
    
    # Sélection des colonnes finales
    final_columns = ['numero_portefeuille', 'nom_Entreprise', 'nom_Responsable', 
                    'fonction', 'telephone', 'email', 'ville', 
                    'activite_Principale', 'IFU', 'QIP', 'RCCM', 'portefeuille']
    df = df[final_columns]
    
    # Remplacement des NaN par None
    df = df.where(pd.notnull(df), None)
    
    # 5. Connexion à la base de données
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor()
        print("Connexion à la base de données réussie")
        
        # 6. Requête d'insertion
        insert_query = """
        INSERT INTO portefeuilles (
            numero_portefeuille, nom_Entreprise, nom_Responsable, fonction,
            telephone, email, ville, activite_Principale, IFU, QIP, RCCM, portefeuille
        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        
        # Conversion du DataFrame en liste de tuples
        data_tuples = [tuple(x) for x in df.to_records(index=False)]
        
        # 7. Exécution de l'insertion
        cursor.executemany(insert_query, data_tuples)
        conn.commit()
        print(f"{cursor.rowcount} lignes insérées avec succès")
        
    except Error as e:
        print(f"Erreur MySQL: {e}")
    finally:
        if conn.is_connected():
            cursor.close()
            conn.close()
            
except Exception as e:
    print(f"Erreur lors de la lecture du fichier Excel: {e}")