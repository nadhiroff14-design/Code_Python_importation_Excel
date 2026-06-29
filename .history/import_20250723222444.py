import pandas as pd
import mysql.connector
from mysql.connector import Error

# 1. Configuration
EXCEL_FILE = "BASE_DE_PROSPECTION_CELTIIS_ABDALLAH-SURFACE-PRO8-GZ.xlsx"
SHEET_NAME = "1 CONSEILS COM"
DB_CONFIG = {
    'host': 'localhost',
    'user': 'root',
    'password': '',
    'database': 'planning'
}
try:
    # Lire sans dtype=str pour mieux voir les problèmes
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    print(f"Lecture réussie : {len(df)} lignes trouvées")
    
    # Afficher les colonnes réelles pour diagnostic
    print("Colonnes réelles dans le fichier Excel:")
    print(df.columns.tolist())
    
    # Nettoyage des noms de colonnes
    df.columns = [col.strip() for col in df.columns]
    
    # 2. Mapping des colonnes (vérifier avec les noms réels)
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
    
    # Vérifier les colonnes manquantes
    missing_columns = [excel_col for excel_col in COLUMN_MAPPING.keys() if excel_col not in df.columns]
    if missing_columns:
        print("\nERREUR: Colonnes manquantes dans le fichier Excel:")
        print(missing_columns)
        print("Colonnes disponibles:", df.columns.tolist())
        exit(1)
    
    # Renommage des colonnes
    df = df.rename(columns=COLUMN_MAPPING)
    
    # 3. Ajout des colonnes fixes
    df['numero_portefeuille'] = '1'
    df['portefeuille'] = 'Conseils et Communication_Isnelle HOUEKPONHOUNDE'
    
    # 4. Sélection des colonnes finales
    final_columns = [
        'numero_portefeuille', 'nom_Entreprise', 'nom_Responsable', 
        'fonction', 'telephone', 'email', 'ville', 
        'activite_Principale', 'IFU', 'QIP', 'RCCM', 'portefeuille'
    ]
    df = df[final_columns]
    
    # 5. Connexion à la base de données
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor()
        
        # 6. Préparation de la requête
        insert_query = """
        INSERT INTO portefeuilles (
            numero_portefeuille, nom_Entreprise, nom_Responsable, fonction,
            telephone, email, ville, activite_Principale, IFU, QIP, RCCM, portefeuille
        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        
        # Conversion des données
        data = [tuple(x) for x in df.to_numpy()]
        
        # 7. Exécution
        cursor.executemany(insert_query, data)
        conn.commit()
        print(f"\nSUCCÈS: {cursor.rowcount} lignes insérées dans la base de données")
        
    except Error as e:
        print(f"Erreur MySQL: {e}")
    finally:
        if conn.is_connected():
            cursor.close()
            conn.close()
            
except Exception as e:
    print(f"Erreur: {e}")