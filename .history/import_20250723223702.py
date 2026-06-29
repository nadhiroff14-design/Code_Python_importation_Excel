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
# Définition des colonnes attendues
EXPECTED_COLUMNS = [
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

try:
    # Lecture du fichier Excel en spécifiant la ligne 1 comme en-tête (index 1)
    df = pd.read_excel(
        EXCEL_FILE, 
        sheet_name=SHEET_NAME,
        header=1,  # Utiliser la deuxième ligne comme en-têtes
        dtype=str
    )
    print(f"Lecture réussie : {len(df)} lignes trouvées")
    
    # Nettoyage des noms de colonnes
    df.columns = [str(col).strip() for col in df.columns]
    
    # Vérification des colonnes
    print("Colonnes trouvées dans le fichier Excel:")
    print(df.columns.tolist())
    
    # Vérification des colonnes attendues
    missing_columns = [col for col in EXPECTED_COLUMNS if col not in df.columns]
    if missing_columns:
        print("\nERREUR: Colonnes manquantes dans le fichier Excel:")
        print(missing_columns)
        exit(1)
    
    # Préparation des données
    df = df[EXPECTED_COLUMNS]  # Sélection des colonnes nécessaires
    df = df.where(pd.notnull(df), None)  # Remplacement des NaN par None
    
    # Ajout des colonnes fixes
    df['numero_portefeuille'] = '1'
    df['portefeuille'] = 'Conseils et Communication_Isnelle HOUEKPONHOUNDE'
    
    # Renommage des colonnes pour correspondre à la base de données
    df = df.rename(columns={
        'NOM ENTREPRISES': 'nom_Entreprise',
        'NOM DU POINT FOCAL': 'nom_Responsable',
        'FONCTION': 'fonction',
        'Télépo': 'telephone',
        'EMAIL': 'email',
        'VILLE': 'ville',
        'ACTIVITE PRINCIPALE': 'activite_Principale',
        'IFU': 'IFU',
        'QIP': 'QIP',
        'RCCM': 'RCCM'
    })
    
    # Connexion à la base de données
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor()
        print("Connexion à la base de données réussie")
        
        # Requête d'insertion
        insert_query = """
        INSERT INTO portefeuilles (
            numero_portefeuille, nom_Entreprise, nom_Responsable, fonction,
            telephone, email, ville, activite_Principale, IFU, QIP, RCCM, portefeuille
        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        
        # Conversion des données
        data_tuples = [tuple(x) for x in df.to_numpy()]
        
        # Exécution de l'insertion
        cursor.executemany(insert_query, data_tuples)
        conn.commit()
        print(f"{cursor.rowcount} lignes insérées avec succès")
        
    except Error as e:
        print(f"Erreur MySQL: {e}")
        conn.rollback()
    finally:
        if conn.is_connected():
            cursor.close()
            conn.close()
            
except Exception as e:
    print(f"Erreur lors de la lecture du fichier Excel: {e}")
