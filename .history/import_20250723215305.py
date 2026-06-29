import pandas as pd
from sqlalchemy import create_engine
import pymysql

# 1. Configuration de la connexion MySQL
db_config = {
    'host': 'localhost',  # ou l'adresse de votre serveur MySQL
    'user': 'votre_utilisateur',
    'password': 'votre_mot_de_passe',
    'database': 'nom_de_votre_base',
    'port': 3306  # port par défaut de MySQL
}

# 2. Chargement du fichier Excel
try:
    df_excel = pd.read_excel(
        'BASE_DE_PROSPECTION_CELTIIS_ABDALLAH-SURFACE-PRO8-GZ.xlsx',
        sheet_name='1 CONSEILS COM',
        engine='openpyxl'
    )
    
    # 3. Mapping des colonnes et création du DataFrame final
    df_final = pd.DataFrame({
        'numero_portefeuille': 1,  # Valeur fixe pour cette feuille
        'portefeuille': "Conseils et Communication_Isnelle HOUEKPONHOUNDE",
        'nom_Entreprise': df_excel['NOM ENTREPRISES'],
        'nom_Responsable': df_excel['NOM DU POINT FOCALE'],
        'fonction': df_excel.get('FONCTION', None),  # .get() pour gérer les colonnes optionnelles
        'telephone': df_excel['TELEPHONE'],
        'email': df_excel.get('EMAIL', None),
        'ville': df_excel['VILLE'],
        'activite_Principale': df_excel['ACTIVITE PRINCIPALE'],
        'IFU': df_excel['IFU'],
        'QIP': df_excel['QIP'],
        'RCCM': df_excel['RCCM']
    })
    
    # 4. Nettoyage des données (remplacement des NaN par None)
    df_final = df_final.where(pd.notnull(df_final), None)
    
    # 5. Connexion à MySQL et insertion des données
    engine = create_engine(f"mysql+pymysql://{db_config['user']}:{db_config['password']}@{db_config['host']}/{db_config['database']}")
    
    df_final.to_sql(
        name='portefeuilles',
        con=engine,
        if_exists='append',  # Options: 'fail', 'replace', 'append'
        index=False,
        chunksize=1000  # Pour les gros fichiers
    )
    
    print(f"Succès: {len(df_final)} lignes insérées dans MySQL")
    
except FileNotFoundError:
    print("Erreur: Fichier Excel introuvable")
except KeyError as e:
    print(f"Erreur: Colonne manquante dans Excel - {str(e)}")
except Exception as e:
    print(f"Erreur inattendue: {str(e)}")
finally:
    if 'engine' in locals():
        engine.dispose()