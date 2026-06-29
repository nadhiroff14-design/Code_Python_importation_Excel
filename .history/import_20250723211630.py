import pandas as pd
import mysql.connector
from mysql.connector import Error

# Configuration
excel_file = 'BASE_DE_PROSPECTION_CELTIIS_ABDALLAH-SURFACE-PRO8-GZ.xlsx'
sheet_name = '1 CONSEILS COM'
portefeuille_num = 1
portefeuille_nom = 'Conseils et Communication_Isnelle HOUEKPONHOUNDE'

try:
    # 1. Lecture du fichier Excel
    df = pd.read_excel(
        excel_file,
        sheet_name=sheet_name,
        usecols=[
            'NOM ENTREPRISES',        # A
            'NOM DU POINT FOCALE',    # B (nom_Responsable)
            'FONCTION',               # C
            'TELEPHONE',              # D
            'EMAIL',                  # E
            'VILLE',                  # F
            'ACTIVITE PRINCIPALE',    # G
            'IFU',                    # H
            'QIP',                    # I
            'RCCM'                    # J
        ],
        dtype={
            'TELEPHONE': str,  # Pour conserver les zéros initiaux
            'IFU': str,
            'QIP': str,
            'RCCM': str
        }
    )

    # Nettoyage des données
    df = df.where(pd.notnull(df), None)  # Remplace NaN par None pour NULL en SQL

    # 2. Connexion MySQL
    conn = mysql.connector.connect(
        host="localhost",
        user="root",
        password="",
        database="planning",
        charset='utf8mb4'
    )
    
    cursor = conn.cursor()

    # 3. Préparation de la requête SQL
    insert_query = """
    INSERT INTO portefeuilles (
        numero_portefeuille, 
        nom_portefeuille, 
        nom_Entreprise, 
        nom_Responsable, 
        fonction, 
        telephone, 
        email, 
        ville, 
        activite_Principale, 
        IFU, 
        QIP, 
        RCCM, 
        portefeuille
    ) VALUES (
        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
    )
    """

    # 4. Insertion des données
    total_rows = len(df)
    success_count = 0
    
    for index, row in df.iterrows():
        try:
            data = (
                portefeuille_num,                      # numero_portefeuille
                portefeuille_nom,                      # nom_portefeuille
                row['NOM ENTREPRISES'],                # nom_Entreprise
                row['NOM DU POINT FOCALE'],            # nom_Responsable
                row['FONCTION'],                       # fonction
                row['TELEPHONE'],                      # telephone
                row['EMAIL'],                          # email
                row['VILLE'],                          # ville
                row['ACTIVITE PRINCIPALE'],            # activite_Principale
                row['IFU'],                            # IFU
                row['QIP'],                            # QIP
                row['RCCM'],                           # RCCM
                portefeuille_nom                        # portefeuille
            )
            
            cursor.execute(insert_query, data)
            success_count += 1
            
            # Affichage de progression
            if (index + 1) % 100 == 0:
                print(f"Traitement ligne {index + 1}/{total_rows}")
                
        except Error as e:
            print(f"Erreur ligne {index + 1}: {e}")
            conn.rollback()  # Annulation de la transaction en cas d'erreur
            continue
    
    # Validation des changements
    conn.commit()
    
    # Résumé
    print(f"\nImportation terminée avec succès!")
    print(f"- Lignes traitées: {total_rows}")
    print(f"- Lignes importées: {success_count}")
    print(f"- Erreurs: {total_rows - success_count}")

except Error as e:
    print(f"Erreur de connexion MySQL: {e}")
    
except Exception as e:
    print(f"Erreur générale: {e}")

finally:
    # Fermeture propre des connexions
    if 'conn' in locals() and conn.is_connected():
        cursor.close()
        conn.close()
        print("Connexion MySQL fermée")