import pandas as pd
import mysql.connector
from mysql.connector import Error
import tkinter as tk
from tkinter import ttk, messagebox
from pandastable import Table

# Configuration
EXCEL_FILE = "BASE_DE_PROSPECTION_CELTIIS_ABDALLAH-SURFACE-PRO8-GZ.xlsx"
SHEET_NAME = "1 CONSEILS COM"
DB_CONFIG = {
    'host': 'localhost',
    'user': 'root',
    'password': '',
    'database': 'planning'
}

class DataImporterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Importation de données")
        self.root.geometry("1200x800")
        
        # Cadre principal
        main_frame = ttk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Titre
        ttk.Label(main_frame, text="Aperçu des données avant importation", 
                 font=("Arial", 14, "bold")).pack(pady=10)
        
        # Cadre pour le tableau
        table_frame = ttk.Frame(main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # Chargement des données
        self.df = self.load_and_prepare_data()
        
        if self.df is not None:
            # Affichage du tableau
            self.table = self.create_table(table_frame)
            
            # Boutons
            btn_frame = ttk.Frame(main_frame)
            btn_frame.pack(pady=10)
            
            ttk.Button(btn_frame, text="Valider l'importation", 
                      command=self.confirm_import).pack(side=tk.LEFT, padx=5)
            ttk.Button(btn_frame, text="Annuler", 
                      command=root.destroy).pack(side=tk.LEFT, padx=5)
            
            # Informations
            info_text = f"Fichier: {EXCEL_FILE}\nFeuille: {SHEET_NAME}\n{len(self.df)} lignes prêtes à l'importation"
            ttk.Label(main_frame, text=info_text).pack(pady=5)
        else:
            ttk.Label(main_frame, text="Erreur lors du chargement des données", 
                     foreground="red").pack(pady=20)
            ttk.Button(main_frame, text="Fermer", 
                      command=root.destroy).pack(pady=10)

    def load_and_prepare_data(self):
        try:
            # Lecture du fichier Excel
            df = pd.read_excel(
                EXCEL_FILE, 
                sheet_name=SHEET_NAME,
                header=1,  # Utiliser la deuxième ligne comme en-têtes
                dtype=str
            )
            
            # Nettoyage des noms de colonnes
            df.columns = [str(col).strip() for col in df.columns]
            
            # Vérification des colonnes attendues
            expected_columns = [
                'NOM ENTREPRISES', 
                'NOM DU POINT FOCAL', 
                'FONCTION', 
                'Téléphone', 
                'EMAIL', 
                'VILLE', 
                'ACTIVITE PRINCIPALE', 
                'IFU', 
                'QIP', 
                'RCCM'
            ]
            
            # Sélection des colonnes nécessaires
            df = df[expected_columns]
            
            # Remplacement des valeurs manquantes
            df = df.fillna('')
            
            # Ajout des colonnes fixes
            df['numero_portefeuille'] = '1'
            df['portefeuille'] = 'Conseils et Communication_Isnelle HOUEKPONHOUNDE'
            
            return df
        
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du chargement des données:\n{str(e)}")
            return None

    def create_table(self, parent):
        # Création d'un cadre pour la table
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Création de la table avec pandastable
        pt = Table(
            frame, 
            dataframe=self.df,
            showtoolbar=True,
            showstatusbar=True,
            width=1100,
            height=600
        )
        pt.show()
        return pt

    def confirm_import(self):
        if messagebox.askyesno(
            "Confirmation",
            f"Êtes-vous sûr de vouloir importer {len(self.df)} enregistrements?\n"
            "Cette action est irréversible."
        ):
            self.import_to_database()
    
    def import_to_database(self):
        try:
            # Renommage des colonnes pour correspondre à la base de données
            df_import = self.df.rename(columns={
                'NOM ENTREPRISES': 'nom_Entreprise',
                'NOM DU POINT FOCAL': 'nom_Responsable',
                'FONCTION': 'fonction',
                'Té': 'telephone',
                'EMAIL': 'email',
                'VILLE': 'ville',
                'ACTIVITE PRINCIPALE': 'activite_Principale',
                'IFU': 'IFU',
                'QIP': 'QIP',
                'RCCM': 'RCCM'
            })
            
            # Connexion à la base de données
            conn = mysql.connector.connect(**DB_CONFIG)
            cursor = conn.cursor()
            
            # Requête d'insertion
            insert_query = """
            INSERT INTO portefeuilles (
                numero_portefeuille, nom_Entreprise, nom_Responsable, fonction,
                telephone, email, ville, activite_Principale, IFU, QIP, RCCM, portefeuille
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """
            
            # Conversion des données
            data_tuples = [tuple(x) for x in df_import.to_numpy()]
            
            # Exécution de l'insertion
            cursor.executemany(insert_query, data_tuples)
            conn.commit()
            
            messagebox.showinfo(
                "Succès",
                f"{cursor.rowcount} lignes importées avec succès!"
            )
            
        except Error as e:
            messagebox.showerror(
                "Erreur MySQL",
                f"Erreur lors de l'importation:\n{str(e)}"
            )
        except Exception as e:
            messagebox.showerror(
                "Erreur",
                f"Erreur inattendue:\n{str(e)}"
            )
        finally:
            if 'conn' in locals() and conn.is_connected():
                cursor.close()
                conn.close()
            self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = DataImporterApp(root)
    root.mainloop()