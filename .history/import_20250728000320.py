import pandas as pd
import mysql.connector
from mysql.connector import Error
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import os

# Configuration
EXCEL_FILE = "BASE_DE_PROSPECTION_CELTIIS_ABDALLAH-SURFACE-PRO8-GZ.xlsx"
SHEET_NAME = "2 BTP"
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
        self.root.state('zoomed')  # Ouvrir en plein écran
        
        # Cadre principal avec barre de défilement
        main_frame = ttk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Titre
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(title_frame, text="Aperçu des données avant importation", 
                 font=("Arial", 14, "bold")).pack(side=tk.LEFT)
        
        # Informations sur le fichier
        file_info = ttk.Label(title_frame, 
                             text=f"Fichier: {os.path.basename(EXCEL_FILE)} | Feuille: {SHEET_NAME}",
                             font=("Arial", 10))
        file_info.pack(side=tk.RIGHT, padx=10)
        
        # Cadre pour le tableau avec barre de défilement
        table_frame = ttk.LabelFrame(main_frame, text="Données à importer")
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Chargement des données
        self.df = self.load_and_prepare_data()
        
        if self.df is not None:
            # Affichage du tableau dans un widget Text
            self.text_widget = scrolledtext.ScrolledText(table_frame, wrap=tk.NONE)
            self.text_widget.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            self.display_dataframe()
            
            # Boutons
            btn_frame = ttk.Frame(main_frame)
            btn_frame.pack(fill=tk.X, pady=(10, 0))
            
            ttk.Button(btn_frame, text="Valider l'importation", 
                      command=self.confirm_import, width=20).pack(side=tk.LEFT, padx=5)
            ttk.Button(btn_frame, text="Annuler", 
                      command=root.destroy, width=10).pack(side=tk.RIGHT, padx=5)
            
            # Informations sur le nombre de lignes
            ttk.Label(btn_frame, 
                     text=f"{len(self.df)} lignes prêtes à l'importation",
                     font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=20)
            
        else:
            error_frame = ttk.Frame(main_frame)
            error_frame.pack(fill=tk.BOTH, expand=True, pady=20)
            ttk.Label(error_frame, text="Erreur lors du chargement des données", 
                     foreground="red", font=("Arial", 12, "bold")).pack()
            ttk.Label(error_frame, text="Veuillez vérifier le fichier et réessayer", 
                     font=("Arial", 10)).pack(pady=10)
            ttk.Button(error_frame, text="Fermer", 
                      command=root.destroy, width=15).pack(pady=20)

    def display_dataframe(self):
        """Affiche le DataFrame dans un widget Text"""
        self.text_widget.delete(1.0, tk.END)
        
        # Création d'une représentation texte du DataFrame
        columns = self.df.columns.tolist()
        col_widths = [max(self.df[col].astype(str).map(len).max(), len(col)) for col in columns]
        total_width = sum(col_widths) + 3 * len(columns) + 1
        
        # En-têtes
        header = " | ".join([col.ljust(width) for col, width in zip(columns, col_widths)])
        separator = "-" * total_width
        
        self.text_widget.insert(tk.END, header + "\n")
        self.text_widget.insert(tk.END, separator + "\n")
        
        # Données
        for _, row in self.df.iterrows():
            row_str = " | ".join([str(row[col]).ljust(width) for col, width in zip(columns, col_widths)])
            self.text_widget.insert(tk.END, row_str + "\n")
        
        self.text_widget.config(state=tk.DISABLED)

    def load_and_prepare_data(self):
        try:
            # Lecture du fichier Excel
            df = pd.read_excel(
    EXCEL_FILE, 
    sheet_name=SHEET_NAME,
    header=1,  # ou header=None si les entêtes sont mal positionnées
    dtype=str
)
print("Colonnes trouvées :", df.columns.tolist())

            
            # Nettoyage des noms de colonnes
            df.columns = [str(col).strip() for col in df.columns]
            
            # Vérification des colonnes attendues
            expected_columns = [
                'NOM ENTREPRISES', 
                'Téléphone', 
                'EMAIL',
                'VILLE',
                'NOM DU RESPONSABLE',  
                'ACTIVITE PRINCIPALE', 
                'IFU', 
                'QIP', 
                'RCCM'
            ]
            
            # Sélection des colonnes nécessaires
            df = df[expected_columns]
            
            # Remplacement des valeurs manquantes
            df = df.fillna('')
            
            # Ajout des colonnes fixes - numero_portefeuille en PREMIER
            df.insert(0, 'numero_portefeuille', '2')  # Ajout en première position
            df['portefeuille'] = 'BTP_Franck AROUNA'
            
            return df
        
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du chargement des données:\n{str(e)}")
            return None

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
                'NOM DU RESPONSABLE': 'nom_Responsable',
                'Téléphone': 'telephone',
                'EMAIL': 'email',
                'VILLE': 'ville',
                'ACTIVITE PRINCIPALE': 'activite_Principale',
                'IFU': 'IFU',
                'QIP': 'QIP',
                'RCCM': 'RCCM'
            })
            
            # Réorganisation des colonnes pour l'insertion
            # numero_portefeuille est déjà en première position
            final_columns = [
                'numero_portefeuille', 'nom_Entreprise', 'nom_Responsable', 
                'telephone', 'email', 'ville', 
                'activite_Principale', 'IFU', 'QIP', 'RCCM', 'portefeuille'
            ]
            df_import = df_import[final_columns]
            
            # Connexion à la base de données
            conn = mysql.connector.connect(**DB_CONFIG)
            cursor = conn.cursor()
            
            # Requête d'insertion
            insert_query = """
            INSERT INTO portefeuilles (
                numero_portefeuille, nom_Entreprise, nom_Responsable,
                telephone, email, ville, activite_Principale, IFU, QIP, RCCM, portefeuille
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
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