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

    print(f"Erreur: {e}")