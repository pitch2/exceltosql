import mysql.connector
import pandas as pd
import numpy as np

def obtenir_noms_feuilles(fichier_path):
    noms_feuilles = pd.ExcelFile(fichier_path).sheet_names
    return noms_feuilles

def traitement_feuille(df): # changer les nom, etude ... en suivant les informations que vous avez besoins de recupérer 
    if 'nom' in df.columns:
        nom = df['nom'].tolist()
        nb_lignes = len(nom)

        if 'etude' in df.columns:
            parcoursup = df['etude'].tolist()
        else:
            parcoursup = [''] * nb_lignes

        if 'email' in df.columns:
            AdresseMail = df['email'].tolist()
        else:
            AdresseMail = [''] * nb_lignes

        if 'bac' in df.columns:
            bac = df['bac'].tolist()
        else:
            bac = [''] * nb_lignes

        if 'classe' in df.columns:
            classe = df['classe'].tolist()
        else:
            classe = [''] * nb_lignes

        # Insérer les données dans la base de données
        for i in range(nb_lignes):
            value = (nom[i], parcoursup[i], AdresseMail[i], bac[i], classe[i])
            ajout(value)
    else:
        print(f"La feuille '{feuille}' ne contient pas la colonne 'nom'.")

def ajout(value):
    try:
        value = tuple('x' if pd.isnull(val) else val for val in value)
    
        sql = "INSERT INTO - (-, -, -, -, -) VALUES (%s, %s , %s , %s , %s)" #changer les - par les colonnes de votre base SQL
        cur.execute(sql, value)
        db.commit()
        print(cur.rowcount, "ligne insérée.")
    except mysql.connector.Error as err:
        print("Erreur d'insertion :", err)
        db.rollback()

fichier_path = "ancien.xls"
feuilles = obtenir_noms_feuilles(fichier_path)

#connexion aux bases de données 
db = mysql.connector.connect(
  host = "-",
  user = "-",
  password = "-",
  database = "-"
)
cur = db.cursor()

# Boucle à travers chaque feuille
for feuille in feuilles:
    df = pd.read_excel(fichier_path, sheet_name=feuille)
    traitement_feuille(df)
