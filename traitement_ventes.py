import pandas as pd
import re
from difflib import get_close_matches
from pyxlsb import open_workbook

# Liste de couleurs courantes à adapter si besoin
COULEURS = [
    'blanc', 'noir', 'gris', 'bleu', 'rouge', 'vert', 'jaune', 'marron', 'beige', 'rose', 'violet', 'orange', 'doré', 'argent', 'cuivre', 'taupe', 'chêne', 'noyer', 'pin', 'anthracite'
]

def lire_xlsb(file_path, sheet=0):
    # Lecture d'un fichier xlsb en DataFrame
    with open_workbook(file_path) as wb:
        with wb.get_sheet(wb.sheets[sheet]) as sheet:
            data = [row for row in sheet.rows()]
            columns = [c.v for c in data[0]]
            rows = [[c.v for c in row] for row in data[1:]]
            return pd.DataFrame(rows, columns=columns)

def extraire_dimension(description):
    # Extrait la première dimension trouvée (ex: 140x190, 200*90, 90 x 200)
    if not isinstance(description, str):
        return ''
    match = re.search(r'(\d{2,4})\s*[xX*]\s*(\d{2,4})', description)
    if match:
        return f"{match.group(1)}x{match.group(2)}"
    return ''

def extraire_couleur(description):
    if not isinstance(description, str):
        return ''
    desc = description.lower()
    for couleur in COULEURS:
        if couleur in desc:
            return couleur
    return ''

def recategoriser_nature(row, natures_valides):
    # Vérifie si la nature est bien présente dans la liste des natures valides
    nature = str(row['Nature']).strip()
    if nature in natures_valides:
        return nature
    # Sinon, cherche la catégorie la plus proche (fuzzy matching)
    proches = get_close_matches(nature, natures_valides, n=1, cutoff=0.6)
    if proches:
        return proches[0]
    return nature  # Si rien trouvé, garde la valeur d'origine

def main():
    # Charger les données
    df = lire_xlsb('data/ecommerce_sales.xlsb')
    
    # Nettoyage des colonnes (adapter si besoin)
    df.columns = [col.strip().lower() for col in df.columns]
    
    # Liste des catégories valides
    natures_valides = set(df['Nature'].dropna().astype(str).str.strip())
    
    # Recatégorisation
    df['nature_corrigee'] = df.apply(lambda row: recategoriser_nature(row, natures_valides), axis=1)
    
    # Extraction dimensions et couleurs
    df['dimension'] = df['Libellé produit'].apply(extraire_dimension)
    df['couleur'] = df['Libellé produit'].apply(extraire_couleur)
    
    # Export
    df.to_csv('ecommerce_sales_clean.csv', index=False)
    print("Traitement terminé. Fichier 'ecommerce_sales_clean.csv' généré.")

if __name__ == "__main__":
    main()
