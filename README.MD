# Introduction

Ce fichier ``/data/ecommerce_sales.xlsb`` contient les lignes de ventes d’un site e-commerce (marketplace) spécialisé dans l’ameublement.

**Principe d’une marketplace :**  
Le site vend les produits de différents vendeurs tiers (voir colonne C).  
Les produits sont regroupés par univers (niveau 1, le plus macro) et par nature (niveau 2).

## Objectifs

- **Détection et correction de catégorisation :**  
    Écrire un algorithme pour identifier les lignes mal catégorisées et les recatégoriser dans tout le dataset. Les catégories cibles sont celles déjà présentes dans la colonne "nature". Il n’est pas autorisé de créer de nouvelles catégories.

- **Extraction de dimensions et couleurs :**  
    Écrire un algorithme pour extraire la dimension et la couleur (si présente dans la description).  
    Exemple : pour « Matelas mousse 140x190 cm », l’algorithme doit extraire la dimension 140x190.

## Structure du projet

- `data/`  
  Contient les fichiers de données sources, notamment `ecommerce_sales.xlsb`.

- `notebooks/`  
  Contient des notebooks Jupyter pour l’exploration, la visualisation et la validation des résultats.

## Exécution du code

1. **Creer python env** 

   Depuis la racine du projet, exécutez:
   ```bash
   python3 -m venv venv
   ```

2. **Installer les dépendances**  
   Depuis la racine du projet, exécutez :
   ```bash
   pip install -r requirements.txt
   ```

3. **Lancer l’analyse**  
 ouvrez et exécutez les notebooks dans le dossier `notebooks/` pour une analyse interactive.
