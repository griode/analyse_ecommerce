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