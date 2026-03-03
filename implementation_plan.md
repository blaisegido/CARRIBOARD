# Logiciel Python de Suivi de Carrière

## Objectif

Créer une application logicielle Python avec interface utilisateur (UI) permettant au gestionnaire de la carrière de :
1. Charger et visualiser les données de l'extraction de la bascule.
2. Afficher un tableau de bord interactif avec KPIs et graphiques (CA, Tonnage, etc.).
3. Filtrer les données (par période, client, produit).
4. (Optionnel/Plus tard) Exécuter le module d'exportation vers Excel vu précédemment.

## Phase 1 : Choix de la Stack Technique

- **Langage** : Python 3
- **Interface Utilisateur (UI)** : `Streamlit` ou `Dash` pour la création rapide d'une interface web interactive, ou `CustomTkinter` pour une application bureau locale lourde. 
  *Recommandation : Streamlit est idéal, très rapide à développer, interactif, et offre de magnifiques capacités de visualisation dès la première ligne.*
- **Manipulation des Données** : `pandas`
- **Visualisation** : `plotly` (graphiques interactifs et modernes).

## Phase 2 : Fonctionnalités de l'Application

### Écran 1 : Chargement et Paramètres
- Bouton pour sélectionner le fichier Excel source.
- Section pour définir quelques paramètres si besoin (ex: objectifs de CA).

### Écran 2 : Tableau de Bord Principal (Dashboard)
- **Filtres (Sidebar)** : Plage de dates, Client, Produit.
- **Cartes KPIs** : CA Total, Tonnage Total, Nb de Livraisons, Prix Moyen/Tonne.
- **Graphiques Interactifs (Plotly)** :
  - Évolution du CA et du Tonnage dans le temps.
  - Répartition du CA par Produit (Pie Chart).
  - Ventes par Produit (Tonnage).
  - Top Clients (Bar Chart horizontale).

### Écran 3 : Vue Détaillée des Données
- Affichage du tableau de données nettoyées et filtrées, avec possibilité de trier les colonnes et de faire une recherche textuelle.

## Phase 3 : Nettoyage et Traitement (Backend Python)

Cette partie réutilisera la logique que nous avons commencé à élaborer pour l'Excel :
- Nettoyer les en-têtes (trouver la ligne des colonnes).
- Caster les types de données (dates, formats numériques).
- Remplir les valeurs manquantes (ex: prix) ou calculer le CA si nécessaire.

## Phase 4 : Déploiement Local

L'application sera livrée sous forme d'un script `app.py`.
Je fournirai une commande simple (fichier `.bat` sous Windows) pour que l'utilisateur puisse lancer l'application en un double-clic sans avoir à manipuler le terminal.

---

**Êtes-vous d'accord pour utiliser *Streamlit* ?** C'est la technologie standard aujourd'hui pour construire rapidement des tableaux de bord de données magnifiques en Python.
