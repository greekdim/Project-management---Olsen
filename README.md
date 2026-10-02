# Réunion engineering (Odoo)

Lancer : `lancer_extraction.bat` (fait un `git pull` puis démarre le script).

Le script lit Odoo (projets étiquetés « PRO (LIG) » **et** « Engineering », hors étapes Annulé / Cloturé /
Autres / Template / Canceled) et ouvre une page dans le navigateur :

- **Filtres multi-sélection** : étapes, chefs de projet, étiquettes, responsables d'actions
  (clic sur une pastille = ajouter/retirer ; « Tout afficher » = vider un groupe ; « Effacer tous les filtres »).
- **Par projet** : sujets à discuter + actions à entreprendre (case à cocher, responsable, échéance).
- **Sauvegarde automatique** dans `notes_reunion.json` (même dossier que le script), retrouvé à chaque lancement.
  Faites des copies de ce fichier si les notes sont importantes.
- **Export Excel** des projets affichés (onglets Réunion, Synthèse, Projets archivés).
- Bouton « Quitter » (ou Ctrl+C) pour arrêter le serveur local (accessible uniquement depuis cet ordinateur).

## Identifiants Odoo

Au lancement, si aucun `.env` n'existe, le script demande l'identifiant et le mot de passe.
Pour ne plus les saisir : copier `.env.example` en `.env` et remplir `ODOO_USER` et `ODOO_PASSWORD`.
