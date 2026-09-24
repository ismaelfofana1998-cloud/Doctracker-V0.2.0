# Stabilisation 0.8.1

Cette livraison corrige des causes de ralentissement et des chemins d'erreur identifiés dans le code. Elle n'ajoute aucune fonctionnalité métier et ne garantit pas l'absence de bugs dans un véritable hôte Excel.

- Ouverture : les index embarqués sont extraits à la demande. Une simple ouverture ne réécrit plus le projet et ne marque plus systématiquement le classeur modifié.
- Recherche et matching : préparation limitée au dossier sélectionné. Les documents ayant déjà échoué sont signalés sans relancer leur OCR à chaque requête ; Réindexer permet une nouvelle tentative. Les anciens résultats disparaissent avant une nouvelle recherche.
- Navigation : plus de reconstitution des listes à chaque preuve, ni de rechargement après chaque sauvegarde. Enregistrer sous conserve la sélection et les plages de matching. Survoler le PDF ne prend plus le clavier à Excel.
- Mémoire : libération des index après une extraction de snip. Fermer le volet ne bloque plus le thread Excel en attendant la fin de l'OCR natif.
- Concurrence : toutes les fenêtres du même classeur désactivent leurs commandes pendant une opération. L'annulation d'une réindexation conserve les index précédents.
- Sauvegarde : les nouveaux index ne deviennent définitifs qu'après écriture des métadonnées. Un échec rétablit leurs identifiants, la révision et la date, et permet une nouvelle tentative. Une erreur de notification ou de nettoyage postérieur ne provoque plus un faux retour arrière après un enregistrement réussi.
- Allègement : suppression des menus et boutons cachés, des relais d'événements devenus inutiles et du mode Lecture inaccessible. Les fonctions utiles restent dans le ruban et les menus du PDF.

## Vérifications

Tests de régression : échec d'écriture réelle avec fichier verrouillé, nouvelle tentative, notification défaillante, ouverture/sauvegarde sans lecture des index, récupération d'une sauvegarde avec index non ouverts, nettoyage des pièces différé. Les essais Windows couvrent également le dossier actif, les fichiers en erreur et l'annulation avant réindexation, en plus des parcours PDF/OCR/interface/export existants.

Validation manuelle encore nécessaire dans Excel : deux fenêtres d'un même classeur, Enregistrer sous, changement de classeur pendant OCR, sauvegarde/réouverture et recherche sur un lot représentatif des documents utilisés. La CI n'exécute pas l'hôte Excel ni Word.
