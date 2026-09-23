# Recette Excel Windows 0.3

À exécuter dans Excel de bureau après installation de l'artefact **du bon commit**.
Ne pas utiliser de pièces confidentielles pour la première recette.

1. Fermer Excel, installer, ouvrir un classeur, l'enregistrer localement.
   Vérifier le ruban et l'ouverture du volet. Un double-clic sur une cellule
   ordinaire ne doit rien déclencher dans Doctracker.
2. Importer PDF natif paysage, scan PDF, PNG et TIFF multipage. Le premier document
   doit apparaître immédiatement. Vérifier toutes les pages, le zoom, le déplacement
   et la sélection dans les quatre coins. Le crop doit correspondre à la zone tracée.
3. Annuler un import/indexage puis relancer une recherche : les pièces restantes
   doivent être reprises. Importer une pièce incorrecte, constater son erreur,
   puis la retirer de la liste pour débloquer la recherche automatique.
4. Extraire texte, `12 345,67`, `(1 250,50)` et une date. Vérifier valeurs/formats,
   commentaires et avance de ligne. Refaire la date dans un classeur 1904.
5. Essayer une zone sans texte et deux nombres en mode Nombre : aucune preuve
   ne doit être créée. Essayer un texte commençant par `=1+1` : il reste du texte.
6. Tableau : vérifier les colonnes dans l'aperçu, corriger une valeur, insérer,
   puis double-cliquer sur plusieurs cellules (pas seulement la première).
   Annuler l'aperçu : aucune cellule/preuve créée.
7. Somme : sélectionner une cellule vide, dessiner successivement `10,00`,
   `5,25`, `1,25`. Attendu : `16,50` dans la même cellule avec trois preuves.
   Double-cliquer et vérifier chacune. Une modification manuelle de la somme
   doit empêcher l'ajout silencieux d'une nouvelle composante.
8. Validation/Exception : cibler une cellule contenant une formule, ajouter deux
   zones. La formule reste identique. Deux preuves sont accessibles au double-clic.
9. Insertion sur cellule ou tableau rempli : refuser l'écrasement et vérifier
   l'absence de changement. Accepter, puis vérifier que les notes personnelles
   antérieures restent présentes. Tester une feuille protégée et des cellules
   fusionnées : message explicite, aucune preuve créée.
10. Matching : plage sans en-têtes avec référence, date, montant. Définir une
    sortie vide. Vérifier une réussite, un montant absent, deux pièces identiques,
    une référence proche mais différente et des critères sur des pages différentes.
    Seule la correspondance exacte unique sur une même page crée les preuves.
11. Relancer après modification d'une entrée : les anciens résultats non confirmés
    sont vidés après confirmation, sans marqueur obsolète. Refuser le remplacement
    conserve toutes les cellules. Une plage de sortie chevauchant l'entrée est refusée.
12. Ouvrir deux classeurs et deux fenêtres d'un même classeur : sélectionner leurs
    pièces, lancer l'OCR, basculer. Aucune écriture ne doit atteindre le mauvais
    classeur. Enregistrer/fermer pendant une opération demande d'attendre ou d'annuler.
13. Revue : Reviewed, Rejected, puis Prepared. Vérifier commentaire et couleur ;
    Prepared remet les informations de relecteur à zéro.
14. Enregistrer, fermer et rouvrir classeur + dossier : vérifier les preuves.
    Faire Enregistrer sous avec volet actif vers un nouveau nom : dossier copié,
    ancien dossier conservé. Un dossier cible existant est signalé sans écrasement.
15. Ouvrir une mission 0.2, réindexer, vérifier les anciens snips et les nouvelles
    positions. Conserver une sauvegarde de recette avant migration.

La compilation GitHub et les smoke tests ne constituent pas l'exécution de cette
recette. Consigner la version d'Excel, son architecture, le commit de l'installateur,
les étapes réussies et les captures des échecs éventuels.
