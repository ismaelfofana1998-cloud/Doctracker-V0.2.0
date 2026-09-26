# Recette Excel Windows 0.4

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
   puis cliquer une fois sur plusieurs cellules (pas seulement la première).
   Annuler l'aperçu : aucune cellule/preuve créée.
7. Somme : sélectionner une cellule vide, dessiner successivement `10,00`,
   `5,25`, `1,25`. Attendu : `16,50` dans la même cellule avec trois preuves.
   Cliquer une fois et choisir chacune dans la liste de preuves. Une modification manuelle de la somme
   doit empêcher l'ajout silencieux d'une nouvelle composante.
8. Validation/Exception : cibler une cellule contenant une formule, ajouter deux
   zones. La formule reste identique. Deux preuves sont accessibles dans la liste de preuves.
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
13. Revue : Reviewed, Rejected, puis Prepared. Vérifier commentaire et statut ; la couleur du type de snip est conservée ;
    Prepared remet les informations de relecteur à zéro.
14. Enregistrer, fermer et rouvrir classeur + dossier : vérifier les preuves.
    Faire Enregistrer sous avec volet actif vers un nouveau nom : dossier copié,
    ancien dossier conservé. Un dossier cible existant est signalé sans écrasement.
15. Ouvrir une mission 0.2, réindexer, vérifier les anciens snips et les nouvelles
    positions. Conserver une sauvegarde de recette avant migration.

16. À 100 %, 150 % et 200 % dans Windows, vérifier un volet étroit et large : titre,
    recherche, mode et statut lisibles, documents en haut à droite, aucune colonne gauche.
    Les résultats sont sous la recherche ; Masquer / Échap redonne la place au document.
17. Passer d'une cellule liée à une autre à la souris et au clavier. La bonne pièce,
    la bonne page et la zone apparaissent sans dialogue ni déplacement de la sélection
    Excel. À zoom manuel, le zoom reste inchangé et la preuve est ramenée dans la vue.
18. Vérifier sept couleurs distinctes : texte bleu, nombre turquoise, date indigo,
    somme jaune, tableau violet, validation verte, exception rouge. Les cellules d'un
    tableau restent violettes même quand leurs valeurs sont des nombres ou dates.

La compilation GitHub et les smoke tests ne constituent pas l'exécution de cette
recette. Consigner la version d'Excel, son architecture, le commit de l'installateur,
les étapes réussies et les captures des échecs éventuels.

## Recette 0.8

- Fermer et rouvrir un classeur contenant des PDF scannés et natifs. Depuis une cellule contenant `BFA`, `X300` ou `6489,83`, utiliser **Rechercher la cellule**. Contrôler chaque occurrence, y compris les répétitions sur une page. Tester aussi la saisie libre, une cellule vide, un dossier vide, et la limite de 200 résultats.
- Ouvrir **Documents > Classer les documents**. Créer `BL`, `Factures`, puis un client et ses sous-dossiers. Déplacer plusieurs pièces par glisser-déposer et avec **Déplacer vers**. Renommer/supprimer un dossier puis rouvrir le classeur : les documents et les preuves doivent rester disponibles.
- Cliquer sur un commentaire, déplacer son cadre puis ses huit poignées à différents zooms. Tester Échap pendant un glissement, les bords de page, un cadre trop petit (ancien cadre conservé), puis double-cliquer pour changer texte/police. Sauvegarder, rouvrir et exporter.
- Créer puis modifier une Xref (préfixe et numéro). Vérifier première page, nom affiché, export et notes Excel. Refuser le numéro d'une autre pièce, autoriser le retour à son propre ancien numéro. Tester une feuille protégée : aucune modification partielle ne doit subsister.
