# Doctracker 0.6 — édition et espace de lecture

- **Supprimer un snip** : clic droit sur sa zone dans le document, ou corbeille près de
  la liste des preuves de la cellule. Tous ses liens dans le classeur sont retirés,
  y compris les copies de cellules. Les autres snips, valeurs et formules sont conservés.
  Un total Somme n'est pas recalculé par cette suppression de preuve : le vérifier.
  Une feuille protégée bloque la suppression avant toute modification. Un échec
  d'écriture déclenche la restauration des commentaires ; toute restauration incomplète est signalée.
- **Xref** : après attribution, retour à la première page et affichage en haut à droite,
  en rouge et en gras, sur PDF et images. Cette annotation est incluse dans l'export.
- **Commenter** : choisir la bulle rouge, dessiner un cadre et saisir le texte.
  Le cadre s'agrandit en hauteur si nécessaire ; un texte dépassant la page est refusé.
  Clic droit sur le commentaire pour le modifier ou le supprimer. Les commentaires sont
  conservés avec le classeur, les sauvegardes et les exports PDF annotés.
- **Word** : `.doc` et `.docx` acceptés dans les imports de fichiers et dossiers.
  Microsoft Word doit être installé et utilisable sur le poste qui importe. Une copie
  PDF est créée en lecture seule, avec macros désactivées et sans mise à jour des liens.
  Le PDF est intégré et indexé comme les autres pièces ; le destinataire n'a pas besoin
  de Word pour lire cette copie. Le fichier Word original reste à son emplacement :
  les sauvegardes Doctracker contiennent sa copie PDF, pas le Word modifiable.
  Les documents protégés/non accessibles sont signalés en erreur. L'annulation attend
  la fin de l'appel Word en cours. La conversion réelle exige une recette sur un poste Word.
- **Dernier import** : sélection automatique du dernier fichier importé avec succès,
  y compris lors d'un import multiple, de dossier ou d'un réimport dédupliqué.
  Le filtre de catégorie est réinitialisé si nécessaire pour rendre la pièce visible.
- **Lecture** : ajustement à la largeur par défaut et défilement vertical des pages longues.
  Le zoom est exprimé par rapport à la taille de page à 96 dpi, pas au bitmap de rendu PDF.
  Clic droit sur « Largeur » pour « Page entière » ou « 100 % » ; Ctrl + molette pour zoomer.
  Le bouton « Lecture » masque recherche, catégories et outils ; « Outils » les réaffiche.
  Les outils passent en icônes avec infobulles dans un volet étroit.

Les annotations sont une couche Doctracker visible immédiatement et intégrée aux
PDF exportés. Les octets des originaux importés restent intacts. Pour voir les ajouts
hors de Doctracker, utiliser l'export PDF annoté ; celui-ci reste une copie rasterisée.

Enregistrer Excel après modification. Schéma 4 : tous les destinataires doivent utiliser
Doctracker 0.6 ; une ancienne version refuse ce schéma pour éviter la perte des commentaires.
Avant remplacement de la version, conserver une sauvegarde complète `.dtpack`.

Recette Excel/Word : créer deux preuves sur une cellule et copier cette cellule ailleurs ;
supprimer une preuve, contrôler les copies, sauvegarder/rouvrir. Ajouter une Xref puis
un commentaire sur un PDF et une image, modifier/supprimer le commentaire, exporter et
rouvrir le PDF. Importer un Word de plusieurs pages et vérifier sa mise en page, ses
caractères accentués, l'extraction et les erreurs de fichiers protégés. Tester le volet
étroit, le mode lecture et la navigation vers un snip en bas d'une page longue.
