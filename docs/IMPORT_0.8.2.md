# Import et panneau compact — 0.8.2

## Pendant un import

Un seul traitement séquentiel calcule les empreintes pour éviter les doublons, copie les PDF/images dans le cache de travail et convertit les Word en PDF via Word installé. Les métadonnées sont enregistrées tous les 25 imports réussis et à la fin. Une annulation conserve les fichiers terminés. Si un checkpoint échoue, les métadonnées reviennent au dernier checkpoint enregistré.

L'import ne lit aucun index, ne lance aucun OCR, ne réindexe aucun ancien document. Le panneau est actualisé une seule fois en sélectionnant directement le dernier fichier importé. Les fichiers sont intégrés au classeur lors de l'enregistrement Excel, selon le mode autonome/partagé choisi.

La première recherche prépare le texte du dossier sélectionné : texte natif pour les PDF qui en disposent, OCR pour les scans. Cette opération est explicite et affiche sa progression ; elle peut prendre du temps sur un gros dossier. Un premier snip sur un PDF non indexé lit uniquement sa page native ; l'OCR de la zone sert de recours. Réindexer reste une action explicite du ruban.

## Interface

- Ligne supérieure du panneau : Doctracker et recherche.
- Ligne suivante : dossier à gauche et document à droite.
- Aucun bandeau de mode « Dessinez une zone ». Le ruban indique le mode actif.
- État sur une ligne en bas, message complet au survol et annulation à côté ; aucun agrandissement du pied de page pour un long message.
- Compteur de pièces dans l'infobulle du sélecteur de documents.

La barre de titre native du volet est gérée par Excel/VSTO et reste distincte du contenu du panneau.

## Vérification

Tests des lots de 61 fichiers (3 sauvegardes), annulation, échecs de fichiers, rollback du checkpoint incluant les doublons, absence de chargement d'anciens index. Essais Windows : deux lignes de contrôles, texte agrandi, aucune perte de surface lors d'un changement de mode, pied de page borné, lecture native avant indexation et parcours PDF/OCR/export existants. Les scénarios COM dans Excel/Word restent à vérifier dans un véritable poste Office.
