# Doctracker 0.9.7

## Recherche depuis une cellule

La barre suit la cellule sélectionnée (clavier ou souris), y compris les cellules sans snip. Une cellule vide efface la recherche précédente. Une saisie manuelle reste conservée tant que la cellule et sa valeur ne changent pas. Le bouton du panneau relit la sélection avant de lancer la recherche, même si le délai de navigation des preuves n’est pas encore écoulé. Un résultat devenu périmé après un changement de cellule ou de texte n’est pas affiché.

## Tableau sur le PDF

Dessiner la zone en mode Tableau ouvre le document source avec une grille proposée et un aperçu des cellules en dessous. La grille se règle **avant insertion dans Excel** :

- Cliquer sur le bord supérieur du cadre pour ajouter une séparation de colonnes ; sur le bord gauche pour une séparation de lignes.
- Utiliser « + Colonne » ou « + Ligne » pour placer une séparation en cliquant dans la zone.
- Déplacer les points violets aux extrémités des séparations ; clic droit sur un point pour supprimer sa séparation.
- Zoomer avec Ctrl + molette et déplacer la vue avec le bouton central de la souris.
- Corriger les valeurs dans l’aperçu, puis « Insérer dans Excel ».

L’aperçu se recalcule au relâchement du point. Modifier la grille ne relance pas l’OCR : chaque mot est affecté à une seule cellule selon son centre. Les corrections manuelles sont conservées pour les cellules dont les limites n’ont pas changé ; les cellules redécoupées sont recalculées depuis le texte source. La grille est limitée à 10 000 cellules, le tri de l’aperçu est désactivé pour conserver le lien entre lignes et preuves.

Chaque valeur insérée possède une preuve correspondant à sa cellule dans la grille. Les cellules vides de la grille vident aussi leur destination, après confirmation si celle-ci contient des données. Annuler ferme l’éditeur sans modifier Excel. Le PDF d’origine reste intact. Après insertion, les snips individuels restent déplaçables/redimensionnables ; la grille complète de préparation n’est pas enregistrée comme un objet rééditable.

## Validation et Anomalie

« Exception » est renommé « Anomalie » dans le ruban et les légendes. L’identifiant interne est conservé pour ouvrir les anciens classeurs.

Ces deux snips lisent désormais le texte natif, l’index ou l’OCR de la zone sélectionnée et l’insèrent comme texte dans Excel, avec leur couleur. Une cellule occupée demande confirmation ; aucun libellé artificiel ne remplace une zone vide. Déplacer ou redimensionner le snip actualise le texte extrait et remet la preuve à vérifier. Une valeur/formule modifiée manuellement demande confirmation avant remplacement.

## Vérification

Tests du cœur : changement de cellule, recherche manuelle, valeurs modifiées, grille, cellules vides, séparations, géométrie et réextraction des annotations. Tests Windows x86/x64 : vrais événements de souris sur le PDF, aperçu, corrections, zoom, extraction du texte natif et capture de l’éditeur.

Les interactions COM avec Excel (navigation depuis une cellule, insertion/confirmation, sauvegarde et réouverture) nécessitent une vérification dans Excel installé ; le runner Windows n’en dispose pas. Le réglage OCR parallèle de 0.9.6 est conservé.
