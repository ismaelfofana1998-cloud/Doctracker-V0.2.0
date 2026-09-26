# Architecture Doctracker 0.3

- **Core** : modèles XML, parsing des valeurs, rectangles normalisés, extraction
  géométrique des tableaux, recherche/matching et persistance. Cible `net48` pour
  VSTO et `net8.0` pour les tests portables.
- **AddIn** : ruban VSTO, volet WinForms, interactions COM Excel, PDFium et Tesseract.
- **Stockage** : `project.xml` et copies des pièces dans un dossier adjacent au
  classeur. Écriture par fichier temporaire, remplacement et sauvegarde `.bak`.

## Isolation et traitements

Un volet est attaché à chaque fenêtre Excel ; les fenêtres du même classeur
partagent un contexte et un verrou logique d'opération. Toutes les interactions
COM restent sur le thread Excel. L'import, l'indexation et le matching passent sur
un worker ; les mises à jour de statut reviennent par `BeginInvoke`. Le classeur,
la cellule, la page et la zone sont capturés avant l'attente. L'insertion contrôle
à nouveau le classeur actif. Fermeture et sauvegarde sont suspendues pendant
l'opération ; l'utilisateur peut l'annuler.

## Extraction et preuve

Les PDF natifs fournissent texte et boîtes PDFium converties de l'origine bas-gauche
vers des coordonnées normalisées haut-gauche. Les scans fournissent les mêmes
informations via Tesseract. L'index est construit à part puis remplacé après
succès. Les images TIFF sont parcourues par frames.

Les captures utilisent les mots indexés présents dans la zone, puis l'OCR du crop
si nécessaire. Les tableaux passent par une reconstruction des lignes/colonnes et
un aperçu éditable. La preuve conserve la zone réelle de chaque cellule produite.

Le snip est préparé sans persistance. L'add-in capture la destination, écrit la
valeur et le commentaire, puis enregistre la preuve. Si cette séquence échoue, il
restaure la destination et retire les nouveaux objets en mémoire. Ce mécanisme
n'est pas une transaction distribuée résistante à un arrêt brutal d'Excel : le
classeur doit être enregistré par l'utilisateur.

Les commentaires contiennent un ou plusieurs marqueurs `DOCTRACKER-SNIP:`. Ils
permettent le retour à la preuve après réouverture, déplacement ou copie de la
cellule. L'adresse dans le journal reste l'adresse d'origine. Les notes personnelles
sont séparées de la section Doctracker. La Somme garde chaque composante comme
preuve distincte et calcule la valeur cumulée à l'insertion.

## Matching

La recherche interactive peut renvoyer des candidats partiels. Le matching exige
tous les critères non vides sur la même page et un seul candidat. Il conserve les
valeurs trouvées dans la pièce et les zones disponibles. Une page legacy sans
positions produit une preuve de page entière explicitement signalée.

## Limites

Aucun stockage cloud, OCR distant ou envoi de pièces. Les dossiers ne sont pas
embarqués dans le classeur. La revue et le journal sont locaux et modifiables sur
le disque, sans garantie cryptographique. Les états COM, les opérations de revue,
les changements de classeur et l'installation doivent être vérifiés dans Excel
selon `VALIDATION_EXCEL.md`.
