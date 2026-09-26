# Corrections 0.3.0

## Défauts trouvés dans 0.2

| Défaut dans le code | Correction |
| --- | --- |
| PictureBox en mode Normal : zoom du cadre, pas de l'image | Image étirée à la surface ; coordonnées normalisées et test de crop |
| SplitContainer initialisé trop petit pour ses limites | Taille initiale définie avant les contraintes |
| Premier document sélectionné pendant la suspension de l'événement | Chargement explicite après liaison de la liste |
| PDF toujours rendu en 1800 × 2400 | Proportions de chaque page préservées, paysage compris |
| TIFF limité à la première page | Navigation et indexation des pages TIFF |
| OCR basé sur BaseDirectory malgré le shadow-copy VSTO | Recherche depuis le CodeBase du complément et chemin natif x86/x64 explicite |
| Nouveau moteur Tesseract pour chaque sélection | Réutilisation contrôlée par verrou et libération du moteur |
| OCR obligatoire même sur les PDF natifs | Extraction native avec positions ; OCR pour les scans |
| Index vidé avant traitement, page partielle considérée prête | Construction isolée, drapeau de fin explicite, annulation et erreurs par pièce |
| Traitements asynchrones utilisant le classeur/la page courants après attente | Classeur et page capturés, contrôles de contexte, exclusivité par classeur |
| Un volet unique pour toutes les fenêtres Excel | Volet par fenêtre et contexte partagé par classeur |
| Enregistrer sous ouvre un dossier vide | Copie du dossier suivi vers le nouveau nom, refus de collision |
| Double-clic normal créant un dossier ou demandant de sauvegarder | Vérification du marqueur avant accès au projet |
| Preuve enregistrée avant l'écriture Excel | Préparation, capture de la destination, écriture, puis commit avec restauration sur erreur |
| Écrasement des cellules et commentaires sans contrôle | Confirmation, refus de fusion/protection/matrice, conservation des notes |
| Texte OCR susceptible d'être interprété comme formule | Insertion texte forcée |
| Date envoyée dans Value2 comme objet DateTime | Numéro de série et calendrier 1904 |
| Table liée uniquement à la cellule supérieure gauche | Aperçu et preuve pour chaque cellule produite |
| Plusieurs nombres concaténés ou premier nombre choisi arbitrairement | Parsing strict et séparation des montants |
| Parenthèses comptables et espace fine ignorés | Signe et groupes de milliers reconnus |
| Référence facture interprétée comme montant ; sous-chaînes numériques | Correspondances exactes, limites de tokens, critères typés |
| Premier candidat accepté malgré ambiguïté | Tous les candidats contrôlés ; lignes ambiguës sans preuve |
| Matching remplaçant des entrées ou conservant d'anciens marqueurs | Interdiction du chevauchement et nettoyage explicite lors des relances |
| Résultat de matching = nom de fichier et score | Valeur source et preuve individuelle par critère |
| Statut Prepared conservant l'identité de l'ancien relecteur | Informations de revue réinitialisées |

## Fonctions ajoutées

- Validation et Exception sans modifier la valeur Excel.
- Plusieurs preuves pour une cellule et choix de la preuve à ouvrir.
- Somme de plusieurs zones successives sur une même cellule.
- Matching exact avec 1 à 10 critères, tous requis sur la même page.
- Annulation des opérations longues, réindexation et retrait des pièces sans preuve.
- Tests Windows de rendu et d'OCR sur les deux architectures.

## Différences connues avec DataSnipper

La parité complète n'est pas atteinte. Le stockage reste adjacent au classeur.
Le matching multicritère se limite à une page commune ; il ne rapproche pas
encore facture, commande et bon de livraison en groupes séparés. Il n'existe pas
d'extraction récurrente par modèle de formulaire, ni de tolérances de matching
configurables, ni de comparaison de versions documentaire.

Le tableau est une reconstruction géométrique avec aperçu ; il ne garantit pas
une reconnaissance correcte de tous les tableaux complexes. La modification
manuelle d'une valeur Excel exige la revue de sa preuve. Les lignes déplacées
ou copiées gardent leur marqueur, mais l'adresse historique du journal est
celle de la création. Le journal est une trace locale, pas une garantie
cryptographique d'inaltérabilité.

## Validation

Au 23 septembre 2026 : **44 tests moteur réussis** sous .NET 8 et sous .NET
Framework 4.8 ; compilation VSTO sans erreur ; smoke tests x86 et x64 réussis
pour affichage/crop, OCR natif français/anglais, PDF paysage et positions des
mots ; packaging ClickOnce et vérification du contenu réussis.

Les tests automatiques vérifient des documents fictifs, les formats de nombres,
les faux positifs, la persistance et le rendu/OCR. Ils ne simulent pas une session
Excel complète. La recette `VALIDATION_EXCEL.md` couvre les opérations COM,
plusieurs classeurs, les fichiers enregistrés puis rouverts, l'installation et
les cas d'écrasement. Aucun résultat non exécuté dans Excel ne doit être présenté
comme validé.
