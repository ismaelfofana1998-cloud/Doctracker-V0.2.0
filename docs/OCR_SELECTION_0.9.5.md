# Doctracker 0.9.5 — OCR choisi et groupé

## Utilisation

- **OCR** dans le ruban ouvre une liste à cocher. Le filtre dossier inclut ses sous-dossiers. « Tout le dossier » coche les documents affichés ; changer de dossier conserve les coches, dont le total reste visible. « Lancer l’OCR » traite uniquement les fichiers cochés.
- Dans **Documents → Classer les documents**, sélectionner des fichiers avec Ctrl/Maj puis utiliser **OCR de la sélection**.
- Le texte natif des PDF est utilisé en priorité. Les index complets et valides sont réutilisés. La réindexation OCR forcée du document actif reste accessible pour corriger une reconnaissance existante.
- Le matching travaille d’abord sur les documents déjà préparés du dossier actif. Il affiche le nombre de lignes rapprochées avant de proposer la préparation des autres fichiers. **Non** conserve la recherche limitée ; **Oui** permet de choisir les fichiers supplémentaires ; **Annuler** interrompt sans insérer de résultats.
- Après préparation supplémentaire, le matching réévalue tous les documents prêts du périmètre : un second candidat devient une ambiguïté à revoir. La barre d’état indique le nombre de documents exclus.
- Les index terminés sont conservés dans le cache de travail puis avec le classeur lors de son enregistrement. Une annulation conserve l’index précédent du document en cours et les documents déjà terminés.

## Optimisation

Avant cette version, chaque page reconnaissable entraînait un lancement de processus et une initialisation Tesseract, avec rendu côté hôte puis encodage/décodage PNG. Désormais, les pages manquantes d’un même document sont regroupées par lots de 16 au maximum. Chaque lot ouvre une seule fois le fichier et réutilise un moteur Tesseract ; les images sont rendues et libérées une page à la fois **dans le processus enfant**. La résolution OCR n’est pas abaissée pour les PDF.

La limite de 16 borne la taille des réponses et renouvelle périodiquement le moteur natif. Un délai de 90 secondes sans page terminée, l’annulation et la surveillance du processus parent restent actifs. Les pages traitées font progresser l’affichage. Le journal enregistre les durées de lots et les numéros de pages, sans contenu reconnu.

Ce changement supprime des coûts fixes. Le gain dépend du nombre de pages, du texte et des scans : aucune accélération chiffrée sur les documents utilisateur n’est encore établie. Le traitement reste séquentiel pour limiter la mémoire, notamment avec Excel 32 bits.

## Édition Excel

Le double-clic n’est plus détourné pour ouvrir une preuve : il conserve son action Excel d’édition. Le bouton **Ouvrir la preuve** reste disponible. Le suivi de la sélection est différé de 220 ms, regroupe les événements rapides et vérifie qu’Excel est prêt. Une preuve déjà affichée n’est pas rechargée.

Le journal `session-18768.log` (0.9.4) montre des snips terminés mais aucune exception explicite lors de l’édition d’une cellule. Ces corrections ciblent les comportements identifiés dans le code ; la disparition du symptôme exact nécessite une vérification dans Excel installé.

## Validation à exécuter

- Tests du cœur : périmètre BL/Factures, sous-dossiers, sélection explicite et ambiguïté après extension.
- Windows x86/x64 : reconnaissance de pages non consécutives, équivalence texte et positions, réutilisation d’un index, échecs/annulation/délai du moteur, sélection de fichiers et capture du dialogue.
- Excel installé : double-clic/F2 et saisie dans une cellule avec snip, changement rapide de sélection, matching avec Oui/Non/Annuler, sauvegarde/réouverture du classeur. L’environnement CI ne dispose pas d’Excel.

## Résultats locaux de cette préparation

- Suite .NET 8 sous Linux : **163 succès / 165 tests**. Les cinq nouveaux tests passent.
- Deux tests d'injection d'échec par `FileShare.None` échouent parce que le remplacement du fichier ne déclenche pas l'exception attendue. Les deux mêmes échecs ont été reproduits sur le commit 0.9.4 `fea4846`, dans un checkout séparé, sans modifier les tests.
- Analyse syntaxique Roslyn : 64 fichiers C#, aucune erreur de syntaxe. Cette analyse ne remplace pas la compilation VSTO ni les essais Windows.
- Après autorisation explicite, la version 0.9.5 a été envoyée sur `fix/excel-audit-workflows`. Le run Windows `36196035654` a réussi : 165 tests du cœur, tests x86/x64 et installateur produit.
- La version 0.9.6 ajoute le parallélisme et remplace le rejet d'ambiguïté par le premier résultat texte ; voir `OCR_PARALLELE_0.9.6.md`.
