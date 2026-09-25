# Doctracker 0.9 — lecture continue et reconnaissance à la demande

## Utilisation

- Texte est le premier outil du ruban et le mode actif dès l'ouverture du panneau. Désactiver un autre outil revient au Texte.
- Défilement vertical continu des PDF ; le champ de numéro de page permet l'accès direct. Les flèches restent des raccourcis facultatifs. Seules les pages visibles et la page en cours d'édition sont rendues, puis les images hors écran sont libérées.
- Cliquer un snip pour le sélectionner, tirer son cadre pour le déplacer ou ses huit poignées pour le redimensionner. Échap annule le geste. Alt permet de dessiner une nouvelle sélection à l'intérieur d'un snip existant. Les commentaires conservent leur propre déplacement/redimensionnement.
- Une modification de snip relit la zone et met à jour les cellules liées. Les sommes de plusieurs snips sont recalculées ensemble. Une formule ou une valeur modifiée manuellement nécessite confirmation avant remplacement. Les snips Validation/Exception changent de zone sans écraser la valeur contrôlée. La modification remet la preuve au statut à revoir.

## Reconnaissance et matching

Choisir le dossier dans le sélecteur supérieur avant de lancer la recherche ou le matching. Le sous-dossier permet de travailler seulement sur une classe de documents (BL, factures, etc.). Le moteur réutilise les index existants.

Pour les pièces restantes, le texte natif est lu sans OCR. Les pages natives d'un PDF mixte sont réutilisées et seules les pages scannées passent en reconnaissance. Une confirmation indique le nombre de documents restant à reconnaître : « Reconnaître le texte de N document(s) du dossier sélectionné ? Cela peut prendre du temps. » Refuser annule cette recherche, sans lancer l'OCR. Les lectures natives déjà effectuées restent dans le cache.

Le snip ponctuel lit d'abord le texte indexé/natif et, si nécessaire, reconnaît uniquement la zone sélectionnée. Une barre animée et le statut indiquent l'opération en cours. Annuler stoppe avant la prochaine page ou avant insertion ; un appel natif OCR déjà en cours termine avant de rendre la main.

## Liens sans notes visibles

Les nouvelles preuves utilisent des noms masqués du classeur, sans note de cellule. Les anciennes notes Doctracker sont migrées au chargement ; les notes personnelles antérieures sont conservées. Une feuille protégée peut conserver sa note jusqu'à déprotection, car le complément ne contourne pas sa protection.

Les noms sont enregistrés avec Excel et suivent les insertions/déplacements de cellules gérés par Excel. L'index de navigation est invalidé lors des changements de feuille, plutôt que parcouru à chaque sélection. Le tri et le copier-coller de valeurs ne constituent pas une duplication de preuve : les noms sont des références de cellules. Vérifier les liens après réorganisation et utiliser Réparer les liens si nécessaire. Les scénarios de déplacement, tri et copier-coller demandent une validation sur Excel installé.

## Où sont les fichiers ?

- Les documents et index nécessaires à l'affichage/reconnaissance sont matérialisés dans `%LOCALAPPDATA%\Doctracker\Cache\<session>`. C'est un cache de travail, pas une nouvelle archive de chaque PDF.
- Les changements de liens, snips, Xref et commentaires restent en mémoire jusqu'à l'enregistrement du classeur. Ils sont alors intégrés au fichier Excel (ou à son stockage partagé configuré). Aucun historique XML supplémentaire n'est écrit pour chaque opération. Les nouveaux index compressés restent écrits dans le cache pour libérer la mémoire et éviter une nouvelle reconnaissance.
- Les caches de cette session sont supprimés à la fermeture normale, si aucun fichier n'est verrouillé. Les anciens dossiers `%LOCALAPPDATA%\Doctracker\Recovery` ne sont pas effacés automatiquement.
- L'archive `.dtpack` reste disponible sur demande et contient **documents + index + liens + annotations**. L'export PDF est une restitution des documents annotés ; ce n'est pas une sauvegarde des liens Excel.

Enregistrer le classeur pour conserver les changements : en cas de fermeture sans enregistrement ou de crash, les modifications depuis le dernier enregistrement peuvent être perdues. Les originaux ne sont pas modifiés ; écrire directement dans le PDF ne conserverait ni les cellules Excel liées ni les données nécessaires au recalcul des snips.

## Validation

Tests du moteur : persistance différée, index réutilisable, un seul changement de lot importé, nouvelle tentative après erreur d'enregistrement, recalcul texte/nombre/somme, retour arrière de géométrie/valeur/statut.

Essais Windows : chargement PDF natif et OCR, scans en attente sans OCR puis reconnaissance autorisée, PDF synthétique de 30 pages, accès direct, défilement vers la page 21, cache borné, activation/annulation d'une édition graphique. La compilation vérifie les interfaces Office ; le runner n'a pas Excel. La migration des notes, les noms masqués, la navigation Excel et la mise à jour effective des cellules doivent donc être vérifiées sur un poste Office.
