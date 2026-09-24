# Interface 0.4

Cette version répond aux coupures de texte et à la navigation signalées sur la
capture utilisateur. La démonstration « Introduction - DataSnipper »
(https://www.youtube.com/watch?v=eqFpLXNY26o) a été étudiée via sa transcription
anglaise automatique, notamment 0:30–0:38, 1:21–1:50 et 2:20–2:29.

- Documents dans une liste en haut à droite, noms longs visibles dans la liste ouverte.
- Suppression de la colonne gauche. Recherche pleine largeur, résultats escamotables.
- Hauteurs calculées sur le contenu, retours à la ligne et commandes qui se réorganisent.
- Icônes vectorielles originales dans le ruban et le volet, fond clair et typographie Segoe UI.
- Texte bleu, somme jaune, tableau violet, validation verte, exception rouge ;
  nombre turquoise et date indigo. Couleur cohérente entre outil, cellule et preuve.
- Preuve ouverte par SheetSelectionChange (clic simple ou clavier), sans dialogue.
  Une liste donne accès aux différentes preuves d'une cellule. Les erreurs de cette
  navigation passive s'affichent dans le statut, sans interrompre Excel.
- Zoom manuel conservé ; une même page n'est pas rendue à nouveau à chaque cellule.
  La zone sélectionnée est ramenée dans la vue. Les autres preuves restent visibles.
- Le mode choisi reste actif. Les actions concurrentes sont désactivées pendant l'OCR.

La propriété optionnelle SourceType conserve l'origine Tableau indépendamment du
format Nombre/Date/Texte de sa cellule. Les anciens projets sans cette propriété
utilisent leur type existant ; la lecture seule des preuves ne recolore pas les
anciennes cellules.

## Vérifications

La CI Windows compile le complément VSTO et teste le moteur net48, PDFium et l'OCR
en x86/x64. Elle vérifie aussi la conservation du zoom et du rendu, la disposition
à 420/760 px et avec texte agrandi. Les images du véritable contrôle WinForms sont
publiées dans Doctracker-UI-Previews. Les vérifications COM Excel et le changement
réel de DPI entre écrans restent à exécuter selon VALIDATION_EXCEL.md.

Cette version rapproche les interactions demandées ; elle ne revendique pas une
parité totale avec DataSnipper. Les fonctions Find All Sums, Form Extraction,
les imports Word/Excel et l'export PDF annoté présentés dans la vidéo ne sont pas
implémentés par cette refonte d'interface.
