# Doctracker 0.7 — espace de travail et recherche

- Les outils sont dans le ruban Excel. Importer, Xref et Exporter ont le même format. Le menu Documents rassemble dossiers, catégories, référence du test, réindexation et retrait ; Récupération rassemble partage et sauvegardes.
- Le panneau affiche le sélecteur de document encadré à gauche, la recherche à droite, la catégorie et l'état d'indexation. Les deux titres internes et la barre des snips ont été supprimés. La largeur reste le mode de lecture par défaut.
- Validation et Exception gardent leur libellé dans le ruban. Dans Excel, elles renseignent une cellule vide ; une valeur ou une formule existante est préservée.
- Un clic simple sur un snip sélectionne sa cellule Excel. Le marqueur réel est vérifié ; si la cellule a été déplacée, les notes du classeur sont parcourues. Une copie valide à l'adresse d'origine est prioritaire. Une feuille masquée ou un lien absent est signalé.
- Clic droit sur un commentaire : modifier le texte, sa taille (6–72 points de référence), la largeur et la hauteur du cadre (en pourcentage de page). La hauteur minimale est ajustée pour inclure le texte ; si le commentaire dépasse la page, l'éditeur reste ouvert pour corriger. La taille de référence est calculée sur une page de 595 points de large et reste stable au zoom et à l'export.
- La recherche accepte un terme dans un libellé ou une référence plus longue. Les montants restent comparés exactement, y compris leur signe, avec espaces de milliers et virgule/point décimal. Un libellé et son montant regroupés dans un bloc OCR gardent une localisation exploitable. Un bloc absent du texte global n'est plus exclu si ses mots positionnés contiennent la recherche.
- Aucun résultat : message visible, avec vérification de la catégorie. Si la couche texte du PDF est incomplète, Documents > Réindexer le document actif par OCR reconstruit son index. Cette opération peut être lente ; elle est ciblée, annulable, et conserve l'ancien index en cas d'échec. Une reconnaissance OCR erronée ne peut pas être réparée par une recherche approximative sans risque pour les montants.
- La référence de test est demandée une fois par classeur. Les Xref suivantes réutilisent ce préfixe et proposent uniquement les numéros disponibles. Documents > Référence du test permet de changer le préfixe pour les prochaines attributions ; les références déjà posées ne sont pas réécrites silencieusement. Les anciens dossiers avec une seule référence sont repris automatiquement.
- L'export porte `DAC B 30 040 - Feuil!A1` auprès du snip et conserve `DAC B 30 040 / 02` en haut à droite du document. Les originaux restent intacts.

Le schéma 5 conserve le préfixe commun et la taille des commentaires. Les dossiers antérieurs sont migrés ; les versions précédentes doivent être remplacées avant de rouvrir un classeur enregistré par cette version.

## Vérification

Les tests couvrent notamment `BFA`, `X300`, `6489,83`, `6 489,83`, les montants entourés d'astérisques et le refus des valeurs voisines ou négatives, ainsi que la persistance du cadre, du texte et de la référence commune. Le banc Windows rend le panneau à 420/760 pixels et avec un texte agrandi, vérifie l'espace de lecture, le message sans résultat, PDFium, OCR et les exports en x86/x64.

La navigation COM, les cellules protégées/formules et l'import Word restent à vérifier dans Excel/Word installés sur Windows. Les captures utilisateur ne sont pas des PDF source ; elles ne permettent pas de vérifier leur couche texte réelle. Aucun document client n'est inclus dans les tests ni dans le dépôt.
