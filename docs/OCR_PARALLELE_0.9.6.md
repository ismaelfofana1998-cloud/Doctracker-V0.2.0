# Doctracker 0.9.6 — OCR parallèle local et premier résultat texte

## Installation et réglage

Extraire tout le ZIP puis lancer `Install-Doctracker.cmd`. L'assistant affiche les processeurs logiques, la mémoire physique et une proposition de nombre de moteurs. Entrée conserve le choix existant ou applique la proposition. Le ruban **Doctracker → Réglages OCR** permet de modifier ce nombre pour le prochain traitement, sans réinstaller.

Pour un i5-1235U avec 16 Go, commencer avec **4 moteurs**. La limite proposée est le nombre de processeurs logiques détectés, plafonné à 16. Ce maximum n'est pas une recommandation de performance. Essayer 4 puis 6 sur le même lot avec la réindexation OCR forcée ; conserver le réglage qui finit le plus vite tout en laissant Excel utilisable. Relancer le bouton OCR sur des documents déjà indexés réutilise leur index : ce n'est pas une mesure du temps de reconnaissance.

Le conseil initial utilise le minimum de la moitié des processeurs logiques et d'une limite mémoire : 1 moteur sous 6 Go, 2 sous 12 Go, 4 à partir de 12 Go. Si la mémoire ne peut pas être détectée, la limite indicative est 2. C'est une règle de départ, pas un benchmark de la machine. Elle ne change pas automatiquement le nombre de moteurs pendant une opération.

Le choix est conservé par compte Windows dans `%LOCALAPPDATA%\Doctracker\ocr-workers.txt`. Une valeur illisible revient au conseil matériel ; une valeur hors limite est bornée au matériel actuel.

## Traitement

- Le bouton **OCR**, les fichiers cochés et les dossiers BL/Factures restent disponibles.
- Plusieurs processus OCR travaillent simultanément ; un PDF unique peut lui aussi être réparti entre plusieurs moteurs.
- Chaque tâche traite un petit lot de pages (8 maximum) avec un même moteur Tesseract et un même fichier ouvert. Les tâches suivantes occupent les places libérées dans la fenêtre courante de documents. Le nombre de processus actifs reste borné au choix utilisateur.
- Le rendu des pages a lieu dans les processus enfants. Le texte natif d'un PDF reste prioritaire et les index valides sont réutilisés.
- Les traitements groupés fixent `OMP_THREAD_LIMIT=1` pour limiter la concurrence interne à Tesseract lorsqu'OpenMP est présent. Les snips ponctuels restent indépendants de ce réglage.
- Les résultats sont réassemblés dans l'ordre des pages. Les sauvegardes du projet sont séquentielles, hors des moteurs. Une erreur annule les tâches du document concerné, sans annuler les autres documents.
- Une annulation arrête les processus en cours. Les documents entièrement terminés sont enregistrés ; les documents interrompus conservent leur ancien index. Il faut enregistrer le classeur pour embarquer ses nouveaux index.

Les journaux distinguent préparation, rendu PDF, reconnaissance et sauvegarde, sans enregistrer le texte des documents. Aucune compression avec perte ni baisse de résolution PDF n'est ajoutée. Le gain réel sur les documents utilisateur reste à mesurer ; davantage de moteurs peut augmenter la contention, la mémoire utilisée et la chauffe.

## Matching automatique

La recherche automatique traite chaque critère comme le **texte affiché dans la cellule Excel**. Aucune interprétation comme montant ou date n'est appliquée. Les cellules affichant uniquement `###` doivent être élargies. Les résultats sont également insérés comme texte, ce qui conserve les zéros initiaux et évite la conversion en formule.

Le premier résultat est retenu : documents par date d'ajout, identifiant pour départager, puis pages croissantes et ordre du texte fourni par le PDF/OCR. Une seconde correspondance ne bloque plus une ligne. Les références partielles ne demandent plus une confirmation supplémentaire ; la preuve conserve le texte source et sa localisation pour vérification humaine.

La recherche ignore la casse, les accents et les espaces de mise en page ; virgule et point sont rapprochés. Elle n'invente pas de variantes de date : `26/09/2026` n'est pas transformé en `2026-09-26`. Une recherche `123` peut correspondre à `FA-00123`, sans validation numérique. Les critères non vides d'une même ligne doivent toujours être présents sur la même page. Les lignes trouvées sont insérées même si d'autres lignes n'ont aucun résultat.

Le matching commence par les documents prêts du dossier actif, puis propose la reconnaissance des autres. Après extension du périmètre, il reprend le même ordre de premier résultat. Les snips restent à vérifier et les destinations déjà occupées demandent toujours une confirmation de remplacement.

Cette description remplace les règles de rejet d'ambiguïté et de typage décrites dans les notes des versions précédentes.

## Vérification

Les tests du cœur couvrent les doublons, l'arrêt au premier document, les fragments numériques, les dates littérales, les zéros initiaux, les lignes partiellement trouvées, les coordonnées et le conseil matériel. Les tests Windows x86/x64 vérifient le chevauchement réel de deux processus, le plafond de concurrence, la sélection de documents, l'ordre des pages, les sauvegardes, les erreurs et l'annulation. Ils exercent aussi PDFium/Tesseract sur un PDF synthétique de 30 pages et la persistance du choix d'installation.

La validation des événements Excel (double-clic/F2, édition d'une cellule avec snip, valeurs affichées et enregistrement/réouverture du classeur) nécessite Excel installé ; le runner CI n'en dispose pas.
