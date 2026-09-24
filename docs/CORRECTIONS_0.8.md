# Doctracker 0.8 — recherche et dossiers

## Recherche

La recherche interactive retrouve les occurrences dans les documents du dossier affiché, sans case de mode. Les caractères autour du terme ne bloquent pas la recherche : `BFA` dans `Facture de BFA`, `X300` dans `500X300Z35`, ou une suite de chiffres dans une référence. Espaces, casse, accents et virgule/point décimal sont normalisés. Chaque occurrence localisée est proposée, y compris les répétitions sur une page. La liste affiche au plus 200 résultats et signale la limite.

Depuis Excel : sélectionner la cellule puis **Doctracker > Rechercher la cellule**. Il n'est pas nécessaire de retaper le contenu. Le champ du panneau reste utilisable ; s'il est vide, sa loupe recherche la cellule active. Le dossier affiché limite la portée ; choisir **Tous les documents** pour élargir.

Le matching automatique est séparé de cette recherche : il garde les contrôles exacts des montants/dates et la détection des ambiguïtés. Un résultat de recherche ne vaut pas validation automatique.

## Réparation XML

Certains caractères de contrôle extraits des PDF peuvent être sérialisés sous une forme interdite par XML 1.0, puis provoquer une erreur à la relecture. L'écriture filtre ces caractères sans toucher au texte imprimable ; le lecteur reprend les anciennes références numériques et nettoie les index. Les DTD et les entités externes restent interdites.

Les index tronqués/manquants sont détectés avant recherche et reconstruits à partir des pièces originales, en conservant les preuves. Une pièce non réindexable est signalée ; la recherche continue sur les autres documents. Les index vérifiés sont mémorisés tant que leur taille/date ne changent pas. La capture d'erreur seule ne prouve pas que cette cause est celle du fichier utilisateur : les erreurs détaillées sont désormais affichées si la récupération ne suffit pas.

## Dossiers de travail

**Documents > Classer les documents** ouvre le gestionnaire : créer, renommer ou supprimer un dossier et ses sous-dossiers, sélectionner plusieurs documents (Ctrl/Maj), puis les glisser vers un dossier. **Déplacer vers…** offre une alternative sans glisser-déposer. Double-clic sur une pièce pour l'ouvrir. Les modifications sont enregistrées immédiatement dans le projet ; enregistrer Excel pour les transmettre.

Les dossiers sont logiques, intégrés au classeur. Aucun déplacement du fichier source ni changement d'identifiant de preuve. Les anciennes catégories sont reprises. Déplacer un document remplace son ancien classement ; supprimer un dossier remonte ses documents au parent, ou dans **Non classés**, sans supprimer les pièces.

## Commentaires et Xref

Cliquer sur un commentaire pour afficher ses huit poignées. Glisser le cadre pour le déplacer, une poignée pour le redimensionner. Échap annule le geste. Un cadre qui couperait le texte est refusé et la géométrie précédente est conservée. Double-clic pour modifier le texte et la police ; clic droit pour modifier/supprimer. Les mêmes positions sont conservées au zoom, à la sauvegarde et à l'export.

**Créer / modifier Xref** permet de modifier le préfixe et le numéro de la pièce active. La case de réutilisation définit le préfixe des prochaines attributions. Les liens de cellules et les snips conservent leurs identifiants ; leurs notes Excel sont actualisées. Les anciens numéros restent réservés au document et peuvent lui être réattribués, jamais à une autre pièce. Les Xref existantes d'autres documents ne sont pas réécrites.

Schéma 6 : les destinataires doivent utiliser la version 0.8 pour conserver les dossiers vides. Sauvegarder les missions avant remplacement du complément. La recette Excel/Word sur un poste Office reste nécessaire ; le runner Windows ne contient pas ces applications.
