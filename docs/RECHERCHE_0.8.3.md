# Recherche et matching — 0.8.3

La recherche libre et le matching des références utilisent le même moteur de caractères. Exemples : `12345` dans `FAC12345A`, `X300` ou `500` dans `500X300Z35`, `500X300Z35` réparti entre de nombreux blocs de texte PDF. Aucun minimum de trois caractères, aucune limite de huit mots. La casse, les accents, les espaces horizontaux, les ligatures et les traits d'union typographiques sont normalisés ; les séparateurs différents (`/` et `-`) ne deviennent pas arbitrairement équivalents pour les références.

Les positions sont associées au texte natif ; les occurrences sans rectangle restent accessibles à la page. Une occurrence positionnée ne masque plus les autres occurrences du texte brut dont les rectangles manquent. Les occurrences qui se chevauchent sont conservées. Les sauts de ligne restent des frontières : deux lignes de tableau ne forment pas une référence. Pas de correction approximative O/0 ou I/1.

La recherche depuis une cellule Excel passe par ce même moteur. La recherche libre retrouve aussi une portion de montant. En rapprochement automatique, les nombres entiers peuvent être des références ; les valeurs explicitement décimales, monétaires ou signées gardent la comparaison financière stricte (ne pas valider 100,00 dans 1100,00). Pour rechercher un fragment de montant, utiliser la recherche libre. Les dates conservent leurs équivalences de format.

Le matching prépare chaque page une fois et réutilise le résultat de chaque critère identique entre les lignes Excel. Il conserve la localisation et la valeur réellement lue. Plusieurs pages/documents candidats restent signalés comme ambigus. Les critères d'une même ligne doivent toujours coexister sur une même page : les joindre arbitrairement entre plusieurs factures regroupées dans un PDF pourrait créer une fausse preuve. Le moteur n'invente pas un résultat en cas d'absence.

## Indexation et sauvegarde

L'indexation extrait le texte natif ou effectue l'OCR, puis écrit un index compressé pour le document. Cette écriture est nécessaire pour éviter de recalculer le texte aux prochaines recherches. Elle conserve aussi les métadonnées du projet et leur version précédente (.bak), par remplacement atomique.

La copie historique supplémentaire dans recovery n'est plus créée pour chaque document indexé ni pour les seuls changements d'erreurs d'indexation. Les modifications de preuves, liens, dossiers et annotations conservent leurs sauvegardes habituelles. Les sources PDF ne sont pas recopiées à chaque indexation. L'intégration dans Excel intervient lors de l'enregistrement du classeur, pas pour chaque page indexée.

Il s'agit d'une réduction ciblée des écritures, pas d'une mesure démontrant que les sauvegardes dominaient le temps d'indexation. Le coût restant dépend de la taille du document, du nombre de blocs PDF, de l'OCR éventuel et du disque. Rien n'est envoyé à Azure.

## Vérification

Tests de références numériques incluses, fragmentation en caractères, positions manquantes/désordonnées, texte absent des blocs, faux rapprochements entre lignes, montants, dates, ambiguïté, index chargé une seule fois, récupération après enregistrement. Scénario de charge reproductible : 1 000 lignes Excel identiques sur une page synthétique de 10 000 références positionnées.

Les tests Windows exercent un PDF natif réel avec références entourées de caractères et glyphes séparés, en x86/x64. Les interactions COM dans Excel installé restent à vérifier sur poste Office ; aucun gain chiffré sur les documents du cabinet n'est annoncé sans les mesurer.
