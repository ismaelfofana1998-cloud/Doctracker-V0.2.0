# Doctracker 0.9.3 — OCR isolé après analyse des journaux

## Ce que montrent les traces reçues

Le journal de session 22336 s'arrête le 25 septembre 2026 à 15:43:45 UTC (17:43:45 à Paris), sur `SnipOcr`, pour un snip Nombre, page 2, sous Excel x64 et Doctracker 0.9.2. La lecture du texte natif a rendu la main. Il n'y a ensuite ni jalon d'écriture Excel ni arrêt normal. Ce jalon précédait **la rasterisation/découpe et l'OCR** : il ne permet pas de désigner précisément le module natif fautif, ni de dater la fermeture à la milliseconde.

Les erreurs antérieures de conversion de nombres étaient interceptées et suivies de snips réussis. La session 14912 comporte un arrêt normal ; la session 9536 ne contient qu'un démarrage. Aucune trace brute fournie par l'utilisateur n'est ajoutée au dépôt.

## Changement

Le découpage et la reconnaissance d'un snip sans texte exploitable s'exécutent maintenant dans un processus enfant x86/x64. La source est lue en lecture seule ; Excel ne reçoit que le résultat reconnu. Le même moteur isolé est utilisé pour l'OCR de l'indexation. La préparation PDF de l'indexation, la recherche native et le lecteur PDF restent dans l'add-in.

En cas d'arrêt du moteur, de PDF invalide, d'annulation ou de dépassement du délai (90 secondes par reconnaissance), aucun résultat partiel n'est inséré. Le processus est arrêté et les fichiers temporaires de la requête sont supprimés. Un nouvel essai démarre un moteur neuf. Le moteur enfant a également une limite de vie de deux minutes et s'arrête si son processus parent se termine.

Le lancement d'un moteur neuf par reconnaissance coûte un peu de temps ; c'est un choix de stabilité. Aucun service cloud, aucune connexion réseau, aucun droit administrateur supplémentaire. Si une politique du poste bloque le lancement de l'exécutable OCR, une erreur est affichée : aucun retour automatique au moteur dans Excel.

Les temporaires de requête/résultat sont sous `%LOCALAPPDATA%\Doctracker\OcrWork`. Ils contiennent transitoirement les paramètres et le texte reconnu (ou les pixels pour l'indexation d'une image), sont locaux au compte Windows et supprimés après l'opération. Un arrêt brutal d'Excel peut laisser des résidus ; après fermeture de toutes ses instances, ce dossier peut être vidé. Les originaux et le classeur ne sont jamais modifiés par le moteur.

Les journaux distinguent désormais : préparation de la zone, zone prête, reconnaissance démarrée, reconnaissance terminée, résultat prêt et code de sortie du moteur. Aucun contenu reconnu n'est écrit dans les journaux.

Une reconnaissance sans texte retourne immédiatement un résultat vide, sans parcourir les positions de mots dans le moteur natif. Le format Nombre continue de refuser une extraction non numérique plutôt que d’inventer une valeur.

Le bouton de suppression reconnaît une sélection de plusieurs cellules et passe à la suppression des snips de cette plage. Le clic droit sur une preuve conserve la suppression de cette preuve précise.

## Validation et limites

Tests Windows x86/x64 : vraie reconnaissance de zone image/PDF dans un processus séparé, récupération après source invalide, arrêt inattendu d'un enfant, arrêt d'un enfant bloqué, annulation, absence de processus et fichiers temporaires résiduels. Les vérifications de l'installateur contrôlent les deux exécutables et leur présence au manifeste ClickOnce.

Cette isolation contient les défaillances du processus OCR ; elle ne prétend pas protéger Excel de tout plantage provenant du lecteur PDF, d'un autre complément ou d'Excel lui-même. Sans rapport de fiabilité Windows, le module responsable de la fermeture initiale n'est pas prouvé. Les écritures COM dans Excel restent à valider sur un poste Office.
