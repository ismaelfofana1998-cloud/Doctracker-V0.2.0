# Doctracker 0.9.2 — fermeture d'Excel pendant les snips

Cette version corrige des risques trouvés dans le code. Sans journal Windows ou fichier déclencheur, elle ne prouve pas la cause du plantage observé et ne garantit pas l'absence de défaillance native.

## Corrections

- Le lecteur rendait chaque page visible à une résolution pouvant atteindre 3 000 pixels de côté. À faible zoom, beaucoup de pages tiennent à l'écran : les images pouvaient donc consommer plusieurs centaines de Mo, indépendamment de leur petite taille à l'affichage. La résolution des aperçus dépend désormais de leur taille affichée. Les images raster affichées sont également bornées.
- La découpe du snip relit les pixels source, indépendamment de l'aperçu. Elle conserve le numéro de page capturé au début, même si un redimensionnement du panneau change la page visible pendant l'attente. La sortie OCR est limitée à 3 000 pixels sur son côté le plus long ; les sources restent intactes.
- Les erreurs gérées du lecteur (souris, défilement, peinture, zoom) sont remontées dans le panneau plutôt que laissées au traitement global d'Excel. Après une erreur de rendu, recharger le document. Les appels différés après destruction du contrôle sont ignorés.
- La fin d'une création/modification de snip est protégée, y compris les erreurs pendant la restauration de l'état Excel et la navigation finale.
- Les bitmaps temporaires sont libérés aussi en cas d'échec de découpe.

PDFium n'est pas thread-safe, mais le wrapper PdfiumViewer utilisé sérialise déjà ses appels natifs au moyen d'un verrou partagé. Aucun verrou bloquant supplémentaire n'a été ajouté sur tout le document. Les moteurs natifs restent chargés dans Excel : une violation d'accès native peut encore fermer Excel. L'isolation dans un processus séparé serait une évolution distincte, non réalisée ici.

## Diagnostic local

Ruban **Doctracker > Récupération > Diagnostic des erreurs**, ou `%LOCALAPPDATA%\Doctracker\Diagnostics`.

Le journal écrit quelques jalons par snip (début, texte natif, OCR, écriture Excel, commit), l'heure UTC, la version au démarrage et l'architecture du processus. En cas d'erreur gérée, il conserve le type, le code et la pile d'appels, sans message d'exception susceptible de contenir les données client. Aucun contenu PDF, texte reconnu, chemin de document ou valeur Excel n'est journalisé. Aucun envoi réseau, aucune copie de document, aucun timer.

Rotation par processus à 256 Kio avec une version précédente. Au démarrage, les anciennes sessions de plus d'un jour sont purgées au-delà des dix fichiers les plus récents. Les sessions récentes sont conservées pour ne pas toucher à d'autres instances Excel actives.

Si Excel se ferme encore : conserver le journal correspondant et relever l'événement de fermeture dans le Moniteur de fiabilité Windows (`perfmon /rel`) : module défaillant, code d'exception, version d'Excel et heure. Une absence de ligne d'erreur après le dernier jalon est possible lors d'un arrêt natif brutal ; le dernier jalon indique la zone à investiguer, pas une cause certaine.

## Validation

Les essais Windows x86/x64 vérifient le budget de bitmaps à 2 % de zoom, quarante découpes successives en résolution source, la découpe d'une page indépendante de la page affichée et huit cycles ouverture/fermeture d'un PDF multipage. Les tests existants de navigation, OCR natif, annotations, export et recherche restent requis. Les scénarios COM Excel réels nécessitent un poste Office.
