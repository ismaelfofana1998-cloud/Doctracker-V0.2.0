# Sécurité et confidentialité

## Garanties de cette version

- toutes les pièces restent dans le dossier local de la mission ;
- l'OCR Tesseract s'exécute dans le processus local ;
- le moteur de matching lit uniquement l'index OCR local ;
- aucune télémétrie n'est activée ;
- aucune clé privée, clé API ou chaîne de connexion n'est attendue ;
- la copie locale est nommée avec un identifiant aléatoire, pas avec une donnée
  client ;
- les chemins relatifs sont contrôlés avant toute ouverture de pièce ;
- l'écriture de `project.xml` est atomique et conserve une sauvegarde
  `project.xml.bak`.

## Risques à traiter avant production

- signer l'installateur avec un certificat de code-signing durable ;
- ajouter un contrôle d'intégrité au chargement du projet ;
- chiffrer le dossier local lorsque la politique du cabinet l'exige ;
- définir une politique de purge et d'archivage ;
- journaliser les modifications de valeurs Excel après création d'un snip ;
- réaliser une revue COM/VSTO et un test antivirus sur le package final ;
- sécuriser la future licence sans collecter de contenu de mission.
