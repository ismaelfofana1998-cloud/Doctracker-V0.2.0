# 0.9.6 — OCR parallèle et matching texte

- Moteurs OCR simultanés configurables à l'installation et dans le ruban, avec conseil selon le processeur et la RAM.
- Répartition des pages d'un PDF entre plusieurs moteurs, résultats ordonnés et sauvegarde séquentielle.
- Matching automatique : premier résultat texte, sans rejet d'ambiguïté ni interprétation montant/date.
- Détails et limites : `docs/OCR_PARALLELE_0.9.6.md`.

# 0.9.5 — OCR sélectionné et traitement groupé

- Bouton OCR, sélection multiple et traitement par dossier.
- Matching initial sur les documents prêts, extension choisie sans OCR imposé.
- Lots isolés de 16 pages maximum avec moteur réutilisé et rendu hors Excel.
- Édition Excel au double-clic restaurée ; suivi de sélection regroupé.
- Détails et limites : `docs/OCR_SELECTION_0.9.5.md`.

# 0.3.0 — fiabilisation des parcours Excel

## 0.9.3
- Préparation de la zone et OCR des snips dans un processus externe, annulable et limité en durée.
- Erreur contrôlée si le moteur s’arrête ; aucun résultat partiel envoyé aux cellules.
- Journaux distinguant rasterisation, reconnaissance et sortie du processus ; temporaires supprimés après requête.
- Bouton de suppression adapté aux plages ; tests d’arrêt du moteur, timeout, annulation et livraison ClickOnce des exécutables x86/x64.

## 0.9.2
- Aperçus PDF/images proportionnés à leur taille affichée pour limiter la mémoire à faible zoom.
- Découpe OCR depuis la source et depuis la page capturée, indépendamment du zoom et du défilement.
- Protection des événements du lecteur et des fins de snip ; libération des images en cas d’échec.
- Diagnostic local borné et sans contenu documentaire ; tests de découpes répétées et cycles de documents.

## 0.9.1
- Défilement horizontal stable après zoom et changement de page.
- Suppression des snips d’une plage, sans supprimer les liens hors sélection ni les valeurs Excel.
- Relecture des références Excel lors de la navigation/réparation ; rattachement manuel des preuves déliées ; conservation des destinations de récupération et du nom interne de feuille.
- Tests de suppression transactionnelle et du lecteur Windows ; documentation des caches et des limites du tri/coller.

Rendu/zoom et coordonnées corrigés, OCR ClickOnce x86/x64, PDF natif positionnel,
TIFF multipage, annulation, isolation des classeurs, Enregistrer sous, écritures
avec confirmation/restauration, validation/exception, preuves multiples, somme
multi-zone, aperçu des tableaux, matching exact multicritère et tests de régression.
Voir `docs/CORRECTIONS_0.3.md` pour les défauts, les tests et les limites réelles.

# Journal des versions

## 0.2.0 — 31 juillet 2026

- refonte de la surface PDF : zoom non destructif, barres de défilement et déplacement au bouton central ;
- navigation de page et commandes de zoom regroupées dans une barre fixe ;
- modes de snip persistants dans le ruban, avec capture automatique après chaque zone ;
- liste déroulante des pièces placée en haut du volet ;
- indexation OCR automatique au moment de l'import ;
- recherche du contenu dans tous les PDF et images indexés ;
- matching Excel par colonne de recherche et colonne de résultat ;
- résultats de matching écrits ligne par ligne avec preuve et commentaire de traçabilité.

## 0.1.8 — 30 juillet 2026

- correction de la résolution de `pdfium.dll` dans Excel : VSTO charge les
  assemblies managées depuis son cache fantôme `AppData\Local\assembly\dl3`,
  tandis que les bibliothèques natives restent dans le dossier ClickOnce ;
- recherche prioritaire à partir du `CodeBase` d'origine de l'add-in, avec
  solutions de repli contrôlées vers les autres dossiers d'exécution ;
- contrôle du manifeste applicatif ClickOnce pour exiger les références
  `x86\pdfium.dll` et `x64\pdfium.dll`, au-delà de leur simple présence dans
  l'artefact.

## 0.1.7 — 30 juillet 2026

- chargement explicite de `pdfium.dll` depuis le dossier ClickOnce de
  Doctracker au lieu de dépendre du dossier du processus `Excel.exe` ;
- sélection automatique de la bibliothèque x86 ou x64 selon l'architecture
  réelle d'Excel ;
- inclusion obligatoire des deux bibliothèques PDFium dans le manifeste
  ClickOnce et contrôle de leur présence après compilation ;
- message de diagnostic précis en cas de fichier absent, de dépendance Visual
  C++ manquante ou d'architecture incompatible.

## 0.1.6 — 30 juillet 2026

- export dans l'artefact de la partie publique du certificat temporaire ayant
  signé la compilation ;
- ajout de `Install-Doctracker.cmd`, qui vérifie la correspondance entre le
  certificat et `setup.exe` avant d'approuver l'éditeur pour l'utilisateur
  Windows courant ;
- contrôle de l'absence de clé privée et de la cohérence des empreintes dans la
  chaîne GitHub Actions et le build local ;
- correction de l'échec d'installation VSTO indiquant que le certificat ou
  l'emplacement de déploiement n'est pas approuvé.

## 0.1.5 — 30 juillet 2026

- correction de l'échec `FindRibbons` restant : la bibliothèque
  `Microsoft.Office.Tools.Common.v4.0.Utilities.dll`, utilisée par la collection
  de ruban VSTO générée, est désormais copiée localement à côté de l'add-in ;
- ajout d'un contrôle MSBuild exécuté avant `FindRibbons` afin de produire un
  message explicite si cette dépendance n'est pas disponible ;
- mise à niveau des actions de préparation (`checkout`, `setup-dotnet` et
  `setup-msbuild`) vers leurs versions Node.js 24 actuelles ;
- confirmation par le journal GitHub Actions que `Doctracker.AddIn.dll`,
  `Doctracker.Core.dll`, Pdfium et Tesseract étaient déjà compilés et copiés
  correctement avant cette correction.

## 0.1.4 — 30 juillet 2026

- alignement de `Doctracker.Core` et de ses tests sur .NET Framework 4.8 afin
  d'éviter l'échec connu de la tâche VSTO `FindRibbons` avec les dépendances
  .NET Standard 2.0 ;
- exécution des étapes MSBuild VSTO sur un seul nœud, sans réutilisation de
  nœud, pour isoler le chargement des assemblies pendant la génération du
  manifeste ;
- vérification explicite de la production du moteur `net48` avant la
  compilation de l'add-in ;
- conservation des correctifs de signature, de publication et de dépendances
  natives introduits en 0.1.3.

## 0.1.3 — 30 juillet 2026

- correction de l'ordre de signature VSTO : le certificat est créé avant le
  premier build et transmis au build comme à la publication ;
- activation explicite de la signature des manifestes ClickOnce ;
- vérification de la variante .NET Framework du package Tesseract : elle
  contient directement `PixConverter` et évite l'ajout incompatible de
  `Tesseract.Drawing` ;
- génération automatique d'un certificat de développement local réutilisable ;
- vérification du contenu et de la signature de l'installateur avant son
  téléversement ;
- inclusion explicite des DLL natives Tesseract x86/x64 dans le manifeste
  ClickOnce, avec contrôle de leur présence dans le package ;
- message explicite lorsque la charge de travail VSTO est absente ;
- script local unique pour restaurer, tester, compiler et publier ;
- normalisation défensive des anciens fichiers `project.xml` ;
- correction du matching qui pouvait accepter une requête composée uniquement
  de ponctuation ;
- validation des pages de snip, des nombres et des dates avant écriture Excel ;
- sécurisation du chargement de documents et du déchargement de l'add-in.

## 0.1.2

- correction de la référence du contrôleur du volet ;
- qualification explicite de `Microsoft.Office.Tools.CustomTaskPane`.

## 0.1.1

- suppression des ambiguïtés entre le namespace interne Excel et
  `Microsoft.Office.Interop.Excel`.
