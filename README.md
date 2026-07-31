# Doctracker V0.2.0

Doctracker est un add-in Excel Windows, local-first, conçu pour formaliser les
preuves d'audit directement dans un classeur. Il importe des PDF et images,
permet de sélectionner une zone, reconnaît son contenu et relie la donnée
extraite à la cellule Excel cible.

## Ce que contient cette première version

- onglet **Doctracker** dans le ruban Excel ;
- volet latéral PDF/image avec liste déroulante des pièces ;
- surface PDF zoomable et réellement déplaçable avec barres de défilement ;
- sélection graphique d'une zone de preuve ;
- OCR français et anglais exécuté localement ;
- OCR automatique dès l'import pour rendre la recherche disponible immédiatement ;
- recherche du contenu dans toutes les pièces indexées depuis une cellule Excel ou le volet ;
- **Text Snip**, **Number Snip**, **Date Snip**, **Sum Snip** et **Table Snip** ;
- mode de snip persistant : choisir un type dans le ruban puis dessiner plusieurs zones à la suite ;
- insertion du résultat dans la cellule active ;
- double-clic sur une cellule liée pour revenir à la pièce et à la zone ;
- commentaires et statuts `Prepared`, `Reviewed`, `Rejected` ;
- piste d'audit horodatée ;
- import local avec détection des doublons par empreinte SHA-256 ;
- Document Matching par colonne de recherche et colonne de résultat, ligne par ligne ;
- tests automatisés du parsing, du matching et de la persistance ;
- compilation Windows et génération d'un installateur ClickOnce.

## Fonctionnement local

Pour un classeur `Mission_Audit.xlsx`, Doctracker crée à côté du classeur :

```text
.Mission_Audit.doctracker/
├── project.xml
└── documents/
```

`project.xml` contient les coordonnées des snips, les cellules cibles, les
statuts et le journal. Les pièces sont copiées dans `documents/`. Aucun document
ou texte OCR n'est envoyé vers un serveur.

Pour déplacer ou archiver une mission, il faut conserver ensemble le classeur
et son dossier `.doctracker`.

## Installer la version compilée

1. Ouvrir l'exécution **Build Doctracker** dans GitHub Actions.
2. Télécharger l'artefact `Doctracker-Windows-Installer`.
3. Décompresser entièrement l'artefact.
4. Fermer Excel puis lancer `Install-Doctracker.cmd`.
5. Vérifier les informations du certificat affichées et répondre `O`.
6. L'assistant approuve le certificat pour l'utilisateur Windows courant et
   lance automatiquement `setup.exe`.
7. Fermer puis rouvrir Excel.

Le certificat produit par la chaîne actuelle est un certificat de
développement temporaire. `Install-Doctracker.cmd` vérifie qu'il correspond
exactement à la signature de `setup.exe`, qu'il ne contient aucune clé privée,
puis l'ajoute uniquement aux magasins de confiance de l'utilisateur courant.
La diffusion commerciale demandera un certificat de signature de code durable.

La chaîne crée le certificat **avant** la compilation VSTO et le transmet à
toutes les étapes MSBuild. Elle vérifie ensuite que `setup.exe`, les manifestes,
les deux langues OCR et les bibliothèques PDF sont réellement présents dans
l'installateur avant de publier l'artefact.

## Développer sous Windows

Prérequis :

- Windows 10 ou 11 ;
- Excel de bureau ;
- Visual Studio 2022 avec la charge **Développement Office/SharePoint** ;
- .NET Framework 4.8 et .NET 8 SDK ;
- redistribuables Microsoft Visual C++ 2015-2022 **x86 et x64** pour le moteur
  OCR natif.

Procédure :

1. ouvrir `Doctracker.sln` ;
2. accepter l'installation des composants proposés par `.vsconfig` ;
3. choisir `Doctracker.AddIn` comme projet de démarrage ;
4. appuyer sur `F5` : le projet télécharge les deux modèles OCR s'ils manquent,
   prépare un certificat de développement local si nécessaire, puis Visual
   Studio lance Excel avec Doctracker.

Pour reproduire localement toute la chaîne de release, y compris les tests et
la génération de l'installateur, lancer `scripts\build-release.cmd`.

La charge Office/SharePoint est la charge Microsoft prévue pour créer des
add-ins VSTO. Voir la
[documentation Visual Studio](https://learn.microsoft.com/visualstudio/install/workload-component-id-vs-community).

## Parcours de test rapide

1. Créer et enregistrer un classeur Excel.
2. Ouvrir l'onglet **Doctracker**, puis **Ajouter des pièces**.
3. Sélectionner une pièce dans le volet.
4. Dessiner une zone sur un montant ou une date.
5. Sélectionner une cellule Excel.
6. Cliquer sur le type de snip correspondant.
7. Choisir un mode de snip dans le ruban, puis dessiner plusieurs zones : chaque zone est traitée automatiquement.
8. Cliquer sur **Rechercher** depuis une cellule contenant une référence ou un montant pour rechercher dans toutes les pièces.
9. Pour le matching, sélectionner la plage d'entrée, cliquer sur **Définir recherche**, sélectionner la plage de sortie, cliquer sur **Définir résultat**, puis cliquer sur **Lancer le matching**.
10. Double-cliquer sur une cellule liée pour revenir à la preuve.
11. Cliquer sur **Revoir** pour renseigner le statut et le commentaire.

Le scénario complet est détaillé dans
[`docs/VALIDATION_EXCEL.md`](docs/VALIDATION_EXCEL.md).

## Limites explicites de la V0.2.0

- la validation finale de l'interface VSTO doit être effectuée dans Excel
  Windows ;
- le Table Snip reconstruit les colonnes à partir des espacements OCR : les
  tableaux complexes ou sans alignement net demanderont une phase
  d'amélioration ;
- le Document Matching crée actuellement une référence vers la page candidate,
  puis exige une validation humaine ; la localisation automatique au mot près
  sera ajoutée avec l'index des boîtes OCR ;
- la licence commerciale, l'administration des abonnements et la signature de
  code de production ne font pas partie de cette version ;
- les fichiers Office protégés ou les PDF chiffrés ne sont pas pris en charge.

## Structure

```text
src/Doctracker.Core/       moteur, stockage, parsing, matching
src/Doctracker.AddIn/      VSTO, ruban, volet, Excel, PDF, OCR
tests/                     tests du moteur indépendant d'Excel
scripts/                   préparation des données OCR et build local
.github/workflows/         compilation et publication Windows
docs/                      architecture et validation
```

Doctracker est un produit indépendant. Le projet reproduit des usages métier
généraux de traçabilité documentaire sans utiliser le code, la marque ou
l'identité graphique d'un logiciel tiers.
