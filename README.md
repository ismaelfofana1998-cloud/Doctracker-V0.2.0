# Doctracker 0.8.1 — complément Excel Windows

Doctracker importe des PDF, images et fichiers Word (via Word installé), les classe par catégories et intègre les pièces
au classeur Excel, ou les référence dans un emplacement réseau partagé. Il permet
de sélectionner des zones, extraire leur contenu dans Excel, associer plusieurs
preuves à une cellule et revenir à la source pour la revue.

Voir [la stabilisation 0.8.1](docs/STABILISATION_0.8.1.md), [les corrections 0.8 : recherche, dossiers et manipulation directe](docs/CORRECTIONS_0.8.md), [les améliorations 0.7 : espace, recherche et annotations](docs/ERGONOMIE_0.7.md), [les ajouts 0.6 : édition, Word et lecture](docs/EDITION_0.6.md) et [le guide 0.5 : partage, catégories, Xref et récupération](docs/PORTABILITE_0.5.md).

## Installer et utiliser

1. Dans **GitHub Actions → Build Doctracker**, ouvrir l'exécution réussie correspondant
   à la branche `fix/excel-audit-workflows` (ou à sa fusion ultérieure).
2. Télécharger l'artefact **Doctracker-Windows-Installer** et le décompresser entièrement.
3. Fermer Excel. Si une ancienne compilation de test est installée, désinstaller
   uniquement le complément dans les applications Windows (conserver les missions).
   Lancer `Install-Doctracker.cmd` et vérifier le certificat affiché.
4. Ouvrir Excel de bureau Windows, puis enregistrer un classeur local.
5. Onglet **Doctracker → Importer**. Le dernier document importé s'affiche ;
   l'indexation utilise le texte natif des PDF ou l'OCR français/anglais pour les scans.
6. Sélectionner une cellule, choisir **Texte**, **Nombre**, **Date**, **Somme**,
   **Tableau**, **Validation** ou **Exception**, puis dessiner une zone.
7. Cliquer une fois sur une cellule liée pour afficher sa preuve. S'il y en a plusieurs,
   choisir la preuve dans la liste au-dessus du document. Utiliser **Revoir** pour le statut et le commentaire.

Prérequis : Windows, Excel de bureau, .NET Framework 4.8, runtime VSTO,
redistribuables Microsoft Visual C++ 2015–2022 correspondant à l'architecture d'Excel.
L'installateur contient les moteurs PDF/OCR et leurs modèles français et anglais.
La signature actuelle repose sur un certificat de développement temporaire,
que le script d'installation fait vérifier avant de l'approuver pour l'utilisateur.
Un déploiement permanent en cabinet nécessite une signature de production.

## Snips

| Mode | Résultat |
| --- | --- |
| Texte | Texte exact du PDF natif, ou résultat OCR ; jamais exécuté comme une formule |
| Nombre | Un seul montant, négatifs et séparateurs français/anglais pris en charge |
| Date | Date Excel, avec prise en compte des classeurs utilisant le calendrier 1904 |
| Somme | Somme de la zone ; les sélections suivantes s'ajoutent dans la même cellule avec leurs propres preuves |
| Tableau | Aperçu modifiable construit à partir des positions des mots, puis insertion et preuve par cellule |
| Validation / Exception | Preuve colorée ; libellé Validation/Exception dans une cellule vide, valeur ou formule existante conservée |

Texte, Nombre, Date et Tableau avancent vers la ligne suivante après insertion.
Validation, Exception et Somme restent sur la même cellule. Une insertion qui
remplacerait des données demande confirmation. Les commentaires personnels sont
conservés. Si l'écriture ou la sauvegarde des preuves échoue, Doctracker tente de
restaurer la destination et signale explicitement une restauration incomplète.

## Recherche et rapprochement

- **Rechercher la cellule** utilise la cellule sélectionnée ; le champ du volet permet aussi une saisie libre.
- Les résultats partiels restent consultables dans la recherche interactive.
- Pour un rapprochement, sélectionner **sans en-têtes** une plage de 1 à 10 colonnes,
  puis **Définir recherche**. Chaque colonne non vide est un critère obligatoire.
- Sélectionner une cellule de départ ou une plage de même dimension, puis
  **Définir résultat → Lancer le matching**.
- Tous les critères d'une ligne doivent être trouvés sur une même page. Il
  retrouve les fragments alphanumériques malgré les espaces ;
  montants et dates restent stricts. Les résultats partiels demandent confirmation.
  Plusieurs pages ou documents possibles restent ambigus : aucune preuve automatique.
- Les sorties sont les valeurs trouvées dans les pièces, avec un lien individuel.
  Les preuves restent au statut **Prepared**, à revoir.
- Lors d'une relance, le remplacement des résultats existants et l'effacement des
  anciens résultats non confirmés demandent une confirmation globale.
- La réindexation est accessible dans le menu **Documents** du ruban. Une pièce incorrecte sans preuve
  peut être retirée de la liste. Un échec d'indexation bloque le matching et nomme la pièce.

Les opérations longues peuvent être annulées. Il faut attendre leur fin ou les
annuler avant de fermer/enregistrer le classeur. Le complément ne sauvegarde pas
le fichier Excel à votre place : enregistrer le classeur après le travail.

## Conservation des missions

Les pièces et leurs liens sont intégrés au classeur lors de son enregistrement
(format `.xlsx`, `.xlsm` ou `.xlsb`). Il suffit de transmettre ce classeur à un
utilisateur équipé du complément. Aucun dossier adjacent n'est nécessaire.
La limite de sécurité est de 256 Mo de pièces et d'index intégrés par classeur.

Pour les volumes importants, un stockage réseau UNC partagé conserve les sources
une seule fois. Tous les utilisateurs doivent avoir accès au même emplacement ;
le classeur transporte les liens et les preuves. Ce mode ne fusionne pas les
modifications de copies indépendantes du classeur.

Un cache de travail et vingt versions des métadonnées restent sous
`%LOCALAPPDATA%\Doctracker\Recovery`. Le menu **Récupération → Sauvegarder les pièces et liens** exporte
une archive `.dtpack` complète : conservez-la ailleurs pour récupérer les pièces
et liens si le classeur et son cache sont perdus. Les anciens dossiers adjacents
sont migrés à l'ouverture ; conservez-les jusqu'à vérification de la migration.
Utilisez ensuite la version 0.8, les versions antérieures ne comprenant pas le
nouveau stockage. Aucun document n'est envoyé à un service OCR cloud.

## Portée réelle et validation

Cette version corrige les parcours documentaires principaux. **Elle n'est pas
une reproduction intégrale de DataSnipper.** Restent notamment absents : coédition des preuves entre copies du classeur,
connecteur documentaire cloud, Form Extraction par modèle, matching entre plusieurs
groupes de documents et pages, tolérances paramétrables, comparaison de versions,
reconnaissance avancée des tableaux et administration d'entreprise.

L'OCR et la reconstruction de tableaux demandent une revue humaine. Une valeur
modifiée manuellement dans Excel ne modifie pas automatiquement la preuve source.
Le lien conservé doit être revu avant de conclure le contrôle.

- Moteur : tests automatisés sous .NET Framework 4.8 ; cible .NET 8 disponible pour les tests portables.
- Windows : compilation VSTO et tests de rendu, OCR, positions PDF en x86/x64.
- Installation : contrôle des manifestes, signature et bibliothèques livrées.
- **Excel installé** : la recette COM/interaction utilisateur doit être effectuée
  sur un poste Windows avec Excel ; le runner de compilation n'héberge pas Excel.

Voir [la nouvelle interface 0.4](docs/INTERFACE_0.4.md), [les corrections et limites](docs/CORRECTIONS_0.3.md) et
[la recette Excel](docs/VALIDATION_EXCEL.md).

## Développement

Visual Studio 2022 avec charge **Développement Office/SharePoint**, SDK .NET 8,
.NET Framework 4.8. Ouvrir `Doctracker.sln`. Le moteur cible `net48` pour VSTO
et `net8.0` pour les tests portables. Lancer :

```powershell
dotnet test tests/Doctracker.Core.Tests -f net48
scripts\build-release.cmd
```

Sous Linux, seuls les tests moteur sont exécutables :

```sh
dotnet test tests/Doctracker.Core.Tests -f net8.0
```

Doctracker est un produit indépendant utilisant ses propres code et interface.
