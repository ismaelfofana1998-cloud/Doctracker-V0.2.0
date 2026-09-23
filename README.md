# Doctracker 0.3.0 — complément Excel Windows

Doctracker importe des PDF et images dans un dossier de mission local. Il permet
de sélectionner des zones, extraire leur contenu dans Excel, associer plusieurs
preuves à une cellule et revenir à la source pour la revue.

## Installer et utiliser

1. Dans **GitHub Actions → Build Doctracker**, ouvrir l'exécution réussie correspondant
   à la branche `fix/excel-audit-workflows` (ou à sa fusion ultérieure).
2. Télécharger l'artefact **Doctracker-Windows-Installer** et le décompresser entièrement.
3. Fermer Excel. Lancer `Install-Doctracker.cmd` et vérifier le certificat affiché.
4. Ouvrir Excel de bureau Windows, puis enregistrer un classeur local.
5. Onglet **Doctracker → Ajouter des pièces**. Le premier document s'affiche ;
   l'indexation utilise le texte natif des PDF ou l'OCR français/anglais pour les scans.
6. Sélectionner une cellule, choisir **Texte**, **Nombre**, **Date**, **Somme**,
   **Tableau**, **Validation** ou **Exception**, puis dessiner une zone.
7. Double-cliquer sur une cellule liée pour revenir à sa preuve. S'il y en a plusieurs,
   choisir la preuve dans la liste. Utiliser **Revoir** pour le statut et le commentaire.

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
| Validation / Exception | Ajout d'une preuve colorée sans remplacer la valeur de la cellule |

Texte, Nombre, Date et Tableau avancent vers la ligne suivante après insertion.
Validation, Exception et Somme restent sur la même cellule. Une insertion qui
remplacerait des données demande confirmation. Les commentaires personnels sont
conservés. Si l'écriture ou la sauvegarde des preuves échoue, Doctracker tente de
restaurer la destination et signale explicitement une restauration incomplète.

## Recherche et rapprochement

- **Rechercher** utilise la cellule sélectionnée ; le champ du volet permet aussi une saisie libre.
- Les résultats partiels restent consultables dans la recherche interactive.
- Pour un rapprochement, sélectionner **sans en-têtes** une plage de 1 à 10 colonnes,
  puis **Définir recherche**. Chaque colonne non vide est un critère obligatoire.
- Sélectionner une cellule de départ ou une plage de même dimension, puis
  **Définir résultat → Lancer le matching**.
- Tous les critères d'une ligne doivent être trouvés exactement sur une même page.
  Plusieurs pages ou documents possibles restent ambigus : aucune preuve automatique.
- Les sorties sont les valeurs trouvées dans les pièces, avec un lien individuel.
  Les preuves restent au statut **Prepared**, à revoir.
- Lors d'une relance, le remplacement des résultats existants et l'effacement des
  anciens résultats non confirmés demandent une confirmation globale.
- La réindexation est accessible dans le volet. Une pièce incorrecte sans preuve
  peut être retirée de la liste. Un échec d'indexation bloque le matching et nomme la pièce.

Les opérations longues peuvent être annulées. Il faut attendre leur fin ou les
annuler avant de fermer/enregistrer le classeur. Le complément ne sauvegarde pas
le fichier Excel à votre place : enregistrer le classeur après le travail.

## Conservation des missions

Pour `Mission_Audit.xlsx`, le dossier adjacent est :

```text
.Mission_Audit.doctracker/
├── project.xml
├── project.xml.bak
└── documents/
```

Les documents, coordonnées, statuts et événements restent sur l'ordinateur.
Aucun document n'est transmis à un service OCR ou à une base cloud.
**Conserver le classeur et son dossier ensemble.** Les preuves ne sont pas
embarquées dans le fichier `.xlsx`. Un « Enregistrer sous » effectué pendant que
le complément suit ce classeur copie son dossier, sans supprimer l'original.
Un dossier de destination existant n'est pas écrasé.

Les projets 0.2 sont lisibles et migrent vers le schéma 2 à l'ouverture.
La sauvegarde précédente reste dans `project.xml.bak`. Après migration,
utiliser 0.3 pour éviter de perdre les positions de mots ou les nouveaux types
de preuves en rouvrant avec 0.2. Cliquer **Réindexer les pièces** pour bénéficier
des positions dans les anciens projets.

## Portée réelle et validation

Cette version corrige les parcours documentaires principaux. **Elle n'est pas
une reproduction intégrale de DataSnipper.** Restent notamment absents : pièces
embarquées dans le classeur, Form Extraction par modèle, matching entre plusieurs
groupes de documents et pages, tolérances paramétrables, comparaison de versions,
reconnaissance avancée des tableaux et administration d'entreprise.

L'OCR et la reconstruction de tableaux demandent une revue humaine. Une valeur
modifiée manuellement dans Excel ne modifie pas automatiquement la preuve source.
Le lien conservé doit être revu avant de conclure le contrôle.

- Moteur : tests automatisés sous .NET 8 et .NET Framework 4.8.
- Windows : compilation VSTO et tests de rendu, OCR, positions PDF en x86/x64.
- Installation : contrôle des manifestes, signature et bibliothèques livrées.
- **Excel installé** : la recette COM/interaction utilisateur doit être effectuée
  sur un poste Windows avec Excel ; le runner de compilation n'héberge pas Excel.

Voir [les corrections et limites](docs/CORRECTIONS_0.3.md) et
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
