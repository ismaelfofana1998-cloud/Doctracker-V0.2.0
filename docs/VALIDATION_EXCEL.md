# Validation Excel Windows

Cette recette est volontairement courte. Elle sert à distinguer un défaut de
chargement VSTO d'un défaut fonctionnel Doctracker.

## 1. Chargement

- fermer Excel ;
- décompresser entièrement l'artefact d'installation ;
- lancer `Install-Doctracker.cmd`, vérifier l'empreinte affichée et confirmer ;
- laisser le script lancer `setup.exe` ;
- ou lancer le projet avec `F5` depuis Visual Studio ;
- ouvrir Excel de bureau ;
- vérifier la présence de l'onglet **Doctracker** ;
- cliquer sur **Ouvrir Doctracker** ;
- confirmer que le volet se fixe à droite.

Résultat attendu : le volet affiche l'en-tête orange, la liste des pièces et la
zone de visualisation.

## 2. Création du dossier local

- créer un nouveau classeur ;
- l'enregistrer sous `Test_Doctracker.xlsx` ;
- ouvrir Doctracker.

Résultat attendu : un dossier `.Test_Doctracker.doctracker` apparaît à côté du
classeur avec `project.xml` et `documents`.

## 3. PDF et image

- importer un PDF natif ou scanné ;
- importer une image JPG ou PNG ;
- changer de pièce ;
- naviguer entre les pages du PDF ;
- utiliser le zoom sans perdre la navigation horizontale ou verticale ;
- déplacer la surface zoomée avec les barres de défilement ou le bouton central de la souris ;
- sélectionner la pièce dans la liste déroulante située en haut du volet.

Résultat attendu : chaque pièce s'affiche, sans figer Excel.

## 4. Snips

Choisir un mode dans le ruban, sélectionner une cellule vide, puis dessiner une
zone. L'extraction doit être écrite automatiquement. Dessiner ensuite plusieurs
zones supplémentaires sans recliquer sur le mode, puis le désactiver dans le
ruban lorsque la série est terminée.

| Type | Exemple | Résultat attendu |
| --- | --- | --- |
| Texte | nom d'un fournisseur | texte nettoyé |
| Nombre | `12 345,67 €` | valeur numérique `12345,67` |
| Date | `31/12/2025` | vraie date Excel |
| Somme | trois montants | somme numérique |
| Tableau | lignes alignées | données réparties sur plusieurs cellules |

Résultat complémentaire : la cellule prend la couleur de preuve et comporte un
commentaire commençant par `DOCTRACKER-SNIP:`.

## 5. Retour à la preuve

- double-cliquer sur une cellule snippée.

Résultat attendu : le volet s'ouvre, sélectionne le bon document, rejoint la
bonne page et surligne la zone.

## 6. Revue

- sélectionner une cellule snippée ;
- cliquer sur **Revoir** ;
- choisir `Reviewed` et saisir un commentaire ;
- enregistrer.

Résultat attendu : le statut et le commentaire sont présents dans `project.xml`
et le commentaire Excel reflète le statut.

## 7. Matching

- importer plusieurs factures ;
- sélectionner la plage de recherche ;
- cliquer sur **Définir recherche** ;
- sélectionner la plage de résultat ;
- cliquer sur **Définir résultat** ;
- cliquer sur **Lancer le matching** ;
- cliquer sur **Rechercher** depuis une cellule pour vérifier la recherche dans toutes les pièces OCR.

Résultat attendu : chaque ligne reçoit dans la colonne de sortie le document,
la page et le score lorsqu'un rapprochement est trouvé. Une preuve page entière
et un commentaire sont associés au résultat ; le relecteur doit ensuite valider
ou rejeter le rapprochement.

## Informations à transmettre en cas d'erreur

Envoyer :

1. la capture du message ;
2. l'étape exacte de cette recette ;
3. la version d'Excel et son architecture 32/64 bits ;
4. le journal de build GitHub si l'échec intervient avant l'installation.

Ne pas transmettre les pièces confidentielles. Un PDF ou une image fictive
suffit pour reproduire un défaut d'interface.

Si seul l'OCR échoue avec un message relatif au moteur natif, vérifier la
présence des redistribuables Microsoft Visual C++ 2015-2022 x86 et x64, puis
relancer Excel.
