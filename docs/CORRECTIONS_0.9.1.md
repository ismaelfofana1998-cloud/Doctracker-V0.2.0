# Doctracker 0.9.1 — zoom et liens Excel

Le lecteur conserve les pages dans un conteneur continu. La barre de défilement horizontale permet d'atteindre le bord droit après zoom ; le glissement avec le bouton central reste disponible. Changer de page conserve le déplacement horizontal.

## Supprimer les snips d'une plage

Sélectionner la plage Excel puis **Doctracker > Supprimer les snips de la plage**. La commande retire tous les liens de cette sélection et conserve les valeurs/formules. Une preuve encore utilisée hors de la sélection reste disponible. Une preuve sans autre cellule liée est supprimée du PDF. Les cellules protégées, fusionnées ou matricielles bloquent l'opération avant toute modification. Les totaux existants ne sont pas recalculés : leurs valeurs/formules sont préservées.

La suppression par clic droit sur un snip garde son sens : supprimer cette preuve et tous ses liens dans le classeur.

## Pourquoi un lien peut disparaître

Les liens sont des noms Excel masqués. Excel les ajuste lors des insertions de lignes/colonnes. Le cache des adresses pouvait rester périmé après certaines opérations sans événements Excel, et l'ancienne réparation pouvait réattacher une preuve déjà déplacée à son ancienne adresse. La navigation explicite, la suppression et la réparation relisent désormais les références réelles. Le nom interne stable de la feuille est également conservé lors de l'enregistrement.

Une cellule ou feuille supprimée produit une référence Excel `#REF!` : aucune adresse fiable ne peut en être déduite. La réparation signale ces preuves au lieu de les attribuer silencieusement à une autre cellule. Sélectionner la bonne cellule puis cliquer sur la preuve dans le PDF propose un rattachement, sans changer sa valeur. Les anciennes destinations des preuves entièrement déliées restent conservées lors de l'enregistrement pour la récupération.

**Limites :** un copier-coller de valeurs seul ne copie pas les noms Excel. Le tri peut déplacer les valeurs sans déplacer leurs références nommées ; vérifier les preuves après tri. Les noms supprimés manuellement et les anciennes métadonnées ne permettent pas de deviner une nouvelle destination avec certitude. Ces situations nécessitent une vérification humaine. Aucun diagnostic du classeur de l'utilisateur n'est possible sans le fichier.

## Dossiers de cache

Un seul emplacement racine : `%LOCALAPPDATA%\Doctracker\Cache`. Ses sous-dossiers isolent les sessions de classeurs, notamment quand plusieurs classeurs ou processus Excel sont ouverts. Ils contiennent les documents extraits du classeur et les index réutilisables. Leur présence ne lance aucune OCR ni tâche d'indexation. Les fusionner peut mélanger les projets ; cela n'apporterait pas un gain de calcul.

Les caches créés par la session sont supprimés à la libération du contexte (nettoyage des classeurs fermés ou arrêt normal d'Excel). Une interruption brutale peut laisser des dossiers résiduels. Après enregistrement des classeurs et fermeture de toutes les instances Excel, les anciens sous-dossiers de `Cache` peuvent être supprimés. Ne pas supprimer les anciens dossiers `Recovery` sans vérifier s'ils contiennent une récupération nécessaire. Aucun nettoyage destructif des anciennes récupérations n'est ajouté.

## Vérifications

Tests automatiques du lecteur WinForms sous Windows : défilement horizontal réel après zoom, stabilité après plusieurs rafraîchissements, maintien du décalage lors du changement de page, accès au bord droit, lecture verticale continue et cache de rendu borné. Tests Core : suppression par plage, preuve partagée, conservation des autres cellules, identité de feuille et retour arrière sur échec d'écriture.

À valider dans Excel de bureau : insérer une ligne, renommer une feuille, enregistrer/rouvrir, cliquer sur la preuve ; supprimer sa ligne puis la rattacher à une destination explicite ; retirer les snips d'une plage contenant une preuve partagée. La CI ne dispose pas d'Excel pour exécuter ces scénarios COM.
