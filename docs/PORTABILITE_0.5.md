# Doctracker 0.5 — pièces, classement et récupération

## Utilisation

Le menu **…** à côté d’Importer rassemble les nouvelles actions : import de dossier,
catégorisation, Xref, stockage, sauvegarde, restauration, réparation et export annoté.
Le sélecteur de catégorie filtre les documents, la recherche, le matching et l’export.
Une pièce peut appartenir à plusieurs catégories (séparées par `;`). L’import récursif
conserve la catégorie principale sur chaque pièce et les sous-dossiers sous forme
de catégories supplémentaires, ignore les formats non pris en
charge et les jonctions, signale les accès refusés et déduplique les contenus identiques.

## Transmettre seulement Excel

Par défaut, les pièces et leurs index sont intégrés dans les parties XML personnalisées
du classeur lors de l’enregistrement. Utiliser `.xlsx`, `.xlsm` ou `.xlsb`, enregistrer,
puis transmettre le classeur. Le destinataire ouvre ce fichier avec Doctracker 0.5.
Les GUID des preuves et les empreintes SHA-256 des sources sont indépendants du
nom du classeur, des chemins utilisateur et des noms affichés des documents.

Il n’y a plus de nouveau dossier de mission à côté d’Excel. L’application utilise
un cache de travail et de récupération dans `%LOCALAPPDATA%\Doctracker\Recovery`.
Ce cache technique est nécessaire aux moteurs PDF/OCR ; il n’est pas à transmettre.
Les sources intégrées sont extraites à la demande. Les anciens dossiers 0.4 sont
copiés vers le cache, puis intégrés au prochain enregistrement ; l’ancien dossier
reste intact. Vérifier la copie transférée avant de l’archiver. Tous les intervenants
doivent passer en 0.5 : le schéma des métadonnées change.

L’intégration augmente le poids du classeur. Limite volontaire : **256 Mo de pièces
et index intégrés**, avant l’encodage XML. Ce mode ne promet pas un classeur léger
avec des gigaoctets de documents. Un enregistrement incomplet ou en conflit est
interrompu, avec un message, plutôt que de supprimer silencieusement des preuves.
Les modifications ne voyagent qu’après enregistrement d’Excel. Ne pas convertir en
CSV / ancien XLS, ni supprimer les données XML via l’inspecteur de document.

## Documents partagés pour les gros dossiers

Choisir **Stockage autonome / partagé → Oui**, puis un emplacement UNC commun
(`\\serveur\partage\mission`) accessible à tous les intervenants. L’application y
publie les sources sous leur empreinte ; Excel conserve les références, les preuves
et les index. Les collègues téléchargent la pièce dans leur cache à sa première
consultation, sans recevoir un ZIP de toutes les pièces à chaque transfert.
Le passage en mode autonome récupère les sources, puis les intègre au prochain Save.

Les droits réseau restent ceux du cabinet : le destinataire doit y avoir accès,
par son réseau/VPN si nécessaire. Une lettre de lecteur ou un chemin OneDrive
personnel est refusé. Il n’y a pas, dans cette version, de connecteur SharePoint,
de serveur internet ni de publication publique automatique des documents.

Ce partage concerne **les sources documentaires**. Il ne fusionne pas les snips
créés dans plusieurs copies divergentes d’Excel. Les conflits de manifeste observés
sont refusés ; la coédition simultanée Excel/OneDrive n’est pas certifiée. Travailler
à tour de rôle sur le classeur de référence. Les réservations Xref sur le partage
sont atomiques et empêchent deux documents de prendre le même numéro.

## Sauvegarder et récupérer

Chaque modification validée conserve `project.xml`, sa version précédente et
20 instantanés des métadonnées dans le cache. Les sources et versions d’index
existantes ne sont pas écrasées. Une corruption XML simple déclenche une récupération
de la version précédente, signalée dans le projet. **Ouvrir les sauvegardes automatiques**
permet de retrouver ces fichiers ; **Restaurer une sauvegarde** accepte `.xml` et `.dtpack`.
Un XML seul peut récupérer les liens, mais il faut encore les sources du cache ou du
partage. Ce n’est pas une protection contre la perte du disque ou de tout le classeur.

Pour une sauvegarde indépendante, utiliser **Sauvegarder toutes les pièces et liens** :
le `.dtpack` contient les pièces originales, les index, catégories, Xref et preuves.
Le stocker sur un autre support/emplacement. Sa restauration se fait dans un nouveau
cache et vérifie les empreintes, les chemins et la présence des sources avant bascule.

**Réparer les liens des cellules** réattache les commentaires Doctracker sans modifier
les valeurs. Les positions des commentaires sont relevées à l’enregistrement, avant sauvegarde
complète et avant export PDF pour
suivre les déplacements et les copies de cellules. En cas de perte antérieure à ce
relevé, les dernières positions connues sont utilisées : vérifier les lignes déplacées.
Les feuilles absentes sont signalées. Une restauration ne recrée pas les valeurs,
formules ou la mise en page d’un classeur Excel perdu.

## Xref et export PDF

Attribuer une référence, par exemple `DAC B 30 040`, puis un numéro disponible `01`.
Le nom affiché et le nom d’export deviennent `DAC B 30 040 - 01 - Nom initial.pdf`.
Les numéros déjà réservés ne sont pas proposés et ne sont pas recyclés lors d’un
retrait. L’identifiant de la pièce et les octets de l’original restent inchangés.

**Exporter la catégorie en PDF annotés** exporte les pièces de la catégorie active
(toutes si aucune catégorie filtrée). Les cadres colorés et Xref/cellules sont dessinés
sur les pages. Ce sont des **copies aplaties et rasterisées** : la présentation est
conservée, mais le texte natif sélectionnable et les signatures numériques originales
ne sont pas conservés dans ces copies. L’original reste disponible dans la sauvegarde.
Les noms déjà présents dans la destination ne sont pas écrasés.

## Recherche et performance

L’option **Références partielles** trouve `X300`, `X300Z35`, `500X300Z35` dans
`500X300 Z35` et tient compte des préfixes/suffixes. Les fragments de moins de trois
caractères sont exclus de ce mode. Les nombres et dates restent comparés comme des
valeurs exactes. Les résultats partiels demandent confirmation avant insertion ;
les ambiguïtés entre documents/pages restent non résolues.

Les index sont compressés et chargés par document, puis libérés après recherche.
Le matching parcourt les pages une fois pour le lot de lignes Excel ; deux candidats
suffisent à signaler une ambiguïté. Les pièces inchangées ne sont pas réintégrées
à chaque Save. L’import de dossier valide les métadonnées par lots de 25 pièces.
Cela réduit la mémoire et les écritures ; aucune promesse de volume illimité n’est faite.

## Validation et recette indispensable

Tests automatiques : transfert du contenu intégré entre deux caches, sources intactes,
index à la demande, erreur d’écriture, conflits, corruption, restauration, traversée de
chemin interdite, Xref, mode partagé, recherche partielle et annulation ; tests Windows
PDF/OCR et export annoté rouvert par PDFium. Les parties Excel sont simulées dans les
tests du moteur : le vrai stockage COM doit encore être testé dans Excel.

Recette sur deux PC : importer dossier et catégories, créer des preuves/Xref, déplacer
et copier quelques cellules, sauvegarder/fermer, envoyer seulement Excel, vérifier
les preuves, modifier/retransmettre, exporter les PDF. Ensuite tester le partage UNC
avec deux comptes autorisés, la coupure réseau et la récupération `.dtpack` après
suppression d’une copie de test du classeur/cache. Tester séparément les changements
de nom et Enregistrer sous, sans écraser la seule copie de production.

Référence technique : Microsoft, Custom XML parts overview et Add custom XML parts
using VSTO Add-ins : https://learn.microsoft.com/en-us/visualstudio/vsto/custom-xml-parts-overview
et https://learn.microsoft.com/en-us/visualstudio/vsto/how-to-add-custom-xml-parts-to-documents-by-using-vsto-add-ins
