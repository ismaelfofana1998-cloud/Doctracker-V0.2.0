# Doctracker 0.9.4 — affichage après modification d’un snip

Le journal de la version 0.9.3 signale une `InvalidOperationException` dans
`Image.get_Width` / `GetPagePicture`, après un redimensionnement. L’OCR isolé
avait terminé. Ce journal confirme une erreur du lecteur, mais ne contient pas
les identifiants de threads permettant de prouver l’origine exacte du conflit GDI+.

## Corrections

- Les continuations des opérations asynchrones repassent explicitement sur le
  thread du panneau, y compris en cas d’annulation ou d’erreur. Les callbacks
  Office peuvent fournir un contexte de synchronisation absent ou non graphique.
  La reconnaissance reste dans le processus séparé ; Excel et les images du
  lecteur sont manipulés sur le thread de l’interface.
- Les dimensions des aperçus sont conservées séparément : les rafraîchissements
  ne relisent plus `Image.Size` dans GDI+. Le rendu des annotations ne modifie
  plus temporairement la page active et ne restaure plus une ancienne référence
  d’image. Les changements d’image reportent les demandes de layout imbriquées.
- Le journal indique désormais le thread et son appartement, ainsi que la fin
  d’une modification de snip. Il reste borné et n’enregistre pas les documents,
  le texte reconnu ou les valeurs Excel.

## Performance et validation

Cette version évite des accès GDI+ et des rafraîchissements imbriqués inutiles.
Elle ne change ni la résolution OCR ni les protections de la version 0.9.3.
Le lancement du processus OCR séparé garde un coût ; aucun gain de vitesse de
reconnaissance n’est revendiqué.

Les tests Windows vérifient les retours asynchrones avec un contexte absent ou
incorrect (succès, résultat immédiat, erreur, annulation), ainsi que 32 cycles de
navigation/modification de zone/zoom avec layout imbriqué et contrôle des images
encore affichées. Les essais PDF, OCR, mémoire et interruption des workers restent
actifs en x86 et x64. Le runner ne dispose pas d’Excel : la mise à jour réelle des
cellules et le scénario exact sur le document de l’utilisateur restent à confirmer
sur un poste Office.
