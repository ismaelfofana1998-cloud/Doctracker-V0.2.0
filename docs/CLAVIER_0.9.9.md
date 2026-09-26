# Édition immédiate après un snip — 0.9.9

Le saut automatique à la prochaine cellule reste actif. Le correctif porte sur le transfert du clavier entre le lecteur PDF et Excel, ainsi que sur le réaffichage automatique de la preuve quand on revient à une cellule.

Avant de désactiver le lecteur pour extraire le texte, le focus natif est rendu à la feuille. Le suivi différé de sélection ne rafraîchit la preuve que si le clavier appartient encore à la grille Excel. Il mémorise le contrôle exact et restaure son focus si le réaffichage de la preuve le déplace dans le panneau ou le laisse vide. Il ne force pas le focus sur un éditeur ou une autre fenêtre et n’envoie aucune touche simulée. Le panneau n’est rendu visible que s’il était masqué.

Une modification de la valeur de la cellule pendant l’OCR empêche l’écriture du résultat tardif. Une saisie Excel en cours détectée à la fin de l’OCR laisse la main à l’utilisateur.

## Vérification sur Excel

1. Faire un snip et laisser la sélection passer à la cellule suivante.
2. Revenir immédiatement sur la cellule précédente, puis taper du texte pour remplacer son contenu.
3. Refaire avec F2 puis saisie, et avec un double-clic puis saisie, sans attendre.
4. Répéter avec une preuve située sur une autre page ou un autre document, panneau déjà ouvert puis masqué.
5. Vérifier que le texte reste dans la cellule, que les premières touches ne déclenchent pas le ruban et que le saut à la cellule suivante fonctionne toujours.

Les tests Windows automatisés exercent les API natives sur des contrôles WinForms : transfert, premières frappes, réaffichage, focus nul, éditeur concurrent, fenêtre externe et contrôles masqués/détruits. Le runner n’a pas Excel : ils ne remplacent pas la validation de ce parcours dans le complément VSTO installé.
