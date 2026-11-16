# Règles de travail du cours N3 Autonomie

## Autorité éditoriale

Le responsable éditorial est le seul à valider les modifications de fond. Le contenu extrait du PowerPoint est une source originale : ne pas le reformuler, le corriger, le compléter ni le supprimer sans instruction explicite.

## Sources et traçabilité

- `sources/` est immuable. Ne jamais modifier, renommer ni remplacer son contenu.
- `slides/` contient la transcription éditable. Distinguer toute proposition du contenu original.
- `assets/original/` contient les actifs extraits et ne doit pas être retouché.
- `assets/svg/` reçoit uniquement les futurs visuels vectoriels validés.
- `dist/` reçoit les sorties générées ; aucune sortie ne doit écraser une source.
- Ne jamais inventer un texte manquant : marquer l’élément comme nécessitant une vérification.

## Statuts des diapositives

- `extracted` : contenu brut importé.
- `in_review` : contenu en cours de travail.
- `validated` : contenu approuvé ; ne jamais le modifier automatiquement.
- `needs_verification` : extraction incomplète ou techniquement douteuse.

Avant toute écriture dans un fichier Markdown, lire son statut. Toute modification d’une diapositive `validated` exige une instruction explicite du responsable éditorial.

## Visuels et génération

- Ne pas convertir automatiquement les images matricielles en SVG.
- Ne pas remplacer un schéma ou une illustration sans validation.
- Les nouveaux SVG doivent rester structurés et éditables dans VS Code et Inkscape.
- Prévisualiser tout nouveau visuel avant validation.
- La future génération PowerPoint doit conserver textes et tableaux éditables et ne jamais modifier les Markdown.

## Git et validation

- Présenter les fichiers concernés et les changements proposés sous forme de diff.
- Ne pas créer de commit sans accord explicite.
- Ne pas effectuer de recherche externe ou de vérification réglementaire sans demande explicite.
