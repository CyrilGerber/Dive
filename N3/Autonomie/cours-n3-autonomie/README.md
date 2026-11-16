# Cours N3 Autonomie FFESSM

Ce dépôt contient une extraction de travail fidèle de la présentation source de 27 diapositives. Le PowerPoint original est conservé dans `sources/` et ne doit jamais être modifié.

## Organisation

- `sources/` : PowerPoint source immuable.
- `slides/` : un fichier Markdown éditable par diapositive.
- `assets/original/` : médias et composants graphiques extraits sans retouche.
- `assets/svg/` : futurs schémas SVG validés.
- `assets/images/` : futures images de travail, distinctes des originaux.
- `theme/` : charte et réglages de présentation.
- `scripts/` : extraction et prévisualisation reproductibles.
- `dist/` : futures présentations générées ; ignorées par Git.

## Réextraire le contenu

Depuis la racine du projet :

```powershell
python scripts/extract_pptx.py "sources/Préparation NIVEAU III_Autonomie_PG+CG.pptx"
```

L’extraction réécrit les fichiers générés dans `slides/`, `manifest.yml`, `EXTRACTION_REPORT.md` et les copies d’actifs. Ne l’exécuter qu’en connaissance de cause si des diapositives ont déjà été éditées.

Les fichiers au statut `validated` sont protégés et ignorés par l’extracteur. Pour contrôler la cohérence du projet :

```powershell
python scripts/validate_project.py
```

## Visualiser

Pour une prévisualisation locale sans dépendance :

```powershell
python scripts/preview.py
```

Ouvrir ensuite `http://127.0.0.1:8000/slides/`. Les SVG peuvent être ouverts directement dans le navigateur, VS Code ou Inkscape.

## Éditer une première diapositive

1. Ouvrir son fichier dans `slides/` et vérifier son statut YAML.
2. Ne pas modifier un statut `validated` sans instruction explicite.
3. Conserver le texte original ; placer les propositions dans une branche ou un diff distinct selon les consignes éditoriales.
4. Vérifier les références de visuels dans la section `Visuels`.
5. Examiner le diff Git avant validation.

## Préparation PptxGenJS

`package.json` réserve la dépendance PptxGenJS et une commande de contrôle. Node.js/npm ne sont pas installés dans l’environnement actuel. Aucune présentation n’est générée à ce stade.

Après installation future de Node.js :

```powershell
npm install
npm run pptx:check
```
