# Bilan mensuel opérateurs — QR codes recto-verso

Chaque mois, en 2 commandes, depuis le classeur `Suivi_Op_PP_OFFICIEL` :

```bash
pip install openpyxl segno
python3 outils/bilan_mensuel.py Suivi_Op_PP_OFFICIEL.xlsx --mois 10 --annee 2026 --sortie bilan
node    outils/html_vers_pdf.js bilan
```

Sortie dans `bilan/` (jamais versionné, données personnelles) :

| Fichier | Usage |
|---|---|
| `cartes_OCT2026.pdf` | À imprimer **recto-verso, bord long**, puis couper en 4 (A6). Recto : QR perso → bilan du mois. Verso : QR → les 18 standards. |
| `synthese_OCT2026.pdf` | Classement complet, **confidentiel superviseur / RH**. |
| `donnees_OCT2026.json` | Contrôle des calculs. |

## Ce que voit l'opérateur (fiche.html)
Note /20 et rang, chiffres du mois face à la moyenne de l'équipe et au mois précédent,
note par critère, courbe jour par jour, record du mois, 10 badges, vie au poste (journal RH),
défi personnalisé avec lien direct vers le standard à revoir (`index.html#P9`).

## Confidentialité
Les chiffres sont **dans le QR** (après le `#` de l'URL, jamais transmis au serveur GitHub) :
rien de personnel n'est publié sur le site. La page efface le lien de la barre d'adresse
dès l'ouverture.

## Notation — identique pour tous
- Classés : actifs, hors pilotes, soumis au quota (case vide = oui), ≥ 74 h de picking dans le mois.
- Par critère : `note = 10 + 4 × (valeur − moyenne équipe) / écart-type`, bornée 0–20 → moyenne = 10/20.
- Poids : productivité (quota) 40 % · qualité décelée (PPM décelée) 20 % · qualité avérée (PPM avérée) 20 % · RH & savoir-être (positifs − écarts du journal SUIVI_RH, ramené à 20 jours) 20 %.
- Équipe confirmée : note finale × 1,10 (10 % moins sévère), plafonnée à 20.
- Nouveaux (liste `NOUVEAUX` dans le script : Kellerer, Dembele, Hamdani, Robin) : productivité, avérée et décelée notées sur leur objectif d'intégration — 500 lignes/jour, PPM avérée 300, PPM décelée 3 200 (objectif atteint = 12/20) — puis +2 points de bonus. Classement séparé.
- Jamais pris en compte : absences, maladie, AT, délégation, congés, formation, contrat, ancienneté.

Réglages (seuils, poids, URL du site) en tête de `bilan_mensuel.py`.
