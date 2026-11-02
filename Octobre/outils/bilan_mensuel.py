#!/usr/bin/env python3
"""
BILAN MENSUEL PICK & PACK — générateur des QR codes opérateurs (recto-verso)
=============================================================================

Lit le classeur de suivi (Suivi_Op_PP_OFFICIEL) et produit, pour un mois donné :

  1. cartes_<MOIS>.html / .pdf   A4, 4 cartes A6 par feuille, recto-verso
                                 recto  = QR personnel -> bilan du mois (fiche.html)
                                 verso  = QR standards -> index.html
  2. synthese_<MOIS>.html / .pdf Classement complet, réservé au superviseur
  3. donnees_<MOIS>.json         Données calculées (contrôle / archivage)

Aucune donnée personnelle n'est publiée sur le site : chaque QR porte les
chiffres de l'opérateur dans l'URL (après le #, jamais envoyé au serveur).

Utilisation
-----------
  python3 outils/bilan_mensuel.py Suivi_Op_PP_OFFICIEL.xlsx --mois 10 --annee 2026 --sortie ./bilan
  node    outils/html_vers_pdf.js ./bilan            # convertit les .html en .pdf

Méthode de notation (impartiale, identique pour tous)
-----------------------------------------------------
Population classée : opérateurs actifs, hors pilotes, soumis au quota (case vide =
oui), avec au moins MIN_HEURES heures de picking dans le mois (sinon « non classé »).
Pour chaque critère, la note /20 dépend de l'écart à la moyenne de l'équipe :
    note = 10 + 4 × (valeur − moyenne) / écart-type      (bornée 0 – 20)
    => pile dans la moyenne = 10/20 ; +1 écart-type = 14/20 ; −1 = 6/20.
Critères et poids :
    Productivité        40 %  quota moyen du mois (%)
    Qualité décelée     20 %  PPM décelée (erreurs au rejet / lignes × 1e6) — plus bas = mieux
    Qualité avérée      20 %  PPM avérée  (litiges / lignes × 1 000 000)  — plus bas = mieux
    RH & savoir-être    20 %  (points positifs − écarts) du journal SUIVI_RH, ramené à 20 jours
Coup de pouce : note finale de l'équipe confirmée × COUP_DE_POUCE (1,10), plafonnée à 20.
Nouveaux (NOUVEAUX) : prod., avérée, décelée notées sur OBJ_NOUVEAU (objectif atteint
= 12/20), puis + BONUS_NOUVEAU points  ; aucun classement affiché.
Ne sont JAMAIS pris en compte : absences, maladie, AT, délégation, congés,
contrat (CDI / intérim), ancienneté.
"""
from __future__ import annotations

import argparse
import base64
import datetime as dt
import html
import json
import math
import statistics
import warnings
import zlib
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl
import segno

warnings.filterwarnings("ignore")

# ─────────────────────────── réglages ───────────────────────────
BASE_URL = "https://legros123.github.io/Standard-Pick-Pack/"          # appli Standards (verso)
SITE_BILAN = "https://legros123.github.io/Suivi-P-P-/"                 # pages perso (recto), 1 dossier par mois
GABARIT = Path(__file__).resolve().parent.parent / "fiche.html"        # modèle de page bilan
MIN_HEURES = 74.0          # = 10 jours pleins de picking pour être classé (note fiable)
BADGE_HEURES = 37.0        # volume mini pour décrocher un badge qualité (5 jours)
MIN_LIGNES_CARTE = 100     # en dessous : pas de carte (passage trop ponctuel)
MIN_HEURES_CARTE = 15.0    # heures de picking réelles (hors HG, renforts…) pour recevoir une carte
POIDS = {"p": 0.40, "d": 0.20, "a": 0.20, "r": 0.20}
COUP_DE_POUCE = 1.10       # équipe confirmée : note finale × 1,10 (10 % moins sévère), plafonnée à 20
# Parcours nouveaux : notés sur LEUR objectif, pas sur la moyenne des confirmés
NOUVEAUX = {"KELLERER MEGANE", "DEMBELE TAMBA", "HAMDANI AMINE", "ROBIN ELODIE"}
OBJ_NOUVEAU = {"lj": 500, "pa": 300, "pd": 3200}   # lignes/jour, PPM avérée, PPM décelée
BONUS_NOUVEAU = 2.0        # points ajoutés à la note finale d'un nouveau
PPM_DEC_VERT = 1000
SEUIL_OR = 78              # carte or (et badge « Carte Or ») à partir de ce niveau OVR
QUOTA_JOUR, H_REF, QUOTA_MAX = 620, 7.4, 150   # PARAMETRES B5, B6, B13

# Corrections validées par le superviseur, en attendant leur report dans le classeur.
# H. HORS PICK (heures passées hors picking, ex. renfort HG) par opérateur et par jour.
CORRECTIONS_HHP = {
    "GRANGER JEAN PHILIPPE": {   # mouvements vers HG non saisis — oct. 2026
        "2026-10-13": 3.50, "2026-10-14": 3.00, "2026-10-15": 3.00, "2026-10-20": 2.00,
        "2026-10-21": 3.00, "2026-10-22": 2.30, "2026-10-23": 2.50, "2026-10-26": 4.40,
        "2026-10-27": 8.40, "2026-10-28": 4.40, "2026-10-29": 4.40,
    },
}
# Points du journal RH ajoutés par le superviseur : (opérateur, date, type, sens, famille, nombre)
AJOUTS_RH = [
    ("GRANGER JEAN PHILIPPE", "2026-10-31", "Annulation de RTT (forte activité)", "+", "SAVOIR-ÊTRE", 1),
    ("GRANGER JEAN PHILIPPE", "2026-10-31", "Polyvalence : renfort expédition en heures sup", "+", "SAVOIR-ÊTRE", 3),
]        # PARAMETRES : PPM décelée verte sous ce seuil
MOIS_FR = ["", "JANVIER", "FÉVRIER", "MARS", "AVRIL", "MAI", "JUIN", "JUILLET",
           "AOÛT", "SEPTEMBRE", "OCTOBRE", "NOVEMBRE", "DÉCEMBRE"]
MOIS_DOSSIER = ["", "Janvier", "Fevrier", "Mars", "Avril", "Mai", "Juin", "Juillet", "Aout", "Septembre", "Octobre", "Novembre", "Decembre"]
MOIS_COURT = ["", "JAN", "FEV", "MARS", "AVR", "MAI", "JUIN", "JUIL", "AOUT", "SEPT", "OCT", "NOV", "DEC"]

# ─────────────────────────── lecture ───────────────────────────

def norm_mat(v):
    if v is None:
        return ""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s.lstrip("0") or "0"


def prenom_joli(p: str) -> str:
    return " ".join("-".join(x.capitalize() for x in w.split("-")) for w in str(p or "").split())


def lire(xlsx: Path):
    wb = openpyxl.load_workbook(xlsx, data_only=True, read_only=True)

    # registre
    ops = {}
    for r in wb["OPERATEURS"].iter_rows(min_row=7, values_only=True):
        if not r[1]:
            continue
        ops[r[3]] = dict(mat=norm_mat(r[0]), nom=r[1], prenom=r[2], np=r[3], role=r[4] or "",
                         poste=r[5] or "", contrat=r[6] or "", actif=(r[9] == "OUI"),
                         quota=(r[12] != "NON"))   # vide = OUI (seul NON exclut)

    # production (une ligne = un opérateur x un jour)
    prod = []
    for r in wb["PROD"].iter_rows(min_row=5, values_only=True):
        if not isinstance(r[0], dt.datetime) or not r[2]:
            continue
        if (r[25] or 0):          # DOUBLON
            continue
        x = dict(date=r[0].date(), np=r[2], h=r[3] or 0, lig=r[5] or 0, cont=r[10] or 0,
                 ed=r[11] or 0, ea=r[12] or 0, q=r[13], reg=r[23])
        hhp = CORRECTIONS_HHP.get(r[2], {}).get(x["date"].isoformat())
        x["hhp"] = hhp if hhp is not None else (r[4] or 0)
        if hhp is not None and r[18]:   # recalcul du quota du jour avec les vraies heures de picking
            h_reel = max(0.0, x["h"] - hhp)
            x["q"] = round(min(QUOTA_MAX, r[18] / (h_reel / H_REF * QUOTA_JOUR) * 100), 1) if h_reel > 0 else None
        prod.append(x)

    # erreurs décelées (types) et avérées (motifs)
    err = [(r[0].date(), r[2], r[3] or "NON PRÉCISÉ") for r in wb["ERREURS"].iter_rows(min_row=4, values_only=True)
           if isinstance(r[0], dt.datetime) and r[2]]
    ave = [(r[0].date(), r[2], r[6] or "AUTRE") for r in wb["AVEREES"].iter_rows(min_row=4, values_only=True)
           if isinstance(r[0], dt.datetime) and r[2]]

    # journal RH : bloc « JOURNAL » de SUIVI_RH (colonnes B..H)
    rh, dans_journal = [], False
    for r in wb["SUIVI_RH"].iter_rows(values_only=True):
        if r[1] == "OPÉRATEUR  (liste)":
            dans_journal = True
            continue
        if dans_journal and r[1] and isinstance(r[2], dt.datetime):
            sens = "+" if "POSITIF" in str(r[4]) else "-"
            rh.append(dict(np=r[1], date=r[2].date(), type=r[3], sens=sens, fam=r[5] or ""))
    for np_, d, t, sens, fam, n in AJOUTS_RH:
        rh += [dict(np=np_, date=dt.date.fromisoformat(d), type=t, sens=sens, fam=fam)] * n

    # pointage : jours / heures / h. supp du mois
    pt = {}
    rows = list(wb["POINTAGE"].iter_rows(values_only=True))
    return wb, ops, prod, err, ave, rh, rows


def pointage_mois(rows, mois, annee):
    titre = f"{MOIS_FR[mois].replace('É', 'E').replace('Û', 'U')} {annee}"
    out, actif = {}, False
    for r in rows:
        if isinstance(r[0], str) and r[0].strip().upper() == titre:
            actif = True
            continue
        if actif and isinstance(r[0], str) and r[0].strip().endswith(str(annee)) and r[0].strip().upper() != titre:
            break
        if actif and r[1] and r[1] not in ("NOM PRENOM", "TOTAL DU JOUR") and isinstance(r[34], (int, float)):
            out[r[1]] = dict(h=float(r[34] or 0), j=int(r[35] or 0), hs=float(r[37] or 0))
    return out

# ─────────────────────────── calculs ───────────────────────────

def agreger(prod, np_, mois, annee):
    d = [x for x in prod if x["np"] == np_ and x["date"].month == mois and x["date"].year == annee]
    if not d:
        return None
    h = sum(max(0.0, x["h"] - x["hhp"]) for x in d)   # heures réellement passées au picking
    lig = sum(x["lig"] for x in d)
    qs = [x["q"] for x in d if isinstance(x["q"], (int, float))]
    regs = [x["reg"] for x in d if isinstance(x["reg"], (int, float))]
    jours_pick = [x for x in d if x["h"] > 0]
    best = max(d, key=lambda x: x["lig"])
    return dict(
        h_pick=round(h, 2), lig=int(lig), cont=int(sum(x["cont"] for x in d)),
        ed=int(sum(x["ed"] for x in d)), ea=int(sum(x["ea"] for x in d)),
        q=round(statistics.mean(qs), 1) if qs else None,
        reg=round(100 * statistics.mean(regs)) if regs else None,
        jp=len(jours_pick),
        lj=round(lig / len(jours_pick)) if jours_pick else None,
        cad=round(lig / h, 1) if h else None,
        best=(int(best["lig"]), best["date"].day),
        serie=[(x["date"].day, round(x["q"])) for x in sorted(d, key=lambda x: x["date"])
               if isinstance(x["q"], (int, float)) and x["h"] > 0],
        jours_det=[[x["date"].day, round(x["h"], 2), round(max(0.0, x["h"] - x["hhp"]), 2), int(x["lig"]),
                    int(x["ed"]), int(x["ea"]), round(x["q"]) if isinstance(x["q"], (int, float)) else None]
                   for x in sorted(d, key=lambda x: x["date"])],
    )


def ppm(n, lig):
    return round(n / lig * 1_000_000) if lig else None


def note_objectif(v, obj, plus_bas_mieux=False):
    """Nouveaux : objectif atteint = 12/20 ; chaque écart de 100 % de l'objectif vaut 8 points."""
    if v is None:
        return None
    x = (obj - v) / obj if plus_bas_mieux else (v - obj) / obj
    return round(max(0.0, min(20.0, 12 + 8 * x)), 1)


def zscore_note(v, moy, sd, plus_bas_mieux=False):
    if v is None:
        return None
    if not sd:
        return 10.0
    z = (v - moy) / sd
    if plus_bas_mieux:
        z = -z
    return round(max(0.0, min(20.0, 10 + 4 * z)), 1)


STD_LIEN = {  # type d'écart / erreur -> standard à revoir
    "Carton non ramassé": ("P9", "Gestion des déchets"),
    "Prise de poste KO": ("P1", "Prise de poste"),
    "Fin de poste KO": ("P11", "Fin de poste"),
    "EPI non porté": ("P1", "Prise de poste"),
    "Glissement de stop": ("P10", "Circulation & sécurité"),
    "Autre écart sécurité": ("P18", "Sécurité au quotidien"),
    "Non-respect du standard": ("P14", "Les interdits"),
}


def construire(xlsx: Path, mois: int, annee: int):
    wb, ops, prod, err, ave, rh, prow = lire(xlsx)
    pt = pointage_mois(prow, mois, annee)
    mprec, aprec = (mois - 1, annee) if mois > 1 else (12, annee - 1)

    fiches = []
    for np_, o in ops.items():
        if not o["actif"] or o["role"].lower().startswith("pilote"):
            continue
        a = agreger(prod, np_, mois, annee)
        if not a or a["lig"] < MIN_LIGNES_CARTE or a["h_pick"] < MIN_HEURES_CARTE:
            continue
        p = agreger(prod, np_, mprec, aprec)
        rh_m = [x for x in rh if x["np"] == np_ and x["date"].month == mois and x["date"].year == annee]
        pos = sum(1 for x in rh_m if x["sens"] == "+")
        ec = sum(1 for x in rh_m if x["sens"] == "-")
        sec = sum(1 for x in rh_m if x["sens"] == "-" and "SÉCU" in x["fam"].upper())
        ptg = pt.get(np_, {})
        jours = ptg.get("j") or a["jp"]
        f = dict(
            np=np_, mat=o["mat"], nom=o["nom"], prenom=prenom_joli(o["prenom"]),
            quota_ok=o["quota"], **a,
            jours=jours, heures=round(ptg.get("h") or a["h_pick"], 1), hs=round(ptg.get("hs") or 0, 1),
            pa=ppm(a["ea"], a["lig"]), pd=ppm(a["ed"], a["lig"]),
            prec=None if not p else dict(q=p["q"], lig=p["lig"], pa=ppm(p["ea"], p["lig"]), pd=ppm(p["ed"], p["lig"])),
            pos=pos, ec=ec, sec=sec,
            rh_types=Counter((x["sens"], x["type"]) for x in rh_m).most_common(),
            err_types=Counter(t for (d, n, t) in err if n == np_ and d.month == mois and d.year == annee).most_common(4),
            ave_types=Counter(t for (d, n, t) in ave if n == np_ and d.month == mois and d.year == annee).most_common(3),
        )
        f["rhv"] = round((pos - ec) / max(jours, 1) * 20, 2)
        f["classe"] = bool(o["quota"] and a["h_pick"] >= MIN_HEURES and a["q"] is not None)
        fiches.append(f)

    # statistiques de l'équipe classée
    cl = [f for f in fiches if f["classe"]]
    crit = {"p": ("q", False), "a": ("pa", True), "d": ("pd", True), "r": ("rhv", False)}
    equipe = {}
    for k, (champ, _) in crit.items():
        vals = [f[champ] for f in cl if f[champ] is not None]
        equipe[k] = (statistics.mean(vals), statistics.pstdev(vals)) if vals else (0, 0)
    for f in fiches:
        f["nouveau"] = f["np"] in NOUVEAUX
        notes = {}
        for k, (champ, inv) in crit.items():
            m, sd = equipe[k]
            notes[k] = zscore_note(f[champ], m, sd, inv)
        if f["nouveau"]:   # notés sur leur objectif d'intégration
            notes["p"] = note_objectif(f["lj"], OBJ_NOUVEAU["lj"])
            notes["a"] = note_objectif(f["pa"], OBJ_NOUVEAU["pa"], True)
            notes["d"] = note_objectif(f["pd"], OBJ_NOUVEAU["pd"], True)
        f["notes"] = notes
        f["note_brute"] = round(sum(POIDS[k] * notes[k] for k in POIDS), 1) if f["classe"] else None
        if f["note_brute"] is None:
            f["note"] = None
        elif f["nouveau"]:
            f["note"] = round(min(20.0, f["note_brute"] + BONUS_NOUVEAU), 1)
        else:
            f["note"] = round(min(20.0, f["note_brute"] * COUP_DE_POUCE), 1)
    groupes = {False: [f for f in cl if not f["nouveau"]], True: [f for f in cl if f["nouveau"]]}
    for g in groupes.values():
        g.sort(key=lambda f: (-f["note"], f["pa"] or 0, -(f["q"] or 0)))
        for i, f in enumerate(g, 1):
            f["rang"], f["N"] = i, len(g)
    for f in fiches:
        f.setdefault("rang", None)
        f.setdefault("N", None)

    # ── carte joueur (style jeu de foot) : 6 attributs sur 99, moyenne équipe ≈ 70 ──
    for f in fiches:
        f["secv"] = round(f["sec"] / max(f["jours"], 1) * 20, 2)
        f["posv"] = round(f["pos"] / max(f["jours"], 1) * 20, 2)
    extra = {}
    for k, champ in (("vol", "lj"), ("sec", "secv"), ("esp", "posv")):
        vals = [f[champ] for f in cl if f[champ] is not None]
        extra[k] = (statistics.mean(vals), statistics.pstdev(vals)) if vals else (0, 0)
    sur99 = lambda n: None if n is None else int(round(45 + 2.5 * n))   # 10/20 -> 70, 20/20 -> 95
    for f in fiches:
        n = f["notes"]
        f["attr"] = {
            "VIT": sur99(n["p"]),                                          # vitesse : quota (ou objectif nouveau)
            "VOL": sur99(zscore_note(f["lj"], *extra["vol"])),             # volume : lignes / jour
            "PRÉ": sur99(n["d"]),                                          # précision : PPM décelée
            "FIA": sur99(n["a"]),                                          # fiabilité : PPM avérée
            "SÉC": sur99(zscore_note(f["secv"], *extra["sec"], True)),      # sécurité : écarts sécu
            "ESP": sur99(zscore_note(f["posv"], *extra["esp"])),           # esprit d'équipe : points positifs
        }
        f["ovr"] = sur99(f["note"]) if f["note"] is not None else None
        f["carte"] = (None if f["ovr"] is None else "icone" if f["ovr"] >= 85 else "or" if f["ovr"] >= SEUIL_OR
                      else "argent" if f["ovr"] >= 65 else "bronze")
    attr_moy = {k: int(round(statistics.mean([f["attr"][k] for f in cl]))) for k in fiches[0]["attr"]} if cl else {}

    moy = dict(
        q=round(equipe["p"][0], 1), pa=round(equipe["a"][0]), pd=round(equipe["d"][0]),
        lj=round(statistics.mean([f["lj"] for f in cl if f["lj"]])) if cl else 0,
        n=len(cl), nc=len(groupes[False]), nn=len(groupes[True]), at=attr_moy,
    )

    # badges (règles objectives, identiques pour tous)
    for f in fiches:
        ok_vol = f["h_pick"] >= BADGE_HEURES
        b = []
        if f["q"] is not None and f["q"] >= 100 and ok_vol: b.append("c100")
        if f["q"] is not None and f["q"] >= 120 and ok_vol: b.append("turbo")
        if f["ea"] == 0 and ok_vol: b.append("zero")
        if f["pd"] is not None and f["pd"] < PPM_DEC_VERT and ok_vol: b.append("lynx")
        if f["sec"] == 0 and ok_vol: b.append("secu")
        if f["pos"] >= 1: b.append("team")
        if f["prec"] and f["prec"]["q"] and f["q"] and f["q"] - f["prec"]["q"] >= 3 and ok_vol: b.append("up")
        if f["reg"] is not None and f["reg"] >= 80 and ok_vol: b.append("reg")
        if f["lig"] >= 10000: b.append("k10")
        if f["ovr"] and f["ovr"] >= SEUIL_OR: b.append("or")
        f["badges"] = b

    # défi du mois suivant : le critère le plus faible
    for f in fiches:
        f["defi"] = defi(f, moy)

    return fiches, moy


def defi(f, moy):
    if f["sec"] or f["ec"]:
        t = next((t for (s, t), n in f["rh_types"] if s == "-"), None)
        sid, stitre = STD_LIEN.get(t, ("P18", "Sécurité au quotidien"))
        if f["ec"] >= 2 or not f["classe"]:
            return dict(k="r", t=f"Zéro écart « {t} »", d=f"{f['ec']} écart(s) ce mois. Objectif : 0 le mois prochain.", s=sid, st=stitre)
    if not f["classe"]:
        return dict(k="p", t="Prendre tes marques", d="Pas assez d'heures de picking ce mois pour être classé. On se retrouve le mois prochain.", s="P4", st="Préparation détail")
    n = f["notes"]
    if f.get("nouveau"):
        o = OBJ_NOUVEAU
        manque = [(n["p"], "p"), (n["a"], "a"), (n["d"], "d")]
        k = min(manque)[1]
        if k == "p":
            return dict(k="p", t=f"Atteindre {o['lj']} lignes / jour", d=f"Tu es à {f['lj']} lignes / jour. Ton objectif d'intégration : {o['lj']}.", s="P4", st="Préparation détail")
        if k == "a":
            return dict(k="a", t=f"PPM avérée sous {o['pa']}", d=f"Tu es à {f['pa']}. Contrôle référence + quantité avant de valider.", s="P8", st="Ruptures & validation")
        return dict(k="d", t=f"PPM décelée sous {o['pd']}", d=f"Tu es à {f['pd']}. Vérifie l'étiquette et la quantité à chaque ligne.", s="P4", st="Préparation détail")
    k = min(("p", "a", "d"), key=lambda x: n[x] if n[x] is not None else 99)
    if all((n[x] or 0) >= 13 for x in ("p", "a", "d")):
        return dict(k="ok", t="Garder le cap", d="Tu es au-dessus de la moyenne partout. Montre la voie, aide un collègue.", s="P8", st="Ruptures & validation")
    if k == "p":
        cible = max(100, round((f["q"] or 0) + 5))
        return dict(k="p", t=f"Viser {cible} % de quota", d=f"Tu es à {f['q']} % (équipe : {moy['q']} %). Gagne 5 points de rythme.", s="P4", st="Préparation détail")
    if k == "a":
        return dict(k="a", t="Zéro litige client", d=f"{f['ea']} litige(s) ce mois. Contrôle référence + quantité avant de valider.", s="P8", st="Ruptures & validation")
    return dict(k="d", t="Moins d'erreurs au rejet", d=f"{f['ed']} erreur(s) décelée(s). Vérifie l'étiquette et la quantité à chaque ligne.", s="P4", st="Préparation détail")

# ─────────────────────────── charge utile QR ───────────────────────────

def charge_utile(f, moy, mois, annee):
    d = {
        "v": 1, "m": f"{annee}-{mois:02d}",
        "p": f["prenom"], "n": f["nom"], "i": f["mat"],
        "j": f["jours"], "h": f["heures"], "x": f["hs"],
        "l": f["lig"], "lj": f["lj"], "c": f["cont"], "cd": f["cad"],
        "q": f["q"], "ed": f["ed"], "ea": f["ea"], "pd": f["pd"], "pa": f["pa"],
        "rg": f["reg"], "b": f["badges"], "bt": list(f["best"]),
        "s": [v for _, v in f["serie"]],
        "jd": f["jours_det"],   # [jour, heures, heures pick, lignes, décelées, avérées, quota]
        "pr": f["prec"],
        "no": f["note"], "ns": f["notes"] if f["classe"] else None, "ov": f["ovr"], "ca": f["carte"], "at": f["attr"],
        "nv": f["nouveau"], "ob": OBJ_NOUVEAU if f["nouveau"] else None, "bn": BONUS_NOUVEAU if f["nouveau"] else None,
        "cp": None if f["nouveau"] else round((COUP_DE_POUCE - 1) * 100),
        "rh": [f["pos"], f["ec"], f["sec"]],
        "rt": [[s, t, n] for (s, t), n in f["rh_types"][:4]],
        "er": [[t, n] for t, n in f["err_types"]],
        "av": [[t, n] for t, n in f["ave_types"]],
        "df": f["defi"],
        "T": moy,
        "cl": f["classe"], "qo": f["quota_ok"],
    }
    brut = json.dumps(d, ensure_ascii=False, separators=(",", ":")).encode()
    comp = zlib.compressobj(9, zlib.DEFLATED, -15)
    z = comp.compress(brut) + comp.flush()
    return "z" + base64.urlsafe_b64encode(z).decode().rstrip("="), d


def qr_svg(url, couleur="#0B0C0E", fond="#ffffff"):
    q = segno.make(url, error="m", micro=False)
    return q.svg_inline(scale=1, border=0, dark=couleur, light=fond, omitsize=True), q.version

# ─────────────────────────── rendu HTML / PDF ───────────────────────────

CSS_COMMUN = """
@import url('https://fonts.googleapis.com/css2?family=Archivo:wght@400;500;600;700;800;900&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
:root{--beton:#0B0C0E;--plate:#1B1E24;--craie:#F4F5F6;--acier:#98A0AC;--jaune:#FFD400;--rouge:#E5332A;--vert:#3FBF6E}
*{margin:0;padding:0;box-sizing:border-box}
html,body{font-family:'Archivo',Inter,Arial,sans-serif;-webkit-print-color-adjust:exact;print-color-adjust:exact}
.mono{font-family:'IBM Plex Mono',monospace}
"""


def nom_fichier(f):
    return f"{f['mat'].zfill(3)}_{f['nom'].lower().replace(' ', '-')}.html"


def page_perso(f, moy, mois, annee):
    """Page autonome : la fiche.html avec les données intégrées (marche sur tous les téléphones)."""
    _, d = charge_utile(f, moy, mois, annee)
    gab = GABARIT.read_text(encoding="utf-8")
    inj = ("<script>window.BILAN=" + json.dumps(d, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
           + ";window.STD_URL=" + json.dumps(BASE_URL + "index.html") + ";</script>\n<script>")
    assert '<script>\n"use strict";' in gab
    gab = gab.replace('<script>\n"use strict";', inj + '\n"use strict";', 1)
    return gab.replace("<title>Mon bilan Pick &amp; Pack</title>",
                       f"<title>Bilan {MOIS_FR[mois].lower()} — {html.escape(f['prenom'])}</title>", 1)


def html_cartes(fiches, moy, mois, annee, pages=None):
    lib = f"{MOIS_FR[mois]} {annee}"
    std_svg, _ = qr_svg(BASE_URL)
    cartes = []
    for f in sorted(fiches, key=lambda f: (f["nom"], f["prenom"])):
        fn = nom_fichier(f)
        if pages is not None:
            pages[fn] = page_perso(f, moy, mois, annee)
        url = f"{SITE_BILAN}{MOIS_DOSSIER[mois]}/{fn}"
        svg, ver = qr_svg(url)
        f["_url"], f["_qrv"] = url, ver
        nb = len(f["badges"])
        cartes.append(f"""
<div class="carte recto">
  <div class="bande"><span class="mono">CLR CHAMPIGNEULLES · PICK &amp; PACK</span><span class="mono">{html.escape(lib)}</span></div>
  <div class="hachures"></div>
  <div class="corps">
    <div class="eyebrow mono">TON BILAN DU MOIS</div>
    <div class="prenom">{html.escape(f['prenom'])}</div>
    <div class="nom">{html.escape(f['nom'])}</div>
    <div class="plaque mono">MAT. {html.escape(f['mat'])}</div>
    <div class="qrbox"><div class="coin c1"></div><div class="coin c2"></div><div class="coin c3"></div><div class="coin c4"></div>{svg}</div>
    <div class="cta">Scanne&nbsp;: ta carte joueur, tes&nbsp;badges</div>
    <ul class="teaser mono"><li><b>/20</b>note</li><li><b>OVR</b>niveau</li><li><b>{nb}</b>badge{'s' if nb > 1 else ''}</li><li><b>1</b>défi</li></ul>
  </div>
  <div class="pied mono"><span>DOCUMENT PERSONNEL — NE PAS PARTAGER</span></div>
</div>""")
    verso = f"""
<div class="carte verso">
  <div class="hachures haut"></div>
  <div class="corps">
    <div class="eyebrow mono">CLR CHAMPIGNEULLES</div>
    <div class="titre-std">LES 18<br>STANDARDS</div>
    <div class="sous mono">PICK &amp; PACK · P1 → P18</div>
    <div class="qrbox jaune"><div class="coin c1"></div><div class="coin c2"></div><div class="coin c3"></div><div class="coin c4"></div>{std_svg}</div>
    <div class="cta">Scanne&nbsp;: standards, quiz &amp; badges</div>
    <div class="chips mono"><span>SÉCURITÉ</span><span>QUALITÉ</span><span>PROCESS</span><span>ENGIN</span></div>
  </div>
  <div class="hachures bas"></div>
</div>"""
    pages = []
    for i in range(0, len(cartes), 4):
        lot = cartes[i:i + 4]
        n = len(lot)
        lot += ['<div class="carte vide"></div>'] * (4 - n)
        pages.append('<section class="feuille">' + "".join(lot) + '<i class="trait v"></i><i class="trait h"></i></section>')
        # verso : même grille ; en recto-verso bord long, la colonne gauche devient la droite
        vs = [verso if k < n else '<div class="carte vide"></div>' for k in range(4)]
        vs = [vs[1], vs[0], vs[3], vs[2]]
        pages.append('<section class="feuille">' + "".join(vs) + '<i class="trait v"></i><i class="trait h"></i></section>')

    css = CSS_COMMUN + """
@page{size:A4;margin:0}
body{background:#fff}
.feuille{width:210mm;height:297mm;display:grid;grid-template-columns:105mm 105mm;grid-template-rows:148.5mm 148.5mm;position:relative;page-break-after:always;overflow:hidden}
.trait{position:absolute;background:transparent;border:0 dashed #c9ccd1}
.trait.v{left:105mm;top:0;bottom:0;border-left-width:.2mm}
.trait.h{top:148.5mm;left:0;right:0;border-top-width:.2mm}
.carte{position:relative;overflow:hidden;margin:4mm;border-radius:4mm;display:flex;flex-direction:column}
.carte.vide{background:none}
.recto{background:#0B0C0E;color:#F4F5F6}
.recto::before{content:'';position:absolute;inset:0;background:radial-gradient(120% 70% at 100% 0%,rgba(255,212,0,.16),transparent 60%),radial-gradient(90% 60% at 0% 100%,rgba(76,154,255,.10),transparent 60%)}
.bande{position:relative;display:flex;justify-content:space-between;align-items:center;background:#FFD400;color:#0B0C0E;padding:2.6mm 4mm;font-size:6.4pt;font-weight:600;letter-spacing:.08em}
.hachures{position:relative;height:2.2mm;background:repeating-linear-gradient(-45deg,#FFD400 0 2mm,#0B0C0E 2mm 4mm)}
.corps{position:relative;flex:1;display:flex;flex-direction:column;align-items:center;padding:4mm 6mm 0;text-align:center}
.eyebrow{font-size:6.6pt;letter-spacing:.22em;color:#FFD400;font-weight:600}
.prenom{font-size:21pt;font-weight:900;letter-spacing:-.01em;line-height:1;margin-top:2.2mm}
.nom{font-size:10.5pt;font-weight:700;letter-spacing:.16em;color:#98A0AC;margin-top:1mm}
.plaque{margin-top:2mm;font-size:6.6pt;letter-spacing:.14em;border:.25mm solid rgba(255,255,255,.25);padding:.8mm 2.4mm;border-radius:1mm;color:#F4F5F6}
.qrbox{position:relative;margin-top:4.5mm;width:62mm;height:62mm;background:#fff;border-radius:2.5mm;padding:3.2mm}
.qrbox svg{width:100%;height:100%;display:block}
.qrbox .coin{position:absolute;width:5mm;height:5mm;border:0 solid #FFD400}
.c1{top:-1.6mm;left:-1.6mm;border-top-width:.9mm!important;border-left-width:.9mm!important}
.c2{top:-1.6mm;right:-1.6mm;border-top-width:.9mm!important;border-right-width:.9mm!important}
.c3{bottom:-1.6mm;left:-1.6mm;border-bottom-width:.9mm!important;border-left-width:.9mm!important}
.c4{bottom:-1.6mm;right:-1.6mm;border-bottom-width:.9mm!important;border-right-width:.9mm!important}
.cta{margin-top:4.5mm;font-size:9.5pt;font-weight:700}
.teaser{display:flex;gap:1.6mm;margin-top:2.6mm;list-style:none}
.teaser li{display:flex;flex-direction:column;align-items:center;min-width:13mm;padding:1.2mm 1mm;border-radius:1.4mm;background:rgba(255,255,255,.06);border:.2mm solid rgba(255,255,255,.12);font-size:5.4pt;letter-spacing:.12em;color:#98A0AC;text-transform:uppercase}
.teaser b{font-family:'Archivo',sans-serif;font-size:10pt;color:#FFD400;font-weight:800;letter-spacing:0}
.pied{position:relative;text-align:center;font-size:5.4pt;letter-spacing:.16em;color:#5E6672;padding:2.6mm 0 3mm}
.verso{background:#FFD400;color:#0B0C0E}
.verso .hachures{height:6mm;background:repeating-linear-gradient(-45deg,#0B0C0E 0 3mm,#FFD400 3mm 6mm)}
.verso .corps{justify-content:center;padding-top:2mm}
.verso .eyebrow{color:#0B0C0E}
.titre-std{font-size:27pt;font-weight:900;line-height:.92;letter-spacing:-.02em;margin-top:2mm}
.sous{font-size:7pt;letter-spacing:.2em;margin-top:2mm;font-weight:600}
.qrbox.jaune{background:#fff;margin-top:5mm;width:54mm;height:54mm}
.qrbox.jaune .coin{border-color:#0B0C0E}
.chips{display:flex;gap:1.4mm;margin-top:3mm}
.chips span{font-size:5.6pt;letter-spacing:.12em;background:#0B0C0E;color:#FFD400;padding:1mm 1.8mm;border-radius:1mm;font-weight:600}
"""
    return f"<!doctype html><html lang='fr'><head><meta charset='utf-8'><title>QR opérateurs — {lib}</title><style>{css}</style></head><body>{''.join(pages)}</body></html>"


def html_synthese(fiches, moy, mois, annee):
    lib = f"{MOIS_FR[mois]} {annee}"
    cl = sorted([f for f in fiches if f["classe"] and not f["nouveau"]], key=lambda f: (f["nom"], f["prenom"]))
    nv = sorted([f for f in fiches if f["classe"] and f["nouveau"]], key=lambda f: (f["nom"], f["prenom"]))
    nc = sorted([f for f in fiches if not f["classe"]], key=lambda f: f["nom"])

    def cell_note(v):
        if v is None:
            return "<td class='n'>—</td>"
        c = "hi" if v >= 13 else "lo" if v < 7 else "mid"
        return f"<td class='n {c}'>{v:.1f}</td>"

    def delta(f):
        if not f["prec"] or f["prec"]["q"] is None or f["q"] is None:
            return "<td class='d'>—</td>"
        d = f["q"] - f["prec"]["q"]
        return f"<td class='d {'up' if d >= 0 else 'dn'}'>{'+' if d >= 0 else ''}{d:.1f}</td>"

    def rangee(f, nouveau=False):
        med = f["carte"] or ""
        ajust = f"+{BONUS_NOUVEAU:.0f} bonus" if nouveau else f"×{COUP_DE_POUCE:.2f}".replace(".", ",")
        return (
            f"<tr><td class='rk {med}'>{f['ovr']}</td><td class='np'><b>{html.escape(f['nom'])}</b> {html.escape(f['prenom'])}<span class='mono'>{html.escape(f['mat'])}</span></td>"
            f"<td class='g'><b>{f['note']:.1f}</b><i style='width:{f['note']*5:.0f}%'></i><small>{ajust}</small></td>"
            + cell_note(f["notes"]["p"]) + cell_note(f["notes"]["d"]) + cell_note(f["notes"]["a"]) + cell_note(f["notes"]["r"])
            + f"<td>{f['q']:.1f}</td>{delta(f)}<td>{f['lig']:,}</td><td>{f['pa']}</td><td>{f['pd']}</td><td>{f['pos']} / {f['ec']}</td><td>{len(f['badges'])}</td></tr>".replace(",", " ")
        )
    lignes = [rangee(f) for f in cl]
    lignes_nv = [rangee(f, True) for f in nv]

    def ligne_nc(f):
        raison = "hors quota" if not f["quota_ok"] else f"{f['h_pick']:.0f} h de picking (&lt; {MIN_HEURES:.0f} h)"
        q = "" if f["q"] is None else f"{f['q']:.1f}"
        return (f"<tr><td class='rk'>·</td><td class='np'><b>{html.escape(f['nom'])}</b> {html.escape(f['prenom'])}"
                f"<span class='mono'>{html.escape(f['mat'])}</span></td><td colspan='5' class='raison'>{raison}</td>"
                f"<td>{q}</td><td></td><td>{f['lig']}</td><td>{f['pa']}</td><td>{f['pd']}</td>"
                f"<td>{f['pos']} / {f['ec']}</td><td>{len(f['badges'])}</td></tr>")
    lnc = "".join(ligne_nc(f) for f in nc)
    css = CSS_COMMUN + """
@page{size:A4 landscape;margin:8mm 10mm}
body{color:#0B0C0E;font-size:7.8pt;zoom:.92}
header{display:flex;justify-content:space-between;align-items:flex-end;border-bottom:1mm solid #0B0C0E;padding-bottom:3mm}
h1{font-size:20pt;font-weight:900;letter-spacing:-.01em}
h1 em{font-style:normal;background:#FFD400;padding:0 2mm}
.conf{font-size:7pt;letter-spacing:.18em;color:#E5332A;font-weight:700;text-align:right}
.kpis{display:flex;gap:3mm;margin:3mm 0}
.kpi{flex:1;border:.3mm solid #d7dade;border-radius:2mm;padding:2.5mm 3mm}
.kpi span{display:block;font-size:6.4pt;letter-spacing:.14em;color:#5E6672}
.kpi b{font-size:15pt;font-weight:800}
table{width:100%;border-collapse:collapse}
th{font-size:6.2pt;letter-spacing:.1em;text-align:right;color:#5E6672;font-weight:600;padding:1.6mm 1.4mm;border-bottom:.3mm solid #0B0C0E}
th.l,td.np{text-align:left}
td{padding:1.05mm 1.4mm;text-align:right;border-bottom:.2mm solid #eceef0;font-variant-numeric:tabular-nums}
td.np span{margin-left:2mm;color:#98A0AC;font-size:6.6pt}
td.rk{text-align:center;font-weight:800;width:7mm}
td.rk.icone{background:#0B0C0E;color:#FFD400}td.rk.or{background:#FFD400}td.rk.argent{background:#D9DCE1}td.rk.bronze{background:#E8C39E}
td.g{position:relative;width:26mm;text-align:left}
td.g b{position:relative;z-index:1;font-size:9.5pt}
td.g small{position:relative;z-index:1;margin-left:2mm;font-size:6pt;color:#5E6672}
td.g i{position:absolute;left:1.4mm;top:50%;height:2.4mm;transform:translateY(-50%);background:rgba(255,212,0,.55);border-radius:.6mm;max-width:24mm}
td.n.hi{color:#1d8a4a;font-weight:700}td.n.lo{color:#E5332A;font-weight:700}
td.d.up{color:#1d8a4a}td.d.dn{color:#E5332A}
td.raison{text-align:left;color:#98A0AC;font-style:italic}
tr.sep td{background:#f4f5f6;text-align:left;font-size:6.4pt;letter-spacing:.14em;color:#5E6672;font-weight:700}
.methode{margin-top:4mm;display:grid;grid-template-columns:1.2fr 1fr;gap:5mm;font-size:7pt;color:#3a3f47;line-height:1.45}
.methode h3{font-size:7pt;letter-spacing:.16em;margin-bottom:1mm}
"""
    tete = ("<tr><th>OVR</th><th class='l'>OPÉRATEUR</th><th class='l'>NOTE /20</th><th>PROD. 40%</th><th>DÉCELÉE 20%</th>"
            "<th>AVÉRÉE 20%</th><th>RH & SAVOIR-ÊTRE 20%</th><th>QUOTA %</th><th>Δ M-1</th><th>LIGNES</th><th>PPM AVÉRÉE</th>"
            "<th>PPM DÉCELÉE</th><th>RH + / −</th><th>BADGES</th></tr>")
    tot_lig = sum(f["lig"] for f in fiches)
    return f"""<!doctype html><html lang='fr'><head><meta charset='utf-8'><title>Synthèse superviseur — {lib}</title><style>{css}</style></head><body>
<header><div><div class='mono' style='font-size:7pt;letter-spacing:.2em;color:#5E6672'>CLR CHAMPIGNEULLES · PICK &amp; PACK · SYNTHÈSE MENSUELLE</div>
<h1>Synthèse opérateurs — <em>{lib}</em></h1></div>
<div class='conf mono'>CONFIDENTIEL — SUPERVISEUR / RH<br><span style='color:#5E6672;font-weight:500'>généré le {dt.date.today():%d/%m/%Y}</span></div></header>
<div class='kpis'>
<div class='kpi'><span class='mono'>NOTÉS · CONFIRMÉS + NOUVEAUX</span><b>{moy['nc']} + {moy['nn']}</b> / {len(fiches)}</div>
<div class='kpi'><span class='mono'>QUOTA MOYEN ÉQUIPE</span><b>{moy['q']} %</b></div>
<div class='kpi'><span class='mono'>PPM AVÉRÉE MOYENNE</span><b>{moy['pa']}</b></div>
<div class='kpi'><span class='mono'>PPM DÉCELÉE MOYENNE</span><b>{moy['pd']}</b></div>
<div class='kpi'><span class='mono'>LIGNES PRÉPARÉES</span><b>{tot_lig:,}</b></div>
</div>
<table>{tete}{''.join(lignes)}<tr class='sep'><td colspan='14'>NOUVEAUX — PARCOURS D'INTÉGRATION · notés sur leur objectif : {OBJ_NOUVEAU['lj']} lignes/jour · PPM avérée {OBJ_NOUVEAU['pa']} · PPM décelée {OBJ_NOUVEAU['pd']:,} · bonus nouveau +{BONUS_NOUVEAU:.0f} pts</td></tr>{''.join(lignes_nv)}<tr class='sep'><td colspan='14'>NON NOTÉS — VOLUME INSUFFISANT OU HORS QUOTA (fiche personnelle quand même remise)</td></tr>{lnc}</table>
<div class='methode'>
<div><h3 class='mono'>MÉTHODE — IDENTIQUE POUR TOUS</h3>
Chaque critère est noté /20 selon l'écart à la moyenne de l'équipe classée : <b>note = 10 + 4 × (valeur − moyenne) / écart-type</b>, bornée de 0 à 20.
Pile dans la moyenne = 10/20. Pour les PPM, plus bas = meilleure note. Note finale = 40 % productivité (quota) + 20 % qualité décelée (PPM décelée)
+ 20 % qualité avérée (PPM avérée) + 20 % RH &amp; savoir-être (points positifs − écarts du journal SUIVI_RH, ramené à 20 jours travaillés).
Équipe confirmée : note finale × {COUP_DE_POUCE:.2f} (10 % moins sévère), plafonnée à 20.
<b>Nouveaux</b> ({', '.join(sorted(n.split()[0].capitalize() for n in NOUVEAUX))}) : prod., avérée et décelée notées sur leur objectif
({OBJ_NOUVEAU['lj']} l/j, PPM avérée {OBJ_NOUVEAU['pa']}, PPM décelée {OBJ_NOUVEAU['pd']}) — objectif atteint = 12/20 — puis +{BONUS_NOUVEAU:.0f} pts de bonus.
Ordre alphabétique : aucun classement affiché. OVR (niveau carte joueur) = 45 + 2,5 × note /20 ; bronze &lt; 65 · argent 65-77 · or 78-84 · icône ≥ 85.</div>
<div><h3 class='mono'>CE QUI N'EST JAMAIS PRIS EN COMPTE</h3>
Absences, maladie, accident du travail, heures de délégation, congés, formation, type de contrat (CDI / intérim), ancienneté.
Population classée : actifs hors pilotes, soumis au quota, ≥ {MIN_HEURES:.0f} h de picking dans le mois. Source : Suivi_Op_PP_OFFICIEL, onglets PROD, SUIVI_RH, POINTAGE.</div>
</div></body></html>""".replace(f"{tot_lig:,}", f"{tot_lig:,}".replace(",", " "))

# ─────────────────────────── main ───────────────────────────

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("classeur", type=Path)
    ap.add_argument("--mois", type=int, required=True)
    ap.add_argument("--annee", type=int, default=dt.date.today().year)
    ap.add_argument("--sortie", type=Path, default=Path("bilan"))
    a = ap.parse_args()
    a.sortie.mkdir(parents=True, exist_ok=True)
    fiches, moy = construire(a.classeur, a.mois, a.annee)
    tag = f"{MOIS_COURT[a.mois]}{a.annee}"
    pages = {}
    (a.sortie / f"cartes_{tag}.html").write_text(html_cartes(fiches, moy, a.mois, a.annee, pages), encoding="utf-8")
    dossier = a.sortie / MOIS_DOSSIER[a.mois]
    dossier.mkdir(exist_ok=True)
    for fn, contenu in pages.items():
        (dossier / fn).write_text(contenu, encoding="utf-8")
    print(f"{len(pages)} pages perso -> {dossier}  (à publier dans {SITE_BILAN}{MOIS_DOSSIER[a.mois]}/)")
    (a.sortie / f"synthese_{tag}.html").write_text(html_synthese(fiches, moy, a.mois, a.annee), encoding="utf-8")
    export = [{k: v for k, v in f.items() if k != "serie"} for f in fiches]
    (a.sortie / f"donnees_{tag}.json").write_text(json.dumps({"moyennes": moy, "fiches": export}, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    print(f"{len(fiches)} fiches · {moy['n']} classés · quota moyen {moy['q']} % · PPM avérée moy. {moy['pa']}")
    for f in sorted(fiches, key=lambda f: (f["rang"] or 999, f["nom"])):
        print(f"  {str(f['rang'] or '-'):>3}  {f['nom']:<18} {f['prenom']:<14} note {f['note'] if f['note'] is not None else '—':>5}  "
              f"q {f['q']!s:>6}  pa {f['pa']!s:>5}  pd {f['pd']!s:>5}  rh {f['pos']}/{f['ec']}  badges {len(f['badges'])}  QR v{f['_qrv']}  url {len(f['_url'])} car.")


if __name__ == "__main__":
    main()
