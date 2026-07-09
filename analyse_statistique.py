#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
analyse_statistique.py — Statistiques completes de l'etude INSi.

Lit « etude_insi_resultats.json » (produit par etude_insi.py) et produit :
  * un rapport complet a l'ecran ;
  * etude_insi_statistiques.json   (toutes les mesures, exploitables) ;
  * etude_insi_rapport.md          (rapport lisible / partageable) ;
  * etude_insi_01_non_trouve.csv   (patients dont les traits ne concordent pas) ;
  * etude_insi_a_reprendre.csv     (incidents techniques a relancer) ;
  * etude_insi_sexe_corrige.csv    (fiches dont le sexe a ete corrige).

Aucune dependance : bibliotheque standard uniquement.

Usage :
    python -u analyse_statistique.py
    python -u analyse_statistique.py --fichier autre_resultats.json
    python -u analyse_statistique.py --total 3500      (pour le taux de couverture)
    python -u analyse_statistique.py --sans-csv
"""

import os
import sys
import csv
import json
import argparse
from collections import Counter, defaultdict
from datetime import date

try:
    sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
except Exception:
    pass

FICHIER_DEFAUT = "etude_insi_resultats.json"
OUT_STATS   = "etude_insi_statistiques.json"
OUT_RAPPORT = "etude_insi_rapport.md"
CSV_01      = "etude_insi_01_non_trouve.csv"
CSV_REPRISE = "etude_insi_a_reprendre.csv"
CSV_SEXE    = "etude_insi_sexe_corrige.csv"

# Classifications « definitives » (l'appel a abouti ou le cas est tranche)
ABOUTIS     = ("00_trouve", "01_non_trouve", "02_plusieurs")
DEFINITIVES = set(ABOUTIS) | {"pas_sexe_pas_numSS", "pas_sexe_enfant", "sexe_non_resolu"}

LIBELLES = {
    "00_trouve":          "00 — patient trouvé",
    "01_non_trouve":      "01 — patient NON trouvé (traits d'identité divergents)",
    "02_plusieurs":       "02 — plusieurs patients trouvés",
    "pas_sexe_pas_numSS": "sexe absent, pas de n°SS (non résolu)",
    "pas_sexe_enfant":    "sexe absent, enfant (non résolu)",
    "sexe_non_resolu":    "sexe absent, aucun sexe ne fonctionne",
    "sans_reponse":       "aucune réponse du téléservice",
    "sexe_manquant":      "sexe manquant (ancienne classification)",
    "navigation_echec":   "échec d'ouverture de la fiche",
    "sous_formulaire_absent": "sous-formulaire non ouvert (incident Access)",
    "erreur_appel":       "erreur lors de l'activation du bouton INSi",
    "reponse_illisible":  "fenêtre réponse illisible",
    "exception":          "exception Python",
    "sexe_non_ecrit":     "écriture du sexe impossible",
    "":                   "(non traité)",
}

TRANCHES = [(0, 17, "0–17 ans"), (18, 39, "18–39 ans"), (40, 59, "40–59 ans"),
            (60, 74, "60–74 ans"), (75, 200, "75 ans et +")]


def out(msg=""):
    print(msg, flush=True)


# ── utilitaires ──────────────────────────────────────────────────────────────
def pct(x, n):
    return (100.0 * x / n) if n else 0.0


def f1(x):
    return ("%.1f" % x).replace(".", ",")


def barre(x, n, largeur=38):
    if not n:
        return ""
    plein = int(round(largeur * x / n))
    return "█" * plein + "·" * (largeur - plein)


def norm_code(v):
    s = str(v if v is not None else "").strip()
    return s[:-2] if s.endswith(".0") else s


def parse_date(s):
    s = str(s or "")[:10]
    try:
        a, m, j = s.split("-")
        return date(int(a), int(m), int(j))
    except Exception:
        return None


def age_a(ddn, ref):
    if not ddn or not ref:
        return None
    a = ref.year - ddn.year - ((ref.month, ref.day) < (ddn.month, ddn.day))
    return a if 0 <= a <= 120 else None


def tranche(age):
    if age is None:
        return "âge inconnu"
    for lo, hi, lib in TRANCHES:
        if lo <= age <= hi:
            return lib
    return "âge inconnu"


def sexe_lisible(v):
    v = str(v or "").strip().upper()
    if v in ("1", "M"):
        return "M"
    if v in ("2", "F"):
        return "F"
    return "?"


def nb_appels(e):
    n = e.get("nb_appels_insi") or 0
    try:
        n = int(n)
    except Exception:
        n = 0
    if n == 0 and e.get("classification") in ABOUTIS:
        n = 1
    return n


def ecrire_csv(chemin, entetes, lignes):
    # utf-8-sig + ';' : ouverture directe dans Excel francais
    with open(chemin, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(entetes)
        w.writerows(lignes)


COURTS = {"00_trouve": "00 trouvé", "01_non_trouve": "01 non trouvé",
          "02_plusieurs": "02 plusieurs"}


def tableau_croise(titre, lignes_cle, entries, cle_fn, classifs, lignes_out):
    """Affiche + renvoie un croisement classification x cle_fn."""
    croise = defaultdict(Counter)
    for e in entries:
        croise[cle_fn(e)][e.get("classification", "")] += 1
    lignes_out.append("")
    lignes_out.append("  " + titre)
    entete = "    %-18s %6s" % ("", "total") + "".join("%16s" % COURTS.get(c, c)[:15]
                                                       for c in classifs)
    lignes_out.append(entete)
    lignes_out.append("    " + "-" * (24 + 16 * len(classifs)))
    for k in lignes_cle:
        c = croise.get(k)
        if not c:
            continue
        tot = sum(c.values())
        ligne = "    %-18s %6d" % (k[:18], tot)
        for cl in classifs:
            v = c.get(cl, 0)
            ligne += "%16s" % ("%d (%s%%)" % (v, f1(pct(v, tot))) if v else "–")
        lignes_out.append(ligne)
    return croise


# ── programme principal ──────────────────────────────────────────────────────
def main(argv):
    p = argparse.ArgumentParser(description="Statistiques de l'etude INSi.")
    p.add_argument("--fichier", default=FICHIER_DEFAUT, help="fichier de resultats JSON")
    p.add_argument("--total", type=int, default=None,
                   help="nombre total de patients de l'annee (taux de couverture)")
    p.add_argument("--sans-csv", dest="sans_csv", action="store_true",
                   help="ne pas produire les fichiers CSV")
    args = p.parse_args(argv)

    if not os.path.exists(args.fichier):
        out("Fichier introuvable : %s" % os.path.abspath(args.fichier))
        out("Lancez d'abord l'étude (même partielle) pour produire ce fichier.")
        return 2
    with open(args.fichier, encoding="utf-8") as f:
        try:
            data = json.load(f)
        except Exception as e:
            out("JSON illisible : %s" % e)
            return 2
    if not isinstance(data, list) or not data:
        out("Aucun enregistrement exploitable.")
        return 2

    # dédoublonnage par code patient (le dernier l'emporte)
    par_code = {}
    for e in data:
        par_code[norm_code(e.get("code_patient"))] = e
    entries = list(par_code.values())
    n_lignes, n = len(data), len(entries)
    doublons = n_lignes - n

    classif = Counter(e.get("classification", "") for e in entries)
    aboutis = sum(classif.get(c, 0) for c in ABOUTIS)
    n00, n01, n02 = (classif.get(c, 0) for c in ABOUTIS)
    a_reprendre = [e for e in entries if e.get("classification", "") not in DEFINITIVES]

    # dates & horodatages
    dcons = [parse_date(e.get("date_consultation")) for e in entries]
    dcons = [d for d in dcons if d]
    horo = sorted(str(e.get("horodatage") or "") for e in entries if e.get("horodatage"))

    R = []          # lignes du rapport (ecran + markdown)
    A = R.append

    A("=" * 72)
    A("  ÉTUDE INSi — RAPPORT STATISTIQUE")
    A("=" * 72)
    A("  Fichier ................ %s" % os.path.abspath(args.fichier))
    A("  Enregistrements ....... %d  (patients distincts : %d%s)"
      % (n_lignes, n, (", %d doublon(s) ignoré(s)" % doublons) if doublons else ""))
    if args.total:
        A("  Couverture ............ %d / %d patients de l'année  (%s %%)"
          % (n, args.total, f1(pct(n, args.total))))
    if dcons:
        A("  Consultations ......... du %s au %s" % (min(dcons).isoformat(), max(dcons).isoformat()))
    if horo:
        A("  Exécution ............. du %s au %s" % (horo[0], horo[-1]))

    # ── 1. Résultat principal ────────────────────────────────────────────────
    A("")
    A("-" * 72)
    A("  1. RÉSULTAT PRINCIPAL — appels INSi ayant reçu une réponse")
    A("-" * 72)
    A("  Appels aboutis ........ %d  (%s %% des patients analysés)" % (aboutis, f1(pct(aboutis, n))))
    A("")
    for c in ABOUTIS:
        v = classif.get(c, 0)
        A("   %-50s %5d  %6s %%" % (LIBELLES[c][:50], v, f1(pct(v, aboutis))))
        A("      %s" % barre(v, aboutis))
    A("")
    if aboutis:
        A("  ➜ TAUX DE RÉUSSITE INSi (00 / aboutis) ......... %s %%" % f1(pct(n00, aboutis)))
        A("  ➜ Traits d'identité divergents (01 / aboutis) .. %s %%" % f1(pct(n01, aboutis)))
        A("  ➜ Ambiguïté (02 / aboutis) .................... %s %%" % f1(pct(n02, aboutis)))

    # ── 2. Toutes les classifications ────────────────────────────────────────
    A("")
    A("-" * 72)
    A("  2. RÉPARTITION COMPLÈTE (tous les cas)")
    A("-" * 72)
    for c, v in classif.most_common():
        A("   %-50s %5d  %6s %%" % (LIBELLES.get(c, c)[:50], v, f1(pct(v, n))))
        A("      %s" % barre(v, n))

    # ── 3. Cas « sexe absent » ───────────────────────────────────────────────
    sexe_abs = [e for e in entries if str(e.get("sexe_source") or "") not in ("", "present")]
    corriges = [e for e in sexe_abs if str(e.get("sexe_corrige") or "").strip()]
    A("")
    A("-" * 72)
    A("  3. FICHES SANS SEXE — résolution automatique")
    A("-" * 72)
    A("  Fiches au sexe vide ... %d  (%s %% des patients analysés)" % (len(sexe_abs), f1(pct(len(sexe_abs), n))))
    if sexe_abs:
        A("  Sexe corrigé .......... %d  (%s %% des fiches sans sexe)"
          % (len(corriges), f1(pct(len(corriges), len(sexe_abs)))))
        A("  Restées sans sexe ..... %d" % (len(sexe_abs) - len(corriges)))
        A("")
        A("  Par méthode de déduction :")
        src = Counter(e.get("sexe_source") for e in sexe_abs)
        src_ok = Counter(e.get("sexe_source") for e in corriges)
        noms_src = {"numSS": "déduit du n°SS", "enfant": "enfant (essais F/M)",
                    "sans_numSS": "pas de n°SS (essais F/M)",
                    "numSS_invalide": "n°SS non exploitable (essais F/M)"}
        for s, v in src.most_common():
            ok = src_ok.get(s, 0)
            A("   %-34s %5d   résolus : %4d  (%5s %%)"
              % (noms_src.get(s, s), v, ok, f1(pct(ok, v))))
        if corriges:
            rep = Counter(sexe_lisible(e.get("sexe_corrige")) for e in corriges)
            A("")
            A("  Sexe finalement retenu : " + "  ".join("%s = %d" % (k, v) for k, v in rep.most_common()))

    # ── 4. Croisements ───────────────────────────────────────────────────────
    A("")
    A("-" * 72)
    A("  4. CROISEMENTS (sur les appels aboutis)")
    A("-" * 72)
    ab = [e for e in entries if e.get("classification") in ABOUTIS]

    tableau_croise("Par sexe de la fiche :", ["M", "F", "?"], ab,
                   lambda e: sexe_lisible(e.get("sexe_fiche")), list(ABOUTIS), R)

    def _tr(e):
        return tranche(age_a(parse_date(e.get("date_naissance_fiche")),
                             parse_date(e.get("date_consultation")) or date.today()))

    ordre_tr = [t[2] for t in TRANCHES] + ["âge inconnu"]
    tableau_croise("Par tranche d'âge :", ordre_tr, ab, _tr, list(ABOUTIS), R)

    tableau_croise("Selon la présence du n°SS :", ["n°SS présent", "n°SS absent"], ab,
                   lambda e: "n°SS présent" if e.get("num_ss_present") else "n°SS absent",
                   list(ABOUTIS), R)

    tableau_croise("Selon le titre :", ["Enfant", "(autre)"], ab,
                   lambda e: "Enfant" if str(e.get("titre") or "").strip().lower() == "enfant"
                   else "(autre)", list(ABOUTIS), R)

    # ── 5. INS obtenus ───────────────────────────────────────────────────────
    ins_ok = sum(1 for e in entries if e.get("ins_present"))
    A("")
    A("-" * 72)
    A("  5. IDENTIFIANT NATIONAL DE SANTÉ (INS)")
    A("-" * 72)
    A("  INS obtenu ............ %d  (%s %% des patients analysés%s)"
      % (ins_ok, f1(pct(ins_ok, n)),
         (", %s %% des 00" % f1(pct(ins_ok, n00))) if n00 else ""))

    # ── 6. Appels au téléservice ─────────────────────────────────────────────
    total_appels = sum(nb_appels(e) for e in entries)
    deux = sum(1 for e in entries if nb_appels(e) >= 2)
    A("")
    A("-" * 72)
    A("  6. APPELS AU TÉLÉSERVICE")
    A("-" * 72)
    A("  Appels INSi émis ...... %d" % total_appels)
    A("  Patients à 2 appels ... %d  (résolution du sexe)" % deux)

    # ── 7. Qualité / incidents ───────────────────────────────────────────────
    incoherences = [e for e in entries if "ATTENTION nom reponse" in str(e.get("detail") or "")]
    A("")
    A("-" * 72)
    A("  7. QUALITÉ DE L'EXÉCUTION")
    A("-" * 72)
    A("  Incidents à reprendre . %d  (%s %%)" % (len(a_reprendre), f1(pct(len(a_reprendre), n))))
    if a_reprendre:
        for c, v in Counter(e.get("classification", "") for e in a_reprendre).most_common():
            A("     %-46s %5d" % (LIBELLES.get(c, c), v))
        A("     ➜ relancer « 5 - Etude INSi (complete) » : ces cas seront refaits.")
    A("  Incohérences nom/fiche  %d" % len(incoherences))

    A("")
    A("=" * 72)

    # ── affichage ────────────────────────────────────────────────────────────
    for l in R:
        out(l)

    # ── fichiers produits ────────────────────────────────────────────────────
    stats = {
        "fichier_source": os.path.abspath(args.fichier),
        "enregistrements": n_lignes, "patients_distincts": n, "doublons_ignores": doublons,
        "couverture_total_annee": args.total,
        "consultations_du": min(dcons).isoformat() if dcons else None,
        "consultations_au": max(dcons).isoformat() if dcons else None,
        "appels_aboutis": aboutis,
        "trouve_00": n00, "trouve_00_pct_aboutis": round(pct(n00, aboutis), 2),
        "non_trouve_01": n01, "non_trouve_01_pct_aboutis": round(pct(n01, aboutis), 2),
        "plusieurs_02": n02, "plusieurs_02_pct_aboutis": round(pct(n02, aboutis), 2),
        "taux_reussite_insi_pct": round(pct(n00, aboutis), 2),
        "ins_obtenus": ins_ok,
        "appels_insi_total": total_appels, "patients_deux_appels": deux,
        "fiches_sans_sexe": len(sexe_abs), "sexe_corriges": len(corriges),
        "sexe_corriges_pct": round(pct(len(corriges), len(sexe_abs)), 2) if sexe_abs else 0,
        "incidents_a_reprendre": len(a_reprendre),
        "incoherences_nom": len(incoherences),
        "detail_par_classification": dict(classif),
    }
    with open(OUT_STATS, "w", encoding="utf-8") as f:
        json.dump(stats, f, ensure_ascii=False, indent=2)

    with open(OUT_RAPPORT, "w", encoding="utf-8") as f:
        f.write("# Étude INSi — rapport statistique\n\n```\n")
        f.write("\n".join(R))
        f.write("\n```\n")

    produits = [OUT_STATS, OUT_RAPPORT]
    if not args.sans_csv:
        l01 = [(e.get("code_patient"), e.get("nom_fiche"), e.get("prenom_fiche"),
                sexe_lisible(e.get("sexe_fiche")), e.get("date_naissance_fiche"),
                e.get("date_consultation"), e.get("titre"),
                "oui" if e.get("num_ss_present") else "non")
               for e in entries if e.get("classification") == "01_non_trouve"]
        if l01:
            ecrire_csv(CSV_01, ["Code patient", "Nom", "Prénom", "Sexe", "Naissance",
                                "Dernière consultation", "Titre", "n°SS présent"], l01)
            produits.append(CSV_01)
        if a_reprendre:
            ecrire_csv(CSV_REPRISE, ["Code patient", "Classification", "Détail"],
                       [(e.get("code_patient"), e.get("classification"), e.get("detail"))
                        for e in a_reprendre])
            produits.append(CSV_REPRISE)
        if corriges:
            ecrire_csv(CSV_SEXE, ["Code patient", "Nom", "Prénom", "Sexe retenu",
                                  "Méthode", "Appels INSi", "Classification"],
                       [(e.get("code_patient"), e.get("nom_fiche"), e.get("prenom_fiche"),
                         sexe_lisible(e.get("sexe_corrige")), e.get("sexe_source"),
                         nb_appels(e), e.get("classification")) for e in corriges])
            produits.append(CSV_SEXE)

    out("")
    out("  Fichiers produits :")
    for p_ in produits:
        out("    • %s" % os.path.abspath(p_))
    out("")
    if aboutis:
        out("  En une phrase : sur %d appels aboutis, %s %% des patients ont été retrouvés"
            % (aboutis, f1(pct(n00, aboutis))))
        out("  dans le référentiel INSi ; %s %% présentent des traits d'identité divergents."
            % f1(pct(n01, aboutis)))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
