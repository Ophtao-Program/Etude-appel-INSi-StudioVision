#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
test_resolution_sexe.py — applique la methode de resolution du sexe aux patients
laisses en « sans_reponse » (sexe absent) dans un fichier de resultats existant.

Objectif : verifier que la resolution du sexe fonctionne AVANT de relancer l'etude
complete. Le programme :
  1. lit etude_insi_resultats.json ;
  2. reprend UNIQUEMENT les entrees « classification == sans_reponse » ;
  3. pour chacune : recharge la fiche, remplit le sexe (deduit du n°SS si possible,
     sinon essais F puis M), appelle INSi, garde le sexe si un appel trouve le patient
     (00/02), sinon remet le champ sexe a vide ;
  4. MET A JOUR l'entree dans le JSON (une sauvegarde .bak est faite au prealable).

Ensuite, relancer « 5 - Etude INSi (complete).bat » : l'etude REPREND la ou elle
s'etait arretee (les patients deja traites sont ignores).

Usage : StudioVision ouvert avec une fiche patient affichee, puis :
    python -u test_resolution_sexe.py
    python -u test_resolution_sexe.py --limit 20        (ne traite que 20 cas)
    python -u test_resolution_sexe.py --fichier autre.json
"""

import os
import sys
import json
import time
import shutil
import argparse

try:
    sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
    sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
except Exception:
    pass

import etude_insi as E   # reutilise toute la logique (connexion, navigation, INSi, resolution)


def out(msg=""):
    print(msg, flush=True)


# Classifications a reprendre (le sexe etait absent) :
A_REPRENDRE = {"sans_reponse", "sexe_manquant"}


def main(argv):
    p = argparse.ArgumentParser(description="Resolution du sexe pour les patients 'sans_reponse'.")
    p.add_argument("--fichier", default=E.OUT_RESULTATS, help="fichier de resultats JSON")
    p.add_argument("--limit", type=int, default=None, help="ne traiter que les N premiers cas")
    args = p.parse_args(argv)

    if not E.HAS_WIN:
        out("Programme reserve a Windows (StudioVision + pywin32).")
        return 2
    if not os.path.exists(args.fichier):
        out("Fichier introuvable : %s" % os.path.abspath(args.fichier))
        out("Lancez d'abord l'etude (meme partielle) pour produire ce fichier.")
        return 2

    with open(args.fichier, encoding="utf-8") as f:
        try:
            results = json.load(f)
        except Exception as e:
            out("JSON illisible : %s" % e)
            return 2
    if not isinstance(results, list):
        out("Format inattendu (liste attendue).")
        return 2

    indices = [i for i, e in enumerate(results)
               if str(e.get("classification", "")).strip() in A_REPRENDRE]
    if args.limit:
        indices = indices[:args.limit]

    out("Fichier : %s" % os.path.abspath(args.fichier))
    out("%d entree(s) au total, %d a reprendre (sexe absent)." % (len(results), len(indices)))
    if not indices:
        out("Rien a faire : aucun patient 'sans_reponse'.")
        return 0

    acc = E.access_app()
    if acc is None:
        out("StudioVision (Access) introuvable. Ouvrez StudioVision puis relancez.")
        return 2
    if E._form(acc, E.FORM_PATIENT) is None:
        out("Ouvrez une fiche patient (n'importe laquelle) dans StudioVision puis relancez.")
        return 2
    out("StudioVision detecte.")

    # Sauvegarde avant modification
    bak = args.fichier + ".bak"
    try:
        shutil.copyfile(args.fichier, bak)
        out("Sauvegarde : %s" % os.path.abspath(bak))
    except Exception as e:
        out("Impossible de sauvegarder (%s) — on continue tout de meme." % e)

    rs_initial = E.lire_recordsource(acc)
    resolus = non_resolus = erreurs = 0
    t0 = time.time()
    try:
        for n, idx in enumerate(indices, 1):
            e = results[idx]
            code = e.get("code_patient")
            out("")
            out("[%d/%d] code=%s (ancien : %s)" % (n, len(indices), code, e.get("classification")))
            try:
                r = E.traiter_patient(acc, code, verbeux=True)
            except Exception as ex:
                r = {"classification": "exception", "detail": str(ex)}
            r["horodatage"] = time.strftime("%Y-%m-%d %H:%M:%S")
            e.update(r)
            cls = e.get("classification", "")
            if cls in ("00_trouve", "02_plusieurs"):
                resolus += 1
            elif cls in ("pas_sexe_enfant", "pas_sexe_pas_numSS", "sexe_non_resolu",
                         "01_non_trouve"):
                non_resolus += 1
            else:
                erreurs += 1
            out("   => %s%s" % (cls, ("  (%s)" % e.get("detail")) if e.get("detail") else ""))
            # sauvegarde incrementale
            tmp = args.fichier + ".tmp"
            with open(tmp, "w", encoding="utf-8") as fw:
                json.dump(results, fw, ensure_ascii=False, indent=2)
            os.replace(tmp, args.fichier)
            time.sleep(E.INTER_APPEL)
    except KeyboardInterrupt:
        out("\nInterruption : modifications sauvegardees.")

    E.restaurer_recordsource(acc, rs_initial)
    out("")
    out("==================== RESOLUTION DU SEXE ====================")
    out("  Repris ................. %d" % len(indices))
    out("  Resolus (00/02) ....... %d   (sexe corrige + INSi ok)" % resolus)
    out("  Non resolus ........... %d   (sexe remis vide : enfant / pas de n°SS / introuvable)" % non_resolus)
    out("  Erreurs ............... %d" % erreurs)
    out("  Fichier mis a jour .... %s" % os.path.abspath(args.fichier))
    out("  Duree ................. %.0f s" % (time.time() - t0))
    out("===========================================================")
    out("")
    out("Vous pouvez maintenant relancer « 5 - Etude INSi (complete).bat » :")
    out("l'etude reprendra la ou elle s'etait arretee.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
