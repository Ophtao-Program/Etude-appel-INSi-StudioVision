#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
correction_nom_jeune_fille.py — Rejoue l'appel INSi pour les patientes en echec,
en substituant temporairement leur NOM par leur NOM DE JEUNE FILLE.

HYPOTHESE TESTEE
----------------
Le teleservice INSi interroge le referentiel national sur le NOM DE NAISSANCE.
Or StudioVision stocke dans « NOM » le nom legal (souvent le nom marital pour les
femmes mariees), et le nom de naissance dans « NomJeuneFille ». Cela expliquerait
l'ecart observe : ~98 % de reussite chez les hommes contre ~59 % chez les femmes.

CE QUE FAIT LE PROGRAMME
------------------------
  1. lit etude_insi_resultats.json ;
  2. selectionne les PATIENTES (sexe F) dont l'appel a ECHOUE (01, et 02 en option) ;
  3. pour chacune, si « NomJeuneFille » est renseigne :
        - memorise le NOM actuel ;
        - ecrit patients.NOM = NomJeuneFille ;
        - recharge la fiche et VERIFIE que le nom affiche est bien le nom de jeune fille ;
        - ouvre CARACTERISTIQUES PATIENT et appelle INSi ;
        - RESTAURE le NOM d'origine (systematiquement, meme en cas d'erreur) ;
  4. ecrit les resultats dans etude_nom_jeune_fille_resultats.json (+ CSV).

SECURITE
--------
  * Le NOM d'origine est restaure dans un bloc « finally » : erreur, exception ou
    Ctrl+C, la fiche est toujours remise en etat.
  * Un fichier sentinelle (_nom_en_cours.json) memorise toute modification en cours.
    Si le poste plante (Access sature, coupure), le programme RESTAURE le nom au
    lancement suivant, avant toute autre operation.
  * Par defaut AUCUNE correction n'est conservee : le programme MESURE, il ne modifie
    pas durablement les fiches. L'option --conserver-si-trouve permet, a l'inverse,
    de garder le nom de jeune fille quand le teleservice confirme l'identite.

USAGE
-----
    python -u correction_nom_jeune_fille.py
    python -u correction_nom_jeune_fille.py --limit 20        # valider sur un petit lot
    python -u correction_nom_jeune_fille.py --inclure-02      # rejouer aussi les 02
    python -u correction_nom_jeune_fille.py --conserver-si-trouve
    python -u correction_nom_jeune_fille.py --fichier autre.json

Prerequis : Windows, StudioVision ouvert avec une fiche patient affichee, pywin32.
"""

import os
import sys
import csv
import json
import time
import argparse

try:
    sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
    sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
except Exception:
    pass

import etude_insi as E   # connexion COM, navigation, appel INSi

# ── Configuration ────────────────────────────────────────────────────────────
CHAMP_NOM_TABLE  = "NOM"              # nom legal dans la table patients
CHAMP_NJF        = "NomJeuneFille"    # nom de naissance (None si absent)

OUT_RESULTATS = "etude_nom_jeune_fille_resultats.json"
OUT_CSV       = "etude_nom_jeune_fille.csv"
SENTINELLE    = "_nom_en_cours.json"  # trace d'une modification non restauree

ECHECS = ["01_non_trouve"]            # + "02_plusieurs" avec --inclure-02


def out(msg=""):
    print(msg, flush=True)


# ── Acces a la table patients ────────────────────────────────────────────────
def _recordset(acc, code, champs):
    db = E._db(acc)
    cols = ", ".join("[%s]" % c for c in champs)
    sql = ("SELECT %s FROM %s WHERE [%s] = %s"
           % (cols, E.TABLE_PATIENTS, E.CTRL_CODE, E._valeur_sql(E._norm_code(code))))
    return db.OpenRecordset(sql)


def lire_champs(acc, code, champs):
    """-> dict {champ: valeur ('' si NULL)} ou None si patient absent."""
    try:
        rs = _recordset(acc, code, champs)
    except Exception as e:
        E.log("lecture %s : %s" % (champs, e))
        return None
    try:
        if rs.EOF:
            rs.Close()
            return None
        res = {}
        for c in champs:
            v = rs.Fields(c).Value
            res[c] = "" if v is None else str(v).strip()
        rs.Close()
        return res
    except Exception as e:
        try:
            rs.Close()
        except Exception:
            pass
        E.log("lecture %s : %s" % (champs, e))
        return None


def ecrire_nom(acc, code, valeur):
    """Ecrit patients.NOM = valeur. -> (True, None) | (False, err)."""
    try:
        rs = _recordset(acc, code, [CHAMP_NOM_TABLE])
    except Exception as e:
        return False, "OpenRecordset: %s" % e
    try:
        if rs.EOF:
            rs.Close()
            return False, "patient %s absent de la table" % code
        rs.Edit()
        rs.Fields(CHAMP_NOM_TABLE).Value = valeur
        rs.Update()
        rs.Close()
        return True, None
    except Exception as e:
        try:
            rs.Close()
        except Exception:
            pass
        return False, "ecriture NOM: %s" % e


def recharger_fiche(acc, code):
    """Recharge fPATIENTS et FORCE un Requery (indispensable : si la fiche affiche
    deja ce patient, ouvrir_fiche_directe ne rafraichit pas les donnees)."""
    nav = E.ouvrir_fiche_directe(acc, code)
    f = E._form(acc, E.FORM_PATIENT)
    if f is not None:
        try:
            f.Requery()
            time.sleep(0.3)
        except Exception as e:
            E.log("Requery: %s" % e)
    return nav


# ── Sentinelle : reparation apres plantage ───────────────────────────────────
def sentinelle_ecrire(code, ancien_nom):
    try:
        with open(SENTINELLE, "w", encoding="utf-8") as f:
            json.dump({"code_patient": code, "nom_origine": ancien_nom,
                       "horodatage": time.strftime("%Y-%m-%d %H:%M:%S")}, f,
                      ensure_ascii=False, indent=2)
    except Exception:
        pass


def sentinelle_effacer():
    try:
        if os.path.exists(SENTINELLE):
            os.remove(SENTINELLE)
    except Exception:
        pass


def sentinelle_reparer(acc):
    """Si une execution precedente a ete interrompue avant restauration, remet le nom."""
    if not os.path.exists(SENTINELLE):
        return
    try:
        with open(SENTINELLE, encoding="utf-8") as f:
            s = json.load(f)
    except Exception:
        sentinelle_effacer()
        return
    code, nom = s.get("code_patient"), s.get("nom_origine")
    if not code:
        sentinelle_effacer()
        return
    out("")
    out("!! Une execution precedente s'est interrompue en cours de modification.")
    out("   Patient %s : restauration du nom d'origine « %s »..." % (code, nom))
    ok, err = ecrire_nom(acc, code, nom if nom else None)
    if ok:
        out("   -> nom restaure.")
        sentinelle_effacer()
    else:
        out("   -> ECHEC de la restauration : %s" % err)
        out("   Corrigez manuellement la fiche %s avant de continuer." % code)
        raise SystemExit(2)
    out("")


# ── Selection des patientes en echec ─────────────────────────────────────────
def est_femme(e):
    for v in (e.get("sexe_fiche"), e.get("sexe_corrige"), e.get("sexe")):
        s = str(v or "").strip().upper()
        if s in ("2", "F"):
            return True
        if s in ("1", "M"):
            return False
    return False


def ecrire_csv(chemin, entetes, lignes):
    with open(chemin, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(entetes)
        w.writerows(lignes)


# ── Traitement d'une patiente ────────────────────────────────────────────────
def traiter(acc, code, verbeux=True):
    """-> dict resultat. Le NOM d'origine est TOUJOURS restaure."""
    r = {"code_patient": code, "nom_legal": "", "nom_jeune_fille": "", "prenom": "",
         "date_naissance": "", "classification": "", "code_insi": "",
         "ins_present": False, "nom_restaure": True, "nom_conserve": False, "detail": ""}

    champs = lire_champs(acc, code, [CHAMP_NOM_TABLE, CHAMP_NJF])
    if champs is None:
        r["classification"] = "patient_absent"
        r["detail"] = "code introuvable dans la table patients"
        return r

    nom_legal = champs[CHAMP_NOM_TABLE]
    njf = champs[CHAMP_NJF]
    r["nom_legal"] = nom_legal
    r["nom_jeune_fille"] = njf

    if not njf:
        r["classification"] = "pas_de_nom_jeune_fille"
        r["detail"] = "champ NomJeuneFille vide"
        return r
    if njf.strip().upper() == nom_legal.strip().upper():
        r["classification"] = "nom_identique"
        r["detail"] = "le nom de jeune fille est identique au nom legal"
        return r

    modifie = False
    try:
        # 1) substitution du nom
        sentinelle_ecrire(code, nom_legal)
        ok, err = ecrire_nom(acc, code, njf)
        if not ok:
            sentinelle_effacer()
            r["classification"] = "nom_non_ecrit"
            r["detail"] = err
            return r
        modifie = True

        # 2) rechargement + VERIFICATION que la fiche affiche le nom de jeune fille
        nav = recharger_fiche(acc, code)
        if not nav["charge"]:
            r["classification"] = "navigation_echec"
            r["detail"] = nav["detail"]
            return r
        affiche = E._ctrl_value(acc, E.FORM_PATIENT, E.CTRL_NOM)
        if affiche.strip().upper() != njf.strip().upper():
            r["classification"] = "substitution_non_effective"
            r["detail"] = "la fiche affiche « %s » au lieu de « %s »" % (affiche, njf)
            return r
        r["prenom"] = nav["prenom"]
        r["date_naissance"] = nav["ddn"]
        if verbeux:
            out("   %s -> %s  (%s, née %s)" % (nom_legal, njf, nav["prenom"], nav["ddn"] or "?"))

        # 3) appel INSi
        oks, errs = E.ouvrir_sous_formulaire(acc)
        if not oks:
            r["classification"] = "sous_formulaire_absent"
            r["detail"] = errs
            return r
        res = E.appel_insi(acc)
        E.fermer_sous_formulaire(acc)

        code_insi = (res.get("code") or "").strip()
        r["code_insi"] = code_insi
        r["ins_present"] = bool(res.get("ins_present"))
        if code_insi == "00":
            r["classification"] = "00_trouve_avec_njf"
        elif code_insi == "02":
            r["classification"] = "02_plusieurs_avec_njf"
        elif code_insi == "01":
            r["classification"] = "01_toujours_non_trouve"
        else:
            r["classification"] = res.get("classification") or "sans_reponse"
        r["detail"] = res.get("detail", "")
        return r

    finally:
        # 4) RESTAURATION systematique (sauf conservation explicite d'un succes)
        if modifie:
            garder = (CONSERVER_SI_TROUVE
                      and r.get("classification") in ("00_trouve_avec_njf",
                                                      "02_plusieurs_avec_njf"))
            if garder:
                r["nom_conserve"] = True
                r["nom_restaure"] = False
                sentinelle_effacer()
                if verbeux:
                    out("   nom de jeune fille CONSERVE dans la fiche (option active)")
            else:
                okr, errr = ecrire_nom(acc, code, nom_legal)
                r["nom_restaure"] = bool(okr)
                if okr:
                    sentinelle_effacer()
                else:
                    r["detail"] = (r.get("detail", "") + " | ECHEC RESTAURATION : %s" % errr).strip(" |")
                    out("   !! ECHEC de la restauration du nom pour %s : %s" % (code, errr))
            try:
                recharger_fiche(acc, code)
            except Exception:
                pass


# ── Programme principal ──────────────────────────────────────────────────────
CONSERVER_SI_TROUVE = False


def main(argv):
    global CONSERVER_SI_TROUVE
    p = argparse.ArgumentParser(
        description="Rejoue l'appel INSi des patientes en echec avec leur nom de jeune fille.")
    p.add_argument("--fichier", default=E.OUT_RESULTATS, help="resultats de l'etude INSi")
    p.add_argument("--limit", type=int, default=None, help="ne traiter que les N premieres")
    p.add_argument("--inclure-02", dest="inclure_02", action="store_true",
                   help="rejouer aussi les reponses 02 (plusieurs identites)")
    p.add_argument("--conserver-si-trouve", dest="conserver", action="store_true",
                   help="conserver le nom de jeune fille dans la fiche si l'appel reussit")
    args = p.parse_args(argv)
    CONSERVER_SI_TROUVE = args.conserver

    if not E.HAS_WIN:
        out("Programme reserve a Windows (StudioVision + pywin32).")
        return 2
    if not os.path.exists(args.fichier):
        out("Fichier introuvable : %s" % os.path.abspath(args.fichier))
        return 2
    with open(args.fichier, encoding="utf-8") as f:
        try:
            data = json.load(f)
        except Exception as e:
            out("JSON illisible : %s" % e)
            return 2
    if not isinstance(data, list):
        out("Format inattendu (liste attendue).")
        return 2

    echecs = list(ECHECS) + (["02_plusieurs"] if args.inclure_02 else [])
    cibles = [e for e in data
              if str(e.get("classification", "")).strip() in echecs and est_femme(e)]
    # dedoublonnage par code
    vus, liste = set(), []
    for e in cibles:
        c = E._norm_code(e.get("code_patient"))
        if c not in vus:
            vus.add(c)
            liste.append(e)
    if args.limit:
        liste = liste[:args.limit]

    out("Fichier ............... %s" % os.path.abspath(args.fichier))
    out("Echecs retenus ........ %s" % ", ".join(echecs))
    out("Patientes a rejouer ... %d" % len(liste))
    if CONSERVER_SI_TROUVE:
        out("MODE : le nom de jeune fille sera CONSERVE dans la fiche en cas de succes.")
    else:
        out("MODE : mesure seule — le nom legal est systematiquement restaure.")
    if not liste:
        out("Rien a faire.")
        return 0

    acc = E.access_app()
    if acc is None:
        out("StudioVision (Access) introuvable. Ouvrez StudioVision puis relancez.")
        return 2
    if E._form(acc, E.FORM_PATIENT) is None:
        out("Ouvrez une fiche patient (n'importe laquelle) dans StudioVision puis relancez.")
        return 2
    out("StudioVision detecte.")
    sentinelle_reparer(acc)

    rs_initial = E.lire_recordsource(acc)
    resultats = []
    t0 = time.time()
    try:
        for i, e in enumerate(liste, 1):
            code = E._norm_code(e.get("code_patient"))
            out("")
            out("[%d/%d] code=%s  (%s)" % (i, len(liste), code, e.get("classification")))
            if not E.studiovision_vivant(acc):
                acc = E.maintenance(acc, reacquerir=True)
                if not E.studiovision_vivant(acc):
                    out("StudioVision ne repond plus — arret (avancement sauvegarde).")
                    break
            try:
                r = traiter(acc, code)
            except Exception as ex:
                r = {"code_patient": code, "classification": "exception",
                     "detail": str(ex), "nom_restaure": True}
            r["classification_avant"] = e.get("classification")
            r["horodatage"] = time.strftime("%Y-%m-%d %H:%M:%S")
            resultats.append(r)
            E.sauver_json(resultats, OUT_RESULTATS)
            out("   => %s%s" % (r.get("classification"),
                                ("  (%s)" % r["detail"]) if r.get("detail") else ""))
            if i % E.MAINTENANCE_TOUS == 0:
                out("   … maintenance memoire apres %d patientes" % i)
                acc = E.maintenance(acc)
            time.sleep(E.INTER_APPEL)
    except KeyboardInterrupt:
        out("")
        out("Interruption : le nom de la patiente en cours a ete restaure.")

    if E.studiovision_vivant(acc):
        E.restaurer_recordsource(acc, rs_initial)

    # ── Synthese ─────────────────────────────────────────────────────────────
    from collections import Counter
    c = Counter(r.get("classification") for r in resultats)
    n = len(resultats)
    avec_njf = sum(v for k, v in c.items()
                   if k not in ("pas_de_nom_jeune_fille", "nom_identique", "patient_absent"))
    trouves = c.get("00_trouve_avec_njf", 0)
    plusieurs = c.get("02_plusieurs_avec_njf", 0)
    non_restaures = [r for r in resultats if not r.get("nom_restaure")
                     and not r.get("nom_conserve")]

    def pc(x, tot):
        return ("%.1f" % (100.0 * x / tot)).replace(".", ",") if tot else "0,0"

    out("")
    out("=" * 66)
    out("  CORRECTION PAR LE NOM DE JEUNE FILLE")
    out("=" * 66)
    out("  Patientes traitees ............ %d" % n)
    out("  Sans nom de jeune fille ....... %d" % c.get("pas_de_nom_jeune_fille", 0))
    out("  Nom identique au nom legal .... %d" % c.get("nom_identique", 0))
    out("  Appels INSi rejoues ........... %d" % avec_njf)
    out("")
    out("  ➜ RETROUVEES avec le nom de naissance (00) ... %d  (%s %% des rejouees)"
        % (trouves, pc(trouves, avec_njf)))
    out("    Plusieurs identites (02) .................... %d" % plusieurs)
    out("    Toujours introuvables (01) .................. %d" % c.get("01_toujours_non_trouve", 0))
    autres = {k: v for k, v in c.items()
              if k not in ("00_trouve_avec_njf", "02_plusieurs_avec_njf",
                           "01_toujours_non_trouve", "pas_de_nom_jeune_fille",
                           "nom_identique")}
    if autres:
        out("")
        out("  Autres cas :")
        for k, v in sorted(autres.items(), key=lambda kv: -kv[1]):
            out("     %-32s %d" % (k, v))
    out("")
    if non_restaures:
        out("  !! %d fiche(s) dont le nom N'A PAS ete restaure — a corriger :" % len(non_restaures))
        for r in non_restaures:
            out("     code=%s  nom d'origine « %s »" % (r["code_patient"], r.get("nom_legal")))
    else:
        out("  Tous les noms legaux ont ete restaures." if not CONSERVER_SI_TROUVE
            else "  Restaurations effectuees (hors succes conserves).")
    out("")
    out("  Duree ......................... %.0f s" % (time.time() - t0))
    out("=" * 66)

    ecrire_csv(OUT_CSV,
               ["Code patient", "Nom légal", "Nom de jeune fille", "Prénom", "Naissance",
                "Avant", "Après", "Code INSi", "Nom restauré"],
               [(r.get("code_patient"), r.get("nom_legal"), r.get("nom_jeune_fille"),
                 r.get("prenom"), r.get("date_naissance"), r.get("classification_avant"),
                 r.get("classification"), r.get("code_insi"),
                 "oui" if r.get("nom_restaure") else ("conservé" if r.get("nom_conserve") else "NON"))
                for r in resultats])
    out("")
    out("  Résultats : %s" % os.path.abspath(OUT_RESULTATS))
    out("  Tableau   : %s" % os.path.abspath(OUT_CSV))
    if trouves:
        out("")
        out("  Interprétation : %s %% des patientes rejouées sont retrouvées dès lors que"
            % pc(trouves, avec_njf))
        out("  le téléservice est interrogé avec le nom de naissance. L'hypothèse du nom")
        out("  d'usage transmis à la place du nom de naissance est confortée.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
