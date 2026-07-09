#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
test_charger_patient.py — TEST DECISIF : peut-on changer de patient en
reecrivant le RecordSource de fPATIENTS ?  (aucun clavier, COM pur)

  RecordSource actuel : select * from patients where [Code patient] = 5182
  On le remplace par  : select * from patients where [Code patient] = <code>
  puis on verifie que la fiche affiche bien le nouveau patient.

Le patient initial est RESTAURE a la fin (sauf refus explicite).

IMPORTANT : ce fichier est en UTF-8 (PAS de declaration cp1252) — c'est ce
mauvais encodage qui faisait planter le test precedent sur « Prénom ».

Usage : StudioVision ouvert, une fiche patient affichee, puis :
    python -u test_charger_patient.py
"""

import re
import sys
import time

try:
    sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
    sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
except Exception:
    pass

FORM_PATIENT = "fPATIENTS"
CTRL_CODE = "Code patient"
CTRL_NOM = "NOM"
CTRL_PRENOM = "Prénom"


def out(msg=""):
    print(msg, flush=True)


# ── Connexion COM (robuste multi-instances, comme WebDMP Assistant) ──────────
def access_app():
    import pythoncom
    import win32com.client as win32
    pythoncom.CoInitialize()
    found = []
    try:
        rot = pythoncom.GetRunningObjectTable()
        for moniker in rot.EnumRunning():
            try:
                raw = rot.GetObject(moniker)
                app = win32.Dispatch(raw.QueryInterface(pythoncom.IID_IDispatch))
                _ = app.CurrentProject
                fc = int(app.Forms.Count)
            except Exception:
                continue
            found.append((fc, app))
    except Exception:
        pass
    if found:
        found.sort(key=lambda t: t[0], reverse=True)
        return found[0][1]
    try:
        return win32.GetActiveObject("Access.Application")
    except Exception:
        return None


def lire_ctrl(f, nom):
    """Lit un controle ; en cas de souci (accents), repli sur le Recordset."""
    try:
        v = f.Controls(nom).Value
        return "" if v is None else str(v).strip()
    except Exception:
        pass
    try:
        flds = f.Recordset.Fields
        cible = nom.strip().lower()
        for i in range(int(flds.Count)):
            if str(flds(i).Name).strip().lower() == cible:
                v = flds(i).Value
                return "" if v is None else str(v).strip()
    except Exception:
        pass
    return "(illisible)"


def lire_champ_rs(f, nom):
    try:
        v = f.Recordset.Fields(nom).Value
        return "" if v is None else str(v).strip()
    except Exception:
        return ""


def lire_label(f, nom):
    try:
        return str(f.Controls(nom).Caption)
    except Exception:
        return ""


def etat(f, titre):
    out("=" * 70)
    out(titre)
    out("=" * 70)
    try:
        out("RecordSource : %s" % f.RecordSource)
    except Exception as e:
        out("RecordSource : (illisible : %s)" % e)
    out("Code patient : %s" % lire_ctrl(f, CTRL_CODE))
    out("NOM          : %s" % lire_ctrl(f, CTRL_NOM))
    out("Prénom       : %s" % lire_ctrl(f, CTRL_PRENOM))
    out("SEXE (table) : %s" % lire_champ_rs(f, "SEXE"))
    out("Naissance    : %s" % lire_champ_rs(f, "Date de Naissance")[:10])
    for lbl in ("LblDate1", "LblDate2", "LblDate3"):
        c = lire_label(f, lbl)
        if c:
            out("%s     : %s" % (lbl, c))
    out("")


def norm(v):
    s = str(v).strip()
    return s[:-2] if s.endswith(".0") else s


def main():
    acc = access_app()
    if acc is None:
        out("StudioVision (Access) introuvable. Ouvrez StudioVision puis relancez.")
        return 2
    try:
        f = acc.Forms(FORM_PATIENT)
    except Exception:
        out("La fiche patient (fPATIENTS) n'est pas ouverte. Ouvrez une fiche puis relancez.")
        return 2

    etat(f, "PATIENT ACTUEL")
    try:
        rs_initial = str(f.RecordSource)
    except Exception:
        rs_initial = ""

    code = input("Nouveau code patient a charger : ").strip()
    if not code.isdigit():
        out("Code invalide (chiffres attendus).")
        return 2

    # Nouveau SQL : on garde EXACTEMENT le gabarit actuel, seul le code change.
    m = re.search(r"^(.*\[Code patient\]\s*=\s*)(\d+)(\s*)$", rs_initial,
                  re.IGNORECASE | re.DOTALL)
    sql = (m.group(1) + code + m.group(3)) if m \
        else "select * from patients where [Code patient] = %s" % code
    out("")
    out("Nouveau RecordSource : %s" % sql)

    out("Application (f.RecordSource = ...) ...")
    try:
        f.RecordSource = sql
        out("  -> accepte.")
    except Exception as e:
        out("  -> REFUSE : %s" % e)
        out("Conclusion : la voie RecordSource n'est pas disponible ; on gardera la")
        out("recherche par la fenetre (frappe corrigee). Rien n'a ete modifie.")
        input("Entrée pour terminer...")
        return 1

    # Attente courte : en Access, changer le RecordSource requiert le formulaire.
    def charge():
        return norm(lire_ctrl(f, CTRL_CODE)) == norm(code)

    fin = time.time() + 4
    while time.time() < fin and not charge():
        time.sleep(0.3)
    if not charge():
        out("Pas encore charge -> Requery() ...")
        try:
            f.Requery()
        except Exception as e:
            out("Requery : %s" % e)
        fin = time.time() + 4
        while time.time() < fin and not charge():
            time.sleep(0.3)

    out("")
    etat(f, "APRES CHANGEMENT")
    v = norm(lire_ctrl(f, CTRL_CODE))
    if v == norm(code):
        out(">>> SUCCES : la fiche affiche le patient %s. La navigation directe" % code)
        out(">>> (sans clavier ni fenetre de recherche) FONCTIONNE.")
        out(">>> Verifiez aussi A L'ECRAN que la fiche affichee est la bonne")
        out(">>> (résumés de consultations, remarques, etc.).")
    elif v == "":
        out(">>> Fiche VIDE : le code %s n'existe probablement pas dans la table." % code)
        out(">>> (Le mecanisme RecordSource fonctionne mecaniquement ; refaites le test")
        out(">>>  avec un code patient existant.)")
    else:
        out(">>> ECHEC : la fiche affiche toujours %s (changement ignore)." % v)

    out("")
    rep = input("Entrée pour RESTAURER le patient initial (ou N pour laisser tel quel) : ").strip().lower()
    if rep != "n" and rs_initial:
        try:
            f.RecordSource = rs_initial
            try:
                f.Requery()
            except Exception:
                pass
            time.sleep(0.5)
            out("Patient initial restaure (%s)." % lire_ctrl(f, CTRL_CODE))
        except Exception as e:
            out("Restauration impossible : %s" % e)
    input("Entrée pour terminer...")
    return 0


if __name__ == "__main__":
    sys.exit(main())
