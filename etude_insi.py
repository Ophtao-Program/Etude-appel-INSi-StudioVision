#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
etude_insi.py — Etude de fiabilite de l'appel INSi depuis StudioVision.  (v7)

NAVIGATION (v7) : UNIQUEMENT par reecriture du RecordSource de fPATIENTS
    « select * from patients where [Code patient] = <code> »  (+ Requery).
COM pur, aucun clavier, aucun aller-retour par le MENU GENERAL. Le programme
verifie a chaque fois que la fiche affiche bien le bon « Code patient » avant
de poursuivre ; l'INSi n'est JAMAIS appele sur un mauvais patient.

MODES :
  --inspecter            liste tous les formulaires ouverts + controles
  --lister               liste les patients de l'annee (aucun appel)
  --test-ouverture [N]   teste l'OUVERTURE de N fiches (RecordSource), SANS INSi
                         (pour reperer les codes a valeurs speciales). N absent = toutes.
  --patient CODE         teste UN patient : ouverture + INSi, en detail
  --test                 etude complete limitee a 10 patients
  (defaut)               etude complete

Chaque enregistrement JSON contient la reponse INSi (00/01/02/erreurs) et les
traits lus dans la fiche (nom, prenom, sexe, date de naissance). Le numero INS
complet et le N° de securite sociale ne sont PAS enregistres.
"""

import sys
import os
import re
import json
import gc
import time
import argparse

try:
    sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
    sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
except Exception:
    pass


def out(msg):
    print(msg, flush=True)


def log(msg):
    sys.stderr.write("[etude] " + msg + "\n")
    sys.stderr.flush()


# ══════════════════════════════════════════════════════════════════════════════
#  CONFIGURATION
# ══════════════════════════════════════════════════════════════════════════════
PUBLIC_MDB    = r"M:\fichier\PUBLIC.mdb"
TABLE_CONSULT = "Consultation"
CHAMP_CODE    = "Code patient"
CHAMP_DATE    = "Date"
ANNEE_DEFAUT  = 2026

FORM_PATIENT = "fPATIENTS"
CTRL_CODE    = "Code patient"     # controle sur fPATIENTS
CTRL_NOM     = "NOM"
CTRL_PRENOM  = "Prénom"

# Gabarit du RecordSource (valide le 06/07/2026) :
RS_GABARIT = "select * from patients where [Code patient] = %s"

# Table + champs de la table patients (confirmes par exploration COM le 06/07/2026) :
TABLE_PATIENTS   = "patients"
CHAMP_SEXE_TABLE = "SEXE"              # entier : 1 = M, 2 = F  (vide = non renseigne)
CHAMP_DDN_TABLE  = "Date de Naissance"
CHAMP_NUMSS      = "SS"                # numero de securite sociale (NIR)
CHAMP_TITRE      = "Titre"             # 'Enfant' pour un enfant (SS = celui d'un parent)
VAL_TITRE_ENFANT = "Enfant"

# Correspondance sexe -> valeur ecrite dans le champ SEXE (a ajuster si besoin) :
SEXE_M = 1
SEXE_F = 2

NAV_ATTENTE = 4      # s max apres reecriture du RecordSource
INTER_APPEL = 0.15   # pause entre deux patients (le teleservice domine le temps total)

OUT_RESULTATS  = "etude_insi_resultats.json"
OUT_STATS      = "etude_insi_statistiques.json"
OUT_OUVERTURE  = "etude_ouverture_resultats.json"

# ── Cartographie INSi (identique a WebDMP Assistant, eprouvee) ────────────────
SUBFORM_FORM_NAME = "CARACTERISTIQUES PATIENT"
INSI_BTN_NAME     = "Commande136"
INSI_BTN_CAPTION  = "INSi"

OMAIN_CLASS   = "OMain"
SUBFORM_CLASS = "OFormPopup"
DIALOG_HINT   = "Identifiant National de Sant"
RESPONSE_HINT = "ponse INSi"
DLG_CLASS     = "#32770"

VAL_OK_ID = 1
VAL_NOM = 500
VAL_PRENOM = 501
VAL_LIEU = 502
VAL_SEXE = 503
VAL_DOB_D = 900
VAL_DOB_M = 901
VAL_DOB_Y = 902
RESP_OK_ID = 2
RESP_TEXT_ID = 65535

CLICK_CONFIRM = 10
WAIT_REPONSE  = 60

VK_SPACE = 0x20
VK_RETURN = 0x0D
KEYEVENTF_KEYUP = 0x0002
WM_GETTEXT = 0x000D
WM_GETTEXTLENGTH = 0x000E
WM_COMMAND = 0x0111
BM_CLICK = 0x00F5

ACC_TYPES = {100: "Label", 101: "Rectangle", 102: "Line", 103: "Image",
             104: "CommandButton", 105: "OptionGroup", 106: "OptionButton",
             107: "CheckBox", 108: "CheckBox", 109: "TextBox", 110: "ListBox",
             111: "ComboBox", 112: "SubForm", 118: "TabCtl", 119: "ActiveX/Custom",
             122: "ToggleButton", 123: "Hyperlink"}

HAS_WIN = sys.platform.startswith("win")


# ══════════════════════════════════════════════════════════════════════════════
#  Couche Win32 (ctypes) — fenetres INSi uniquement
# ══════════════════════════════════════════════════════════════════════════════
if HAS_WIN:
    import ctypes
    from ctypes import wintypes

    user32   = ctypes.WinDLL("user32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    user32.EnumWindows.argtypes = [EnumWindowsProc, wintypes.LPARAM]
    user32.EnumWindows.restype = wintypes.BOOL
    user32.EnumChildWindows.argtypes = [wintypes.HWND, EnumWindowsProc, wintypes.LPARAM]
    user32.EnumChildWindows.restype = wintypes.BOOL
    user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user32.GetClassNameW.restype = ctypes.c_int
    user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user32.GetWindowTextW.restype = ctypes.c_int
    user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
    user32.GetWindowTextLengthW.restype = ctypes.c_int
    user32.IsWindowVisible.argtypes = [wintypes.HWND]
    user32.IsWindowVisible.restype = wintypes.BOOL
    user32.GetDlgItem.argtypes = [wintypes.HWND, ctypes.c_int]
    user32.GetDlgItem.restype = wintypes.HWND
    user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, ctypes.c_size_t, ctypes.c_void_p]
    user32.SendMessageW.restype = ctypes.c_ssize_t
    user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, ctypes.c_size_t, ctypes.c_void_p]
    user32.PostMessageW.restype = wintypes.BOOL
    user32.SetForegroundWindow.argtypes = [wintypes.HWND]
    user32.SetForegroundWindow.restype = wintypes.BOOL
    user32.BringWindowToTop.argtypes = [wintypes.HWND]
    user32.BringWindowToTop.restype = wintypes.BOOL
    user32.GetForegroundWindow.restype = wintypes.HWND
    user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, wintypes.LPDWORD]
    user32.GetWindowThreadProcessId.restype = wintypes.DWORD
    user32.AttachThreadInput.argtypes = [wintypes.DWORD, wintypes.DWORD, wintypes.BOOL]
    user32.AttachThreadInput.restype = wintypes.BOOL
    user32.keybd_event.argtypes = [wintypes.BYTE, wintypes.BYTE, wintypes.DWORD, ctypes.c_void_p]
    kernel32.GetCurrentThreadId.restype = wintypes.DWORD

    def _enum(parent=0):
        res = []

        @EnumWindowsProc
        def _cb(h, _l):
            res.append(h)
            return True

        if parent:
            user32.EnumChildWindows(parent, _cb, 0)
        else:
            user32.EnumWindows(_cb, 0)
        return res

    def get_class(h):
        buf = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(h, buf, 256)
        return buf.value or ""

    def get_title(h):
        n = user32.GetWindowTextLengthW(h)
        if n <= 0:
            return ""
        buf = ctypes.create_unicode_buffer(n + 1)
        user32.GetWindowTextW(h, buf, n + 1)
        return buf.value or ""

    def wm_get_text(h):
        if not h:
            return ""
        n = user32.SendMessageW(h, WM_GETTEXTLENGTH, 0, 0)
        if n <= 0:
            return ""
        buf = ctypes.create_unicode_buffer(int(n) + 1)
        user32.SendMessageW(h, WM_GETTEXT, int(n) + 1, buf)
        return buf.value or ""

    def find_window(class_sub="", title_sub="", visible_only=True):
        for h in _enum():
            if visible_only and not user32.IsWindowVisible(h):
                continue
            if class_sub and class_sub.lower() not in get_class(h).lower():
                continue
            if title_sub and title_sub.lower() not in get_title(h).lower():
                continue
            return h
        return 0

    def foreground(hwnd):
        try:
            fg = user32.GetForegroundWindow()
            if fg == hwnd:
                return True
            cur = kernel32.GetCurrentThreadId()
            fg_thr = user32.GetWindowThreadProcessId(fg, None) if fg else 0
            tgt_thr = user32.GetWindowThreadProcessId(hwnd, None)
            if fg_thr:
                user32.AttachThreadInput(cur, fg_thr, True)
            user32.AttachThreadInput(cur, tgt_thr, True)
            user32.BringWindowToTop(hwnd)
            user32.SetForegroundWindow(hwnd)
            if fg_thr:
                user32.AttachThreadInput(cur, fg_thr, False)
            user32.AttachThreadInput(cur, tgt_thr, False)
            return user32.GetForegroundWindow() == hwnd
        except Exception as e:
            log("foreground: %s" % e)
            return False

    def foreground_sv(title_sub=""):
        h = 0
        if title_sub:
            h = find_window(class_sub=SUBFORM_CLASS, title_sub=title_sub)
        if not h:
            h = find_window(class_sub=SUBFORM_CLASS) or find_window(class_sub=OMAIN_CLASS)
        if h:
            foreground(h)
            time.sleep(0.15)
        return h

    def press_key(vk):
        user32.keybd_event(vk, 0, 0, 0)
        time.sleep(0.03)
        user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)

    def dlg_item_text(hdlg, ctrl_id):
        return wm_get_text(user32.GetDlgItem(hdlg, ctrl_id)).strip()

    def click_dlg_button(hdlg, ctrl_id):
        h = user32.GetDlgItem(hdlg, ctrl_id)
        if h:
            user32.SendMessageW(h, BM_CLICK, 0, 0)
            return True
        return bool(user32.PostMessageW(hdlg, WM_COMMAND, ctrl_id, 0))

    def read_response_static(hresp):
        txt = wm_get_text(user32.GetDlgItem(hresp, RESP_TEXT_ID))
        if txt and "INS=" in txt:
            return txt
        for c in _enum(hresp):
            t = wm_get_text(c)
            if t and "INS=" in t:
                return t
        return txt or ""

    def texte_dialogue_erreur():
        for h in _enum():
            if not user32.IsWindowVisible(h):
                continue
            if DLG_CLASS.lower() not in get_class(h).lower():
                continue
            t = get_title(h)
            low = t.lower()
            if DIALOG_HINT.lower() in low or RESPONSE_HINT.lower() in low:
                continue
            parts = [t] if t else []
            for c in _enum(h):
                s = wm_get_text(c)
                if s and len(s.strip()) > 2:
                    parts.append(s.strip())
            txt = " | ".join(p for p in parts if p).strip()
            if txt:
                return txt[:300]
        return ""

    def fermer_msgbox():
        for h in _enum():
            if not user32.IsWindowVisible(h):
                continue
            if DLG_CLASS.lower() not in get_class(h).lower():
                continue
            low = get_title(h).lower()
            if DIALOG_HINT.lower() in low or RESPONSE_HINT.lower() in low:
                continue
            for cid in (1, 2, 6, 7):
                if user32.GetDlgItem(h, cid):
                    click_dlg_button(h, cid)
                    time.sleep(0.2)
                    return True
        return False

    def fermer_dialogues_insi():
        for hint, cid in ((RESPONSE_HINT, RESP_OK_ID), (DIALOG_HINT, VAL_OK_ID)):
            h = find_window(class_sub=DLG_CLASS, title_sub=hint)
            if h:
                for c in (cid, 1, 2):
                    if click_dlg_button(h, c):
                        break
                time.sleep(0.2)


def _attendre(pred, timeout, interval=0.35):
    end = time.time() + timeout
    while time.time() < end:
        v = pred()
        if v:
            return v
        time.sleep(interval)
    return pred()


# ══════════════════════════════════════════════════════════════════════════════
#  Couche COM (Access)
# ══════════════════════════════════════════════════════════════════════════════
def co_init():
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass


def _is_access_app(obj):
    try:
        _ = obj.CurrentProject
        _ = obj.Forms.Count
        return True
    except Exception:
        return False


def access_app():
    co_init()
    found = []
    try:
        import pythoncom
        import win32com.client as win32
        rot = pythoncom.GetRunningObjectTable()
        for moniker in rot.EnumRunning():
            try:
                raw = rot.GetObject(moniker)
                app = win32.Dispatch(raw.QueryInterface(pythoncom.IID_IDispatch))
            except Exception:
                continue
            if not _is_access_app(app):
                continue
            try:
                fc = int(app.Forms.Count)
            except Exception:
                fc = -1
            try:
                proj = str(app.CurrentProject.Name)
            except Exception:
                proj = ""
            found.append({"app": app, "forms": fc, "project": proj})
    except Exception as e:
        log("ROT: %s" % e)
    if found:
        found.sort(key=lambda d: (d["forms"] > 0, bool(d["project"]), d["forms"]), reverse=True)
        best = found[0]
        if best["forms"] > 0 or best["project"]:
            return best["app"]
    try:
        import win32com.client as win32
        return win32.GetActiveObject("Access.Application")
    except Exception:
        return found[0]["app"] if found else None


# ── Anti-saturation d'Access : cache DB, detection de fermeture, maintenance ──
_DB = None


def _db(acc):
    """Renvoie une reference CurrentDb() mise en cache (evite de creer des milliers
    d'objets Database, cause classique de saturation d'Access)."""
    global _DB
    try:
        if _DB is not None:
            _ = _DB.Name          # teste la validite de la reference
            return _DB
    except Exception:
        _DB = None
    _DB = acc.CurrentDb()
    return _DB


def _reset_db():
    global _DB
    _DB = None


def studiovision_vivant(acc):
    """True si StudioVision (Access) repond encore."""
    try:
        _ = int(acc.Forms.Count)
        return True
    except Exception:
        return False


def maintenance(acc, reacquerir=False):
    """Libere les references, force le GC et fait une courte pause pour laisser
    Access respirer. Peut re-acquerir l'objet COM (proxy neuf) periodiquement."""
    _reset_db()
    try:
        gc.collect()
    except Exception:
        pass
    time.sleep(0.6)
    if reacquerir:
        try:
            acc2 = access_app()
            if acc2 is not None and studiovision_vivant(acc2):
                return acc2
        except Exception:
            pass
    return acc


def forme_active(acc):
    try:
        return str(acc.Screen.ActiveForm.Name)
    except Exception:
        return "(aucun)"


def controle_actif(acc):
    try:
        return str(acc.Screen.ActiveControl.Name)
    except Exception:
        return ""


def _form(acc, nom):
    try:
        return acc.Forms(nom)
    except Exception:
        return None


def _ctrl(form, nom):
    try:
        return form.Controls(nom)
    except Exception:
        return None


def _ctrl_value(acc, form_name, ctrl_name):
    try:
        v = acc.Forms(form_name).Controls(ctrl_name).Value
        return "" if v is None else str(v).strip()
    except Exception:
        return ""


def _rs_value(acc, form_name, field_name):
    """Lit un champ du Recordset du formulaire (table patients : SEXE, DDN...)."""
    try:
        v = acc.Forms(form_name).Recordset.Fields(field_name).Value
        return "" if v is None else str(v).strip()
    except Exception:
        return ""


def _norm_code(v):
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def focus_verifie(acc, form, ctrl_name, essais=3):
    """SetFocus sur le controle + VERIFICATION Screen.ActiveControl. True/False."""
    for _ in range(essais):
        try:
            form.SetFocus()
        except Exception:
            pass
        try:
            form.Controls(ctrl_name).SetFocus()
        except Exception as e:
            log("SetFocus %s: %s" % (ctrl_name, e))
        time.sleep(0.15)
        if controle_actif(acc).strip().lower() == ctrl_name.strip().lower():
            return True
        time.sleep(0.2)
    return False


# ══════════════════════════════════════════════════════════════════════════════
#  NAVIGATION v7 — ouverture d'une fiche par reecriture du RecordSource
# ══════════════════════════════════════════════════════════════════════════════
def fermer_sous_formulaire(acc):
    try:
        acc.DoCmd.Close(2, SUBFORM_FORM_NAME)   # 2 = acForm
        time.sleep(0.3)
    except Exception:
        pass


def _valeur_sql(cible):
    """Valeur SQL du code : numerique si -?\\d+, sinon texte quote (defensif)."""
    if re.match(r"^-?\d+$", cible):
        return cible
    return "'%s'" % cible.replace("'", "''")


def ouvrir_fiche_directe(acc, code):
    """Charge le patient dans fPATIENTS en reecrivant le RecordSource.
    -> dict {charge, code_charge, nom, prenom, sexe, ddn, detail}."""
    res = {"charge": False, "code_charge": "", "nom": "", "prenom": "",
           "sexe": "", "ddn": "", "detail": ""}
    f = _form(acc, FORM_PATIENT)
    if f is None:
        res["detail"] = "fiche fPATIENTS non ouverte (ouvrez une fiche patient dans StudioVision)"
        return res
    cible = _norm_code(code)
    val = _valeur_sql(cible)
    try:
        rs_actuel = str(f.RecordSource)
    except Exception:
        rs_actuel = ""
    m = re.search(r"^(.*\[Code patient\]\s*=\s*)(.+?)(\s*)$", rs_actuel, re.IGNORECASE | re.DOTALL)
    sql = (m.group(1) + val + m.group(3)) if m else (RS_GABARIT % val)
    try:
        f.RecordSource = sql
    except Exception as e:
        res["detail"] = "RecordSource refuse : %s" % e
        return res

    def _charge():
        return 1 if _norm_code(_ctrl_value(acc, FORM_PATIENT, CTRL_CODE)) == cible else 0

    if not _attendre(_charge, NAV_ATTENTE):
        try:
            f.Requery()
        except Exception:
            pass
        _attendre(_charge, NAV_ATTENTE)

    code_charge = _norm_code(_ctrl_value(acc, FORM_PATIENT, CTRL_CODE))
    res["code_charge"] = code_charge
    if code_charge == cible:
        res["charge"] = True
        res["nom"] = _ctrl_value(acc, FORM_PATIENT, CTRL_NOM)
        res["prenom"] = _ctrl_value(acc, FORM_PATIENT, CTRL_PRENOM)
        res["sexe"] = _rs_value(acc, FORM_PATIENT, CHAMP_SEXE_TABLE)
        res["ddn"] = _rs_value(acc, FORM_PATIENT, CHAMP_DDN_TABLE)[:10]
    elif code_charge == "":
        res["detail"] = "fiche vide (code %s absent de la table patients)" % cible
    else:
        res["detail"] = "code charge = %s au lieu de %s" % (code_charge, cible)
    return res


def lire_recordsource(acc):
    f = _form(acc, FORM_PATIENT)
    if f is None:
        return ""
    try:
        return str(f.RecordSource)
    except Exception:
        return ""


def restaurer_recordsource(acc, rs):
    if not rs:
        return
    f = _form(acc, FORM_PATIENT)
    if f is None:
        return
    try:
        f.RecordSource = rs
        try:
            f.Requery()
        except Exception:
            pass
        time.sleep(0.3)
    except Exception as e:
        log("restauration RecordSource: %s" % e)


def ouvrir_sous_formulaire(acc, essais=3):
    def _present():
        try:
            forms = acc.Forms
            for i in range(int(forms.Count)):
                if str(forms(i).Name).strip().upper() == SUBFORM_FORM_NAME.upper():
                    return forms(i)
        except Exception:
            pass
        return None

    if _present() is not None:
        return True, None
    derniere = ""
    for k in range(essais):
        try:
            acc.DoCmd.OpenForm(SUBFORM_FORM_NAME)
            time.sleep(0.6)
            if _present() is not None:
                return True, None
            derniere = "formulaire absent apres OpenForm"
        except Exception as e:
            derniere = str(e)
            log("OpenForm(%s) essai %d KO: %s" % (SUBFORM_FORM_NAME, k + 1, e))
        # recuperation entre essais : fermer un residu, purger les dialogues, pause croissante
        fermer_sous_formulaire(acc)
        fermer_dialogues_insi()
        if k == 0:
            _reset_db()
            try:
                gc.collect()
            except Exception:
                pass
        time.sleep(0.7 + 0.7 * k)
    return False, "ouverture %s : %s" % (SUBFORM_FORM_NAME, derniere)


# ══════════════════════════════════════════════════════════════════════════════
#  Appel INSi (sequence eprouvee ; deux flux ; focus verifie)
# ══════════════════════════════════════════════════════════════════════════════
def _form_caracteristiques(acc):
    try:
        forms = acc.Forms
        for i in range(int(forms.Count)):
            f = forms(i)
            if str(f.Name).strip().upper() == SUBFORM_FORM_NAME.upper():
                return f
    except Exception:
        pass
    return None


def _bouton_insi(form):
    c = _ctrl(form, INSI_BTN_NAME)
    if c is not None:
        return c, INSI_BTN_NAME
    try:
        ctrls = form.Controls
        for i in range(int(ctrls.Count)):
            c = ctrls(i)
            try:
                if str(getattr(c, "Caption", "")).strip().upper() == INSI_BTN_CAPTION.upper():
                    return c, str(c.Name)
            except Exception:
                continue
    except Exception:
        pass
    return None, ""


def _dialogue_insi():
    h = find_window(class_sub=DLG_CLASS, title_sub=RESPONSE_HINT)
    if h:
        return ("reponse", h)
    h = find_window(class_sub=DLG_CLASS, title_sub=DIALOG_HINT)
    if h:
        return ("validation", h)
    return None


def parse_response(text):
    fields = {}
    norm = text.replace("\\r\\n", "\n").replace("\r\n", "\n").replace("\r", "\n")
    for line in norm.split("\n"):
        if "=" in line:
            k, v = line.split("=", 1)
            fields[k.strip()] = v.strip()
    return fields


def read_validation_traits(hval):
    f = {}
    f["Nom"] = dlg_item_text(hval, VAL_NOM)
    f["Prenoms"] = dlg_item_text(hval, VAL_PRENOM)
    f["Lieu_naissance"] = dlg_item_text(hval, VAL_LIEU)
    f["Sexe"] = dlg_item_text(hval, VAL_SEXE)
    j = dlg_item_text(hval, VAL_DOB_D)
    m = dlg_item_text(hval, VAL_DOB_M)
    a = dlg_item_text(hval, VAL_DOB_Y)
    f["Date_naissance"] = ("%s/%s/%s" % (j, m, a)) if (j or m or a) else ""
    return f


def appel_insi(acc):
    r = {
        "classification": "", "reponse": "", "code": "",
        "nom": "", "prenoms": "", "sexe": "", "date_naissance": "", "lieu_naissance": "",
        "sexe_manquant": False, "ins_present": False, "ins_masque": "", "detail": "",
    }
    form = _form_caracteristiques(acc)
    if form is None:
        r["classification"] = "erreur_appel"
        r["detail"] = "sous-formulaire %s non ouvert" % SUBFORM_FORM_NAME
        return r
    btn, btn_name = _bouton_insi(form)
    if btn is None:
        r["classification"] = "erreur_appel"
        r["detail"] = "bouton INSi introuvable"
        return r

    got = None
    if focus_verifie(acc, form, btn_name):
        foreground_sv("PATIENTS")
        for vk in (VK_SPACE, VK_RETURN):
            if controle_actif(acc).strip().lower() != btn_name.strip().lower():
                if not focus_verifie(acc, form, btn_name):
                    break
            press_key(vk)
            got = _attendre(_dialogue_insi, CLICK_CONFIRM)
            if got:
                break
    else:
        r["classification"] = "erreur_appel"
        r["detail"] = "focus non obtenu sur le bouton INSi — aucune touche envoyee"
        return r
    if not got:
        r["classification"] = "erreur_appel"
        r["detail"] = "aucun dialogue INSi apres activation"
        t = texte_dialogue_erreur()
        if t:
            r["detail"] += " | " + t
            fermer_msgbox()
        return r

    kind, hwin = got
    hresp = 0
    if kind == "validation":
        traits = read_validation_traits(hwin)
        r["nom"] = traits.get("Nom", "")
        r["prenoms"] = traits.get("Prenoms", "")
        r["sexe"] = traits.get("Sexe", "")
        r["date_naissance"] = traits.get("Date_naissance", "")
        r["lieu_naissance"] = traits.get("Lieu_naissance", "")
        r["sexe_manquant"] = (r["sexe"].strip() == "")
        click_dlg_button(hwin, VAL_OK_ID)
        hresp = _attendre(lambda: find_window(class_sub=DLG_CLASS, title_sub=RESPONSE_HINT),
                          WAIT_REPONSE)
    else:
        hresp = hwin

    if not hresp:
        t = texte_dialogue_erreur()
        r["detail"] = t or "aucune fenetre reponse (teleservice sans retour)"
        if t:
            fermer_msgbox()
        r["classification"] = "sexe_manquant" if r["sexe_manquant"] else "sans_reponse"
        fermer_dialogues_insi()
        return r

    time.sleep(0.5)
    text = read_response_static(hresp)
    click_dlg_button(hresp, RESP_OK_ID)
    time.sleep(0.2)

    if not text:
        r["classification"] = "reponse_illisible"
        fermer_dialogues_insi()
        return r

    f = parse_response(text)
    r["reponse"] = f.get("Reponse", "")
    r["code"] = f.get("code", "").strip()
    for src, dst in (("Nom", "nom"), ("Prenoms", "prenoms"), ("Prenom", "prenoms"),
                     ("Sexe", "sexe"), ("Date_naissance", "date_naissance"),
                     ("Lieu_naissance", "lieu_naissance")):
        if f.get(src):
            r[dst] = f[src]
    ins = re.sub(r"\D", "", f.get("INS", ""))
    if len(ins) == 15:
        r["ins_present"] = True
        r["ins_masque"] = ins[0] + "X" * 11 + ins[-3:]
    if not r["sexe"].strip():
        r["sexe_manquant"] = True

    code = r["code"]
    if code == "00":
        r["classification"] = "00_trouve"
    elif code == "01":
        r["classification"] = "01_non_trouve"
    elif code == "02":
        r["classification"] = "02_plusieurs"
    elif code == "":
        r["classification"] = "sexe_manquant" if r["sexe_manquant"] else "reponse_illisible"
    else:
        r["classification"] = "code_%s" % code

    fermer_dialogues_insi()
    return r


def _autre_sexe(s):
    return "F" if s == "M" else "M"


def _sexe_valeur(s):
    return SEXE_M if s == "M" else SEXE_F


def lire_sexe_table(acc, code):
    """Lit patients.SEXE directement dans la table (etat persiste)."""
    try:
        db = _db(acc)
        rs = db.OpenRecordset("SELECT [%s] FROM %s WHERE [%s] = %s"
                              % (CHAMP_SEXE_TABLE, TABLE_PATIENTS, CTRL_CODE,
                                 _valeur_sql(_norm_code(code))))
        v = "" if rs.EOF else rs.Fields(0).Value
        rs.Close()
        return "" if v is None else str(v).strip()
    except Exception:
        return ""


def ecrire_sexe_table(acc, code, valeur):
    """Ecrit patients.SEXE = valeur (SEXE_M, SEXE_F, ou None pour vider).
    -> (True, None) | (False, err). Corrige DURABLEMENT la fiche patient."""
    try:
        db = _db(acc)
    except Exception as e:
        return False, "CurrentDb: %s" % e
    cible = _norm_code(code)
    try:
        rs = db.OpenRecordset("SELECT [%s] FROM %s WHERE [%s] = %s"
                              % (CHAMP_SEXE_TABLE, TABLE_PATIENTS, CTRL_CODE, _valeur_sql(cible)))
    except Exception as e:
        return False, "OpenRecordset: %s" % e
    try:
        if rs.EOF:
            rs.Close()
            return False, "patient %s absent de la table" % cible
        rs.Edit()
        rs.Fields(CHAMP_SEXE_TABLE).Value = valeur
        rs.Update()
        rs.Close()
        return True, None
    except Exception as e:
        try:
            rs.Close()
        except Exception:
            pass
        return False, "ecriture SEXE: %s" % e


def resoudre_sexe_et_appeler(acc, code, num_ss, titre, verbeux=False):
    """Sexe absent : on le remplit (deduit du n°SS si possible, sinon essais F puis M),
    on appelle INSi (jusqu'a 2 fois), on GARDE le sexe si un appel trouve le patient
    (code 00/02), sinon on REMET le champ sexe a vide. On n'appelle JAMAIS INSi avec un
    champ sexe vide."""
    r = {
        "classification": "", "reponse": "", "code": "",
        "nom": "", "prenoms": "", "sexe": "", "date_naissance": "", "lieu_naissance": "",
        "sexe_manquant": True, "ins_present": False, "ins_masque": "", "detail": "",
        "num_ss_present": bool(re.sub(r"\D", "", num_ss or "")),
        "titre": (titre or "").strip(), "sexe_source": "", "sexe_corrige": "",
        "nb_appels_insi": 0,
    }
    ss = re.sub(r"\D", "", num_ss or "")
    est_enfant = ((titre or "").strip().lower() == VAL_TITRE_ENFANT.lower())

    sexe_ss = None
    if ss and not est_enfant:
        if ss[0] == "1":
            sexe_ss = "M"
        elif ss[0] == "2":
            sexe_ss = "F"

    if sexe_ss:
        ordre, cas = [sexe_ss, _autre_sexe(sexe_ss)], "numSS"
    elif est_enfant:
        ordre, cas = ["F", "M"], "enfant"
    elif not ss:
        ordre, cas = ["F", "M"], "sans_numSS"
    else:
        ordre, cas = ["F", "M"], "numSS_invalide"
    r["sexe_source"] = cas
    if verbeux:
        out("   sexe ABSENT -> cas=%s, essais=%s (n°SS %s, titre=%s)"
            % (cas, "/".join(ordre), ("present" if ss else "absent"), r["titre"] or "-"))

    meilleur = None
    for sexe in ordre:
        okw, errw = ecrire_sexe_table(acc, code, _sexe_valeur(sexe))
        if not okw:
            r["classification"] = "sexe_non_ecrit"
            r["detail"] = "ecriture du sexe impossible : %s" % errw
            return r
        ouvrir_fiche_directe(acc, code)          # recharge la fiche (reflete le sexe)
        oks, errs = ouvrir_sous_formulaire(acc)
        if not oks:
            ecrire_sexe_table(acc, code, None)   # on ne laisse pas un sexe non valide
            r["classification"] = "sous_formulaire_absent"
            r["detail"] = errs
            return r
        if verbeux:
            out("   essai sexe=%s (SEXE=%s) -> appel INSi..." % (sexe, _sexe_valeur(sexe)))
        res = appel_insi(acc)
        fermer_sous_formulaire(acc)
        r["nb_appels_insi"] += 1
        code_insi = (res.get("code") or "").strip()
        if code_insi in ("00", "02"):
            for k in ("reponse", "code", "nom", "prenoms", "sexe",
                      "date_naissance", "lieu_naissance", "ins_present", "ins_masque"):
                r[k] = res.get(k, r.get(k))
            r["classification"] = "00_trouve" if code_insi == "00" else "02_plusieurs"
            r["sexe_corrige"] = sexe
            r["sexe_manquant"] = False
            r["detail"] = "sexe rempli = %s (SEXE=%s) via %s" % (sexe, _sexe_valeur(sexe), cas)
            return r                              # correction CONSERVEE dans la table
        if meilleur is None:
            meilleur = res

    # Aucun appel concluant -> on REMET le champ sexe a vide
    ecrire_sexe_table(acc, code, None)
    r["sexe_corrige"] = ""
    if meilleur:
        r["reponse"] = meilleur.get("reponse", "")
        r["code"] = (meilleur.get("code") or "").strip()
    detail_base = meilleur.get("detail", "") if meilleur else ""
    if cas == "enfant":
        r["classification"] = "pas_sexe_enfant"
        r["detail"] = "enfant : essais F/M sans succes — champ sexe remis vide"
    elif cas == "sans_numSS":
        r["classification"] = "pas_sexe_pas_numSS"
        r["detail"] = "pas de n°SS : essais F/M sans succes — champ sexe remis vide"
    else:
        r["classification"] = "sexe_non_resolu"
        r["detail"] = "aucun sexe (essais %s) ne donne 00/02 — champ sexe remis vide" % "/".join(ordre)
    if detail_base:
        r["detail"] += " | " + detail_base
    return r


def traiter_patient(acc, code, verbeux=False):
    if verbeux:
        out("   [1] ouverture de la fiche %s (RecordSource)..." % code)
    fermer_dialogues_insi()
    fermer_sous_formulaire(acc)
    nav = ouvrir_fiche_directe(acc, code)
    if not nav["charge"]:
        return {"classification": "navigation_echec", "detail": nav["detail"]}
    num_ss = _rs_value(acc, FORM_PATIENT, CHAMP_NUMSS)
    titre = _rs_value(acc, FORM_PATIENT, CHAMP_TITRE)
    sexe_present = (nav["sexe"].strip() != "")
    if verbeux:
        out("   fiche : %s %s  (sexe=%s, ne(e) %s, n°SS=%s, titre=%s)"
            % (nav["nom"], nav["prenom"], nav["sexe"] or "VIDE", nav["ddn"] or "?",
               "oui" if num_ss.strip() else "non", titre or "-"))

    if sexe_present:
        if verbeux:
            out("   [2] ouverture de %s..." % SUBFORM_FORM_NAME)
        ok2, err2 = ouvrir_sous_formulaire(acc)
        if not ok2:
            r = {"classification": "sous_formulaire_absent", "detail": err2}
        else:
            if verbeux:
                out("   [3] appel INSi (sexe present)...")
            r = appel_insi(acc)
            fermer_sous_formulaire(acc)
        r.setdefault("sexe_source", "present")
        r["sexe_fiche"] = nav["sexe"]
    else:
        # Sexe ABSENT : on ne fait JAMAIS d'appel avec un sexe vide -> resolution
        if verbeux:
            out("   [2-3] sexe ABSENT -> remplissage + appel(s) INSi...")
        r = resoudre_sexe_et_appeler(acc, code, num_ss, titre, verbeux)
        r["sexe_fiche"] = lire_sexe_table(acc, code)   # etat final (corrige ou remis vide)

    r["nom_fiche"] = nav["nom"]
    r["prenom_fiche"] = nav["prenom"]
    r["date_naissance_fiche"] = nav["ddn"]
    r.setdefault("num_ss_present", bool(num_ss.strip()))
    r.setdefault("titre", titre.strip())
    if r.get("nom") and nav["nom"] and r["nom"].strip().upper() != nav["nom"].strip().upper():
        r["detail"] = (r.get("detail", "") + " | ATTENTION nom reponse (%s) != fiche (%s)"
                       % (r["nom"], nav["nom"])).strip(" |")
    return r


# ══════════════════════════════════════════════════════════════════════════════
#  Lecture PUBLIC.mdb (patients de l'annee)
# ══════════════════════════════════════════════════════════════════════════════
def _ouvrir_base(acc, chemin):
    if acc is not None:
        try:
            db = acc.CurrentDb()
            log("Base : CurrentDb() de StudioVision.")
            return db
        except Exception as e:
            log("CurrentDb KO (%s)" % e)
    for prog in ("DAO.DBEngine.120", "DAO.DBEngine.36"):
        try:
            import win32com.client as w
            eng = w.Dispatch(prog)
            db = eng.OpenDatabase(chemin, False, True)
            log("Base : %s -> %s" % (prog, chemin))
            return db
        except Exception as e:
            log("%s KO: %s" % (prog, e))
    if acc is not None:
        try:
            return acc.DBEngine.OpenDatabase(chemin, False, True)
        except Exception as e:
            log("DBEngine(Access) KO: %s" % e)
    raise RuntimeError("Impossible d'ouvrir la base (CurrentDb / PUBLIC.mdb).")


def _annee(d):
    try:
        return int(d.year)
    except Exception:
        try:
            return int(str(d)[:4])
        except Exception:
            return None


def _date_str(d):
    try:
        return d.strftime("%Y-%m-%d")
    except Exception:
        return str(d)[:19]


def lire_patients_annee(acc, chemin, annee):
    out("Lecture des consultations (table %s)..." % TABLE_CONSULT)
    db = _ouvrir_base(acc, chemin)
    vus = {}
    d1 = "#01/01/%d#" % annee
    d2 = "#01/01/%d#" % (annee + 1)
    sql = ("SELECT [%s], MAX([%s]) FROM [%s] WHERE [%s] >= %s AND [%s] < %s GROUP BY [%s]"
           % (CHAMP_CODE, CHAMP_DATE, TABLE_CONSULT, CHAMP_DATE, d1, CHAMP_DATE, d2, CHAMP_CODE))
    try:
        out("  requete des patients %d..." % annee)
        rs = db.OpenRecordset(sql)
        while not rs.EOF:
            code = rs.Fields(0).Value
            d = rs.Fields(1).Value
            if code is not None:
                vus[_norm_code(code)] = (_norm_code(code), _date_str(d))
            rs.MoveNext()
        rs.Close()
    except Exception as e:
        log("Requete groupee KO (%s) -> lecture par blocs." % e)
        out("  (filtrage direct impossible, lecture complete de la table...)")
        rs = db.OpenRecordset("SELECT [%s], [%s] FROM [%s]" % (CHAMP_CODE, CHAMP_DATE, TABLE_CONSULT))
        total = 0
        while not rs.EOF:
            data = rs.GetRows(2000)
            if not data or not len(data[0]):
                break
            for i in range(len(data[0])):
                total += 1
                codev = data[0][i]
                d = data[1][i]
                if codev is not None and _annee(d) == annee:
                    ds = _date_str(d)
                    key = _norm_code(codev)
                    if key not in vus or ds > vus[key][1]:
                        vus[key] = (key, ds)
            out("    ... %d lignes lues" % total)
        rs.Close()
    try:
        db.Close()
    except Exception:
        pass
    out("  -> %d patient(s) distinct(s) ayant consulte en %d." % (len(vus), annee))
    # tri numerique quand possible, sinon alphabetique
    def _cle(t):
        c = t[0]
        try:
            return (0, int(c))
        except Exception:
            return (1, c)
    return sorted(vus.values(), key=_cle)


# ══════════════════════════════════════════════════════════════════════════════
#  Sauvegarde + statistiques
# ══════════════════════════════════════════════════════════════════════════════
def sauver_json(obj, chemin):
    tmp = chemin + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    os.replace(tmp, chemin)


def calculer_stats(results, annee):
    from collections import Counter
    n = len(results)
    c = Counter(r.get("classification", "?") for r in results)
    n_sexe = sum(1 for r in results if r.get("sexe_manquant"))
    succes = c.get("00_trouve", 0)

    def pct(x):
        return round(100.0 * x / n, 1) if n else 0.0

    return {
        "annee": annee, "total_patients": n,
        "trouve_00": succes, "trouve_00_pct": pct(succes),
        "non_trouve_01": c.get("01_non_trouve", 0), "non_trouve_01_pct": pct(c.get("01_non_trouve", 0)),
        "plusieurs_02": c.get("02_plusieurs", 0), "plusieurs_02_pct": pct(c.get("02_plusieurs", 0)),
        "sexe_manquant": n_sexe, "sexe_manquant_pct": pct(n_sexe),
        "total_echecs": n - succes, "total_echecs_pct": pct(n - succes),
        "detail_par_classification": dict(c),
    }


def afficher_stats(s):
    out("")
    out("==================== STATISTIQUES INSi %s ====================" % s["annee"])
    out("  Patients traites .................. %d" % s["total_patients"])
    out("  00  patient trouve ................ %d  (%.1f %%)" % (s["trouve_00"], s["trouve_00_pct"]))
    out("  01  patient NON trouve ............ %d  (%.1f %%)" % (s["non_trouve_01"], s["non_trouve_01_pct"]))
    out("  02  plusieurs patients ............ %d  (%.1f %%)" % (s["plusieurs_02"], s["plusieurs_02_pct"]))
    out("  sexe non renseigne ................ %d  (%.1f %%)" % (s["sexe_manquant"], s["sexe_manquant_pct"]))
    out("  ---------------------------------------------------------")
    out("  TOTAL echecs (hors 00) ............ %d  (%.1f %%)" % (s["total_echecs"], s["total_echecs_pct"]))
    out("")
    out("  Detail par classification :")
    for k, v in sorted(s["detail_par_classification"].items(), key=lambda kv: -kv[1]):
        out("     %-22s %d" % (k, v))
    out("=============================================================")


# ══════════════════════════════════════════════════════════════════════════════
#  Modes
# ══════════════════════════════════════════════════════════════════════════════
def _dump_controls(form):
    try:
        ctrls = form.Controls
        n = int(ctrls.Count)
    except Exception as e:
        out("    (controles illisibles: %s)" % e)
        return
    out("    %d controle(s) :" % n)
    for i in range(n):
        try:
            c = ctrls(i)
            name = str(getattr(c, "Name", ""))
            try:
                ct = int(c.ControlType)
            except Exception:
                ct = -1
            tname = ACC_TYPES.get(ct, "type%s" % ct)
            extra = []
            try:
                cap = str(getattr(c, "Caption", "") or "")
                if cap:
                    extra.append("Caption=%r" % cap[:50])
            except Exception:
                pass
            try:
                v = getattr(c, "Value", None)
                if v not in (None, ""):
                    extra.append("Value=%r" % str(v)[:30])
            except Exception:
                pass
            try:
                extra.append("T=%s L=%s" % (int(c.Top), int(c.Left)))
            except Exception:
                pass
            out("      [%-14s] %-26s %s" % (tname, name, "  ".join(extra)))
        except Exception as e:
            out("      (controle %d illisible: %s)" % (i, e))


def mode_inspecter(acc):
    out("Formulaire actif : %s" % forme_active(acc))
    try:
        forms = acc.Forms
        nf = int(forms.Count)
    except Exception as e:
        out("acc.Forms illisible: %s" % e)
        return 2
    noms = []
    for fi in range(nf):
        try:
            noms.append(str(forms(fi).Name))
        except Exception:
            noms.append("(?)")
    out("%d formulaire(s) ouvert(s) : %s" % (nf, " | ".join(noms)))
    for fi in range(nf):
        try:
            f = forms(fi)
            fname = str(f.Name)
        except Exception as e:
            out("  (formulaire %d illisible: %s)" % (fi, e))
            continue
        out("")
        out("========== Formulaire : %s ==========" % fname)
        try:
            out("    RecordSource : %s" % f.RecordSource)
        except Exception:
            pass
        _dump_controls(f)
    return 0


def mode_patient(acc, code):
    out("=== Test d'un seul patient : code=%s ===" % code)
    r = traiter_patient(acc, _norm_code(code), verbeux=True)
    out("")
    out("  classification : %s" % r.get("classification"))
    for k in ("reponse", "code", "nom", "prenoms", "sexe", "date_naissance", "lieu_naissance",
              "nom_fiche", "prenom_fiche", "sexe_fiche", "date_naissance_fiche",
              "num_ss_present", "titre", "sexe_source", "sexe_corrige", "nb_appels_insi",
              "sexe_manquant", "ins_present", "ins_masque", "detail"):
        if k in r and r[k] not in ("", False, None):
            out("  %-20s : %s" % (k, r[k]))
    fermer_dialogues_insi()
    return 0


def mode_test_ouverture(acc, n, annee, chemin_mdb, out_path):
    """Teste UNIQUEMENT l'ouverture des fiches (RecordSource), sans appel INSi."""
    out("=== Test d'ouverture des fiches (RecordSource, SANS INSi) ===")
    if _form(acc, FORM_PATIENT) is None:
        out("Ouvrez une fiche patient (n'importe laquelle) dans StudioVision puis relancez.")
        return 2
    rs_initial = lire_recordsource(acc)
    code_initial = _ctrl_value(acc, FORM_PATIENT, CTRL_CODE)

    try:
        patients = lire_patients_annee(acc, chemin_mdb, annee)
    except Exception as e:
        out("ERREUR lecture base : %s" % e)
        return 2
    total = len(patients)
    if n and 0 < n < total:
        patients = patients[:n]
    out("Test de %d fiche(s) (sur %d patient(s) %d)." % (len(patients), total, annee))
    out("Patient initial : %s (restaure a la fin)." % (code_initial or "?"))
    out("")

    results = []
    nb_ok = nb_vide = nb_echec = 0
    t0 = time.time()
    try:
        for i, (code, dcons) in enumerate(patients, 1):
            r = ouvrir_fiche_directe(acc, code)
            entry = {"index": i, "code_patient": code, "date_consultation": dcons,
                     "charge": r["charge"], "code_charge": r["code_charge"],
                     "nom": r["nom"], "prenom": r["prenom"], "sexe": r["sexe"],
                     "date_naissance": r["ddn"], "detail": r["detail"]}
            results.append(entry)
            sauver_json(results, out_path)
            if r["charge"]:
                nb_ok += 1
                out("  [%d/%d] code=%s -> OK   %s %s" % (i, len(patients), code, r["nom"], r["prenom"]))
            elif "vide" in r["detail"]:
                nb_vide += 1
                out("  [%d/%d] code=%s -> FICHE VIDE   (%s)" % (i, len(patients), code, r["detail"]))
            else:
                nb_echec += 1
                out("  [%d/%d] code=%s -> ECHEC   (%s)" % (i, len(patients), code, r["detail"]))
            time.sleep(0.12)
    except KeyboardInterrupt:
        out("\nInterruption : %d fiche(s) testee(s)." % len(results))

    restaurer_recordsource(acc, rs_initial)
    out("")
    out("Patient initial restaure : %s" % (_ctrl_value(acc, FORM_PATIENT, CTRL_CODE) or "?"))
    out("")
    out("==================== RESULTAT OUVERTURE ====================")
    out("  Fiches testees ......... %d" % len(results))
    out("  Ouvertes (OK) .......... %d" % nb_ok)
    out("  Fiches vides ........... %d   (code absent de la table patients)" % nb_vide)
    out("  Echecs (probleme) ...... %d" % nb_echec)
    problematiques = [e for e in results if not e["charge"]]
    if problematiques:
        out("")
        out("  Codes NON ouverts :")
        for e in problematiques:
            out("     code=%-14s %s" % (e["code_patient"], e["detail"]))
    out("")
    out("  Fichier : %s" % os.path.abspath(out_path))
    out("  Duree : %.0f s" % (time.time() - t0))
    out("===========================================================")
    return 0


# Classifications DEFINITIVES : seules celles-ci sont conservees a la reprise.
# Toutes les autres (erreurs transitoires) sont automatiquement RETRAITEES.
CLASSIFS_DEFINITIVES = {"00_trouve", "01_non_trouve", "02_plusieurs",
                        "pas_sexe_pas_numSS", "pas_sexe_enfant", "sexe_non_resolu"}

MAINTENANCE_TOUS = 50   # maintenance memoire tous les N patients


def _est_erreur(cls):
    return cls in ("exception", "sous_formulaire_absent", "navigation_echec",
                   "erreur_appel", "sans_reponse", "reponse_illisible", "sexe_non_ecrit", "")


def _nouvel_entry(code, dcons, index):
    return {
        "index": index, "code_patient": code, "date_consultation": dcons,
        "classification": "", "reponse": "", "code": "",
        "nom": "", "prenoms": "", "sexe": "", "date_naissance": "", "lieu_naissance": "",
        "nom_fiche": "", "prenom_fiche": "", "sexe_fiche": "", "date_naissance_fiche": "",
        "num_ss_present": False, "titre": "", "sexe_source": "", "sexe_corrige": "",
        "nb_appels_insi": 0, "sexe_manquant": False, "ins_present": False,
        "ins_masque": "", "detail": "", "horodatage": time.strftime("%Y-%m-%d %H:%M:%S"),
    }


def _traiter_avec_reprise(acc, code):
    """Traite un patient ; un seul reessai en cas d'echec transitoire (apres maintenance)."""
    try:
        r = traiter_patient(acc, code)
    except Exception as e:
        r = {"classification": "exception", "detail": str(e)}
    if _est_erreur(r.get("classification", "")) and studiovision_vivant(acc):
        maintenance(acc)
        try:
            r2 = traiter_patient(acc, code)
        except Exception as e:
            r2 = {"classification": "exception", "detail": str(e)}
        if not _est_erreur(r2.get("classification", "")):
            r2["detail"] = (r2.get("detail", "") + " | reessai OK apres %s"
                            % r.get("classification", "?")).strip(" |")
            return r2
    return r


def run_etude(annee, limit, chemin_mdb, out_res, out_stats,
              lister, patient, inspecter, test_ouverture):
    if not HAS_WIN:
        out("Ce programme fonctionne uniquement sous Windows (StudioVision + pywin32).")
        return 2
    acc = access_app()
    if acc is None:
        out("StudioVision (Access) introuvable. Ouvrez StudioVision puis relancez.")
        return 2
    out("StudioVision detecte (formulaire actif : %s)." % forme_active(acc))

    if inspecter:
        return mode_inspecter(acc)
    if test_ouverture is not None:
        return mode_test_ouverture(acc, test_ouverture, annee, chemin_mdb, OUT_OUVERTURE)
    if patient is not None:
        return mode_patient(acc, patient)

    try:
        patients = lire_patients_annee(acc, chemin_mdb, annee)
    except Exception as e:
        out("ERREUR de lecture de la base : %s" % e)
        return 2
    out("%d patient(s) sur l'annee %d." % (len(patients), annee))

    # REPRISE robuste : on conserve UNIQUEMENT les resultats definitifs ; les erreurs
    # (transitoires) sont automatiquement retraitees. Les entrees sont indexees par code
    # pour pouvoir en remplacer une ancienne sans creer de doublon.
    par_code = {}
    if os.path.exists(out_res):
        try:
            with open(out_res, encoding="utf-8") as _f:
                anciens = json.load(_f)
            if isinstance(anciens, list):
                for e in anciens:
                    par_code[_norm_code(e.get("code_patient"))] = e
        except Exception as e:
            log("Resultats existants illisibles (%s) — on repart de zero." % e)
            par_code = {}
    definitifs = set(c for c, e in par_code.items()
                     if e.get("classification") in CLASSIFS_DEFINITIVES)
    a_retraiter = sum(1 for c, e in par_code.items()
                      if e.get("classification") and c not in definitifs)
    if par_code:
        out("Reprise : %d resultat(s) definitif(s) conserves ; %d erreur(s) a retraiter."
            % (len(definitifs), a_retraiter))

    a_faire = [(c, d) for (c, d) in patients if _norm_code(c) not in definitifs]
    if limit:
        a_faire = a_faire[:limit]
    out("%d patient(s) a traiter maintenant." % len(a_faire))

    def _liste():
        return list(par_code.values())

    rs_initial = lire_recordsource(acc)
    t0 = time.time()
    interrompu = False
    try:
        for i, (code, dcons) in enumerate(a_faire, 1):
            # StudioVision repond-il toujours ? sinon on tente une reconnexion, puis on s'arrete.
            if not studiovision_vivant(acc):
                acc = maintenance(acc, reacquerir=True)
                if not studiovision_vivant(acc):
                    interrompu = True
                    out("")
                    out("StudioVision ne repond plus — etude INTERROMPUE (avancement sauvegarde).")
                    out("Rouvrez StudioVision (avec une fiche patient affichee) et relancez :")
                    out("la reprise repartira automatiquement des patients non traites.")
                    break

            entry = _nouvel_entry(code, dcons, len(par_code) + 1)
            entry.update(_traiter_avec_reprise(acc, code))
            par_code[_norm_code(code)] = entry
            sauver_json(_liste(), out_res)
            out("  [%d/%d] code=%s -> %s%s"
                % (i, len(a_faire), code, entry.get("classification", ""),
                   ("  (%s)" % entry["detail"]) if entry.get("detail") else ""))
            # maintenance memoire periodique (anti-saturation d'Access)
            if i % MAINTENANCE_TOUS == 0:
                out("  … maintenance memoire apres %d patients" % i)
                acc = maintenance(acc, reacquerir=(i % (MAINTENANCE_TOUS * 4) == 0))
            time.sleep(INTER_APPEL)
    except KeyboardInterrupt:
        interrompu = True
        out("\nInterruption : avancement sauvegarde (%d resultat(s) au total)." % len(par_code))

    if studiovision_vivant(acc):
        restaurer_recordsource(acc, rs_initial)
    results = _liste()
    stats = calculer_stats(results, annee)
    sauver_json(stats, out_stats)
    afficher_stats(stats)
    out("Resultats  : %s" % os.path.abspath(out_res))
    out("Stats      : %s" % os.path.abspath(out_stats))
    out("Duree      : %.0f s" % (time.time() - t0))
    if interrompu:
        out("(Etude interrompue : relancez pour reprendre la ou vous en etiez.)")
    return 0


def main(argv):
    p = argparse.ArgumentParser(description="Etude de fiabilite de l'appel INSi depuis StudioVision (v9).")
    p.add_argument("--inspecter", action="store_true", help="lister tous les formulaires + controles")
    p.add_argument("--lister", action="store_true", help="lister les patients (aucun appel)")
    p.add_argument("--test-ouverture", dest="test_ouverture", type=int, nargs="?", const=0,
                   default=None, metavar="N",
                   help="tester l'ouverture de N fiches (RecordSource) SANS INSi (N absent = toutes)")
    p.add_argument("--patient", default=None, help="tester UN patient (ouverture + INSi)")
    p.add_argument("--test", action="store_true", help="etude complete limitee a 10 patients")
    p.add_argument("--limit", type=int, default=None, help="limiter a N patients")
    p.add_argument("--annee", type=int, default=ANNEE_DEFAUT, help="annee etudiee")
    p.add_argument("--public-mdb", default=PUBLIC_MDB, help="chemin de PUBLIC.mdb")
    p.add_argument("--output", default=OUT_RESULTATS, help="fichier resultats JSON")
    p.add_argument("--stats", default=OUT_STATS, help="fichier statistiques JSON")
    args = p.parse_args(argv)
    limit = 10 if args.test else args.limit
    return run_etude(args.annee, limit, args.public_mdb, args.output, args.stats,
                     args.lister, args.patient, args.inspecter, args.test_ouverture)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
