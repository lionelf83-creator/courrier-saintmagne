#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Courrier Saint-Magne V9
- PyQt6 WebEngine (beau design HTML dans un exe)
- Base SQLite sur disque réseau partagé
- Drag & drop PDF/scans stockés sur le réseau
- Import Excel existant + scans
- Export Excel
- Connexion par agent avec création de compte
"""

import sys, os, json, sqlite3, hashlib, shutil, base64, threading
from datetime import datetime
from pathlib import Path

from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWebEngineCore import QWebEngineProfile, QWebEnginePage
from PyQt6.QtCore import QUrl, pyqtSlot, QObject, Qt
from PyQt6.QtWebChannel import QWebChannel
from PyQt6.QtGui import QIcon

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font as XLFont, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────
#  CHEMINS
# ─────────────────────────────────────────
APP_DIR   = Path(os.path.abspath(__file__)).parent / "CourrierSaintMagne_data"
SCANS_DIR = APP_DIR / "scans"
DB_PATH   = APP_DIR / "courrier.db"
APP_DIR.mkdir(parents=True, exist_ok=True)
SCANS_DIR.mkdir(parents=True, exist_ok=True)

# ─────────────────────────────────────────
#  BASE DE DONNÉES
# ─────────────────────────────────────────
db_lock = threading.Lock()

def get_con():
    con = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    con.row_factory = sqlite3.Row
    return con

def init_db():
    with db_lock:
        con = get_con()
        con.executescript("""
        CREATE TABLE IF NOT EXISTS agents (
            id TEXT PRIMARY KEY, prenom TEXT, nom TEXT,
            role TEXT DEFAULT 'agent', code TEXT,
            created_at TEXT DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS arrivee (
            id TEXT PRIMARY KEY, expediteur TEXT, objet TEXT,
            date_courrier TEXT, recu_le TEXT, destinataire TEXT,
            commentaires TEXT, scan_name TEXT, agent TEXT,
            created_at TEXT DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS depart (
            id TEXT PRIMARY KEY, agent_nom TEXT, destinataire TEXT,
            objet TEXT, date_envoi TEXT, heure_envoi TEXT,
            reference TEXT, commentaires TEXT, scan_name TEXT,
            created_at TEXT DEFAULT (datetime('now','localtime'))
        );
        """)
        con.commit(); con.close()

def db_q(sql, params=()):
    with db_lock:
        con = get_con()
        rows = [dict(r) for r in con.execute(sql, params).fetchall()]
        con.close()
        return rows

def db_x(sql, params=()):
    with db_lock:
        con = get_con()
        con.execute(sql, params)
        con.commit(); con.close()

# ─────────────────────────────────────────
#  EXPORT EXCEL
# ─────────────────────────────────────────
HDR_FILL = PatternFill("solid", fgColor="1A3A52")
HDR_FONT = XLFont(name="Calibri", bold=True, color="FFFFFF", size=11)
ALT_FILL = PatternFill("solid", fgColor="EEF4F8")
THIN     = Border(
    left=Side(style="thin", color="C8DCEA"),
    right=Side(style="thin", color="C8DCEA"),
    top=Side(style="thin", color="C8DCEA"),
    bottom=Side(style="thin", color="C8DCEA"),
)

def style_xl(ws, headers, widths):
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=1, column=i, value=h)
        c.fill = HDR_FILL; c.font = HDR_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = THIN
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

def xl_row(ws, ridx, ncols):
    fill = ALT_FILL if ridx % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
    for c in range(1, ncols+1):
        cell = ws.cell(row=ridx, column=c)
        cell.fill = fill; cell.border = THIN
        cell.alignment = Alignment(vertical="center", wrap_text=True)

def export_arrivee(path):
    rows = db_q("SELECT * FROM arrivee ORDER BY created_at DESC")
    wb = Workbook(); ws = wb.active; ws.title = "Courrier Arrivée"
    hdrs  = ["#","Expéditeur","Objet","Date courrier","Reçu le","Destinataire","Commentaires","Scan"]
    widths= [5,  24,          34,      14,             14,        22,             36,            26]
    style_xl(ws, hdrs, widths)
    for i, r in enumerate(rows, 2):
        vals = [i-1, r["expediteur"], r["objet"], r["date_courrier"],
                r["recu_le"], r["destinataire"], r["commentaires"],
                Path(r["scan_name"]).name if r["scan_name"] else ""]
        for j, v in enumerate(vals, 1):
            ws.cell(row=i, column=j, value=v or "")
        xl_row(ws, i, len(hdrs))
    wb.save(path)

def export_depart(path):
    rows = db_q("SELECT * FROM depart ORDER BY created_at DESC")
    wb = Workbook(); ws = wb.active; ws.title = "Courrier Départ"
    hdrs  = ["#","Agent","Destinataire","Objet","Date envoi","Heure","Référence","Commentaires","Scan"]
    widths= [5,  20,     24,            34,      13,          8,      16,         34,            26]
    style_xl(ws, hdrs, widths)
    for i, r in enumerate(rows, 2):
        vals = [i-1, r["agent_nom"], r["destinataire"], r["objet"],
                r["date_envoi"], r["heure_envoi"], r["reference"],
                r["commentaires"],
                Path(r["scan_name"]).name if r["scan_name"] else ""]
        for j, v in enumerate(vals, 1):
            ws.cell(row=i, column=j, value=v or "")
        xl_row(ws, i, len(hdrs))
    wb.save(path)

# ─────────────────────────────────────────
#  BRIDGE Python ↔ JavaScript
# ─────────────────────────────────────────
class Bridge(QObject):

    def __init__(self, window):
        super().__init__()
        self._win = window

    def _ok(self, **kw):
        return json.dumps({"ok": True, **kw}, ensure_ascii=False)

    def _err(self, msg):
        return json.dumps({"ok": False, "error": msg}, ensure_ascii=False)

    # ── Agents ──
    @pyqtSlot(result=str)
    def getAgents(self):
        rows = db_q("SELECT id,prenom,nom,role FROM agents ORDER BY prenom,nom")
        return json.dumps(rows, ensure_ascii=False)

    @pyqtSlot(str, result=str)
    def createAgent(self, data):
        d = json.loads(data)
        prenom = d.get("prenom","").strip()
        nom    = d.get("nom","").strip()
        role   = d.get("role","agent")
        code   = d.get("code","")
        if not prenom or not nom or len(code) < 4:
            return self._err("Données invalides")
        ag_id = datetime.now().strftime("%Y%m%d%H%M%S%f")
        hcode = hashlib.sha256(code.encode()).hexdigest()
        db_x("INSERT INTO agents (id,prenom,nom,role,code) VALUES (?,?,?,?,?)",
             (ag_id, prenom, nom, role, hcode))
        return self._ok(agent={"id":ag_id,"prenom":prenom,"nom":nom,"role":role})

    @pyqtSlot(str, result=str)
    def login(self, data):
        d = json.loads(data)
        rows = db_q("SELECT * FROM agents WHERE id=?", (d.get("id",""),))
        if not rows:
            return self._err("Compte introuvable")
        ag = rows[0]
        if ag["code"] != hashlib.sha256(d.get("code","").encode()).hexdigest():
            return self._err("Code incorrect")
        return self._ok(agent={"id":ag["id"],"prenom":ag["prenom"],"nom":ag["nom"],"role":ag["role"]})

    # ── Arrivée ──
    @pyqtSlot(result=str)
    def getArrivee(self):
        return json.dumps(db_q("SELECT * FROM arrivee ORDER BY created_at DESC"), ensure_ascii=False)

    @pyqtSlot(str, result=str)
    def saveArrivee(self, data):
        d = json.loads(data)
        eid = datetime.now().strftime("%Y%m%d%H%M%S%f")
        # Si le scan a déjà été copié via drag&drop, on utilise le chemin existant
        scan_path = d.get("scan_path","")
        if scan_path and Path(scan_path).exists():
            scan_name = scan_path
        else:
            scan_name = self._save_scan(d.get("scan_name",""), d.get("scan_data",""))
        db_x("""INSERT INTO arrivee
                (id,expediteur,objet,date_courrier,recu_le,destinataire,commentaires,scan_name,agent)
                VALUES (?,?,?,?,?,?,?,?,?)""",
             (eid, d.get("expediteur",""), d.get("objet",""),
              d.get("date_courrier",""), d.get("recu_le",""),
              d.get("destinataire",""), d.get("commentaires",""),
              scan_name, d.get("agent","")))
        return self._ok(id=eid)

    # ── Départ ──
    @pyqtSlot(result=str)
    def getDepart(self):
        return json.dumps(db_q("SELECT * FROM depart ORDER BY created_at DESC"), ensure_ascii=False)

    @pyqtSlot(str, result=str)
    def saveDepart(self, data):
        d = json.loads(data)
        eid = datetime.now().strftime("%Y%m%d%H%M%S%f")
        scan_path = d.get("scan_path","")
        if scan_path and Path(scan_path).exists():
            scan_name = scan_path
        else:
            scan_name = self._save_scan(d.get("scan_name",""), d.get("scan_data",""))
        db_x("""INSERT INTO depart
                (id,agent_nom,destinataire,objet,date_envoi,heure_envoi,reference,commentaires,scan_name)
                VALUES (?,?,?,?,?,?,?,?,?)""",
             (eid, d.get("agent_nom",""), d.get("destinataire",""),
              d.get("objet",""), d.get("date_envoi",""), d.get("heure_envoi",""),
              d.get("reference",""), d.get("commentaires",""), scan_name))
        return self._ok(id=eid)

    # ── Import Excel ──
    @pyqtSlot(str, result=str)
    def importData(self, data):
        d = json.loads(data)
        rows  = d.get("rows", [])
        scans = d.get("scans", {})
        count = 0
        for i, r in enumerate(rows):
            eid = datetime.now().strftime("%Y%m%d%H%M%S") + str(i).zfill(4)
            sn  = r.get("scan_name","")
            scan_name = self._save_scan(sn, scans.get(sn,"")) if sn in scans else (sn or "")
            db_x("""INSERT OR IGNORE INTO arrivee
                    (id,expediteur,objet,date_courrier,recu_le,destinataire,commentaires,scan_name,agent)
                    VALUES (?,?,?,?,?,?,?,?,?)""",
                 (eid, r.get("expediteur",""), r.get("objet",""),
                  r.get("date_courrier",""), r.get("recu_le",""),
                  r.get("destinataire",""), r.get("commentaires",""),
                  scan_name, "Import"))
            count += 1
        return self._ok(imported=count)

    # ── Export Excel ──
    @pyqtSlot(str, result=str)
    def exportExcel(self, type_):
        default = f"courrier_{type_}_{datetime.now().strftime('%Y%m')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(
            self._win, "Enregistrer le fichier Excel",
            str(Path.home() / default), "Excel (*.xlsx)")
        if not path:
            return self._err("Annulé")
        try:
            if type_ == "arrivee":
                export_arrivee(path)
            else:
                export_depart(path)
            return self._ok(path=path)
        except Exception as e:
            return self._err(str(e))

    # ── Ouvrir un scan (dialogue fichier) ──
    @pyqtSlot(result=str)
    def pickScan(self):
        path, _ = QFileDialog.getOpenFileName(
            self._win, "Choisir un scan",
            str(Path.home()),
            "Fichiers (*.pdf *.png *.jpg *.jpeg *.bmp);;Tous (*)")
        if not path:
            return json.dumps({"name":"","b64":""})
        try:
            data = base64.b64encode(Path(path).read_bytes()).decode()
            return json.dumps({"name": Path(path).name, "b64": data})
        except Exception as e:
            return json.dumps({"name":"","b64":"","error":str(e)})

    # ── Ouvrir Excel pour import ──
    @pyqtSlot(result=str)
    def pickExcel(self):
        path, _ = QFileDialog.getOpenFileName(
            self._win, "Choisir le fichier Excel",
            str(Path.home()), "Excel (*.xlsx *.xls)")
        if not path:
            return json.dumps({"name":"","b64":""})
        try:
            data = base64.b64encode(Path(path).read_bytes()).decode()
            return json.dumps({"name": Path(path).name, "b64": data})
        except Exception as e:
            return json.dumps({"name":"","b64":"","error":str(e)})

    # ── Ouvrir dossier scans pour import ──
    @pyqtSlot(result=str)
    def pickScansFolder(self):
        folder = QFileDialog.getExistingDirectory(
            self._win, "Choisir le dossier des scans", str(Path.home()))
        if not folder:
            return json.dumps({"files":[]})
        files = []
        for p in Path(folder).iterdir():
            if p.suffix.lower() in (".pdf",".png",".jpg",".jpeg",".bmp"):
                try:
                    data = base64.b64encode(p.read_bytes()).decode()
                    files.append({"name": p.name, "b64": data})
                except Exception:
                    pass
        return json.dumps({"files": files})

    # ── Sauvegarder scan sur disque réseau ──
    def _save_scan(self, name, b64):
        if not name or not b64:
            return ""
        try:
            safe = datetime.now().strftime("%Y%m%d_%H%M%S_") + Path(name).name
            dest = SCANS_DIR / safe
            dest.write_bytes(base64.b64decode(b64))
            return str(dest)
        except Exception:
            return ""

    # ── Infos dossier scans ──
    @pyqtSlot(result=str)
    def getScansDir(self):
        return str(SCANS_DIR)

    # ── Suppression arrivée ──
    @pyqtSlot(str, result=str)
    def deleteArrivee(self, rid):
        rows = db_q("SELECT scan_name FROM arrivee WHERE id=?", (rid,))
        if rows and rows[0]["scan_name"]:
            try: Path(rows[0]["scan_name"]).unlink(missing_ok=True)
            except: pass
        db_x("DELETE FROM arrivee WHERE id=?", (rid,))
        return self._ok()

    # ── Suppression départ ──
    @pyqtSlot(str, result=str)
    def deleteDepart(self, rid):
        rows = db_q("SELECT scan_name FROM depart WHERE id=?", (rid,))
        if rows and rows[0]["scan_name"]:
            try: Path(rows[0]["scan_name"]).unlink(missing_ok=True)
            except: pass
        db_x("DELETE FROM depart WHERE id=?", (rid,))
        return self._ok()

    # ── Copier fichier local directement (drag & drop) ──
    @pyqtSlot(str, result=str)
    def copyScanFromPath(self, src_path):
        src = Path(src_path.replace("file:///","").replace("file://",""))
        if not src.exists():
            return json.dumps({"ok": False, "error": "Fichier introuvable"})
        try:
            safe = datetime.now().strftime("%Y%m%d_%H%M%S_") + src.name
            dest = SCANS_DIR / safe
            shutil.copy2(str(src), str(dest))
            return json.dumps({"ok": True, "name": src.name, "path": str(dest)})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    # ── Ouvrir scan avec le programme système ──
    @pyqtSlot(str, result=str)
    def openScan(self, scan_path):
        import subprocess, platform
        try:
            p = Path(scan_path)
            if not p.exists():
                return self._err("Fichier introuvable")
            if platform.system() == "Windows":
                os.startfile(str(p))
            elif platform.system() == "Darwin":
                subprocess.run(["open", str(p)])
            else:
                subprocess.run(["xdg-open", str(p)])
            return self._ok()
        except Exception as e:
            return self._err(str(e))

    # ── Export Excel automatique par année (comme l'ancien logiciel) ──
    @pyqtSlot(result=str)
    def autoExportAnnee(self):
        """Crée/met à jour le fichier Excel de l'année courante automatiquement"""
        annee = datetime.now().year
        path_a = APP_DIR / f"courrier_arrivee_{annee}.xlsx"
        path_d = APP_DIR / f"courrier_depart_{annee}.xlsx"
        try:
            export_arrivee(str(path_a))
            export_depart(str(path_d))
            return self._ok(arrivee=str(path_a), depart=str(path_d))
        except Exception as e:
            return self._err(str(e))

# ─────────────────────────────────────────
#  HTML DE L'APPLICATION
# ─────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Courrier Saint-Magne</title>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root{
  --sky:#e8f4fd;--sky2:#d1e9f9;
  --blue:#2980b9;--blue-d:#1f6b9a;--blue-dd:#155478;
  --navy:#1a3a52;--navy-d:#112637;
  --teal:#1abc9c;--teal-d:#17a589;
  --green:#27ae60;--green-bg:#eafaf1;
  --red:#e74c3c;--red-bg:#fdf2f1;
  --bg:#f0f5f9;--surface:#fff;
  --border:#dde8f0;--border2:#c8dcea;
  --ink:#0f2535;--ink2:#2c4a60;--muted:#6b8496;--muted2:#94aab8;
  --sh-sm:0 2px 6px rgba(15,37,53,.07),0 1px 2px rgba(15,37,53,.05);
  --sh-md:0 6px 20px rgba(15,37,53,.09),0 2px 6px rgba(15,37,53,.06);
  --sh-lg:0 20px 50px rgba(15,37,53,.14),0 4px 14px rgba(15,37,53,.08);
  --r-sm:9px;--r-md:13px;--r-lg:18px;--r-xl:24px;
  --sb:240px;
}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%;overflow:hidden}
body{font-family:'Plus Jakarta Sans',system-ui,sans-serif;background:var(--bg);color:var(--ink);font-size:14px;line-height:1.55;-webkit-font-smoothing:antialiased}

/* ── Loading ── */
#loading{position:fixed;inset:0;z-index:9999;background:linear-gradient(145deg,var(--navy-d),var(--navy) 60%,var(--blue-dd));display:flex;flex-direction:column;align-items:center;justify-content:center;gap:18px;transition:opacity .5s}
.ld-brand{display:flex;align-items:center;gap:14px}
.ld-icon{width:58px;height:58px;border-radius:var(--r-md);background:rgba(255,255,255,.13);display:flex;align-items:center;justify-content:center;font-size:28px;box-shadow:0 4px 16px rgba(0,0,0,.2)}
.ld-title{font-size:22px;font-weight:800;color:#fff}
.ld-sub{font-size:13px;color:rgba(255,255,255,.45);margin-top:2px}
.spin{width:28px;height:28px;border-radius:50%;border:3px solid rgba(255,255,255,.15);border-top-color:rgba(255,255,255,.8);animation:spin .65s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}

/* ── App ── */
#app{display:none;height:100vh}

/* ── Écrans plein page ── */
.fs{display:none;position:fixed;inset:0;z-index:100;background:linear-gradient(145deg,var(--navy-d),var(--navy) 60%,var(--blue-dd));align-items:center;justify-content:center;overflow-y:auto;padding:24px}
.fs.on{display:flex}

/* Carte centrale */
.fc{background:var(--surface);border-radius:var(--r-xl);padding:48px;width:100%;max-width:460px;box-shadow:var(--sh-lg)}
.fc-wide{max-width:540px}
.fc-head{display:flex;align-items:center;gap:14px;margin-bottom:30px}
.fc-head-icon{width:50px;height:50px;border-radius:var(--r-sm);background:linear-gradient(135deg,var(--blue),var(--blue-dd));display:flex;align-items:center;justify-content:center;font-size:22px;box-shadow:0 4px 14px rgba(41,128,185,.35);flex-shrink:0}
.fc-head-h{font-size:18px;font-weight:800;color:var(--ink);line-height:1.2}
.fc-head-s{font-size:12px;color:var(--muted);margin-top:2px}
.fc-div{height:1px;background:var(--border);margin:0 0 26px}
.fc-title{font-size:16px;font-weight:700;color:var(--ink2);margin-bottom:5px}
.fc-sub{font-size:13px;color:var(--muted);margin-bottom:20px;line-height:1.5}

/* ── Champs ── */
.f{margin-bottom:16px}.f:last-child{margin-bottom:0}
.f-row{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.lbl{display:block;font-size:11.5px;font-weight:700;color:var(--muted);letter-spacing:.02em;margin-bottom:6px}
.inp{width:100%;appearance:none;border:1.5px solid var(--border2);border-radius:var(--r-sm);padding:11px 14px;font-size:13.5px;font-family:inherit;font-weight:500;color:var(--ink);background:#f7fbfd;outline:none;transition:all .18s}
.inp::placeholder{color:var(--muted2);font-weight:400}
.inp:focus{border-color:var(--blue);box-shadow:0 0 0 3.5px rgba(41,128,185,.13);background:#fff}
select.inp{background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='16' height='16' viewBox='0 0 24 24' fill='none' stroke='%236b8496' stroke-width='2'%3E%3Cpath d='m6 9 6 6 6-6'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 12px center;padding-right:36px;cursor:pointer}
textarea.inp{resize:vertical;min-height:74px}
.ferr{font-size:12px;color:var(--red);font-weight:600;margin-top:5px;display:none}
.ferr.on{display:block}
.login-err{display:none;background:var(--red-bg);border:1px solid #f5b7b1;border-radius:var(--r-sm);padding:10px 14px;font-size:13px;font-weight:600;color:var(--red);margin-bottom:14px}
.login-err.on{display:block}

/* ── Boutons ── */
.btn{display:inline-flex;align-items:center;justify-content:center;gap:8px;padding:11px 22px;border:none;border-radius:var(--r-sm);font-size:13.5px;font-weight:700;font-family:inherit;cursor:pointer;transition:all .18s;white-space:nowrap;line-height:1}
.btn-full{width:100%;padding:13px 22px;font-size:14.5px}
.btn-sm{padding:7px 14px;font-size:12.5px}
.btn-primary{background:linear-gradient(135deg,var(--blue),var(--blue-d));color:#fff;box-shadow:0 3px 10px rgba(41,128,185,.3)}
.btn-primary:hover{background:linear-gradient(135deg,var(--blue-d),var(--blue-dd));box-shadow:0 4px 14px rgba(41,128,185,.38);transform:translateY(-1px)}
.btn-teal{background:linear-gradient(135deg,var(--teal),var(--teal-d));color:#fff;box-shadow:0 3px 10px rgba(26,188,156,.28)}
.btn-teal:hover{filter:brightness(1.06);transform:translateY(-1px)}
.btn-green{background:var(--green-bg);color:var(--green);border:1.5px solid #b2dfca}
.btn-green:hover{background:#d5f0e4}
.btn-ghost{background:var(--bg);color:var(--muted);border:1.5px solid var(--border)}
.btn-ghost:hover{background:var(--sky);color:var(--ink2);border-color:var(--border2)}
.btn:disabled{opacity:.45;cursor:not-allowed;transform:none!important}
.link{color:var(--blue);font-weight:600;cursor:pointer;font-size:13px}
.link:hover{text-decoration:underline}
.btn-row{display:flex;gap:10px;margin-top:20px}

/* ── Étapes import ── */
.steps{display:flex;margin-bottom:26px;border-radius:var(--r-sm);overflow:hidden;border:1px solid var(--border)}
.step{flex:1;padding:10px 6px;text-align:center;font-size:12px;font-weight:600;color:var(--muted);background:#f7fbfd;border-right:1px solid var(--border);transition:all .2s}
.step:last-child{border-right:none}
.step.active{background:var(--blue);color:#fff}
.step.done{background:var(--green-bg);color:var(--green)}
.step-icon{font-size:16px;display:block;margin-bottom:3px}

.info-pill{background:var(--sky);border:1px solid var(--border2);border-radius:var(--r-sm);padding:12px 15px;font-size:12.5px;color:var(--ink2);line-height:1.6;margin-bottom:16px}
.info-pill strong{color:var(--blue-d)}

.imp-drop{border:2px dashed var(--border2);border-radius:var(--r-md);background:var(--sky);padding:26px 18px;text-align:center;cursor:pointer;transition:all .2s;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:8px;min-height:110px}
.imp-drop:hover,.imp-drop.over{border-color:var(--blue);background:var(--sky2)}
.imp-drop.done{border-color:var(--green);background:var(--green-bg);border-style:solid;cursor:default}
.imp-drop .di{font-size:28px}
.imp-drop .dt{font-size:14px;font-weight:700;color:var(--ink2)}
.imp-drop .ds{font-size:12px;color:var(--muted);line-height:1.5}
.imp-drop.done .dt{color:var(--green)}.imp-drop.done .ds{color:var(--teal-d)}
.prog-w{background:var(--border);border-radius:99px;height:7px;margin:10px 0 4px}
.prog-f{background:var(--blue);border-radius:99px;height:7px;transition:width .3s}
.prog-l{font-size:12px;color:var(--muted);text-align:center}
.imp-sum{background:var(--green-bg);border:1px solid #b2dfca;border-radius:var(--r-sm);padding:16px 18px;font-size:13px;color:var(--ink2);line-height:1.9}
.imp-sum strong{color:var(--green);font-size:15px}

/* ── Sidebar ── */
#ml{display:none;height:100vh;flex-direction:row}
#ml.on{display:flex}
#sb{width:var(--sb);flex-shrink:0;background:linear-gradient(180deg,var(--navy-d),var(--navy));display:flex;flex-direction:column;overflow:hidden}
.sb-top{padding:24px 18px 18px}
.sb-brand{display:flex;align-items:center;gap:11px}
.sb-icon{width:38px;height:38px;border-radius:10px;background:rgba(255,255,255,.13);display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0}
.sb-name{font-size:14.5px;font-weight:800;color:#fff;line-height:1.25}
.sb-ver{font-size:11px;color:rgba(255,255,255,.3);margin-top:1px}
.sb-nav{padding:6px 12px;flex:1}
.ni{display:flex;align-items:center;gap:10px;width:100%;padding:10px 12px;border:none;border-radius:10px;background:transparent;color:rgba(255,255,255,.6);font-size:13.5px;font-weight:600;font-family:inherit;cursor:pointer;text-align:left;transition:all .15s;margin-bottom:2px}
.ni:hover{background:rgba(255,255,255,.09);color:rgba(255,255,255,.9)}
.ni.active{background:rgba(255,255,255,.16);color:#fff;box-shadow:inset 2px 0 0 var(--teal)}
.ni-ico{font-size:16px;width:22px;text-align:center;flex-shrink:0}
.sb-sep{height:1px;background:rgba(255,255,255,.08);margin:6px 14px}
.sb-bot{padding:10px 12px 20px}
.sb-agent{background:rgba(255,255,255,.08);border-radius:11px;padding:12px 14px;margin-bottom:10px;display:flex;align-items:center;gap:10px}
.sb-av{width:34px;height:34px;border-radius:99px;background:linear-gradient(135deg,var(--blue),var(--teal));display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:800;color:#fff;flex-shrink:0}
.sb-an{font-size:13px;font-weight:700;color:#fff;line-height:1.2}
.sb-ar{font-size:11px;color:rgba(255,255,255,.45);margin-top:2px}
.btn-lo{display:flex;align-items:center;gap:9px;width:100%;padding:9px 13px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.1);border-radius:10px;color:rgba(255,255,255,.55);font-size:13px;font-family:inherit;cursor:pointer;transition:all .15s}
.btn-lo:hover{background:rgba(220,60,60,.22);color:#ffb3b3;border-color:rgba(200,50,50,.3)}

/* ── Contenu ── */
#ct{flex:1;overflow-y:auto;display:flex;flex-direction:column;min-width:0}
.pg{display:none;flex-direction:column;min-height:100%}
.pg.on{display:flex}
.pg-hdr{padding:24px 28px 0;flex-shrink:0}
.pg-tr{display:flex;align-items:center;justify-content:space-between;gap:14px;margin-bottom:4px;flex-wrap:wrap}
.pg-title{font-size:22px;font-weight:800;color:var(--ink)}
.pg-sub{font-size:13px;color:var(--muted)}
.pg-act{display:flex;gap:10px;flex-wrap:wrap}
.pg-body{padding:16px 28px 26px;flex:1}

/* ── Cartes ── */
.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--r-md);box-shadow:var(--sh-sm);overflow:hidden}
.card+.card{margin-top:16px}
.ch{padding:13px 18px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;gap:10px}
.ch-t{font-size:13.5px;font-weight:700;color:var(--blue-d);display:flex;align-items:center;gap:8px}
.cb{padding:20px}

/* ── Formulaire ── */
.fg{display:grid;gap:14px}
.fg2{grid-template-columns:1fr 1fr}
.fg3{grid-template-columns:1fr 1fr 1fr}
.s2{grid-column:span 2}.s3{grid-column:span 3}
.fi{width:100%;appearance:none;border:1.5px solid var(--border2);border-radius:var(--r-sm);padding:10px 13px;font-size:13.5px;font-family:inherit;font-weight:500;color:var(--ink);background:#f8fbfd;outline:none;transition:all .18s}
.fi::placeholder{color:var(--muted2);font-weight:400}
.fi:focus{border-color:var(--blue);box-shadow:0 0 0 3.5px rgba(41,128,185,.12);background:#fff}
textarea.fi{resize:vertical;min-height:70px}
.fa{display:flex;justify-content:flex-end;gap:10px;margin-top:16px}

/* ── Zone scan ── */
.dz{border:2px dashed var(--border2);border-radius:var(--r-sm);background:var(--sky);padding:13px 16px;display:flex;align-items:center;gap:12px;cursor:pointer;transition:all .18s;min-height:58px}
.dz:hover,.dz.over{border-color:var(--blue);background:var(--sky2)}
.dz.has{border-color:var(--teal);background:#edfaf7;border-style:solid}
.dz-ico{font-size:20px;flex-shrink:0}
.dz-txt{font-size:13px;color:var(--muted)}
.dz.has .dz-txt{color:var(--teal-d);font-weight:600}

/* ── Tableau ── */
.tw{overflow-x:auto;max-height:320px;overflow-y:auto}
table{width:100%;border-collapse:collapse;font-size:13px}
thead{position:sticky;top:0;z-index:2}
thead th{background:var(--navy);color:rgba(255,255,255,.85);font-size:11px;font-weight:700;letter-spacing:.04em;padding:11px 14px;text-align:left;white-space:nowrap;border-right:1px solid rgba(255,255,255,.07)}
thead th:last-child{border-right:none}
tbody tr{border-bottom:1px solid #eef4f8;transition:background .1s}
tbody tr:last-child{border-bottom:none}
tbody tr:hover{background:#f2f8fc}
tbody tr.mine{background:#edfaf7}
tbody tr.mine:hover{background:#e0f6f2}
tbody td{padding:10px 14px;vertical-align:middle}
.tm{color:var(--muted)}.ts{font-weight:600;color:var(--ink2)}
.badge{display:inline-flex;align-items:center;gap:4px;padding:3px 9px;border-radius:99px;font-size:11px;font-weight:700}
.bb{background:var(--sky2);color:var(--blue-d)}
.bt{background:#e0f7f3;color:var(--teal-d)}
.bn{background:rgba(26,58,82,.1);color:var(--navy)}
.bg{background:var(--green-bg);color:var(--green)}

/* ── Sous-onglets ── */
.itabs{display:flex;gap:6px;margin-bottom:14px}
.itab{padding:8px 18px;border-radius:var(--r-sm);font-size:13px;font-weight:700;cursor:pointer;background:var(--surface);color:var(--muted);border:1.5px solid var(--border);transition:all .15s}
.itab.on{background:var(--navy);color:#fff;border-color:var(--navy);box-shadow:var(--sh-sm)}
.itp{display:none}.itp.on{display:block}

/* ── Bouton supprimer ── */
.btn-del{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border:1.5px solid #f5b7b1;border-radius:6px;background:#fdf2f1;color:var(--red);font-size:11.5px;font-weight:700;font-family:inherit;cursor:pointer;transition:all .15s}
.btn-del:hover{background:var(--red);color:#fff;border-color:var(--red)}
/* ── Lien scan cliquable ── */
.scan-link{display:inline-flex;align-items:center;gap:5px;padding:3px 9px;border-radius:99px;font-size:11px;font-weight:700;background:var(--sky2);color:var(--blue-d);cursor:pointer;border:none;font-family:inherit;transition:all .15s}
.scan-link:hover{background:var(--blue-d);color:#fff}
/* ── Empty ── */
.empty{text-align:center;padding:44px 20px;color:var(--muted)}
.empty-ico{font-size:36px;margin-bottom:12px;opacity:.45}
.empty p{font-size:13.5px}

/* ── Statusbar ── */
.stbar{padding:7px 28px;flex-shrink:0;background:var(--sky);border-top:1px solid var(--border);font-size:12px;color:var(--muted);display:flex;align-items:center;gap:8px}
.stdot{width:7px;height:7px;border-radius:99px;background:var(--teal);flex-shrink:0}

/* ── Toast ── */
#toast{position:fixed;bottom:22px;left:50%;transform:translateX(-50%);padding:12px 22px;border-radius:var(--r-sm);font-size:13.5px;font-weight:700;box-shadow:var(--sh-md);z-index:9999;pointer-events:none;opacity:0;transition:opacity .25s;max-width:440px;text-align:center;background:var(--navy);color:#fff}
#toast.on{opacity:1}
#toast.s{background:#1a6b42}
#toast.e{background:#a93226}
#toast.i{background:var(--blue-d)}
</style>
</head>
<body>

<div id="loading">
  <div class="ld-brand">
    <div class="ld-icon">📬</div>
    <div><div class="ld-title">Courrier Saint-Magne</div><div class="ld-sub">Mairie de Saint-Magne · v9</div></div>
  </div>
  <div class="spin"></div>
</div>

<div id="app">

<!-- ══ CONNEXION ══ -->
<div class="fs" id="sc-login">
  <div class="fc">
    <div class="fc-head">
      <div class="fc-head-icon">📬</div>
      <div><div class="fc-head-h">Courrier Saint-Magne</div><div class="fc-head-s">Mairie de Saint-Magne · Version 9</div></div>
    </div>
    <div class="fc-div"></div>
    <div class="fc-title">Connexion à votre espace</div>
    <div class="fc-sub">Sélectionnez votre nom et entrez votre code personnel.</div>
    <div class="login-err" id="lerr"></div>
    <div class="f"><label class="lbl">Votre nom</label><select class="inp" id="sel-ag"></select></div>
    <div class="f"><label class="lbl">Code d'accès</label><input class="inp" type="password" id="inp-code" placeholder="Entrez votre code…"></div>
    <div style="margin-top:20px"><button class="btn btn-primary btn-full" onclick="tryLogin()">Se connecter</button></div>
    <p style="text-align:center;font-size:12.5px;color:var(--muted);margin-top:16px">Première fois ? <span class="link" onclick="showSc('sc-reg')">Créer votre compte</span></p>
  </div>
</div>

<!-- ══ CRÉATION COMPTE ══ -->
<div class="fs" id="sc-reg">
  <div class="fc">
    <div class="fc-head">
      <div class="fc-head-icon">👤</div>
      <div><div class="fc-head-h">Créer votre compte</div><div class="fc-head-s">Accès immédiat après création</div></div>
    </div>
    <div class="fc-div"></div>
    <div class="f f-row">
      <div><label class="lbl">Prénom *</label><input class="inp" type="text" id="rp" placeholder="Marie"></div>
      <div><label class="lbl">Nom *</label><input class="inp" type="text" id="rn" placeholder="Dupont"></div>
    </div>
    <div class="f"><label class="lbl">Rôle</label>
      <select class="inp" id="rr">
        <option value="agent">Agent</option>
        <option value="admin">Administrateur (gère le courrier arrivée)</option>
      </select>
    </div>
    <div class="f"><label class="lbl">Choisissez votre code *</label><input class="inp" type="password" id="rc1" placeholder="Minimum 4 caractères"></div>
    <div class="f"><label class="lbl">Confirmez le code *</label><input class="inp" type="password" id="rc2" placeholder="Répétez le code"><div class="ferr" id="rerr"></div></div>
    <div class="btn-row"><button class="btn btn-ghost" onclick="showSc('sc-login')" style="flex:1">← Retour</button><button class="btn btn-primary" onclick="doCreateAccount()" style="flex:2">Créer mon compte</button></div>
  </div>
</div>

<!-- ══ IMPORT ══ -->
<div class="fs" id="sc-import">
  <div class="fc fc-wide">
    <div class="fc-head">
      <div class="fc-head-icon">📥</div>
      <div><div class="fc-head-h">Import des archives existantes</div><div class="fc-head-s">Importez vos courriers 2025 et 2026</div></div>
    </div>
    <div class="fc-div"></div>
    <div class="steps">
      <div class="step active" id="si1"><span class="step-icon">📊</span>Fichier Excel</div>
      <div class="step" id="si2"><span class="step-icon">🗂️</span>Dossier scans</div>
      <div class="step" id="si3"><span class="step-icon">✅</span>Confirmation</div>
    </div>

    <div id="is1">
      <div class="info-pill"><strong>Cliquez pour sélectionner votre fichier Excel.</strong><br>Toutes les feuilles sont lues automatiquement (2025, 2026…). Colonnes reconnues : Expéditeur, Objet, Date du Courrier, Reçu le, Destinataire, Commentaires, Fichier scanné.</div>
      <div class="imp-drop" id="id-xl" onclick="pickExcel()">
        <div class="di">📊</div>
        <div class="dt" id="idt-xl">Cliquer pour sélectionner le fichier Excel</div>
        <div class="ds" id="ids-xl">Formats .xlsx et .xls acceptés</div>
      </div>
      <div id="xlp" style="display:none;margin-top:12px"><div class="prog-w"><div class="prog-f" id="xlb" style="width:0%"></div></div><div class="prog-l" id="xll">Lecture…</div></div>
      <div style="display:flex;justify-content:space-between;align-items:center;margin-top:16px">
        <button class="btn btn-ghost btn-sm" onclick="skipImport()">Passer et démarrer →</button>
        <button class="btn btn-primary btn-sm" id="bi1" onclick="goStep(2)" disabled>Étape suivante →</button>
      </div>
    </div>

    <div id="is2" style="display:none">
      <div class="info-pill"><strong>Cliquez pour sélectionner votre dossier de scans.</strong><br>Tous les fichiers PDF et images du dossier seront chargés automatiquement.<br><span style="color:var(--teal-d);font-weight:600">Les scans seront reliés aux courriers par leur nom de fichier.</span></div>
      <div class="imp-drop" id="id-sc" onclick="pickFolder()">
        <div class="di">🗂️</div>
        <div class="dt" id="idt-sc">Cliquer pour sélectionner le dossier des scans</div>
        <div class="ds" id="ids-sc">Sélectionnez le dossier contenant tous vos scans</div>
      </div>
      <div id="scp" style="display:none;margin-top:12px"><div class="prog-w"><div class="prog-f" id="scb" style="width:0%"></div></div><div class="prog-l" id="scl">Chargement…</div></div>
      <div style="display:flex;justify-content:space-between;align-items:center;margin-top:16px">
        <button class="btn btn-ghost btn-sm" onclick="goStep(1)">← Retour</button>
        <div style="display:flex;gap:8px">
          <button class="btn btn-ghost btn-sm" onclick="goStep(3)">Passer →</button>
          <button class="btn btn-primary btn-sm" id="bi2" onclick="goStep(3)" disabled>Étape suivante →</button>
        </div>
      </div>
    </div>

    <div id="is3" style="display:none">
      <div id="imp-sum"></div>
      <div style="display:flex;justify-content:space-between;align-items:center;margin-top:20px">
        <button class="btn btn-ghost btn-sm" onclick="goStep(2)">← Retour</button>
        <button class="btn btn-teal" onclick="doImport()">✅ Importer et démarrer</button>
      </div>
    </div>
  </div>
</div>

<!-- ══ APP ══ -->
<div id="ml">
  <div id="sb">
    <div class="sb-top">
      <div class="sb-brand">
        <div class="sb-icon">📬</div>
        <div><div class="sb-name">Courrier<br>Saint-Magne</div><div class="sb-ver">v9.0</div></div>
      </div>
    </div>
    <div class="sb-nav" id="sb-nav"></div>
    <div class="sb-sep"></div>
    <div class="sb-bot">
      <div class="sb-agent"><div class="sb-av" id="sb-av">?</div><div><div class="sb-an" id="sb-an">—</div><div class="sb-ar" id="sb-ar">—</div></div></div>
      <button class="btn-lo" onclick="doLogout()">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
        <span>Se déconnecter</span>
      </button>
    </div>
  </div>

  <div id="ct">
    <!-- ARRIVÉE -->
    <div class="pg" id="pg-a">
      <div class="pg-hdr">
        <div class="pg-tr">
          <div><div class="pg-title">📥 Courrier Arrivée</div><div class="pg-sub">Enregistrement et archivage du courrier reçu</div></div>
          <div class="pg-act"><button class="btn btn-green btn-sm" onclick="doExport('arrivee')">📊 Exporter Excel</button></div>
        </div>
      </div>
      <div class="pg-body">
        <div class="card">
          <div class="ch"><span class="ch-t">✉️ Nouveau courrier arrivée</span><button class="btn btn-ghost btn-sm" onclick="clearA()">Réinitialiser</button></div>
          <div class="cb">
            <div class="fg fg2">
              <div><label class="lbl">Expéditeur *</label><input type="text" class="fi" id="ae" placeholder="Organisme ou personne expéditrice"></div>
              <div><label class="lbl">Destinataire</label><input type="text" class="fi" id="ad" placeholder="Service ou agent destinataire"></div>
              <div class="s2"><label class="lbl">Objet *</label><input type="text" class="fi" id="ao" placeholder="Objet du courrier"></div>
              <div><label class="lbl">Date du courrier</label><input type="date" class="fi" id="adc"></div>
              <div><label class="lbl">Reçu le</label><input type="date" class="fi" id="adr"></div>
              <div class="s2"><label class="lbl">Commentaires</label><textarea class="fi" id="ac" placeholder="Références, actions à mener…"></textarea></div>
              <div class="s2">
                <label class="lbl">Scan / Pièce jointe</label>
                <div class="dz" id="adz" onclick="pickScanFor('arrivee')"
                  ondragover="dzO(event,'adz')" ondragleave="dzL('adz')" ondrop="dzD(event,'adz','arrivee')">
                  <div class="dz-ico">📎</div>
                  <div>
                    <div class="dz-txt" id="adzt">Glisser-déposer un PDF ou cliquer pour importer</div>
                    <div style="font-size:11px;color:var(--muted2);margin-top:3px">Le fichier sera copié dans le dossier scans réseau</div>
                  </div>
                </div>
              </div>
            </div>
            <div class="fa"><button class="btn btn-primary" onclick="saveA()">💾 Archiver le courrier</button></div>
          </div>
        </div>
        <div class="card" style="margin-top:16px">
          <div class="ch"><span class="ch-t">📚 Archives <span class="badge bb" id="a-cnt">0</span></span></div>
          <div class="tw"><table><thead><tr><th>#</th><th>Expéditeur</th><th>Objet</th><th>Date courrier</th><th>Reçu le</th><th>Destinataire</th><th>Commentaires</th><th>Scan</th><th>Action</th></tr></thead><tbody id="tb-a"></tbody></table></div>
        </div>
      </div>
      <div class="stbar"><div class="stdot"></div><span id="st-a">Prêt</span></div>
    </div>

    <!-- DÉPART -->
    <div class="pg" id="pg-d">
      <div class="pg-hdr">
        <div class="pg-tr">
          <div><div class="pg-title">📤 Courrier Départ</div><div class="pg-sub" id="dsub">Envoi et suivi du courrier sortant</div></div>
          <div class="pg-act"><button class="btn btn-green btn-sm" onclick="doExport('depart')">📊 Exporter Excel</button></div>
        </div>
      </div>
      <div class="pg-body">
        <div class="card">
          <div class="ch"><span class="ch-t">📤 Envoyer un courrier</span><button class="btn btn-ghost btn-sm" onclick="clearD()">Réinitialiser</button></div>
          <div class="cb">
            <div class="fg fg2">
              <div><label class="lbl">Destinataire *</label><input type="text" class="fi" id="dd" placeholder="Nom, service ou organisme destinataire"></div>
              <div><label class="lbl">Référence / N° dossier</label><input type="text" class="fi" id="dr" placeholder="Référence ou numéro de dossier"></div>
              <div class="s2"><label class="lbl">Objet *</label><input type="text" class="fi" id="do" placeholder="Objet du courrier envoyé"></div>
              <div><label class="lbl">Date d'envoi</label><input type="date" class="fi" id="ddate"></div>
              <div><label class="lbl">Heure d'envoi</label><input type="time" class="fi" id="dheure"></div>
              <div class="s2"><label class="lbl">Commentaires</label><textarea class="fi" id="dc" placeholder="Notes, suivi, accusé de réception attendu…"></textarea></div>
              <div class="s2">
                <label class="lbl">Scan / Pièce jointe</label>
                <div class="dz" id="ddz" onclick="pickScanFor('depart')"
                  ondragover="dzO(event,'ddz')" ondragleave="dzL('ddz')" ondrop="dzD(event,'ddz','depart')">
                  <div class="dz-ico">📎</div>
                  <div>
                    <div class="dz-txt" id="ddzt">Glisser-déposer un PDF ou cliquer pour importer</div>
                    <div style="font-size:11px;color:var(--muted2);margin-top:3px">Le fichier sera copié dans le dossier scans réseau</div>
                  </div>
                </div>
              </div>
            </div>
            <div class="fa"><button class="btn btn-primary" onclick="saveD()">📤 Enregistrer l'envoi</button></div>
          </div>
        </div>
        <div style="margin-top:16px">
          <div class="itabs">
            <div class="itab on" id="it-my" onclick="switchIt('my')">📤 Mes envois</div>
            <div class="itab" id="it-all" onclick="switchIt('all')">👥 Tous les envois</div>
          </div>
          <div class="itp on" id="itp-my">
            <div class="card"><div class="ch"><span class="ch-t">Mes envois <span class="badge bt" id="d-cmy">0</span></span></div>
            <div class="tw"><table><thead><tr><th>#</th><th>Destinataire</th><th>Objet</th><th>Date</th><th>Heure</th><th>Référence</th><th>Commentaires</th><th>Scan</th><th>Action</th></tr></thead><tbody id="tb-dmy"></tbody></table></div></div>
          </div>
          <div class="itp" id="itp-all">
            <div class="card"><div class="ch"><span class="ch-t">Tous les envois <span class="badge bt" id="d-call">0</span></span></div>
            <div class="tw"><table><thead><tr><th>#</th><th>Agent</th><th>Destinataire</th><th>Objet</th><th>Date</th><th>Heure</th><th>Référence</th><th>Commentaires</th><th>Scan</th><th>Action</th></tr></thead><tbody id="tb-dall"></tbody></table></div></div>
          </div>
        </div>
      </div>
      <div class="stbar"><div class="stdot"></div><span id="st-d">Prêt</span></div>
    </div>
  </div>
</div>
</div>

<div id="toast"></div>

<script src="qrc:///qtwebchannel/qwebchannel.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script>
let py = null;
let agent = null;
let allA = [], allD = [];
let scanData = { arrivee: null, depart: null };
let impRows = [], impScans = {};
let refreshT = null;

// ── Init WebChannel ──
new QWebChannel(qt.webChannelTransport, ch => {
  py = ch.objects.bridge;
  init();
});

async function call(fn, ...args) {
  return new Promise(res => {
    if (args.length) py[fn](args[0], r => res(typeof r === 'string' ? JSON.parse(r) : r));
    else py[fn](r => res(typeof r === 'string' ? JSON.parse(r) : r));
  });
}

async function init() {
  await loadAgents();
  document.getElementById('inp-code').addEventListener('keydown', e => { if (e.key==='Enter') tryLogin(); });
  document.getElementById('rc2').addEventListener('keydown', e => { if (e.key==='Enter') doCreateAccount(); });
  // Export Excel automatique à l'ouverture
  try { await call('autoExportAnnee'); } catch(e) {}
  setTimeout(() => {
    document.getElementById('loading').style.opacity = '0';
    setTimeout(() => {
      document.getElementById('loading').style.display = 'none';
      document.getElementById('app').style.display = 'block';
      showSc('sc-login');
    }, 450);
  }, 600);
}

// ── Agents ──
let agents = [];
async function loadAgents() {
  agents = await call('getAgents');
  buildSel();
}
function buildSel() {
  const sel = document.getElementById('sel-ag');
  sel.innerHTML = '';
  if (!agents.length) { sel.innerHTML = '<option value="">— Aucun compte, créez le vôtre —</option>'; return; }
  agents.forEach(a => { const o=document.createElement('option'); o.value=a.id; o.textContent=a.prenom+' '+a.nom; sel.appendChild(o); });
}

async function doCreateAccount() {
  const p=v('rp'),n=v('rn'),r=document.getElementById('rr').value,c1=v('rc1'),c2=v('rc2');
  const e=document.getElementById('rerr'); e.classList.remove('on');
  if(!p||!n){e.textContent='Prénom et nom obligatoires.';e.classList.add('on');return}
  if(c1.length<4){e.textContent='Code minimum 4 caractères.';e.classList.add('on');return}
  if(c1!==c2){e.textContent='Les codes ne correspondent pas.';e.classList.add('on');return}
  const res = await call('createAgent', JSON.stringify({prenom:p,nom:n,role:r,code:c1}));
  if(!res.ok){e.textContent=res.error;e.classList.add('on');return}
  agents.push(res.agent); buildSel();
  document.getElementById('sel-ag').value = res.agent.id;
  toast('Compte créé avec succès !','s');
  ['rp','rn','rc1','rc2'].forEach(id=>document.getElementById(id).value='');
  showSc('sc-login');
}

// ── Connexion ──
async function tryLogin() {
  const agId=document.getElementById('sel-ag').value, code=v('inp-code');
  const err=document.getElementById('lerr'); err.classList.remove('on');
  if(!agId){err.textContent="Créez d'abord votre compte.";err.classList.add('on');return}
  const res = await call('login', JSON.stringify({id:agId,code}));
  document.getElementById('inp-code').value='';
  if(!res.ok){err.textContent=res.error;err.classList.add('on');return}
  agent = res.agent;
  const imported = localStorage.getItem('sm_imp');
  if(agent.role==='admin'&&!imported) showSc('sc-import');
  else startApp();
}

function doLogout() {
  agent=null; clearInterval(refreshT);
  document.getElementById('ml').classList.remove('on');
  showSc('sc-login');
}

// ── App ──
function startApp() {
  hideSc(); document.getElementById('ml').classList.add('on');
  const init=(agent.prenom[0]+agent.nom[0]).toUpperCase();
  document.getElementById('sb-av').textContent=init;
  document.getElementById('sb-an').textContent=agent.prenom+' '+agent.nom;
  document.getElementById('sb-ar').textContent=agent.role==='admin'?'⭐ Administrateur':'👤 Agent';
  document.getElementById('dsub').textContent='Connecté : '+agent.prenom+' '+agent.nom;
  buildNav(); setDates(); loadData();
  refreshT = setInterval(loadData, 15000);
  switchPg(agent.role==='admin'?'a':'d');
}

function buildNav() {
  const nav=document.getElementById('sb-nav'); nav.innerHTML='';
  const items=[];
  if(agent.role==='admin') items.push({id:'a',ico:'📥',lbl:'Courrier Arrivée'});
  items.push({id:'d',ico:'📤',lbl:'Courrier Départ'});
  items.forEach(it=>{
    const b=document.createElement('button'); b.className='ni'; b.id='ni-'+it.id;
    b.innerHTML=`<span class="ni-ico">${it.ico}</span><span>${it.lbl}</span>`;
    b.onclick=()=>switchPg(it.id); nav.appendChild(b);
  });
}

function switchPg(id) {
  document.querySelectorAll('.pg').forEach(p=>p.classList.remove('on'));
  document.querySelectorAll('.ni').forEach(b=>b.classList.remove('active'));
  const p=document.getElementById('pg-'+id); if(p) p.classList.add('on');
  const b=document.getElementById('ni-'+id); if(b) b.classList.add('active');
}
function switchIt(t) {
  ['my','all'].forEach(x=>{
    document.getElementById('it-'+x).classList.toggle('on',x===t);
    document.getElementById('itp-'+x).classList.toggle('on',x===t);
  });
}
function setDates() {
  const today=new Date().toISOString().split('T')[0], now=new Date().toTimeString().slice(0,5);
  ['adc','adr','ddate'].forEach(id=>{const e=document.getElementById(id);if(e)e.value=today});
  const h=document.getElementById('dheure');if(h)h.value=now;
}

// ── Données ──
async function loadData() {
  if(agent.role==='admin') allA = await call('getArrivee');
  allD = await call('getDepart');
  renderA(); renderD();
}

// ── Sauvegarde ──
async function saveA() {
  const exp=v('ae'),obj=v('ao');
  if(!exp||!obj){toast('Expéditeur et objet obligatoires','e');return}
  const btn=document.querySelector('#pg-a .btn-primary'); btn.disabled=true;
  const sd=scanData.arrivee;
  const res = await call('saveArrivee', JSON.stringify({
    expediteur:exp, objet:obj,
    destinataire:v('ad'), date_courrier:v('adc'), recu_le:v('adr'),
    commentaires:v('ac'), agent:agent.prenom+' '+agent.nom,
    scan_name:sd?.name||'', scan_data:sd?.b64||'',
    scan_path:sd?.path||''   // chemin réseau si déjà copié
  }));
  await loadData(); clearA();
  // Export auto à chaque archivage
  try { await call('autoExportAnnee'); } catch(e){}
  document.getElementById('st-a').textContent='✅ Courrier archivé — '+exp;
  toast('Courrier arrivée archivé','s'); btn.disabled=false;
}

async function saveD() {
  const dest=v('dd'),obj=v('do');
  if(!dest||!obj){toast('Destinataire et objet obligatoires','e');return}
  const btn=document.querySelector('#pg-d .btn-primary'); btn.disabled=true;
  const sd=scanData.depart;
  const res = await call('saveDepart', JSON.stringify({
    agent_nom:agent.prenom+' '+agent.nom,
    destinataire:dest, objet:obj, reference:v('dr'),
    date_envoi:v('ddate'), heure_envoi:v('dheure'),
    commentaires:v('dc'),
    scan_name:sd?.name||'', scan_data:sd?.b64||'',
    scan_path:sd?.path||''
  }));
  await loadData(); clearD();
  // Export auto à chaque envoi
  try { await call('autoExportAnnee'); } catch(e){}
  document.getElementById('st-d').textContent='✅ Envoi enregistré — '+v('ddate')+' à '+v('dheure');
  toast('Courrier départ enregistré','s'); btn.disabled=false;
}

// ── Rendu ──
function fd(s){if(!s)return'';const m=s.match(/^(\d{4})-(\d{2})-(\d{2})$/);return m?`${m[3]}/${m[2]}/${m[1]}`:s}
function esc(s){return(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}
function tr(s,n=40){return s&&s.length>n?s.slice(0,n)+'…':(s||'')}

function scanCell(scan_name) {
  if (!scan_name) return '<span style="color:#ccc">—</span>';
  const fname = scan_name.split(/[\/\\]/).pop();
  return `<button class="scan-link" onclick="ouvrirScan('${scan_name.replace(/\\/g,'\\\\').replace(/'/g,"\\'")}')">📎 ${esc(fname)}</button>`;
}

function renderA() {
  const tb=document.getElementById('tb-a');
  document.getElementById('a-cnt').textContent=allA.length;
  if(!allA.length){tb.innerHTML=`<tr><td colspan="9"><div class="empty"><div class="empty-ico">📭</div><p>Aucun courrier arrivée</p></div></td></tr>`;return}
  tb.innerHTML=allA.map((r,i)=>`
    <tr>
      <td class="tm">${allA.length-i}</td>
      <td class="ts">${esc(r.expediteur)}</td>
      <td>${esc(tr(r.objet))}</td>
      <td class="tm">${fd(r.date_courrier)}</td>
      <td class="tm">${fd(r.recu_le)}</td>
      <td>${esc(r.destinataire||'')}</td>
      <td class="tm">${esc(tr(r.commentaires))}</td>
      <td>${scanCell(r.scan_name)}</td>
      <td><button class="btn-del" onclick="supprimerArrivee('${r.id}')">🗑 Supprimer</button></td>
    </tr>`).join('');
}

async function supprimerArrivee(id) {
  if (!confirm('Supprimer ce courrier arrivée ? Cette action est irréversible.')) return;
  await call('deleteArrivee', id);
  await loadData();
  toast('Courrier arrivée supprimé','i');
}

async function supprimerDepart(id) {
  if (!confirm('Supprimer ce courrier départ ? Cette action est irréversible.')) return;
  await call('deleteDepart', id);
  await loadData();
  toast('Courrier départ supprimé','i');
}

async function ouvrirScan(path) {
  const res = await call('openScan', path);
  if (!res.ok) toast('❌ ' + res.error, 'e');
}

function renderD() {
  const me=agent.prenom+' '+agent.nom;
  const mine=allD.filter(r=>r.agent_nom===me);
  document.getElementById('d-cmy').textContent=mine.length;
  document.getElementById('d-call').textContent=allD.length;
  const tbMy=document.getElementById('tb-dmy');
  if(!mine.length){tbMy.innerHTML=`<tr><td colspan="9"><div class="empty"><div class="empty-ico">📤</div><p>Aucun envoi pour l'instant</p></div></td></tr>`}
  else tbMy.innerHTML=mine.map((r,i)=>`
    <tr>
      <td class="tm">${mine.length-i}</td>
      <td class="ts">${esc(r.destinataire)}</td>
      <td>${esc(tr(r.objet))}</td>
      <td class="tm">${fd(r.date_envoi)}</td>
      <td class="tm">${esc(r.heure_envoi||'')}</td>
      <td class="tm">${esc(r.reference||'')}</td>
      <td class="tm">${esc(tr(r.commentaires))}</td>
      <td>${scanCell(r.scan_name)}</td>
      <td><button class="btn-del" onclick="supprimerDepart('${r.id}')">🗑 Supprimer</button></td>
    </tr>`).join('');
  const tbAll=document.getElementById('tb-dall');
  if(!allD.length){tbAll.innerHTML=`<tr><td colspan="10"><div class="empty"><div class="empty-ico">📤</div><p>Aucun courrier départ</p></div></td></tr>`}
  else tbAll.innerHTML=allD.map((r,i)=>`
    <tr class="${r.agent_nom===me?'mine':''}">
      <td class="tm">${allD.length-i}</td>
      <td><span class="badge bn">${esc(r.agent_nom)}</span></td>
      <td class="ts">${esc(r.destinataire)}</td>
      <td>${esc(tr(r.objet))}</td>
      <td class="tm">${fd(r.date_envoi)}</td>
      <td class="tm">${esc(r.heure_envoi||'')}</td>
      <td class="tm">${esc(r.reference||'')}</td>
      <td class="tm">${esc(tr(r.commentaires))}</td>
      <td>${scanCell(r.scan_name)}</td>
      <td><button class="btn-del" onclick="supprimerDepart('${r.id}')">🗑 Supprimer</button></td>
    </tr>`).join('');
}

// ── Reset ──
function clearA(){
  ['ae','ad','ao','ac'].forEach(id=>document.getElementById(id).value='');
  const t=new Date().toISOString().split('T')[0];
  document.getElementById('adc').value=t; document.getElementById('adr').value=t;
  scanData.arrivee=null; resetDz('adz','adzt');
}
function clearD(){
  ['dd','dr','do','dc'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('ddate').value=new Date().toISOString().split('T')[0];
  document.getElementById('dheure').value=new Date().toTimeString().slice(0,5);
  scanData.depart=null; resetDz('ddz','ddzt');
}
function resetDz(id,tid){
  document.getElementById(id).classList.remove('has','over');
  document.getElementById(tid).textContent='Glisser-déposer ou cliquer pour importer';
}

// ── Scan ──
async function pickScanFor(type) {
  const res = await call('pickScan');
  if(!res.name) return;
  scanData[type]={name:res.name,b64:res.b64};
  const dzId=type==='arrivee'?'adz':'ddz', txtId=type==='arrivee'?'adzt':'ddzt';
  document.getElementById(dzId).classList.add('has');
  document.getElementById(txtId).textContent='✅ '+res.name;
}
function dzO(e,id){e.preventDefault();document.getElementById(id).classList.add('over')}
function dzL(id){document.getElementById(id).classList.remove('over')}
async function dzD(e,id,type){
  e.preventDefault();dzL(id);
  const f=e.dataTransfer.files[0]; if(!f) return;
  const txtId=type==='arrivee'?'adzt':'ddzt';
  document.getElementById(txtId).textContent='⏳ Copie en cours…';
  // Essai copie directe via chemin local (drag depuis l'explorateur Windows)
  const localPath = e.dataTransfer.getData('text/plain') || '';
  if(localPath && localPath.match(/\.(pdf|png|jpg|jpeg|bmp)$/i)) {
    const res = await call('copyScanFromPath', localPath);
    if(res.ok) {
      scanData[type]={name:res.name, path:res.path, b64:''};
      document.getElementById(id).classList.add('has');
      document.getElementById(txtId).textContent='✅ '+res.name+' (copié sur le réseau)';
      return;
    }
  }
  // Fallback : lecture base64 via FileReader
  const r=new FileReader();
  r.onload=()=>{
    const b64=r.result.split(',')[1];
    scanData[type]={name:f.name,b64,path:''};
    document.getElementById(id).classList.add('has');
    document.getElementById(txtId).textContent='✅ '+f.name;
  };
  r.readAsDataURL(f);
}

// ── Import ──
async function pickExcel() {
  const res = await call('pickExcel');
  if(!res.name) return;
  document.getElementById('xlp').style.display='block';
  document.getElementById('xll').textContent='Lecture en cours…';
  document.getElementById('xlb').style.width='15%';
  try {
    const buf=Uint8Array.from(atob(res.b64),c=>c.charCodeAt(0)).buffer;
    const wb=XLSX.read(buf,{type:'array'});
    impRows=[];
    wb.SheetNames.forEach(sh=>{
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[sh],{defval:''});
      rows.forEach(row=>{
        const m={
          expediteur:row['Expediteur']||row['Expéditeur']||row['expediteur']||'',
          objet:row['Objet']||row['objet']||'',
          date_courrier:xlDate(row['Date du Courrier']||row['date_courrier']||''),
          recu_le:xlDate(row['Reçu le']||row['Recu le']||row['recu_le']||''),
          destinataire:row['Destinataire']||row['destinataire']||'',
          commentaires:row['Commentaires']||row['commentaires']||'',
          scan_name:row['Fichier scanné']||row['Fichier scanne']||row['scan']||'',
        };
        if(m.expediteur||m.objet) impRows.push(m);
      });
    });
    document.getElementById('xlb').style.width='100%';
    document.getElementById('xll').textContent=`✅ ${impRows.length} courriers lus (${wb.SheetNames.length} feuille(s))`;
    document.getElementById('id-xl').classList.add('done');
    document.getElementById('idt-xl').textContent='✅ '+res.name;
    document.getElementById('ids-xl').textContent=impRows.length+' courriers trouvés';
    document.getElementById('bi1').disabled=false;
  } catch(err) {
    document.getElementById('xll').textContent='❌ Erreur : '+err.message;
  }
}

function xlDate(v){
  if(!v)return'';
  if(typeof v==='number'){const d=new Date(Math.round((v-25569)*86400*1000));return d.toISOString().split('T')[0]}
  const m=(v+'').match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if(m){const yr=m[3].length===2?'20'+m[3]:m[3];return`${yr}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`}
  return v+'';
}

async function pickFolder() {
  const res = await call('pickScansFolder');
  const files = res.files || [];
  if(!files.length) return;
  impScans={};
  document.getElementById('scp').style.display='block';
  let done=0;
  files.forEach(f=>{ impScans[f.name]=f.b64; done++; });
  document.getElementById('scb').style.width='100%';
  document.getElementById('scl').textContent=`✅ ${files.length} scans chargés`;
  document.getElementById('id-sc').classList.add('done');
  document.getElementById('idt-sc').textContent=`✅ ${files.length} scans chargés`;
  document.getElementById('ids-sc').textContent='Reliés automatiquement aux courriers';
  document.getElementById('bi2').disabled=false;
}

function goStep(n){
  [1,2,3].forEach(i=>{
    document.getElementById('is'+i).style.display=i===n?'block':'none';
    const s=document.getElementById('si'+i);
    s.classList.remove('active','done');
    if(i<n)s.classList.add('done'); else if(i===n)s.classList.add('active');
  });
  if(n===3)buildSum();
}

function buildSum(){
  let linked=0;
  impRows.forEach(r=>{
    if(r.scan_name&&impScans[r.scan_name])linked++;
    else if(r.scan_name){
      const base=r.scan_name.split('.')[0].toLowerCase();
      const match=Object.keys(impScans).find(k=>k.toLowerCase().startsWith(base));
      if(match){r.scan_name=match;linked++;}
    }
  });
  document.getElementById('imp-sum').innerHTML=`
    <div class="imp-sum">
      <strong>Résumé de l'import</strong><br><br>
      📊 ${impRows.length} courriers lus dans le fichier Excel<br>
      🗂️ ${Object.keys(impScans).length} scans chargés<br>
      🔗 ${linked} scans reliés automatiquement<br><br>
      <span style="font-size:12px;color:var(--muted)">Ces données seront ajoutées à vos archives.</span>
    </div>`;
}

async function doImport(){
  const btn=document.querySelector('#is3 .btn-teal');
  btn.disabled=true;btn.textContent='⏳ Import en cours…';
  const res = await call('importData', JSON.stringify({rows:impRows,scans:impScans}));
  localStorage.setItem('sm_imp','1');
  toast(`✅ ${res.imported} courriers importés !`,'s');
  startApp();
}
function skipImport(){localStorage.setItem('sm_imp','1');startApp()}

// ── Export ──
async function doExport(type) {
  const res = await call('exportExcel', type);
  if(res.ok) toast('✅ Export Excel enregistré','s');
  else if(res.error!=='Annulé') toast('❌ '+res.error,'e');
}

// ── Nav ──
function showSc(id){hideSc();const e=document.getElementById(id);if(e)e.classList.add('on')}
function hideSc(){document.querySelectorAll('.fs').forEach(s=>s.classList.remove('on'))}
function v(id){return document.getElementById(id)?.value?.trim()||''}

// ── Toast ──
let tT;
function toast(msg,type=''){
  const e=document.getElementById('toast');
  e.textContent=msg;e.className='on'+(type?' '+type:'');
  clearTimeout(tT);tT=setTimeout(()=>{e.className=''},3400);
}
</script>
</body>
</html>"""

# ─────────────────────────────────────────
#  FENÊTRE PRINCIPALE
# ─────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Courrier Saint-Magne V9")
        self.setMinimumSize(1100, 720)
        self.resize(1280, 800)

        # WebEngine
        self.view = QWebEngineView()
        self.setCentralWidget(self.view)

        # WebChannel
        self.channel = QWebChannel()
        self.bridge  = Bridge(self)
        self.channel.registerObject("bridge", self.bridge)
        self.view.page().setWebChannel(self.channel)

        # Charger HTML
        self.view.setHtml(HTML, QUrl("qrc:///"))

if __name__ == "__main__":
    if hasattr(Qt.ApplicationAttribute, "AA_EnableHighDpiScaling"):
        QApplication.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
    if hasattr(Qt.ApplicationAttribute, "AA_UseHighDpiPixmaps"):
        QApplication.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)

    init_db()
    app = QApplication(sys.argv)
    app.setApplicationName("Courrier Saint-Magne")

    win = MainWindow()
    win.show()
    sys.exit(app.exec())
