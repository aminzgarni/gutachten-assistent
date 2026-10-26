#!/usr/bin/env python3
"""Lokale Browser-Oberfläche. Start: python programm/web.py  (nur auf diesem Rechner erreichbar, 127.0.0.1)."""
import json, os, sys, threading, subprocess, secrets, webbrowser, time, codecs, re, shutil, base64, importlib, importlib.util, platform
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse, parse_qs

if sys.stdout is None: sys.stdout = open(os.devnull, "w")
if sys.stderr is None: sys.stderr = open(os.devnull, "w")
TOOL = Path(__file__).resolve().parent
sys.path.insert(0, str(TOOL))
import gutachten as G  # nutzt dieselben Funktionen wie das Terminal-Menü

TOKEN = secrets.token_urlsafe(16)
LAST = [time.time()]
JOBS, LOCK = {}, threading.Lock()
KEYS = ["betroffene_person", "adresse", "richter", "untersuchungsort", "gutachten_auftrag", "auftrag_text", "honorar_text", "honorargruppe",
        "beschluss_datum", "untersuchungsdatum", "gutachten_datum", "rechnungsdatum", "abteilung", "auftrag_details", "auftrag_datum"]
SUBS = G.SUB
TEXT_EINST = ["gutachter_name", "gutachter_titel", "fachgebiet", "gutachten_titel", "strasse", "plz_ort", "ort", "kontoinhaber", "iban", "bic", "bank",
              "abteilung_standard", "whisper_modell", "ki_prompt_zusatz", "design"]
ZAHL_EINST = ["km_satz", "schreibgebuehr_pro_einheit", "anschlaege_pro_einheit"]
PAKETE = [  # id, Modul, pip-Name, Pflicht, Beschreibung
    ("docx", "python-docx", True, "Word-Dateien erzeugen (Gutachten, Rechnung)"),
    ("openpyxl", "openpyxl", True, "Excel-Übersicht der Rechnungen"),
    ("pymupdf", "pymupdf", True, "PDFs lesen und zusammenführen"),
    ("pymupdf4llm", "pymupdf4llm", False, "Bessere Textumwandlung (empfohlen)"),
    ("faster_whisper", "faster-whisper", False, "Gespräche lokal transkribieren"),
    ("av", "av", False, "Audio-Dateien lesen (für die Transkription)"),
]

def py_exe():
    exe = sys.executable
    if os.name == "nt" and exe.lower().endswith("pythonw.exe"):
        c = exe[:-len("pythonw.exe")] + "python.exe"
        if os.path.exists(c): return c
    return exe

def fall_pfad(fid):
    if not G.BASE: raise ValueError("Es ist noch kein Arbeitsordner festgelegt.")
    p = (G.BASE / fid).resolve()
    if G.BASE.resolve() not in p.parents or not (p / "fall.json").exists(): raise ValueError("Fall nicht gefunden")
    return p

def lies(p): return json.loads((p / "fall.json").read_text(encoding="utf-8"))
def schreibe(p, d): G.write_atomic(p / "fall.json", lambda t: t.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8"))

def dateien(p, sub):
    d = p / SUBS[sub]
    return [f.name for f in sorted(d.iterdir()) if f.is_file() and not f.name.startswith(("~", "."))] if d.exists() else []

def start_job(argv, cwd, stdin_text=""):
    with LOCK:
        if any(not j["done"] for j in JOBS.values()): raise RuntimeError("Es läuft bereits ein Vorgang.")
        jid = secrets.token_hex(6); job = {"text": "", "done": False, "code": None}; JOBS[jid] = job
    env = dict(os.environ, GUTACHTEN_WEB="1", PYTHONIOENCODING="utf-8", PYTHONUTF8="1", PYTHONUNBUFFERED="1")
    def run():
        try:
            proc = subprocess.Popen(argv, cwd=str(cwd), stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, env=env,
                                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            proc.stdin.write(stdin_text.encode("utf-8")); proc.stdin.close()
            dec = codecs.getincrementaldecoder("utf-8")("replace")
            while True:
                chunk = os.read(proc.stdout.fileno(), 4096)
                if not chunk: break
                job["text"] += dec.decode(chunk, False)
            job["code"] = proc.wait()
        except Exception as e:
            job["text"] += f"\nFehler: {e}"; job["code"] = 1
        job["done"] = True; importlib.invalidate_caches()
    threading.Thread(target=run, daemon=True).start()
    return jid

def gjob(cmd, cwd, stdin_text=""): return start_job([py_exe(), "-u", str(TOOL / "gutachten.py")] + cmd, cwd, stdin_text)

def kandidaten():
    h = Path.home(); c = []
    for pat in ("OneDrive*", "Google Drive*", "GoogleDrive*", "Dropbox*", "iCloud*"): c += sorted(h.glob(pat))
    cs = h / "Library" / "CloudStorage"
    if cs.exists(): c += sorted(cs.iterdir())
    if os.name == "nt":
        for l in "DEFGHIJKLMNOPQRSTUVWXYZ":
            for n in ("My Drive", "Meine Ablage"):
                if Path(f"{l}:/{n}").exists(): c.append(Path(f"{l}:/{n}"))
    for n in ("Documents", "Dokumente"):
        if (h / n).exists(): c.append(h / n)
    out = []
    for p in c:
        if not p.is_dir(): continue
        low = str(p).lower()
        et = "OneDrive" if "onedrive" in low else "Google Drive" if ("google" in low or "my drive" in low or "meine ablage" in low) else "Dropbox" if "dropbox" in low else "iCloud" if "icloud" in low else "Dieser Computer"
        out.append({"label": et + (" (wird synchronisiert)" if et != "Dieser Computer" else " (nicht synchronisiert)"), "pfad": str(p / "Gutachten")})
    return out

def einstellungen_sicht():
    c = G.CFG
    return {**{k: c.get(k, "") for k in TEXT_EINST + ZAHL_EINST}, "honorar_pro_stunde": c["honorar_pro_stunde"],
            "standard_positionen": c["standard_positionen"], "gutachten_module_IV": c["gutachten_module_IV"], "ki_tools": c.get("ki_tools", [])}

def saubere_tools(tools):
    """KI-Dienste aus den Einstellungen: nur Name und http(s)-Adresse zulassen."""
    out = []
    for t in tools:
        name, url = str(t.get("name", "")).strip(), str(t.get("url", "")).strip()
        if not name: continue
        if url and not re.match(r"^https?://", url, re.I): url = "https://" + url
        out.append({"name": name, "url": url, "aktiv": bool(t.get("aktiv"))})
    return out

def lies_expl(liste):
    return [{k: str(x.get(k, "")).strip() for k in ("art", "datum", "beteiligte", "ort", "notiz", "quelle")} for x in liste
            if any(str(x.get(k, "")).strip() for k in ("art", "datum", "beteiligte", "ort", "notiz", "quelle"))]

def vorlagen_status():
    import vorlagen_engine as VE
    out = {}
    for art in ("rechnung", "gutachten"):
        eigene = G.BASE / "_einstellungen" / "vorlagen" / f"{art.capitalize()}.docx" if G.BASE else None
        pf = G.vorlage_pfad(art); st = {"eigene": bool(eigene and eigene.exists())}
        try: st["pruefung"] = VE.pruefe(pf, art)
        except Exception as e: st["pruefung"] = None; st["hinweis"] = "Prüfung nicht möglich (python-docx fehlt?)"
        out[art] = st
    return out

def komponenten():
    pk = []
    for modul, pip, pflicht, text in PAKETE:
        pk.append({"modul": modul, "pip": pip, "pflicht": pflicht, "text": text, "vorhanden": importlib.util.find_spec(modul) is not None})
    return {"pakete": pk, "tesseract": G.tesseract_status(), "os": sys.platform, "python": platform.python_version(),
            "winget": bool(shutil.which("winget")), "brew": bool(shutil.which("brew") or os.path.exists("/opt/homebrew/bin/brew")),
            "modell": G.CFG.get("whisper_modell")}

class H(BaseHTTPRequestHandler):
    def log_message(self, *a): pass
    def _ok(self):
        host = (self.headers.get("Host") or "").split(":")[0]
        if host not in ("127.0.0.1", "localhost"): return False
        q = parse_qs(urlparse(self.path).query)
        ok = self.headers.get("X-Token") == TOKEN or q.get("t", [""])[0] == TOKEN
        if ok: LAST[0] = time.time()
        return ok
    def _send(self, code, body, ctype="application/json"):
        b = body if isinstance(body, bytes) else json.dumps(body, ensure_ascii=False).encode("utf-8")
        self.send_response(code); self.send_header("Content-Type", ctype + "; charset=utf-8")
        self.send_header("Content-Length", str(len(b))); self.send_header("Cache-Control", "no-store"); self.end_headers(); self.wfile.write(b)
    def do_GET(self):
        if not self._ok(): return self._send(403, {"fehler": "Zugriff verweigert"})
        u = urlparse(self.path); q = {k: v[0] for k, v in parse_qs(u.query).items()}
        try:
            if u.path == "/": return self._send(200, (TOOL / "web" / "index.html").read_bytes(), "text/html")
            if u.path == "/api/ping": return self._send(200, {"ok": True})
            if u.path == "/api/start":
                return self._send(200, {"konfiguriert": bool(G.BASE), "arbeitsordner": str(G.BASE or ""), "kandidaten": kandidaten(), "os": sys.platform})
            if u.path == "/api/einstellungen":
                import vorlagen_engine as VE
                return self._send(200, {"einstellungen": einstellungen_sicht(), "vorlagen": vorlagen_status() if G.BASE else {}, "beschreibung": VE.BESCHREIBUNG,
                                        "arbeitsordner": str(G.BASE or "")})
            if u.path == "/api/komponenten": return self._send(200, komponenten())
            if u.path == "/api/faelle":
                try: G.sync_rechnungsliste()
                except Exception: pass
                out = []
                for p in G.faelle():
                    f = lies(p); out.append({"id": f"{p.parent.name}/{p.name}", "az": f["aktenzeichen"], "gericht": f["gericht"], "gruppe": f["honorargruppe"]})
                return self._send(200, out)
            if u.path == "/api/fall":
                p = fall_pfad(q["id"]); G.migriere(p); f = lies(p)
                pf = G.pseudo_file(p); eintraege = json.loads(pf.read_text(encoding="utf-8")) if pf.exists() else []
                return self._send(200, {"fall": f, "pseudonyme": eintraege, "module": G.CFG["gutachten_module_IV"],
                                        "dateien": {k: dateien(p, k) for k in SUBS}, "ocr": G.tesseract_status(), "ordner": str(p),
                                        "gutachter": G.CFG.get("gutachter_name", ""),
                                        "ki_dateien": [x.name for x in G.ki_dateien(p)], "gespraech_quellen": [x.name for x in G.gespraech_quellen(p)],
                                        "expl_arten": G.EXPL_ARTEN, "arbeitsweisen": {k: v[0] for k, v in G.ARBEITSWEISEN.items()}})
            if u.path == "/api/job":
                j = JOBS.get(q["id"]); pos = int(q.get("pos", 0))
                if not j: return self._send(404, {"fehler": "unbekannt"})
                return self._send(200, {"text": j["text"][pos:], "pos": len(j["text"]), "done": j["done"], "code": j["code"]})
            return self._send(404, {"fehler": "nicht gefunden"})
        except Exception as e:
            return self._send(400, {"fehler": str(e)})
    def do_POST(self):
        if not self._ok(): return self._send(403, {"fehler": "Zugriff verweigert"})
        n = int(self.headers.get("Content-Length") or 0); path = urlparse(self.path).path
        if n > (20_000_000 if path == "/api/vorlage" else 2_000_000): return self._send(413, {"fehler": "zu groß"})
        b = json.loads(self.rfile.read(n) or b"{}")
        try:
            return self._post(path, b)
        except Exception as e:
            return self._send(400, {"fehler": str(e)})
    def _post(self, path, b):
        if path == "/api/quit":
            self._send(200, {"ok": True}); H.beenden(); return
        if path == "/api/einstellungen":
            ao = (b.get("arbeitsordner") or "").strip()
            if ao:
                pa = Path(ao).expanduser(); pa.mkdir(parents=True, exist_ok=True)
                G.HOME.mkdir(parents=True, exist_ok=True)
                (G.HOME / "start.json").write_text(json.dumps({"arbeitsordner": str(pa)}, ensure_ascii=False), encoding="utf-8")
                (pa / "_einstellungen" / "vorlagen").mkdir(parents=True, exist_ok=True); G.neu_laden()
            if not G.BASE: raise ValueError("Bitte zuerst einen Arbeitsordner festlegen.")
            f = b.get("felder", {}); datei = G.einstellungen_datei(); datei.parent.mkdir(parents=True, exist_ok=True)
            cur = json.loads(datei.read_text(encoding="utf-8")) if datei.exists() else {}
            for k in TEXT_EINST:
                if k in f: cur[k] = str(f[k]).strip()
            if cur.get("design") not in (None, "light", "dark", "system"): cur["design"] = "system"
            for k in ZAHL_EINST:
                if k in f: cur[k] = float(str(f[k]).replace(",", ".")) if k != "anschlaege_pro_einheit" else int(float(str(f[k]).replace(",", ".")))
            if "honorar_pro_stunde" in f: cur["honorar_pro_stunde"] = {g: float(str(v).replace(",", ".")) for g, v in f["honorar_pro_stunde"].items() if g in ("M2", "M3")}
            if "ki_tools" in f: cur["ki_tools"] = saubere_tools(f["ki_tools"])
            if "gutachten_module_IV" in f: cur["gutachten_module_IV"] = [str(m).strip() for m in f["gutachten_module_IV"] if str(m).strip()]
            if "standard_positionen" in f: cur["standard_positionen"] = [[str(a).strip(), str(c).strip()] for a, c in f["standard_positionen"] if str(a).strip()]
            G.write_atomic(datei, lambda t: t.write_text(json.dumps(cur, ensure_ascii=False, indent=2), encoding="utf-8")); G.neu_laden()
            return self._send(200, {"ok": True, "arbeitsordner": str(G.BASE)})
        if path == "/api/neu":
            if not G.BASE: raise ValueError("Bitte zuerst die Einrichtung abschließen.")
            zeilen = [b["az"], b["gericht"], b["beschluss"], b["untersuchung"], b["gutachten_datum"], b.get("gruppe", "M3")]
            if not all(z.strip() for z in zeilen[:5]): raise ValueError("Bitte alle Felder ausfüllen (auch das Datum der Gutachtenerstattung).")
            return self._send(200, {"job": gjob(["neu"], G.BASE, "\n".join(z.strip() for z in zeilen) + "\n")})
        if path == "/api/setup":
            a = b.get("aktion")
            if a == "pip":
                erlaubt = {pip for _, pip, _, _ in PAKETE}; pak = [x for x in b.get("pakete", []) if x in erlaubt]
                if not pak: raise ValueError("Nichts zu installieren.")
                return self._send(200, {"job": start_job([py_exe(), "-m", "pip", "install", "--upgrade"] + pak, Path.home())})
            if a == "tessdata": return self._send(200, {"job": gjob(["tessdata"], Path.home())})
            if a == "modell": return self._send(200, {"job": gjob(["modell"], Path.home())})
            if a == "tesseract":
                if sys.platform.startswith("win") and shutil.which("winget"):
                    return self._send(200, {"job": start_job(["winget", "install", "-e", "--id", "UB-Mannheim.TesseractOCR", "--silent",
                                                              "--accept-package-agreements", "--accept-source-agreements"], Path.home())})
                brew = shutil.which("brew") or ("/opt/homebrew/bin/brew" if os.path.exists("/opt/homebrew/bin/brew") else None)
                if sys.platform == "darwin" and brew: return self._send(200, {"job": start_job([brew, "install", "tesseract"], Path.home())})
                raise ValueError("Automatische Installation hier nicht möglich. Bitte die Anleitung in den Einstellungen (Programme) befolgen.")
            raise ValueError("unbekannte Aktion")
        if path == "/api/vorlage":
            art = b["art"]
            if art not in ("rechnung", "gutachten"): raise ValueError("unbekannte Vorlage")
            if not G.BASE: raise ValueError("Bitte zuerst einen Arbeitsordner festlegen.")
            data = base64.b64decode(b["daten"])
            import zipfile, io
            try:
                if "word/document.xml" not in zipfile.ZipFile(io.BytesIO(data)).namelist(): raise ValueError
            except Exception: raise ValueError("Das ist keine gültige Word-Datei (.docx).")
            ziel = G.BASE / "_einstellungen" / "vorlagen" / f"{art.capitalize()}.docx"; ziel.parent.mkdir(parents=True, exist_ok=True)
            ziel.write_bytes(data)
            import vorlagen_engine as VE
            try: pr = VE.pruefe(ziel, art)
            except Exception as e: pr = {"gefunden": [], "unbekannt": [], "fehlt": [], "hinweis": "Prüfung nicht möglich: " + str(e)}
            return self._send(200, {"ok": True, "pruefung": pr})
        if path == "/api/vorlage_reset":
            f = G.BASE / "_einstellungen" / "vorlagen" / f"{b['art'].capitalize()}.docx"
            if f.exists(): f.unlink()
            return self._send(200, {"ok": True})
        if path == "/api/vorlage_oeffnen":
            art = b["art"]
            if art not in ("rechnung", "gutachten"): raise ValueError("unbekannte Vorlage")
            eigene = G.BASE / "_einstellungen" / "vorlagen" / f"{art.capitalize()}.docx"
            if not eigene.exists():
                eigene.parent.mkdir(parents=True, exist_ok=True); shutil.copyfile(G.TOOL / "vorlagen" / f"{art.capitalize()}_Standard.docx", eigene)
            G.oeffne(eigene); return self._send(200, {"ok": True})
        p = fall_pfad(b["id"])
        if path == "/api/pseudonyme":
            eintr = [{"name": e["name"].strip(), "platzhalter": e["platzhalter"].strip(), "art": "person" if e.get("art") == "person" else "sonst"}
                     for e in b["eintraege"] if e.get("name", "").strip() and e.get("platzhalter", "").strip()]
            (p / "_intern").mkdir(exist_ok=True)
            G.write_atomic(G.pseudo_file(p), lambda t: t.write_text(json.dumps(eintr, ensure_ascii=False, indent=2), encoding="utf-8"))
            return self._send(200, {"ok": True, "anzahl": len(eintr)})
        if path == "/api/fall":
            f = lies(p)
            for k in KEYS:
                if k in b.get("felder", {}): f[k] = str(b["felder"][k]).strip()
            if "positionen" in b:
                f["positionen"] = [{"titel": str(x["titel"]).strip(), "text": str(x.get("text", "")).strip(), "stunden": float(str(x.get("stunden") or 0).replace(",", "."))}
                                   for x in b["positionen"] if str(x.get("titel", "")).strip()]
            if "explorationen" in b: f["explorationen"] = lies_expl(b["explorationen"])
            if "km" in b: f["fahrt"]["km_einfach"] = float(str(b["km"]).replace(",", ".") or 0)
            if "ziel" in b: f["fahrt"]["ziel"] = str(b["ziel"]).strip()
            if "anschlaege" in b: f["anschlaege"] = int(re.sub(r"\D", "", str(b["anschlaege"])) or 0)
            schreibe(p, f); return self._send(200, {"ok": True})
        if path == "/api/run":
            cmd = b["cmd"]
            if cmd not in ("akte", "transkript", "pseudo", "entwurf", "rechnung"): raise ValueError("unbekannter Befehl")
            args = [cmd]; stdin = ""
            if cmd in ("akte", "pseudo"):
                pf = G.pseudo_file(p); stdin = ("n\n" if pf.exists() and json.loads(pf.read_text(encoding="utf-8")) else "\n")
            if cmd == "rechnung": stdin = "\n" * (len(lies(p)["positionen"]) + 5)
            if cmd == "entwurf":
                datei = b.get("datei")
                if datei:
                    q = (p / SUBS["gutachten"] / datei).resolve()
                    if p.resolve() not in q.parents or not q.exists(): raise ValueError("Datei nicht gefunden")
                    args.append(str(q))
                if not (datei or "").lower().endswith(".docx"):
                    f = lies(p)
                    for k in ("betroffene_person", "adresse", "richter"):
                        if not f.get(k): raise ValueError("Bitte Betroffene Person, Adresse und Richter/in ausfüllen.")
                    if not f.get("untersuchungsort"): f["untersuchungsort"] = f["adresse"]
                    if not f.get("gutachten_auftrag"):
                        f["gutachten_auftrag"] = G.CFG.get("gutachten_auftrag_vorlage", "").format(gericht=f["gericht"], beschluss_datum=f.get("beschluss_datum", ""))
                    schreibe(p, f)
                if not datei:
                    sel = b.get("module", []); stdin = (",".join(str(int(i)) for i in sel) if sel else "0") + "\n"
            return self._send(200, {"job": gjob(args, p, stdin)})
        if path == "/api/rechnungsliste":
            f = lies(p); G.sync_rechnungsliste(); x = G.rechnungsliste_pfad(G.rechnungsjahr(f, (sorted(f.get("rechnungen", {})) or [None])[-1]))
            if b.get("oeffnen"):
                if not x.exists(): raise ValueError("Für dieses Jahr gibt es noch keine Rechnungsliste.")
                G.oeffne(x)
            return self._send(200, {"ok": True, "datei": x.name})
        if path == "/api/ki_prompt":
            f = lies(p); mod = [int(i) for i in b.get("module", []) if str(i).isdigit()] or None
            r = G.ki_prompt(p, f, module_idx=mod, arbeitsweise=b.get("arbeitsweise", "schrittweise"),
                            dateien=b.get("dateien") if isinstance(b.get("dateien"), list) else None)
            return self._send(200, r)
        if path in ("/api/ki_dateien_kopieren", "/api/ki_dateien_zeigen"):
            ok = {x.name: x for x in G.ki_dateien(p)}   # nur pseudonymisierte Dateien, nie Originale
            pf = [ok[n] for n in b.get("dateien", []) if n in ok]
            if not pf: raise ValueError("Keine pseudonymisierten Dateien ausgewählt oder vorhanden.")
            if path == "/api/ki_dateien_kopieren": return self._send(200, {"ok": G.kopiere_dateien_in_zwischenablage(pf), "anzahl": len(pf)})
            G.zeige_dateien(p / SUBS["ki"], pf); return self._send(200, {"ok": True})
        if path == "/api/ki_text_kopieren":
            return self._send(200, {"ok": G.kopiere_text(str(b.get("text", "")))})
        if path == "/api/ki_antwort":
            text = str(b.get("text", "")).strip()
            if len(text) < 50: raise ValueError("Der eingefügte Text ist zu kurz für einen Entwurf.")
            d = p / SUBS["gutachten"]; d.mkdir(exist_ok=True); n = 1
            while (d / f"KI-Entwurf_v{n}.md").exists(): n += 1
            name = f"KI-Entwurf_v{n}.md"; G.write_atomic(d / name, lambda t: t.write_text(text + "\n", encoding="utf-8"))
            return self._send(200, {"ok": True, "datei": name})
        if path == "/api/open":
            sub = b.get("sub"); name = b.get("name")
            ziel = p if not sub else p / SUBS[sub]
            if name:
                ziel = (ziel / name).resolve()
                if p.resolve() not in ziel.parents or not ziel.exists(): raise ValueError("Datei nicht gefunden")
            G.oeffne(ziel); return self._send(200, {"ok": True})
        return self._send(404, {"fehler": "nicht gefunden"})
    def do_PUT(self): self._send(405, {})

def main():
    for port in range(8765, 8800):
        try: srv = ThreadingHTTPServer(("127.0.0.1", port), H); break
        except OSError: continue
    else: sys.exit("Kein freier Port.")
    url = f"http://127.0.0.1:{port}/?t={TOKEN}"
    def beenden(): time.sleep(0.3); srv.shutdown()
    H.beenden = staticmethod(lambda: threading.Thread(target=beenden, daemon=True).start())
    def wache():  # beendet sich selbst, wenn der Browser lange nichts mehr abgefragt hat und kein Vorgang läuft
        while True:
            time.sleep(30)
            if time.time() - LAST[0] > 300 and not any(not j["done"] for j in JOBS.values()): srv.shutdown(); return
    threading.Thread(target=wache, daemon=True).start()
    print(f"Gutachten-Assistent läuft: {url}\nZum Beenden dieses Fenster schließen oder im Browser 'Beenden' klicken.")
    threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try: srv.serve_forever()
    except KeyboardInterrupt: pass

if __name__ == "__main__":
    main()
