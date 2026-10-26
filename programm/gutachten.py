#!/usr/bin/env python3
"""Gutachten-Assistent: Kernfunktionen (Akte, Transkription, Entwurf, Rechnung). Wird von web.py (Browser-Oberfläche) genutzt;
Einzelbefehle auch im Terminal: python gutachten.py [menu|neu|akte|transkript|pseudo|entwurf|rechnung]"""
import json, math, re, sys, os, time, datetime as dt
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

TOOL = Path(__file__).resolve().parent
HOME = Path(os.environ.get("GUTACHTEN_HOME_DIR") or (Path.home() / ".gutachten-assistent"))
DEFAULTS = json.loads((TOOL / "config.json").read_text(encoding="utf-8"))
BASE = None
CFG = dict(DEFAULTS)

def arbeitsordner():
    env = os.environ.get("GUTACHTEN_ARBEITSORDNER")
    if env: return Path(env)
    try:
        p = json.loads((HOME / "start.json").read_text(encoding="utf-8")).get("arbeitsordner")
        return Path(p) if p else None
    except Exception:
        return None

def einstellungen_datei(): return BASE / "_einstellungen" / "einstellungen.json" if BASE else None

def neu_laden():
    """Liest Arbeitsordner und Einstellungen neu ein (auch nach Änderungen im Browser)."""
    global BASE, CFG
    BASE = arbeitsordner(); c = dict(DEFAULTS)
    f = einstellungen_datei()
    if f and f.exists():
        try: c.update(json.loads(f.read_text(encoding="utf-8")))
        except Exception: pass
    c["gutachter_adresse"] = ", ".join(x for x in (c.get("strasse", ""), c.get("plz_ort", "")) if x)
    if not c.get("ort"): c["ort"] = re.sub(r"^\d{4,5}\s*", "", c.get("plz_ort", "")).strip()
    CFG = c
neu_laden()
D = Decimal

def eur(x):  return f"{D(x).quantize(D('0.01'), ROUND_HALF_UP):.2f}".replace(".", ",") + " EUR"
def num(x): return f"{D(str(x)).normalize():f}".replace(".", ",")
def std(h):  s = f"{D(str(h)).normalize():f}".replace(".", ","); return f"{s} Std."
WEB = bool(os.environ.get("GUTACHTEN_WEB"))
def ask(q, default=""):
    a = input("" if WEB else f"{q}" + (f" [{default}]" if default else "") + ": ").strip()
    return a or default
def safe(s): return re.sub(r"[^\w\-]+", "-", s).strip("-")
def write_atomic(path, saver):
    tmp = path.with_name("~tmp_" + path.name); saver(tmp)
    for _ in range(10):
        try: os.replace(tmp, path); return
        except PermissionError: time.sleep(0.5)
    import shutil
    shutil.copyfile(tmp, path)
    try: os.remove(tmp)
    except OSError: pass

def brauche_ordner():
    if not BASE: sys.exit("Es ist noch kein Arbeitsordner festgelegt. Bitte zuerst die Einrichtung (Einstellungen) im Assistenten abschließen.")

def brauche(modul, pip):
    try: __import__(modul)
    except ImportError: sys.exit(f"Das Programm-Paket '{pip}' fehlt. Bitte im Assistenten unter Einstellungen > Programme installieren.")
def next_version(folder, stem):
    n = 1
    while (folder / f"{stem}_v{n}.docx").exists(): n += 1
    return folder / f"{stem}_v{n}.docx"

# ---------- Ordnerstruktur ----------
# 01_Akte        Gerichts-PDFs (Eingang)
# 02_Gespraeche  optional: Aufnahmen (nur mit Einwilligung), Transkripte (.md, Korrektur lesen), eigene Notizen/E-Mails als .txt/.md. Enthält Klarnamen: NIE hochladen.
# 03_Gutachten   Entwürfe und Gutachten      04_Rechnung   Rechnungen
# 05_KI-Upload   NUR hier liegt, was zur KI hochgeladen wird: Akte_pseudonymisiert*.md und Transkript_*_pseudonymisiert.md      _intern   Originaltext, Pseudonym-Liste (NIE hochladen)
SUB = {"akte": "01_Akte", "gespraeche": "02_Gespraeche", "gutachten": "03_Gutachten", "rechnung": "04_Rechnung", "ki": "05_KI-Upload"}
AUDIO_ENDUNGEN = {".m4a", ".mp3", ".wav", ".mp4", ".aac", ".flac", ".ogg", ".opus", ".wma", ".caf", ".amr"}
TEXT_ENDUNGEN = {".md", ".txt"}

def _verschiebe_inhalt(quelle, ziel):
    ziel.mkdir(exist_ok=True)
    for f in sorted(quelle.iterdir()):
        z = ziel / f.name; n = 1
        while z.exists():
            z = ziel / f"{f.stem}_alt{n}{f.suffix}"; n += 1
        f.rename(z)
    try: quelle.rmdir()
    except OSError: pass

def migriere(ordner):
    """Bringt Fälle im alten Aufbau (02_Audio, 03_Transkript, 04_Gutachten, 05_Rechnung, 06_KI-Upload) auf die neue Struktur.
    Läuft still, ist wiederholbar und überschreibt nie etwas."""
    try:
        for alt, neu in (("02_Audio", "02_Gespraeche"), ("04_Gutachten", "03_Gutachten"), ("05_Rechnung", "04_Rechnung")):
            a, n = ordner / alt, ordner / neu
            if a.is_dir():
                if n.exists(): _verschiebe_inhalt(a, n)
                else: a.rename(n)
        a = ordner / "03_Transkript"
        if a.is_dir(): _verschiebe_inhalt(a, ordner / "02_Gespraeche")
        a, n = ordner / "06_KI-Upload", ordner / "05_KI-Upload"
        if a.is_dir():
            if n.exists(): _verschiebe_inhalt(a, n)
            else: a.rename(n)
        # pseudonymisierte Dateien, die eine Zwischenversion in 01_Akte abgelegt hat, zurück in den KI-Upload-Ordner
        if (ordner / "01_Akte").is_dir():
            (ordner / "05_KI-Upload").mkdir(exist_ok=True)
            for f in list((ordner / "01_Akte").iterdir()):
                if f.is_file() and "_pseudonymisiert" in f.name and f.suffix.lower() in (".md", ".txt"):
                    z = ordner / "05_KI-Upload" / f.name
                    if z.exists(): z.unlink()
                    f.rename(z)
    except OSError as e:
        print(f"Hinweis: Alte Ordner konnten nicht umbenannt werden ({e}). Bitte Dateien im Explorer/Finder schließen und erneut öffnen.")
    for s in SUB.values(): (ordner / s).mkdir(exist_ok=True)
    (ordner / "_intern").mkdir(exist_ok=True)

# ---------- neu ----------
def cmd_neu():
    brauche_ordner()
    az = ask("Aktenzeichen (z. B. XVII 123/45)")
    gericht = ask("Amtsgericht (z. B. Dachau)")
    heute = dt.date.today()
    fall = {
        "aktenzeichen": az, "gericht": gericht,
        "abteilung": CFG.get("abteilung_standard", "Abteilung für Betreuungssachen"),
        "beschluss_datum": ask("Beschluss vom (TT.MM.JJJJ)"),
        "untersuchungsdatum": ask("Untersuchungsdatum (TT.MM.JJJJ)"),
        "gutachten_datum": ask("Datum der Gutachtenerstattung (TT.MM.JJJJ)", heute.strftime("%d.%m.%Y")),
        "honorargruppe": ask("Honorargruppe (M2 oder M3)", "M3").upper().replace(" ", ""),
        "auftrag_text": CFG.get("auftrag_text_vorlage", "Gutachten gemäß Beschluss vom {beschluss_datum}"),
        "honorar_text": CFG.get("honorar_text_vorlage", ""),
        "gutachten_auftrag": CFG.get("gutachten_auftrag_vorlage", "").format(gericht=gericht, beschluss_datum=""),
        "auftrag_details": "", "auftrag_datum": "", "explorationen": [],
        "fahrt": {"ziel": "", "km_einfach": 0},
        "anschlaege": 0,
        "positionen": [{"titel": t, "text": x, "stunden": 0} for t, x in CFG["standard_positionen"]],
        "rechnungsdatum": "",
    }
    fall["explorationen"] = [{"art": "Persönliche Exploration der betroffenen Person", "datum": fall["untersuchungsdatum"], "beteiligte": "",
                              "ort": "", "notiz": "", "quelle": ""}]
    jahr = fall["beschluss_datum"][-4:] if re.fullmatch(r"\d\d\.\d\d\.\d{4}", fall["beschluss_datum"]) else str(heute.year)
    ordner = BASE / jahr / f"{safe(gericht)}_{safe(az)}"
    ordner.mkdir(parents=True, exist_ok=True); migriere(ordner)
    write_atomic(ordner / "fall.json", lambda p: p.write_text(json.dumps(fall, ensure_ascii=False, indent=2), encoding="utf-8"))
    print(f"Fall angelegt: {ordner}" + ("" if WEB else "\nNächster Schritt: im Menü den Fall wählen."))

# ---------- rechnung ----------
def find_fall():
    p = Path.cwd()
    for c in [p, *p.parents]:
        if (c / "fall.json").exists(): migriere(c); return c
    sys.exit("fall.json nicht gefunden. Bitte im Fallordner ausführen.")

def cmd_rechnung():
    ordner = find_fall(); fp = ordner / "fall.json"
    fall = json.loads(fp.read_text(encoding="utf-8"))
    print("Stunden je Position eingeben (Enter = beibehalten, 0 = Position weglassen):")
    for pos in fall["positionen"]:
        v = ask(f"  {pos['titel']}", str(pos["stunden"]))
        pos["stunden"] = float(v.replace(",", "."))
    k = ask("Einfache Strecke in km (0 = keine Fahrtkosten)", str(fall["fahrt"]["km_einfach"]))
    fall["fahrt"]["km_einfach"] = float(k.replace(",", "."))
    if fall["fahrt"]["km_einfach"]:
        fall["fahrt"]["ziel"] = ask("Fahrtziel (Ort, Adresse)", fall["fahrt"]["ziel"])
    a = ask("Anschläge inkl. Leerzeichen (0 = keine Schreibgebühr)", str(fall["anschlaege"]))
    fall["anschlaege"] = int(a.replace(".", ""))
    write_atomic(fp, lambda p: p.write_text(json.dumps(fall, ensure_ascii=False, indent=2), encoding="utf-8"))

    gruppe = fall["honorargruppe"]; satz = D(str(CFG["honorar_pro_stunde"][gruppe]))
    pos = [p for p in fall["positionen"] if p["stunden"] > 0]
    stunden = sum(D(str(p["stunden"])) for p in pos)
    honorar = stunden * satz
    km = D(str(fall["fahrt"]["km_einfach"])) * 2
    fahrt = km * D(str(CFG["km_satz"]))
    einheiten = math.ceil(fall["anschlaege"] / CFG["anschlaege_pro_einheit"]) if fall["anschlaege"] else 0
    schreib = einheiten * D(str(CFG["schreibgebuehr_pro_einheit"]))
    gesamt = honorar + fahrt + schreib

    datei = build_docx(fall, pos, gruppe, satz, stunden, honorar, km, fahrt, einheiten, schreib, gesamt, ordner)
    rdatum = fall.get("rechnungsdatum") or fall.get("gutachten_datum", "")
    fall.setdefault("rechnungen", {})[datei.name] = {"datum": rdatum, "gruppe": gruppe, "stunden": float(stunden), "gesamt": float(gesamt)}
    write_atomic(fp, lambda p: p.write_text(json.dumps(fall, ensure_ascii=False, indent=2), encoding="utf-8"))
    jahre = sync_rechnungsliste()
    print(f"\nRechnung: {datei.name}   Gesamt: {eur(gesamt)}\nListe für den Steuerberater: Rechnungen_{rechnungsjahr(fall, datei.name)}.xlsx im Arbeitsordner" + ("" if jahre else " (Datei war nicht änderbar, bitte schließen und unter Rechnung 'Liste aktualisieren' klicken)"))

def vorlage_pfad(art):
    eigene = BASE / "_einstellungen" / "vorlagen" / f"{art.capitalize()}.docx" if BASE else None
    return eigene if eigene and eigene.exists() else TOOL / "vorlagen" / f"{art.capitalize()}_Standard.docx"

def basis_werte(fall, art):
    c = CFG; gr = fall.get("honorargruppe", "M3"); ctx = {k: fall.get(k, "") for k in ("beschluss_datum", "untersuchungsdatum", "gericht")}
    if art == "rechnung": auftrag = fall.get("auftrag_text", "").format(**ctx)
    else:
        auftrag = fall.get("gutachten_auftrag", "")
        if fall.get("richter") and "Richter" not in auftrag: auftrag += f" (Richter am Amtsgericht {fall['richter']})"
    return {"GUTACHTER_NAME": c.get("gutachter_name", ""), "GUTACHTER_TITEL": c.get("gutachter_titel", ""),
            "GUTACHTER_STRASSE": c.get("strasse", ""), "GUTACHTER_PLZ_ORT": c.get("plz_ort", ""), "GUTACHTER_ADRESSE": c.get("gutachter_adresse", ""),
            "FACHGEBIET": c.get("fachgebiet", ""), "GUTACHTEN_TITEL": c.get("gutachten_titel", "Gutachten"), "ORT": c.get("ort", ""),
            "KONTOINHABER": c.get("kontoinhaber", ""), "IBAN": c.get("iban", ""), "BIC": c.get("bic", ""), "BANK": c.get("bank", ""),
            "GERICHT": fall["gericht"], "ABTEILUNG": fall.get("abteilung", ""), "AKTENZEICHEN": fall["aktenzeichen"],
            "BESCHLUSS_DATUM": fall.get("beschluss_datum", ""), "UNTERSUCHUNGSDATUM": fall.get("untersuchungsdatum", ""),
            "GUTACHTEN_DATUM": fall.get("gutachten_datum", ""), "RECHNUNGSDATUM": fall.get("rechnungsdatum") or fall.get("gutachten_datum", ""),
            "AUFTRAG": auftrag, "HONORAR_TEXT": fall.get("honorar_text", ""), "HONORARGRUPPE": f"{gr[0]} {gr[1:]}",
            "BETROFFENE_PERSON": fall.get("betroffene_person", ""), "ADRESSE_BETROFFENE": fall.get("adresse", ""),
            "RICHTER": fall.get("richter", ""), "UNTERSUCHUNGSORT": fall.get("untersuchungsort", "")}

def build_docx(fall, pos, gruppe, satz, stunden, honorar, km, fahrt, einheiten, schreib, gesamt, ordner):
    brauche("docx", "python-docx")
    from docx import Document
    import vorlagen_engine as VE
    d = Document(str(vorlage_pfad("rechnung")))
    VE.bloecke(d, {"FAHRT": bool(km), "SCHREIB": bool(einheiten)})
    VE.wiederhole_zeilen(d, [{"POS_TITEL": p["titel"].format(**{k: fall.get(k, "") for k in ("untersuchungsdatum", "beschluss_datum")}),
                              "POS_TEXT": p["text"], "POS_STUNDEN": std(p["stunden"])} for p in pos])
    VE.nummeriere(d)
    w = basis_werte(fall, "rechnung")
    w.update({"HONORARSATZ": eur(satz), "STUNDEN_GESAMT": std(stunden), "STUNDEN_ZAHL": num(stunden), "HONORAR_BETRAG": eur(honorar),
              "FAHRT_ZIEL": fall["fahrt"].get("ziel", ""), "FAHRT_KM_EINFACH": num(fall["fahrt"]["km_einfach"]), "KM_GESAMT": num(km),
              "KM_SATZ": f"{D(str(CFG['km_satz'])):.2f}".replace(".", ","), "FAHRTKOSTEN": eur(fahrt),
              "SCHREIB_EINHEITEN": str(einheiten), "ANSCHLAEGE": f"{fall['anschlaege']:,}".replace(",", "."),
              "SCHREIB_SATZ": eur(CFG["schreibgebuehr_pro_einheit"]), "SCHREIBGEBUEHR": eur(schreib), "GESAMTBETRAG": eur(gesamt)})
    rest = VE.fuelle(d, w)
    if rest: print("Hinweis: Diese Platzhalter der Vorlage sind unbekannt und blieben stehen:", ", ".join("{{" + k + "}}" for k in rest))
    out = next_version(ordner / SUB["rechnung"], f"Rechnung_{safe(fall['aktenzeichen'])}")
    write_atomic(out, lambda p: d.save(p))
    return out

# ---------- Rechnungsliste (Excel je Jahr, für den Steuerberater) ----------
RE_KOPF = ["Aktenzeichen", "Gericht", "Rechnungsdatum", "Honorargruppe", "Stunden", "Gesamtbetrag EUR", "Status", "Datei"]

def _d(s):
    try: return dt.datetime.strptime(str(s).strip(), "%d.%m.%Y").date()
    except ValueError: return None

def _zelle_datum(v):
    if isinstance(v, (dt.datetime, dt.date)): return v.strftime("%d.%m.%Y")
    return str(v or "").strip()

def _zahl(v):
    try: return float(str(v).replace(",", "."))
    except ValueError: return 0.0

def rechnungsjahr(fall, datei=None):
    e = fall.get("rechnungen", {}).get(datei, {}) if datei else {}
    d = _d(e.get("datum") or fall.get("rechnungsdatum") or fall.get("gutachten_datum"))
    return d.year if d else dt.date.today().year

def rechnungsliste_pfad(jahr): return BASE / f"Rechnungen_{jahr}.xlsx"

def sync_rechnungsliste():
    """Gleicht die Excel-Listen Rechnungen_JJJJ.xlsx im Arbeitsordner mit den vorhandenen Rechnungs-Dateien ab:
    neue Rechnung -> Zeile (Liste für das Jahr wird bei Bedarf angelegt), gelöschte Rechnungs-Datei -> Zeile entfällt,
    geändertes Datum -> Zeile wandert ins richtige Jahr. Die Spalte Status und von dir selbst ergänzte Zeilen bleiben erhalten.
    Schreibt nur bei Änderungen. Gibt die Jahre zurück, deren Liste geändert wurde."""
    if not BASE or not BASE.exists(): return []
    try: from openpyxl import Workbook, load_workbook
    except ImportError: return []
    from openpyxl.styles import Font
    faelle_ = {}
    for o in faelle():
        try: f = json.loads((o / "fall.json").read_text(encoding="utf-8"))
        except Exception: continue
        faelle_[(f.get("aktenzeichen", ""), f.get("gericht", ""))] = (o, f)
    # vorhandene Listen einlesen
    books = {}
    for x in BASE.glob("Rechnungen_*.xlsx"):
        m = re.fullmatch(r"Rechnungen_(\d{4})\.xlsx", x.name)
        if not m: continue
        try: wb = load_workbook(x)
        except Exception: continue
        ws = wb.active
        rows = [[c.value for c in r[:8]] + [None] * (8 - len(r[:8])) for r in ws.iter_rows(min_row=2) if any(c.value not in (None, "") for c in r)]
        books[int(m.group(1))] = (x, wb, ws, rows)
    # Zeilen aus Listen der alten Programmversion übernehmen (Fall kennt die Rechnung noch nicht)
    geaendert = set()
    for jahr, (x, wb, ws, rows) in books.items():
        for r in rows:
            az, ge, datum, gr, std_, ges, status, datei = r
            key = (str(az or ""), str(ge or ""))
            if key in faelle_ and datei:
                o, f = faelle_[key]
                if (o / SUB["rechnung"] / str(datei)).exists() and str(datei) not in f.get("rechnungen", {}):
                    f.setdefault("rechnungen", {})[str(datei)] = {"datum": _zelle_datum(datum), "gruppe": str(gr or ""), "stunden": _zahl(std_), "gesamt": _zahl(ges)}
                    geaendert.add(key)
    for key in geaendert:
        o, f = faelle_[key]; write_atomic(o / "fall.json", lambda p, f=f: p.write_text(json.dumps(f, ensure_ascii=False, indent=2), encoding="utf-8"))
    # Sollzustand: alle Rechnungen, deren Datei existiert
    soll = {}
    for (az, ge), (o, f) in faelle_.items():
        for datei, e in f.get("rechnungen", {}).items():
            if not (o / SUB["rechnung"] / datei).exists(): continue
            soll.setdefault(rechnungsjahr(f, datei), {})[(az, ge, datei)] = e
    alt_status = {(str(r[0] or ""), str(r[1] or ""), str(r[7] or "")): r[6] for _, (_, _, _, rows) in books.items() for r in rows}
    geschrieben = []
    for jahr in sorted(set(soll) | set(books)):
        x = rechnungsliste_pfad(jahr)
        alt_rows = books[jahr][3] if jahr in books else []
        neu = []
        for r in alt_rows:  # eigene Zeilen behalten (Fall oder Datei unbekannt)
            if (str(r[0] or ""), str(r[1] or "")) not in faelle_ or not r[7]: neu.append(r)
        for (az, ge, datei), e in soll.get(jahr, {}).items():
            neu.append([az, ge, _d(e.get("datum")) or e.get("datum", ""), e.get("gruppe", ""), _zahl(e.get("stunden")), _zahl(e.get("gesamt")),
                        alt_status.get((az, ge, datei)) or "offen", datei])
        neu.sort(key=lambda r: ((_d(_zelle_datum(r[2])) or dt.date.max), str(r[0] or "")))
        vgl = lambda rows: [[_zelle_datum(r[2]) if i == 2 else (round(_zahl(v), 2) if i in (4, 5) else (v or "")) for i, v in enumerate(r)] for r in rows]
        if jahr in books and vgl(alt_rows) == vgl(neu): continue
        if not neu and jahr not in soll and jahr not in books: continue
        if jahr in books: wb, ws = books[jahr][1], books[jahr][2]; ws.delete_rows(2, max(ws.max_row, 2))
        else:
            wb = Workbook(); ws = wb.active; ws.title = f"Rechnungen {jahr}"; ws.append(RE_KOPF)
            for c in ws[1]: c.font = Font(bold=True)
            for col, w in zip("ABCDEFGH", (18, 18, 16, 15, 10, 18, 12, 38)): ws.column_dimensions[col].width = w
            ws.freeze_panes = "A2"
        for r in neu:
            ws.append(r); z = ws.max_row
            ws.cell(z, 3).number_format = "DD.MM.YYYY"; ws.cell(z, 5).number_format = "0.0#"; ws.cell(z, 6).number_format = '#,##0.00'
        try: write_atomic(x, lambda p, wb=wb: wb.save(p)); geschrieben.append(jahr)
        except Exception as e: print(f"Hinweis: {x.name} konnte nicht aktualisiert werden (in Excel geöffnet?): {e}")
    return geschrieben

# ---------- Pseudonymisierung ----------
AUTO = [  # (Muster, Ersatz) – automatisch, nicht umkehrbar
    (re.compile(r"[\w.+-]+@[\w-]+(\.[\w-]+)+"), "[E-MAIL]"),
    (re.compile(r"\bDE\d{2}(?:\s?\d{4}){4}\s?\d{2}\b"), "[IBAN]"),
    (re.compile(r"(?<!\d)(?:\+49|0049|0)[\s/()-]*\d(?:[\s/()-]*\d){6,12}(?!\d)"), "[TELEFON]"),
    (re.compile(r"(geb\.|geboren am|\*)\s*\d{1,2}\.\s?\d{1,2}\.\s?(\d{4})"), r"\1 [TT.MM.]\2"),
]
def pseudo_file(ordner): return ordner / "_intern" / "pseudonyme.json"

def lade_pseudonyme(ordner):
    p = pseudo_file(ordner)
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else []

def pseudonyme_pflegen(ordner):
    eintraege = lade_pseudonyme(ordner)
    if eintraege:
        print("Vorhandene Pseudonyme:", ", ".join(f"{e['name']} -> {e['platzhalter']}" for e in eintraege))
        if ask("Weitere Namen/Begriffe hinzufügen? (j/n)", "n").lower() != "j": return eintraege
    else:
        print("Pseudonymisierung: Namen eingeben, die nie zu KI-Diensten gelangen sollen.")
        print("Vollständiger Name ohne Titel (z. B. Anna Müller), auch Straßen/Orte/Arbeitgeber möglich. Leere Eingabe = fertig.")
    buchst = iter("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
    for _ in eintraege: next(buchst)
    while True:
        name = ask("Name/Begriff")
        if not name: break
        art = ask("Art: 1 Person, 2 Ort/Straße/Sonstiges", "1")
        vorschlag = (("Betroffene" if not eintraege else "Person") if art == "1" else "Ort") + "_" + next(buchst)
        eintraege.append({"name": name, "platzhalter": ask("Platzhalter", vorschlag), "art": "person" if art == "1" else "sonst"})
    write_atomic(pseudo_file(ordner), lambda p: p.write_text(json.dumps(eintraege, ensure_ascii=False, indent=2), encoding="utf-8"))
    return eintraege

def alle_varianten(eintraege):
    """Liste (Original, Platzhalter, teilweise, umkehrbar). Ganze Namen zuerst, Teilnamen danach."""
    voll, teil = [], []
    for e in eintraege:
        n, ph = e["name"].strip(), e["platzhalter"]
        voll.append((n, ph))
        t = n.split()
        if e.get("art") == "person" and len(t) >= 2:
            vor, nach = " ".join(t[:-1]), t[-1]
            voll += [(f"{nach}, {vor}", ph), (f"{nach} {vor}", ph)]
            if len(nach) >= 3: teil.append((nach, ph + "_N"))
            if len(vor) >= 3: teil.append((vor, ph + "_V"))
    zaehl = {}
    for o, _ in teil: zaehl[o.lower()] = zaehl.get(o.lower(), 0) + 1
    out = [(o, p, False, True) for o, p in sorted(voll, key=lambda x: -len(x[0]))]
    gesehen = set()
    for o, p in sorted(teil, key=lambda x: -len(x[0])):
        if zaehl[o.lower()] > 1:  # gleicher Name bei mehreren Personen: nicht eindeutig
            if o.lower() not in gesehen: out.append((o, "[NAME]", True, False)); gesehen.add(o.lower())
        else: out.append((o, p, True, True))
    return out

def pseudonymisieren(text, eintraege):
    zaehler = {}
    for orig, ph, _, _ in alle_varianten(eintraege):
        text, k = re.subn(r"(?<![\wäöüÄÖÜß])" + re.escape(orig) + r"(?![\wäöüÄÖÜß])", ph, text, flags=re.I)
        zaehler[ph] = zaehler.get(ph, 0) + k
    for muster, ers in AUTO:
        text, k = muster.subn(ers, text)
        if k: zaehler[ers if "\\" not in ers else "geb.-Datum"] = k
    return text, zaehler

def rueckwaerts(text, eintraege):
    text = text.replace("\\_", "_")
    paare = {}
    for e in eintraege:
        n, ph = e["name"].strip(), e["platzhalter"]; paare[ph] = n
        t = n.split()
        if e.get("art") == "person" and len(t) >= 2:
            if len(t[-1]) >= 3: paare.setdefault(ph + "_N", t[-1])
            if len(" ".join(t[:-1])) >= 3: paare.setdefault(ph + "_V", " ".join(t[:-1]))
    for ph in sorted(paare, key=len, reverse=True): text = text.replace(ph, paare[ph])
    return text

# ---------- akte ----------
def natsort(p): return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", p.name)]

def cmd_akte():
    brauche_ordner()
    brauche("pymupdf", "pymupdf")
    import pymupdf
    ordner = find_fall()
    pdfs = sorted([p for p in (ordner / SUB["akte"]).glob("*.pdf") if not p.name.startswith(("_", "~"))], key=natsort)
    if not pdfs: sys.exit("Keine PDFs in 01_Akte gefunden.")
    print("Reihenfolge der Zusammenführung:")
    for i, p in enumerate(pdfs, 1): print(f"  {i}. {p.name}")
    print("(Reihenfolge ändern: Dateien mit 01_, 02_ ... beginnen lassen)")
    gesamt_doc, seiten_start, fehler = pymupdf.open(), {}, []
    for p in pdfs:
        try:
            q = pymupdf.open(str(p))
            if q.needs_pass and not q.authenticate(""):
                fehler.append((p.name, "durch Passwort geschützt")); continue
            seiten_start[len(gesamt_doc)] = p.name
            gesamt_doc.insert_pdf(q); q.close()
        except Exception as e:
            fehler.append((p.name, str(e)))
    for n, f in fehler: print(f"FEHLER: {n} konnte nicht gelesen werden ({f}).")
    if len(gesamt_doc) == 0: sys.exit("Keine PDF konnte gelesen werden.")
    gesamt = ordner / "_intern" / "Akte_gesamt.pdf"
    write_atomic(gesamt, lambda t: gesamt_doc.save(str(t), garbage=3, deflate=True))
    print(f"{len(gesamt_doc)} Seiten zusammengeführt (lokal): _intern/Akte_gesamt.pdf")

    seiten = text_seiten(gesamt)
    ocr_set = set()
    leer = [i + 1 for i, t in enumerate(seiten) if len(t.strip()) < 30]
    if leer:
        print(f"{len(leer)} Seite(n) ohne Text (Scans) - Texterkennung (OCR) läuft lokal ...")
        erg = ocr_seiten(gesamt, leer)
        for n, t in erg.items():
            if len(t.strip()) >= 30: seiten[n - 1] = t; ocr_set.add(n)
        leer = [n for n in leer if n not in ocr_set]
        if ocr_set: print(f"{len(ocr_set)} Seite(n) per OCR erkannt. Sie sind im Text markiert; Fehler sind möglich, bitte stichprobenartig prüfen.")
    if leer:
        print(f"ACHTUNG: {len(leer)} Seite(n) ohne lesbaren Text: {leer[:20]}{' ...' if len(leer) > 20 else ''}")

    md = ""
    for i, t in enumerate(seiten):
        if i in seiten_start: md += f"\n\n=== DOKUMENT: {seiten_start[i]} ===\n\n"
        md += (f"\n[Seite {i + 1} - per OCR erkannt, Fehler möglich]\n" if (i + 1) in ocr_set else f"\n[Seite {i + 1}]\n") + f"{t}\n"
    original = ordner / "_intern" / "Akte_original.md"
    write_atomic(original, lambda t: t.write_text(md, encoding="utf-8"))

    eintraege = pseudonyme_pflegen(ordner)
    ps, zaehler = pseudonymisieren(md, eintraege)
    # Kandidaten für weitere Namen
    bekannt = " ".join(e["name"] for e in eintraege).lower()
    kand = {}
    for m in re.finditer(r"\b(?:(?:Herrn?|Frau|Fr\.)\s+(?:(?:Dr\.|Prof\.)\s+)*|(?:Dr\.|Prof\.)\s+)((?:[A-ZÄÖÜ][\wäöüß-]+)(?:\s[A-ZÄÖÜ][\wäöüß-]+)?)", ps):
        n = m.group(1)
        if "_" not in n and not n.startswith("[") and n.lower() not in bekannt: kand[n] = kand.get(n, 0) + 1
    # Teile schreiben
    maxz = CFG.get("ki_max_zeichen", 60000)
    teile, akt = [], ""
    for block in re.split(r"(?=\n\[Seite \d+)", ps):
        if len(akt) + len(block) > maxz and akt: teile.append(akt); akt = ""
        akt += block
    if akt: teile.append(akt)
    (ordner / SUB["ki"]).mkdir(exist_ok=True)
    for alt in (ordner / SUB["ki"]).glob("Akte_pseudonymisiert*.md"): alt.unlink()
    for i, t in enumerate(teile, 1):
        name = "Akte_pseudonymisiert.md" if len(teile) == 1 else f"Akte_pseudonymisiert_Teil{i}.md"
        write_atomic(ordner / SUB["ki"] / name, lambda p, t=t: p.write_text(t, encoding="utf-8"))
    print(f"\nFertig: {len(teile)} Datei(en) in 05_KI-Upload ({len(ps):,} Zeichen), erkennbar am Namen *_pseudonymisiert.md.".replace(",", "."))
    print("Ersetzungen:", ", ".join(f"{k}: {v}" for k, v in zaehler.items() if v) or "keine")
    if kand:
        print("Mögliche weitere Namen im Text (bitte prüfen und ggf. mit erneutem `akte` ergänzen):",
              ", ".join(f"{k} ({v}x)" for k, v in sorted(kand.items(), key=lambda x: -x[1])[:15]))
    print("Bitte vor dem Hochladen kurz überfliegen: seltene Diagnosen + Ort können trotzdem identifizieren.")

# ---------- OCR (Tesseract, lokal) ----------
OCR_HINWEIS = ("Für gescannte Seiten wird Tesseract mit deutschem Sprachpaket benötigt:\n"
               "  Windows: Installer von github.com/UB-Mannheim/tesseract/wiki laden, bei 'Additional language data' 'German' anhaken.\n"
               "  Mac:     im Terminal  brew install tesseract tesseract-lang\n"
               "Danach 'Akte vorbereiten' erneut ausführen.")

def tesseract_pfad():
    import shutil
    p = shutil.which("tesseract")
    if p: return p
    for c in [r"C:\Program Files\Tesseract-OCR\tesseract.exe", r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
              "/opt/homebrew/bin/tesseract", "/usr/local/bin/tesseract", "/usr/bin/tesseract"]:
        if os.path.exists(c): return c
    return None

def tessdata_dir(): return HOME / "tessdata"

def tesseract_status():
    import subprocess
    exe = tesseract_pfad(); eigene = (tessdata_dir() / "deu.traineddata").exists()
    if not exe: return {"tesseract": False, "deu": False, "eigene_daten": eigene}
    langs = set()
    try:
        r = subprocess.run([exe, "--list-langs"], capture_output=True, text=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        langs = set((r.stdout or r.stderr).split()[1:])
    except Exception:
        pass
    return {"tesseract": True, "deu": ("deu" in langs) or eigene, "exe": exe, "eigene_daten": eigene}

def ocr_seiten(pdf, nummern):
    """OCR für Seitennummern (1-basiert). Gibt {Seite: Text} zurück."""
    import subprocess, tempfile, pymupdf
    st = tesseract_status()
    if not st["tesseract"]: print(OCR_HINWEIS); return {}
    lang = os.environ.get("GUTACHTEN_OCR_LANG") or ("deu" if st["deu"] else None)
    if not lang:
        print("Tesseract ist installiert, aber das deutsche Sprachpaket fehlt.\n" + OCR_HINWEIS); return {}
    extra = ["--tessdata-dir", str(tessdata_dir())] if lang == "deu" and st["eigene_daten"] else []
    doc = pymupdf.open(str(pdf)); out = {}
    with tempfile.TemporaryDirectory() as tmp:
        for k, n in enumerate(nummern, 1):
            img = Path(tmp) / f"s{n}.png"
            doc[n - 1].get_pixmap(dpi=300).save(str(img))
            r = subprocess.run([st["exe"], str(img), "stdout", "-l", lang, "--psm", "3"] + extra, capture_output=True,
                               text=True, encoding="utf-8", errors="replace", creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            out[n] = r.stdout
            print(f"\r   OCR {k}/{len(nummern)}", end="", flush=True)
    print()
    return out

def cmd_tessdata():
    import urllib.request
    ziel = tessdata_dir(); ziel.mkdir(parents=True, exist_ok=True)
    url = "https://github.com/tesseract-ocr/tessdata_best/raw/main/deu.traineddata"
    print("Lade deutsches Sprachpaket für die Texterkennung (ca. 15 MB) ...")
    req = urllib.request.Request(url, headers={"User-Agent": "gutachten-assistent"})
    with urllib.request.urlopen(req, timeout=60) as r, open(ziel / "~deu.part", "wb") as f:
        gesamt = int(r.headers.get("Content-Length") or 0); got = 0
        while True:
            b = r.read(65536)
            if not b: break
            f.write(b); got += len(b)
            if gesamt: print(f"\r   {got * 100 // gesamt} %", end="", flush=True)
    os.replace(ziel / "~deu.part", ziel / "deu.traineddata")
    print("\nFertig: Deutsches Sprachpaket installiert.")

def cmd_modell():
    try: from faster_whisper import WhisperModel
    except ImportError: sys.exit("faster-whisper ist nicht installiert (Einstellungen > Programme).")
    print(f"Lade Sprachmodell '{CFG.get('whisper_modell')}' (einmalig, mehrere GB, bitte Geduld) ...")
    WhisperModel(CFG.get("whisper_modell", "large-v3-turbo"), device="cpu", compute_type="int8")
    print("Fertig: Sprachmodell bereit (danach läuft die Transkription offline).")

def text_seiten(pdf):
    import pymupdf
    doc = pymupdf.open(str(pdf))
    try:
        import pymupdf4llm
        try: chunks = pymupdf4llm.to_markdown(doc, page_chunks=True, use_ocr=False)  # eigene OCR (deutsch) statt Standard-OCR
        except TypeError: chunks = pymupdf4llm.to_markdown(doc, page_chunks=True)
        out = [c["text"] for c in chunks]
        if len(out) == len(doc): return out
    except Exception:
        pass
    return [pg.get_text() for pg in doc]

# ---------- entwurf ----------
def inline_runs(p, text, size=None):
    for teil in re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*)", text):
        if not teil: continue
        if teil.startswith("**"): r = p.add_run(teil[2:-2]); r.bold = True
        elif teil.startswith("*") and len(teil) > 2: r = p.add_run(teil[1:-1]); r.italic = True
        else: r = p.add_run(teil)
        if size: r.font.size = size

H1 = re.compile(r"^(?:#{1,2}\s+)?((?:I|II|III|IV|V|VI|VII|VIII|IX|X)\.\s+\S.*)$")
H2 = re.compile(r"^(?:#{3,4}\s+)?((?:I|II|III|IV|V|VI|VII|VIII|IX|X)\.\d+\.?\s+\S.*)$")

def skelett(mit_modulen):
    z = [("## I. Untersuchungsumstände und Grundlagen", ""), ("## II. Anamnese", ""),
         ("### II.1. Eigenanamnese (Angaben des Betroffenen)", ""), ("### II.2. Fremdanamnese", ""),
         ("### II.3. Tragende Vorbefunde", ""), ("## III. Psychopathologischer Befund", ""), ("## IV. Beurteilung", "")]
    z += [(f"### IV.{i}. {t}", "") for i, t in enumerate(mit_modulen, 1)]
    z.append(("## V. Zusammenfassung", ""))
    return "\n\n".join(h for h, _ in z) + "\n"

def cmd_entwurf(quelle_arg="__argv__"):
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
    ordner = find_fall(); fp = ordner / "fall.json"
    if quelle_arg == "__argv__": quelle_arg = sys.argv[2] if len(sys.argv) > 2 else None
    if quelle_arg:
        q = Path(quelle_arg); q = q if q.exists() else ordner / SUB["gutachten"] / str(quelle_arg)
        if not q.exists(): sys.exit(f"Datei nicht gefunden: {quelle_arg}")
        if q.suffix.lower() == ".docx": return klarnamen_in_docx(q, ordner)
    fall = json.loads(fp.read_text(encoding="utf-8"))
    for key, frage, default in [
        ("betroffene_person", "Betroffene Person (Name, geb. TT.MM.JJJJ)", ""),
        ("adresse", "Adresse der betroffenen Person", ""),
        ("richter", "Richter/in am Amtsgericht (Name)", ""),
        ("untersuchungsort", "Untersuchungsort", "@adresse"),
        ("gutachten_auftrag", "Gutachtenauftrag (Text)", f"Beschluss des AG {fall['gericht']} vom {fall['beschluss_datum']} zu den Voraussetzungen einer Betreuerbestellung nach § 1814 Abs. 5 BGB")]:
        if not fall.get(key): fall[key] = ask(frage, fall.get("adresse", "") if default == "@adresse" else default)
    write_atomic(fp, lambda p: p.write_text(json.dumps(fall, ensure_ascii=False, indent=2), encoding="utf-8"))

    if quelle_arg:
        text = q.read_text(encoding="utf-8")
    else:
        mods = CFG["gutachten_module_IV"]
        if not WEB:
            print("Kein Textentwurf angegeben – es wird ein Gerüst erstellt. Module für Abschnitt IV (Nummern, Enter = alle):")
            for i, m in enumerate(mods, 1): print(f"  {i}. {m}")
        wahl = ask("Auswahl, z. B. 1,2,3,5", "")
        idx = [int(x) for x in re.findall(r"\d+", wahl)] if wahl else range(1, len(mods) + 1)
        text = skelett([mods[i - 1] for i in idx if 1 <= i <= len(mods)])

    text = rueckwaerts(text, lade_pseudonyme(ordner))
    offen = sorted(set(re.findall(r"\b(?:Betroffene|Person|Ort)_[A-Z](?:_[NV])?\b|\[NAME\]|\[TELEFON\]|\[E-MAIL\]|\[IBAN\]", text)))

    import vorlagen_engine as VE
    d = Document(str(vorlage_pfad("gutachten")))
    marker = VE.finde_marker(d, "TEXT")
    def para(text="", bold=False, size=None, italic=False, align=None, before=None, after=None, indent=None):
        p = d.add_paragraph()
        if marker is not None: marker._p.addprevious(p._p)
        if text: r = p.add_run(text); r.bold = bold; r.italic = italic; r.font.size = Pt(size) if size else None
        if align: p.alignment = align
        if before is not None: p.paragraph_format.space_before = Pt(before)
        if after is not None: p.paragraph_format.space_after = Pt(after)
        if indent: p.paragraph_format.left_indent = Cm(indent); p.paragraph_format.first_line_indent = Cm(-0.5)
        return p
    if marker is None: print("HINWEIS: In der Gutachten-Vorlage fehlt der Platzhalter {{TEXT}} - der Text konnte nicht eingefügt werden.")
    else:
        zeilen = text.splitlines(); puffer = []
        def absatz():
            if puffer: inline_runs(para(after=6), " ".join(puffer)); puffer.clear()
        for z in zeilen:
            s_ = z.strip()
            if not s_: absatz(); continue
            if re.match(r"^#\s+\w*\s*Gutachten", s_, re.I): continue
            m1, m2 = H1.match(s_), H2.match(s_)
            if m1: absatz(); para(m1.group(1).replace("**", ""), bold=True, size=14, before=14, after=4)
            elif m2: absatz(); para(m2.group(1).replace("**", ""), bold=True, before=8, after=3)
            elif s_.startswith("#"): absatz(); para(s_.lstrip("# ").replace("**", ""), bold=True, before=8, after=3)
            elif re.match(r"^[-*•]\s+", s_): absatz(); inline_runs(para(after=2, indent=1.0), "•  " + re.sub(r"^[-*•]\s+", "", s_))
            else: puffer.append(s_)
        absatz()
        marker._p.getparent().remove(marker._p)
    VE.bloecke(d, {}); VE.nummeriere(d)
    rest = VE.fuelle(d, basis_werte(fall, "gutachten"))
    if rest: print("Hinweis: Diese Platzhalter der Vorlage sind unbekannt und blieben stehen:", ", ".join("{{" + k + "}}" for k in rest))

    if not fall.get("gutachten_datum"): print("HINWEIS: Das Datum der Gutachtenerstattung fehlt in den Falldaten (Unterschrift).")
    out = next_version(ordner / SUB["gutachten"], f"Gutachten_{safe(fall['aktenzeichen'])}")
    write_atomic(out, lambda p: d.save(p))
    print(f"Entwurf erstellt: {out.name}")
    if offen: print("ACHTUNG: Diese Platzhalter wurden nicht zurückgetauscht (nicht in pseudonyme.json):", ", ".join(offen))
    print("Hinweis: Geburtsdaten im Fließtext sind nicht umkehrbar ([TT.MM.]JJJJ) – bitte prüfen.")


# ---------- transkript ----------
def zeit(sek):
    sek = int(sek); return f"{sek // 3600:02d}:{sek % 3600 // 60:02d}:{sek % 60:02d}"

def pseudo_name(name, eintraege):
    """Dateiname für die KI ohne Klarnamen (Namen und Teilnamen werden ersetzt)."""
    stem = re.sub(r"[_\-]+", " ", Path(name).stem)
    t, _ = pseudonymisieren(stem, eintraege)
    return safe(t) or "Dokument"

def transkript_ausgabename(quelle, eintraege): return f"Transkript_{pseudo_name(quelle, eintraege)}_pseudonymisiert.md"

def gespraech_quellen(ordner):
    """Texte in 02_Gespraeche (Transkripte, eigene Notizen, E-Mails als .md/.txt), die pseudonymisiert werden können."""
    d = ordner / SUB["gespraeche"]
    if not d.exists(): return []
    return [p for p in sorted(d.iterdir(), key=natsort) if p.is_file() and p.suffix.lower() in TEXT_ENDUNGEN
            and not p.stem.endswith(("_mit_Zeit", "_pseudonymisiert")) and not p.name.startswith(("~", "."))]

def cmd_transkript():
    ordner = find_fall(); g = ordner / SUB["gespraeche"]
    offen = [p for p in sorted(g.iterdir(), key=natsort)
             if p.is_file() and p.suffix.lower() in AUDIO_ENDUNGEN and not (g / (p.stem + ".md")).exists() and not (g / (p.stem + ".txt")).exists()]
    if not offen:
        print("Keine neuen Aufnahmen in 02_Gespraeche. (Zum Wiederholen die zugehörige .md-Datei löschen.)"); return
    try:
        from faster_whisper import WhisperModel
    except ImportError:
        sys.exit("faster-whisper fehlt. Bitte im Assistenten unter Einstellungen > Programme installieren.")
    modell = CFG.get("whisper_modell", "large-v3-turbo")
    print(f"{len(offen)} Aufnahme(n). Modell '{modell}' wird geladen (beim ersten Mal einmaliger Download, danach läuft alles offline) ...")
    wm = WhisperModel(modell, device="cpu", compute_type="int8")
    for p in offen:
        print(f"\n>> {p.name}")
        try:
            segs, info = wm.transcribe(str(p), language="de", beam_size=5, vad_filter=True,
                                       condition_on_previous_text=False, initial_prompt=CFG.get("whisper_prompt"))
        except TypeError as e:
            if "metadata_errors" in str(e):
                sys.exit("Die Audio-Bibliothek 'av' ist zu alt. Einmalig im Terminal/PowerShell ausführen:\n"
                         "    pip install --upgrade av faster-whisper\nDann erneut transkribieren.")
            raise
        absaetze, aktuell, start, letztes_ende = [], [], 0.0, 0.0
        for sg in segs:
            t = sg.text.strip()
            if not t: continue
            if aktuell and sg.start - letztes_ende > 2.0: absaetze.append((start, " ".join(aktuell))); aktuell = []
            if not aktuell: start = sg.start
            aktuell.append(t); letztes_ende = sg.end
            if info.duration: print(f"\r   {min(100, int(sg.end / info.duration * 100))} %", end="", flush=True)
        if aktuell: absaetze.append((start, " ".join(aktuell)))
        print()
        md = (f"# Transkript: {p.name}\n\n*Automatische Transkription (Whisper), lokal erstellt. Sprecher nicht getrennt, Fehler möglich. "
              "Zeitmarke [hh:mm:ss] am Beginn jedes Absatzes.*\n\n" + "\n\n".join(f"[{zeit(s)}] {t}" for s, t in absaetze) + "\n")
        write_atomic(g / (p.stem + ".md"), lambda q, md=md: q.write_text(md, encoding="utf-8"))
        print(f"   fertig: 02_Gespraeche/{p.stem}.md")
    print("\nBitte die Transkripte jetzt Korrektur lesen (Namen und Fachbegriffe werden oft falsch geschrieben).")
    print("Danach: 'Transkripte pseudonymisieren' ausführen, bevor etwas hochgeladen wird.")

def cmd_pseudo():
    ordner = find_fall()
    quellen = gespraech_quellen(ordner)
    if not quellen: sys.exit("Keine Transkripte oder Texte (.md/.txt) in 02_Gespraeche gefunden.")
    eintraege = pseudonyme_pflegen(ordner)
    for p in quellen:
        t, zaehler = pseudonymisieren(p.read_text(encoding="utf-8-sig", errors="replace"), eintraege)
        if p.suffix.lower() == ".txt": t = f"# Gesprächs-/Kommunikationsdokument: {pseudo_name(p.name, eintraege)}\n\n" + t
        (ordner / SUB["ki"]).mkdir(exist_ok=True)
        out = ordner / SUB["ki"] / transkript_ausgabename(p.name, eintraege)
        write_atomic(out, lambda q, t=t: q.write_text(t, encoding="utf-8"))
        print(f"{p.name} -> 05_KI-Upload/{out.name}: " + (", ".join(f"{k}: {v}" for k, v in zaehler.items() if v) or "keine Ersetzungen"))
    print("Fertig. Die Dateien in 05_KI-Upload sind die einzigen, die für die KI gedacht sind. Vor dem Hochladen kurz überfliegen.")

# ---------- KI-Prompt für den Entwurf ----------
EXPL_ARTEN = ["Persönliche Exploration der betroffenen Person", "Fremdanamnese (persönlich)", "Fremdanamnese (telefonisch)",
              "E-Mail-Kommunikation", "Schriftlicher Kontakt (Brief, Fax)", "Telefonat (z. B. Terminvereinbarung)",
              "Anhörung / Termin beim Gericht", "Gespräch mit Betreuer/Behörde/Einrichtung", "Sonstiges"]

ARBEITSWEISEN = {
    "schrittweise": ("Schrittweise mit Rückfragen",
        "Gehe schrittweise vor:\n"
        "1. Lies zuerst alle Unterlagen. Fasse in höchstens 10 Zeilen zusammen, was du verstanden hast (nur Sachstand, keine Bewertung).\n"
        "2. Nenne Lücken, Widersprüche und Rückfragen (z. B. fehlende Befunde, unklare Zeitangaben) und warte auf meine Antwort.\n"
        "3. Erst danach erstellst du den Entwurf, und zwar abschnittsweise (ein Abschnitt je Nachricht). Warte nach jedem Abschnitt auf meine Korrekturen."),
    "komplett": ("Kompletter Entwurf sofort",
        "Erstelle direkt einen vollständigen Entwurf aller Abschnitte. Markiere Lücken im Text mit [ANGABE FEHLT: ...] und liste am Ende alle offenen Fragen auf."),
}

def ki_dateien(ordner):
    """Für die KI bestimmte Dateien: pseudonymisierte Texte in 05_KI-Upload (Akte und Transkripte)."""
    d = ordner / SUB["ki"]
    if not d.exists(): return []
    return [p for p in sorted(d.iterdir(), key=natsort) if p.is_file() and "_pseudonymisiert" in p.name
            and p.suffix.lower() in TEXT_ENDUNGEN and not p.name.startswith(("~", "."))]

def _datum_key(s):
    m = re.fullmatch(r"\s*(\d{1,2})\.\s?(\d{1,2})\.\s?(\d{4})\s*", s or "")
    return (0, int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else (1, 0, 0, 0)

def sortierte_explorationen(fall):
    return sorted([e for e in fall.get("explorationen", []) if any(str(e.get(k, "")).strip() for k in ("art", "datum", "beteiligte", "notiz", "quelle"))],
                  key=lambda e: _datum_key(e.get("datum")))

def ki_prompt(ordner, fall, module_idx=None, arbeitsweise="schrittweise", dateien=None, extra=None):
    """Baut den Prompt für die KI. Alle freien Texte werden mit den geschützten Namen des Falls pseudonymisiert.
    Es werden bewusst weder Aktenzeichen noch Namen von Gericht, Richter/in, Betroffenen oder Gutachter eingefügt.
    Gibt {"prompt", "dateien" (Namen), "hinweise"} zurück."""
    eint = lade_pseudonyme(ordner)
    P = lambda s: pseudonymisieren(str(s or ""), eint)[0].strip()
    hinweise = []
    if not any(e.get("name", "").strip() for e in eint):
        hinweise.append("Es sind keine geschützten Namen eingetragen: Auftrag, Notizen und Dateien werden NICHT pseudonymisiert.")
    verfuegbar = ki_dateien(ordner)
    gewaehlt = [p for p in verfuegbar if dateien is None or p.name in dateien]
    if not verfuegbar: hinweise.append("In 05_KI-Upload liegen noch keine pseudonymisierten Dateien (Schritt 'Akte vorbereiten' bzw. 'Transkripte pseudonymisieren').")
    elif not gewaehlt: hinweise.append("Es ist keine Datei ausgewählt.")
    expl = sortierte_explorationen(fall)
    if not expl: hinweise.append("Es sind keine Explorationen/Kontakte eingetragen.")
    if any(not _datum_key(e.get("datum"))[0] == 0 for e in expl): hinweise.append("Bei mindestens einer Exploration fehlt das Datum (Format TT.MM.JJJJ).")
    auftrag = P(fall.get("gutachten_auftrag")); details = P(fall.get("auftrag_details"))
    if not (auftrag or details): hinweise.append("Der Auftrag ist leer.")

    mods = CFG.get("gutachten_module_IV", [])
    idx = module_idx if module_idx else range(1, len(mods) + 1)
    gliederung = skelett([mods[i - 1] for i in idx if 1 <= i <= len(mods)]).strip()
    titel = CFG.get("gutachten_titel") or "Gutachten"; fach = CFG.get("fachgebiet") or ""

    # Zuordnung Transkript-Datei -> Exploration
    nach_datei = {}
    for n, e in enumerate(expl, 1):
        if str(e.get("quelle", "")).strip(): nach_datei[transkript_ausgabename(e["quelle"], eint)] = n
    z = []
    z.append(f"# Aufgabe\nDu unterstützt einen ärztlichen Sachverständigen ({fach or 'Fachgebiet nicht angegeben'}) beim Entwurf eines Gutachtens („{titel}“) "
             "im Betreuungsverfahren (Gutachten zu den Voraussetzungen einer Betreuerbestellung nach § 1814 BGB; Anforderungen an das Gutachten nach § 280 FamFG). "
             "Verantwortlich für Inhalt und Beurteilung ist ausschließlich der Sachverständige; du lieferst einen Entwurf zur fachlichen Prüfung und Überarbeitung.")
    ph = [f"- {e['platzhalter']}: {'Person' if e.get('art') == 'person' else 'Ort / Sonstiges'}" for e in eint if e.get("name", "").strip() and e.get("platzhalter", "").strip()]
    z.append("# Pseudonymisierung\nAlle Namen, Orte und Kontaktdaten in den Unterlagen sind durch Platzhalter ersetzt. Verwende ausschließlich diese Platzhalter, "
             "versuche nicht, sie aufzulösen, und erfinde keine Namen. Bei Personen gibt es Varianten für Nachname (…_N) und Vorname (…_V). "
             "Geburtsdaten sind teilweise unkenntlich gemacht ([TT.MM.]JJJJ), Telefonnummern, E-Mail-Adressen und IBAN sind durch [TELEFON], [E-MAIL], [IBAN] ersetzt."
             + ("\nPlatzhalter in diesem Fall:\n" + "\n".join(ph) if ph else ""))
    a = ["# Auftrag"]
    if fall.get("beschluss_datum"): a.append(f"- Beschluss des Betreuungsgerichts vom: {fall['beschluss_datum']}")
    if fall.get("auftrag_datum"): a.append(f"- Auftrag eingegangen am: {fall['auftrag_datum']}")
    if fall.get("untersuchungsdatum"): a.append(f"- Untersuchung (persönliche Exploration) am: {fall['untersuchungsdatum']}")
    if auftrag: a.append(f"- Auftrag (Kurzfassung): {auftrag}")
    if details: a.append("- Fragestellung und Vorgaben des Gerichts (Wortlaut bzw. Zusammenfassung):\n" + "\n".join("  " + l for l in details.splitlines()))
    z.append("\n".join(a))
    if expl:
        l = ["# Explorationen und Kontakte (chronologisch)",
             "Gib im Gutachten bei jeder Information an, aus welcher Exploration bzw. Quelle (mit Datum) sie stammt."]
        for n, e in enumerate(expl, 1):
            s = f"{n}. {str(e.get('datum', '')).strip() or '[Datum fehlt]'} – {str(e.get('art', '')).strip() or 'Kontakt'}"
            if str(e.get("beteiligte", "")).strip(): s += f"; Beteiligte: {P(e['beteiligte'])}"
            if str(e.get("ort", "")).strip(): s += f"; Ort: {P(e['ort'])}"
            quelle = next((k for k, v in nach_datei.items() if v == n), "")
            if quelle: s += f"; Datei: {quelle}" + ("" if quelle in [p.name for p in gewaehlt] else " (noch nicht vorhanden/ausgewählt)")
            if str(e.get("notiz", "")).strip(): s += "\n   Notiz des Sachverständigen: " + P(e["notiz"]).replace("\n", "\n   ")
            l.append(s)
        z.append("\n".join(l))
    if gewaehlt:
        l = ["# Unterlagen (als Dateien angehängt)"]
        for p in gewaehlt:
            if p.name.startswith("Akte_pseudonymisiert"):
                s = "Gerichtsakte, zusammengeführt; Seitenmarken [Seite N]; per OCR gelesene Seiten sind markiert und können Lesefehler enthalten. Zitiere Fundstellen als [Seite N]."
            else:
                n = nach_datei.get(p.name)
                s = ("Transkript, Notiz oder E-Mail" + (f" zu Exploration Nr. {n}" if n else "") + ". Bei automatischen Transkripten: Zeitmarken [hh:mm:ss] je Absatz, Sprecher nicht getrennt, Erkennungsfehler möglich.")
            l.append(f"- {p.name}: {s}")
        z.append("\n".join(l))
    else:
        z.append("# Unterlagen\nEs sind keine Dateien angehängt. Arbeite nur mit den Angaben in diesem Prompt und frage nach.")
    z.append("# Arbeitsweise\n" + ARBEITSWEISEN.get(arbeitsweise, ARBEITSWEISEN["schrittweise"])[1])
    z.append("# Regeln für den Entwurf\n"
             "- Verwende nur Informationen aus den Unterlagen und diesem Prompt. Erfinde nichts; wo etwas fehlt, schreibe [ANGABE FEHLT: ...].\n"
             "- Trenne klar: Angaben der betroffenen Person, Angaben Dritter (Fremdanamnese; indirekte Rede/Konjunktiv), Befunde aus der Akte und eigene Beobachtungen des Sachverständigen.\n"
             "- Den psychopathologischen Befund und die Beurteilung stützt du ausschließlich auf dokumentierte Beobachtungen aus der Untersuchung (Transkript, Notizen). "
             "Liegen keine vor, setze [VOM SACHVERSTÄNDIGEN ZU ERGÄNZEN].\n"
             "- Sachlich, neutral, medizinisch-fachlich. Keine rechtlichen Entscheidungen; Aussagen zu Aufgabenkreisen, Einwilligungsvorbehalt und Dauer sind fachliche Einschätzungen, die der Sachverständige prüft.\n"
             "- Weise auf Widersprüche zwischen Quellen hin, statt sie zu glätten.")
    z.append("# Gliederung und Ausgabeformat\nGliedere den Entwurf so:\n\n```\n" + gliederung + "\n```\n"
             "Gib den Text als Markdown im Chat aus (Überschriften wie oben mit ## und ###, Absätze, Aufzählungen mit „-“; keine Tabellen, keine Dateien), "
             "damit ich ihn kopieren und in meine Word-Vorlage übernehmen kann.")
    zus = (extra if extra is not None else CFG.get("ki_prompt_zusatz", "")).strip()
    if zus: z.append("# Eigene Hinweise des Sachverständigen\n" + zus)
    text = "\n\n".join(z) + "\n"
    verdacht = sorted(set(re.findall(r"\b(?:Herrn?|Frau|Fr\.)\s+(?:(?:Dr\.|Prof\.)\s+)*([A-ZÄÖÜ][\wäöüß-]+)", text)))
    verdacht = [v for v in verdacht if "_" not in v]
    if verdacht: hinweise.append("Im Prompt stehen mögliche Namen, die nicht geschützt sind: " + ", ".join(verdacht[:10]) + ". Bitte unter 'Geschützte Namen' ergänzen und Prompt neu erzeugen.")
    return {"prompt": text, "dateien": [p.name for p in gewaehlt], "hinweise": hinweise}

# ---------- Namen in Word-Datei zurücktauschen ----------
def paare_rueckwaerts(eintraege):
    paare = {}
    for e in eintraege:
        n, ph = e["name"].strip(), e["platzhalter"]; paare[ph] = n
        t = n.split()
        if e.get("art") == "person" and len(t) >= 2:
            if len(t[-1]) >= 3: paare.setdefault(ph + "_N", t[-1])
            if len(" ".join(t[:-1])) >= 3: paare.setdefault(ph + "_V", " ".join(t[:-1]))
    return paare

def ersetze_im_absatz(p, paare):
    n = 0
    while True:
        runs = p.runs; voll = "".join(r.text for r in runs)
        best = None
        for ph, orig in paare.items():
            i = voll.find(ph)
            if i >= 0 and (best is None or i < best[0] or (i == best[0] and len(ph) > len(best[1]))): best = (i, ph, orig)
        if not best: return n
        i, ph, orig = best; ende = i + len(ph); pos = 0; erster = True
        for r in runs:
            t = r.text; rs, re_ = pos, pos + len(t); pos = re_
            if re_ <= i or rs >= ende: continue
            a, b = max(i, rs) - rs, min(ende, re_) - rs
            if erster: r.text = t[:a] + orig + t[b:]; erster = False
            else: r.text = t[:a] + t[b:]
        n += 1

def klarnamen_in_docx(quelle, ordner):
    from docx import Document
    eintraege = lade_pseudonyme(ordner)
    if not eintraege: sys.exit("Keine Pseudonyme gespeichert (_intern/pseudonyme.json). Zuerst 'akte' ausführen.")
    paare = paare_rueckwaerts(eintraege)
    d = Document(str(quelle)); n = 0
    def tabellen(tbls):
        k = 0
        for t in tbls:
            for row in t.rows:
                for c in row.cells:
                    for p in c.paragraphs: k += ersetze_im_absatz(p, paare)
                    k += tabellen(c.tables)
        return k
    for p in d.paragraphs: n += ersetze_im_absatz(p, paare)
    n += tabellen(d.tables)
    for s in d.sections:
        for teil in (s.header, s.footer):
            for p in teil.paragraphs: n += ersetze_im_absatz(p, paare)
    rest = set()
    for p in d.paragraphs:
        rest |= set(re.findall(r"\b(?:Betroffene|Person|Ort)_[A-Z](?:_[NV])?\b|\[NAME\]|\[TELEFON\]|\[E-MAIL\]|\[IBAN\]|\[TT\.MM\.\]", p.text))
    out = next_version(quelle.parent, quelle.stem + "_Klarnamen")
    write_atomic(out, lambda q: d.save(q))
    print(f"{n} Platzhalter zurückgetauscht -> {out.name} (Formatierung von Claude unverändert)")
    if rest: print("ACHTUNG, noch nicht ersetzt (bitte von Hand ergänzen):", ", ".join(sorted(rest)))

# ---------- Menü ----------
def oeffne(pfad):
    import subprocess
    if sys.platform.startswith("win"): os.startfile(str(pfad))
    elif sys.platform == "darwin": subprocess.run(["open", str(pfad)])
    else: subprocess.run(["xdg-open", str(pfad)])

def _still(): return getattr(__import__("subprocess"), "CREATE_NO_WINDOW", 0)

def kopiere_dateien_in_zwischenablage(pfade):
    """Legt DATEIEN (nicht Text) in die Zwischenablage, sodass sie im Chat der KI mit Strg+V / Cmd+V eingefügt werden können.
    Windows: PowerShell Set-Clipboard. Mac: AppleScript. Gibt True zurück, wenn der Befehl ohne Fehler lief."""
    import subprocess
    pfade = [str(Path(p).resolve()) for p in pfade]
    if not pfade: return False
    try:
        if sys.platform.startswith("win"):
            lst = ",".join("'" + p.replace("'", "''") + "'" for p in pfade)
            r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", f"Set-Clipboard -LiteralPath {lst}"],
                               capture_output=True, text=True, creationflags=_still())
            return r.returncode == 0
        if sys.platform == "darwin":
            lst = ", ".join('(POSIX file "' + p.replace("\\", "\\\\").replace('"', '\\"') + '") as alias' for p in pfade)
            r = subprocess.run(["osascript", "-e", f"set the clipboard to {{{lst}}}"], capture_output=True, text=True)
            return r.returncode == 0
    except Exception:
        return False
    return False

def zeige_dateien(ordner, pfade):
    """Öffnet den Ordner und markiert (soweit das System es kann) die Dateien, damit man sie in den Chat ziehen kann."""
    import subprocess
    ordner = Path(ordner); pfade = [Path(p).resolve() for p in pfade]
    try:
        if sys.platform.startswith("win") and pfade:
            import ctypes
            sh = ctypes.windll.shell32; sh.ILCreateFromPathW.restype = ctypes.c_void_p; sh.ILFree.argtypes = [ctypes.c_void_p]
            ctypes.windll.ole32.CoInitialize(None)
            fo = sh.ILCreateFromPathW(str(ordner.resolve())); items = [sh.ILCreateFromPathW(str(p)) for p in pfade]
            arr = (ctypes.c_void_p * len(items))(*items)
            sh.SHOpenFolderAndSelectItems.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_void_p, ctypes.c_ulong]
            hr = sh.SHOpenFolderAndSelectItems(fo, len(items), arr, 0)
            for x in [fo, *items]: sh.ILFree(x)
            if hr == 0: return
        elif sys.platform == "darwin" and pfade:
            lst = ", ".join('POSIX file "' + str(p).replace("\\", "\\\\").replace('"', '\\"') + '"' for p in pfade)
            r = subprocess.run(["osascript", "-e", f'tell application "Finder" to reveal {{{lst}}}', "-e", 'tell application "Finder" to activate'], capture_output=True)
            if r.returncode == 0: return
    except Exception:
        pass
    oeffne(ordner)

def kopiere_text(text):
    """Text in die Zwischenablage (Rückfall, falls der Browser das nicht erlaubt)."""
    import subprocess
    try:
        if sys.platform.startswith("win"): subprocess.run(["clip"], input=text.encode("utf-16"), creationflags=_still()); return True
        if sys.platform == "darwin": subprocess.run(["pbcopy"], input=text.encode("utf-8")); return True
        for cmd in (["xclip", "-selection", "clipboard"], ["wl-copy"]):
            try: subprocess.run(cmd, input=text.encode("utf-8")); return True
            except FileNotFoundError: continue
    except Exception:
        pass
    return False

def faelle():
    if not BASE or not BASE.exists(): return []
    out = [f.parent for j in BASE.iterdir() if j.is_dir() and j.name.isdigit() for f in j.glob("*/fall.json")]
    return sorted(out, key=lambda p: (p / "fall.json").stat().st_mtime, reverse=True)

def sicher(fn, *a):
    try: fn(*a)
    except SystemExit as e:
        if e.code not in (None, 0): print(e.code)
    except Exception as e:
        print(f"Fehler: {e}")

def wahl_entwurf(ordner):
    dateien = [p for p in sorted((ordner / SUB["gutachten"]).iterdir())
               if p.suffix.lower() in (".docx", ".md", ".txt") and not p.name.startswith(("~", "Gutachten_"))]
    print("\nWoher kommt der Entwurf?\n  0  Leeres Gerüst nach meinem Muster")
    for i, p in enumerate(dateien, 1):
        art = "Word von Claude: Namen zurücktauschen" if p.suffix.lower() == ".docx" else "Text: in mein Muster setzen"
        print(f"  {i}  {p.name}   ({art})")
    w = ask("Nummer", "0")
    return None if w in ("", "0") or not w.isdigit() or int(w) > len(dateien) else dateien[int(w) - 1]

def cmd_menu():
    brauche_ordner()
    print("\n===== GUTACHTEN-ASSISTENT =====")
    while True:
        liste = faelle()
        print("\nFall wählen:")
        for i, p in enumerate(liste[:15], 1): print(f"  {i:2d}  {p.parent.name}/{p.name}")
        print("   n  Neuen Fall anlegen\n   q  Beenden")
        w = ask("Auswahl").lower()
        if w == "q": return
        if w == "n": sicher(cmd_neu); continue
        if not w.isdigit() or not 1 <= int(w) <= len(liste[:15]): continue
        fall = liste[int(w) - 1]; os.chdir(fall)
        while True:
            print(f"\n--- {fall.name} ---\n  1  Akte vorbereiten (PDFs zusammenführen, Text, pseudonymisieren)"
                  "\n  2  Audio transkribieren (lokal)\n  3  Transkript pseudonymisieren (nach Korrekturlesen)"
                  "\n  4  Gutachten-Entwurf erstellen / Namen zurücktauschen\n  5  Rechnung erstellen"
                  "\n  6  Fallordner öffnen\n  0  Zurück")
            a = ask("Auswahl")
            if a == "0": break
            elif a == "1": sicher(cmd_akte)
            elif a == "2": sicher(cmd_transkript)
            elif a == "3": sicher(cmd_pseudo)
            elif a == "4": sicher(cmd_entwurf, wahl_entwurf(fall))
            elif a == "5": sicher(cmd_rechnung)
            elif a == "6": oeffne(fall)
            input("\n(Enter drücken, um fortzufahren)")

if __name__ == "__main__":
    cmds = {"neu": cmd_neu, "rechnung": cmd_rechnung, "akte": cmd_akte, "entwurf": cmd_entwurf, "transkript": cmd_transkript, "pseudo": cmd_pseudo, "menu": cmd_menu, "tessdata": cmd_tessdata, "modell": cmd_modell}
    if len(sys.argv) < 2: cmd_menu()
    elif sys.argv[1] not in cmds: sys.exit("Aufruf: python gutachten.py [menu|neu|akte|transkript|pseudo|entwurf|rechnung]")
    else:
        try: cmds[sys.argv[1]]()
        except SystemExit: raise
        except Exception:
            import traceback; print("\nEin Fehler ist aufgetreten. Bitte diesen Text kopieren und weitergeben:\n"); traceback.print_exc(); sys.exit(1)
