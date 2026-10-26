"""Einfache Vorlagen-Engine für Word-Dateien (.docx) mit {{PLATZHALTERN}}.
Unterstützt: Felder {{NAME}}, Wiederholzeile {{POS_...}}, Bedingungen {{#NAME}}...{{/NAME}} (Absätze) bzw. {{#NAME}} in einer Tabellenzeile,
Nummerierung {{NR}} und einen Textbereich {{TEXT}}."""
import re
from copy import deepcopy
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

PH = re.compile(r"\{\{\s*([#/]?[A-Za-z0-9_]+)\s*\}\}")

BESCHREIBUNG = {
    "GUTACHTER_NAME": "Name des Gutachters", "GUTACHTER_TITEL": "Qualifikation / Titel (z. B. Facharzt für ...)",
    "GUTACHTER_STRASSE": "Straße und Nummer", "GUTACHTER_PLZ_ORT": "PLZ und Ort", "GUTACHTER_ADRESSE": "Straße, PLZ Ort in einer Zeile",
    "FACHGEBIET": "Fachgebiet (z. B. Psychiatrie)", "GUTACHTEN_TITEL": "Titel des Gutachtens (z. B. Psychiatrisches Gutachten)", "ORT": "Ort für die Unterschrift",
    "KONTOINHABER": "Kontoinhaber", "IBAN": "IBAN", "BIC": "BIC", "BANK": "Bank",
    "GERICHT": "Amtsgericht (nur der Ort)", "ABTEILUNG": "Abteilung des Gerichts", "AKTENZEICHEN": "Aktenzeichen",
    "BESCHLUSS_DATUM": "Datum des Beschlusses", "UNTERSUCHUNGSDATUM": "Datum der Untersuchung",
    "GUTACHTEN_DATUM": "Datum der Gutachtenerstattung (Unterschrift)", "RECHNUNGSDATUM": "Datum der Rechnung",
    "AUFTRAG": "Gutachtenauftrag (Text)", "HONORAR_TEXT": "Einleitungstext zum Honorar",
    "HONORARGRUPPE": "z. B. M 3", "HONORARSATZ": "Stundensatz, z. B. 131,00 EUR",
    "STUNDEN_GESAMT": "Gesamtzeit, z. B. 9 Std.", "STUNDEN_ZAHL": "Gesamtzeit als Zahl, z. B. 9", "HONORAR_BETRAG": "Honorar in EUR",
    "POS_TITEL": "Leistung: Bezeichnung (Wiederholzeile)", "POS_TEXT": "Leistung: Beschreibung (Wiederholzeile)", "POS_STUNDEN": "Leistung: Dauer (Wiederholzeile)",
    "FAHRT_ZIEL": "Fahrtziel", "FAHRT_KM_EINFACH": "Kilometer einfache Strecke", "KM_GESAMT": "Kilometer gesamt (Hin und Zurück)",
    "KM_SATZ": "Kilometersatz", "FAHRTKOSTEN": "Fahrtkosten in EUR",
    "SCHREIB_EINHEITEN": "Einheiten Schreibgebühr", "ANSCHLAEGE": "Anschläge des Gutachtens", "SCHREIB_SATZ": "Satz je Einheit", "SCHREIBGEBUEHR": "Schreibgebühr in EUR",
    "GESAMTBETRAG": "Gesamtbetrag der Rechnung",
    "BETROFFENE_PERSON": "Betroffene Person (Name, geb.)", "ADRESSE_BETROFFENE": "Adresse der betroffenen Person",
    "RICHTER": "Richter/in", "UNTERSUCHUNGSORT": "Untersuchungsort",
    "TEXT": "Gutachten-Text wird hier eingefügt (nur Gutachten-Vorlage)",
    "NR": "Fortlaufende Abschnittsnummer (1, 2, 3 ...)",
    "#FAHRT ... /FAHRT": "Block nur wenn Fahrtkosten vorhanden", "#SCHREIB ... /SCHREIB": "Block nur wenn Schreibgebühr vorhanden",
}
PFLICHT = {"rechnung": {"POS_TITEL", "GESAMTBETRAG"}, "gutachten": {"TEXT"}}

def alle_absaetze(doc):
    teile = [doc.element.body]
    for s in doc.sections:
        for hf in (s.header, s.footer):
            try:
                if not hf.is_linked_to_previous: teile.append(hf._element)
            except Exception: pass
    for t in teile:
        for p in list(t.iter(qn("w:p"))): yield Paragraph(p, doc)

def _span(runs, a, b, neu):
    pos, erster = 0, True
    for r in runs:
        t = r.text; rs, re_ = pos, pos + len(t); pos = re_
        if re_ <= a or rs >= b: continue
        x, y = max(a, rs) - rs, min(b, re_) - rs
        if erster: r.text = t[:x] + neu + t[y:]; erster = False
        else: r.text = t[:x] + t[y:]

def ersetze_absatz(p, fn, marker=False):
    pos = 0
    while True:
        runs = p.runs; voll = "".join(r.text for r in runs)
        m = PH.search(voll, pos)
        if not m: return
        key = m.group(1)
        if key[0] in "#/" and not marker: pos = m.end(); continue
        val = fn(key)
        if val is None: pos = m.end(); continue
        val = str(val); _span(runs, m.start(), m.end(), val); pos = m.start() + len(val)

def _txt(el): return "".join(t.text or "" for t in el.iter(qn("w:t")))

def bloecke(doc, aktiv):
    body = doc.element.body; start = None
    for el in list(body):
        t = _txt(el).strip() if el.tag == qn("w:p") else ""
        if start is None:
            m = re.fullmatch(r"\{\{#(\w+)\}\}", t)
            if m: start = (m.group(1), [el])
            continue
        start[1].append(el)
        if re.fullmatch(r"\{\{/%s\}\}" % start[0], t):
            name, els = start; start = None
            if not aktiv.get(name, True):
                for e in els: body.remove(e)
            else:
                body.remove(els[0]); body.remove(els[-1])
    for tr in list(body.iter(qn("w:tr"))):
        namen = re.findall(r"\{\{#(\w+)\}\}", _txt(tr))
        if not namen: continue
        if any(not aktiv.get(n, True) for n in namen): tr.getparent().remove(tr); continue
        for p in tr.iter(qn("w:p")):
            ersetze_absatz(Paragraph(p, doc), lambda k: "" if k.startswith("#") and k[1:] in namen else None, marker=True)

def wiederhole_zeilen(doc, positionen, prefix="POS_"):
    for tr in list(doc.element.body.iter(qn("w:tr"))):
        if "{{" + prefix not in _txt(tr).replace("{{ ", "{{"): continue
        eltern = tr.getparent()
        for pos in positionen:
            klon = deepcopy(tr); eltern.insert(eltern.index(tr), klon)
            for p in klon.iter(qn("w:p")): ersetze_absatz(Paragraph(p, doc), lambda k, pos=pos: pos.get(k))
        eltern.remove(tr)

def nummeriere(doc):
    n = 0
    for p in doc.element.body.iterchildren(qn("w:p")):
        if re.search(r"\{\{\s*NR\s*\}\}", _txt(p)):
            n += 1; ersetze_absatz(Paragraph(p, doc), lambda k, n=n: str(n) if k == "NR" else None)

def finde_marker(doc, key):
    for p in doc.element.body.iterchildren(qn("w:p")):
        if re.fullmatch(r"\{\{\s*%s\s*\}\}" % key, _txt(p).strip()): return Paragraph(p, doc)
    return None

def fuelle(doc, werte):
    for p in alle_absaetze(doc): ersetze_absatz(p, lambda k: werte.get(k))
    rest = set()
    for p in alle_absaetze(doc):
        for m in PH.finditer("".join(r.text for r in p.runs)):
            if m.group(1)[0] not in "#/": rest.add(m.group(1))
    return sorted(rest)

def platzhalter_in(pfad):
    d = Document(str(pfad)); keys = set()
    for p in alle_absaetze(d):
        for m in PH.finditer("".join(r.text for r in p.runs)): keys.add(m.group(1))
    return sorted(keys)

def pruefe(pfad, art):
    keys = set(platzhalter_in(pfad))
    einfach = {k.lstrip("#/") for k in keys}
    bekannt = {k.split(" ")[0] for k in BESCHREIBUNG} | {"FAHRT", "SCHREIB"}
    return {"gefunden": sorted(einfach), "unbekannt": sorted(einfach - bekannt), "fehlt": sorted(PFLICHT[art] - einfach)}
