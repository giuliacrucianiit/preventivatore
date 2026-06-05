"""
Preventivatore Web — backend Flask
"""
import json
import io
import os
from pathlib import Path
from datetime import date

from flask import Flask, jsonify, request, send_file, render_template_string

BASE = Path(__file__).parent
DATA = BASE / "data"
DATA.mkdir(exist_ok=True)

CLIENTI_FILE = DATA / "clienti.json"
IMPOSTAZIONI_FILE = DATA / "impostazioni.json"

IMPOSTAZIONI_DEFAULT = {
    "azienda": {
        "nome": "", "indirizzo": "", "partita_iva": "",
        "telefono": "", "email": ""
    },
}

app = Flask(__name__)
_voci_cache: list[dict] = []

def carica_prezziario() -> list[dict]:
    global _voci_cache
    if _voci_cache:
        return _voci_cache
    try:
        import openpyxl
    except ImportError:
        return []

    # Legge TUTTI gli xlsx nella cartella data/, qualunque nome abbiano
    files = sorted(DATA.glob("*.xlsx"))
    voci = []
    for percorso in files:
        nome_sorgente = percorso.stem
        try:
            wb = openpyxl.load_workbook(percorso, read_only=True, data_only=True)
        except Exception:
            continue
        for nome_foglio in wb.sheetnames:
            ws = wb[nome_foglio]
            col = _rileva_colonne(ws)
            for row in ws.iter_rows(values_only=True):
                if not row or all(v is None for v in row):
                    continue
                codice = _cell(row, col["codice"])
                descrizione = _cell(row, col["descrizione"])
                if not descrizione or descrizione.lower() in ("descrizione","voce","lavorazione"):
                    continue
                prezzo = _parse_prezzo(_cell(row, col["prezzo"]))
                if prezzo == 0 and not codice:
                    continue
                um = _cell(row, col["um"]) or "cad"
                voci.append({
                    "codice": codice,
                    "descrizione": descrizione,
                    "prezzo": prezzo,
                    "um": um,
                    "sorgente": nome_sorgente,
                    "foglio": nome_foglio,
                })
        wb.close()
    _voci_cache = voci
    return voci

def _rileva_colonne(ws) -> dict:
    col = {"codice": 0, "descrizione": 1, "prezzo": 2, "um": 3}
    kw = {
        "codice": ["codice","cod.","numero d'ordine","numero","n.","articolo"],
        "descrizione": ["descrizione dell'articolo","descrizione","desc.","lavorazione","voce"],
        "prezzo": ["prezzo €","prezzo","importo","€"],
        "um": ["u.m.","um","unità","misura"],
    }
    for row in ws.iter_rows(min_row=1, max_row=5):
        for cell in row:
            if not cell.value: continue
            v = str(cell.value).strip().lower().replace("\n"," ")
            for campo, keys in kw.items():
                if any(k in v for k in keys):
                    col[campo] = cell.column - 1
    return col

def _cell(row, idx):
    if idx is None or idx >= len(row): return ""
    v = row[idx]
    return str(v).strip() if v is not None else ""

def _parse_prezzo(v) -> float:
    if not v: return 0.0
    try:
        return float(str(v).replace("€","").replace(" ","").replace(".","").replace(",","."))
    except: return 0.0

def leggi_clienti() -> list:
    if not CLIENTI_FILE.exists(): return []
    return json.loads(CLIENTI_FILE.read_text(encoding="utf-8"))

def scrivi_clienti(clienti: list):
    CLIENTI_FILE.write_text(json.dumps(clienti, ensure_ascii=False, indent=2), encoding="utf-8")

def leggi_impostazioni() -> dict:
    if not IMPOSTAZIONI_FILE.exists(): return dict(IMPOSTAZIONI_DEFAULT)
    try:
        d = json.loads(IMPOSTAZIONI_FILE.read_text(encoding="utf-8"))
        merged = dict(IMPOSTAZIONI_DEFAULT)
        merged.update(d)
        merged["azienda"] = {**IMPOSTAZIONI_DEFAULT["azienda"], **d.get("azienda", {})}
        return merged
    except: return dict(IMPOSTAZIONI_DEFAULT)

def scrivi_impostazioni(s: dict):
    IMPOSTAZIONI_FILE.write_text(json.dumps(s, ensure_ascii=False, indent=2), encoding="utf-8")

def genera_pdf(doc: dict, impostazioni: dict) -> bytes:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_RIGHT, TA_LEFT
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable

    buf = io.BytesIO()
    az = impostazioni.get("azienda", {})
    righe = doc.get("righe", [])
    cliente = doc.get("cliente", {})
    tipo = doc.get("tipo", "Preventivo").upper()
    numero = doc.get("numero", "")
    note = doc.get("note", "")

    totale = sum(r["quantita"] * r["prezzo_unitario"] for r in righe)

    def fmt(v): return f"€ {v:,.2f}".replace(",","X").replace(".",",").replace("X",".")
    def p(testo, fs=9, bold=False, color="#333333", align=TA_LEFT):
        return Paragraph(str(testo), ParagraphStyle("_",
            fontSize=fs, fontName="Helvetica-Bold" if bold else "Helvetica",
            textColor=colors.HexColor(color), leading=fs+4, alignment=align))

    pdf = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm, topMargin=2*cm, bottomMargin=2.5*cm)
    W = A4[0] - 4*cm
    el = []

    # Header azienda + tipo documento
    az_lines = [p(az.get("nome",""), 13, True, "#1a1a2e")]
    for k in ("indirizzo","partita_iva","telefono","email"):
        if az.get(k): az_lines.append(p(az[k], 8, color="#666"))
    doc_lines = [p(tipo, 18, True, "#1a1a2e"),
                 p(f"N. {numero}" if numero else "", 9, color="#555"),
                 p(f"Data: {date.today().strftime('%d/%m/%Y')}", 9, color="#555")]

    t = Table([[az_lines, doc_lines]], colWidths=[W*0.55, W*0.45])
    t.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),("ALIGN",(1,0),(1,0),"RIGHT")]))
    el += [t, Spacer(1,8), HRFlowable(width="100%",thickness=0.5,color=colors.HexColor("#ccc")), Spacer(1,10)]

    # Cliente
    if cliente.get("nome"):
        el.append(p("DESTINATARIO", 7.5, True, "#888"))
        el.append(p(cliente["nome"], 10, True))
        if cliente.get("indirizzo"): el.append(p(cliente["indirizzo"], 9))
        dettagli = [f"P.IVA: {cliente['partita_iva']}" if cliente.get("partita_iva") else "",
                    cliente.get("telefono",""), cliente.get("email","")]
        dettagli = [d for d in dettagli if d]
        if dettagli: el.append(p("  |  ".join(dettagli), 8, color="#777"))
        el.append(Spacer(1,12))

    # Tabella voci
    intestazioni = ["Codice","Descrizione","U.M.","Qnt.","Prezzo u.","Totale"]
    dati = [intestazioni]
    for r in righe:
        tot = r["quantita"] * r["prezzo_unitario"]
        dati.append([
            p(r.get("codice",""), 8),
            p(r.get("descrizione",""), 9),
            p(r.get("um",""), 8),
            p(f"{r['quantita']:g}", 9, align=TA_RIGHT),
            p(fmt(r["prezzo_unitario"]), 9, align=TA_RIGHT),
            p(fmt(tot), 9, True, align=TA_RIGHT),
        ])
    cw = [W*0.14, W*0.38, W*0.06, W*0.08, W*0.17, W*0.17]
    tv = Table(dati, colWidths=cw, repeatRows=1)
    stile = [
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1a1a2e")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,0),8.5),
        ("BOTTOMPADDING",(0,0),(-1,0),7),("TOPPADDING",(0,0),(-1,0),7),
        ("FONTSIZE",(0,1),(-1,-1),9),
        ("TOPPADDING",(0,1),(-1,-1),5),("BOTTOMPADDING",(0,1),(-1,-1),5),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("LINEBELOW",(0,0),(-1,-1),0.3,colors.HexColor("#e0e0e0")),
        ("ALIGN",(3,1),(-1,-1),"RIGHT"),("ALIGN",(3,0),(-1,0),"RIGHT"),
    ]
    for i in range(1, len(dati)):
        if i % 2 == 0: stile.append(("BACKGROUND",(0,i),(-1,i),colors.HexColor("#f8f8f8")))
    tv.setStyle(TableStyle(stile))
    el += [tv, Spacer(1,10)]

    # Totale finale (senza IVA)
    tt = Table([
        ["","","","", p("TOTALE",11,True), p(fmt(totale),11,True,align=TA_RIGHT)],
    ], colWidths=cw)
    tt.setStyle(TableStyle([
        ("ALIGN",(4,0),(-1,-1),"RIGHT"),
        ("BACKGROUND",(4,0),(-1,0),colors.HexColor("#f0f0f0")),
        ("LINEABOVE",(4,0),(-1,0),0.8,colors.HexColor("#555")),
        ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
    ]))
    el.append(tt)

    if note:
        el += [Spacer(1,14), HRFlowable(width="100%",thickness=0.3,color=colors.HexColor("#ddd")),
               Spacer(1,6), p("Note",7.5,True,"#888"), p(note,8.5,color="#555")]

    def footer(c, d):
        c.saveState()
        c.setFont("Helvetica",7.5); c.setFillColor(colors.HexColor("#aaa"))
        c.drawString(2*cm,1.2*cm, az.get("nome",""))
        c.drawRightString(A4[0]-2*cm,1.2*cm,f"Pagina {d.page}")
        c.restoreState()

    pdf.build(el, onFirstPage=footer, onLaterPages=footer)
    return buf.getvalue()

# ── Routes ────────────────────────────────────────────────────────

@app.route("/")
def index():
    return open(BASE / "index.html", encoding="utf-8").read()

@app.route("/api/prezziario")
def api_prezziario():
    q = request.args.get("q","").lower().strip()
    sorgente = request.args.get("sorgente","")
    voci = carica_prezziario()
    if q:
        tokens = q.split()
        voci = [v for v in voci
                if all(t in f"{v['codice']} {v['descrizione']}".lower() for t in tokens)
                and (not sorgente or v.get("sorgente") == sorgente)]
    elif sorgente:
        voci = [v for v in voci if v.get("sorgente") == sorgente]
    return jsonify(voci[:300])

@app.route("/api/prezziario/sorgenti")
def api_sorgenti():
    voci = carica_prezziario()
    sorgenti = sorted(set(v.get("sorgente","") for v in voci if v.get("sorgente")))
    return jsonify(sorgenti)

@app.route("/api/clienti", methods=["GET"])
def api_clienti_get():
    return jsonify(leggi_clienti())

@app.route("/api/clienti", methods=["POST"])
def api_clienti_post():
    clienti = leggi_clienti()
    clienti.append(request.json)
    scrivi_clienti(clienti)
    return jsonify({"ok": True, "index": len(clienti)-1})

@app.route("/api/clienti/<int:idx>", methods=["PUT"])
def api_clienti_put(idx):
    clienti = leggi_clienti()
    if 0 <= idx < len(clienti):
        clienti[idx] = request.json
        scrivi_clienti(clienti)
    return jsonify({"ok": True})

@app.route("/api/clienti/<int:idx>", methods=["DELETE"])
def api_clienti_delete(idx):
    clienti = leggi_clienti()
    if 0 <= idx < len(clienti):
        del clienti[idx]
        scrivi_clienti(clienti)
    return jsonify({"ok": True})

@app.route("/api/impostazioni", methods=["GET"])
def api_impostazioni_get():
    return jsonify(leggi_impostazioni())

@app.route("/api/impostazioni", methods=["POST"])
def api_impostazioni_post():
    scrivi_impostazioni(request.json)
    return jsonify({"ok": True})

@app.route("/api/pdf", methods=["POST"])
def api_pdf():
    doc = request.json
    impostazioni = leggi_impostazioni()
    try:
        pdf_bytes = genera_pdf(doc, impostazioni)
        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype="application/pdf",
            as_attachment=True,
            download_name=f"preventivo_{date.today().strftime('%Y%m%d')}.pdf"
        )
    except Exception as e:
        return jsonify({"errore": str(e)}), 500

@app.route("/api/reload", methods=["POST"])
def api_reload():
    global _voci_cache
    _voci_cache = []
    carica_prezziario()
    return jsonify({"ok": True, "voci": len(_voci_cache)})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"\n✓ Preventivatore avviato → http://localhost:{port}\n")
    app.run(host="0.0.0.0", port=port)
