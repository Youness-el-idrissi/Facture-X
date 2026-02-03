#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Factur-X PDF XML Editor — version deux interfaces (Upload + Update)
===================================================================
Dépendances :
  pip install flask pymupdf lxml
"""

import os, secrets
from datetime import datetime
from pathlib import Path
from flask import Flask, request, redirect, url_for, send_file, flash, session, render_template_string
from werkzeug.utils import secure_filename
import fitz  # PyMuPDF
from lxml import etree

# ===============================
# CONFIG
# ===============================
BASE_DIR = Path(__file__).resolve().parent
WORK_DIR = BASE_DIR / "workspace"
WORK_DIR.mkdir(parents=True, exist_ok=True)
ALLOWED_EXTS = {".pdf"}

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", secrets.token_hex(16))


# ===============================
# HELPERS
# ===============================
def ensure_job() -> Path:
    """Crée un espace de travail unique par session navigateur"""
    if "job_id" not in session:
        session["job_id"] = secrets.token_hex(8)
    job = WORK_DIR / session["job_id"]
    job.mkdir(exist_ok=True)
    return job


def find_facturx_xml(doc):
    xmls = [(i, n, doc.embfile_get(i)) for i, n in enumerate(doc.embfile_names()) if n.lower().endswith(".xml")]
    if not xmls:
        return None
    for i, n, b in xmls:
        if "factur" in n.lower() or "zugferd" in n.lower():
            return i, n, b
    return xmls[0]


def clean_xml(raw: bytes) -> bytes:
    """Supprime BOM + espaces avant <?xml>"""
    txt = raw.decode("utf-8", errors="ignore").lstrip("\ufeff\r\n\t ")
    return txt.encode("utf-8")


def validate_xml(data: bytes):
    etree.fromstring(data)  # lève exception si invalide


def inject_facturx_xml(pdf_in: Path, xml_bytes: bytes, xml_name: str | None, out_path: Path):
    """Remplace le XML Factur-X / ZUGFeRD dans le PDF"""
    doc = fitz.open(str(pdf_in))
    to_del = [i for i, n in enumerate(doc.embfile_names())
              if n.lower().endswith(".xml") and ("factur" in n.lower() or "zugferd" in n.lower() or "invoice" in n.lower())]
    for i in reversed(to_del):
        doc.embfile_del(i)
    target = (xml_name or "factur-x.xml").strip() or "factur-x.xml"
    doc.embfile_add(target, xml_bytes, ufilename=target, desc="Updated Factur-X XML")
    doc.save(str(out_path), garbage=3, deflate=True)
    doc.close()


# ===============================
# HTML TEMPLATES INLINE
# ===============================

layout = """
<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8">
<title>{{ title }}</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
<style>
body{background:#0f172a;color:#e5e7eb}
.container-box{
 max-width:950px;margin:40px auto;padding:24px;
 background:rgba(15,23,42,.97);
 border-radius:22px;border:1px solid rgba(129,140,248,.35);
 box-shadow:0 18px 40px rgba(15,23,42,.9);
}
textarea{font-family:monospace;font-size:.85rem}
</style>
</head>
<body>
<div class="container-box">
  {% with messages = get_flashed_messages(with_categories=true) %}
  {% for cat,msg in messages %}
  <div class="alert alert-{{'warning' if cat=='message' else cat}} py-2 mb-2">{{msg|safe}}</div>
  {% endfor %}
  {% endwith %}
  {{ body|safe }}
</div>
</body></html>
"""


def render_page(title, body_html):
    return render_template_string(layout, title=title, body=body_html)


# ===============================
# ROUTES
# ===============================

@app.route("/")
def index():
    """Interface 1 — Upload"""
    body = f"""
    <h4>🧾 Importer un PDF Factur-X</h4>
    <p class="text-secondary small mb-3">
      Cette étape extrait automatiquement le XML intégré pour modification.
    </p>
    <form action="{url_for('upload')}" method="post" enctype="multipart/form-data">
      <div class="mb-3">
        <input type="file" name="pdf" accept="application/pdf" class="form-control" required>
      </div>
      <button class="btn btn-primary">Extraire le XML</button>
    </form>
    """
    return render_page("Upload PDF", body)


@app.route("/upload", methods=["POST"])
def upload():
    job = ensure_job()
    f = request.files.get("pdf")
    if not f or f.filename == "":
        flash("Aucun fichier sélectionné.", "danger")
        return redirect(url_for("index"))

    name = secure_filename(f.filename)
    if Path(name).suffix.lower() not in ALLOWED_EXTS:
        flash("Seuls les fichiers .pdf sont acceptés.", "danger")
        return redirect(url_for("index"))

    pdf_path = job / "original.pdf"
    f.save(pdf_path)
    doc = fitz.open(str(pdf_path))
    fx = find_facturx_xml(doc)
    if not fx:
        flash("Aucun XML Factur-X / ZUGFeRD détecté.", "danger")
        return redirect(url_for("index"))

    _, xml_name, xml_raw = fx
    xml_clean = clean_xml(xml_raw)
    try:
        validate_xml(xml_clean)
        flash(f"XML « {xml_name} » extrait et valide ✅", "success")
    except Exception as e:
        flash(f"XML extrait mais invalide : {e}", "warning")

    (job / "factur-x.xml").write_bytes(xml_clean)
    (job / "xml_name.txt").write_text(xml_name, encoding="utf-8")
    doc.close()
    return redirect(url_for("edit"))


@app.route("/edit")
def edit():
    """Interface 2 — Édition / Update"""
    job = ensure_job()
    xml_file = job / "factur-x.xml"
    if not xml_file.exists():
        flash("Aucun XML trouvé pour cette session.", "danger")
        return redirect(url_for("index"))

    xml_text = xml_file.read_text(encoding="utf-8", errors="ignore")
    xml_name = (job / "xml_name.txt").read_text(encoding="utf-8") if (job / "xml_name.txt").exists() else "factur-x.xml"

    body = f"""
    <h4>✏️ Éditer le XML intégré</h4>
    <p class="text-secondary small">Nom de la pièce jointe : <b>{xml_name}</b></p>
    <form action="{url_for('save_xml')}" method="post">
      <textarea name="xml_text" rows="22" class="form-control mb-2">{xml_text}</textarea>
      <div class="d-flex gap-2">
        <button class="btn btn-success btn-sm">💾 Enregistrer</button>
        <a href="{url_for('download_xml')}" class="btn btn-outline-light btn-sm">Télécharger XML</a>
      </div>
    </form>
    <form action="{url_for('build')}" method="post" class="mt-3">
      <button class="btn btn-primary btn-sm">📦 Générer le PDF mis à jour</button>
      <p class="text-secondary small mt-1">
        Le fichier PDF original sera copié et le XML remplacé.
      </p>
    </form>
    <div class="mt-3">
      <a href="{url_for('index')}" class="btn btn-outline-secondary btn-sm">← Revenir à l’import</a>
    </div>
    """
    return render_page("Édition XML", body)


@app.route("/save", methods=["POST"])
def save_xml():
    job = ensure_job()
    xml_file = job / "factur-x.xml"
    content = request.form.get("xml_text", "").lstrip("\ufeff\r\n\t ")
    try:
        validate_xml(content.encode("utf-8"))
    except Exception as e:
        flash(f"XML invalide : {e}", "danger")
        return redirect(url_for("edit"))
    xml_file.write_text(content, encoding="utf-8")
    flash("XML enregistré et valide ✅", "success")
    return redirect(url_for("edit"))


@app.route("/download/xml")
def download_xml():
    job = ensure_job()
    xml_file = job / "factur-x.xml"
    if not xml_file.exists():
        flash("Aucun XML à télécharger.", "danger")
        return redirect(url_for("index"))
    return send_file(str(xml_file), as_attachment=True, download_name="factur-x.xml")


@app.route("/build", methods=["POST"])
def build():
    job = ensure_job()
    pdf_in, xml_file = job / "original.pdf", job / "factur-x.xml"
    if not pdf_in.exists() or not xml_file.exists():
        flash("Fichiers manquants.", "danger")
        return redirect(url_for("index"))
    xml_name = (job / "xml_name.txt").read_text(encoding="utf-8") if (job / "xml_name.txt").exists() else "factur-x.xml"
    xml_bytes = xml_file.read_bytes()
    try:
        validate_xml(xml_bytes)
    except Exception as e:
        flash(f"XML invalide : {e}", "danger")
        return redirect(url_for("edit"))

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = job / f"updated_{ts}.pdf"
    inject_facturx_xml(pdf_in, xml_bytes, xml_name, out_path)
    flash("✅ Nouveau PDF généré avec succès !", "success")
    return redirect(url_for("result", fname=out_path.name))


@app.route("/result/<fname>")
def result(fname):
    """Page de résultat après génération"""
    job = ensure_job()
    path = job / fname
    if not path.exists():
        flash("Fichier introuvable.", "danger")
        return redirect(url_for("index"))
    body = f"""
    <h4>✅ PDF mis à jour</h4>
    <p>Le XML Factur-X a été remplacé avec succès.</p>
    <a href="{url_for('download_pdf', fname=fname)}" class="btn btn-primary btn-sm">Télécharger le nouveau PDF</a>
    <div class="mt-3">
      <a href="{url_for('index')}" class="btn btn-outline-light btn-sm">Importer un autre PDF</a>
    </div>
    """
    return render_page("Résultat", body)


@app.route("/download/pdf/<fname>")
def download_pdf(fname):
    job = ensure_job()
    path = job / fname
    return send_file(str(path), as_attachment=True, download_name=fname, max_age=0)


# ===============================
# RUN
# ===============================
if __name__ == "__main__":
    print("➡️  Ouvre http://127.0.0.1:5000 dans ton navigateur.")
    app.run(host="127.0.0.1", port=5000, debug=True)
