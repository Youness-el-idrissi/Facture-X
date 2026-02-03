#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Flask application to edit the embedded Factur‑X XML inside a PDF.

This app extracts the first XML attachment from an uploaded PDF, allows
the user to edit it in a browser, and then writes the modified XML back
into the PDF.  Existing XML attachments are removed before the new
attachment is injected.  A unique filename is generated for each
updated PDF to avoid caching issues.

Dependencies:
  - Flask
  - PyMuPDF (imported as ``fitz``)
  - lxml

Templates live in the ``templates`` directory; see ``index.html`` and
``edit.html`` for the UI.
"""

import os
import secrets
from datetime import datetime
from pathlib import Path

from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    send_file,
    flash,
    session,
    abort,
)
from werkzeug.utils import secure_filename
import fitz  # PyMuPDF
from lxml import etree


# Namespaces used in Factur-X XML
NAMESPACES = {
    'rsm': 'urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100',
    'ram': 'urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100',
    'qdt': 'urn:un:unece:uncefact:data:standard:QualifiedDataType:100',
    'udt': 'urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100',
}


def extract_invoice_data(xml_bytes: bytes) -> dict:
    """Extract key invoice data from Factur-X XML.

    Parameters
    ----------
    xml_bytes: bytes
        The XML content to parse.

    Returns
    -------
    dict
        Dictionary containing the extracted invoice data.
    """
    root = etree.fromstring(xml_bytes)
    data = {}

    # Helper function to get text from xpath
    def get_text(xpath: str, default: str = "") -> str:
        elems = root.xpath(xpath, namespaces=NAMESPACES)
        return elems[0].text if elems and elems[0].text else default

    # Invoice number and dates
    data['invoice_number'] = get_text('.//rsm:ExchangedDocument/ram:ID')
    data['invoice_date'] = get_text('.//rsm:ExchangedDocument/ram:IssueDateTime/udt:DateTimeString')
    data['due_date'] = get_text('.//ram:SpecifiedTradePaymentTerms/ram:DueDateDateTime/udt:DateTimeString')

    # Seller information
    data['seller_name'] = get_text('.//ram:ApplicableHeaderTradeAgreement/ram:SellerTradeParty/ram:Name')
    data['seller_siren'] = get_text('.//ram:ApplicableHeaderTradeAgreement/ram:SellerTradeParty/ram:SpecifiedLegalOrganization/ram:ID')
    data['seller_vat'] = get_text('.//ram:ApplicableHeaderTradeAgreement/ram:SellerTradeParty/ram:SpecifiedTaxRegistration/ram:ID')
    data['seller_street'] = get_text('.//ram:ApplicableHeaderTradeAgreement/ram:SellerTradeParty/ram:PostalTradeAddress/ram:LineOne')
    data['seller_city'] = get_text('.//ram:ApplicableHeaderTradeAgreement/ram:SellerTradeParty/ram:PostalTradeAddress/ram:CityName')
    data['seller_postal_code'] = get_text('.//ram:ApplicableHeaderTradeAgreement/ram:SellerTradeParty/ram:PostalTradeAddress/ram:PostcodeCode')
    data['seller_country'] = get_text('.//ram:ApplicableHeaderTradeAgreement/ram:SellerTradeParty/ram:PostalTradeAddress/ram:CountryID')
    # Seller electronic address (BT-34 Adresse FE)
    data['seller_fe_address'] = get_text('.//ram:ApplicableHeaderTradeAgreement/ram:SellerTradeParty/ram:URIUniversalCommunication/ram:URIID')

    # Buyer information
    data['buyer_name'] = get_text('.//ram:ApplicableHeaderTradeAgreement/ram:BuyerTradeParty/ram:Name')
    data['buyer_siren'] = get_text('.//ram:ApplicableHeaderTradeAgreement/ram:BuyerTradeParty/ram:SpecifiedLegalOrganization/ram:ID')
    data['buyer_vat'] = get_text('.//ram:ApplicableHeaderTradeAgreement/ram:BuyerTradeParty/ram:SpecifiedTaxRegistration/ram:ID')
    data['buyer_street'] = get_text('.//ram:ApplicableHeaderTradeAgreement/ram:BuyerTradeParty/ram:PostalTradeAddress/ram:LineOne')
    data['buyer_city'] = get_text('.//ram:ApplicableHeaderTradeAgreement/ram:BuyerTradeParty/ram:PostalTradeAddress/ram:CityName')
    data['buyer_postal_code'] = get_text('.//ram:ApplicableHeaderTradeAgreement/ram:BuyerTradeParty/ram:PostalTradeAddress/ram:PostcodeCode')
    data['buyer_country'] = get_text('.//ram:ApplicableHeaderTradeAgreement/ram:BuyerTradeParty/ram:PostalTradeAddress/ram:CountryID')
    # Buyer electronic address (BT-49 Adresse FE)
    data['buyer_fe_address'] = get_text('.//ram:ApplicableHeaderTradeAgreement/ram:BuyerTradeParty/ram:URIUniversalCommunication/ram:URIID')

    # Billing address (BT-0225 - ShipToTradeParty / BillToTradeParty)
    data['billing_name'] = get_text('.//ram:ApplicableHeaderTradeDelivery/ram:ShipToTradeParty/ram:Name')
    data['billing_street'] = get_text('.//ram:ApplicableHeaderTradeDelivery/ram:ShipToTradeParty/ram:PostalTradeAddress/ram:LineOne')
    data['billing_city'] = get_text('.//ram:ApplicableHeaderTradeDelivery/ram:ShipToTradeParty/ram:PostalTradeAddress/ram:CityName')
    data['billing_postal_code'] = get_text('.//ram:ApplicableHeaderTradeDelivery/ram:ShipToTradeParty/ram:PostalTradeAddress/ram:PostcodeCode')
    data['billing_country'] = get_text('.//ram:ApplicableHeaderTradeDelivery/ram:ShipToTradeParty/ram:PostalTradeAddress/ram:CountryID')

    return data


def update_invoice_xml(xml_bytes: bytes, data: dict) -> bytes:
    """Update Factur-X XML with new invoice data.

    Parameters
    ----------
    xml_bytes: bytes
        The original XML content.
    data: dict
        Dictionary containing the new invoice data.

    Returns
    -------
    bytes
        The updated XML content.
    """
    root = etree.fromstring(xml_bytes)

    # Helper function to set text in xpath (creates element if needed)
    def set_text(xpath: str, value: str) -> None:
        if not value:
            return
        elems = root.xpath(xpath, namespaces=NAMESPACES)
        if elems:
            elems[0].text = value

    # Invoice number and dates
    set_text('.//rsm:ExchangedDocument/ram:ID', data.get('invoice_number', ''))
    set_text('.//rsm:ExchangedDocument/ram:IssueDateTime/udt:DateTimeString', data.get('invoice_date', ''))
    set_text('.//ram:SpecifiedTradePaymentTerms/ram:DueDateDateTime/udt:DateTimeString', data.get('due_date', ''))

    # Seller information
    set_text('.//ram:ApplicableHeaderTradeAgreement/ram:SellerTradeParty/ram:Name', data.get('seller_name', ''))
    set_text('.//ram:ApplicableHeaderTradeAgreement/ram:SellerTradeParty/ram:SpecifiedLegalOrganization/ram:ID', data.get('seller_siren', ''))
    set_text('.//ram:ApplicableHeaderTradeAgreement/ram:SellerTradeParty/ram:SpecifiedTaxRegistration/ram:ID', data.get('seller_vat', ''))
    set_text('.//ram:ApplicableHeaderTradeAgreement/ram:SellerTradeParty/ram:PostalTradeAddress/ram:LineOne', data.get('seller_street', ''))
    set_text('.//ram:ApplicableHeaderTradeAgreement/ram:SellerTradeParty/ram:PostalTradeAddress/ram:CityName', data.get('seller_city', ''))
    set_text('.//ram:ApplicableHeaderTradeAgreement/ram:SellerTradeParty/ram:PostalTradeAddress/ram:PostcodeCode', data.get('seller_postal_code', ''))
    set_text('.//ram:ApplicableHeaderTradeAgreement/ram:SellerTradeParty/ram:PostalTradeAddress/ram:CountryID', data.get('seller_country', ''))
    # Seller electronic address (BT-34 Adresse FE)
    set_text('.//ram:ApplicableHeaderTradeAgreement/ram:SellerTradeParty/ram:URIUniversalCommunication/ram:URIID', data.get('seller_fe_address', ''))

    # Buyer information
    set_text('.//ram:ApplicableHeaderTradeAgreement/ram:BuyerTradeParty/ram:Name', data.get('buyer_name', ''))
    set_text('.//ram:ApplicableHeaderTradeAgreement/ram:BuyerTradeParty/ram:SpecifiedLegalOrganization/ram:ID', data.get('buyer_siren', ''))
    set_text('.//ram:ApplicableHeaderTradeAgreement/ram:BuyerTradeParty/ram:SpecifiedTaxRegistration/ram:ID', data.get('buyer_vat', ''))
    set_text('.//ram:ApplicableHeaderTradeAgreement/ram:BuyerTradeParty/ram:PostalTradeAddress/ram:LineOne', data.get('buyer_street', ''))
    set_text('.//ram:ApplicableHeaderTradeAgreement/ram:BuyerTradeParty/ram:PostalTradeAddress/ram:CityName', data.get('buyer_city', ''))
    set_text('.//ram:ApplicableHeaderTradeAgreement/ram:BuyerTradeParty/ram:PostalTradeAddress/ram:PostcodeCode', data.get('buyer_postal_code', ''))
    set_text('.//ram:ApplicableHeaderTradeAgreement/ram:BuyerTradeParty/ram:PostalTradeAddress/ram:CountryID', data.get('buyer_country', ''))
    # Buyer electronic address (BT-49 Adresse FE)
    set_text('.//ram:ApplicableHeaderTradeAgreement/ram:BuyerTradeParty/ram:URIUniversalCommunication/ram:URIID', data.get('buyer_fe_address', ''))

    # Billing address (BT-0225 - ShipToTradeParty)
    set_text('.//ram:ApplicableHeaderTradeDelivery/ram:ShipToTradeParty/ram:Name', data.get('billing_name', ''))
    set_text('.//ram:ApplicableHeaderTradeDelivery/ram:ShipToTradeParty/ram:PostalTradeAddress/ram:LineOne', data.get('billing_street', ''))
    set_text('.//ram:ApplicableHeaderTradeDelivery/ram:ShipToTradeParty/ram:PostalTradeAddress/ram:CityName', data.get('billing_city', ''))
    set_text('.//ram:ApplicableHeaderTradeDelivery/ram:ShipToTradeParty/ram:PostalTradeAddress/ram:PostcodeCode', data.get('billing_postal_code', ''))
    set_text('.//ram:ApplicableHeaderTradeDelivery/ram:ShipToTradeParty/ram:PostalTradeAddress/ram:CountryID', data.get('billing_country', ''))

    return etree.tostring(root, xml_declaration=True, encoding='UTF-8')


# Base directory of this file and workspace for uploads/temporary data
BASE_DIR = Path(__file__).resolve().parent
WORK_DIR = BASE_DIR / "workspace"
WORK_DIR.mkdir(parents=True, exist_ok=True)

# Allowed upload extensions
ALLOWED_EXTS = {".pdf"}

# Flask app configuration
app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", secrets.token_hex(16))


def ensure_job() -> Path:
    """Ensure there is a unique working directory for the current session.

    Each browser session is given a unique job_id stored in the Flask
    session.  Files related to the upload and XML editing are stored
    under ``workspace/<job_id>``.  This isolates concurrent users and
    prevents collisions.

    Returns the Path object for the job directory.
    """
    if "job_id" not in session:
        session["job_id"] = secrets.token_hex(8)
    job_dir = WORK_DIR / session["job_id"]
    job_dir.mkdir(parents=True, exist_ok=True)
    return job_dir


def extract_first_xml(pdf_path: Path) -> tuple[str, bytes]:
    """Extract the first XML attachment from the given PDF.

    Preference is given to attachments whose names contain "factur" or
    "zugferd"; if none match those terms, the first ``.xml`` attachment
    encountered is used.

    Parameters
    ----------
    pdf_path: Path
        Path to the PDF file to inspect.

    Returns
    -------
    (xml_name, xml_bytes): tuple of str and bytes
        The filename of the XML attachment and its binary contents.

    Raises
    ------
    ValueError
        If no attachments or no XML attachments are found in the PDF.
    """
    doc = fitz.open(str(pdf_path))
    names = list(doc.embfile_names())
    if not names:
        doc.close()
        raise ValueError("Aucune pièce jointe trouvée dans le PDF.")
    xml_name: str | None = None
    # Prefer names containing 'factur' or 'zugferd'
    for n in names:
        if n.lower().endswith(".xml"):
            xml_name = n
            if "factur" in n.lower() or "zugferd" in n.lower():
                break
    if not xml_name:
        doc.close()
        raise ValueError("Aucun attachement XML trouvé dans le PDF.")
    idx = names.index(xml_name)
    xml_bytes = doc.embfile_get(idx)
    doc.close()
    return xml_name, xml_bytes


def inject_xml(pdf_path: Path, xml_bytes: bytes, xml_name: str, out_path: Path) -> None:
    """Inject an XML attachment into a PDF, replacing existing XML attachments.

    Existing ``.xml`` attachments are removed before adding the new
    attachment.  The name of the new attachment matches the original
    attachment name when available.  The resulting PDF is written to
    ``out_path``.

    Parameters
    ----------
    pdf_path: Path
        The source PDF file.
    xml_bytes: bytes
        The modified XML content to embed.
    xml_name: str
        The filename to use for the attachment.  If ``xml_name`` is
        falsy, ``factur-x.xml`` will be used.
    out_path: Path
        Destination path for the updated PDF.
    """
    doc = fitz.open(str(pdf_path))
    # Remove all existing attachments.  Some PDF generators embed multiple
    # XML files or other artefacts; to ensure only the corrected XML is
    # retained we delete every attachment.  Removing from the end
    # preserves indices during deletion.
    try:
        for i in reversed(range(doc.embfile_count())):
            doc.embfile_del(i)
    except Exception:
        # Ignore any errors encountered during deletion (e.g. if there
        # are no attachments).
        pass
    target_name = xml_name or "factur-x.xml"
    # In PyMuPDF 1.26+, data goes as second positional argument
    doc.embfile_add(target_name, xml_bytes, ufilename=target_name, desc="Factur-X embedded XML")
    # Save the updated PDF; garbage=3 eliminates unused objects, deflate compresses streams
    doc.save(str(out_path), garbage=3, deflate=True)
    doc.close()


@app.route("/")
def index():
    """Landing page with PDF upload form."""
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    """Handle PDF upload and extract the first XML attachment."""
    job_dir = ensure_job()
    f = request.files.get("pdf")
    if not f or f.filename == "":
        flash("Veuillez sélectionner un PDF.", "danger")
        return redirect(url_for("index"))
    name = secure_filename(f.filename)
    if Path(name).suffix.lower() not in ALLOWED_EXTS:
        flash("Seuls les fichiers .pdf sont acceptés.", "danger")
        return redirect(url_for("index"))
    pdf_path = job_dir / "original.pdf"
    f.save(str(pdf_path))
    try:
        xml_name, xml_bytes = extract_first_xml(pdf_path)
    except Exception as e:
        flash(str(e), "danger")
        return redirect(url_for("index"))
    # Persist original attachment name and extracted XML to the job directory
    (job_dir / "xml_name.txt").write_text(xml_name, encoding="utf-8")
    (job_dir / "factur-x.xml").write_bytes(xml_bytes)
    flash("XML extrait avec succès. Vous pouvez modifier les données ci-dessous.", "success")
    return redirect(url_for("edit_form"))


@app.route("/edit")
def edit():
    """Show a text area to edit the extracted XML."""
    job_dir = ensure_job()
    xml_file = job_dir / "factur-x.xml"
    if not xml_file.exists():
        flash("Aucun XML chargé pour ce job.", "danger")
        return redirect(url_for("index"))
    xml_text = xml_file.read_text(encoding="utf-8", errors="ignore")
    xml_name_path = job_dir / "xml_name.txt"
    xml_name = xml_name_path.read_text(encoding="utf-8") if xml_name_path.exists() else "factur-x.xml"
    return render_template("edit.html", xml_text=xml_text, xml_name=xml_name)


@app.route("/download/xml")
def download_xml():
    """Download the current XML file."""
    job_dir = ensure_job()
    xml_file = job_dir / "factur-x.xml"
    if not xml_file.exists():
        abort(404)
    return send_file(str(xml_file), as_attachment=True, download_name="factur-x.xml")


@app.route("/save", methods=["POST"])
def save_xml():
    """Save the edited XML back to the job directory, validating its well-formedness."""
    job_dir = ensure_job()
    xml_file = job_dir / "factur-x.xml"
    content = request.form.get("xml_text", "")
    try:
        etree.fromstring(content.encode("utf-8"))
    except Exception as e:
        flash(f"XML invalide : {e}", "danger")
        return redirect(url_for("edit"))
    xml_file.write_text(content, encoding="utf-8")
    flash("XML enregistré.", "success")
    return redirect(url_for("edit"))


@app.route("/edit_form")
def edit_form():
    """Show a form to edit key invoice fields (SIREN, addresses, etc.)."""
    job_dir = ensure_job()
    xml_file = job_dir / "factur-x.xml"
    if not xml_file.exists():
        flash("Aucun XML chargé pour ce job.", "danger")
        return redirect(url_for("index"))
    xml_bytes = xml_file.read_bytes()
    try:
        data = extract_invoice_data(xml_bytes)
    except Exception as e:
        flash(f"Erreur lors de l'extraction des données : {e}", "danger")
        return redirect(url_for("edit"))
    xml_name_path = job_dir / "xml_name.txt"
    xml_name = xml_name_path.read_text(encoding="utf-8") if xml_name_path.exists() else "factur-x.xml"
    return render_template("edit_form.html", data=data, xml_name=xml_name)


@app.route("/save_form", methods=["POST"])
def save_form():
    """Save the form data back to the XML file."""
    job_dir = ensure_job()
    xml_file = job_dir / "factur-x.xml"
    if not xml_file.exists():
        flash("Aucun XML chargé pour ce job.", "danger")
        return redirect(url_for("index"))
    # Collect form data
    form_data = {
        'invoice_number': request.form.get('invoice_number', ''),
        'invoice_date': request.form.get('invoice_date', ''),
        'due_date': request.form.get('due_date', ''),
        'seller_name': request.form.get('seller_name', ''),
        'seller_siren': request.form.get('seller_siren', ''),
        'seller_vat': request.form.get('seller_vat', ''),
        'seller_street': request.form.get('seller_street', ''),
        'seller_city': request.form.get('seller_city', ''),
        'seller_postal_code': request.form.get('seller_postal_code', ''),
        'seller_country': request.form.get('seller_country', ''),
        'seller_fe_address': request.form.get('seller_fe_address', ''),
        'buyer_name': request.form.get('buyer_name', ''),
        'buyer_siren': request.form.get('buyer_siren', ''),
        'buyer_vat': request.form.get('buyer_vat', ''),
        'buyer_street': request.form.get('buyer_street', ''),
        'buyer_city': request.form.get('buyer_city', ''),
        'buyer_postal_code': request.form.get('buyer_postal_code', ''),
        'buyer_country': request.form.get('buyer_country', ''),
        'buyer_fe_address': request.form.get('buyer_fe_address', ''),
        # Billing address (BT-0225)
        'billing_name': request.form.get('billing_name', ''),
        'billing_street': request.form.get('billing_street', ''),
        'billing_city': request.form.get('billing_city', ''),
        'billing_postal_code': request.form.get('billing_postal_code', ''),
        'billing_country': request.form.get('billing_country', ''),
    }
    # Update the XML
    xml_bytes = xml_file.read_bytes()
    try:
        updated_xml = update_invoice_xml(xml_bytes, form_data)
        xml_file.write_bytes(updated_xml)
        flash("Données de la facture enregistrées.", "success")
    except Exception as e:
        flash(f"Erreur lors de la mise à jour du XML : {e}", "danger")
    return redirect(url_for("edit_form"))


@app.route("/build", methods=["POST"])
def build():
    """Build a new PDF with the modified XML and offer it for download."""
    job_dir = ensure_job()
    pdf_path = job_dir / "original.pdf"
    xml_file = job_dir / "factur-x.xml"
    if not (pdf_path.exists() and xml_file.exists()):
        flash("Fichiers manquants.", "danger")
        return redirect(url_for("index"))
    # Use the original attachment name if available
    xml_name_path = job_dir / "xml_name.txt"
    xml_name = xml_name_path.read_text(encoding="utf-8") if xml_name_path.exists() else "factur-x.xml"
    # Create a unique filename to avoid browser caching
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = job_dir / f"updated_{ts}.pdf"
    inject_xml(pdf_path, xml_file.read_bytes(), xml_name, out_path)
    # Extract the resulting XML for debugging/verification purposes
    try:
        doc = fitz.open(str(out_path))
        for i, n in enumerate(doc.embfile_names()):
            if n.lower().endswith(".xml"):
                xml_bytes = doc.embfile_get(i)
                (job_dir / "verify.xml").write_bytes(xml_bytes)
                break
        doc.close()
    except Exception:
        pass
    download_name = f"invoice.updated.{ts}.pdf"
    # ``max_age=0`` prevents caching the file in the browser (Flask 3+)
    return send_file(str(out_path), as_attachment=True, download_name=download_name, max_age=0)


if __name__ == "__main__":
    # Run the Flask development server
    app.run(host="127.0.0.1", port=5000, debug=True)