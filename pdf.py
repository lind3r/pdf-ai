from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, Spacer, SimpleDocTemplate
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from pathlib import Path
from io import BytesIO
import subprocess, tempfile, json, shutil


# === Helper: Inject title into any PDF ===
def inject_pdf_title(pdf_path, title):
    """Add a title (filename) to the first page of an existing PDF."""
    try:
        reader = PdfReader(str(pdf_path))
        writer = PdfWriter()

        # Create overlay with title
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)
        can.setFont("Helvetica-Bold", 14)
        can.drawCentredString(A4[0] / 2, A4[1] - 50, title)
        can.save()
        packet.seek(0)
        overlay_pdf = PdfReader(packet)

        # Merge overlay with first page
        first_page = reader.pages[0]
        first_page.merge_page(overlay_pdf.pages[0])
        writer.add_page(first_page)

        # Copy remaining pages
        for page in reader.pages[1:]:
            writer.add_page(page)

        # Write new titled PDF
        titled_path = Path(pdf_path)
        with open(titled_path, "wb") as f_out:
            writer.write(f_out)
        return titled_path
    except Exception as e:
        print(f"⚠️ Failed to inject title for {pdf_path}: {e}")
        return pdf_path


# === PDF Converters ===
def convert_docx_to_pdf(docx_path, output_folder):
    """Converts DOCX to PDF using docx2pdf or LibreOffice fallback."""
    output_pdf = Path(output_folder) / (Path(docx_path).stem + ".pdf")
    try:
        from docx2pdf import convert
        convert(input_path=str(docx_path), output_path=str(output_pdf))
    except Exception:
        try:
            subprocess.run(
                [
                    "libreoffice",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(output_folder),
                    str(docx_path),
                ],
                check=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        except Exception as e:
            print(f"Failed to convert {docx_path}: {e}")
            return None
    return output_pdf


def txt_to_pdf(txt_path, output_folder):
    """Convert a plain text file into a PDF and include title directly."""
    output_pdf = Path(output_folder) / (Path(txt_path).stem + ".pdf")
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(str(output_pdf), pagesize=A4)
    story = []

    # Inject title (filename)
    title = Path(txt_path).name
    story.append(Paragraph(f"<b>{title}</b>", styles["Title"]))
    story.append(Spacer(1, 0.3 * inch))

    # Read text file
    with open(txt_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                story.append(Paragraph(line, styles["BodyText"]))
                story.append(Spacer(1, 2))

    doc.build(story)
    return output_pdf


def image_to_pdf(image_path, output_path, title=None):
    """Converts an image to a PDF page."""
    c = canvas.Canvas(str(output_path), pagesize=A4)
    width, height = A4
    title_height = 50

    if title:
        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(width / 2, height - title_height, title)

    img = ImageReader(image_path)
    iw, ih = img.getSize()
    max_width, max_height = width - 100, height - title_height - 100
    scale = min(max_width / iw, max_height / ih, 1)

    img_w, img_h = iw * scale, ih * scale
    x, y = (width - img_w) / 2, (height - img_h) / 2 - 20

    c.drawImage(img, x, y, width=img_w, height=img_h, preserveAspectRatio=True)
    c.showPage()
    c.save()
    return output_path


def copy_pdf(pdf_path, output_folder):
    """Copies an existing PDF into the converted folder."""
    dest = Path(output_folder) / Path(pdf_path).name
    shutil.copy2(pdf_path, dest)
    return dest


# === Conversion Dispatcher ===
PDFConverterRegistry = {
    ".docx": convert_docx_to_pdf,
    ".pdf": copy_pdf,
    ".txt": txt_to_pdf,
    ".jpg": image_to_pdf,
    ".jpeg": image_to_pdf,
    ".png": image_to_pdf,
    ".tif": image_to_pdf,
    ".tiff": image_to_pdf,
}


def convert_all_to_pdfs(source_folder, ordered_files):
    """Converts all files to PDFs, injects titles, and collects metadata."""
    converted_dir = Path(source_folder) / "converted"
    converted_dir.mkdir(exist_ok=True)
    pdf_infos = []

    for entry in ordered_files:
        filename = entry.get("filnamn")
        src_path = Path(source_folder) / filename
        if not src_path.exists():
            print(f"Skipping missing file: {src_path}")
            continue

        ext = src_path.suffix.lower()
        converter = PDFConverterRegistry.get(ext)
        if not converter:
            print(f"Unsupported file type: {filename}")
            continue

        output_pdf = None
        if converter == image_to_pdf:
            output_pdf = converted_dir / f"{src_path.stem}.pdf"
            output_pdf = image_to_pdf(src_path, output_pdf, title=src_path.name)
        else:
            output_pdf = converter(src_path, converted_dir)

        # Inject title (filename) into first page — but skip .txt since it already has one
        if output_pdf and Path(output_pdf).exists():
            if ext != ".txt":
                output_pdf = inject_pdf_title(output_pdf, src_path.name)

            reader = PdfReader(str(output_pdf))
            pdf_infos.append(
                {
                    "title": src_path.name,
                    "path": Path(output_pdf),
                    "pages": len(reader.pages),
                }
            )
            print(f"Converted {filename} → {output_pdf} ({len(reader.pages)} pages)")

    return pdf_infos


# === TOC and Summary Creation ===
def build_toc_pdf(pdf_infos, output_path):
    """Creates a TOC PDF with filenames including extensions and summary entry."""
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.units import inch

    styles = getSampleStyleSheet()
    title_style = styles["Title"]
    entry_style = ParagraphStyle(
        "TOCEntry", parent=styles["Normal"], fontSize=11, spaceAfter=4
    )

    doc = SimpleDocTemplate(str(output_path), pagesize=A4)
    story = [Paragraph("TOC", title_style), Spacer(1, 0.2 * inch)]

    page_counter = 2  # TOC starts at page 1
    for info in pdf_infos:
        text = info["title"]
        dots = "." * (80 - len(text))
        entry = f"{text} {dots} {page_counter}"
        story.append(Paragraph(entry, entry_style))
        story.append(Spacer(1, 0.05 * inch))
        page_counter += info["pages"]

    # Add summary entry
    dots = "." * (80 - len("Sammanfattning"))
    entry = f"Sammanfattning {dots} {page_counter}"
    story.append(Paragraph(entry, entry_style))
    story.append(Spacer(1, 0.05 * inch))

    doc.build(story)
    return output_path


def create_summary_pdf(summary_text, output_path):
    """Creates a summary PDF."""
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(str(output_path), pagesize=A4)
    story = [
        Paragraph("Sammanfattning av hela materialet", styles["Heading1"]),
        Spacer(1, 0.2 * inch),
        Paragraph(summary_text, styles["BodyText"]),
    ]
    doc.build(story)
    return output_path


# === Merging ===
def merge_pdfs_with_toc(toc_pdf, pdf_infos, summary_pdf, output_pdf_path):
    merger = PdfMerger()
    merger.append(str(toc_pdf))
    for info in pdf_infos:
        merger.append(str(info["path"]))
    merger.append(str(summary_pdf))
    merger.write(str(output_pdf_path))
    merger.close()
    print(f"Final PDF created at: {output_pdf_path}")


# === Orchestration ===
def generate_pdf_report(summary_json_path, output_pdf_path, source_folder):
    with open(summary_json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    ordered_files = data.get("ordered_files", [])
    overall_summary = data.get("overall_summary", "Ingen sammanfattning tillgänglig.")
    source_folder = Path(source_folder)
    temp_dir = Path(tempfile.gettempdir())

    pdf_infos = convert_all_to_pdfs(source_folder, ordered_files)

    toc_pdf = temp_dir / "toc.pdf"
    summary_pdf = temp_dir / "summary.pdf"

    build_toc_pdf(pdf_infos, toc_pdf)
    create_summary_pdf(overall_summary, summary_pdf)
    merge_pdfs_with_toc(toc_pdf, pdf_infos, summary_pdf, output_pdf_path)
