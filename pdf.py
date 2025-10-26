from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, Spacer, SimpleDocTemplate
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from PyPDF2 import PdfMerger, PdfReader
from pathlib import Path
import subprocess, tempfile, json, shutil

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
                    "libreoffice", "--headless", "--convert-to", "pdf",
                    "--outdir", str(output_folder), str(docx_path)
                ],
                check=True,
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
            )
        except Exception as e:
            print(f"Failed to convert {docx_path}: {e}")
            return None
    return output_pdf

def txt_to_pdf(txt_path, output_folder):
    """Convert a plain text file into a PDF."""
    output_pdf = Path(output_folder) / (Path(txt_path).stem + ".pdf")
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(str(output_pdf), pagesize=A4)
    story = []

    # Read the text file
    with open(txt_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                story.append(Paragraph(line, styles["BodyText"]))
                story.append(Spacer(1, 2))  # small spacing

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
    """Converts all files to PDFs and collects metadata."""
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

        if output_pdf and Path(output_pdf).exists():
            reader = PdfReader(str(output_pdf))
            pdf_infos.append({
                "title": src_path.name,
                "path": Path(output_pdf),
                "pages": len(reader.pages)
            })
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

    doc = SimpleDocTemplate(str(output_path), pagesize=(595.27, 841.89))
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
