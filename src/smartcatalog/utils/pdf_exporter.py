from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import io

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER


def export_product_blocks_to_pdf(blocks, output_path):
    """
    blocks: list of dicts with keys: page, image_index, image_bytes, texts (list of strings)
    """
    c = canvas.Canvas(output_path, pagesize=A4)
    page_w, page_h = A4

    max_image_width = 250
    max_image_height = 300
    top_margin = 100

    for block in blocks:
        # Centered image at top with preserved aspect ratio
        try:
            img_data = io.BytesIO(block["image_bytes"])
            img = ImageReader(img_data)
            img_width, img_height = img.getSize()

            scale = min(max_image_width / img_width, max_image_height / img_height, 1.0)
            display_width = img_width * scale
            display_height = img_height * scale

            x = (page_w - display_width) / 2
            y = page_h - top_margin - display_height

            c.drawImage(img, x, y, width=display_width, height=display_height, preserveAspectRatio=True)
        except:
            c.setFont("Helvetica", 10)
            c.drawString(100, page_h - top_margin, "[Image Error]")

        # Add metadata: Page
        y_text = y - 30
        c.setFont("Helvetica-Bold", 11)
        c.drawString(80, y_text, f"Page: {block['page']}")

        # Add Catalog Code
        y_text -= 20
        catalog_text = ", ".join(block.get("codes", [])) or "N/A"
        c.drawString(80, y_text, f"Catalog code: {catalog_text}")

        # Add product group (if available)
        y_text -= 20
        product_group = block.get("product_group", "N/A")
        c.drawString(80, y_text, f"Product Group: {product_group}")

        # Add Brand (if available)
        y_text -= 20
        brand_text = block.get("brand", "N/A")
        c.drawString(80, y_text, f"Brand: {brand_text}")

        # Add Description Title
        y_text -= 20
        c.drawString(80, y_text, "Description:")
        y_text -= 20

        # Add associated text lines as description
        c.setFont("Helvetica", 11)
        for line in block["texts"]:
            if y_text < 50:
                c.showPage()
                y_text = page_h - top_margin
                c.setFont("Helvetica", 11)
            c.drawString(100, y_text, line.strip())  # Indent for cleaner layout
            y_text -= 18
            
        # Word Item Info
        if "item" in block and isinstance(block["item"], dict):
            y_text -= 30
            c.setFont("Helvetica-Bold", 11)
            c.drawString(80, y_text, "Word Item Description:")
            c.setFont("Helvetica", 11)
            y_text -= 20

            for key, value in block["item"].items():
                if value and str(value).strip():
                    if y_text < 50:
                        c.showPage()
                        y_text = page_h - top_margin
                        c.setFont("Helvetica", 11)
                    c.drawString(100, y_text, f"{key.capitalize()}: {value}")
                    y_text -= 18

        c.showPage()

    c.save()


def export_match_results_table_format(matches, output_path):
    """
    Export matches to a table-based reference sheet like your example.
    Each row = (catalog code, page number)
    """
    doc = SimpleDocTemplate(output_path, pagesize=A4)
    styles = getSampleStyleSheet()
    content = []

    # Title
    title_style = ParagraphStyle(name="CenteredTitle", fontSize=14, alignment=TA_CENTER, spaceAfter=20)
    content.append(Paragraph("Tài liệu tham chiếu", title_style))
    content.append(Spacer(1, 12))

    # Build each match block
    for match in matches:
        block = match["matched_block"]
        item = match["item"]

        code = item.get("code", "N/A")  # You need to include 'code' in your parsed word items
        page = block["page"] if block else "N/A"

        table_data = [
            [Paragraph(f"<b>Catalogue mã hàng:</b><br/>{code}", styles["Normal"])],
            [Paragraph(f"<b>Trang số:</b> {page}", styles["Normal"])]
        ]

        table = Table(table_data, colWidths=[450])
        table.setStyle(TableStyle([
            ('BOX', (0, 0), (-1, -1), 1, colors.black),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))

        content.append(table)
        content.append(Spacer(1, 12))

    doc.build(content)

