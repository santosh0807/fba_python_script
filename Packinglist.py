import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import letter, landscape
from PyPDF2 import PdfMerger
import os

# Load Excel
file_path = "packing_list.xlsx"
df = pd.read_excel(file_path)

# Group by Box No
grouped = df.groupby("Box No")

styles = getSampleStyleSheet()

generated_files = []

# 🔹 Step 1: Generate PDFs
for box_no, data in grouped:
    shipment_id = data["Shipment id"].iloc[0]

    file_name = f"Packing_Slip_Box_{int(box_no)}.pdf"
    generated_files.append(file_name)

    doc = SimpleDocTemplate(file_name, pagesize=landscape(letter))
    elements = []

    # ✅ FC Value from Excel (Column G)
    fc_value = ""
    if "FC" in data.columns and pd.notna(data["FC"].iloc[0]):
        fc_value = str(data["FC"].iloc[0]).strip()

    fc_text = f"FC - {fc_value}" if fc_value else "FC"

    # Title
    elements.append(Paragraph(f"<para align='center'><b>{fc_text}</b></para>", styles['Title']))
    elements.append(Spacer(1, 10))

    elements.append(Paragraph("<para align='center'><b>Packing Slip</b></para>", styles['Title']))
    elements.append(Spacer(1, 10))

    # Shipment Details
    elements.append(Paragraph(f"<b>Shipment ID:</b> {shipment_id}", styles['Normal']))
    elements.append(Paragraph(f"<b>Box No:</b> {int(box_no)}", styles['Normal']))
    elements.append(Spacer(1, 10))

    # Table Header
    table_data = [["SKU", "Title", "FNSKU", "Qty"]]

    # Table Rows
    for _, row in data.iterrows():
        table_data.append([
            str(row["SKU"]),
            str(row["Title"]),
            str(row["FNSKU"]),
            str(int(row["Qty"]))
        ])

    # Table Design
    row_heights = [50] + [45] * (len(table_data) - 1)

    table = Table(
        table_data,
        colWidths=[150, 350, 150, 70],
        rowHeights=row_heights
    )

    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.black),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),

        ('FONTSIZE', (0, 0), (-1, -1), 11),
        ('FONTSIZE', (3, 1), (3, -1), 15),

        ('ALIGN', (3, 1), (-1, -1), 'CENTER'),
        ('VALIGN', (3, 1), (3, -1), 'MIDDLE'),
    ]))

    elements.append(table)

    # Build PDF
    doc.build(elements)

# 🔹 Step 2: Merge PDFs
merger = PdfMerger()

# Sort files by box number
generated_files.sort(key=lambda x: int(x.split("_")[-1].replace(".pdf", "")))

for pdf in generated_files:
    merger.append(pdf)

merger.write("Final_Packing_Slips.pdf")
merger.close()

print("✅ All PDFs merged into Final_Packing_Slips.pdf")