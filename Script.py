import openpyxl
from docx import Document

EXCEL_PATH = r"C:\Users\AceGr\Desktop\EdProject\Test.xlsx"
OUTPUT_DOCX = r"C:\Users\AceGr\Downloads\Group_Status_Report.docx"

def norm(v):
    return "" if v is None else str(v).strip()


def display_value(status):
    return norm(status)


def payer_label(header, fallback_letter):
    h = norm(header)
    if h:
        return h
    return f"Payer {fallback_letter}"


wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
ws = wb.active

payer_labels = []
letters = "ABCDEFGH"
for i, c in enumerate(range(2, 10)):
    header = ws.cell(2, c).value
    payer_labels.append(payer_label(header, letters[i]))

doc = Document()

first_group = True
for r in range(4, ws.max_row + 1):
    group = norm(ws.cell(r, 1).value)
    if not group:
        continue

    if not first_group:
        doc.add_page_break()
    first_group = False

    doc.add_paragraph(f"Group Name {group}")
    doc.add_paragraph("Received Plans")

    statuses = [ws.cell(r, c).value for c in range(2, 10)]
    for label, status in zip(payer_labels, statuses):
        value = display_value(status)
        line = f"{label} {value}".rstrip()
        doc.add_paragraph(line)

doc.save(OUTPUT_DOCX)
print(f"Saved: {OUTPUT_DOCX}")
