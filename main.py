from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
import io
import openpyxl
from docx import Document

app = FastAPI()

TEMPLATE_PATH = "template.docx"
RECEIVED_MARK = "X"


def normalize(v):
    return "" if v is None else str(v).strip()


def mark_for_status(status):
    s = normalize(status).lower()
    if s == "received":
        return RECEIVED_MARK
    if s == "not eligible":
        return "Not Eligible"
    return ""


def fill_template(group_name, payer_labels, statuses):
    doc = Document(TEMPLATE_PATH)

    for p in doc.paragraphs:
        t = p.text.strip()

        if t.startswith("Group Name"):
            p.text = f"Group Name {group_name}"
            continue

        if t.startswith("Payer "):
            label = t.split("\t", 1)[0].strip()
            if label in payer_labels:
                idx = payer_labels.index(label)
                p.text = f"{label}\t{mark_for_status(statuses[idx])}"

    return doc


@app.post("/generate")
async def generate(file: UploadFile = File(...)):
    data = await file.read()
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    ws = wb.active

    payer_letters = [normalize(ws.cell(2, c).value) for c in range(2, 10)]
    payer_labels = [f"Payer {x}" for x in payer_letters]

    out = Document()

    for r in range(4, ws.max_row + 1):
        group_name = normalize(ws.cell(r, 1).value)
        if not group_name:
            continue

        statuses = [ws.cell(r, c).value for c in range(2, 10)]
        temp = fill_template(group_name, payer_labels, statuses)

        for p in temp.paragraphs:
            out.add_paragraph(p.text)
        out.add_page_break()

    buf = io.BytesIO()
    out.save(buf)
    buf.seek(0)

    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": "attachment; filename=Received_Plans_Report.docx"},
    )
