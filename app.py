from fastapi import FastAPI, UploadFile, Form
from fastapi.responses import StreamingResponse, JSONResponse
from pptx import Presentation
import io
import json

app = FastAPI()

def replace_textframe(tf, mapping: dict):
    for p in tf.paragraphs:
        for run in p.runs:
            text = run.text or ""
            for k, v in mapping.items():
                token = "{{ " + k + " }}"
                if token in text:
                    text = text.replace(token, str(v or ""))
            run.text = text

def replace_in_shape(shape, mapping: dict):
    # Text frames
    if hasattr(shape, "has_text_frame") and shape.has_text_frame:
        replace_textframe(shape.text_frame, mapping)

    # Tables
    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                if cell.text_frame:
                    replace_textframe(cell.text_frame, mapping)

@app.get("/health")
def health():
    return {"ok": True}

@app.post("/fill")
async def fill(
    template: UploadFile,
    data: UploadFile | None = None,
    json_text: str = Form(None)
):
    """
    Accept either:
    - template: .pptx file (required)
    - data:     .json file OR
    - json_text: string field containing JSON

    Returns the filled .pptx with {{ tokens }} replaced.
    """
    # Get mapping
    mapping = {}
    if data is not None:
        mapping = json.loads((await data.read()).decode("utf-8"))
    elif json_text:
        mapping = json.loads(json_text)

    # Load PPTX
    prs_bytes = await template.read()
    prs = Presentation(io.BytesIO(prs_bytes))

    # Replace tokens across slides
    for slide in prs.slides:
        # Slide-level shapes
        for shape in slide.shapes:
            replace_in_shape(shape, mapping)
        # Slide master placeholders (safe pass)
        if hasattr(slide, "placeholders"):
            for ph in slide.placeholders:
                try:
                    replace_in_shape(ph, mapping)
                except Exception:
                    pass

    # Stream back
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return StreamingResponse(
        out,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": 'attachment; filename="filled.pptx"'}
    )
