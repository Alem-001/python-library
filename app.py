# app.py
from fastapi import FastAPI, UploadFile, Form
from fastapi.responses import StreamingResponse, JSONResponse
from pptx import Presentation
import io, json

app = FastAPI()

def replace_in_textframe(tf, mapping: dict):
    """Replace {{tokens}} in an entire text frame (robust vs. split runs)."""
    text = tf.text or ""
    for k, v in mapping.items():
        text = text.replace(f"{{{{{k}}}}}", "" if v is None else str(v))
    tf.text = text  # setting text rebuilds runs; fine for placeholders

def replace_in_shape(shape, mapping: dict):
    # normal text boxes, titles, placeholders
    if hasattr(shape, "has_text_frame") and shape.has_text_frame:
        replace_in_textframe(shape.text_frame, mapping)

    # tables
    if hasattr(shape, "has_table") and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                replace_in_textframe(cell.text_frame, mapping)

    # grouped shapes
    if hasattr(shape, "shape_type") and shape.shape_type == 6:  # MSO_SHAPE_TYPE.GROUP
        for shp in shape.group_shapes:
            replace_in_shape(shp, mapping)

@app.get("/health")
def health():
    return {"ok": True}

@app.post("/fill")
async def fill(
    template: UploadFile,
    data: UploadFile | None = None,
    json_text: str = Form(default=None),
):
    """
    Accepts either:
      - template: UploadFile (.pptx)
      - data: UploadFile (.json) OR
      - json_text: form field containing JSON
    Returns a filled .pptx.
    """
    try:
        # parse mapping
        if json_text:
            mapping = json.loads(json_text)
        elif data:
            mapping = json.loads((await data.read()).decode("utf-8"))
        else:
            return JSONResponse({"error": "No JSON provided"}, status_code=400)

        # load pptx
        prs = Presentation(io.BytesIO(await template.read()))

        # replace across all shapes on all slides
        for slide in prs.slides:
            for shape in slide.shapes:
                replace_in_shape(shape, mapping)

        # stream result
        out = io.BytesIO()
        prs.save(out)
        out.seek(0)
        return StreamingResponse(
            out,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": 'attachment; filename="filled.pptx"'},
        )
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)
