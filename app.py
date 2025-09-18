from fastapi import FastAPI, UploadFile, Form
from fastapi.responses import StreamingResponse, JSONResponse
from pptx import Presentation
import io, json

app = FastAPI()

def replace_placeholders(prs, mapping: dict):
    for slide in prs.slides:
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            text = shape.text_frame.text or ""
            for k, v in mapping.items():
                token = "{{" + k + "}}"
                if token in text:
                    text = text.replace(token, str(v))
            shape.text_frame.clear()
            shape.text_frame.paragraphs[0].text = text

@app.get("/health")
def health():
    return {"ok": True}

@app.post("/fill")
async def fill(template: UploadFile, json_text: str = Form(...)):
    try:
        mapping = json.loads(json_text)
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=400)

    prs = Presentation(io.BytesIO(await template.read()))
    replace_placeholders(prs, mapping)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return StreamingResponse(
        out,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": 'attachment; filename="filled.pptx"'}
    )
