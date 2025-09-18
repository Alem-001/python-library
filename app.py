from fastapi import FastAPI, UploadFile, Form
from fastapi.responses import StreamingResponse, JSONResponse
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt
import io, json

app = FastAPI()

# -------------------------------
# Helpers
# -------------------------------

def _replace_in_textframe(tf, mapping: dict):
    """Replace {{tokens}} in a full text frame (handles split runs)."""
    text = tf.text or ""
    for k, v in mapping.items():
        if isinstance(v, (dict, list)):
            continue
        text = text.replace(f"{{{{{k}}}}}", str(v or ""))
    tf.text = text

def _replace_in_table_cells(table, mapping: dict):
    """Replace {{tokens}} inside an existing table."""
    for row in table.rows:
        for cell in row.cells:
            if cell.text and "{{" in cell.text:
                new_text = cell.text
                for k, v in mapping.items():
                    if isinstance(v, (dict, list)):
                        continue
                    new_text = new_text.replace(f"{{{{{k}}}}}", str(v or ""))
                if new_text != cell.text:
                    cell.text = new_text

def _remove_shape(slide, shape):
    """Delete a shape from a slide."""
    sp = shape._element
    sp.getparent().remove(sp)

def _rebuild_table_from_json(slide, placeholder_shape, financials: dict):
    """
    Build a new table where the placeholder text is {{TABLE:financials}}.
    JSON format:
    {
      "years": ["2022","2023","2024"],
      "rows": [
        {"item":"Revenue","values":["3.85m","4.27m","6.53m"]},
        {"item":"EBITDA","values":["15.9%","18.1%","25.1%"]}
      ]
    }
    """
    years = financials.get("years", [])
    rows = financials.get("rows", [])

    n_rows = 1 + len(rows)       # header + data
    n_cols = 1 + len(years)      # item + values

    # Position from placeholder
    left, top, width, height = placeholder_shape.left, placeholder_shape.top, placeholder_shape.width, placeholder_shape.height
    _remove_shape(slide, placeholder_shape)

    # Add new table
    table_shape = slide.shapes.add_table(n_rows, n_cols, left, top, width, height)
    table = table_shape.table

    # Header row
    table.cell(0, 0).text = "Item"
    for c, y in enumerate(years, start=1):
        if c < len(table.columns):
            table.cell(0, c).text = str(y)

    # Data rows
    for r, rowdata in enumerate(rows, start=1):
        table.cell(r, 0).text = str(rowdata.get("item", ""))
        vals = rowdata.get("values", [])
        for c, val in enumerate(vals, start=1):
            if c < len(table.columns):
                table.cell(r, c).text = str(val)

    # normalize font size
    for row in table.rows:
        for cell in row.cells:
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(12)

def _walk_shapes(slide, mapping: dict):
    for shp in slide.shapes:
        # Rebuild table from marker like {{TABLE:financials}}
        if getattr(shp, "has_text_frame", False) and shp.text and "{{TABLE:" in shp.text:
            key = shp.text.strip().replace("{{TABLE:", "").replace("}}", "").strip()
            if key in mapping and isinstance(mapping[key], dict):
                _rebuild_table_from_json(slide, shp, mapping[key])
            continue

        # Normal text boxes
        if getattr(shp, "has_text_frame", False):
            _replace_in_textframe(shp.text_frame, mapping)

        # Tables with token placeholders
        if getattr(shp, "has_table", False):
            _replace_in_table_cells(shp.table, mapping)

        # Groups (recurse)
        if shp.shape_type == MSO_SHAPE_TYPE.GROUP:
            _walk_shapes(shp, mapping)

# -------------------------------
# Routes
# -------------------------------

@app.get("/health")
def health():
    return {"ok": True}

@app.post("/fill")
async def fill(
    template: UploadFile,
    data: UploadFile | None = None,
    json_text: str = Form(None),
):
    try:
        # load mapping
        if data:
            mapping = json.loads((await data.read()).decode("utf-8"))
        elif json_text:
            mapping = json.loads(json_text)
        else:
            return JSONResponse({"error": "No JSON provided"}, status_code=400)

        # load pptx
        prs = Presentation(io.BytesIO(await template.read()))

        # walk all slides/shapes
        for slide in prs.slides:
            _walk_shapes(slide, mapping)

        # output pptx
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
