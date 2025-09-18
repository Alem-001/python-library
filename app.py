# app.py
import io, json, math
from fastapi import FastAPI, UploadFile, Form
from fastapi.responses import StreamingResponse, PlainTextResponse
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from pptx.oxml.xmlchemy import OxmlElement

app = FastAPI()

# --------- helpers ---------
def iter_text_shapes(slide):
    for sh in slide.shapes:
        if hasattr(sh, "text_frame") and sh.text_frame is not None:
            yield sh

def replace_text_tokens(prs, data):
    """
    Replace {{token}} in all text boxes and table cells with strings from data dict (flat only).
    Nested dicts are ignored here—use table handler for them.
    """
    def replace_in_textframe(tf, mapping):
        if not tf: return
        # Work per paragraph to preserve formatting reasonably
        for p in tf.paragraphs:
            text = "".join(run.text for run in p.runs) if p.runs else p.text
            changed = False
            for k, v in mapping.items():
                token = "{{" + k + "}}"
                if isinstance(v, (str, int, float)) and token in text:
                    text = text.replace(token, str(v))
                    changed = True
            if changed:
                # rewrite runs minimally: one run with the new text
                for _ in range(len(p.runs)):
                    p.runs[0]._element.getparent().remove(p.runs[0]._element)
                p.text = text

    # flatten only 1st-level scalars
    flat = {k: v for k, v in data.items() if isinstance(v, (str, int, float))}
    for slide in prs.slides:
        for sh in iter_text_shapes(slide):
            replace_in_textframe(sh.text_frame, flat)

        # also replace inside tables
        for sh in slide.shapes:
            if sh.has_table:
                tbl = sh.table
                for r in tbl.rows:
                    for c in r.cells:
                        replace_in_textframe(c.text_frame, flat)

def _add_table_at_placeholder(slide, ph_shape, headers, rows):
    # Create a table that exactly fits the placeholder box, then delete the box
    cols = len(headers) if headers else len(rows[0])
