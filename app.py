# --- add near your other imports ---
from pptx import Presentation
from pptx.util import Inches

def safe_fill_quality_table(prs, quality):
    """Fill a table named 'quality' without deleting rows/columns."""
    if not quality:
        return

    headers = quality.get("headers", [])
    rows = quality.get("rows", [])

    for slide in prs.slides:
        for shape in slide.shapes:
            # must be a table and named 'quality'
            if not getattr(shape, "has_table", False):
                continue
            if getattr(shape, "name", "").strip().lower() != "quality":
                continue

            tbl = shape.table

            # --- ensure enough columns (only ADD, never delete) ---
            need_cols = max(len(headers), max((len(r) for r in rows), default=0))
            have_cols = len(tbl.columns)
            if need_cols > have_cols:
                for _ in range(need_cols - have_cols):
                    tbl.add_column(Inches(1.5))  # width heuristic

            # --- ensure enough rows (only ADD, never delete) ---
            need_rows = 1 + len(rows)  # header + data rows
            have_rows = len(tbl.rows)
            if need_rows > have_rows:
                for _ in range(need_rows - have_rows):
                    tbl.add_row()

            # --- write header (row 0) ---
            for c in range(need_cols):
                text = headers[c] if c < len(headers) else ""
                cell = tbl.cell(0, c)
                tf = cell.text_frame
                tf.clear()
                tf.paragraphs[0].text = str(text)

            # clear any extra header cells beyond need_cols
            for c in range(need_cols, len(tbl.columns)):
                tf = tbl.cell(0, c).text_frame
                tf.clear()
                tf.paragraphs[0].text = ""

            # --- write data rows ---
            for r_idx, row in enumerate(rows, start=1):
                for c in range(need_cols):
                    txt = row[c] if c < len(row) else ""
                    cell = tbl.cell(r_idx, c)
                    tf = cell.text_frame
                    tf.clear()
                    tf.paragraphs[0].text = str(txt)

            # clear any leftover existing rows (keep structure intact)
            for r in range(1 + len(rows), len(tbl.rows)):
                for c in range(len(tbl.columns)):
                    tf = tbl.cell(r, c).text_frame
                    tf.clear()
                    tf.paragraphs[0].text = ""

            return  # stop after first matching table

# --- inside your /fill handler, after prs = Presentation(...) and data = json.loads(...) ---
safe_fill_quality_table(prs, data.get("quality"))
