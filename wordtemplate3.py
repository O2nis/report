# autofill_word_pptx.py  –  2025-05-21
# pip install streamlit python-docx python-pptx pandas openpyxl
import streamlit as st, pandas as pd
from io import BytesIO

# ---------- Word ----------
from docx import Document

# ---------- PowerPoint ----------
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn as qn_pptx


# ───────────────────────── WORD ──────────────────────────
def replace_in_word(doc: Document, placeholders: dict):
    # paragraphs
    for para in doc.paragraphs:
        for run in para.runs:
            for k, v in placeholders.items():
                tok = f'{{{k}}}'
                if tok in run.text:
                    run.text = run.text.replace(tok, str(v))

    # tables (borders preserved)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        for k, v in placeholders.items():
                            tok = f'{{{k}}}'
                            if tok in run.text:
                                run.text = run.text.replace(tok, str(v))
    return doc


def save_word(doc):
    buf = BytesIO(); doc.save(buf); buf.seek(0); return buf


# ───────────────────────  PPTX HELPERS  ───────────────────────
def _strip_table_borders(shape):
    """Remove borders inside a PPT table."""
    if not shape.has_table:
        return
    tbl = shape.table
    for row in tbl.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            for tag in (
                "a:lnL", "a:lnR", "a:lnT", "a:lnB",
                "a:lnTlToBr", "a:lnBlToTr"
            ):
                ln = tcPr.find(qn_pptx(tag))
                if ln is not None:
                    tcPr.remove(ln)


def _purge_dashed_shapes(shapes):
    """Delete dashed-line shapes from a pptx.shapes collection (recursive for groups)."""
    doomed = []
    for shp in shapes:
        if shp.shape_type == MSO_SHAPE_TYPE.GROUP:
            _purge_dashed_shapes(shp.shapes)  # recurse into group shapes
        elif shp.shape_type in (MSO_SHAPE_TYPE.LINE, MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.FREEFORM):
            try:
                ln = shp.line
                if ln and ln.dash_style is not None:  # dash_style is None for solid lines
                    doomed.append(shp)
            except Exception:
                pass  # Skip shapes that don't have line properties
    # Remove after iteration to avoid modifying collection during loop
    for shp in doomed:
        try:
            shp._element.getparent().remove(shp._element)
        except Exception:
            pass  # Skip if removal fails (e.g., shape already removed)


def _process_shape_text(shape, placeholders):
    """Replace tokens inside a shape’s text frame (keeps run formatting)."""
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    for para in tf.paragraphs:
        for run in para.runs:
            for k, v in placeholders.items():
                tok = f'{{{k}}}'
                if tok in run.text:
                    run.text = run.text.replace(tok, str(v))
    # kill outline
    if shape.line:
        shape.line.color.rgb = RGBColor(255, 255, 255)
        shape.line.width = Pt(0)


def _process_shapes_collection(shapes, placeholders):
    """Handle text, table borders & dashed lines inside a shapes collection."""
    _purge_dashed_shapes(shapes)  # first, delete dashed lines

    for shp in list(shapes):      # list() to avoid iterator issues if we removed shapes
        if shp.shape_type == MSO_SHAPE_TYPE.GROUP:
            _process_shapes_collection(shp.shapes, placeholders)  # recurse into group
        elif shp.shape_type == MSO_SHAPE_TYPE.TABLE:
            _strip_table_borders(shp)
            tbl = shp.table
            for row in tbl.rows:
                for cell in row.cells:
                    for para in cell.text_frame.paragraphs:
                        for run in para.runs:
                            for k, v in placeholders.items():
                                tok = f'{{{k}}}'
                                if tok in run.text:
                                    run.text = run.text.replace(tok, str(v))
        else:
            _process_shape_text(shp, placeholders)


def replace_in_pptx(prs: Presentation, placeholders: dict):
    # Slide Masters & Layouts first (they sit “under” pictures/text on slides)
    for master in prs.slide_masters:
        _process_shapes_collection(master.shapes, placeholders)
    for layout in prs.slide_layouts:
        _process_shapes_collection(layout.shapes, placeholders)

    # Normal slides
    for slide in prs.slides:
        _process_shapes_collection(slide.shapes, placeholders)
    return prs


def save_pptx(prs):
    buf = BytesIO(); prs.save(buf); buf.seek(0); return buf


# ───────────────────────── STREAMLIT UI ──────────────────────────
st.set_page_config(page_title="Auto-Fill Word / PPTX", layout="centered")
st.title("📝 Auto-fill Word or PowerPoint templates")

kind = st.radio("Template type:", ("Word (.docx)", "PowerPoint (.pptx)"), horizontal=True)
tfile = st.file_uploader("Upload template",
                         type=["docx"] if kind.startswith("Word") else ["pptx"])
xfile = st.file_uploader("Upload Excel with keywords & values", type=["xlsx"])

if tfile and xfile:
    df = pd.read_excel(xfile)
    if df.empty:
        st.error("Excel is empty"); st.stop()
    cols = df.columns.tolist()
    kw_col  = st.selectbox("Keyword column", cols, key="kw")
    val_col = st.selectbox("Value column",   cols, key="val")

    if st.button("Generate file"):
        keys  = df[kw_col].astype(str).tolist()
        vals  = df[val_col].astype(str).tolist()
        if len(keys) != len(vals):
            st.error("Columns must have same number of rows"); st.stop()

        mapping = dict(zip(keys, vals))

        if kind.startswith("Word"):
            doc = Document(tfile)
            filled = replace_in_word(doc, mapping)
            buf = save_word(filled)
            st.download_button("⬇️ Download filled Word",
                               data=buf, file_name="filled.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        else:
            prs = Presentation(tfile)
            filled = replace_in_pptx(prs, mapping)
            buf = save_pptx(filled)
            st.download_button("⬇️ Download filled PowerPoint",
                               data=buf, file_name="filled.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
else:
    st.info("Upload both template and Excel to begin.")