# autofill_word_pptx.py  –  2025-05-21
# pip install streamlit python-docx python-pptx pandas openpyxl
import streamlit as st, pandas as pd
from io import BytesIO
import zipfile

# ---------- Word ----------
from docx import Document
from docx.oxml.ns import qn as qn_docx
from docx.shared import RGBColor as DocxRGBColor
from docx.oxml import OxmlElement

# ---------- PowerPoint ----------
from pptx import Presentation
from pptx.dml.color import RGBColor as PptxRGBColor
from pptx.util import Pt, Inches
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.oxml.ns import qn as qn_pptx


# ───────────────────────── HELPERS ──────────────────────────
def parse_value_and_color(value):
    """Parse value|color syntax, return (value, color) or (value, None)."""
    if isinstance(value, str) and '|' in value:
        parts = value.split('|', 1)
        if len(parts) == 2 and parts[1].lower() in ('green', 'yellow', 'red'):
            return parts[0], parts[1].lower()
    return value, None


def get_color_rgb(color):
    """Return RGB tuple for color name or None if not recognized."""
    color_map = {
        'green': (0, 255, 0),
        'yellow': (255, 255, 0),
        'red': (255, 0, 0)
    }
    return color_map.get(color)


# ───────────────────────── WORD ──────────────────────────
def replace_in_word(doc: Document, placeholders: dict):
    replacement_count = 0
    # paragraphs
    for para in doc.paragraphs:
        for run in para.runs:
            for k, v in placeholders.items():
                tok = f'{{{k}}}'
                if tok in run.text:
                    value, color = parse_value_and_color(v)
                    replacement_count += run.text.count(tok)  # Count before replacing
                    run.text = run.text.replace(tok, str(value))
                    if color and k.startswith('dot|'):
                        run.font.color.rgb = DocxRGBColor(*get_color_rgb(color))

    # tables (borders preserved)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        for k, v in placeholders.items():
                            tok = f'{{{k}}}'
                            if tok in run.text:
                                value, color = parse_value_and_color(v)
                                replacement_count += run.text.count(tok)  # Count before replacing
                                run.text = run.text.replace(tok, str(value))
                                if color and k.startswith('dot|'):
                                    run.font.color.rgb = DocxRGBColor(*get_color_rgb(color))
                                elif color:  # Apply cell background color
                                    cell_xml = cell._element
                                    tcPr = cell_xml.get_or_add_tcPr()
                                    shd = tcPr.find(qn_docx('w:shd'))
                                    if shd is None:
                                        shd = OxmlElement('w:shd')
                                        tcPr.append(shd)
                                    shd.set(qn_docx('w:fill'), f'{get_color_rgb(color)[0]:02X}{get_color_rgb(color)[1]:02X}{get_color_rgb(color)[2]:02X}')
    return doc, replacement_count


def save_word(doc):
    buf = BytesIO(); doc.save(buf); buf.seek(0); return buf


# ─────────────────────── PPTX HELPERS ───────────────────────
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


def _process_shape_text(shape, placeholders, replacement_count):
    """Replace tokens in a shape’s text frame and handle dots (keeps run formatting)."""
    if not shape.has_text_frame:
        return replacement_count
    tf = shape.text_frame
    for para in tf.paragraphs:
        for run in para.runs:
            for k, v in placeholders.items():
                tok = f'{{{k}}}'
                if tok in run.text:
                    value, color = parse_value_and_color(v)
                    replacement_count += run.text.count(tok)  # Count before replacing
                    run.text = run.text.replace(tok, str(value))
                    if color and k.startswith('dot|'):
                        run.font.color.rgb = PptxRGBColor(*get_color_rgb(color))
    # kill outline
    if shape.line:
        shape.line.color.rgb = PptxRGBColor(255, 255, 255)
        shape.line.width = Pt(0)
    return replacement_count


def _process_shapes_collection(shapes, placeholders, slide=None):
    """Handle text, table borders, dashed lines, and dots in a shapes collection."""
    replacement_count = 0
    _purge_dashed_shapes(shapes)  # first, delete dashed lines

    for shp in list(shapes):      # list() to avoid iterator issues if we removed shapes
        if shp.shape_type == MSO_SHAPE_TYPE.GROUP:
            replacement_count += _process_shapes_collection(shp.shapes, placeholders, slide)  # recurse into group
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
                                    value, color = parse_value_and_color(v)
                                    replacement_count += run.text.count(tok)  # Count before replacing
                                    run.text = run.text.replace(tok, str(value))
                                    if color and not k.startswith('dot|'):  # Apply cell background color
                                        cell.fill.solid()
                                        cell.fill.fore_color.rgb = PptxRGBColor(*get_color_rgb(color))
        else:
            # Handle dots for non-table shapes
            if slide and any(k.startswith('dot|') for k in placeholders.keys()) and shp.has_text_frame:
                for para in shp.text_frame.paragraphs:
                    for run in para.runs:
                        for k, v in placeholders.items():
                            tok = f'{{{k}}}'
                            if tok in run.text and k.startswith('dot|'):
                                value, color = parse_value_and_color(v)
                                if color:
                                    # Add a small circle shape at the shape’s position
                                    left, top, width, height = shp.left, shp.top, Inches(0.1), Inches(0.1)
                                    dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, width, height)
                                    dot.fill.solid()
                                    dot.fill.fore_color.rgb = PptxRGBColor(*get_color_rgb(color))
                                    dot.line.color.rgb = PptxRGBColor(*get_color_rgb(color))
                                    replacement_count += run.text.count(tok)
                                    run.text = run.text.replace(tok, '')
                                continue
                            replacement_count = _process_shape_text(shp, placeholders, replacement_count)
            else:
                replacement_count = _process_shape_text(shp, placeholders, replacement_count)
    return replacement_count


def replace_in_pptx(prs: Presentation, placeholders: dict):
    replacement_count = 0
    # Slide Masters & Layouts first (they sit “under” pictures/text on slides)
    for master in prs.slide_masters:
        replacement_count += _process_shapes_collection(master.shapes, placeholders)
    for layout in prs.slide_layouts:
        replacement_count += _process_shapes_collection(layout.shapes, placeholders)

    # Normal slides
    for slide in prs.slides:
        replacement_count += _process_shapes_collection(slide.shapes, placeholders, slide)
    return prs, replacement_count


def save_pptx(prs):
    buf = BytesIO(); prs.save(buf); buf.seek(0); return buf


# ───────────────────────── STREAMLIT UI ──────────────────────────
st.set_page_config(page_title="Auto-Fill Word / PPTX", layout="centered")
st.title("📝 Auto-fill Word or PowerPoint templates")

kind = st.radio("Template type:", ("Word (.docx)", "PowerPoint (.pptx)"), horizontal=True)
tfile = st.file_uploader("Upload template",
                         type=["docx"] if kind.startswith("Word") else ["pptx"])
xfile = st.file_uploader("Upload Excel with keywords & values", type=["xlsx"])
multi_report = st.checkbox("Generate multiple reports (one per value column)", value=False)

if tfile and xfile:
    df = pd.read_excel(xfile)
    if df.empty:
        st.error("Excel is empty"); st.stop()
    cols = df.columns.tolist()
    kw_col = st.selectbox("Keyword column", cols, key="kw")
    val_col = st.selectbox("Value column (first column for multiple reports)", cols, key="val")

    if st.button("Generate file(s)"):
        # Get value columns: either the selected one or all to its right if multi_report
        if multi_report:
            val_cols = df.columns[df.columns.get_loc(val_col):].tolist()
        else:
            val_cols = [val_col]

        if multi_report:
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                for i, v_col in enumerate(val_cols, 1):
                    keys = df[kw_col].astype(str).tolist()
                    vals = df[v_col].astype(str).tolist()
                    if len(keys) != len(vals):
                        st.error(f"Columns {kw_col} and {v_col} must have same number of rows"); st.stop()

                    mapping = dict(zip(keys, vals))

                    if kind.startswith("Word"):
                        doc = Document(tfile)
                        filled, count = replace_in_word(doc, mapping)
                        buf = save_word(filled)
                        file_name = f"filled_{i}.docx"
                        zf.writestr(file_name, buf.getvalue())
                        st.write(f"Report {i} ({v_col}): Replaced {count} keywords")
                    else:
                        prs = Presentation(tfile)
                        filled, count = replace_in_pptx(prs, mapping)
                        buf = save_pptx(filled)
                        file_name = f"filled_{i}.pptx"
                        zf.writestr(file_name, buf.getvalue())
                        st.write(f"Report {i} ({v_col}): Replaced {count} keywords")
            zip_buffer.seek(0)
            st.download_button("⬇️ Download all reports (ZIP)",
                               data=zip_buffer,
                               file_name="filled_reports.zip",
                               mime="application/zip")
        else:
            # Single report
            keys = df[kw_col].astype(str).tolist()
            vals = df[val_col].astype(str).tolist()
            if len(keys) != len(vals):
                st.error(f"Columns {kw_col} and {val_col} must have same number of rows"); st.stop()

            mapping = dict(zip(keys, vals))

            if kind.startswith("Word"):
                doc = Document(tfile)
                filled, count = replace_in_word(doc, mapping)
                buf = save_word(filled)
                st.download_button("⬇️ Download filled Word",
                                   data=buf, file_name="filled.docx",
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                st.write(f"Report 1 ({val_col}): Replaced {count} keywords")
            else:
                prs = Presentation(tfile)
                filled, count = replace_in_pptx(prs, mapping)
                buf = save_pptx(filled)
                st.download_button("⬇️ Download filled PowerPoint",
                                   data=buf, file_name="filled.pptx",
                                   mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
                st.write(f"Report 1 ({val_col}): Replaced {count} keywords")
else:
    st.info("Upload both template and Excel to begin.")
