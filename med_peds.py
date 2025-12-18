# app.py
import streamlit as st
from datetime import date
from io import BytesIO
from docx import Document

st.set_page_config(page_title="CCC Letter Generator", layout="centered")

TODAY_STR = date.today().strftime("%B %d, %Y")

st.title("CCC Letter Generator")
st.caption(f"Today's date: {TODAY_STR}")

# ----------------------------
# Helper functions
# ----------------------------
def replace_simple_text(paragraph, replacements: dict):
    if not paragraph.text:
        return

    full_text = paragraph.text
    new_text = full_text
    for old, new in replacements.items():
        new_text = new_text.replace(old, new)

    if new_text != full_text:
        paragraph.clear()
        paragraph.add_run(new_text)

def replace_everywhere(doc: Document, replacements: dict):
    for p in doc.paragraphs:
        replace_simple_text(p, replacements)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_simple_text(p, replacements)

    for section in doc.sections:
        for p in section.header.paragraphs:
            replace_simple_text(p, replacements)
        for p in section.footer.paragraphs:
            replace_simple_text(p, replacements)

def insert_date_at_top(doc: Document, date_line: str):
    if doc.paragraphs:
        doc.paragraphs[0].insert_paragraph_before(date_line)
    else:
        doc.add_paragraph(date_line)

def replace_ccc_text_block(doc: Document, replacement_text: str):
    """
    Replaces the paragraph containing CCC_TEXT with
    multiple paragraphs based on textarea newlines.
    """
    blocks = [line.strip() for line in replacement_text.split("\n") if line.strip()]

    for i, p in enumerate(doc.paragraphs):
        if "CCC_TEXT" in p.text:
            parent = p._p.getparent()
            idx = parent.index(p._p)

            # remove placeholder paragraph
            parent.remove(p._p)

            # insert new paragraphs at same location
            for offset, block in enumerate(blocks):
                new_p = Document().add_paragraph(block)._p
                parent.insert(idx + offset, new_p)
            return

# ----------------------------
# UI
# ----------------------------
with st.form("inputs"):
    name = st.text_input("Person's name (e.g., Jane Smith)")
    month = st.selectbox(
        "Month",
        ["January","February","March","April","May","June","July","August",
         "September","October","November","December"],
        index=date.today().month - 1,
    )
    year = st.number_input("Year", min_value=2000, max_value=2100,
                           value=date.today().year, step=1)

    ccc_text = st.text_area(
        "CCC personalized paragraph(s)",
        height=220,
        help="Line breaks will be preserved as separate paragraphs in Word."
    )

    template_file = st.file_uploader("Upload Word template (.docx)", type=["docx"])
    submitted = st.form_submit_button("Generate Word Document")

# ----------------------------
# Generate document
# ----------------------------
if submitted:
    if not template_file or not name.strip() or not ccc_text.strip():
        st.error("Please complete all required fields and upload a template.")
        st.stop()

    month_year = f"{month} {int(year)}"

    doc = Document(BytesIO(template_file.read()))

    # Simple placeholder replacements
    replace_everywhere(
        doc,
        {
            "xxx": name.strip(),
            "Date": TODAY_STR,
            "Month of Year": month_year,
            "Month of  Year": month_year,
        }
    )

    # Replace CCC_TEXT block with multi-paragraph input
    replace_ccc_text_block(doc, ccc_text)

    # Output
    out = BytesIO()
    doc.save(out)
    out.seek(0)

    safe_name = name.strip().replace(" ", "_")
    filename = f"CCC_Letter_{safe_name}.docx"

    st.success("Document created successfully.")
    st.download_button(
        "Download Word Document",
        data=out.getvalue(),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

