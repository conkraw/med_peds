# app.py
import streamlit as st
from datetime import date
from io import BytesIO
from docx import Document

st.set_page_config(page_title="CCC Letter Generator", layout="centered")

TODAY_STR = date.today().strftime("%B %d, %Y")  # e.g., December 18, 2025

st.title("CCC Letter Generator")
st.caption(f"Today's date: {TODAY_STR}")

# ----------------------------
# Helpers
# ----------------------------
def _replace_in_paragraph(paragraph, replacements: dict):
    """
    Robust-ish text replacement that works even if Word split the placeholder across runs.
    Keeps formatting of the first run, but may flatten mixed formatting within the paragraph.
    """
    if paragraph is None or paragraph.text is None:
        return

    full_text = paragraph.text
    new_text = full_text
    for old, new in replacements.items():
        new_text = new_text.replace(old, new)

    if new_text == full_text:
        return

    # If there are runs, keep formatting of first run, clear the rest
    if paragraph.runs:
        paragraph.runs[0].text = new_text
        for r in paragraph.runs[1:]:
            r.text = ""
    else:
        paragraph.add_run(new_text)

def _replace_everywhere(doc: Document, replacements: dict):
    # Body paragraphs
    for p in doc.paragraphs:
        _replace_in_paragraph(p, replacements)

    # Tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _replace_in_paragraph(p, replacements)

    # Headers/footers
    for section in doc.sections:
        header = section.header
        footer = section.footer
        for p in header.paragraphs:
            _replace_in_paragraph(p, replacements)
        for p in footer.paragraphs:
            _replace_in_paragraph(p, replacements)

def _insert_date_at_top(doc: Document, date_line: str):
    # Insert as first paragraph if possible
    if doc.paragraphs:
        doc.paragraphs[0].insert_paragraph_before(date_line)
    else:
        doc.add_paragraph(date_line)

# ----------------------------
# UI
# ----------------------------
with st.form("inputs"):
    name = st.text_input("Person's name (e.g., Jane Smith)", value="")
    month = st.selectbox(
        "Month",
        ["January","February","March","April","May","June","July","August",
         "September","October","November","December"],
        index=date.today().month - 1,
    )
    year = st.number_input("Year", min_value=2000, max_value=2100, value=date.today().year, step=1)

    template_file = st.file_uploader("Upload your Word template (.docx)", type=["docx"])
    submitted = st.form_submit_button("Generate Word Document")

if submitted:
    if not template_file:
        st.error("Please upload a .docx template.")
        st.stop()
    if not name.strip():
        st.error("Please enter the person's name.")
        st.stop()

    month_year = f"{month} {int(year)}"

    # Load template into python-docx
    template_bytes = BytesIO(template_file.read())
    doc = Document(template_bytes)

    # Insert date line at very top (optional but per your request)
    #_insert_date_at_top(doc, TODAY_STR)

    # Replace placeholders
    replacements = {
        "xxx": name.strip(),
        "Date": TODAY_STR,
        "Month of Year": month_year,
        "Month of  Year": month_year,  # common double-space typo guard
        "Month  of Year": month_year,  # another typo guard
    }
    _replace_everywhere(doc, replacements)

    # Save to bytes for download
    out = BytesIO()
    doc.save(out)
    out.seek(0)

    safe_name = "".join(c for c in name.strip() if c.isalnum() or c in (" ", "_", "-")).strip().replace(" ", "_")
    filename = f"CCC_Letter_{safe_name}.docx"

    st.success("Done! Download your generated document below.")
    st.download_button(
        label="Download Word Document",
        data=out.getvalue(),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

