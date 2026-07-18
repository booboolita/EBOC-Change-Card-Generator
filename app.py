import io
import os
import re
from datetime import date

import fitz  # PyMuPDF
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from openai import OpenAI


# ============================================================
# App configuration
# ============================================================

st.set_page_config(
    page_title="EBOC Change Card Generator",
    page_icon="🩺",
    layout="wide",
)

APP_TITLE = "EBOC Change Card Generator"
DEFAULT_LAST_UPDATED_DATE = "October 8, 2019"
MODEL_NAME = "gpt-5.5"


# ============================================================
# Helper functions: API key
# ============================================================

def get_api_key():
    """
    Get API key from Streamlit secrets first.
    If not found, allow temporary manual entry in the sidebar.
    This avoids hard-coding the key into GitHub.
    """
    api_key = None

    try:
        api_key = st.secrets.get("OPENAI_API_KEY", None)
    except Exception:
        api_key = None

    if not api_key:
        api_key = st.sidebar.text_input(
            "OpenAI API key",
            type="password",
            help=(
                "Temporary option for testing. In the deployed app, the key should be stored "
                "in Streamlit Secrets, not pasted into the code."
            ),
        )

    return api_key


# ============================================================
# Helper functions: document reading
# ============================================================

def read_docx(uploaded_file):
    """
    Read text from a Word document uploaded through Streamlit.
    Includes paragraphs and tables.
    """
    doc = Document(uploaded_file)
    parts = []

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text:
            parts.append(text)

    for table in doc.tables:
        for row in table.rows:
            row_text = " | ".join(cell.text.strip() for cell in row.cells)
            if row_text.strip():
                parts.append(row_text)

    return "\n".join(parts)


def read_pdf(uploaded_file):
    """
    Read text from a PDF uploaded through Streamlit.
    """
    pdf_bytes = uploaded_file.read()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages = []

    for page in doc:
        pages.append(page.get_text())

    return "\n".join(pages)


def read_uploaded_file(uploaded_file):
    """
    Route uploaded files to the correct reader.
    Supports DOCX and PDF.
    """
    if uploaded_file is None:
        return ""

    filename = uploaded_file.name.lower()

    if filename.endswith(".docx"):
        return read_docx(uploaded_file)

    if filename.endswith(".pdf"):
        return read_pdf(uploaded_file)

    raise ValueError("Unsupported file type. Please upload a .docx or .pdf file.")


def clip_text(text, max_chars=45000):
    """
    Keep prompts from becoming too large.
    """
    if not text:
        return ""

    if len(text) <= max_chars:
        return text

    return text[:max_chars] + "\n\n[TEXT CLIPPED BECAUSE DOCUMENT WAS LONG.]"


# ============================================================
# Helper functions: Word formatting
# ============================================================

def add_hyperlink(paragraph, url, text=None):
    """
    Add a clickable hyperlink to a Word paragraph.
    """
    if text is None:
        text = url

    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")

    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    r_pr.append(color)

    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    r_pr.append(underline)

    new_run.append(r_pr)

    text_element = OxmlElement("w:t")
    text_element.text = text
    new_run.append(text_element)

    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def add_markdown_runs(paragraph, text):
    """
    Add text to a Word paragraph while converting simple **bold** markdown to real bold.
    Also detects URLs and makes them clickable when possible.
    """
    url_pattern = r"(https?://[^\s]+)"
    parts = re.split(url_pattern, text)

    for part in parts:
        if not part:
            continue

        if re.match(url_pattern, part):
            clean_url = part.rstrip(".,);]")
            trailing = part[len(clean_url):]
            add_hyperlink(paragraph, clean_url, clean_url)
            if trailing:
                paragraph.add_run(trailing)
            continue

        bold_parts = re.split(r"(\*\*.*?\*\*)", part)

        for bold_part in bold_parts:
            if not bold_part:
                continue

            if bold_part.startswith("**") and bold_part.endswith("**"):
                run = paragraph.add_run(bold_part[2:-2])
                run.bold = True
            else:
                paragraph.add_run(bold_part)


def is_bullet_line(line):
    stripped = line.strip()
    return (
        stripped.startswith("- ")
        or stripped.startswith("● ")
        or stripped.startswith("* ")
        or stripped.startswith("• ")
    )


def is_numbered_line(line):
    stripped = line.strip()
    return re.match(r"^\d+\.\s+", stripped) is not None


def clean_bullet_text(line):
    stripped = line.strip()

    for prefix in ["- ", "● ", "* ", "• "]:
        if stripped.startswith(prefix):
            return stripped[len(prefix):].strip()

    return stripped


def clean_numbered_text(line):
    return re.sub(r"^\d+\.\s+", "", line.strip())


def set_normal_spacing(paragraph):
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.line_spacing = 1.08


def create_styled_docx(card_text, topic):
    """
    Create a styled Word document from generated card text.
    This is designed to reduce manual cleanup.
    """
    doc = Document()

    styles = doc.styles

    styles["Normal"].font.name = "Calibri"
    styles["Normal"].font.size = Pt(11)

    styles["Heading 1"].font.name = "Calibri"
    styles["Heading 1"].font.size = Pt(16)
    styles["Heading 1"].font.bold = True

    styles["Heading 2"].font.name = "Calibri"
    styles["Heading 2"].font.size = Pt(13)
    styles["Heading 2"].font.bold = True

    # Date/topic header
    header = doc.add_paragraph()
    header.alignment = WD_ALIGN_PARAGRAPH.LEFT
    header_run = header.add_run(f"{date.today().strftime('%-m/%-d/%Y') if os.name != 'nt' else date.today().strftime('%#m/%#d/%Y')}\n{topic}")
    header_run.bold = True
    header_run.font.size = Pt(10)

    doc.add_paragraph("")

    lines = card_text.splitlines()

    for raw_line in lines:
        line = raw_line.strip()

        if not line:
            doc.add_paragraph("")
            continue

        # Markdown headings
        if line.startswith("# "):
            heading_text = line.replace("# ", "").strip()
            doc.add_heading(heading_text, level=1)
            continue

        if line.startswith("## "):
            heading_text = line.replace("## ", "").strip()
            doc.add_heading(heading_text, level=2)
            continue

        if line.startswith("### "):
            heading_text = line.replace("### ", "").strip()
            doc.add_heading(heading_text, level=3)
            continue

        # Bullets
        if is_bullet_line(line):
            p = doc.add_paragraph(style="List Bullet")
            add_markdown_runs(p, clean_bullet_text(line))
            set_normal_spacing(p)
            continue

        # Numbered lists
        if is_numbered_line(line):
            p = doc.add_paragraph(style="List Number")
            add_markdown_runs(p, clean_numbered_text(line))
            set_normal_spacing(p)
            continue

        # Blockquote style from markdown
        if line.startswith("> "):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Pt(18)
            p.paragraph_format.space_after = Pt(6)
            add_markdown_runs(p, line.replace("> ", "").strip())
            continue

        # Normal paragraph
        p = doc.add_paragraph()
        add_markdown_runs(p, line)
        set_normal_spacing(p)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# ============================================================
# Prompt
# ============================================================

PHYSICIAN_CARD_PROMPT = """
You are assisting with an Evidence-Based Outcomes Center (EBOC) guideline revision.

Your job is to generate a physician-facing Change Card for clinician/workgroup review.

Core rules:
1. Use only the documents and text provided by the user.
2. Do not invent studies, citations, statistics, quotes, DOIs, PMIDs, hyperlinks, or source details.
3. Do not use the labels "AI", "AI-suggested", "AI recommendation", or "evidence packet" anywhere in the output.
4. Refer to the evidence source material as "literature evidence" or "supporting literature evidence."
5. "Update" means revise, clarify, replace, remove, or ADD new guideline text when the current guideline does not answer the evidence question.
6. The proposed update or new section must be explicitly based on the literature evidence provided.
7. Always provide concrete proposed guideline text. Do not ask clinicians to create guidance from scratch.
8. Clinician review should be framed as Approve / Modify / Reject.
9. Include the section "Decision Required Before Approval" only when a true local decision is needed.
10. If a decision is required, include a recommended decision based on the evidence, not an open-ended question.
11. Clearly label confidence exactly as "GRADE-informed Confidence Rating."
12. The GRADE-informed Confidence Rating must be based on evidence type, consistency, directness, precision, pediatric applicability, and local implementation considerations. Do not make the confidence rating arbitrary.
13. Do not include "Proposed Algorithm Impact."
14. Do not include "Draft Change Log Entry."
15. This card is focused on guideline text review. If a separate algorithm/pathway update may be needed, mention it briefly only in Evidence Traceback Notes as a future separate review item.
16. Maintain the guideline topic as the unit of work. The focused topic is one evidence review area within the larger guideline update.
17. The output should be readable for busy physicians. Keep it structured, skimmable, and decision-oriented.
18. Evidence Finding may be more detailed than the other sections because it explains the rationale for the proposed change.
19. Evidence Traceback Notes should include all relevant sources from the literature evidence, with complete citations and hyperlinks/DOIs/PMIDs when available from the source material.
20. Do not create unnecessary clinician questions. If the literature evidence supports a clear recommendation, state the recommended decision and allow Approve / Modify / Reject.
21. The Proposed Guideline Edit or New Section must provide actual proposed wording.
22. In Section 5, always provide implementation placement guidance. Specify whether the proposed update should replace existing text, add a clarifying sentence, add a short subsection, create a new section, or create a new addendum.
23. Do not force an addendum. Recommend an addendum only when the proposed content is long, criteria-based, operational, stepwise, or likely to clutter the main body of the guideline.
24. If an addendum is recommended, provide a suggested addendum title and both brief main-body cross-reference text and detailed addendum text.
25. If an addendum is not recommended, briefly state why the content is better placed in the main guideline body.
26. Use clear formatting with concise paragraphs, bullets, numbered steps, and bolded key labels or conclusions when helpful.
27. Avoid unnecessary verbosity, but do not omit important evidence rationale or citations.

Output exactly these sections:

# Change Card Title

## 1. Evidence Question

## 2. Current Guidance

## 3. Workgroup Concern / Reason for Review

## 4. Evidence Finding

## 5. Proposed Guideline Edit or New Section

## 6. GRADE-informed Confidence Rating

## 7. Decision Required Before Approval

## 8. Evidence Traceback Notes
"""


# ============================================================
# Main app layout
# ============================================================

st.title(APP_TITLE)

st.write(
    "Generate a physician-facing, evidence-based Change Card for guideline revision review."
)

st.info(
    "This prototype is intended for non-PHI guideline materials only. "
    "Final clinical approval remains human-led."
)

api_key = get_api_key()

with st.sidebar:
    st.header("Status")

    if api_key:
        st.success("API key: available")
    else:
        st.error("API key: missing")

    st.caption(
        "In the hosted app, the API key should be stored in Streamlit Secrets. "
        "Temporary manual entry is only for testing."
    )


# ============================================================
# User input fields
# ============================================================

st.header("1. Guideline Update Details")

col1, col2 = st.columns(2)

with col1:
    guideline_name = st.text_input(
        "Guideline name",
        value="Bronchiolitis",
        help="Example: Bronchiolitis",
    )

    focused_topic = st.text_input(
        "Focused topic",
        value="",
        help="Example: Feeding, NG Feeds, and IV Fluids",
    )

with col2:
    last_updated_date = st.text_input(
        "Guideline/pathway last updated date",
        value=DEFAULT_LAST_UPDATED_DATE,
        help="Example: October 8, 2019",
    )

st.write("Evidence question")

evidence_question = st.text_area(
    "Paste the question being answered",
    value="",
    height=130,
    help="Paste the PICO or focused evidence question here.",
)


# ============================================================
# File uploads
# ============================================================

st.header("2. Upload Documents")

st.write(
    "Upload the current guideline materials, workgroup notes, and literature evidence."
)

upload_col1, upload_col2 = st.columns(2)

with upload_col1:
    guideline_file = st.file_uploader(
        "Current guideline draft (.docx or .pdf)",
        type=["docx", "pdf"],
    )

    pathway_file = st.file_uploader(
        "Current algorithm/pathway (.docx or .pdf)",
        type=["docx", "pdf"],
    )

    notes_file = st.file_uploader(
        "Workgroup notes / meeting minutes (.docx or .pdf)",
        type=["docx", "pdf"],
    )

with upload_col2:
    literature_file = st.file_uploader(
        "Literature evidence from OpenEvidence or other source (.docx or .pdf)",
        type=["docx", "pdf"],
    )

    optional_context_file = st.file_uploader(
        "Optional local pathway/context document (.docx or .pdf)",
        type=["docx", "pdf"],
    )


# ============================================================
# Status checks
# ============================================================

st.header("3. Readiness Check")

required_items = {
    "API key": bool(api_key),
    "Guideline name": bool(guideline_name.strip()),
    "Focused topic": bool(focused_topic.strip()),
    "Evidence question": bool(evidence_question.strip()),
    "Last updated date": bool(last_updated_date.strip()),
    "Current guideline draft": guideline_file is not None,
    "Current algorithm/pathway": pathway_file is not None,
    "Workgroup notes": notes_file is not None,
    "Literature evidence": literature_file is not None,
}

ready = True

for item, is_ready in required_items.items():
    if is_ready:
        st.success(f"{item}: ready")
    else:
        st.warning(f"{item}: missing")
        ready = False

if optional_context_file is not None:
    st.success("Optional local context: uploaded")
else:
    st.info("Optional local context: not uploaded")


# ============================================================
# Generate card
# ============================================================

st.header("4. Generate Change Card")

if not ready:
    st.warning("Upload all required documents and complete all required fields before generating the card.")

generate = st.button(
    "Generate Physician-Facing Change Card",
    disabled=not ready,
    type="primary",
)

if generate:
    try:
        with st.spinner("Reading uploaded documents..."):
            guideline_text = read_uploaded_file(guideline_file)
            pathway_text = read_uploaded_file(pathway_file)
            notes_text = read_uploaded_file(notes_file)
            literature_evidence_text = read_uploaded_file(literature_file)

            optional_context_text = ""
            if optional_context_file is not None:
                optional_context_text = read_uploaded_file(optional_context_file)

        with st.expander("Document text extraction summary", expanded=False):
            st.write(f"Guideline characters: {len(guideline_text)}")
            st.write(f"Pathway characters: {len(pathway_text)}")
            st.write(f"Workgroup notes characters: {len(notes_text)}")
            st.write(f"Literature evidence characters: {len(literature_evidence_text)}")
            st.write(f"Optional context characters: {len(optional_context_text)}")

        local_context_text = f"""
CURRENT ALGORITHM/PATHWAY:
{pathway_text}

OPTIONAL LOCAL CONTEXT:
{optional_context_text}
"""

        user_input = f"""
Guideline name:
{guideline_name}

Focused topic:
{focused_topic}

Evidence question / focused review area:
{evidence_question}

Last Updated date for the current guideline/algorithm:
{last_updated_date}

Current algorithm/pathway text and related local pathway context:
{clip_text(local_context_text)}

Current written guideline draft text:
{clip_text(guideline_text)}

Literature evidence:
{clip_text(literature_evidence_text)}

Workgroup notes/context:
{clip_text(notes_text)}
"""

        with st.spinner("Generating physician-facing Change Card..."):
            client = OpenAI(api_key=api_key)

            response = client.responses.create(
                model=MODEL_NAME,
                instructions=PHYSICIAN_CARD_PROMPT,
                input=user_input,
            )

            physician_card = response.output_text

        st.session_state["physician_card"] = physician_card
        st.session_state["focused_topic"] = focused_topic

        st.success("Change Card generated successfully.")

    except Exception as e:
        st.error("The Change Card could not be generated.")
        st.write("Error details:")
        st.code(str(e))


# ============================================================
# Output preview and download
# ============================================================

if "physician_card" in st.session_state:
    st.header("5. Preview and Download")

    physician_card = st.session_state["physician_card"]
    topic_for_file = st.session_state.get("focused_topic", "Change Card")

    st.subheader("Preview")
    st.markdown(physician_card)

    docx_file = create_styled_docx(physician_card, topic_for_file)

    safe_topic = re.sub(r"[^A-Za-z0-9]+", "_", topic_for_file).strip("_")
    if not safe_topic:
        safe_topic = "Change_Card"

    filename = f"{guideline_name}_{safe_topic}_Physician_Facing_Change_Card.docx"
    filename = re.sub(r"[^A-Za-z0-9_.-]+", "_", filename)

    st.download_button(
        label="Download Word Document",
        data=docx_file,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    st.caption(
        "Review the generated card before sharing. Final guideline approval remains human-led."
    )
