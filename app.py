import os
import re
import pythoncom
from flask import Flask, render_template, request, jsonify, send_from_directory
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx2pdf import convert
from google import genai
from google.genai import types

app = Flask(__name__)

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "generated_files")
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.docx")
os.makedirs(OUTPUT_DIR, exist_ok=True)

FONT_FAMILY = "Arial"


def build_human_academic_prompt(question: str) -> str:
    """Engineered to match the exact 4-page layout and density of Week7_ECS.pdf (~1,200-1,350 words)."""
    return f"""Write a comprehensive, formal, human-like academic response for the following university Long Answer Question (LAQ):

"{question}"

STRICT STRUCTURAL & LENGTH REQUIREMENTS (TARGET: 1,200 TO 1,350 WORDS / EXACTLY 4 PAGES):
1. Opening Definition & Core Framing (2 paragraphs):
   - Paragraph 1: Rigorous academic definition with scholarly citation in (Author, Year) format.
   - Paragraph 2: Operational reality explaining how this functions in practice (3-4 sentences).
2. Core Pillars / Foundational Dimensions:
   - Heading: ### Core Pillars of [Topic]
   - Provide 2 or 3 numbered items on single paragraphs (1. **Title:** Explanation... 2. **Title:** Explanation...). Do NOT use subheadings for these pillars.
3. Strategic Importance in Contemporary Business:
   - Heading: ### Importance of [Topic] in Contemporary Business
   - Write 1 short introductory sentence ending with a colon.
   - Provide 4 bullet points (* **Bold Lead-in Title:** 2-3 concise sentences with concrete corporate examples).
4. Comparative Analysis Table:
   - Heading: ### Comparative Analysis: [Topic] vs. Alternative Frameworks
   - Create a compact 5-column Markdown table with 1 header row and 3 to 4 comparison rows (Dimension | Framework 1 | Framework 2 | Framework 3 | Framework 4). Keep cell text concise (8 to 14 words per cell) so the table fits cleanly on one page.
5. Actionable Implementation Strategies:
   - Heading: ### How Leaders Effectively Harness [Topic] Within Organizations
   - Write 1 short introductory sentence ending with a colon.
   - Provide 5 numbered subheadings (#### 1. Strategy Title ... through #### 5. Strategy Title).
   - Under EACH numbered subheading, write 1 short intro sentence followed by 3 concise bullet points (* **Bold Lead-in:** 1-2 sentences).
6. Conclusion:
   - Heading: ### Conclusion
   - Write 2 concise concluding paragraphs (3 sentences each) synthesizing the strategic takeaway.

WRITING STYLE RULES:
- Write in a natural, authoritative academic tone.
- Avoid AI clichés ("delve", "tapestry", "testament", "in today's fast-paced world").
- Do NOT include "Week #" or repeat the question at the start of your output.
- Output clean Markdown."""


def set_cell_border_and_padding(cell, is_header=False):
    """Applies clean thin borders, cell padding (margins), and header shading matching Week7_ECS.pdf."""
    tcPr = cell._tc.get_or_add_tcPr()

    # 1. Borders
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        element = OxmlElement(f'w:{edge}')
        element.set(qn('w:val'), 'single')
        element.set(qn('w:sz'), '4')
        element.set(qn('w:space'), '0')
        element.set(qn('w:color'), 'D0D5DD')
        tcBorders.append(element)
    tcPr.append(tcBorders)

    # 2. Internal Cell Padding (Top/Bottom: 6pt, Left/Right: 8pt in dxa units)
    tcMar = OxmlElement('w:tcMar')
    for side, dxa in (('top', '110'), ('bottom', '110'), ('left', '150'), ('right', '150')):
        node = OxmlElement(f'w:{side}')
        node.set(qn('w:w'), dxa)
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

    # 3. Header Row Background Shading
    if is_header:
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), 'F2F4F7')
        tcPr.append(shd)


def prevent_row_split(row, is_header=False):
    """Prevents Word from cutting a table row in half across a page break."""
    trPr = row._tr.get_or_add_trPr()
    cantSplit = OxmlElement('w:cantSplit')
    trPr.append(cantSplit)
    if is_header:
        tblHeader = OxmlElement('w:tblHeader')
        trPr.append(tblHeader)


def add_formatted_runs(paragraph, text, font_size=10.5, color_rgb=(30, 30, 30)):
    """Parses inline **bold** markdown and applies consistent font family, size, and color."""
    parts = re.split(r'(\*\*.*?\*\*)', text)
    for part in parts:
        if not part:
            continue
        if part.startswith('**') and part.endswith('**'):
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        else:
            run = paragraph.add_run(part)
        run.font.name = FONT_FAMILY
        run.font.size = Pt(font_size)
        run.font.color.rgb = RGBColor(*color_rgb)


def format_base_paragraph(p, space_before=0, space_after=6, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.LEFT):
    """Ensures clean, consistent paragraph spacing and alignment."""
    pf = p.paragraph_format
    pf.alignment = align
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line_spacing
    pf.left_indent = Inches(0)
    pf.right_indent = Inches(0)
    pf.first_line_indent = Inches(0)


def create_laq_docx(week_num: str, question: str, answer_md: str, output_docx_path: str):
    """Generates a Word document matching the exact visual layout of Week7_ECS.pdf."""
    if os.path.exists(TEMPLATE_PATH):
        doc = Document(TEMPLATE_PATH)
        doc._body.clear_content()
    else:
        doc = Document()

    # Enforce uniform 0.85" margins so tables and text align identically to Week7_ECS.pdf
    for section in doc.sections:
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)
        section.top_margin = Inches(0.85)
        section.bottom_margin = Inches(0.85)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)

    # 1. Week Header ("Week 7")
    p_week = doc.add_paragraph()
    format_base_paragraph(p_week, space_before=0, space_after=6, line_spacing=1.1)
    r_week = p_week.add_run(f"Week {week_num}")
    r_week.bold = True
    r_week.font.name = FONT_FAMILY
    r_week.font.size = Pt(16)
    r_week.font.color.rgb = RGBColor(17, 24, 39)

    # 2. Question Header with Clean Bottom Divider Line
    p_question = doc.add_paragraph()
    format_base_paragraph(p_question, space_before=0, space_after=12, line_spacing=1.2)
    r_question = p_question.add_run(question.strip())
    r_question.bold = True
    r_question.font.name = FONT_FAMILY
    r_question.font.size = Pt(11.5)
    r_question.font.color.rgb = RGBColor(17, 24, 39)

    pPr = p_question._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '6')
    bottom.set(qn('w:color'), 'D0D5DD')
    pBdr.append(bottom)
    pPr.append(pBdr)

    # 3. Parse Markdown Answer
    lines = answer_md.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].strip()

        if not line or line in ('---', '***'):
            i += 1
            continue

        # A. Markdown Tables
        if line.startswith('|') and line.endswith('|'):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                row_str = lines[i].strip()
                if not re.match(r'^\|[\s\-:|]+\|$', row_str):
                    cells = [c.strip() for c in row_str.strip('|').split('|')]
                    table_lines.append(cells)
                i += 1

            if table_lines:
                num_cols = max(len(r) for r in table_lines)
                table = doc.add_table(rows=len(table_lines), cols=num_cols)
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.autofit = False

                # Distribute 6.7 inches total width cleanly across columns
                total_width = 6.7
                col_widths = [total_width / num_cols] * num_cols
                if num_cols >= 4:
                    col_widths[0] = 1.15
                    rem = (total_width - 1.15) / (num_cols - 1)
                    for c_i in range(1, num_cols):
                        col_widths[c_i] = rem

                for r_idx, row_data in enumerate(table_lines):
                    row = table.rows[r_idx]
                    prevent_row_split(row, is_header=(r_idx == 0))

                    for c_idx in range(num_cols):
                        cell_text = row_data[c_idx] if c_idx < len(row_data) else ""
                        cell = row.cells[c_idx]
                        cell.width = Inches(col_widths[c_idx])
                        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
                        set_cell_border_and_padding(cell, is_header=(r_idx == 0))

                        cell_p = cell.paragraphs[0]
                        format_base_paragraph(cell_p, space_before=1, space_after=1, line_spacing=1.1, align=WD_ALIGN_PARAGRAPH.LEFT)

                        if r_idx == 0:
                            clean_header = cell_text.replace('**', '')
                            run = cell_p.add_run(clean_header)
                            run.bold = True
                            run.font.name = FONT_FAMILY
                            run.font.size = Pt(9.5)
                            run.font.color.rgb = RGBColor(17, 24, 39)
                        else:
                            add_formatted_runs(cell_p, cell_text, font_size=9.5)

                spacer = doc.add_paragraph()
                format_base_paragraph(spacer, space_before=0, space_after=4)
            continue

        # B. Headings (##, ###, ####)
        if line.startswith('#'):
            level = len(line.split(' ')[0])
            heading_text = line.lstrip('#').strip().replace('**', '')
            p_head = doc.add_paragraph()
            p_head.paragraph_format.keep_with_next = True

            if level <= 3:
                # Major Section Heading (e.g. "Core Pillars of...", "Comparative Analysis...")
                format_base_paragraph(p_head, space_before=11, space_after=5, line_spacing=1.15)
                r_head = p_head.add_run(heading_text)
                r_head.bold = True
                r_head.font.name = FONT_FAMILY
                r_head.font.size = Pt(12.5)
                r_head.font.color.rgb = RGBColor(17, 24, 39)
            else:
                # Numbered Subheading (e.g. "1. Cultivating Effectual Reasoning...")
                format_base_paragraph(p_head, space_before=9, space_after=3, line_spacing=1.15)
                r_head = p_head.add_run(heading_text)
                r_head.bold = True
                r_head.font.name = FONT_FAMILY
                r_head.font.size = Pt(11)
                r_head.font.color.rgb = RGBColor(17, 24, 39)
            i += 1
            continue

        # C. Bullet Points (* or -) with Real Bullet Dot (•) & True Hanging Indent
        if re.match(r'^[\*\-•]\s+', line):
            bullet_text = re.sub(r'^[\*\-•]\s+', '', line)
            p_bul = doc.add_paragraph()
            format_base_paragraph(p_bul, space_before=2, space_after=4, line_spacing=1.15)
            p_bul.paragraph_format.left_indent = Inches(0.32)
            p_bul.paragraph_format.first_line_indent = Inches(-0.18)

            r_dot = p_bul.add_run("•  ")
            r_dot.bold = True
            r_dot.font.name = FONT_FAMILY
            r_dot.font.size = Pt(10.5)

            add_formatted_runs(p_bul, bullet_text, font_size=10.5)
            i += 1
            continue

        # D. Numbered List Items (1. **The Moral Person:** ...)
        num_match = re.match(r'^(\d+\.)\s+(.*)', line)
        if num_match:
            num_prefix, rest_text = num_match.group(1), num_match.group(2)
            p_num = doc.add_paragraph()

            # Check if next line is a paragraph (meaning LLM wrote a numbered heading without ###)
            next_line = lines[i + 1].strip() if (i + 1) < len(lines) else ""
            is_standalone_subhead = (len(rest_text) < 85 and not (':' in rest_text and len(rest_text) > 60) and next_line and not re.match(r'^(\d+\.|[\*\-•])\s+', next_line))

            if is_standalone_subhead:
                p_num.paragraph_format.keep_with_next = True
                format_base_paragraph(p_num, space_before=9, space_after=3, line_spacing=1.15)
                clean_sub = f"{num_prefix} {rest_text.replace('**', '')}"
                r_sub = p_num.add_run(clean_sub)
                r_sub.bold = True
                r_sub.font.name = FONT_FAMILY
                r_sub.font.size = Pt(11)
                r_sub.font.color.rgb = RGBColor(17, 24, 39)
            else:
                format_base_paragraph(p_num, space_before=2, space_after=5, line_spacing=1.15)
                p_num.paragraph_format.left_indent = Inches(0.32)
                p_num.paragraph_format.first_line_indent = Inches(-0.22)

                r_num = p_num.add_run(f"{num_prefix} ")
                r_num.bold = True
                r_num.font.name = FONT_FAMILY
                r_num.font.size = Pt(10.5)
                add_formatted_runs(p_num, rest_text, font_size=10.5)
            i += 1
            continue

        # E. Standard Body Paragraph
        p = doc.add_paragraph()
        format_base_paragraph(p, space_before=2, space_after=6, line_spacing=1.15)
        add_formatted_runs(p, line, font_size=10.5)
        i += 1

    doc.save(output_docx_path)


@app.route("/")
def index():
    has_template = os.path.exists(TEMPLATE_PATH)
    return render_template("index.html", has_template=has_template)


@app.route("/api/generate-answer", methods=["POST"])
def generate_answer():
    import time
    data = request.json
    api_key = data.get("api_key", "").strip()
    question = data.get("question", "").strip()
    requested_model = data.get("model", "gemini-3-flash-preview")

    if not api_key:
        return jsonify({"error": "Please enter your Gemini API Key at the top, or use the 'Copy Prompt for Gemini Pro' button."}), 400
    if not question:
        return jsonify({"error": "Question cannot be empty."}), 400

    # Fallback queue of free-tier models if one has a 503 traffic spike or 429 quota limit
    fallback_models = [
        requested_model,
        "gemini-2.5-flash",
        "gemini-3-flash-preview",
        "gemini-3.1-flash-lite-preview",
        "gemini-2.0-flash"
    ]
    models_to_try = list(dict.fromkeys(fallback_models))

    client = genai.Client(api_key=api_key)
    prompt = build_human_academic_prompt(question)
    last_error = ""

    for model_name in models_to_try:
        for attempt in range(2):  # Try up to 2 times per model on 503 spikes
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        temperature=0.65,
                    )
                )
                return jsonify({
                    "answer": response.text,
                    "model_used": model_name
                })
            except Exception as e:
                last_error = str(e)
                # If 503 High Demand, wait 1.5s and retry once before moving to the next model
                if "503" in last_error or "UNAVAILABLE" in last_error:
                    time.sleep(1.5)
                    continue
                # If 429 Quota or 404 Not Found, immediately try the next model in the list
                elif any(code in last_error for code in ("429", "RESOURCE_EXHAUSTED", "404", "NOT_FOUND", "500")):
                    break
                else:
                    return jsonify({"error": last_error}), 500

    return jsonify({"error": f"All models are currently busy. Last error: {last_error}"}), 500

@app.route("/api/export", methods=["POST"])
def export_files():
    data = request.json
    week_num = str(data.get("week", "1")).strip()
    course_code = data.get("course_code", "LAQ").strip().upper()
    question = data.get("question", "").strip()
    answer = data.get("answer", "").strip()

    if not question or not answer:
        return jsonify({"error": "Both Question and Answer are required to generate Word & PDF files."}), 400

    base_filename = f"Week{week_num}_{course_code}" if course_code else f"Week{week_num}"
    docx_filename = f"{base_filename}.docx"
    pdf_filename = f"{base_filename}.pdf"

    docx_path = os.path.join(OUTPUT_DIR, docx_filename)
    pdf_path = os.path.join(OUTPUT_DIR, pdf_filename)

    try:
        create_laq_docx(week_num, question, answer, docx_path)

        pythoncom.CoInitialize()
        try:
            convert(docx_path, pdf_path)
        finally:
            pythoncom.CoUninitialize()

        return jsonify({
            "success": True,
            "docx_url": f"/download/{docx_filename}",
            "pdf_url": f"/download/{pdf_filename}",
            "docx_filename": docx_filename,
            "pdf_filename": pdf_filename
        })
    except Exception as e:
        return jsonify({"error": f"Export failed: {str(e)}"}), 500


@app.route("/download/<path:filename>")
def download_file(filename):
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True, port=5000)