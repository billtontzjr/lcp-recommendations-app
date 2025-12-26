"""Word document generation service."""
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from datetime import datetime
from app.utils.currency import format_currency

# Constants
GRAY_BACKGROUND = "D3D3D3"
BLACK_BORDER = "000000"
FONT_NAME = "Times New Roman"
DATA_FONT_SIZE = Pt(10)
TITLE_FONT_SIZE = Pt(11)
SMALL_FONT_SIZE = Pt(9)


def set_cell_shading(cell, color):
    """Set cell background color."""
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color)
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_borders(cell, border_size=12, color="000000"):
    """Set cell borders."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    tcBorders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), str(border_size))
        border.set(qn('w:color'), color)
        tcBorders.append(border)

    tcPr.append(tcBorders)


def format_cell(cell, text, font_size=DATA_FONT_SIZE, bold=False, center=False, gray=False):
    """Format a table cell with text and styling."""
    cell.text = str(text) if text else ''
    paragraph = cell.paragraphs[0]
    run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()

    run.font.name = FONT_NAME
    run.font.size = font_size
    run.font.bold = bold

    if center:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if gray:
        set_cell_shading(cell, GRAY_BACKGROUND)

    set_cell_borders(cell)


def format_date(date_val):
    """Format date for display."""
    if not date_val:
        return ''
    if isinstance(date_val, datetime):
        return date_val.strftime('%B %d, %Y')
    if isinstance(date_val, str):
        return date_val
    return str(date_val)


def generate_lcp_document(patient_info, cost_data, output_path):
    """
    Generate the Life Care Plan Recommendations Word document.

    Args:
        patient_info: Patient information dictionary
        cost_data: Cost calculation results dictionary
        output_path: Path to save the generated document

    Returns:
        Path to the generated document
    """
    doc = Document()

    # Set default font
    style = doc.styles['Normal']
    style.font.name = FONT_NAME
    style.font.size = DATA_FONT_SIZE

    # Page 1: Title Page
    add_title_page(doc, patient_info)

    # Page 2: Appendix A Summary
    doc.add_page_break()
    add_appendix_a(doc, patient_info, cost_data)

    # Section Pages (one per category)
    for category, data in cost_data['category_totals'].items():
        doc.add_page_break()
        add_section_page(doc, patient_info, category, data, cost_data['totals'])

    doc.save(output_path)
    return output_path


def add_title_page(doc, patient_info):
    """Add the title page."""
    # Main title
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run('LIFE CARE PLAN RECOMMENDATIONS')
    run.font.name = FONT_NAME
    run.font.size = Pt(16)
    run.font.bold = True

    doc.add_paragraph()

    # Patient name
    name_para = doc.add_paragraph()
    name_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = name_para.add_run(patient_info.get('patient_name', ''))
    run.font.name = FONT_NAME
    run.font.size = Pt(14)
    run.font.bold = True

    doc.add_paragraph()

    # Patient details
    details = [
        ('Date of Birth:', format_date(patient_info.get('date_of_birth'))),
        ('Date of Injury:', format_date(patient_info.get('date_of_injury'))),
        ('Date of Report:', format_date(patient_info.get('date_of_report'))),
        ('Life Expectancy:', f"{patient_info.get('life_expectancy', '')} years"),
    ]

    for label, value in details:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(f'{label} ')
        run.font.name = FONT_NAME
        run.font.size = Pt(11)
        run = para.add_run(str(value))
        run.font.name = FONT_NAME
        run.font.size = Pt(11)

    doc.add_paragraph()
    doc.add_paragraph()

    # Prepared by
    prep = doc.add_paragraph()
    prep.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = prep.add_run('Prepared by:')
    run.font.name = FONT_NAME
    run.font.size = Pt(11)
    run.font.bold = True

    doc.add_paragraph()

    credentials = [
        'William Tontz, Jr., MD',
        'Board-Certified Orthopedic Surgeon',
        'Certified Life Care Planner (CLCP)',
        '',
        'Precision Life Care Planning',
    ]

    for line in credentials:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(line)
        run.font.name = FONT_NAME
        run.font.size = Pt(11)


def add_appendix_a(doc, patient_info, cost_data):
    """Add Appendix A summary page."""
    # Header
    header = doc.add_paragraph()
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = header.add_run('APPENDIX A')
    run.font.name = FONT_NAME
    run.font.size = Pt(14)
    run.font.bold = True

    doc.add_paragraph()

    # Patient info block
    info_lines = [
        f"Patient: {patient_info.get('patient_name', '')}",
        f"Date of Birth: {format_date(patient_info.get('date_of_birth'))}",
        f"Date of Injury: {format_date(patient_info.get('date_of_injury'))}",
        f"Life Expectancy: {patient_info.get('life_expectancy', '')} years",
    ]

    for line in info_lines:
        para = doc.add_paragraph(line)
        para.runs[0].font.name = FONT_NAME
        para.runs[0].font.size = Pt(10)

    doc.add_paragraph()

    # Summary table
    category_totals = cost_data['category_totals']
    totals = cost_data['totals']

    # Create table: header + categories + separator
    num_rows = 2 + len(category_totals) + 1  # title, headers, data rows, separator
    table = doc.add_table(rows=num_rows, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Set column widths
    for row in table.rows:
        row.cells[0].width = Inches(3.5)
        row.cells[1].width = Inches(1.5)
        row.cells[2].width = Inches(1.5)

    # Row 0: Title
    cell = table.rows[0].cells[0]
    cell.merge(table.rows[0].cells[2])
    format_cell(cell, 'Lifetime Projected Costs', TITLE_FONT_SIZE, bold=True, center=True, gray=True)

    # Row 1: Headers
    headers = ['Section', 'Annual Cost', 'One Time Cost']
    for i, header_text in enumerate(headers):
        format_cell(table.rows[1].cells[i], header_text, DATA_FONT_SIZE, bold=True, center=True, gray=True)

    # Data rows
    row_idx = 2
    for category, data in category_totals.items():
        format_cell(table.rows[row_idx].cells[0], category, DATA_FONT_SIZE)
        format_cell(table.rows[row_idx].cells[1], format_currency(data['annual_cost']), DATA_FONT_SIZE, center=True)
        format_cell(table.rows[row_idx].cells[2], format_currency(data['one_time_cost']), DATA_FONT_SIZE, center=True)
        row_idx += 1

    # Separator row
    for i in range(3):
        format_cell(table.rows[row_idx].cells[i], '', DATA_FONT_SIZE, gray=True)

    doc.add_paragraph()

    # Totals table
    totals_table = doc.add_table(rows=4, cols=2)
    totals_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for row in totals_table.rows:
        row.cells[0].width = Inches(3.75)
        row.cells[1].width = Inches(2.75)

    totals_data = [
        ('Total Annual Cost:', format_currency(totals['total_annual'])),
        (f"Lifetime Annual Cost ({totals['life_expectancy']} years):", format_currency(totals['lifetime_annual'])),
        ('Total One-Time Cost:', format_currency(totals['total_one_time'])),
        ('GRAND TOTAL:', format_currency(totals['grand_total'])),
    ]

    for i, (label, value) in enumerate(totals_data):
        format_cell(totals_table.rows[i].cells[0], label, DATA_FONT_SIZE, bold=(i == 3))
        format_cell(totals_table.rows[i].cells[1], value, DATA_FONT_SIZE, bold=(i == 3), center=True)


def add_section_page(doc, patient_info, category, category_data, totals):
    """Add a section page with data table, rationale, and sources."""
    # Patient info header
    info_line = f"Patient: {patient_info.get('patient_name', '')} | DOB: {format_date(patient_info.get('date_of_birth'))} | Life Expectancy: {patient_info.get('life_expectancy', '')} years"
    para = doc.add_paragraph(info_line)
    para.runs[0].font.name = FONT_NAME
    para.runs[0].font.size = Pt(9)

    doc.add_paragraph()

    items = category_data['items']

    # Data table (7 columns)
    num_rows = 2 + len(items) + 1  # title, headers, data rows, separator
    table = doc.add_table(rows=num_rows, cols=7)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Column widths
    col_widths = [2.0, 0.7, 0.7, 0.8, 1.0, 0.9, 0.9]
    for row in table.rows:
        for i, width in enumerate(col_widths):
            row.cells[i].width = Inches(width)

    # Row 0: Category title
    cell = table.rows[0].cells[0]
    cell.merge(table.rows[0].cells[6])
    format_cell(cell, category, TITLE_FONT_SIZE, bold=True, center=True, gray=True)

    # Row 1: Headers
    headers = ['Item', 'Age Init', 'Until Age', 'Cost', 'Frequency', 'Annual Cost', 'One Time']
    for i, header_text in enumerate(headers):
        format_cell(table.rows[1].cells[i], header_text, DATA_FONT_SIZE, bold=True, center=True, gray=True)

    # Data rows
    age_init = patient_info.get('age_initiated', '')
    until_age = patient_info.get('until_age', '')

    for row_idx, item in enumerate(items, start=2):
        item_name = item.get('item', '') or item.get('service_description', '')
        format_cell(table.rows[row_idx].cells[0], item_name, DATA_FONT_SIZE)
        format_cell(table.rows[row_idx].cells[1], age_init, DATA_FONT_SIZE, center=True)
        format_cell(table.rows[row_idx].cells[2], until_age, DATA_FONT_SIZE, center=True)
        format_cell(table.rows[row_idx].cells[3], format_currency(item.get('unit_cost', 0)), DATA_FONT_SIZE, center=True)
        format_cell(table.rows[row_idx].cells[4], item.get('frequency', ''), DATA_FONT_SIZE, center=True)
        format_cell(table.rows[row_idx].cells[5], format_currency(item.get('annual_cost', 0)), DATA_FONT_SIZE, center=True)
        format_cell(table.rows[row_idx].cells[6], format_currency(item.get('one_time_cost', 0)), DATA_FONT_SIZE, center=True)

    # Separator row
    last_row = len(items) + 2
    for i in range(7):
        format_cell(table.rows[last_row].cells[i], '', DATA_FONT_SIZE, gray=True)

    doc.add_paragraph()

    # Rationale table
    rationales = [item.get('rationale', '') for item in items if item.get('rationale')]
    rationale_text = ' '.join(rationales) if rationales else ''

    rationale_table = doc.add_table(rows=2, cols=1)
    rationale_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    rationale_table.rows[0].cells[0].width = Inches(7.0)
    rationale_table.rows[1].cells[0].width = Inches(7.0)

    format_cell(rationale_table.rows[0].cells[0], 'Rationale', DATA_FONT_SIZE, bold=True, gray=True)
    format_cell(rationale_table.rows[1].cells[0], rationale_text, SMALL_FONT_SIZE)

    doc.add_paragraph()

    # Sources table
    sources = [item.get('source', '') for item in items if item.get('source')]
    sources_text = '; '.join(set(sources)) if sources else ''

    sources_table = doc.add_table(rows=2, cols=1)
    sources_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sources_table.rows[0].cells[0].width = Inches(7.0)
    sources_table.rows[1].cells[0].width = Inches(7.0)

    format_cell(sources_table.rows[0].cells[0], 'Sources', DATA_FONT_SIZE, bold=True, gray=True)
    format_cell(sources_table.rows[1].cells[0], sources_text, SMALL_FONT_SIZE)
