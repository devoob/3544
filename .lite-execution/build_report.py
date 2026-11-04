from pathlib import Path

from PIL import Image
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "Group_Project_Report.docx"
FIGURE_SOURCE = Path("/tmp/elec3544_report_assets/cell29_0.png")
FIGURE_CROP = Path("/tmp/elec3544_report_assets/model_results_crop.png")

NAVY = "17324D"
INK = "17222B"
MUTED = "5A6570"
LIGHT_RULE = "D9D9D9"
ROW_TINT = "F3F6F8"
PALE_BLUE = "E8EFF5"
WHITE = "FFFFFF"

BODY_FONT = "Calibri"
BODY_SIZE = 10.5


def set_run_font(run, name=BODY_FONT, size=None, bold=None, color=INK):
    run.font.name = name
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), name)
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), name)
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), name)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if color is not None:
        run.font.color.rgb = RGBColor.from_string(color)


def set_paragraph_spacing(paragraph, before=0, after=6, line=1.12):
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.line_spacing = line
    fmt.widow_control = True


def set_cell_margins(cell, top=90, start=105, bottom=90, end=105):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def shade_cell(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")


def set_cell_borders(cell, color=LIGHT_RULE, size="6"):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), size)
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), color)


def mark_repeat_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)
    cant_split = OxmlElement("w:cantSplit")
    tr_pr.append(cant_split)


def remove_paragraph_borders(paragraph_or_style):
    p_pr = paragraph_or_style._element.get_or_add_pPr()
    for border in p_pr.findall(qn("w:pBdr")):
        p_pr.remove(border)


def prevent_row_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:cantSplit")) is None:
        tr_pr.append(OxmlElement("w:cantSplit"))


def set_table_width(table, widths):
    table.autofit = False
    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = Inches(width)
    grid = table._tbl.tblGrid
    for idx, width in enumerate(widths):
        if idx < len(grid.gridCol_lst):
            grid.gridCol_lst[idx].set(qn("w:w"), str(int(width * 1440)))


def add_table_caption(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text)
    set_run_font(r, size=8.5, bold=True, color=NAVY)


def add_table(doc, headers, rows, widths, font_size=9.0, numeric=None, highlight_rows=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    set_table_width(table, widths)

    header = table.rows[0]
    mark_repeat_header(header)
    for idx, text in enumerate(headers):
        cell = header.cells[idx]
        shade_cell(cell, NAVY)
        set_cell_margins(cell)
        set_cell_borders(cell)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(p, 0, 0, 1.0)
        r = p.add_run(str(text))
        set_run_font(r, size=font_size, bold=True, color=WHITE)
        p.paragraph_format.keep_with_next = True

    numeric = numeric or set()
    highlight_rows = highlight_rows or set()
    for row_idx, values in enumerate(rows):
        row = table.add_row()
        prevent_row_split(row)
        for idx, value in enumerate(values):
            cell = row.cells[idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            shade_cell(cell, PALE_BLUE if row_idx in highlight_rows else (WHITE if row_idx % 2 == 0 else ROW_TINT))
            set_cell_margins(cell)
            set_cell_borders(cell)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if idx not in numeric else WD_ALIGN_PARAGRAPH.CENTER
            set_paragraph_spacing(p, 0, 0, 1.0)
            r = p.add_run(str(value))
            set_run_font(r, size=font_size, bold=(row_idx in highlight_rows), color=INK)
            if row_idx < len(rows) - 1:
                p.paragraph_format.keep_with_next = True
    set_table_width(table, widths)
    return table


def add_body(doc, text, space_after=6, justify=False):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if justify else WD_ALIGN_PARAGRAPH.LEFT
    set_paragraph_spacing(p, after=space_after)
    r = p.add_run(text)
    set_run_font(r, size=BODY_SIZE, color=INK)
    return p


def add_h1(doc, text, page_break=False):
    if page_break:
        doc.add_page_break()
    p = doc.add_paragraph(style="Heading 1")
    p.paragraph_format.keep_with_next = True
    set_paragraph_spacing(p, before=4, after=8, line=1.0)
    r = p.add_run(text)
    set_run_font(r, size=16, bold=True, color=INK)
    return p


def add_h2(doc, text):
    p = doc.add_paragraph(style="Heading 2")
    p.paragraph_format.keep_with_next = True
    set_paragraph_spacing(p, before=10, after=4, line=1.0)
    r = p.add_run(text)
    set_run_font(r, size=12, bold=True, color=INK)
    return p


def add_figure_caption(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_paragraph_spacing(p, before=3, after=9, line=1.0)
    r = p.add_run(text)
    set_run_font(r, size=8.5, color=MUTED)


def add_page_field(paragraph):
    run = paragraph.add_run()
    fld_char_1 = OxmlElement("w:fldChar")
    fld_char_1.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "PAGE"
    fld_char_2 = OxmlElement("w:fldChar")
    fld_char_2.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    fld_char_3 = OxmlElement("w:fldChar")
    fld_char_3.set(qn("w:fldCharType"), "end")
    run._r.extend([fld_char_1, instr_text, fld_char_2, text, fld_char_3])
    set_run_font(run, size=8.5, color=MUTED)


def setup_document():
    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.72)
    section.bottom_margin = Inches(0.68)
    section.left_margin = Inches(0.82)
    section.right_margin = Inches(0.82)
    section.header_distance = Inches(0.32)
    section.footer_distance = Inches(0.32)
    section.different_first_page_header_footer = True

    normal = doc.styles["Normal"]
    normal.font.name = BODY_FONT
    normal._element.rPr.rFonts.set(qn("w:ascii"), BODY_FONT)
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), BODY_FONT)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)
    normal.font.size = Pt(BODY_SIZE)
    normal.font.color.rgb = RGBColor.from_string(INK)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.12
    normal.paragraph_format.widow_control = True

    title = doc.styles["Title"]
    title.font.name = BODY_FONT
    title._element.rPr.rFonts.set(qn("w:ascii"), BODY_FONT)
    title._element.rPr.rFonts.set(qn("w:hAnsi"), BODY_FONT)
    title.font.size = Pt(28)
    title.font.bold = True
    title.font.color.rgb = RGBColor.from_string(INK)
    title.paragraph_format.space_before = Pt(0)
    title.paragraph_format.space_after = Pt(10)
    title.paragraph_format.keep_with_next = True

    for style_name, size in (("Heading 1", 16), ("Heading 2", 12)):
        style = doc.styles[style_name]
        style.font.name = BODY_FONT
        style._element.rPr.rFonts.set(qn("w:ascii"), BODY_FONT)
        style._element.rPr.rFonts.set(qn("w:hAnsi"), BODY_FONT)
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(INK)

    header = section.header
    hp = header.paragraphs[0]
    hp.paragraph_format.space_after = Pt(0)
    hp.paragraph_format.tab_stops.add_tab_stop(Inches(6.86), alignment=2)
    r = hp.add_run("ELEC3544 GROUP PROJECT\t\tGROUP 1")
    set_run_font(r, size=8.5, color=MUTED)

    footer = section.footer
    fp = footer.paragraphs[0]
    fp.paragraph_format.space_before = Pt(0)
    fp.paragraph_format.space_after = Pt(0)
    fp.paragraph_format.tab_stops.add_tab_stop(Inches(6.86), alignment=2)
    r = fp.add_run("Predicting Thirty Day Customer Repeat Purchases\t\tPage ")
    set_run_font(r, size=8.5, color=MUTED)
    add_page_field(fp)

    doc.core_properties.title = "Predicting Thirty Day Customer Repeat Purchases"
    doc.core_properties.subject = "ELEC3544 group project report"
    doc.core_properties.author = "Group 1"
    return doc


def build_cover(doc):
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=14, after=8, line=1.0)
    r = p.add_run("ELEC3544 DATA SCIENCE WITH FOUNDATION MODELS")
    set_run_font(r, size=9, bold=True, color=NAVY)

    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=0, after=28, line=1.0)
    r = p.add_run("GROUP PROJECT REPORT")
    set_run_font(r, size=9, bold=True, color=MUTED)

    title = doc.add_paragraph(style="Title")
    remove_paragraph_borders(doc.styles["Title"])
    remove_paragraph_borders(title)
    title.add_run("Predicting Thirty Day Customer Repeat Purchases")

    subtitle = doc.add_paragraph()
    set_paragraph_spacing(subtitle, before=0, after=37, line=1.05)
    subtitle.paragraph_format.keep_with_next = True
    r = subtitle.add_run("Behavioural, cancellation, TF-IDF and MiniLM features")
    set_run_font(r, size=14, color=NAVY)

    lead = doc.add_paragraph()
    lead.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_paragraph_spacing(lead, before=0, after=9, line=1.14)
    r = lead.add_run(
        "We test whether a customer's previous 90 days of transactions can rank who will make "
        "a valid purchase in the next 30 days. The study compares basic RFM features with "
        "basket, timing, cancellation, TF-IDF and pretrained MiniLM product-history features."
    )
    set_run_font(r, size=11, color=INK)

    finding = doc.add_paragraph()
    finding.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_paragraph_spacing(finding, before=0, after=28, line=1.14)
    r = finding.add_run(
        "The main finding is that rich behavioural features improve ranking, while product "
        "semantics add no measurable gain at three decimal places. The July-selected model "
        "reached a final pooled PR-AUC of 0.637 and lift at 20 percent of 1.77."
    )
    set_run_font(r, size=11, color=INK)

    add_table_caption(doc, "Group details")
    add_table(
        doc,
        ["Item", "Details"],
        [
            ["Course", "ELEC3544 Data Science with Foundation Models"],
            ["Group", "Group 1"],
            ["Members", "Stacia Benita (3036094289)\nSurjadjajadi Clara Natalie (3036029686)\nDevano Jonathan (3036029698)"],
            ["Report date", "4 November 2026"],
        ],
        [1.25, 5.60],
        font_size=9.2,
        numeric=set(),
    )


def build_executive_summary(doc):
    add_h1(doc, "Executive summary", page_break=True)
    add_body(
        doc,
        "The project asks whether recent customer behaviour can help a retailer target customers "
        "who are likely to return. It uses one identified customer at one month-end as the "
        "prediction unit. Each customer appears only when they made a valid purchase in the "
        "previous 90 days. The target is a valid positive purchase in the following 30 days."
    )
    add_body(
        doc,
        "The strongest evidence concerns the value of behavioural detail. On the pooled "
        "August-October snapshots, extended behaviour improved PR-AUC from 0.625 to 0.634 "
        "for logistic regression and from 0.630 to 0.638 for gradient boosting. The July "
        "selection procedure chose Boosting + TF-IDF. It reached 0.616 validation PR-AUC "
        "and 0.637 pooled test PR-AUC with lift at 20 percent of 1.77."
    )
    add_body(
        doc,
        "The foundation-model experiment produced a clear negative result. MiniLM changed "
        "pooled PR-AUC by 0.000 in both model families after rounding to three decimals. "
        "TF-IDF also did not improve on behavioural boosting, which fell by 0.001 when "
        "TF-IDF features were added. We therefore treat product meaning as an exploratory "
        "signal for this prediction task and report the result rather than claiming a gain."
    )

    add_table_caption(doc, "Table 1. Main evidence and interpretation")
    add_table(
        doc,
        ["Question", "Evidence", "Interpretation"],
        [
            ["Does the data support customer-level modelling?", "541,909 invoice lines; 4,338 identified purchasing customers; 16,155 customer-snapshot rows", "The sample is sufficient for eight cohorts and two model families."],
            ["Do behavioural features add value?", "PR-AUC gain of +0.009 for logistic regression and +0.008 for gradient boosting", "The gain is consistent across both model families but modest."],
            ["Does MiniLM add value beyond behaviour?", "Pooled PR-AUC change of 0.000 in both families", "Product semantics add no measurable incremental ranking signal."],
            ["Is ranking useful?", "Pooled lift at 20 percent of 1.77 for the July-selected model", "A high-scoring group has more repeat customers, but the prediction is not causal."],
        ],
        [1.48, 2.55, 2.82],
        font_size=8.8,
    )

    add_h2(doc, "Research question and scope")
    add_body(
        doc,
        "Can the previous 90 days of transaction history predict whether an identified customer "
        "will make another valid purchase within the next 30 days? We compare RFM, extended "
        "behavioural and cancellation features, TF-IDF product groups and frozen MiniLM product "
        "groups. The experiment includes logistic regression and gradient boosting, supported "
        "by majority-class and recency-rule baselines."
    )
    add_body(
        doc,
        "This is a predictive ranking study. A high model score identifies customers who resemble "
        "earlier repeat purchasers. It does not show that an email, discount or other retention "
        "action would cause the customer to return."
    )


def build_data_section(doc):
    add_h1(doc, "Data and preparation")
    add_body(
        doc,
        "The main dataset is the UCI Machine Learning Repository Online Retail dataset, ID 352. "
        "Its workbook contains 541,909 invoice-line records from 1 December 2010 to 9 December "
        "2011. Each row is a product line within an invoice, so a repeated invoice number is "
        "expected. The retailer records customers in multiple countries, with most activity "
        "from the United Kingdom."
    )
    add_body(
        doc,
        "We also inspected Online Retail II. Its 2010-2011 sheet contains the same 541,909 rows "
        "as the main workbook plus one postage line. Its two year sheets repeat 22,523 rows "
        "from 1-9 December 2010. The earlier year falls outside the project's March-October "
        "feature windows, so it does not extend the required outcomes. Using the one-sheet "
        "archive avoids a cross-sheet duplicate risk without reducing model coverage."
    )

    add_table_caption(doc, "Table 2. Dataset audit")
    add_table(
        doc,
        ["Measure", "Value"],
        [
            ["Raw invoice-line records", "541,909"],
            ["Exact duplicate rows removed", "5,268"],
            ["Cleaned invoice-line records", "536,641"],
            ["Rows with missing CustomerID", "135,080"],
            ["Rows with missing product description", "1,454"],
            ["Identified cancellation and return lines after cleaning", "8,872"],
            ["Valid positive-purchase lines after cleaning", "392,692"],
            ["Positive invoices in the modelling table", "18,562"],
            ["Identified customers with a valid purchase", "4,338"],
        ],
        [4.35, 2.50],
        font_size=9.1,
        numeric={1},
    )

    add_h2(doc, "Cleaning and transaction definitions")
    add_body(
        doc,
        "A valid purchase requires an identified customer, an invoice that does not begin with "
        "C, positive quantity, positive unit price and a valid timestamp. Cancellations and "
        "returns remain available for historical features, but they never contribute to "
        "positive spending or a positive target label. Rows without a customer identifier "
        "cannot enter customer-level modelling."
    )
    add_body(
        doc,
        "The notebook removes exact duplicate rows and records the count. Duplicate provenance "
        "is not certain, so the work plan includes a kept-versus-dropped sensitivity test. That "
        "test matters because a change to the main cleaning rule would require every downstream "
        "snapshot and model result to be regenerated."
    )
    add_body(
        doc,
        "Product descriptions are standardised to uppercase, punctuation is replaced with spaces, "
        "and repeated whitespace is removed. The final text vocabulary contains short product "
        "names rather than catalogue descriptions. This limits the information available to "
        "TF-IDF and MiniLM."
    )


def build_snapshot_section(doc):
    add_h1(doc, "Snapshot design and leakage control")
    add_body(
        doc,
        "A snapshot date marks the end of a 90-day feature window. The model observes purchases "
        "up to and including that day. The label is constructed only from the next 30 calendar "
        "days. For the 31 July snapshot, for example, features cover 3 May to 31 July and the "
        "target covers 1 August to 30 August."
    )

    add_table_caption(doc, "Table 3. Customer cohorts and future repeat rates")
    add_table(
        doc,
        ["Snapshot", "Customer rows", "Repeat purchase rate", "Role"],
        [
            ["31 March 2011", "1,776", "27.9%", "Training"],
            ["30 April 2011", "1,891", "35.4%", "Training"],
            ["31 May 2011", "1,959", "32.4%", "Training"],
            ["30 June 2011", "1,969", "31.3%", "Training"],
            ["31 July 2011", "2,025", "31.8%", "Model selection"],
            ["31 August 2011", "1,923", "37.9%", "Rolling test"],
            ["30 September 2011", "2,151", "35.2%", "Rolling test"],
            ["31 October 2011", "2,461", "42.9%", "Rolling test"],
        ],
        [1.55, 1.35, 1.60, 1.25],
        font_size=9.0,
        numeric={1, 2},
        highlight_rows={4},
    )

    add_body(
        doc,
        "The March-June snapshots train the models. The July snapshot selects the gradient-boosting "
        "configuration, feature stage, model and F1 threshold. August, September and October form "
        "the final rolling test. For each test month, the model uses only earlier snapshots, and "
        "the July-selected settings remain fixed."
    )

    add_table_caption(doc, "Table 4. Controls against temporal leakage")
    add_table(
        doc,
        ["Control", "Implementation"],
        [
            ["Feature cutoff", "Predictors use transactions on or before the snapshot date only."],
            ["Label isolation", "The following 30 days construct the outcome and never enter the feature table."],
            ["Chronological split", "Training precedes July selection, and July precedes all final test months."],
            ["Transform fitting", "Imputers and scalers fit on the applicable training rows."],
            ["Text fitting", "TF-IDF vocabulary and cluster centres use product descriptions observed through June only."],
            ["Rolling refit", "Each final test model is refitted with earlier snapshots and excludes the current or later target windows."],
        ],
        [1.52, 5.33],
        font_size=9.0,
    )
    add_body(
        doc,
        "Customers can appear in more than one snapshot because their history and outcome change "
        "over time. The temporal split prevents future outcomes from leaking into training, but "
        "the repeated customer rows are not statistically independent. We therefore report "
        "cohort-level results as well as pooled results and avoid treating the 16,155 rows as "
        "16,155 distinct customers."
    )


def build_features_section(doc):
    add_h1(doc, "Features and models")
    add_table_caption(doc, "Table 5. Feature stages")
    add_table(
        doc,
        ["Stage", "Information used", "Purpose"],
        [
            ["RFM", "Recency, purchase frequency, monetary value and log transforms", "Simple behavioural baseline"],
            ["Extended behaviour", "Units, baskets, product count, active days, recent invoices, purchase gaps, weekend activity and history span", "Describe purchasing rhythm and basket composition"],
            ["Cancellation history", "Cancellation lines and invoices, returned value and units, recent cancellation and rate", "Separate positive purchasing from return behaviour"],
            ["TF-IDF groups", "Cluster spending shares, cluster count and entropy from word and bigram representations", "Conventional text benchmark"],
            ["MiniLM groups", "The same customer summaries from frozen 384-dimensional sentence embeddings", "Test pretrained semantic information"],
        ],
        [1.18, 3.92, 1.75],
        font_size=8.8,
    )

    add_body(
        doc,
        "TF-IDF and MiniLM both assign product descriptions to semantic groups. Candidate cluster "
        "counts are 8, 12, 16 and 20. The score uses cosine silhouette on training-period product "
        "descriptions. Each customer's 90-day history is then summarised by the share of spending "
        "in each group, the number of groups purchased from and Shannon diversity. The two "
        "representations remain separate so their gains can be compared."
    )
    add_body(
        doc,
        "Logistic regression uses median imputation, standardisation and class balancing. "
        "Gradient boosting uses histogram-based trees, median imputation and class balancing. "
        "The notebook tests a small set of boosting settings on July and then uses the selected "
        "configuration for every final comparison. This keeps model tuning separate from the "
        "final test."
    )

    add_table_caption(doc, "Table 6. Evaluation design")
    add_table(
        doc,
        ["Role", "Methods and metrics"],
        [
            ["Weak baselines", "Majority class and recency rule."],
            ["Trained comparisons", "Logistic regression and gradient boosting with RFM, extended behaviour, MiniLM and selected TF-IDF alternatives."],
            ["Primary metric", "PR-AUC, because the positive class is a minority and the retailer needs useful ranking."],
            ["Supporting metrics", "ROC-AUC, F1 at the July-selected threshold and lift among the highest-scored 20 percent."],
            ["Error review", "Confusion type, recency, frequency, spending, product variety and cancellation summaries."],
        ],
        [1.55, 5.30],
        font_size=9.0,
    )


def build_results_section(doc):
    add_h1(doc, "Model results")
    add_body(
        doc,
        "All trained methods clearly beat the recency rule on PR-AUC. The strongest learned models "
        "cluster near 0.63 on the final test. Differences within that group are too small to "
        "support a broad claim that one model family dominates the other."
    )

    if FIGURE_SOURCE.exists():
        image = Image.open(FIGURE_SOURCE)
        crop_box = (0, 0, min(1160, image.width), image.height)
        image.crop(crop_box).save(FIGURE_CROP)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(p, before=2, after=2, line=1.0)
        run = p.add_run()
        inline = run.add_picture(str(FIGURE_CROP), width=Inches(6.55))
        inline._inline.docPr.set(
            "descr",
            "Two horizontal bar charts showing July validation and August-October pooled PR-AUC for nine baselines and learned models.",
        )
        add_figure_caption(
            doc,
            "Figure 1. July validation and final pooled PR-AUC. Behavioural and text-enhanced models "
            "cluster near 0.63, while the recency and majority baselines remain lower.",
        )

    add_table_caption(doc, "Table 7. Validation and final ranking metrics")
    add_table(
        doc,
        ["Model", "July PR-AUC", "Final PR-AUC", "Final ROC-AUC", "Final F1", "Final lift"],
        [
            ["Boosting + MiniLM", "0.615", "0.638", "0.700", "0.567", "1.80"],
            ["Boosting Behaviour", "0.616", "0.638", "0.700", "0.559", "1.76"],
            ["Boosting + TF-IDF", "0.616", "0.637", "0.701", "0.560", "1.77"],
            ["Logistic + MiniLM", "0.612", "0.634", "0.696", "0.567", "1.78"],
            ["Logistic Behaviour", "0.609", "0.634", "0.696", "0.568", "1.76"],
            ["Boosting RFM", "0.598", "0.630", "0.692", "0.544", "1.78"],
            ["Logistic RFM", "0.595", "0.625", "0.687", "0.549", "1.76"],
            ["Recency rule", "0.407", "0.420", "0.552", "0.540", "1.10"],
            ["Majority baseline", "0.318", "0.407", "0.526", "0.000", "1.14"],
        ],
        [1.80, 0.93, 0.95, 1.08, 0.82, 0.80],
        font_size=8.7,
        numeric={1, 2, 3, 4, 5},
        highlight_rows={2},
    )
    add_body(
        doc,
        "Boosting + TF-IDF is highlighted because July selected it before the final months were "
        "inspected. Its 0.637 pooled PR-AUC is not replaced by the slightly higher post-hoc score "
        "of Boosting + MiniLM. This distinction protects the test from model shopping after the "
        "final outcomes are known."
    )


def build_interpretation_section(doc):
    add_h1(doc, "Interpretation and feature ablation")
    add_table_caption(doc, "Table 8. Pooled PR-AUC ablation by model family")
    add_table(
        doc,
        ["Model family", "RFM", "Plus behaviour", "Behaviour gain", "Plus MiniLM", "MiniLM gain"],
        [
            ["Logistic regression", "0.625", "0.634", "+0.009", "0.634", "+0.000"],
            ["Gradient boosting", "0.630", "0.638", "+0.008", "0.638", "+0.000"],
        ],
        [1.58, 0.82, 1.12, 1.08, 1.04, 1.02],
        font_size=9.2,
        numeric={1, 2, 3, 4, 5},
    )
    add_body(
        doc,
        "The behavioural gain is small but appears in both model families. Product semantics do "
        "not produce an additional gain at three decimal places. TF-IDF follows the same pattern: "
        "adding it to behavioural boosting lowers pooled PR-AUC from 0.638 to 0.637."
    )
    add_body(
        doc,
        "Permutation importance for the July-selected model places ActivePurchaseDays first. "
        "Recency, frequency, product variety, purchase-gap statistics and monetary value provide "
        "smaller contributions. This supports a practical reading of the result. The model "
        "learns how regularly and recently a customer purchases before it learns what the "
        "products mean."
    )
    add_body(
        doc,
        "One explanation is feature compression. A 384-dimensional MiniLM embedding is reduced "
        "to cluster shares, count and diversity before it reaches the classifier. Those summaries "
        "may repeat purchase-frequency information already present in the behavioural stage. "
        "This explanation is plausible, but the experiment does not isolate every possible "
        "semantic feature design."
    )
    add_body(
        doc,
        "The negative result is useful for the project. It shows that a pretrained text model "
        "does not automatically improve tabular prediction. The comparison against TF-IDF, "
        "behavioural boosting and two learned model families makes the conclusion stronger than "
        "an isolated MiniLM run."
    )


def build_error_section(doc):
    add_h1(doc, "Error analysis and practical use")
    add_body(
        doc,
        "The July-selected model confuses customers whose recent behaviour does not separate "
        "cleanly into returning and non-returning groups. The final test contains 6,535 customer "
        "snapshots. False positives are commercially active and comparatively recent, so a "
        "simple activity score cannot distinguish all of them from true repeaters."
    )
    add_table_caption(doc, "Table 9. Error profiles for the July-selected model")
    add_table(
        doc,
        ["Outcome", "Rows", "Median recency", "Median frequency", "Median spend", "Median products"],
        [
            ["True positive", "1,378", "17.49 days", "3", "GBP 1,064.20", "46"],
            ["False positive", "998", "24.42 days", "2", "GBP 780.16", "37"],
            ["True negative", "2,995", "32.55 days", "1", "GBP 302.46", "17"],
            ["False negative", "1,164", "36.52 days", "1", "GBP 324.14", "19"],
        ],
        [1.40, 0.72, 1.20, 1.18, 1.20, 1.08],
        font_size=8.9,
        numeric={1, 2, 3, 4, 5},
    )
    add_body(
        doc,
        "True positives have the shortest median recency, highest frequency, highest spending and "
        "largest product variety. False negatives are the most lapsed group by median recency and "
        "have low frequency. False positives sit between true positives and true negatives: they "
        "look active, but not active enough to make the selected threshold reliable."
    )
    add_body(
        doc,
        "Lift at 20 percent is 1.77 for the selected model. The highest-scoring fifth of customers "
        "therefore contains about 1.77 times the repeat-purchase rate of the full final test. "
        "This supports prioritisation for a retention team. It does not justify automatic "
        "exclusion of lower-scored customers or a claim that contact changes their behaviour."
    )
    add_body(
        doc,
        "Cancellation rates have a median of zero in every error group. The current customer-level "
        "summaries do not reveal a strong cancellation pattern in this model. A later experiment "
        "should test cancellation features directly and report whether the sparse signal helps."
    )


def build_assessment_section(doc):
    add_h1(doc, "Assessment alignment and work plan")
    add_body(
        doc,
        "The report is organised around the four assessment areas. The table below identifies "
        "where each area appears and what evidence carries the claim."
    )
    add_table_caption(doc, "Table 10. Assessment coverage")
    add_table(
        doc,
        ["Area", "Coverage in this report", "Evidence"],
        [
            ["Workload", "Dataset sourcing, audit, cleaning, feature construction and model iteration", "Tables 2-8; source files in Appendix A"],
            ["Reasoning quality", "Baselines, staged ablations, pre-test model selection and negative text result", "Tables 1, 6-8; leakage controls"],
            ["Presentation", "A result-first narrative, readable comparison tables and one focused results figure", "Executive summary; Figure 1; result and error sections"],
            ["Report quality", "Linked problem, data, method, results, errors and limits", "Full report structure; source inventory"],
        ],
        [1.15, 3.55, 2.15],
        font_size=8.9,
    )

    add_h2(doc, "Planned contribution")
    add_body(
        doc,
        "The work split remains subject to team confirmation. The table records the current plan "
        "and the evidence each person needs to produce, rather than treating planned work as "
        "completed."
    )
    add_table_caption(doc, "Table 11. Team responsibilities and evidence")
    add_table(
        doc,
        ["Member", "Responsibility", "Required evidence", "Area"],
        [
            ["Stacia Benita", "Stages 1-3: raw data, cleaning and invoice table", "Audit table, revenue reconciliation, duplicate sensitivity and EDA plots", "A"],
            ["Devano Jonathan", "Stages 8-9 and 30-32: text clusters and error analysis", "Reviewed cluster examples, cluster-count silhouette table, error table and limitations", "A, B"],
            ["Surjadjajadi Clara Natalie", "Stages 10-13 and integration", "Leakage check, baseline comparison, July selection, threshold experiment and iteration log", "B"],
        ],
        [1.28, 2.05, 2.83, 0.69],
        font_size=8.7,
    )

    add_h2(doc, "Checkpoint schedule")
    add_table_caption(doc, "Table 12. Next milestones")
    add_table(
        doc,
        ["Dates", "Work", "Dependency or output"],
        [
            ["4-9 November", "Run each owned stage in the environment and add one independent experiment", "Individual evidence for the checkpoint"],
            ["By 7 November", "Submit individual checkpoint versions", "Member name, student ID, hours, evidence and AI checks"],
            ["9-14 November", "Complete duplicate sensitivity, cluster-count review, error analysis and final figures", "Agreed sensitivity result and iteration log"],
            ["15-20 November", "Merge report sections and reconcile all reported numbers", "Group report draft"],
            ["21-25 November", "Build slides, rehearse and complete the presentation", "Final slide deck and rehearsal notes"],
        ],
        [1.22, 3.18, 2.45],
        font_size=8.8,
    )
    add_body(
        doc,
        "The checkpoint requires each member's approximate hours and own evidence. We do not "
        "invent those values at group level. Each member should record the time from their run "
        "log before submission, because the checkpoint asks for work completed and next tasks "
        "from that person's perspective."
    )
    add_h2(doc, "Blockers and AI assistance")
    add_body(
        doc,
        "The duplicate sensitivity can change the main cleaning table and every downstream result. "
        "Stacia should report that result early. The text track also needs a stable cluster count "
        "and manual review of cluster examples before its contribution can be described as "
        "complete. Final course dates remain to be confirmed."
    )
    add_body(
        doc,
        "The integrated pipeline and notebook were generated with Codex assistance. We treat them "
        "as drafts to verify, not as evidence that independent work is complete. Each member will "
        "rerun the owned stage, compare outputs with notebook_outputs, inspect code and data "
        "assumptions, and add an experiment that is not already in the notebook. Any rerun that "
        "changes the results takes precedence over the figures in this report."
    )


def build_conclusion_section(doc):
    add_h1(doc, "Limitations and conclusion")
    add_h2(doc, "Limitations")
    add_body(
        doc,
        "The data are observational. A predicted repeat purchase is an association, not proof "
        "that a retention action causes a return. Rows without a customer identifier cannot enter "
        "the model, so the results apply only to identified customers. The retailer also includes "
        "wholesale-like purchasing patterns, which may not transfer to a typical consumer retailer."
    )
    add_body(
        doc,
        "Product descriptions are short and noisy. MiniLM is a general-domain encoder and was "
        "not fine-tuned on the retail target. Customer text histories compress each embedding "
        "into cluster summaries. These design choices can remove semantic detail that might "
        "matter in another representation."
    )
    add_body(
        doc,
        "The final test covers three monthly snapshots from one year. Customers repeat across "
        "snapshots, so pooled rows are not independent observations. The pooled score should be "
        "read with the monthly variation, and the model should be revalidated if the retailer's "
        "product mix or customer base changes."
    )

    add_h2(doc, "Conclusion")
    add_body(
        doc,
        "The project turns invoice-line records into chronological customer predictions and "
        "compares behavioural, conventional text and pretrained semantic features. Extended "
        "behaviour improves PR-AUC over RFM by 0.009 in logistic regression and 0.008 in gradient "
        "boosting. MiniLM and TF-IDF do not add a measurable gain beyond behaviour."
    )
    add_body(
        doc,
        "The result changes how we should use the model. Purchase rhythm and recent activity "
        "support customer prioritisation. Product meaning is worth studying for other tasks, "
        "such as catalogue search or product grouping, but it is not the main predictive signal "
        "for the next 30-day repeat purchase in this dataset. The final report should retain the "
        "negative text result, the duplicate sensitivity and the error analysis because they give "
        "the team a more credible account of what the pipeline can and cannot do."
    )

    add_h2(doc, "Recommended next actions")
    add_table(
        doc,
        ["Action", "Owner", "Reason"],
        [
            ["Run duplicate kept-versus-dropped sensitivity", "Stacia", "Tests the cleaning rule with the largest possible downstream effect."],
            ["Review cluster names and cluster-count stability", "Devano", "Shows whether the semantic result is an artefact of one arbitrary cluster count."],
            ["Verify leakage and rerun the July-selected test", "Clara", "Protects the central model-selection claim before the report is finalised."],
            ["Reconcile report numbers after all reruns", "Clara", "Keeps the report and notebook evidence consistent."],
        ],
        [2.30, 1.15, 3.40],
        font_size=8.9,
    )


def build_appendix(doc):
    add_h1(doc, "Appendix A. Replication materials", page_break=True)
    add_table(
        doc,
        ["File", "Purpose"],
        [
            ["Online_Retail_Repeat_Purchase_Project.ipynb", "Executed modelling pipeline, outputs, error analysis and conclusions."],
            ["Detailed_Project_Pipeline_Aligned.docx", "Stage-by-stage explanation aligned to the executed notebook."],
            ["data/online+retail.zip", "Raw UCI Online Retail workbook used for the reported results."],
            ["data/online+retail+ii.zip", "Additional workbook inspected for dataset comparison; not used in the main pipeline."],
            ["notebook_outputs/", "Generated data audit, model-ready snapshots, cluster examples, metrics, predictions and figures."],
        ],
        [2.45, 4.40],
        font_size=8.9,
    )
    add_h2(doc, "Replication sequence")
    add_body(
        doc,
        "Extract the source workbook, set DATA_PATH to the extracted file, run the notebook from "
        "the project directory and retain the notebook_outputs folder. The notebook currently "
        "expects Online Retail 2.xlsx, while the archive contains Online Retail.xlsx; use one "
        "file name consistently or update DATA_PATH before execution."
    )
    add_body(
        doc,
        "For each reported result, retain the source table or figure and the command or notebook "
        "cell that produced it. When a sensitivity run changes a number, update the report, the "
        "checkpoint table and the slides together. Do not mix results from different cleaning "
        "rules in the same comparison."
    )
    add_h2(doc, "Interpretation of reported precision")
    add_body(
        doc,
        "PR-AUC values are rounded to three decimal places and lifts to two. A reported gain of "
        "0.000 therefore means no difference at the displayed precision, not a formal statistical "
        "test of equivalence. The project reports the small differences directly and avoids "
        "claims that the two text methods are equivalent."
    )


def main():
    doc = setup_document()
    build_cover(doc)
    build_executive_summary(doc)
    build_data_section(doc)
    build_snapshot_section(doc)
    build_features_section(doc)
    build_results_section(doc)
    build_interpretation_section(doc)
    build_error_section(doc)
    build_assessment_section(doc)
    build_conclusion_section(doc)
    build_appendix(doc)
    doc.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    main()
