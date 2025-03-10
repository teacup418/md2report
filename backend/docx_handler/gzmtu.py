import logging
import pathlib

from docx import Document
from docx.document import Document as TDocument
from docx.oxml.xmlchemy import OxmlElement
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph, Run
from docx.image.image import Image
from docx.enum.text import WD_BREAK
from docx.shared import Inches
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

from docx_handler.utils import get_paras_by_style_name


def add_logo(doc: TDocument):
    paras = get_paras_by_style_name(doc, "Title")
    if len(paras) != 1:
        logging.error("Title of docx not found")
    title: Paragraph = paras[0]
    logo_path: pathlib.Path = (
        pathlib.Path(__file__).parent.resolve() / "assets" / "gzmtu" / "logo.jpg"
    )
    image_para: Paragraph = doc.add_paragraph()
    image_para.style = doc.styles["Figure"]
    r = image_para.add_run()
    for _ in range(3):
        r.add_break(WD_BREAK.LINE)
    r.add_picture(str(logo_path), width=Inches(4))
    for _ in range(2):
        r.add_break(WD_BREAK.LINE)
    title._p.addprevious(image_para._p)


def add_school_name(doc: TDocument):
    paras = get_paras_by_style_name(doc, "School")
    school = doc.add_paragraph("广州航海学院")

def add_course_name(doc: TDocument):
    paras = get_paras_by_style_name(doc, "Course")
    course = doc.add_paragraph("实验报告")

def add_student_info_table(doc: TDocument):
    """
    Add the table of student info.

    There will be two columns in the table, the first
    col is title and the second is content.
    
    The first col has no border and the second col will have
    only the bottom border.

    The style of table is `StudentInfoTable`, and the style
    of text is `StudentInfo`
    """
    # paras = get_paras_by_style_name(doc, "Subtitle")
    # if len(paras) != 1:
    #     logging.error("Subtitle of docx not found")
    # title: Paragraph = paras[0]
    # paras = get_paras_by_style_name(doc, "School")

    school = doc.add_paragraph("广州航海学院")
    # course = doc.add_paragraph("实验报告")
    front_break: Paragraph = doc.add_paragraph()
    r: Run = front_break.add_run()
    r.add_break()
    school.insert_paragraph_before(front_break._p)

    back_break: Paragraph = doc.add_paragraph()
    r: Run = back_break.add_run()
    for _ in range(2):
        r.add_break()

    table: Table = doc.add_table(rows=3, cols=4)
    table.style = doc.styles["StudentInfoTable"]
    cell: _Cell

    table.columns[0].cells[0].text = "专业班级"
    table.columns[2].cells[0].text = "实验日期"
    table.columns[0].cells[1].text = "姓名"
    table.columns[2].cells[1].text = "学号"
    table.columns[0].cells[2].text = "实验名称"
    table.columns[2].cells[2].text = "指导老师"

    for cell in table.columns[1].cells:
        cell.width = Inches(3)
        cell.paragraphs[0].style = doc.styles["StudentInfo"]

    for cell in table.columns[3].cells:
        cell.width = Inches(3)
        cell.paragraphs[0].style = doc.styles["StudentInfo"]

    for cell in table.columns[0].cells:
        cell.width = Inches(2)
        cell.paragraphs[0].style = doc.styles["StudentInfo"]

    for cell in table.columns[2].cells:
        cell.width = Inches(2)
        cell.paragraphs[0].style = doc.styles["StudentInfo"]

    # 遍历表格的每一行
    for row in table.rows:
        # 遍历每一行的单元格
        for idx, cell in enumerate(row.cells):
            # 获取单元格的边框属性
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()

            # 创建边框元素
            tcBorders = OxmlElement('w:tcBorders')

            # 如果是第2列或第4列，设置下边框
            if idx == 1 or idx == 3:
                bottom = OxmlElement('w:bottom')
                bottom.set(qn('w:val'), 'single')
                bottom.set(qn('w:sz'), '4')
                bottom.set(qn('w:space'), '0')
                bottom.set(qn('w:color'), '000000')
                tcBorders.append(bottom)
            else:
                # 如果是第1列或第3列，移除所有边框
                for border_name in ['top', 'left', 'bottom', 'right']:
                    border = OxmlElement(f'w:{border_name}')
                    border.set(qn('w:val'), 'nil')
                    tcBorders.append(border)

            # 将边框元素添加到单元格属性中
            tcPr.append(tcBorders)
            # 如果是第1列或第3列，设置分散对齐
            if idx == 0 or idx == 2:
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.DISTRIBUTE

    # # set table border
    # bottom = table._element.xpath("./w:tblPr/w:tblLook")[0]
    # bottom.set(qn("w:lastColumn"), "1")
    # bottom.set(qn("w:firstRow"), "0")
    school._p.addnext(table._tbl)
    school._p.addnext(back_break._p)

def replace_top_caption(doc:TDocument):
    """
    Replace the caption of "Table of content" by Chinese.

    `sdt` is for "structured document tags", which is the root element TOC.
    The caption of TOC is in the first paragraph of `sdtContent`.
    """

    caption = doc.element.xpath("./w:body/w:sdt/w:sdtContent/w:p/w:r/w:t")[0]
    caption.text = "目录"

def process_gzmtu_docx(filename: str):
    doc: TDocument = Document(filename)
    # add_school_name(doc)
    # add_course_name(doc)
    add_student_info_table(doc)
    doc.save(filename)
