from docx import Document
import shutil
from docx.shared import Inches, Pt
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# 复制文件
shutil.copyfile('./md4report/assets/md4report.docx', './md4report/assets/test.docx')

# 打开复制后的docx文件
doc = Document('./md4report/assets/test.docx')

# 设置页面宽度和页边距
section = doc.sections[0]
section.page_width = Inches(8.27)
section.page_height = Inches(11.69)
section.left_margin = Inches(1)
section.right_margin = Inches(1)
section.top_margin = Inches(1)
section.bottom_margin = Inches(1)

# 在文档开头插入一个标题
title = doc.add_paragraph('实验报告')
title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

# 在标题后插入一个一行两列的表格
score = doc.add_table(rows=1, cols=2)
score.alignment = WD_TABLE_ALIGNMENT.RIGHT

# 设置表格内容
cells = score.rows[0].cells
cells[0].text = '成绩'
cells[1].text = '100'

# 缩小表格
for cell in cells:
    cell.width = Inches(1)

# 设置表格边框为全黑色
tbl = score._element
tblBorders = OxmlElement('w:tblBorders')
for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
    border = OxmlElement(f'w:{border_name}')
    border.set(qn('w:val'), 'single')
    border.set(qn('w:sz'), '4')  # 边框宽度
    border.set(qn('w:space'), '0')
    border.set(qn('w:color'), '000000')  # 黑色
    tblBorders.append(border)
tbl.tblPr.append(tblBorders)

# 将新插入的内容移动到文档开头
starter = doc.paragraphs[0].insert_paragraph_before(text='', style=None)  # 转到文章开头
doc.paragraphs[0]._element.addprevious(title._element)  # 将标题移动到开头
doc.paragraphs[1]._element.addprevious(score._element)  # 将表格移动到标题后

# 保存修改后的文档
doc.save('./md4report/assets/test.docx')