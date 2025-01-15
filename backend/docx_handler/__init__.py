from typing import Callable, Dict
from docx_handler.gzmtu import process_gzmtu_docx
from docx_handler.general import insert_indent

handler_map: Dict[str, Callable] = {
    "GZMTU": process_gzmtu_docx,
    "first_line_indent": insert_indent,
}
