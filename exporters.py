"""
Модуль экспорта таблиц погрешностей в форматы Microsoft Office:
- Microsoft Excel (.xlsx) через openpyxl
- Microsoft Word (.docx) через python-docx
- Буфер обмена Windows (TSV / Excel Clipboard)
- Markdown и CSV
"""

import os
from typing import List, Dict, Any, Optional

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

# Заголовки для 15-колоночной (сводной) таблицы
HEADERS_SUMMARY = [
    ("ik", "№ ИК"),
    ("i1", "I₁, %"),
    ("cos_phi", "cos φ"),
    ("kt_tt", "КТ ТТ"),
    ("kt_th", "КТ ТН"),
    ("kt_meter", "КТ сч."),
    ("delta_i", "δi, %"),
    ("theta_i", "θi, '"),
    ("delta_u", "δu, %"),
    ("theta_u", "θu, '"),
    ("delta_L", "δL, %"),
    ("dW_osn_act", "δW осн (акт), %"),
    ("dW_osn_react", "δW осн (реакт), %"),
    ("dW_rab_act", "δW раб (акт), %"),
    ("dW_rab_react", "δW раб (реакт), %"),
]

# Заголовки для 29-колоночной (полной по МИ 3000) таблицы
HEADERS_FULL = [
    ("ik", "Номер ИК"),
    ("i1", "I₁, % от Iном"),
    ("cos_phi", "cos φ"),
    ("kt_tt", "КТ ТТ"),
    ("kt_th", "КТ ТН"),
    ("delta_i", "δi, %"),
    ("theta_i", "θi, мин"),
    ("delta_u", "δu, %"),
    ("theta_u", "θu, мин"),
    ("delta_theta_act", "δθ (акт), %"),
    ("delta_theta_react", "δθ (реакт), %"),
    ("delta_L", "δL, %"),
    ("kt_meter_act", "КТ сч. акт"),
    ("kt_meter_react", "КТ сч. реакт"),
    ("dco_act", "δс.о (акт), %"),
    ("dco_react", "δс.о (реакт), %"),
    ("delta_t", "Δt, °C"),
    ("dct_act", "δct (акт), %"),
    ("dct_react", "δct (реакт), %"),
    ("dcU_act", "δcU (акт), %"),
    ("dcU_react", "δcU (реакт), %"),
    ("dcf_act", "δcf (акт), %"),
    ("dcf_react", "δcf (реакт), %"),
    ("dmi_act", "δми (акт), %"),
    ("dmi_react", "δми (реакт), %"),
    ("dW_osn_act", "δW осн (акт), %"),
    ("dW_osn_react", "δW осн (реакт), %"),
    ("dW_rab_act", "δW раб (акт), %"),
    ("dW_rab_react", "δW раб (реакт), %"),
]


def get_tsv_text(rows: List[Dict[str, Any]], mode: str = "summary") -> str:
    """Генерирует текст с разделителями-табуляциями (TSV) для копирования в буфер."""
    headers_spec = HEADERS_SUMMARY if mode == "summary" else HEADERS_FULL
    header_line = "\t".join(label for _, label in headers_spec)
    
    data_lines = []
    for r in rows:
        line_vals = [str(r.get(key, "")) for key, _ in headers_spec]
        data_lines.append("\t".join(line_vals))
        
    return header_line + "\n" + "\n".join(data_lines)


def export_to_excel(
    rows: List[Dict[str, Any]],
    filepath: str,
    mode: str = "summary",
    title: Optional[str] = None
) -> str:
    """Экспорт таблицы погрешностей в Microsoft Excel (.xlsx) с профессиональным форматированием."""
    headers_spec = HEADERS_SUMMARY if mode == "summary" else HEADERS_FULL
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Погрешности ИК"

    # Стили
    font_title = Font(name="Calibri", size=13, bold=True, color="1F497D")
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=9, bold=False)
    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_alt = PatternFill(start_color="F2F5F8", end_color="F2F5F8", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_side = Side(style="thin", color="D9D9D9")
    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    current_row = 1

    # Заголовок
    if title:
        ws.cell(row=current_row, column=1, value=title).font = font_title
        ws.row_dimensions[current_row].height = 24
        current_row += 2

    # Строка заголовков таблицы
    header_row_idx = current_row
    ws.row_dimensions[header_row_idx].height = 28
    for col_idx, (_, label) in enumerate(headers_spec, start=1):
        cell = ws.cell(row=header_row_idx, column=col_idx, value=label)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_cell

    current_row += 1

    # Данные
    for r_idx, r in enumerate(rows):
        row_num = current_row
        row_fill = fill_alt if (r_idx % 2 == 1) else fill_white
        ws.row_dimensions[row_num].height = 18

        for col_idx, (key, _) in enumerate(headers_spec, start=1):
            val = r.get(key, "")
            cell = ws.cell(row=row_num, column=col_idx, value=val)
            cell.font = font_data
            cell.fill = row_fill
            cell.alignment = align_center
            cell.border = border_cell
            
        current_row += 1

    # Автоподбор ширины столбцов
    for col_idx, (key, label) in enumerate(headers_spec, start=1):
        max_len = len(label)
        for r in rows:
            max_len = max(max_len, len(str(r.get(key, ""))))
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 9)

    ws.freeze_panes = ws[f"A{header_row_idx + 1}"]
    wb.save(filepath)
    return filepath


def export_to_docx(
    rows: List[Dict[str, Any]],
    filepath: str,
    mode: str = "summary",
    title: Optional[str] = None
) -> str:
    """Экспорт таблицы погрешностей в Microsoft Word (.docx) по ГОСТ."""
    headers_spec = HEADERS_SUMMARY if mode == "summary" else HEADERS_FULL
    doc = docx.Document()

    # Если 29 колонок — ставим альбомную ориентацию (Landscape)
    section = doc.sections[0]
    if mode == "full":
        new_width, new_height = section.page_height, section.page_width
        section.orientation = docx.enum.section.WD_ORIENT.LANDSCAPE
        section.page_width = new_width
        section.page_height = new_height
        section.left_margin = Inches(0.4)
        section.right_margin = Inches(0.4)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
    else:
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)

    # Заголовок документа
    title_text = title or ("Таблица погрешностей измерений" if mode == "summary" else "Таблица 12. Полный расчет погрешностей измерительных каналов")
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run(title_text)
    run_title.font.name = "Times New Roman"
    run_title.font.size = Pt(12)
    run_title.font.bold = True
    p_title.paragraph_format.space_after = Pt(10)

    # Создание таблицы
    num_cols = len(headers_spec)
    table = doc.add_table(rows=1, cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    # Заголовки таблицы
    hdr_cells = table.rows[0].cells
    hdr_tr = table.rows[0]._tr.get_or_add_trPr()
    hdr_tr.append(parse_xml(r'<w:tblHeader %s/>' % nsdecls('w')))
    hdr_tr.append(parse_xml(r'<w:cantSplit %s/>' % nsdecls('w')))

    font_sz = Pt(7.5) if mode == "full" else Pt(8.5)

    for i, (_, label) in enumerate(headers_spec):
        cell = hdr_cells[i]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        # Фон ячейки заголовка
        tcPr = cell._tc.get_or_add_tcPr()
        tcPr.append(parse_xml(r'<w:shd %s w:fill="EBF1F5"/>' % nsdecls('w')))
        
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(label)
        r.font.name = "Times New Roman"
        r.font.size = font_sz
        r.font.bold = True

    # Заполнение данными
    for row_data in rows:
        row_cells = table.add_row().cells
        tr = table.rows[-1]._tr.get_or_add_trPr()
        tr.append(parse_xml(r'<w:cantSplit %s/>' % nsdecls('w')))

        for i, (key, _) in enumerate(headers_spec):
            cell = row_cells[i]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            r = p.add_run(str(row_data.get(key, "")))
            r.font.name = "Times New Roman"
            r.font.size = font_sz

    # Установка границ таблицы (тонкая сетка по ГОСТ)
    tblPr = table._tbl.tblPr
    borders_xml = parse_xml(
        r'<w:tblBorders %s>'
        r'  <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'  <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'  <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="D0D0D0"/>'
        r'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="D0D0D0"/>'
        r'</w:tblBorders>' % nsdecls('w')
    )
    tblPr.append(borders_xml)

    doc.save(filepath)
    return filepath
