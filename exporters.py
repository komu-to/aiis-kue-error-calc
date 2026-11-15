"""
Модуль экспорта таблиц погрешностей в форматы Microsoft Office:
- Microsoft Excel (.xlsx) через openpyxl
- Microsoft Word (.docx) через python-docx
- Буфер обмена Windows (TSV / Excel Clipboard)
- Markdown и CSV
"""

import os
import re
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


def sanitize_filename_part(text: str) -> str:
    """Очищает строку от недопустимых в Windows символов и лишних пробелов."""
    # Заменяем разделители классов счетчика типа ' / ' на '-'
    text = re.sub(r'\s*[\/\\]\s*', '-', text.strip())
    # Заменяем двоеточия на дефис
    text = text.replace(":", "-")
    for bad_char in ['*', '?', '"', '<', '>', '|']:
        text = text.replace(bad_char, '')
    # Убираем запятые с пробелами -> просто запятая
    text = re.sub(r'\s*,\s*', ',', text)
    # Заменяем пробелы вокруг дефиса на одиночный дефис
    text = re.sub(r'\s*-\s*', '-', text)
    # Заменяем оставшиеся группы пробелов на знак подчеркивания
    text = re.sub(r'\s+', '_', text.strip())
    # Убираем возможные комбинации _-_ или -_-
    text = re.sub(r'[_\-]+-[_\-]+', '-', text)
    return text


def make_export_filename(
    ik_str: str = "1",
    tt_str: str = "0.5S",
    th_str: str = "0.5",
    meter_str: str = "0.2S / 0.5",
    delta_t: Any = 20.0,
    ext: str = "xlsx",
    is_all_groups: bool = False
) -> str:
    """
    Генерирует осмысленное имя файла экспорта по шаблону:
    "№каналов-ТТ_ТН_Счетчик_температура.<ext>"
    Например:
    "ИК1,2,15,16-ТТ0.2S_ТН0.2_Сч0.2S-0.5_t12C.xlsx"
    """
    ext = ext.lstrip(".")
    if is_all_groups:
        ik_part = "Все_ИК_1-23"
        t_part = f"t{delta_t}C" if delta_t is not None else "t20C"
        return f"{ik_part}-Сводная_таблица_проекта_{t_part}.{ext}"

    # 1. Номер(а) каналов
    ik_clean = sanitize_filename_part(str(ik_str))
    if not ik_clean.lower().startswith("ик") and not ik_clean.lower().startswith("кан"):
        ik_part = f"ИК{ik_clean}"
    else:
        ik_part = ik_clean

    # 2. ТТ
    tt_raw = str(tt_str).strip()
    if tt_raw in ["—", "-", "", "Без ТТ", "Нет", "Нет (прямое включение)"]:
        tt_part = "безТТ"
    else:
        tt_part = f"ТТ{sanitize_filename_part(tt_raw)}"

    # 3. ТН
    th_raw = str(th_str).strip()
    if th_raw in ["—", "-", "", "Без ТН", "Нет"]:
        th_part = "безТН"
    else:
        th_part = f"ТН{sanitize_filename_part(th_raw)}"

    # 4. Счетчик
    meter_raw = str(meter_str).replace("(Прямое включение)", "").strip()
    meter_part = f"Сч{sanitize_filename_part(meter_raw)}"

    # 5. Температура
    try:
        dt_val = float(str(delta_t).replace(",", "."))
        if dt_val.is_integer():
            t_part = f"t{int(dt_val)}C"
        else:
            t_part = f"t{dt_val}C"
    except Exception:
        t_part = f"t{sanitize_filename_part(str(delta_t))}C"

    return f"{ik_part}-{tt_part}_{th_part}_{meter_part}_{t_part}.{ext}"



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


# =====================================================================
# ЭКСПОРТ В ФОРМАТ ТАБЛИЦЫ 5.1 ПРОЕКТА (МАТРИЦА ПО ТОКАМ И cos φ)
# =====================================================================

def make_channel_label(first_row: Dict[str, Any]) -> str:
    """Формирует наименование канала для шапки группы Таблицы 5.1."""
    ik = str(first_row.get("ik", "")).strip()
    tt = str(first_row.get("kt_tt", "")).strip()
    th = str(first_row.get("kt_th", "")).strip()
    meter_act = str(first_row.get("kt_meter_act", "")).strip()
    meter_react = str(first_row.get("kt_meter_react", "")).strip()
    meter_part = f"Сч. {meter_act}/{meter_react}" if meter_react else f"Сч. {meter_act}"

    parts = []
    if tt and tt not in ["—", "-", "", "Без ТТ", "не норм", "н/н", "None"]:
        parts.append(f"ТТ {tt}")
    if th and th not in ["—", "-", "", "Без ТН", "не норм", "н/н", "None"]:
        parts.append(f"ТН {th}")
    parts.append(meter_part)

    return f"{ik} ({'; '.join(parts)})"


def group_rows_for_table51(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Разбивает плоский список строк на группы по измерительным каналам."""
    groups = []
    current_ik = None
    cur_group_rows = []

    for r in rows:
        ik = str(r.get("ik", "")).strip()
        if ik != current_ik:
            if current_ik is not None:
                groups.append({"label": make_channel_label(cur_group_rows[0]), "rows": cur_group_rows})
            current_ik = ik
            cur_group_rows = [r]
        else:
            cur_group_rows.append(r)

    if cur_group_rows:
        groups.append({"label": make_channel_label(cur_group_rows[0]), "rows": cur_group_rows})

    return groups


def build_table51_data(groups_list: List[Dict[str, Any]], val_type: str = "osn") -> List[Dict[str, Any]]:
    """
    Формирует структуру матрицы Таблицы 5.1 для экспорта в Word / Excel.
    val_type: 'osn' (δW осн) или 'rab' (δW раб)
    """
    matrix_groups = []
    cos_list = ["0,5", "0,8", "0,9", "1,0"]

    key_act = "dW_osn_act" if val_type == "osn" else "dW_rab_act"
    key_react = "dW_osn_react" if val_type == "osn" else "dW_rab_react"

    for grp_info in groups_list:
        label = grp_info["label"]
        rows = grp_info["rows"]

        idx_map = {}
        has_1pct = False
        for r in rows:
            i1 = str(r.get("i1", "")).strip()
            cos_p = str(r.get("cos_phi", "")).strip().replace("1.0", "1,0")
            if cos_p == "1":
                cos_p = "1,0"
            if i1 == "1(2)":
                has_1pct = True
            idx_map[(i1, cos_p)] = r

        group_rows = []
        for cos_p in cos_list:
            row_dict = {
                "cos_phi": cos_p,
                "i_1_act": "-",
                "i_1_react": "-",
                "i_5_act": "-",
                "i_5_react": "-",
                "i_20_act": "-",
                "i_20_react": "-",
                "i_100_act": "-",
                "i_100_react": "-",
            }

            # Точка 1(2)%
            if has_1pct and ("1(2)", cos_p) in idx_map:
                r1 = idx_map[("1(2)", cos_p)]
                row_dict["i_1_act"] = str(r1.get(key_act, "-"))
                r_react = str(r1.get(key_react, "-"))
                row_dict["i_1_react"] = "не норм" if r_react in ["н/н", "не норм"] and cos_p == "1,0" else r_react

            # Точка 5%
            if ("5", cos_p) in idx_map:
                r5 = idx_map[("5", cos_p)]
                row_dict["i_5_act"] = str(r5.get(key_act, "-"))
                r_react = str(r5.get(key_react, "-"))
                row_dict["i_5_react"] = "не норм" if r_react in ["н/н", "не норм"] and cos_p == "1,0" else r_react

            # Точка 20%
            if ("20", cos_p) in idx_map:
                r20 = idx_map[("20", cos_p)]
                row_dict["i_20_act"] = str(r20.get(key_act, "-"))
                r_react = str(r20.get(key_react, "-"))
                row_dict["i_20_react"] = "не норм" if r_react in ["н/н", "не норм"] and cos_p == "1,0" else r_react

            # Точка 100% (или 120%)
            r100 = idx_map.get(("100", cos_p)) or idx_map.get(("120", cos_p))
            if r100:
                row_dict["i_100_act"] = str(r100.get(key_act, "-"))
                r_react = str(r100.get(key_react, "-"))
                row_dict["i_100_react"] = "не норм" if r_react in ["н/н", "не норм"] and cos_p == "1,0" else r_react

            group_rows.append(row_dict)

        matrix_groups.append({
            "label": label,
            "rows": group_rows
        })

    return matrix_groups


def export_table51_to_docx(
    rows: List[Dict[str, Any]],
    filepath: str,
    val_type: str = "osn",
    title: Optional[str] = None
) -> str:
    """Экспорт данных в формате Таблицы 5.1 проекта в Microsoft Word (.docx)."""
    groups_list = group_rows_for_table51(rows)
    groups_data = build_table51_data(groups_list, val_type=val_type)

    doc = docx.Document()

    # Альбомная ориентация
    section = doc.sections[0]
    section.orientation = docx.enum.section.WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

    val_name = "основной" if val_type == "osn" else "в рабочих условиях"
    val_sym = "δWосн" if val_type == "osn" else "δWраб"

    if not title:
        title = f"Таблица 5.1 — Пределы допускаемой {val_name} погрешности ИК при измерении активной (реактивной) электроэнергии, ± {val_sym} (%)"

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(4)
    r_t = p_title.add_run(title)
    r_t.font.name = "Times New Roman"
    r_t.font.size = Pt(11)
    r_t.font.bold = True

    total_data_rows = sum(len(g["rows"]) for g in groups_data)
    table = doc.add_table(rows=3 + total_data_rows, cols=10)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Ширина столбцов
    col_widths = [Inches(2.5), Inches(0.9)] + [Inches(0.9)] * 8
    for row in table.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = width

    def _format_cell(cell, text, bold=False, size_pt=9.5, align=WD_ALIGN_PARAGRAPH.CENTER):
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = align
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.text = ""
        run = p.add_run(text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(size_pt)
        run.font.bold = bold

    # Повторение шапки и cantSplit
    for r_idx in range(3):
        tr = table.rows[r_idx]._tr.get_or_add_trPr()
        tr.append(parse_xml(r'<w:tblHeader %s/>' % nsdecls('w')))
        tr.append(parse_xml(r'<w:cantSplit %s/>' % nsdecls('w')))

    # Шапка Row 0
    c0 = table.cell(0, 0)
    c0.merge(table.cell(2, 0))
    _format_cell(c0, "Номера каналов", bold=True, size_pt=9.5)

    c1 = table.cell(0, 1)
    c1.merge(table.cell(2, 1))
    _format_cell(c1, "Значение cos φ (sin φ)", bold=True, size_pt=9.0)

    c_top = table.cell(0, 2)
    c_top.merge(table.cell(0, 9))
    _format_cell(c_top, f"Пределы допускаемой {val_name} погрешности ИК при измерении активной (реактивной) электроэнергии, ± {val_sym} (%)", bold=True, size_pt=9.5)

    # Шапка Row 1
    ranges = [
        (2, 3, "δ1(2)%, I1(2) ≤ Iизм < I5%"),
        (4, 5, "δ5%, I5% ≤ Iизм < I20%"),
        (6, 7, "δ20%, I20% ≤ Iизм < I100%"),
        (8, 9, "δ100%, I100% ≤ Iизм ≤ I120%"),
    ]
    for c_start, c_end, txt in ranges:
        cell_sub = table.cell(1, c_start)
        cell_sub.merge(table.cell(1, c_end))
        _format_cell(cell_sub, txt, bold=True, size_pt=8.5)

    # Шапка Row 2: А и Р
    for i in range(2, 10):
        t_ar = "А" if i % 2 == 0 else "Р"
        _format_cell(table.cell(2, i), t_ar, bold=True, size_pt=9.5)

    # Заполнение данными
    current_r = 3
    for grp in groups_data:
        label = grp["label"]
        g_rows = grp["rows"]
        grp_len = len(g_rows)
        start_r = current_r
        end_r = current_r + grp_len - 1

        for r_offset, r_data in enumerate(g_rows):
            row_idx = current_r + r_offset
            tr = table.rows[row_idx]._tr.get_or_add_trPr()
            tr.append(parse_xml(r'<w:cantSplit %s/>' % nsdecls('w')))

            _format_cell(table.cell(row_idx, 1), r_data["cos_phi"], size_pt=9.0)

            cols_vals = [
                r_data["i_1_act"], r_data["i_1_react"],
                r_data["i_5_act"], r_data["i_5_react"],
                r_data["i_20_act"], r_data["i_20_react"],
                r_data["i_100_act"], r_data["i_100_react"],
            ]
            for c_idx, val in enumerate(cols_vals, start=2):
                _format_cell(table.cell(row_idx, c_idx), val, size_pt=9.0)

        # Объединение ячейки группы по вертикали
        cell_grp = table.cell(start_r, 0)
        cell_grp.merge(table.cell(end_r, 0))
        _format_cell(cell_grp, label, size_pt=9.0, align=WD_ALIGN_PARAGRAPH.LEFT)

        current_r += grp_len

    # Границы по ГОСТ
    tblPr = table._tbl.tblPr
    borders_xml = parse_xml(
        r'<w:tblBorders %s>'
        r'  <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'  <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'  <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'</w:tblBorders>' % nsdecls('w')
    )
    tblPr.append(borders_xml)

    # Примечания
    p_notes = doc.add_paragraph()
    p_notes.paragraph_format.space_before = Pt(6)
    p_notes.paragraph_format.space_after = Pt(0)
    r_n = p_notes.add_run("Примечания:\n1. Характеристики погрешности ИК даны для измерения электроэнергии и средней мощности (получасовая).\n2. В качестве характеристик относительной погрешности указаны границы интервала, соответствующие вероятности 0,95.")
    r_n.font.name = "Times New Roman"
    r_n.font.size = Pt(8.5)
    r_n.font.italic = True

    doc.save(filepath)
    return filepath


def export_table51_to_excel(
    rows: List[Dict[str, Any]],
    filepath: str,
    val_type: str = "osn",
    title: Optional[str] = None
) -> str:
    """Экспорт данных в формате Таблицы 5.1 проекта в Microsoft Excel (.xlsx)."""
    groups_list = group_rows_for_table51(rows)
    groups_data = build_table51_data(groups_list, val_type=val_type)

    wb = openpyxl.Workbook()
    ws = wb.active
    val_name = "основной" if val_type == "osn" else "в рабочих условиях"
    val_sym = "δWосн" if val_type == "osn" else "δWраб"
    ws.title = f"Таблица 5.1 ({val_sym})"

    if not title:
        title = f"Таблица 5.1 — Пределы допускаемой {val_name} погрешности ИК при измерении активной (реактивной) электроэнергии, ± {val_sym} (%)"

    ws['A1'] = title
    ws['A1'].font = Font(name="Calibri", size=12, bold=True, color="1F497D")
    ws.row_dimensions[1].height = 25

    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=9, bold=False)
    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_group_a = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_group_b = PatternFill(start_color="F5F8FA", end_color="F5F8FA", fill_type="solid")

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin_side = Side(style="thin", color="B0C4DE")
    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # Строки 3, 4, 5 - шапка
    ws.merge_cells('A3:A5')
    ws['A3'] = "Номера каналов"

    ws.merge_cells('B3:B5')
    ws['B3'] = "Значение\ncos φ (sin φ)"

    ws.merge_cells('C3:J3')
    ws['C3'] = f"Пределы допускаемой {val_name} погрешности ИК при измерении активной (реактивной) электроэнергии, ± {val_sym} (%)"

    ws.merge_cells('C4:D4')
    ws['C4'] = "δ1(2)%, I1(2) ≤ Iизм < I5%"
    ws.merge_cells('E4:F4')
    ws['E4'] = "δ5%, I5% ≤ Iизм < I20%"
    ws.merge_cells('G4:H4')
    ws['G4'] = "δ20%, I20% ≤ Iизм < I100%"
    ws.merge_cells('I4:J4')
    ws['I4'] = "δ100%, I100% ≤ Iизм ≤ I120%"

    for col_idx in range(3, 11):
        col_letter = get_column_letter(col_idx)
        ws[f'{col_letter}5'] = "А" if col_idx % 2 == 1 else "Р"

    for r in range(3, 6):
        ws.row_dimensions[r].height = 22
        for c in range(1, 11):
            cell = ws.cell(row=r, column=c)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = border_cell

    current_row = 6
    for grp_idx, grp in enumerate(groups_data):
        label = grp["label"]
        g_rows = grp["rows"]
        grp_len = len(g_rows)
        grp_fill = fill_group_b if grp_idx % 2 == 1 else fill_group_a

        start_r = current_row
        end_r = current_row + grp_len - 1

        for r_offset, r_data in enumerate(g_rows):
            r_num = current_row + r_offset
            ws.row_dimensions[r_num].height = 20

            cell_cos = ws.cell(row=r_num, column=2, value=r_data["cos_phi"])
            cell_cos.font = font_data
            cell_cos.fill = grp_fill
            cell_cos.alignment = align_center
            cell_cos.border = border_cell

            vals = [
                r_data["i_1_act"], r_data["i_1_react"],
                r_data["i_5_act"], r_data["i_5_react"],
                r_data["i_20_act"], r_data["i_20_react"],
                r_data["i_100_act"], r_data["i_100_react"],
            ]
            for c_offset, v in enumerate(vals):
                cell_val = ws.cell(row=r_num, column=3 + c_offset, value=v)
                cell_val.font = font_data
                cell_val.fill = grp_fill
                cell_val.alignment = align_center
                cell_val.border = border_cell

        ws.merge_cells(start_row=start_r, start_column=1, end_row=end_r, end_column=1)
        for r_fill in range(start_r, end_r + 1):
            cell_a = ws.cell(row=r_fill, column=1)
            cell_a.fill = grp_fill
            cell_a.border = border_cell
        cell_first = ws.cell(row=start_r, column=1, value=label)
        cell_first.font = font_data
        cell_first.alignment = align_left

        current_row += grp_len

    ws.column_dimensions['A'].width = 38
    ws.column_dimensions['B'].width = 14
    for c_idx in range(3, 11):
        ws.column_dimensions[get_column_letter(c_idx)].width = 13

    note_r = current_row + 1
    ws.cell(row=note_r, column=1, value="Примечания:").font = Font(name="Calibri", size=9, bold=True, italic=True)
    ws.cell(row=note_r+1, column=1, value="1. Характеристики погрешности ИК даны для измерения электроэнергии и средней мощности (получасовая).").font = Font(name="Calibri", size=8.5, italic=True)
    ws.cell(row=note_r+2, column=1, value="2. В качестве характеристик относительной погрешности указаны границы интервала, соответствующие вероятности 0,95.").font = Font(name="Calibri", size=8.5, italic=True)

    wb.save(filepath)
    return filepath


def get_table51_tsv_text(rows: List[Dict[str, Any]], val_type: str = "osn") -> str:
    """Генерирует текст TSV в формате Таблицы 5.1 для вставки через буфер обмена."""
    groups_list = group_rows_for_table51(rows)
    groups_data = build_table51_data(groups_list, val_type=val_type)

    val_sym = "δWосн" if val_type == "osn" else "δWраб"
    lines = [
        f"Номера каналов\tЗначение cos φ (sin φ)\tδ1(2)%, акт (±{val_sym})\tδ1(2)%, реакт\tδ5%, акт\tδ5%, реакт\tδ20%, акт\tδ20%, реакт\tδ100%, акт\tδ100%, реакт"
    ]
    for grp in groups_data:
        label = grp["label"]
        for r in grp["rows"]:
            lines.append(f"{label}\t{r['cos_phi']}\t{r['i_1_act']}\t{r['i_1_react']}\t{r['i_5_act']}\t{r['i_5_react']}\t{r['i_20_act']}\t{r['i_20_react']}\t{r['i_100_act']}\t{r['i_100_react']}")
    return "\n".join(lines)

