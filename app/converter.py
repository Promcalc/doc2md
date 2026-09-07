import os
import tempfile
import subprocess
import re
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from app.utils import escape_md, get_heading_level, is_heading

# ---------- Работа со списками ----------
def get_list_info(paragraph):
    """
    Определяет, является ли абзац элементом списка.
    Возвращает (is_ordered, level) или (None, None) если не список.
    level начинается с 0.
    """
    p = paragraph._element
    numPr = p.find(qn('w:numPr'))
    if numPr is None:
        return None, None
    # Определяем тип списка по стилю
    style_name = paragraph.style.name.lower() if paragraph.style else ''
    if 'list number' in style_name or 'нумерованный' in style_name:
        is_ordered = True
    elif 'list bullet' in style_name or 'маркированный' in style_name:
        is_ordered = False
    else:
        # По умолчанию считаем маркированным
        is_ordered = False
    # Уровень
    ilvl = numPr.find(qn('w:ilvl'))
    level = int(ilvl.get(qn('w:val'))) if ilvl is not None else 0
    return is_ordered, level

def process_paragraph(para, footnote_map):
    """
    Обрабатывает один параграф, возвращает строку Markdown или None (если пустой).
    footnote_map: dict {footnote_id: sequential_number}
    """
    # Сначала обрабатываем runs: заменяем ссылки на сноски
    new_runs = []
    for run in para.runs:
        # Ищем элемент footnoteReference в run
        refs = run._element.findall(qn('w:footnoteReference'))
        if refs:
            for ref in refs:
                fn_id = ref.get(qn('w:id'))
                if fn_id and fn_id in footnote_map:
                    new_runs.append(f'[^{footnote_map[fn_id]}]')
                # Если id не найден, просто пропускаем
        else:
            if run.text:
                new_runs.append(run.text)
    text = ''.join(new_runs).strip()
    if not text:
        return None

    style_name = para.style.name if para.style else ''
    # Заголовки
    if is_heading(style_name):
        level = get_heading_level(style_name)
        return '#' * level + ' ' + escape_md(text)
    # Списки
    is_ordered, level = get_list_info(para)
    if is_ordered is not None:
        indent = '  ' * level
        if is_ordered:
            return indent + '1. ' + escape_md(text)
        else:
            return indent + '- ' + escape_md(text)
    # Обычный абзац
    return escape_md(text)

def process_table(table):
    """Преобразует таблицу Word в простую Markdown-таблицу."""
    # Строим матрицу, игнорируя вертикальные объединения (упрощаем)
    rows = []
    max_cols = 0
    for row in table.rows:
        row_cells = []
        for cell in row.cells:
            # Объединяем все параграфы ячейки через пробел
            cell_text = ' '.join(p.text.strip() for p in cell.paragraphs)
            if not cell_text:
                cell_text = ' '
            # Учитываем горизонтальное объединение (gridSpan)
            tc = cell._element
            grid_span_elem = tc.find(qn('w:gridSpan'))
            span = int(grid_span_elem.get(qn('w:val'))) if grid_span_elem is not None else 1
            # Добавляем ячейку span раз
            for _ in range(span):
                row_cells.append(escape_md(cell_text))
        rows.append(row_cells)
        if len(row_cells) > max_cols:
            max_cols = len(row_cells)
    # Дополняем строки до максимальной длины
    for row in rows:
        while len(row) < max_cols:
            row.append(' ')
    if not rows:
        return ''
    # Формируем Markdown-таблицу (первая строка — заголовок)
    md_lines = []
    md_lines.append('| ' + ' | '.join(rows[0]) + ' |')
    md_lines.append('| ' + ' | '.join(['---'] * len(rows[0])) + ' |')
    for row in rows[1:]:
        md_lines.append('| ' + ' | '.join(row) + ' |')
    return '\n'.join(md_lines)

def process_footnotes(doc):
    """
    Извлекает сноски из документа через XML.
    Возвращает (footnote_map, footnote_texts):
        footnote_map: dict {footnote_id (str): sequential_number (int)}
        footnote_texts: list of (sequential_number, text)
    """
    footnotes_part = None

    # Способ 1: поиск части по типу связи (если доступен)
    try:
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        footnotes_part = doc.part.package.part_related_by_type(
            RT.FOOTNOTES, doc.part
        )
    except (AttributeError, KeyError):
        pass

    # Способ 2: если не найден, ищем часть по имени файла
    if footnotes_part is None:
        try:
            for part in doc.part.package.iter_parts():
                if part.partname.endswith('/footnotes.xml'):
                    footnotes_part = part
                    break
        except AttributeError:
            pass

    # Если сноски не найдены, возвращаем пустые словари
    if footnotes_part is None:
        return {}, {}

    # Парсим XML
    root = footnotes_part.element
    footnote_elements = root.findall(qn('w:footnote'))
    footnotes = []
    for fn_elem in footnote_elements:
        fn_id = fn_elem.get(qn('w:id'))
        # Собираем текст из всех параграфов внутри сноски
        paragraphs = fn_elem.findall(qn('w:p'))
        text_parts = []
        for p in paragraphs:
            runs = p.findall(qn('w:r'))
            for r in runs:
                t = r.find(qn('w:t'))
                if t is not None and t.text:
                    text_parts.append(t.text)
        full_text = ' '.join(text_parts).strip()
        if full_text:
            footnotes.append((fn_id, full_text))

    # Нумеруем последовательно
    footnote_map = {}
    footnote_texts = []
    for idx, (fn_id, text) in enumerate(footnotes, start=1):
        footnote_map[fn_id] = idx
        footnote_texts.append((idx, escape_md(text)))
    return footnote_map, footnote_texts

def convert_docx_to_md(docx_path):
    """Основная функция конвертации .docx в строку Markdown."""
    doc = Document(docx_path)
    footnote_map, footnote_texts = process_footnotes(doc)

    output_lines = []
    # Обрабатываем параграфы
    for para in doc.paragraphs:
        md_line = process_paragraph(para, footnote_map)
        if md_line is not None:
            output_lines.append(md_line)
        else:
            output_lines.append('')  # пустая строка-разделитель

    # Обрабатываем таблицы
    for table in doc.tables:
        table_md = process_table(table)
        if table_md:
            output_lines.append(table_md)
            output_lines.append('')  # разделитель после таблицы

    # Добавляем сноски в конец
    if footnote_texts:
        output_lines.append('')
        for num, text in footnote_texts:
            output_lines.append(f'[^{num}]: {text}')

    # Убираем лишние пустые строки (не более одной подряд)
    final_lines = []
    prev_empty = False
    for line in output_lines:
        if line == '':
            if not prev_empty:
                final_lines.append('')
                prev_empty = True
        else:
            final_lines.append(line)
            prev_empty = False
    if final_lines and final_lines[-1] == '':
        final_lines.pop()
    return '\n'.join(final_lines)

def convert_doc_to_docx(doc_path):
    """Конвертирует .doc в .docx с помощью LibreOffice."""
    out_dir = tempfile.mkdtemp()
    out_file = os.path.join(out_dir, 'converted.docx')
    cmd = ['soffice', '--headless', '--convert-to', 'docx', '--outdir', out_dir, doc_path]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")
    # Находим созданный файл
    for f in os.listdir(out_dir):
        if f.endswith('.docx'):
            return os.path.join(out_dir, f)
    raise FileNotFoundError("Converted .docx not found")

def convert_file(input_path, output_path=None):
    """Основная точка входа для CLI и API."""
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"File not found: {input_path}")
    ext = os.path.splitext(input_path)[1].lower()
    if ext == '.doc':
        temp_docx = convert_doc_to_docx(input_path)
        md_content = convert_docx_to_md(temp_docx)
        # Удаляем временную папку
        import shutil
        shutil.rmtree(os.path.dirname(temp_docx), ignore_errors=True)
    elif ext == '.docx':
        md_content = convert_docx_to_md(input_path)
    else:
        raise ValueError("Unsupported file type. Only .doc and .docx are supported.")
    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + '.md'
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(md_content)
    return output_path