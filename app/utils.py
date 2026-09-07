import re

def escape_md(text: str) -> str:
    """Экранирование служебных символов Markdown."""
    if not text:
        return text
    chars = r'\`*_{}[]()#+-.!|'
    for ch in chars:
        text = text.replace(ch, '\\' + ch)
    return text

def get_heading_level(style_name: str) -> int:
    """Извлекает уровень заголовка из имени стиля (например, 'Заголовок 1' -> 1)."""
    match = re.search(r'(\d+)$', style_name)
    if match:
        return int(match.group(1))
    return 1 if ('заголовок' in style_name.lower() or 'heading' in style_name.lower()) else 0

def is_heading(style_name: str) -> bool:
    """Проверяет, является ли стиль заголовком (по имени)."""
    lower = style_name.lower()
    return 'заголовок' in lower or 'heading' in lower

def get_heading_level_by_formatting(paragraph):
    """
    Определяет уровень заголовка по форматированию (жирность и размер шрифта).
    Возвращает 1, 2, 3... или 0, если не похоже на заголовок.
    Упрощённо: если текст жирный и размер больше базового (11pt), считаем заголовком.
    """
    if not paragraph.runs:
        return 0
    # Проверяем, есть ли жирное начертание
    is_bold = any(run.bold for run in paragraph.runs if run.bold is not None)
    if not is_bold:
        return 0
    # Проверяем размер шрифта: берём первый run с размером
    font_size = None
    for run in paragraph.runs:
        if run.font.size:
            font_size = run.font.size.pt
            break
    if font_size is None:
        return 0
    # Если размер > 12 pt (условно), считаем заголовком
    # Уровень можно оценить по размеру: 16pt -> h1, 14pt -> h2, 13pt -> h3 и т.д.
    if font_size >= 16:
        return 1
    elif font_size >= 14:
        return 2
    elif font_size >= 13:
        return 3
    else:
        return 0
