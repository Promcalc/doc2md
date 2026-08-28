import re

def escape_md(text: str) -> str:
    """Экранирование служебных символов Markdown."""
    if not text:
        return text
    # Порядок важен: сначала экранируем обратный слеш, затем остальные
    chars = r'\`*_{}[]()#+-.!|'
    for ch in chars:
        text = text.replace(ch, '\\' + ch)
    return text

def get_heading_level(style_name: str) -> int:
    """Извлекает уровень заголовка из имени стиля (например, 'Заголовок 1' -> 1)."""
    match = re.search(r'(\d+)$', style_name)
    if match:
        return int(match.group(1))
    # Если есть слово 'заголовок' но без цифры, считаем уровнем 1
    return 1 if 'заголовок' in style_name.lower() else 0

def is_heading(style_name: str) -> bool:
    """Проверяет, содержит ли имя стиля слово 'заголовок' (регистронезависимо)."""
    return 'заголовок' in style_name.lower()