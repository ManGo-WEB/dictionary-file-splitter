"""
Модуль для парсинга словарного файла
"""
import re
from docx import Document

def is_section_header(text):
    """Проверяет, является ли текст заголовком раздела"""
    return bool(re.match(r'^[А-Я]', text))

def get_section_letter(text):
    """Извлекает букву раздела из заголовка"""
    return text[0]

def get_section_content(paragraph):
    """
    Получает содержимое параграфа, сохраняя форматирование
    
    Args:
        paragraph: Параграф из docx документа
    Returns:
        tuple: (текст без буквы раздела, исходный параграф)
    """
    text = paragraph.text
    return text[1:], paragraph