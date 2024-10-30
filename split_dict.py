"""
Главный модуль для разделения словаря на файлы с сохранением форматирования
"""
import sys
from dictionary_parser import is_section_header, get_section_letter, get_section_content
from file_handler import read_dictionary_file, save_section_to_file

def split_dictionary(input_file, output_dir='.'):
    """
    Разделяет словарь Word на отдельные файлы по разделам
    
    Args:
        input_file (str): Путь к входному файлу .docx
        output_dir (str): Директория для выходных файлов
    """
    try:
        doc = read_dictionary_file(input_file)
    except Exception as e:
        print(f"Ошибка: {str(e)}")
        return

    current_letter = None
    current_content = []
    
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:  # Пропускаем пустые строки
            continue
            
        if is_section_header(text):
            if current_letter:
                save_section_to_file(current_letter, current_content, output_dir)
            current_letter = get_section_letter(text)
            content, para = get_section_content(paragraph)
            current_content = [(content, para)]
        else:
            if current_letter:
                current_content.append((text, paragraph))
    
    # Сохраняем последний раздел
    if current_letter:
        save_section_to_file(current_letter, current_content, output_dir)

def main():
    """Точка входа в программу"""
    if len(sys.argv) != 2:
        print("Использование: python split_dict.py путь_к_файлу.docx")
        sys.exit(1)
        
    input_file = sys.argv[1]
    if not input_file.endswith('.docx'):
        print("Ошибка: файл должен быть в формате .docx")
        sys.exit(1)
        
    split_dictionary(input_file)

if __name__ == "__main__":
    main()