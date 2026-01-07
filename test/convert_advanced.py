# -*- coding: utf-8 -*-
"""
Исправленный скрипт для конвертации PPTX в Markdown с сохранением структуры
"""

from pptx import Presentation
import os
import re

def extract_text_with_structure(pptx_path):
    """
    Извлекает текст с сохранением структуры (списки, отступы)
    """
    prs = Presentation(pptx_path)
    
    markdown_lines = []
    
    # Добавляем заголовок документа
    markdown_lines.append("# Презентация")
    markdown_lines.append("")
    
    for slide_num, slide in enumerate(prs.slides, 1):
        # Заголовок слайда
        markdown_lines.append(f"## Слайд {slide_num}")
        
        # Заголовок слайда из презентации
        if slide.shapes.title:
            title = slide.shapes.title.text.strip()
            if title:
                markdown_lines.append(f"**Заголовок:** {title}")
        
        markdown_lines.append("")
        
        # Собираем весь текст слайда
        slide_text = []
        
        for shape in slide.shapes:
            if hasattr(shape, "text_frame"):
                # Пропускаем заголовок слайда
                if shape == slide.shapes.title:
                    continue
                
                text_frame = shape.text_frame
                
                for paragraph in text_frame.paragraphs:
                    text = paragraph.text.strip()
                    if text:
                        # Определяем уровень отступа (для вложенных списков)
                        level = paragraph.level
                        
                        # ПРАВИЛЬНО: проверяем, является ли это маркированным списком
                        # В объекте paragraph есть свойство bullet
                        bullet_char = ""
                        if hasattr(paragraph, 'bullet') and paragraph.bullet:
                            bullet_char = "-"
                        
                        # Добавляем отступы для вложенных пунктов
                        indent = "  " * level
                        slide_text.append(f"{indent}{bullet_char} {text}")
        
        # Если на слайде есть текст, добавляем его
        if slide_text:
            markdown_lines.extend(slide_text)
        else:
            markdown_lines.append("*(Слайд без текста)*")
        
        markdown_lines.append("")
        markdown_lines.append("---")
        markdown_lines.append("")
    
    return markdown_lines

def clean_markdown_text(lines):
    """
    Очищает и форматирует текст для лучшего восприятия ИИ
    """
    cleaned = []
    
    for line in lines:
        # Удаляем лишние пробелы
        line = re.sub(r'\s+', ' ', line.strip())
        
        # Заменяем несколько дефисов на один
        line = re.sub(r'-{2,}', '---', line)
        
        # Убираем пустые строки подряд (оставляем максимум 1)
        if not line and cleaned and not cleaned[-1]:
            continue
        
        cleaned.append(line)
    
    return cleaned

def save_markdown_file(content, filename):
    """
    Сохраняет результат в файл
    """
    with open(filename, 'w', encoding='utf-8') as f:
        for line in content:
            f.write(line + '\n')
    
    print(f"Файл сохранен: {filename}")
    print(f"Общее количество строк: {len(content)}")

def main():
    # Файлы
    input_file = "presentation.pptx"
    output_file = "presentation_cleaned.md"
    
    # Проверяем существование файла
    if not os.path.exists(input_file):
        print(f"Файл '{input_file}' не найден!")
        return
    
    print("Начинаю обработку презентации...")
    
    # Извлекаем текст
    raw_text = extract_text_with_structure(input_file)
    
    # Очищаем и форматируем
    cleaned_text = clean_markdown_text(raw_text)
    
    # Сохраняем
    save_markdown_file(cleaned_text, output_file)
    
    # Показываем пример
    print("\nПример содержимого (первые 20 строк):")
    print("=" * 50)
    for i, line in enumerate(cleaned_text[:20]):
        print(f"{i+1:3}: {line}")

if __name__ == "__main__":
    main()