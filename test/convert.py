# -*- coding: utf-8 -*-
"""
Скрипт для конвертации PPTX в Markdown
"""

from pptx import Presentation
import os

def extract_text_from_pptx(pptx_path):
    """
    Извлекает текст из презентации PPTX
    """
    print(f"Открываю файл: {pptx_path}")
    
    # Открываем презентацию
    prs = Presentation(pptx_path)
    
    print(f"Найдено слайдов: {len(prs.slides)}")
    
    # Список для хранения содержимого всех слайдов
    all_slides_text = []
    
    # Проходим по всем слайдам
    for i, slide in enumerate(prs.slides, 1):
        print(f"Обрабатываю слайд {i}...")
        
        slide_content = []
        
        # Добавляем заголовок слайда
        slide_content.append(f"# СЛАЙД {i}")
        slide_content.append("")  # Пустая строка
        
        # Проверяем, есть ли у слайда заголовок
        if slide.shapes.title:
            title = slide.shapes.title.text.strip()
            if title:
                slide_content.append(f"## {title}")
                slide_content.append("")
        
        # Проходим по всем фигурам на слайде
        for shape in slide.shapes:
            # Проверяем, есть ли у фигуры текст
            if hasattr(shape, "text_frame"):
                text_frame = shape.text_frame
                
                # Игнорируем заголовок слайда (его мы уже обработали)
                if shape == slide.shapes.title:
                    continue
                
                # Получаем текст из текстового фрейма
                text = text_frame.text.strip()
                
                if text:
                    # Разделяем текст на параграфы
                    paragraphs = text.split('\n')
                    
                    for paragraph in paragraphs:
                        paragraph = paragraph.strip()
                        if paragraph:
                            # Проверяем, является ли абзац маркированным списком
                            # (обычно в PowerPoint маркированные списки начинаются с • или -)
                            if paragraph.startswith('•') or paragraph.startswith('-'):
                                slide_content.append(f"- {paragraph[1:].strip()}")
                            else:
                                slide_content.append(paragraph)
                    
                    slide_content.append("")  # Пустая строка после блока текста
        
        # Добавляем разделитель между слайдами
        slide_content.append("---")
        slide_content.append("")
        
        # Добавляем содержимое слайда к общему результату
        all_slides_text.extend(slide_content)
    
    return all_slides_text

def save_to_markdown(content, output_path):
    """
    Сохраняет текст в Markdown файл
    """
    print(f"Сохраняю результат в: {output_path}")
    
    with open(output_path, 'w', encoding='utf-8') as f:
        for line in content:
            f.write(line + '\n')
    
    print("Готово!")

def main():
    """
    Основная функция
    """
    # Указываем пути к файлам
    pptx_file = "presentation.pptx"
    output_file = "presentation.md"
    
    # Проверяем, существует ли файл презентации
    if not os.path.exists(pptx_file):
        print(f"ОШИБКА: Файл '{pptx_file}' не найден!")
        print("Убедитесь, что:")
        print(f"1. Файл находится в папке: {os.getcwd()}")
        print("2. Файл называется 'presentation.pptx'")
        print("3. Вы запускаете скрипт из той же папки, где лежит файл")
        return
    
    # Извлекаем текст из презентации
    print("=" * 50)
    print("Начинаю конвертацию PPTX в Markdown...")
    print("=" * 50)
    
    markdown_content = extract_text_from_pptx(pptx_file)
    
    # Сохраняем в Markdown файл
    save_to_markdown(markdown_content, output_file)
    
    # Показываем статистику
    print("\n" + "=" * 50)
    print("КОНВЕРТАЦИЯ ЗАВЕРШЕНА!")
    print("=" * 50)
    print(f"Создан файл: {output_file}")
    print(f"Количество строк: {len(markdown_content)}")
    
    # Показываем первые 10 строк для проверки
    print("\nПервые 10 строк результата:")
    print("-" * 30)
    for i, line in enumerate(markdown_content[:10]):
        print(f"{i+1:2}: {line}")

# Запускаем основную функцию
if __name__ == "__main__":
    main()