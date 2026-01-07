"""
Конвертация PPTX в Markdown
Версия: pipeline-based
Цель: подготовка данных для обучения ИИ (RAG / fine-tuning)
"""

from pptx import Presentation
from dataclasses import dataclass
import os


# ============================================================
# МОДЕЛЬ ДАННЫХ
# ============================================================

@dataclass
class TextBlock:
    """
    Базовый текстовый блок с геометрией.
    Это наш "атом" — минимальная единица смысла.
    Текст + геометрия, а не просто строка
    """
    text: str
    left: int
    top: int
    width: int
    height: int


# ============================================================
# ЭТАП 1. ИЗВЛЕЧЕНИЕ СЫРЫХ БЛОКОВ
# ============================================================

def extract_raw_blocks(slide):
    """
    Извлекает все текстовые блоки со слайда вместе с координатами.
    НИКАКОЙ логики порядка или Markdown здесь нет.
    """
    blocks = []

    for shape in slide.shapes:
        # Нас интересуют только фигуры с текстом
        if not shape.has_text_frame:
            continue

        text = shape.text_frame.text.strip()
        if not text:
            continue

        blocks.append(
            TextBlock(
                text=text,
                left=shape.left,
                top=shape.top,
                width=shape.width,
                height=shape.height
            )
        )

    return blocks


# ============================================================
# ЭТАП 2. НОРМАЛИЗАЦИЯ И ПОРЯДОК ЧТЕНИЯ
# ============================================================

def order_blocks(blocks):
    """
    Восстанавливает порядок чтения:
    1) сверху вниз
    2) слева направо

    Используем округление координат,
    чтобы убрать микрошумы PowerPoint.
    """
    return sorted(
        blocks,
        key=lambda b: (round(b.top, -2), round(b.left, -2))
    )


# ============================================================
# ЭТАП 3. ОПРЕДЕЛЕНИЕ ЗАГОЛОВКОВ
# ============================================================

def is_title(block: TextBlock) -> bool:
    """
    Простая, но устойчивая эвристика заголовка.
    Её легко дорабатывать под реальные презентации.
    """
    lines = block.text.splitlines()

    return (
        len(lines) == 1 and          # одна строка
        len(block.text) < 120 and    # не длинный текст
        block.top < 300              # верхняя часть слайда
    )


# ============================================================
# ЭТАП 4. ГРУППИРОВКА В СМЫСЛОВЫЕ СЕКЦИИ
# ============================================================

def group_into_sections(blocks):
    """
    Группирует блоки по принципу:
    ЗАГОЛОВОК → ВСЁ, ЧТО ИДЁТ ПОСЛЕ НЕГО
    """
    sections = []
    current_section = None

    for block in blocks:
        if is_title(block):
            current_section = {
                "title": block.text,
                "content": []
            }
            sections.append(current_section)
        else:
            if current_section:
                current_section["content"].append(block.text)
            else:
                # Текст до первого заголовка (редко, но бывает)
                sections.append({
                    "title": None,
                    "content": [block.text]
                })

    return sections


# ============================================================
# ЭТАП 5. РЕНДЕРИНГ MARKDOWN
# ============================================================

def render_markdown(slide_number, sections):
    """
    Превращает логическую структуру в Markdown.
    Здесь нет PPTX — только текст.
    """
    md = []

    md.append(f"# СЛАЙД {slide_number}")
    md.append("")

    for section in sections:
        if section["title"]:
            md.append(f"## {section['title']}")
            md.append("")

        for text in section["content"]:
            # Разбиваем на строки, чтобы корректно обрабатывать списки
            for line in text.splitlines():
                line = line.strip()
                if not line:
                    continue

                if line.startswith(("•", "-")):
                    md.append(f"- {line[1:].strip()}")
                else:
                    md.append(line)

            md.append("")

    md.append("---")
    md.append("")

    return md


# ============================================================
# ОСНОВНАЯ ФУНКЦИЯ КОНВЕРТАЦИИ
# ============================================================

def convert_pptx_to_markdown(pptx_path):
    """
    Главный управляющий pipeline.
    """
    prs = Presentation(pptx_path)
    all_markdown = []

    for idx, slide in enumerate(prs.slides, start=1):
        print(f"Обрабатываю слайд {idx}")

        # 1. Извлечение
        raw_blocks = extract_raw_blocks(slide)

        # 2. Порядок чтения
        ordered_blocks = order_blocks(raw_blocks)

        # 3. Логическая структура
        sections = group_into_sections(ordered_blocks)

        # 4. Markdown
        slide_md = render_markdown(idx, sections)

        all_markdown.extend(slide_md)

    return all_markdown


# ============================================================
# СОХРАНЕНИЕ В ФАЙЛ
# ============================================================

def save_to_markdown(content, output_path):
    """
    Сохраняет результат в Markdown-файл.
    """
    with open(output_path, "w", encoding="utf-8") as f:
        for line in content:
            f.write(line + "\n")


# ============================================================
# ENTRY POINT
# ============================================================

def main():
    pptx_file = "presentation.pptx"
    output_file = "presentation.md"

    if not os.path.exists(pptx_file):
        print(f"Файл не найден: {pptx_file}")
        return

    print("Начинаю конвертацию PPTX → Markdown")
    print("=" * 50)

    markdown = convert_pptx_to_markdown(pptx_file)
    save_to_markdown(markdown, output_file)

    print("=" * 50)
    print("Готово!")
    print(f"Создан файл: {output_file}")
    print(f"Количество строк: {len(markdown)}")


if __name__ == "__main__":
    main()
