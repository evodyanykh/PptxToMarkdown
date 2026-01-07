from typing import List
from models import Document, TextBlock

class MarkdownRenderer:
    """
    Преобразует Document → Markdown

     Особенности:
    - Заголовки используются как структура (section)
    - Контент выводится под заголовками
    - Поддерживает маркеры списков
    - Добавляет разделитель слайдов
    """

    def render(self, document: Document) -> str:
        lines: List[str] = []

        for slide in document.slides:
            # Разделитель слайдов
            lines.append(f"\n---\n<!-- Slide {slide.slide_index} -->\n")
            
            if slide.title:
                lines.append(f"# {slide.title}\n")

            last_section = None

            for block in slide.blocks:
                # Добавляем заголовок секции один раз
                #section = getattr(block, "section", None)
                #if section and section != last_section and section != "Без заголовка":
                #    lines.append(f"## {section}\n")
                #    last_section = section

                # Добавляем содержимое блока  
                # Рендерим ТОЛЬКО контентные блоки
                lines.extend(self._render_block(block))

        return "\n".join(lines)

    def _render_block(self, block: TextBlock) -> List[str]:
        """
        Преобразует один TextBlock в список строк Markdown.
        Рендерит только КОНТЕНТ.
        Заголовочные блоки (is_title=True) пропускаются,
        т.к. они уже использованы для структуры документа.
        """
        # ⛔ НЕ рендерим title-блоки
        if getattr(block, "is_title", False):
            return []
        
        result = []
        text = block.text.strip()

        if not text:
            return result

        # Разделяем на строки
        for line in text.split("\n"):
            line = line.strip()
            if not line:
                continue
            # Маркированный список
            if line.startswith("•") or line.startswith("-"):
                result.append(f"- {line[1:].strip()}")
            else:
                result.append(line)

        result.append("")  # пустая строка после блока
        return result
