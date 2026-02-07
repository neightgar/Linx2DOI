"""
Парсинг документов Word для извлечения списка статей
"""

import re
from docx import Document
from docx.oxml.ns import qn


class DocumentParser:
    """Парсер для извлечения статей из .docx документов"""

    @staticmethod
    def extract_hyperlinks_from_paragraph(paragraph, doc):
        """Извлекает URL из гиперссылок в параграфе

        Args:
            paragraph: Объект параграфа python-docx
            doc: Document объект для доступа к relationships

        Returns:
            list: Список URL из гиперссылок
        """
        hyperlinks = []

        try:
            # Получаем relationships документа
            rels = doc.part.rels

            # Ищем гиперссылки в XML параграфа
            for hyperlink in paragraph._element.xpath('.//w:hyperlink'):
                # Получаем relationship ID
                rid = hyperlink.get(qn('r:id'))
                if rid and rid in rels:
                    # Получаем целевой URL
                    target_url = rels[rid].target_ref
                    if target_url:
                        hyperlinks.append(target_url)
        except Exception:
            pass

        return hyperlinks

    @staticmethod
    def normalize_title(title):
        """Нормализует заголовок для сравнения"""
        clean = re.sub(r'[^\w\s]', ' ', title.lower())
        clean = re.sub(r'\s+', ' ', clean).strip()
        return clean[:50] if len(clean) > 50 else clean

    @staticmethod
    def clean_url(url):
        """Исправляет пробелы внутри URL"""
        url = re.sub(r'\s*([/:])\s*', r'\1', url)
        url = url.strip().rstrip('.,;:')
        return url

    @staticmethod
    def extract_item_from_text(text, original_index):
        """Извлекает информацию о статье из текста

        Args:
            text: Текст строки из документа
            original_index: Порядковый номер записи

        Returns:
            dict или None: Словарь с данными статьи или None если формат не распознан
        """
        if not text.strip():
            return None

        # Паттерн: номер. название URL (допускаем пробелы в URL)
        pattern = r'^\s*\d+\.\s*(.*?)(https?://.+)$'
        match = re.match(pattern, text.strip())
        if not match:
            return None

        title_raw, url_raw = match.groups()
        url = DocumentParser.clean_url(url_raw)

        # Очистка заголовка от точек в конце
        title_clean = re.sub(r'[\s\.]*\.{2,}[\s\.]*$', '', title_raw.rstrip(' \t.')).strip()
        if not title_clean:
            title_clean = title_raw.strip()

        return {
            'original_index': original_index,
            'title': title_clean,
            'url': url,
            'normalized_title': DocumentParser.normalize_title(title_clean),
            'text': text.strip()
        }

    @staticmethod
    def get_paragraph_number(paragraph):
        """Извлекает номер из встроенной нумерации Word

        Args:
            paragraph: Объект параграфа python-docx

        Returns:
            int или None: Номер элемента списка или None
        """
        try:
            # Проверяем наличие встроенной нумерации Word
            numPr = paragraph._element.pPr.numPr if paragraph._element.pPr is not None else None
            if numPr is not None:
                # Получаем уровень нумерации (ilvl)
                ilvl_element = numPr.ilvl
                if ilvl_element is not None:
                    # Нумерация найдена - используем position_counter снаружи
                    return True
        except (AttributeError, KeyError):
            pass
        return None

    @staticmethod
    def collect_items_from_document(doc):
        """Собирает все элементы из документа с сохранением порядка

        Args:
            doc: Document объект из python-docx

        Returns:
            list: Список словарей с данными статей
        """
        items = []
        position_counter = 1

        # Обработка параграфов
        i = 0
        while i < len(doc.paragraphs):
            paragraph = doc.paragraphs[i]
            text = paragraph.text.strip()

            # Проверка на неполный URL (разбит на несколько параграфов)
            if text and re.search(r'https?://$', text):
                # URL обрывается на протокол - объединяем со следующим параграфом
                if i + 1 < len(doc.paragraphs):
                    next_para = doc.paragraphs[i + 1]
                    next_text = next_para.text.strip()

                    # Проверяем, что следующий параграф не имеет Word нумерации
                    next_has_word_num = DocumentParser.get_paragraph_number(next_para)

                    if not next_has_word_num and next_text:
                        # Объединяем параграфы
                        text = text + next_text
                        i += 1  # Пропускаем следующий параграф

            # Проверяем: есть ли номер в тексте или встроенная нумерация Word
            has_text_number = text and re.match(r'^\s*\d+\.\s', text)
            has_word_numbering = DocumentParser.get_paragraph_number(paragraph)

            # Извлекаем гиперссылки из параграфа (приоритет!)
            hyperlinks = DocumentParser.extract_hyperlinks_from_paragraph(paragraph, doc)

            if has_text_number:
                # Обычная текстовая нумерация
                if hyperlinks:
                    # Есть гиперссылка - используем её URL
                    # Извлекаем название из текста (до URL)
                    title_match = re.match(r'^\s*\d+\.\s*(.+?)(?:\s+https?://.*)?$', text)
                    title = title_match.group(1).strip() if title_match else text
                    # Создаем item с URL из гиперссылки
                    numbered_text = f"{position_counter}. {title} {hyperlinks[0]}"
                    item = DocumentParser.extract_item_from_text(numbered_text, position_counter)
                else:
                    # Нет гиперссылки - парсим текст как обычно
                    item = DocumentParser.extract_item_from_text(text, position_counter)

                if item:
                    items.append(item)
                    position_counter += 1

            elif has_word_numbering and text:
                # Встроенная нумерация Word
                if hyperlinks:
                    # Есть гиперссылка - используем её URL
                    # Извлекаем название из текста (убираем URL если есть)
                    title = re.sub(r'\s+https?://.*$', '', text).strip()
                    # Создаем item с URL из гиперссылки
                    numbered_text = f"{position_counter}. {title} {hyperlinks[0]}"
                    item = DocumentParser.extract_item_from_text(numbered_text, position_counter)
                else:
                    # Нет гиперссылки - добавляем номер к тексту
                    numbered_text = f"{position_counter}. {text}"
                    item = DocumentParser.extract_item_from_text(numbered_text, position_counter)

                if item:
                    items.append(item)
                    position_counter += 1

            i += 1

        # Обработка таблиц
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        text = paragraph.text.strip()

                        has_text_number = text and re.match(r'^\s*\d+\.\s', text)
                        has_word_numbering = DocumentParser.get_paragraph_number(paragraph)

                        # Извлекаем гиперссылки из параграфа
                        hyperlinks = DocumentParser.extract_hyperlinks_from_paragraph(paragraph, doc)

                        if has_text_number:
                            if hyperlinks:
                                # Есть гиперссылка - используем её URL
                                title_match = re.match(r'^\s*\d+\.\s*(.+?)(?:\s+https?://.*)?$', text)
                                title = title_match.group(1).strip() if title_match else text
                                numbered_text = f"{position_counter}. {title} {hyperlinks[0]}"
                                item = DocumentParser.extract_item_from_text(numbered_text, position_counter)
                            else:
                                item = DocumentParser.extract_item_from_text(text, position_counter)

                            if item:
                                items.append(item)
                                position_counter += 1

                        elif has_word_numbering and text:
                            if hyperlinks:
                                # Есть гиперссылка - используем её URL
                                title = re.sub(r'\s+https?://.*$', '', text).strip()
                                numbered_text = f"{position_counter}. {title} {hyperlinks[0]}"
                                item = DocumentParser.extract_item_from_text(numbered_text, position_counter)
                            else:
                                numbered_text = f"{position_counter}. {text}"
                                item = DocumentParser.extract_item_from_text(numbered_text, position_counter)

                            if item:
                                items.append(item)
                                position_counter += 1

        return items
