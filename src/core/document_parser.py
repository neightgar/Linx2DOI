"""
Парсинг документов Word для извлечения списка статей
"""

import re
from docx import Document


class DocumentParser:
    """Парсер для извлечения статей из .docx документов"""

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

        # Паттерн: номер. название URL
        pattern = r'^\s*\d+\.\s*(.*?)(https?://\S.*)$'
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
        for paragraph in doc.paragraphs:
            text = paragraph.text.strip()
            if text and re.match(r'^\s*\d+\.\s', text):
                item = DocumentParser.extract_item_from_text(text, position_counter)
                if item:
                    items.append(item)
                    position_counter += 1

        # Обработка таблиц
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        text = paragraph.text.strip()
                        if text and re.match(r'^\s*\d+\.\s', text):
                            item = DocumentParser.extract_item_from_text(text, position_counter)
                            if item:
                                items.append(item)
                                position_counter += 1

        return items
