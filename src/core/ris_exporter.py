"""
Генератор RIS файлов для экспорта в Mendeley
"""

import re
from typing import List, Dict


class RISExporter:
    """Генератор файлов в формате RIS для менеджеров библиографий"""

    @staticmethod
    def generate_ris(items: List[Dict], filename: str) -> str:
        """
        Генерирует RIS файл из списка обработанных элементов

        Args:
            items: Список элементов с метаданными
            filename: Имя исходного файла

        Returns:
            str: Содержимое RIS файла
        """
        ris_content = []

        for item in items:
            # Пропускаем элементы без DOI
            if not item.get('dois') or len(item['dois']) == 0:
                continue

            # Пропускаем дубли
            if 'ДУБЛЬ' in item.get('status', ''):
                continue

            # Начало записи - тип публикации
            ris_content.append("TY  - JOUR")  # Journal Article

            # Авторы (если есть в метаданных)
            authors = RISExporter._extract_authors(item)
            for author in authors:
                ris_content.append(f"AU  - {author}")

            # Заголовок статьи
            article_title = item.get('article_title', item.get('title', 'Unknown'))
            if article_title and article_title != 'Не найден':
                ris_content.append(f"TI  - {article_title}")

            # Название журнала (из метаданных или оригинального названия)
            journal = RISExporter._extract_journal(item)
            if journal:
                ris_content.append(f"JO  - {journal}")

            # Год публикации
            year = RISExporter._extract_year(item)
            if year:
                ris_content.append(f"PY  - {year}")

            # Том
            volume = RISExporter._extract_volume(item)
            if volume:
                ris_content.append(f"VL  - {volume}")

            # Выпуск
            issue = RISExporter._extract_issue(item)
            if issue:
                ris_content.append(f"IS  - {issue}")

            # Страницы
            pages = RISExporter._extract_pages(item)
            if pages:
                if '-' in pages:
                    start, end = pages.split('-', 1)
                    ris_content.append(f"SP  - {start.strip()}")
                    ris_content.append(f"EP  - {end.strip()}")
                else:
                    ris_content.append(f"SP  - {pages}")

            # DOI
            doi = item['dois'][0]  # Первый DOI
            ris_content.append(f"DO  - {doi}")

            # URL на DOI
            ris_content.append(f"UR  - https://doi.org/{doi}")

            # Исходный URL статьи (если отличается от DOI)
            if item.get('url') and 'doi.org' not in item['url']:
                ris_content.append(f"UR  - {item['url']}")

            # Аннотация (если есть)
            abstract = RISExporter._extract_abstract(item)
            if abstract:
                ris_content.append(f"AB  - {abstract}")

            # Конец записи
            ris_content.append("ER  - ")
            ris_content.append("")  # Пустая строка между записями

        return "\n".join(ris_content)

    @staticmethod
    def _extract_authors(item: Dict) -> List[str]:
        """Извлекает список авторов из метаданных"""
        authors = []
        apa = item.get('apa_citation', '')

        if apa:
            # Парсим авторов из APA цитаты
            # Формат: Author, A. B., & Author2, C. D. (Year). Title...
            match = re.match(r'^([^(]+)\s*\(\d{4}\)', apa)
            if match:
                author_string = match.group(1).strip()
                # Разделяем по ", &" и по ","
                author_parts = re.split(r',\s*&\s*|,\s*(?=[A-Z]\.)', author_string)
                for author in author_parts:
                    author = author.strip()
                    if author and len(author) > 2:
                        authors.append(author)

        return authors

    @staticmethod
    def _extract_journal(item: Dict) -> str:
        """Извлекает название журнала из метаданных"""
        apa = item.get('apa_citation', '')

        if apa:
            # Парсим журнал из APA: ...Title. Journal Name, volume(issue), pages.
            # Ищем паттерн после точки и перед запятой с номером
            match = re.search(r'\.\s+([^.]+?),\s*\d+', apa)
            if match:
                return match.group(1).strip()

        return item.get('title', '')

    @staticmethod
    def _extract_year(item: Dict) -> str:
        """Извлекает год публикации"""
        apa = item.get('apa_citation', '')

        if apa:
            # Год в скобках после авторов: (2024)
            match = re.search(r'\((\d{4})\)', apa)
            if match:
                return match.group(1)

        return ""

    @staticmethod
    def _extract_volume(item: Dict) -> str:
        """Извлекает том публикации"""
        apa = item.get('apa_citation', '')

        if apa:
            # Том: Journal, 15(3), pages - ищем число перед скобкой
            match = re.search(r',\s*(\d+)\(', apa)
            if match:
                return match.group(1)

        return ""

    @staticmethod
    def _extract_issue(item: Dict) -> str:
        """Извлекает номер выпуска"""
        apa = item.get('apa_citation', '')

        if apa:
            # Выпуск в скобках: 15(3)
            match = re.search(r'\((\d+)\)', apa)
            if match:
                # Проверяем, что это не год
                potential_issue = match.group(1)
                if len(potential_issue) <= 3:  # Год обычно 4 цифры
                    return potential_issue

        return ""

    @staticmethod
    def _extract_pages(item: Dict) -> str:
        """Извлекает диапазон страниц"""
        apa = item.get('apa_citation', '')

        if apa:
            # Страницы в конце: , 123-145. или , e12345.
            match = re.search(r',\s*([\d\-e]+)\.\s*(?:https?://|$)', apa)
            if match:
                return match.group(1)

        return ""

    @staticmethod
    def _extract_abstract(item: Dict) -> str:
        """Извлекает аннотацию (если есть в метаданных)"""
        # Пока аннотация не сохраняется, можно добавить в будущем
        return ""

    @staticmethod
    def save_ris_file(items: List[Dict], filepath: str, source_filename: str) -> bool:
        """
        Сохраняет RIS файл на диск

        Args:
            items: Список элементов для экспорта
            filepath: Путь для сохранения
            source_filename: Имя исходного файла

        Returns:
            bool: True если успешно сохранено
        """
        try:
            ris_content = RISExporter.generate_ris(items, source_filename)

            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(ris_content)

            return True
        except Exception as e:
            print(f"Ошибка сохранения RIS файла: {e}")
            return False
