"""
Рабочий поток для фоновой обработки документов
"""

import time
from pathlib import Path
from docx import Document
from PyQt6.QtCore import QThread, pyqtSignal

from .config import RATE_LIMIT_DELAY
from .document_parser import DocumentParser
from .doi_resolver import DOIResolver
from .citation_formatter import CitationFormatter
from .html_generator import HTMLGenerator


class WorkerThread(QThread):
    """Поток для фоновой обработки документов."""
    progress = pyqtSignal(int, str)
    log = pyqtSignal(str)
    finished_success = pyqtSignal(str, list)
    finished_error = pyqtSignal(str)
    table_data = pyqtSignal(list)
    item_updated = pyqtSignal(dict)

    def __init__(self, input_path, user_email, selected_items=None, apply_manual_only=False,
                 previous_items=None):
        super().__init__()
        self.input_path = input_path
        self.user_email = user_email
        self.selected_items = selected_items if selected_items else []
        self.apply_manual_only = apply_manual_only
        self.previous_items = previous_items if previous_items else []

        # Инициализация модулей
        self.doi_resolver = DOIResolver(user_email, self.log.emit)
        self.citation_formatter = CitationFormatter(user_email, self.log.emit)

        self.total_items_to_process = 0
        self.processed_items = 0

    def process_item(self, item, manual_doi=None):
        """Обрабатывает один элемент и возвращает результат

        Args:
            item: Словарь с данными статьи
            manual_doi: Ручной DOI (если указан)

        Returns:
            dict: Результат обработки
        """
        # Разрешаем DOI
        auto_search = not self.apply_manual_only or not manual_doi
        dois_list = self.doi_resolver.resolve_doi(item, manual_doi, auto_search)

        article_title = None
        apa_citation = None

        if dois_list:
            primary_doi = dois_list[0]
            self.log.emit(f"  📊 Получение метаданных для DOI: {primary_doi}")
            article_title = self.citation_formatter.get_article_title_from_doi(primary_doi)
            apa_citation = self.citation_formatter.get_apa_citation(primary_doi)
            time.sleep(RATE_LIMIT_DELAY)  # Задержка для избежания rate limit
        else:
            self.log.emit(f"  ⚠️ DOI не найден")

        has_data = len(dois_list) > 0 or apa_citation is not None

        result = {
            'original_index': item['original_index'],
            'title': item['title'],
            'url': item['url'],
            'dois': dois_list,
            'article_title': article_title,
            'apa_citation': apa_citation,
            'has_data': has_data,
            'manual_doi': manual_doi if manual_doi else None
        }

        status = "✅" if has_data else "⚠️"
        source = "РУЧНОЙ DOI" if manual_doi else ("АВТОМАТИЧЕСКИ" if has_data else "НЕТ ДАННЫХ")
        self.log.emit(f"  {status} Итог: {source}")

        return result

    def process_only_manual_items(self, all_items, manual_dois_map):
        """Обрабатывает только элементы с ручными DOI, сохраняя остальные

        Args:
            all_items: Все элементы из документа
            manual_dois_map: Словарь {original_index: manual_doi}

        Returns:
            list: Список результатов обработки
        """
        self.log.emit(f"\n🔧 ОБРАБОТКА ТОЛЬКО РУЧНЫХ DOI")
        self.log.emit(f"Найдено {len(manual_dois_map)} записей с ручными DOI")

        results = []
        self.total_items_to_process = len(manual_dois_map)
        self.processed_items = 0

        # Обрабатываем только элементы с ручными DOI
        processed_indices = set()

        for original_index, manual_doi in manual_dois_map.items():
            # Находим элемент по оригинальному индексу
            item_to_process = None
            for item in all_items:
                if item['original_index'] == original_index:
                    item_to_process = item
                    break

            if item_to_process:
                self.log.emit(f"\n🔄 ОБРАБОТКА ЗАПИСИ #{original_index} С РУЧНЫМ DOI")
                result = self.process_item(item_to_process, manual_doi)
                results.append(result)
                processed_indices.add(original_index)

                # Обновляем прогресс
                self.processed_items += 1
                progress = int((self.processed_items / self.total_items_to_process) * 100)
                self.progress.emit(progress, f"Ручные DOI: {self.processed_items}/{self.total_items_to_process}")

        # Добавляем необработанные элементы (без изменений)
        for item in all_items:
            if item['original_index'] not in processed_indices:
                # Находим соответствующий результат в предыдущих данных
                existing_result = None
                for prev_item in self.previous_items:
                    if prev_item.get('original_index') == item['original_index']:
                        existing_result = prev_item
                        break

                if existing_result:
                    results.append(existing_result)
                    self.log.emit(f"📋 Сохранена запись #{item['original_index']} (без изменений)")
                else:
                    # Если нет предыдущих данных, создаем новый результат
                    result = self.process_item(item)
                    results.append(result)

        # Сортируем по оригинальному индексу
        results.sort(key=lambda x: x.get('original_index', 0))

        self.log.emit(f"\n✅ ОБРАБОТКА ЗАВЕРШЕНА")
        self.log.emit(f"Обработано записей с ручными DOI: {len(manual_dois_map)}")
        self.log.emit(f"Всего записей в результате: {len(results)}")

        return results

    def run(self):
        """Главная функция обработки"""
        try:
            self.log.emit(f"⚙️ Загрузка документа...")
            doc = Document(str(self.input_path))

            # Собираем все элементы из документа
            all_items = DocumentParser.collect_items_from_document(doc)

            if not all_items:
                self.finished_error.emit("⚠️ Не найдено пунктов для обработки")
                return

            # Если это обработка только ручных DOI
            if self.apply_manual_only and self.selected_items:
                self.log.emit(f"\n🔧 РЕЖИМ: ОБРАБОТКА ТОЛЬКО РУЧНЫХ DOI")

                # Создаем карту ручных DOI
                manual_dois_map = {}
                for selected_item in self.selected_items:
                    if 'original_item' in selected_item and selected_item.get('manual_doi'):
                        original_index = selected_item['original_item'].get('original_index')
                        if original_index:
                            manual_dois_map[original_index] = selected_item['manual_doi']
                            self.log.emit(f"Запись #{original_index}: ручной DOI {selected_item['manual_doi']}")

                if not manual_dois_map:
                    self.finished_error.emit("⚠️ Не найдено ручных DOI для обработки")
                    return

                # Обрабатываем только элементы с ручными DOI
                results = self.process_only_manual_items(all_items, manual_dois_map)

            # Если это обработка выбранных элементов (повторный парсинг)
            elif self.selected_items:
                self.log.emit(f"\n🔄 РЕЖИМ: ПОВТОРНЫЙ ПАРСИНГ ВЫБРАННЫХ")
                self.log.emit(f"Выбрано {len(self.selected_items)} записей")

                results = []
                self.total_items_to_process = len(self.selected_items)
                self.processed_items = 0

                # Создаем карту для быстрого поиска
                items_map = {item['original_index']: item for item in all_items}

                for i, selected_item in enumerate(self.selected_items):
                    if 'original_item' in selected_item:
                        original_index = selected_item['original_item'].get('original_index')
                        if original_index in items_map:
                            item_to_process = items_map[original_index]
                            manual_doi = selected_item.get('manual_doi')

                            self.log.emit(f"\n🔄 ОБРАБОТКА ЗАПИСИ #{original_index}")
                            result = self.process_item(item_to_process, manual_doi)
                            results.append(result)

                            # Обновляем прогресс
                            self.processed_items += 1
                            progress = int((self.processed_items / self.total_items_to_process) * 100)
                            self.progress.emit(progress,
                                               f"Повторный парсинг: {self.processed_items}/{self.total_items_to_process}")

                # Добавляем необработанные элементы из предыдущих результатов
                processed_indices = {item['original_index'] for item in results}
                for prev_item in self.previous_items:
                    if prev_item.get('original_index') not in processed_indices:
                        results.append(prev_item)
                        self.log.emit(f"📋 Сохранена запись #{prev_item.get('original_index')} (без изменений)")

                # Сортируем по оригинальному индексу
                results.sort(key=lambda x: x.get('original_index', 0))

            # Стандартная обработка всего документа
            else:
                self.log.emit(f"\n🚀 РЕЖИМ: ПОЛНАЯ ОБРАБОТКА ДОКУМЕНТА")
                self.log.emit(f"Найдено {len(all_items)} записей")

                results = []
                self.total_items_to_process = len(all_items)
                self.processed_items = 0

                for i, item in enumerate(all_items):
                    self.log.emit(f"\n📄 ЗАПИСЬ #{item['original_index']} ИЗ {len(all_items)}")
                    result = self.process_item(item)
                    results.append(result)

                    # Обновляем прогресс
                    self.processed_items += 1
                    progress = int((self.processed_items / self.total_items_to_process) * 100)
                    self.progress.emit(progress, f"Обработка: {self.processed_items}/{self.total_items_to_process}")

            # Обновляем таблицу
            table_data = HTMLGenerator.generate_table_data(results)
            self.table_data.emit(table_data)

            # Генерируем HTML
            self.progress.emit(90, "Генерация HTML...")
            self.log.emit("\n🎨 ГЕНЕРАЦИЯ HTML-РЕЗУЛЬТАТА...")

            if self.apply_manual_only:
                output_path = self.input_path.parent / f"{self.input_path.stem}_apa_manual.html"
            elif self.selected_items and not self.apply_manual_only:
                output_path = self.input_path.parent / f"{self.input_path.stem}_apa_updated.html"
            else:
                output_path = self.input_path.parent / f"{self.input_path.stem}_apa.html"

            html_content = HTMLGenerator.generate_html_ordered(results, self.input_path.stem, self.user_email)

            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(html_content)

            self.progress.emit(100, "Готово!")
            self.log.emit(f"\n✅ ОБРАБОТКА УСПЕШНО ЗАВЕРШЕНА!")
            self.log.emit(f"📁 Результат сохранен: {output_path.name}")
            self.finished_success.emit(str(output_path), results)

        except Exception as e:
            import traceback
            error_msg = f"❌ КРИТИЧЕСКАЯ ОШИБКА:\n{str(e)}\n\n{traceback.format_exc()}"
            self.log.emit(f"\n{error_msg}")
            self.finished_error.emit(error_msg)
