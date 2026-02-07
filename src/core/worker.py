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

    def __init__(self, input_path, user_email, selected_items=None, previous_items=None):
        super().__init__()
        self.input_path = input_path
        self.user_email = user_email
        self.selected_items = selected_items if selected_items else []
        self.previous_items = previous_items if previous_items else []
        self.is_user_selection = bool(selected_items)  # Флаг явного выбора пользователя

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
        # Проверка: пропускаем автопоиск для нецитируемых записей
        # НО ТОЛЬКО если это НЕ явный выбор пользователя
        if not item.get('is_cited', True) and not manual_doi and not self.is_user_selection:
            self.log.emit(f"  ⚠️ Пропуск автопоиска: запись отсутствует в тексте")
            return {
                'original_index': item['original_index'],
                'title': item['title'],
                'url': item['url'],
                'dois': [],
                'article_title': None,
                'apa_citation': None,
                'has_data': False,
                'manual_doi': None,
                'is_cited': False,
                'uncited_reason': 'NOT_IN_TEXT'
            }

        # Если пользователь явно выбрал запись - обрабатываем независимо от is_cited
        if self.is_user_selection and not item.get('is_cited', True):
            self.log.emit(f"  🎯 Явный выбор пользователя: обработка несмотря на статус 'не цитируется'")

        # Разрешаем DOI
        # Если есть ручной DOI, не делаем автопоиск
        auto_search = not manual_doi

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
            'manual_doi': manual_doi if manual_doi else None,
            'is_cited': item.get('is_cited', True)
        }

        status = "✅" if has_data else "⚠️"
        source = "РУЧНОЙ DOI" if manual_doi else ("АВТОМАТИЧЕСКИ" if has_data else "НЕТ ДАННЫХ")
        self.log.emit(f"  {status} Итог: {source}")

        return result


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

            # Если это обработка выбранных элементов
            if self.selected_items:
                self.log.emit(f"\n🔄 РЕЖИМ: ОБРАБОТКА ВЫБРАННЫХ")
                self.log.emit(f"Выбрано {len(self.selected_items)} записей")

                # Подсчитываем количество с ручными DOI и без
                manual_count = sum(1 for item in self.selected_items if item.get('manual_doi'))
                auto_count = len(self.selected_items) - manual_count
                if manual_count > 0:
                    self.log.emit(f"   💾 С ручными DOI: {manual_count}")
                if auto_count > 0:
                    self.log.emit(f"   🤖 Автопоиск: {auto_count}")

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

                            # Логируем тип обработки
                            if manual_doi:
                                self.log.emit(f"\n💾 ЗАПИСЬ #{original_index} (ручной DOI: {manual_doi})")
                            else:
                                self.log.emit(f"\n🤖 ЗАПИСЬ #{original_index} (автопоиск)")

                            result = self.process_item(item_to_process, manual_doi)
                            results.append(result)

                            # Обновляем прогресс
                            self.processed_items += 1
                            progress = int((self.processed_items / self.total_items_to_process) * 100)
                            self.progress.emit(progress,
                                               f"Обработка: {self.processed_items}/{self.total_items_to_process}")

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

            if self.selected_items:
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
