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
from .ris_exporter import RISExporter


class WorkerThread(QThread):
    """Поток для фоновой обработки документов."""
    progress = pyqtSignal(int, str)
    log = pyqtSignal(str)
    finished_success = pyqtSignal(str, list)
    finished_error = pyqtSignal(str)
    table_data = pyqtSignal(list)
    item_updated = pyqtSignal(dict)

    def __init__(self, input_path, user_email, selected_items=None, previous_items=None, analyze_citations=True):
        super().__init__()
        self.input_path = input_path
        self.user_email = user_email
        self.selected_items = selected_items if selected_items else []
        self.previous_items = previous_items if previous_items else []
        self.is_user_selection = bool(selected_items)
        self.analyze_citations = analyze_citations

        self.doi_resolver = DOIResolver(user_email, self.log.emit)
        self.citation_formatter = CitationFormatter(user_email, self.log.emit)

        self.total_items_to_process = 0
        self.processed_items = 0

    def process_item(self, item, manual_doi=None):
        """Обрабатывает один элемент и возвращает результат

        Args:
            item: Словарь с данными статьи
            manual_doi: Ручной DOI или ISBN (если указан)

        Returns:
            dict: Результат обработки
        """
        # ✅ ПРОВЕРКА: Если ручной ISBN (не DOI) - используем его напрямую
        if manual_doi and not manual_doi.startswith('10.'):
            self.log.emit(f"  📚 Обнаружен ручной ISBN: {manual_doi}")
            # Получаем метаданные книги по ISBN
            book_meta = self.doi_resolver.get_book_metadata_by_isbn(manual_doi)
            if book_meta:
                apa_citation = self.doi_resolver.format_book_apa_from_metadata(book_meta)
                self.log.emit(f"  ✅ Сформирована APA-цитата из метаданных ISBN")
                return {
                    'original_index': item['original_index'],
                    'title': item['title'],
                    'url': item['url'],
                    'dois': [],
                    'article_title': book_meta.get('title'),
                    'apa_citation': apa_citation,
                    'has_data': True,
                    'manual_doi': manual_doi,
                    'is_cited': item.get('is_cited', True),
                    'isbn': manual_doi,
                    'source': 'ISBN_METADATA'
                }
            else:
                # Метаданные не найдены, но ISBN валиден
                return {
                    'original_index': item['original_index'],
                    'title': item['title'],
                    'url': item['url'],
                    'dois': [],
                    'article_title': item.get('title'),
                    'apa_citation': None,
                    'has_data': True,
                    'manual_doi': manual_doi,
                    'is_cited': item.get('is_cited', True),
                    'isbn': manual_doi,
                    'source': 'ISBN_MANUAL'
                }

        # ✅ ПРОВЕРКА: Если ISBN найден в документе (автоматически) - пытаемся получить метаданные
        # Работает даже если есть URL (например, ссылка на магазин книг)
        if not manual_doi and item.get('isbn'):
            self.log.emit(f"  📚 Обнаружен ISBN в документе: {item['isbn']}")
            book_meta = self.doi_resolver.get_book_metadata_by_isbn(item['isbn'])
            if book_meta:
                apa_citation = self.doi_resolver.format_book_apa_from_metadata(book_meta)
                self.log.emit(f"  ✅ Сформирована APA-цитата из метаданных ISBN")
                return {
                    'original_index': item['original_index'],
                    'title': item['title'],
                    'url': item['url'],
                    'dois': [],
                    'article_title': book_meta.get('title'),
                    'apa_citation': apa_citation,
                    'has_data': True,
                    'manual_doi': None,
                    'is_cited': item.get('is_cited', True),
                    'isbn': item['isbn'],
                    'source': 'ISBN_METADATA'
                }

        # Пропуск для нецитируемых записей (только если нет ручного DOI/ISBN)
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
                'uncited_reason': 'NOT_IN_TEXT',
                'isbn': item.get('isbn')
            }

        if self.is_user_selection and not item.get('is_cited', True):
            self.log.emit(f"  🎯 Явный выбор пользователя: обработка несмотря на статус 'не цитируется'")

        # Разрешаем DOI
        auto_search = not manual_doi
        result_data = self.doi_resolver.resolve_doi(item, manual_doi, auto_search)

        # ✅ Обработка результата: может быть list (DOI) или dict (метаданные книги)
        if isinstance(result_data, dict) and result_data.get('source') == 'ISBN_METADATA':
            dois_list = []
            article_title = result_data.get('article_title')
            apa_citation = result_data.get('apa_citation')
            has_data = True
        else:
            dois_list = result_data if isinstance(result_data, list) else []
            article_title = None
            apa_citation = None

            if dois_list:
                primary_doi = dois_list[0]
                self.log.emit(f"  📊 Получение метаданных для DOI: {primary_doi}")
                article_title = self.citation_formatter.get_article_title_from_doi(primary_doi)
                apa_citation = self.citation_formatter.get_apa_citation(primary_doi)
                time.sleep(RATE_LIMIT_DELAY)
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
            'is_cited': item.get('is_cited', True),
            'isbn': item.get('isbn')
        }

        status = "✅" if has_data else "⚠️"
        source = "РУЧНОЙ DOI/ISBN" if manual_doi else ("АВТОМАТИЧЕСКИ" if has_data else "НЕТ ДАННЫХ")
        self.log.emit(f"  {status} Итог: {source}")

        return result

    def run(self):
        """Главная функция обработки"""
        try:
            self.log.emit(f"⚙️ Загрузка документа...")
            doc = Document(str(self.input_path))

            all_items = DocumentParser.collect_items_from_document(doc, analyze_citations=self.analyze_citations)

            if not all_items:
                self.finished_error.emit("⚠️ Не найдено пунктов для обработки")
                return

            if self.selected_items:
                self.log.emit(f"\n🔄 РЕЖИМ: ОБРАБОТКА ВЫБРАННЫХ")
                self.log.emit(f"Выбрано {len(self.selected_items)} записей")

                manual_count = sum(1 for item in self.selected_items if item.get('manual_doi'))
                auto_count = len(self.selected_items) - manual_count
                if manual_count > 0:
                    self.log.emit(f"   💾 С ручными DOI/ISBN: {manual_count}")
                if auto_count > 0:
                    self.log.emit(f"   🤖 Автопоиск: {auto_count}")

                results = []
                self.total_items_to_process = len(self.selected_items)
                self.processed_items = 0

                items_map = {item['original_index']: item for item in all_items}

                for i, selected_item in enumerate(self.selected_items):
                    if 'original_item' in selected_item:
                        original_index = selected_item['original_item'].get('original_index')
                        if original_index in items_map:
                            item_to_process = items_map[original_index]
                            manual_doi = selected_item.get('manual_doi')

                            if manual_doi:
                                self.log.emit(f"\n💾 ЗАПИСЬ #{original_index} (ручной DOI/ISBN: {manual_doi})")
                            else:
                                self.log.emit(f"\n🤖 ЗАПИСЬ #{original_index} (автопоиск)")

                            result = self.process_item(item_to_process, manual_doi)
                            results.append(result)

                            self.processed_items += 1
                            progress = int((self.processed_items / self.total_items_to_process) * 100)
                            self.progress.emit(progress,
                                               f"Обработка: {self.processed_items}/{self.total_items_to_process}")

                processed_indices = {item['original_index'] for item in results}
                for prev_item in self.previous_items:
                    if prev_item.get('original_index') not in processed_indices:
                        results.append(prev_item)
                        self.log.emit(f"📋 Сохранена запись #{prev_item.get('original_index')} (без изменений)")

                results.sort(key=lambda x: x.get('original_index', 0))

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

                    self.processed_items += 1
                    progress = int((self.processed_items / self.total_items_to_process) * 100)
                    self.progress.emit(progress, f"Обработка: {self.processed_items}/{self.total_items_to_process}")

            table_data = HTMLGenerator.generate_table_data(results)
            self.table_data.emit(table_data)

            self.progress.emit(90, "Генерация HTML...")
            self.log.emit("\n🎨 ГЕНЕРАЦИЯ HTML-РЕЗУЛЬТАТА...")

            if self.selected_items:
                output_path = self.input_path.parent / f"{self.input_path.stem}_apa_updated.html"
            else:
                output_path = self.input_path.parent / f"{self.input_path.stem}_apa.html"

            html_content = HTMLGenerator.generate_html_ordered(results, self.input_path.stem, self.user_email)

            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(html_content)

            # ✅ ЭКСПОРТ В RIS
            self.log.emit("\n📑 ГЕНЕРАЦИЯ RIS-ФАЙЛА...")
            ris_path = self.input_path.parent / f"{self.input_path.stem}_bibliography.ris"

            if RISExporter.save_ris_file(results, ris_path, self.input_path.stem):
                self.log.emit(f"✅ RIS-файл сохранен: {ris_path.name}")
            else:
                self.log.emit(f"⚠️ Ошибка экспорта RIS")

            self.progress.emit(100, "Готово!")
            self.log.emit(f"\n✅ ОБРАБОТКА УСПЕШНО ЗАВЕРШЕНА!")
            self.log.emit(f"📁 Результат сохранен: {output_path.name}")
            self.finished_success.emit(str(output_path), results)

        except Exception as e:
            import traceback
            error_msg = f"❌ КРИТИЧЕСКАЯ ОШИБКА:\n{str(e)}\n\n{traceback.format_exc()}"
            self.log.emit(f"\n{error_msg}")
            self.finished_error.emit(error_msg)