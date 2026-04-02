"""
Разрешение DOI через различные источники
"""
import re
import time
import requests
import xml.etree.ElementTree as ET
from difflib import SequenceMatcher
from urllib.parse import urlparse, unquote
from .config import ENTREZ_EFETCH, CROSSREF_WORKS, GOOGLE_BOOKS_API, MAX_RETRIES, REQUEST_TIMEOUT
from .document_parser import DocumentParser
from .pubmed_searcher import search_doi_by_reference


class DOIResolver:
    """Класс для поиска DOI через различные источники"""

    def __init__(self, user_email, log_callback=None):
        """
        Args:
            user_email: Email пользователя для API запросов
            log_callback: Функция обратного вызова для логирования
        """
        self.user_email = user_email
        self.log_callback = log_callback
        self.headers = {
            'User-Agent': f'DOIExtractor/1.0 (mailto:{user_email})',
            'Accept': 'application/json',
        }

    def log(self, message):
        """Логирование сообщений"""
        if self.log_callback:
            self.log_callback(message)

    def extract_doi_from_text(self, text):
        """Извлекает DOI из текста используя regex паттерны

        ✅ ИЗМЕНЕНО: Добавлены строчные буквы a-z в паттерн (Nature DOI)
        ✅ ИЗМЕНЕНО: Удаление дублей (предпочтение более длинным DOI)
        """
        dois = set()
        # ✅ ИЗМЕНЕНО: Добавлены a-z в паттерн (Nature DOI содержат строчные буквы)
        doi_pattern = r'10\.\d{4,9}/[-._;()/:A-Za-z0-9]+'

        for match in re.finditer(r'doi:\s*(' + doi_pattern + r')', text, re.IGNORECASE):
            dois.add(match.group(1).strip())

        for match in re.finditer(r'https?://doi\.org/(' + doi_pattern + r')', text, re.IGNORECASE):
            dois.add(match.group(1).strip())

        for match in re.finditer(doi_pattern, text):
            candidate = match.group(0).strip()
            if candidate.startswith('10.'):
                dois.add(candidate)

        # ✅ Удаление дублей (предпочтение более длинным DOI)
        dois = self._deduplicate_dois(dois)

        return sorted(dois)

    def _deduplicate_dois(self, dois):
        """Удаляет дубликаты DOI, предпочитая более полные версии

        Например:
        - 10.1038/171737 и 10.1038/171737a0 → оставляем 10.1038/171737a0
        """
        if not dois:
            return set()

        sorted_dois = sorted(dois, key=len, reverse=True)
        unique_dois = set()

        for doi in sorted_dois:
            is_subdoi = False
            for existing in list(unique_dois):
                if doi.startswith(existing) and len(doi) > len(existing):
                    unique_dois.discard(existing)
                elif existing.startswith(doi) and len(existing) > len(doi):
                    is_subdoi = True
                    break

            if not is_subdoi:
                unique_dois.add(doi)

        return unique_dois

    def extract_doi_from_doi_url(self, url):
        """Извлекает DOI из URL вида doi.org/10.xxxx/xxxx или doi:10.xxxx/xxxx

        ✅ ИЗМЕНЕНО: Добавлена поддержка doi: протокола
        """
        if not url:
            return None
            
        # ✅ ИЗМЕНЕНО: Паттерн включает a-z для полных DOI
        doi_pattern = r'10\.\d{4,9}/[-._;()/:A-Za-z0-9]+'
        
        # Поддержка doi:10.xxxx/xxxx
        if url.lower().startswith('doi:'):
            doi_number = url[4:].strip()
            match = re.match(doi_pattern, doi_number)
            if match:
                return match.group(0)
            return None
        
        # Поддержка https://doi.org/10.xxxx/xxxx
        if 'doi.org' in url.lower():
            match = re.search(doi_pattern, url)
            if match:
                return match.group(0)
        return None

    def extract_pmid_from_url(self, url):
        """Извлекает PMID из PubMed URL"""
        parsed = urlparse(url)
        if 'pubmed.ncbi.nlm.nih.gov' in parsed.netloc:
            match = re.search(r'/(\d+)/?$', parsed.path)
            if match:
                return match.group(1)
        return None

    def extract_pmid_from_text(self, text):
        """Извлекает PMID из текста"""
        pmid_patterns = [
            r'PMID[:\s]*(\d{6,9})',
            r'pmid=(\d{6,9})',
            r'PubMed[:\s]*(\d{6,9})'
        ]

        for pattern in pmid_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                pmid = match.group(1).strip()
                self.log(f"  📊 Найден PMID в тексте: {pmid}")
                return pmid

        return None

    def extract_pmcid_from_url(self, url):
        """Извлекает PMCID из PMC URL"""
        parsed = urlparse(url)
        if 'pmc.ncbi.nlm.nih.gov' in parsed.netloc:
            match = re.search(r'/articles/(PMC\d+)/?$', parsed.path, re.IGNORECASE)
            if match:
                return match.group(1)
        return None

    def get_doi_from_pubmed_api(self, pmid):
        """Получает DOI через PubMed API по PMID"""
        try:
            params = {
                'db': 'pubmed',
                'id': pmid,
                'retmode': 'xml',
                'email': self.user_email,
                'tool': 'DOIExtractor'
            }

            for attempt in range(MAX_RETRIES):
                try:
                    response = requests.get(
                        ENTREZ_EFETCH,
                        params=params,
                        headers=self.headers,
                        timeout=REQUEST_TIMEOUT
                    )
                    if response.status_code == 200:
                        root = ET.fromstring(response.content)

                        for elocation in root.findall(".//ELocationID[@EIdType='doi']"):
                            doi = elocation.text.strip()
                            if doi.startswith('10.'):
                                return doi

                        for aid in root.findall(".//ArticleId[@IdType='doi']"):
                            doi = aid.text.strip()
                            if doi.startswith('10.'):
                                return doi

                        return None

                    elif response.status_code == 429:
                        wait = 2 ** attempt
                        time.sleep(wait)
                        continue
                    else:
                        return None

                except (requests.exceptions.RequestException, ET.ParseError):
                    if attempt < MAX_RETRIES - 1:
                        time.sleep(2)
                        continue
                    return None

            return None
        except Exception:
            return None

    def get_doi_from_pmc_api(self, pmcid):
        """Получает DOI через PMC API по PMCID"""
        try:
            params = {
                'db': 'pmc',
                'id': pmcid,
                'retmode': 'xml',
                'email': self.user_email,
                'tool': 'DOIExtractor'
            }

            for attempt in range(MAX_RETRIES):
                try:
                    response = requests.get(
                        ENTREZ_EFETCH,
                        params=params,
                        headers=self.headers,
                        timeout=REQUEST_TIMEOUT
                    )
                    if response.status_code == 200:
                        root = ET.fromstring(response.content)

                        for article_id in root.findall(".//article-id[@pub-id-type='doi']"):
                            doi = article_id.text.strip()
                            if doi.startswith('10.'):
                                return doi

                        return None

                    elif response.status_code == 429:
                        wait = 2 ** attempt
                        time.sleep(wait)
                        continue
                    else:
                        return None

                except (requests.exceptions.RequestException, ET.ParseError):
                    if attempt < MAX_RETRIES - 1:
                        time.sleep(2)
                        continue
                    return None

            return None
        except Exception:
            return None

    def get_doi_from_isbn_crossref(self, isbn):
        """Получает DOI через CrossRef API по ISBN

        Args:
            isbn: ISBN книги

        Returns:
            str или None: Найденный DOI или None
        """
        try:
            isbn_clean = re.sub(r'[-\s]', '', isbn)
            params = {
                'query.isbn': isbn_clean,
                'rows': 1,
                'select': 'DOI,title'
            }

            response = requests.get(CROSSREF_WORKS, params=params, headers=self.headers, timeout=REQUEST_TIMEOUT)

            if response.status_code == 200:
                data = response.json()
                if data.get('status') == 'ok' and 'message' in data:
                    items = data['message'].get('items', [])
                    if items:
                        doi = items[0].get('DOI', '').strip()
                        if doi.startswith('10.'):
                            self.log(f"  ✅ DOI найден по ISBN в CrossRef: {doi}")
                            return doi
            return None
        except Exception as e:
            self.log(f"  ❌ Ошибка поиска DOI по ISBN в CrossRef: {str(e)}")
            return None

    def get_book_metadata_by_isbn(self, isbn):
        """Получает метаданные книги по ISBN через Google Books API

        Args:
            isbn: ISBN книги (10 или 13 цифр)

        Returns:
            dict: Метаданные книги или None
        """
        try:
            isbn_clean = re.sub(r'[-\s]', '', isbn)
            url = f"{GOOGLE_BOOKS_API}?q=isbn:{isbn_clean}"

            response = requests.get(url, headers=self.headers, timeout=REQUEST_TIMEOUT)

            if response.status_code == 200:
                data = response.json()
                if data.get('totalItems', 0) > 0:
                    volume = data['items'][0].get('volumeInfo', {})

                    metadata = {
                        'title': volume.get('title', ''),
                        'authors': volume.get('authors', []),
                        'publisher': volume.get('publisher', ''),
                        'published_date': volume.get('publishedDate', ''),
                        'isbn_10': None,
                        'isbn_13': None,
                        'edition': '',
                        'pages': volume.get('pageCount', ''),
                    }

                    identifiers = volume.get('industryIdentifiers', [])
                    for ident in identifiers:
                        if ident.get('type') == 'ISBN_10':
                            metadata['isbn_10'] = ident.get('identifier')
                        elif ident.get('type') == 'ISBN_13':
                            metadata['isbn_13'] = ident.get('identifier')

                    self.log(f"  📚 Метаданные книги получены: {metadata['title'][:50]}...")
                    return metadata

            self.log(f"  ⚠️ Метаданные по ISBN не найдены")
            return None

        except Exception as e:
            self.log(f"  ❌ Ошибка получения метаданных по ISBN: {str(e)}")
            return None

    def format_book_apa_from_metadata(self, metadata):
        """Форматирует книгу в стиле APA по метаданным

        Args:
            metadata: dict с полями title, authors, publisher, published_date

        Returns:
            str: APA-цитата
        """
        authors = metadata.get('authors', [])
        title = metadata.get('title', 'Без названия')
        publisher = metadata.get('publisher', '')
        pub_date = metadata.get('published_date', '')

        if not authors:
            authors_str = "Без автора"
        elif len(authors) == 1:
            parts = authors[0].split(', ')
            authors_str = f"{parts[-1]}, {parts[0][0]}." if len(parts) > 1 else authors[0]
        elif len(authors) == 2:
            a1 = authors[0].split(', ')
            a2 = authors[1].split(', ')
            auth1 = f"{a1[-1]}, {a1[0][0]}." if len(a1) > 1 else authors[0]
            auth2 = f"{a2[-1]}, {a2[0][0]}." if len(a2) > 1 else authors[1]
            authors_str = f"{auth1} & {auth2}"
        else:
            formatted = []
            for a in authors[:6]:
                parts = a.split(', ')
                if len(parts) > 1:
                    formatted.append(f"{parts[-1]}, {parts[0][0]}.")
                else:
                    formatted.append(a)
            authors_str = ", ".join(formatted[:-1]) + ", & " + formatted[-1]

        year = pub_date[:4] if pub_date and len(pub_date) >= 4 else "б.г."

        citation = f"{authors_str} ({year}). <i>{title}</i>. {publisher}."
        return citation.strip()

    def extract_researchgate_title(self, url):
        """Извлекает заголовок из структуры URL ResearchGate"""
        try:
            parsed = urlparse(url)
            if 'researchgate.net' not in parsed.netloc:
                return None

            path = unquote(parsed.path)
            match = re.search(r'/publication/\d+_(.+?)(?:/|$)', path)
            if not match:
                return None

            title_raw = match.group(1)
            title_clean = title_raw.replace('_', ' ')
            title_clean = re.sub(r'[^a-zA-Z0-9\s\-.,:;]', ' ', title_clean)
            title_clean = re.sub(r'\s+', ' ', title_clean).strip()

            if len(title_clean) < 20 or len(title_clean) > 150:
                return None

            self.log(f"  🔍 Извлечен заголовок из ResearchGate: {title_clean[:50]}...")
            return title_clean

        except Exception as e:
            self.log(f"  ❌ Ошибка извлечения заголовка из ResearchGate: {str(e)}")
            return None

    def search_doi_via_crossref(self, title, url=None):
        """Поиск DOI через CrossRef API по заголовку статьи"""
        try:
            self.log(f"  🔍 Поиск DOI по заголовку: {title[:50]}...")
            clean_title = re.sub(r'[^\w\s\-.,:;]', '', title)

            if len(clean_title) < 15:
                self.log(f"  ⚠️ Заголовок слишком короткий для поиска")
                return None

            if len(clean_title) > 100:
                clean_title = clean_title[:100].rsplit(' ', 1)[0]

            params = {
                'query.bibliographic': clean_title,
                'rows': 3,
                'select': 'DOI,title'
            }

            for attempt in range(MAX_RETRIES - 1):
                try:
                    response = requests.get(
                        CROSSREF_WORKS,
                        params=params,
                        headers=self.headers,
                        timeout=REQUEST_TIMEOUT
                    )

                    if response.status_code == 200:
                        data = response.json()
                        if data.get('status') == 'ok' and 'message' in data and data['message'].get('items'):
                            items = data['message']['items']
                            self.log(f"  📊 Найдено {len(items)} возможных совпадений")

                            for item in items:
                                doi = item.get('DOI', '').strip()
                                if not doi.startswith('10.'):
                                    continue

                                found_titles = item.get('title', [])
                                if not found_titles:
                                    continue

                                found_title = found_titles[0].lower() if found_titles else ''
                                clean_found = DocumentParser.normalize_title(found_title)
                                clean_search = DocumentParser.normalize_title(clean_title)

                                if clean_search[:25] in clean_found or clean_found[:25] in clean_search:
                                    self.log(f"  ✅ Найден DOI: {doi}")
                                    return doi

                                if len(clean_search) < 30:
                                    self.log(f"  ✅ Найден DOI: {doi}")
                                    return doi

                                similarity = SequenceMatcher(None, clean_search, clean_found).ratio()
                                if similarity >= 0.75:
                                    self.log(f"  ✅ Найден DOI (сходство: {similarity:.2f}): {doi}")
                                    return doi

                        self.log(f"  ❌ DOI по заголовку не найден")
                        return None

                    elif response.status_code == 429:
                        wait = 2 ** attempt
                        self.log(f"  ⚠️ Rate limit, ожидание {wait} сек...")
                        time.sleep(wait)
                        continue

                except requests.exceptions.RequestException as e:
                    self.log(f"  ❌ Ошибка поиска DOI: {str(e)}")
                    if attempt < MAX_RETRIES - 2:
                        time.sleep(2)
                        continue

            return None
        except Exception as e:
            self.log(f"  ❌ Исключение при поиске DOI: {str(e)}")
            return None

    def resolve_doi(self, item, manual_doi=None, auto_search=True):
        """Разрешает DOI для элемента используя различные стратегии

        Args:
            item: Словарь с данными статьи (title, url, text, isbn)
            manual_doi: Ручной DOI или ISBN (приоритетный)
            auto_search: Выполнять автоматический поиск

        Returns:
            list или dict: Список найденных DOI или dict с метаданными книги
        """
        self.log(f"\n📄 Обработка записи #{item['original_index']}: {item['title'][:50]}...")
        dois = set()

        # ✅ ПРОВЕРКА: Если ручной ISBN (не DOI) - используем его напрямую
        if manual_doi and not manual_doi.startswith('10.'):
            self.log(f"  📚 Обнаружен ручной ISBN: {manual_doi}")
            # Получаем метаданные книги по ISBN
            book_meta = self.get_book_metadata_by_isbn(manual_doi)
            if book_meta:
                apa_citation = self.format_book_apa_from_metadata(book_meta)
                self.log(f"  ✅ Сформирована APA-цитата из метаданных ISBN")
                return {
                    'dois': [],
                    'article_title': book_meta.get('title'),
                    'apa_citation': apa_citation,
                    'has_data': True,
                    'source': 'ISBN_METADATA',
                    'isbn': manual_doi
                }
            else:
                # Метаданные не найдены, но ISBN валиден
                return {
                    'dois': [],
                    'article_title': item.get('title'),
                    'apa_citation': None,
                    'has_data': True,
                    'source': 'ISBN_MANUAL',
                    'isbn': manual_doi
                }

        # Ручной DOI имеет наивысший приоритет
        if manual_doi and manual_doi.strip():
            self.log(f"  ✅ Используется ручной DOI: {manual_doi}")
            dois.add(manual_doi.strip())
            return sorted(dois)

        if not auto_search:
            return []

        # Автоматический поиск DOI
        self.log(f"  🔍 Поиск DOI в тексте...")
        text_dois = self.extract_doi_from_text(item['text'])
        if text_dois:
            dois.update(text_dois)
            self.log(f"  ✅ Найдено DOI в тексте: {', '.join(text_dois)}")

        if not dois:
            self.log(f"  🔍 Поиск DOI в URL...")
            doi_from_url = self.extract_doi_from_doi_url(item['url'])
            if doi_from_url:
                dois.add(doi_from_url)
                self.log(f"  ✅ DOI извлечен из URL: {doi_from_url}")
            
            # ✅ Проверка doi: в тексте URL
            if not doi_from_url and item['url'].lower().startswith('doi:'):
                self.log(f"  🔍 Обнаружен doi: протокол в URL")

        if not dois:
            self.log(f"  🔍 Проверка PubMed...")
            pmid = self.extract_pmid_from_url(item['url'])
            if pmid:
                self.log(f"  📊 Найден PMID: {pmid}")
                doi_from_pubmed = self.get_doi_from_pubmed_api(pmid)
                if doi_from_pubmed:
                    dois.add(doi_from_pubmed)
                    self.log(f"  ✅ PubMed DOI: {doi_from_pubmed}")

        # ✅ Проверка PMID в тексте
        if not dois:
            pmid_from_text = self.extract_pmid_from_text(item['text'])
            if pmid_from_text:
                doi_from_pubmed = self.get_doi_from_pubmed_api(pmid_from_text)
                if doi_from_pubmed:
                    dois.add(doi_from_pubmed)

        if not dois:
            self.log(f"  🔍 Проверка PMC...")
            pmcid = self.extract_pmcid_from_url(item['url'])
            if pmcid:
                self.log(f"  📊 Найден PMCID: {pmcid}")
                doi_from_pmc = self.get_doi_from_pmc_api(pmcid)
                if doi_from_pmc:
                    dois.add(doi_from_pmc)
                    self.log(f"  ✅ PMC DOI: {doi_from_pmc}")

        if not dois and 'researchgate.net' in item['url'].lower():
            self.log(f"  🔍 Обработка ResearchGate...")
            rg_title = self.extract_researchgate_title(item['url'])
            if rg_title:
                rg_doi = self.search_doi_via_crossref(rg_title, item['url'])
                if rg_doi:
                    dois.add(rg_doi)
                    self.log(f"  ✅ ResearchGate DOI: {rg_doi}")

        if not dois and item['title']:
            self.log(f"  🔍 Резервный поиск по заголовку...")
            doi = self.search_doi_via_crossref(item['title'], item['url'])
            if doi:
                dois.add(doi)
                self.log(f"  ✅ DOI найден по заголовку: {doi}")

        # 🔍 НОВЫЙ БЛОК: Поиск DOI через PubMed по библиографической ссылке (RefChecker)
        if not dois and item.get('text'):
            self.log(f"  🔍 Поиск DOI через PubMed (RefChecker)...")
            doi, status = search_doi_by_reference(item['text'], self.user_email)
            if doi:
                dois.add(doi)
                self.log(f"  ✅ DOI найден через PubMed: {doi} ({status})")
            else:
                self.log(f"  ⚠️ DOI не найден через PubMed ({status})")

        # ✅ НОВЫЙ БЛОК: Обработка книг по ISBN (если DOI не найден)
        if not dois and item.get('isbn'):
            self.log(f"  📚 Попытка получить метаданные книги по ISBN: {item['isbn']}...")

            doi_from_isbn = self.get_doi_from_isbn_crossref(item['isbn'])
            if doi_from_isbn:
                dois.add(doi_from_isbn)
                self.log(f"  ✅ DOI найден по ISBN через CrossRef: {doi_from_isbn}")
            else:
                book_meta = self.get_book_metadata_by_isbn(item['isbn'])
                if book_meta:
                    apa_citation = self.format_book_apa_from_metadata(book_meta)
                    self.log(f"  ✅ Сформирована APA-цитата из метаданных ISBN")

                    return {
                        'dois': [],
                        'article_title': book_meta.get('title'),
                        'apa_citation': apa_citation,
                        'has_data': True,
                        'source': 'ISBN_METADATA',
                        'isbn': item['isbn']
                    }

        if not dois:
            self.log(f"  ⚠️ DOI не найден")

        return sorted(dois)