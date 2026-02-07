"""
Конфигурация приложения
"""

# API endpoints
ENTREZ_EFETCH = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"
CROSSREF_WORKS = "https://api.crossref.org/works"

# Настройки приложения
APP_NAME = "DOIExtractor"
APP_VERSION = "1.0"
SETTINGS_ORG = "DOIExtractor"
SETTINGS_APP = "APAFormatter"

# Параметры обработки
MAX_RETRIES = 3
REQUEST_TIMEOUT = 10
RATE_LIMIT_DELAY = 0.5  # секунды между запросами DOI метаданных
