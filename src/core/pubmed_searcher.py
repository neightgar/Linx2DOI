"""
PubMed Searcher - Поиск статей в PubMed по заголовку и авторам
Интегрирует функционал из проекта RefChecker для поиска DOI через PubMed
"""
import re
import time
from typing import Optional, List, Tuple
from Bio import Entrez
from rapidfuzz import fuzz

# ================= НАСТРОЙКИ =================
REQUEST_DELAY = 0.25  # Задержка между запросами к PubMed
FUZZY_THRESHOLD = 85  # Минимальный порог нечёткого совпадения
EXACT_MATCH_THRESHOLD = 95  # Для раннего выхода при точном совпадении
GOOD_MATCH_THRESHOLD = 90  # Для раннего выхода при хорошем совпадении
SEARCH_TIMEOUT = 20  # Максимальное время поиска одной ссылки (сек)

# Настройки поиска (прогрессивная стратегия)
RETMAX_HIGH_CONFIDENCE = 10    # Для первых 2 запросов
RETMAX_MEDIUM_CONFIDENCE = 20  # Для следующих запросов
RETMAX_DESPERATE = 50          # Для последних отчаянных попыток

# Валидация парсинга
MIN_TITLE_LENGTH = 10
MIN_WORDS_IN_TITLE = 2
MAX_TITLE_LENGTH = 300
MAX_KEYWORDS = 7

# Предкомпилированные regex паттерны для парсеров
# ГОСТ
GOST_JOURNAL_SPLIT = re.compile(r'\s*//\s*')
GOST_SLASH_SPLIT = re.compile(r'\s+/\s+')
GOST_ET_AL = re.compile(r',?\s*et\s+al\.?', re.IGNORECASE)
GOST_AUTHOR_START = re.compile(r'^[A-Z][a-z]+\s+[A-Z]\.')
GOST_BOUNDARY_PATTERN = re.compile(r'([A-Z]+)\.\s+([A-Z][a-z]+)')

# APA
APA_YEAR_PATTERN = re.compile(r'\(\d{4}\)\.? ')
APA_YEAR_ALT = re.compile(r'\(\d{4}\)[ ,]')
APA_TITLE_END = re.compile(r'\.\s+[A-Z][a-z]+|,\s*\d+\(')
APA_AUTHOR_SPLIT = re.compile(r',\s*(?=[A-Z][a-z]+,?\s+[A-Z]\.)')

# Vancouver
VANCOUVER_NUMBER = re.compile(r'^\d+[\.\)]\s*')
VANCOUVER_ET_AL = re.compile(r',?\s*et\s+al\.?\s*', re.IGNORECASE)
VANCOUVER_AUTHOR_END = re.compile(
    r'(?:[A-Z]{1,3}|ETAL_MARKER)\.\s+'
    r'(\d|[A-Z]{2,}[:\s\-]|[A-Z][a-z]|[A-Z][\-\d]|A\s+[a-z])'
)
VANCOUVER_AUTHOR_END_ALT = re.compile(
    r'[A-Z]{1,3}\.(\d|[A-Z]{2,}[:\s\-]|[A-Z][a-z]|[A-Z][\-\d]|A\s+[a-z])'
)
VANCOUVER_AUTHOR_END_FALLBACK = re.compile(r'\.\s+([^\s,])')
VANCOUVER_JOURNAL_KEYWORDS = re.compile(
    r'\.\s+(?:[A-Z][a-z]*\s+){0,3}(?:'
    r'J\b|Journal|Int\b|Ann\b|Proc\b|Trans\b|Rev\b|Lett\b|Rep\b|Clin\b|Med\b|Sci\b|'
    r'Biol\b|Chem\b|Phys\b|Eng\b|Tech\b|Res\b|Acta\b|Arch\b|Bull\b|Cancer\b|'
    r'Cell\b|Curr\b|Eur\b|Exp\b|FASEB\b|Front\b|Genet\b|Genom\b|Immunol\b|'
    r'Mol\b|Nat\b|Nucleic\b|Oncol\b|PLOS\b|PLoS\b|Pharm\b|Physiol\b|Radiat\b|'
    r'Biophys\b|Biochem\b|Bioinform\b|Methods\b|Mutat\b|Nucl\b|BMC\b|DNA\b|RNA\b|'
    r'Acids\b|Proteom\b|Genom\b|Toxicol\b|Carcinog\b|Am\b|Br\b|Genes\b|Chromosom|'
    r'Nature\b|Science\b|Lancet\b|Blood\b|Circulation\b|Diabetes\b|Gut\b|Heart\b|'
    r'Brain\b|Thorax\b|Pain\b|Spine\b|Stroke\b|Neurology\b|Oncogene\b|Carcinogenesis\b|'
    r'Leukemia\b|Neoplasia\b|Mutagenesis\b|eLife\b|Life\b|'
    r'Biochemistry\b|Bioessays\b|Biochim\b|Biophysics\b|Biomarkers\b|Biomaterials\b|'
    r'Biomolecules\b|Cancers\b|Cells\b|Chromosoma\b|Development\b|Endocrinology\b|'
    r'Epigenetics\b|Genes\b|Genetics\b|Genome\b|Genomics\b|Hepatology\b|'
    r'Immunity\b|Infection\b|Metabolism\b|Microbiology\b|Mutation\b|Neuron\b|'
    r'Nutrients\b|Oncology\b|Pathology\b|Pediatrics\b|Pharmacology\b|Proteomics\b|'
    r'Radiobiology\b|Toxicology\b|Virology\b|'
    r'Plant\b|Animal\b|Ecology\b|Evolution\b|Marine\b|Freshwater\b|'
    r'Environ\b|Environm\b|Environmental\b|Sustainability\b|Climate\b|'
    r'Comput\b|Computer\b|Inform\b|Data\b|Knowl\b|Syst\b|'
    r'Mater\b|Material\b|Struct\b|Mech\b|Design\b|'
    r'Educ\b|Teach\b|Learn\b|Student\b|'
    r'Public\b|Health\b|Policy\b|Manage\b|Econ\b|'
    r'Surg\b|Orthop\b|Dent\b|Oral\b|Pathol\b|'
    r'Obstet\b|Gynecol\b|Reprod\b|Fertil\b|'
    r'Dermatol\b|Ophthalmol\b|Cardiol\b|Gastroent\b|Nephrol\b|'
    r'Pulmonol\b|Rheumatol\b|Allergy\b|Infect\b|'
    r'Microb\b|Parasitol\b|Entomol\b|Zool\b|Bot\b|'
    r'Geogr\b|Geol\b|Palaeontol\b|Archaeol\b|'
    r'Anthropol\b|Sociol\b|Psychol\b|Psychiat\b|'
    r'Linguist\b|Philos\b|Theol\b|Histor\b|'
    r'Art\b|Music\b|Liter\b|Cult\b|Sport\b|'
    r'Tourism\b|Hospitality\b|Food\b|Agric\b|'
    r'Forest\b|Fish\b|Wildlife\b|Conserv\b|'
    r'Energy\b|Power\b|Fuel\b|Petrol\b|'
    r'Min\b|Metall\b|Ceramic\b|Polym\b|'
    r'Textile\b|Paper\b|Packaging\b|Print\b|'
    r'Optic\b|Laser\b|Spectrosc\b|'
    r'Acoust\b|Vibr\b|Tribol\b|'
    r'Corros\b|Coat\b|Adhes\b|'
    r'Anal\b|Chromatogr\b|Electrochem\b|'
    r'Nanotechnol\b|Biotechnol\b|'
    r'Robot\b|Autom\b|Control\b|'
    r'Network\b|Commun\b|Telecommun\b|'
    r'Signal\b|Image\b|Pattern\b|'
    r'Physica\b|Theor\b|Appl\b|'
    r'Open\b|Access\b|Direct\b|'
    r'Global\b|Regional\b|Local\b|'
    r'Urban\b|Rural\b|Regional\b|'
    r'Comp\b|Toxicol\b|Pharmacol\b|'
    r'Vet\b|Animal\b|Dairy\b|'
    r'Food\b|Sci\b|Technol\b|'
    r'Waste\b|Water\b|Air\b|'
    r'Soil\b|Land\b|Coastal\b|'
    r'Atmos\b|Ocean\b|Polar\b|'
    r'Space\b|Aero\b|Satellit\b|'
    r'Radiat\b|Isot\b|Nucl\b|'
    r'Plasma\b|Fusion\b|'
    r'Crystal\b|Miner\b|'
    r'Palaeo\b|Quatern\b|'
    r'Prehist\b|Ancient\b|'
    r'Mediev\b|Modern\b|'
    r'Contemp\b|Class\b|'
    r'Early\b|Late\b|'
    r'North\b|South\b|East\b|West\b|'
    r'Cent\b|Assoc\b|Natl\b|Int\b|'
    r'World\b|Eur\b|Asian\b|Afr\b|'
    r'Am\b|Can\b|Aust\b|'
    r'Jpn\b|Chin\b|Ind\b|'
    r'Braz\b|Mex\b|Arg\b|'
    r'Rus\b|Sov\b|'
    r'UK\b|US\b|EU\b|'
    r'IEEE\b|ACM\b|'
    r' Springer\b|Elsevier\b|Wiley\b|'
    r'Taylor\b|Francis\b|'
    r'Palgrave\b|Macmillan\b|'
    r'Routledge\b|Sage\b|'
    r'Oxford\b|Cambridge\b|Harvard\b|'
    r'MIT\b|Stanford\b|'
    r'Calif\b|NY\b|'
    r'London\b|Paris\b|Berlin\b|'
    r'Tokyo\b|Beijing\b|'
    r'Moscow\b|Sydney\b|'
    r'Toronto\b|Boston\b|'
    r'Chicago\b|LA\b|'
    r'SF\b|DC\b|'
    r'AB\b|BC\b|'
    r'ON\b|QC\b|'
    r'NS\b|NB\b|'
    r'MB\b|SK\b|'
    r'AK\b|AL\b|AR\b|AZ\b|'
    r'CA\b|CO\b|CT\b|'
    r'DC\b|DE\b|FL\b|GA\b|'
    r'HI\b|IA\b|ID\b|IL\b|IN\b|'
    r'KS\b|KY\b|LA\b|MA\b|MD\b|'
    r'ME\b|MI\b|MN\b|MO\b|MS\b|'
    r'MT\b|NC\b|ND\b|NE\b|NH\b|'
    r'NJ\b|NM\b|NV\b|NY\b|OH\b|'
    r'OK\b|OR\b|PA\b|RI\b|SC\b|'
    r'SD\b|TN\b|TX\b|UT\b|VA\b|'
    r'VT\b|WA\b|WI\b|WV\b|WY\b)',
    re.IGNORECASE
)
VANCOUVER_YEAR_PATTERN = re.compile(r'\.\s*(?:19|20)\d{2}\s*[;:(]')
VANCOUVER_YEAR_ALT = re.compile(r'\.\s*(?:19|20)\d{2}\s')
VANCOUVER_AUTHOR_PATTERN = re.compile(r'([A-Z][a-z]+(?:\s+[a-z]+)?\s+[A-Z]{1,3})')


def extract_journal_and_year(ref: str, style: str) -> Tuple[Optional[str], Optional[str]]:
    """
    Извлекает название журнала и год из библиографической ссылки.
    Используется для улучшения поиска в PubMed.

    Args:
        ref: Текст ссылки
        style: Стиль оформления ("ГОСТ 7.0.100-2018", "APA", "Vancouver", "Hybrid")

    Returns:
        (journal, year) - кортеж с названием журнала и годом (или None если не найдено)
    """
    ref = ref.replace("\xa0", " ").strip()
    journal = None
    year = None

    if style == "ГОСТ 7.0.100-2018":
        year_match = re.search(r'[–-]\s*((19|20)\d{2})', ref)
        if year_match:
            year = year_match.group(1)
        journal_match = re.search(r'//\s*([A-ZА-Я][^.\n–-]{5,80}?)(?:\s*[.–]|\s*$)', ref)
        if journal_match:
            journal = journal_match.group(1).strip().rstrip('.')
        if not journal:
            journal_match = re.search(r'//\s*([A-ZА-Я][A-Za-zА-Яа-я.\s]+?)(?:\s*[.–]|\s*$)', ref)
            if journal_match:
                journal = journal_match.group(1).strip().rstrip('.')

    elif style == "APA":
        year_match = re.search(r'\((\d{4})\)', ref)
        if year_match:
            year = year_match.group(1)
        journal_match = re.search(r'\.\s*([A-Z][a-zA-Z]+(?:\s+[a-zA-Z]+)*?)\s*,\s*\d+', ref)
        if journal_match:
            journal = journal_match.group(1).strip()
        if not journal:
            journal_match = re.search(r'\)\.\s*([A-Z][a-zA-Z]+(?:\s+[a-zA-Z]+)*?)(?:,\s*\d+|\s*$)', ref)
            if journal_match:
                potential = journal_match.group(1).strip()
                if len(potential) < 60 and not potential.endswith(':'):
                    journal = potential

    elif style == "Vancouver":
        year_match = re.search(r'\.\s*((19|20)\d{2});', ref)
        if year_match:
            year = year_match.group(1)
        journal_match = re.search(r'\.\s+([A-Z](?:[a-zA-Z.\s]{2,40}?))\.\s*(?:19|20)\d{2}', ref)
        if journal_match:
            journal = journal_match.group(1).strip()
        if not journal:
            journal_match = re.search(r'\.\s+([A-Z]\s[A-Z][a-z]+(?:\s[A-Z][a-z]+)*)\.\s*(?:19|20)\d{2}', ref)
            if journal_match:
                journal = journal_match.group(1).strip()
        if not journal:
            journal_match = re.search(r'\.\s+([A-Z][a-zA-Z]{2,20}(?:\s[A-Z][a-zA-Z]{1,15}){0,2})\.\s*(?:19|20)\d{2}', ref)
            if journal_match:
                journal = journal_match.group(1).strip()

    elif style == "Hybrid":
        year_match = re.search(r'\.\s+((19|20)\d{2})\.', ref)
        if year_match:
            year = year_match.group(1)
        journal_match = re.search(r'([A-Z][a-z]+\.\s+(?:[A-Z][a-z]+\.\s*)+)\s*V\.?\s*\d+', ref)
        if journal_match:
            journal = journal_match.group(1).strip()
        if not journal:
            journal_match = re.search(r'([A-Z][a-z]+\.\s+(?:[A-Z][a-z]+\.\s*)+)\s*P\.?\s*\d+', ref)
            if journal_match:
                journal = journal_match.group(1).strip()
        if not journal:
            journal_match = re.search(r'\.\s+(Nature|Science|Cell|Genetics|Blood|Brain|Thorax|Cancer|Immunity|Development)\.\s*V\.?\s*\d+', ref)
            if journal_match:
                journal = journal_match.group(1)
        if not journal:
            journal_match = re.search(r'\.\s+([A-Z]\.\s+[A-Z][a-z]+)\.\s*V\.?\s*\d+', ref)
            if journal_match:
                journal = journal_match.group(1).strip()
        if not journal:
            journal_match = re.search(r'([A-Z][a-z]+\s+[a-z]+\.)\s+Bd\.?\s*\d+', ref)
            if journal_match:
                journal = journal_match.group(1).strip()
            if not journal:
                journal_match = re.search(r'([A-Z][a-z]+\s+[a-z]+\.)\s+S\.?\s*\d+', ref)
                if journal_match:
                    journal = journal_match.group(1).strip()

    if journal:
        journal = re.sub(r'\s+', ' ', journal)
        if len(journal) > 80:
            journal = journal[:80].rsplit(' ', 1)[0]

    return journal, year


def detect_reference_style(ref: str) -> Optional[str]:
    """
    Автоматически определяет стиль библиографической ссылки.
    Возвращает: "ГОСТ 7.0.100-2018", "APA", "Vancouver", "Hybrid" или None
    """
    ref = ref.replace("\xa0", " ").strip()

    gost_score = 0
    apa_score = 0
    vancouver_score = 0
    hybrid_score = 0

    if '//' in ref:
        gost_score += 3
    if re.search(r'\s+/\s+', ref):
        gost_score += 2
    if re.search(r'[А-Яа-я]', ref):
        gost_score += 2
    if re.search(r'[A-Z][a-z]+\s+[A-Z]\.\s*[A-Z]\.', ref):
        gost_score += 1
    if re.search(r'–\s*(19|20)\d{2}', ref):
        gost_score += 1

    year_in_parens = re.search(r'\(\d{4}\)', ref)
    if year_in_parens:
        apa_score += 3
        if re.search(r'\(\d{4}\)\.\s+[A-Z]', ref):
            apa_score += 2
    if re.search(r'[A-Z][a-z]+,\s+[A-Z]\.\s+[A-Z]\.', ref):
        apa_score += 2
    if re.search(r'&', ref):
        apa_score += 1
    if re.search(r'https?://doi\.org/', ref):
        apa_score += 1

    if re.search(r'^[0-9]+[\.\)]\s*', ref):
        vancouver_score += 2
    if re.search(r'[A-Z][a-z]+\s+[A-Z]{1,3}[\.,\s]', ref):
        if not re.search(r'[A-Z][a-z]+,\s+[A-Z]\.', ref):
            vancouver_score += 2
    if re.search(r'\.\s*(19|20)\d{2};', ref):
        vancouver_score += 3
    if re.search(r',\s*et\s+al\.?\s*\.', ref):
        vancouver_score += 1
    if not year_in_parens and re.search(r'(19|20)\d{2}', ref):
        vancouver_score += 1

    if re.search(r'[A-Z][a-z]+\s+[A-Z]\.,', ref):
        hybrid_score += 2
    if re.search(r'\.\s+((19|20)\d{2})\.', ref):
        hybrid_score += 3
    if re.search(r'\bV(?:ol)?\.?\s*\d+', ref):
        hybrid_score += 3
    if re.search(r'\b(?:P|S)\.?\s*\d+-\d+', ref):
        hybrid_score += 3
    if re.search(r'\bBd\.?\s*\d+', ref):
        hybrid_score += 2
    if re.search(r'[A-Z][a-z]+\.\s+[A-Z][a-z]+\.', ref):
        hybrid_score += 1

    scores = {
        "ГОСТ 7.0.100-2018": gost_score,
        "APA": apa_score,
        "Vancouver": vancouver_score,
        "Hybrid": hybrid_score
    }

    max_score = max(scores.values())
    if max_score < 2:
        return None

    for style, score in scores.items():
        if score == max_score:
            return style

    return None


def parse_hybrid_reference(ref: str) -> Tuple[Optional[str], List[str]]:
    """
    Парсит гибридный формат ссылок (Vancouver-подобный с элементами ГОСТ).
    """
    ref = ref.replace("\xa0", " ").strip()

    year_match = re.search(r'[A-Z]\.\s+((19|20)\d{2})\.', ref)
    if not year_match:
        year_match = re.search(r'\.\s+((19|20)\d{2})\.', ref)
    if not year_match:
        return None, []

    authors_part = ref[:year_match.start()].strip()
    authors = re.findall(r'([A-Z][a-z]+(?:-[a-z]+)?\s+[A-Z](?:\.?\s*[A-Z])*\.)', authors_part)

    if not authors:
        authors = re.findall(r'([A-Z][a-z]+(?:-[a-z]+)?)', authors_part)
        authors = [a for a in authors if len(a) > 2 and a.lower() not in ['and', 'the', 'of']]

    if not authors:
        return None, []

    after_year = ref[year_match.end():].strip()

    vol_pattern = re.search(r'\s+V\.?\s*\d+', after_year)
    if vol_pattern:
        before_vol = after_year[:vol_pattern.start()].strip()
    else:
        pages_pattern = re.search(r'\s+(?:P|S)\.?\s*\d+', after_year)
        if pages_pattern:
            before_vol = after_year[:pages_pattern.start()].strip()
        else:
            vol_pattern = re.search(r'\s+Bd\.?\s*\d+', after_year)
            if vol_pattern:
                before_vol = after_year[:vol_pattern.start()].strip()
            else:
                before_vol = after_year

    journal_pattern = re.search(r'([A-Z][a-z]+\.\s+(?:[A-Z][a-z]+\.\s*)+)$', before_vol)

    if journal_pattern:
        title = before_vol[:journal_pattern.start()].strip().rstrip('.')
    else:
        single_journal_pattern = re.search(r'\.\s+([A-Z][a-z]+)\.\s*$', before_vol)
        if single_journal_pattern:
            journal_candidate = single_journal_pattern.group(1)
            known_journals = ['Nature', 'Science', 'Cell', 'Genetics', 'Blood', 'Brain', 'Thorax', 'Cancer', 'Immunity', 'Development']
            if journal_candidate in known_journals or len(before_vol) > 80:
                title = before_vol[:single_journal_pattern.start()].strip().rstrip('.')
            else:
                title = before_vol.rstrip('.')
        else:
            single_journal_pattern2 = re.search(r'\.\s+([A-Z][a-z]+)\s*$', before_vol)
            if single_journal_pattern2:
                journal_candidate = single_journal_pattern2.group(1)
                if journal_candidate in known_journals or len(before_vol) > 80:
                    title = before_vol[:single_journal_pattern2.start()].strip().rstrip('.')
                else:
                    title = before_vol.rstrip('.')
            else:
                journal_dot_pattern = re.search(r'\.\s+([A-Z]\.\s+[A-Z][a-z]+)\s*$', before_vol)
                if journal_dot_pattern:
                    journal_candidate = journal_dot_pattern.group(1)
                    if len(before_vol) > 80 or 'J.' in journal_candidate:
                        title = before_vol[:journal_dot_pattern.start()].strip().rstrip('.')
                    else:
                        title = before_vol.rstrip('.')
                else:
                    title = before_vol.rstrip('.')

    if len(title) < 10:
        return None, []

    if re.match(r'^[A-Z][a-z]+\.\s+[A-Z][a-z]+\.$', title):
        return None, []

    return title, authors


def parse_gost_reference(ref: str) -> Tuple[Optional[str], List[str]]:
    """Парсит ГОСТ-ссылку (форматы с // и /)"""
    ref = ref.replace("\xa0", " ").strip()

    parts = GOST_JOURNAL_SPLIT.split(ref, maxsplit=1)
    if not parts:
        return None, []
    before_journal = parts[0].strip()

    slash_match = GOST_SLASH_SPLIT.search(before_journal)
    if slash_match:
        before_slash = before_journal[:slash_match.start()].strip()
        authors_str = before_journal[slash_match.end():].strip()
        authors_str = GOST_ET_AL.sub('', authors_str)
        authors = [a.strip() for a in authors_str.split(",") if a.strip()]

        if GOST_AUTHOR_START.match(before_slash):
            boundary = find_author_title_boundary(before_slash)
            if boundary != -1:
                title = before_slash[boundary:].strip().lstrip(". ")
            else:
                title = before_slash
        else:
            title = before_slash
    else:
        clean_text = GOST_ET_AL.sub(' ', before_journal)
        boundary = find_author_title_boundary(clean_text)
        if boundary == -1:
            return None, []
        title = clean_text[boundary:].strip().lstrip(". ")
        authors_str = clean_text[:boundary].strip().rstrip(".")
        authors = [a.strip() for a in authors_str.split(",") if a.strip()]

    return title if len(title) >= 5 else None, authors


def parse_apa_reference(ref: str) -> Tuple[Optional[str], List[str]]:
    """Парсит APA-ссылку (Автор, A. A. (Year). Title...)"""
    ref = ref.replace("\xa0", " ").strip()

    year_match = APA_YEAR_PATTERN.search(ref)
    if not year_match:
        year_match = APA_YEAR_ALT.search(ref)
        if not year_match:
            return None, []

    authors_part = ref[:year_match.start()].strip()
    after_year = ref[year_match.end():].strip()

    title_end = APA_TITLE_END.search(after_year)
    if title_end:
        title = after_year[:title_end.start()].strip()
    else:
        title_match = re.match(r'([^.]+)', after_year)
        if title_match:
            title = title_match.group(1).strip()
        else:
            return None, []

    authors = []
    author_parts = APA_AUTHOR_SPLIT.split(authors_part)
    for author_part in author_parts:
        author_part = author_part.strip().rstrip(',')
        if not author_part or author_part.lower() in ['et al.', 'and', '&']:
            continue
        authors.append(author_part)

    return title if len(title) >= 5 else None, authors


def parse_vancouver_reference(ref: str) -> Tuple[Optional[str], List[str]]:
    """Парсит Vancouver-ссылку (Автор AA, Автор BB. Title. Journal. Year;...)"""
    ref = ref.replace("\xa0", " ").strip()
    ref = VANCOUVER_NUMBER.sub('', ref)
    ref_clean = VANCOUVER_ET_AL.sub(' ETAL_MARKER ', ref)

    author_end = VANCOUVER_AUTHOR_END.search(ref_clean)
    title_group = 1

    if not author_end:
        author_end = VANCOUVER_AUTHOR_END_ALT.search(ref_clean)
        title_group = 1

    if not author_end:
        author_end = re.search(
            r'([A-Z]\.\s+)(\d|[A-Z][a-z]|[A-Z]{2,}|[A-Z][\-\d]|A\s+[a-z]|alpha|beta|gamma|delta)',
            ref_clean,
            re.IGNORECASE
        )
        title_group = 2

    if not author_end:
        author_end = VANCOUVER_AUTHOR_END_FALLBACK.search(ref_clean)
        title_group = 1
        if not author_end:
            return None, []

    title_start = author_end.start(title_group)
    authors_part = ref_clean[:title_start].strip().rstrip('.')
    authors_part = authors_part.replace('ETAL_MARKER', 'et al.')
    after_authors = ref_clean[title_start:].strip()
    after_authors = after_authors.replace('ETAL_MARKER', 'et al.')

    title = None

    title_end = VANCOUVER_JOURNAL_KEYWORDS.search(after_authors)
    if title_end:
        title = after_authors[:title_end.start()].strip()

    if not title:
        year_match = VANCOUVER_YEAR_PATTERN.search(after_authors)
        if year_match:
            title = after_authors[:year_match.start()].strip()

    if not title:
        year_match = VANCOUVER_YEAR_ALT.search(after_authors)
        if year_match:
            title = after_authors[:year_match.start()].strip()

    if not title:
        sentence_match = re.match(r'([^.]{20,})\.\s+[A-Z]', after_authors)
        if sentence_match:
            title = sentence_match.group(1).strip()

    if not title:
        sentences = re.split(r'\.\s+', after_authors)
        if sentences and len(sentences[0]) >= 15:
            title = sentences[0].strip()

    if not title:
        title_match = re.match(r'([^.]+)', after_authors)
        if title_match:
            title = title_match.group(1).strip()
        else:
            title = after_authors.strip()

    title = title.rstrip('.')

    if len(title) < 10 or len(title) < 30 and VANCOUVER_JOURNAL_KEYWORDS.search(title):
        alt_title = after_authors.split('.')[0].strip()
        if len(alt_title) > len(title):
            title = alt_title

    authors = []
    author_matches = VANCOUVER_AUTHOR_PATTERN.findall(authors_part)
    if author_matches:
        authors = [a.strip() for a in author_matches]
    else:
        for author in re.split(r',\s*', authors_part):
            author = author.strip().rstrip('.')
            if author and len(author) > 2 and not author.isdigit() and 'et al' not in author.lower():
                authors.append(author)

    if not authors:
        first_author_match = re.match(r'([A-Z][a-z]+\s+[A-Z]{1,3})', ref)
        if first_author_match:
            authors = [first_author_match.group(1).strip()]

    return title if len(title) >= 5 else None, authors


def find_author_title_boundary(text: str) -> int:
    """Находит границу между авторами и заголовком для ГОСТ"""
    roman_numerals = {'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X'}
    title_starters = {'The', 'An', 'A', 'On', 'In', 'To'}

    matches = list(GOST_BOUNDARY_PATTERN.finditer(text))
    for match in reversed(matches):
        letter_before = match.group(1)
        word_after = match.group(2)
        if letter_before in roman_numerals:
            continue
        if len(letter_before) != 1:
            continue
        if word_after in title_starters:
            return match.start() + 2
        if len(word_after) >= 3:
            return match.start() + 2

    for match in reversed(list(re.finditer(r'\.\s+', text))):
        pos = match.end()
        remaining = text[pos:]
        is_roman = any(remaining.startswith(r + '.') or remaining.startswith(r + ' ')
                       for r in roman_numerals)
        if is_roman:
            continue
        if re.match(r'^[A-Z]\.', remaining):
            continue
        if re.match(r'^[a-z]', remaining):
            continue
        return match.start()

    return -1


def normalize_title(title: str) -> str:
    """Нормализует заголовок для сравнения"""
    if not title:
        return ""
    title = title.split(":")[0]
    title = re.sub(r"[()\[\]]", "", title)

    greek_map = {
        "α": "alpha", "β": "beta", "γ": "gamma", "δ": "delta",
        "ε": "epsilon", "ζ": "zeta", "η": "eta", "θ": "theta",
        "ι": "iota", "κ": "kappa", "λ": "lambda", "μ": "mu",
        "ν": "nu", "ξ": "xi", "ο": "omicron", "π": "pi",
        "ρ": "rho", "σ": "sigma", "τ": "tau", "υ": "upsilon",
        "φ": "phi", "χ": "chi", "ψ": "psi", "ω": "omega"
    }
    for symbol, word in greek_map.items():
        title = title.replace(symbol, word)

    title = re.sub(r"[-–—‑:]", " ", title)
    title = re.sub(r"\s+", " ", title)
    return title.strip().lower()


def truncate_long_title(title: str) -> str:
    """Обрезает слишком длинный заголовок до разумной длины"""
    if len(title) <= MAX_TITLE_LENGTH:
        return title

    sentences = re.split(r'\.\s+', title)
    if sentences and len(sentences[0]) >= MIN_TITLE_LENGTH:
        return sentences[0]

    if len(title) > MAX_TITLE_LENGTH:
        truncated = title[:MAX_TITLE_LENGTH]
        last_space = truncated.rfind(' ')
        if last_space > MIN_TITLE_LENGTH:
            return truncated[:last_space]

    return title[:MAX_TITLE_LENGTH]


def validate_parsed_data(title: Optional[str], authors: List[str], ref_text: str = "") -> Tuple[bool, Optional[str]]:
    """Проверяет качество распарсенных данных перед поиском в PubMed"""
    if not title:
        return False, None

    if len(title) < MIN_TITLE_LENGTH:
        return False, None

    processed_title = truncate_long_title(title)

    words = [w for w in re.findall(r'\b[A-Za-z]{3,}\b', processed_title)
             if w.lower() not in {'the', 'and', 'for', 'with', 'from', 'that', 'this'}]
    if len(words) < MIN_WORDS_IN_TITLE:
        return False, None

    title_lower = processed_title.lower()

    journal_false_positives = [
        'j biol chem', 'j biol', 'j clin', 'j exp', 'j gen',
        'nature', 'science', 'cell', 'lancet', 'bmj', 'jama',
        'plos one', 'sci rep', 'nat commun', 'front',
        'frontiers', 'bmc', 'journal of', 'int j', 'ann oncol',
        'j name', 'journal name', 'j title',
        'mol biol', 'cell biol', 'biochem biophys',
        'biological journal', 'chemical journal',
    ]

    for fp in journal_false_positives:
        if title_lower == fp or title_lower.startswith(fp + ' '):
            if len(processed_title) < 50:
                return False, None

    if re.match(r'^j\s+[a-z]{3,15}$', title_lower.strip()):
        return False, None

    if re.match(r'^[A-Z][a-z]+\s+[A-Z]\.?\s*[A-Z]?\.?$', processed_title.strip()):
        if len(processed_title) < 25:
            return False, None

    journal_indicators = ['journal', ' j ', ' int ', ' ann ', ' proc ', ' transactions']
    title_padded = f" {title_lower} "
    if any(ind in title_padded for ind in journal_indicators):
        if len(processed_title) < 40:
            return False, None

    if '//' in processed_title:
        return False, None

    return True, processed_title


def normalize_author(author: str) -> str:
    """Нормализует имя автора для поиска"""
    if not author:
        return ""
    author = author.replace("'", "").replace("'", "")
    parts = [p.strip() for p in author.replace(".", " ").split() if p.strip()]
    surnames = [p for p in parts if len(p) > 2]
    return surnames[0] if surnames else (parts[0] if parts else "")


def fetch_article_data(pmid: str) -> Tuple[Optional[str], Optional[str]]:
    """Получает заголовок и DOI статьи по PMID"""
    try:
        handle = Entrez.efetch(db="pubmed", id=pmid, retmode="xml")
        records = Entrez.read(handle)
        handle.close()

        article = records['PubmedArticle'][0]
        title = article['MedlineCitation']['Article']['ArticleTitle']
        doi = None

        for aid in article.get('PubmedData', {}).get('ArticleIdList', []):
            if aid.attributes.get('IdType') == 'doi':
                doi = str(aid)
                break
        if not doi:
            for eid in article['MedlineCitation']['Article'].get('ELocationID', []):
                if eid.attributes.get('EIdType') == 'doi':
                    doi = str(eid)
                    break

        return title, doi
    except Exception:
        return None, None


# Кэш для результатов поиска
_search_cache = {}


def search_pubmed(title: Optional[str], authors: List[str], email: str,
                  journal: Optional[str] = None, year: Optional[str] = None) -> Tuple[Optional[str], Optional[str], str]:
    """
    Ищет статью в PubMed по заголовку и авторам.

    Args:
        title: Заголовок статьи
        authors: Список авторов
        email: Email для PubMed API
        journal: Название журнала (опционально)
        year: Год публикации (опционально)

    Returns:
        (pmid, doi, status) - PMID, DOI и статус поиска
    """
    Entrez.email = email if email else "linx2doi@example.com"

    if not title:
        return None, None, "no title"

    is_valid, processed_title = validate_parsed_data(title, authors)
    if not is_valid:
        _search_cache[title] = (None, None)
        return None, None, "invalid_parse"

    search_title = processed_title or title

    if search_title in _search_cache:
        pmid, doi = _search_cache[search_title]
        return pmid, doi, "cached"

    normalized_title = normalize_title(search_title)
    first_author = normalize_author(authors[0]) if authors else None

    clean_title = search_title
    greek_map = {
        "α": "alpha", "β": "beta", "γ": "gamma", "δ": "delta",
        "ε": "epsilon", "ζ": "zeta", "η": "eta", "θ": "theta",
        "ι": "iota", "κ": "kappa", "λ": "lambda", "μ": "mu",
        "ν": "nu", "ξ": "xi", "ο": "omicron", "π": "pi",
        "ρ": "rho", "σ": "sigma", "τ": "tau", "υ": "upsilon",
        "φ": "phi", "χ": "chi", "ψ": "psi", "ω": "omega"
    }
    for symbol, word in greek_map.items():
        clean_title = clean_title.replace(symbol, word)
    clean_title = re.sub(r'[:\[\]()"\'`/]', ' ', clean_title)
    clean_title = re.sub(r'\s+', ' ', clean_title).strip()

    spelling_variants = {
        'labelling': 'labeling', 'labelled': 'labeled',
        'colour': 'color', 'coloured': 'colored',
        'behaviour': 'behavior', 'favour': 'favor',
        'tumour': 'tumor', 'tumours': 'tumors',
        'centre': 'center', 'centres': 'centers',
        'analyse': 'analyze', 'analysed': 'analyzed',
        'catalyse': 'catalyze', 'catalysed': 'catalyzed',
        'recognise': 'recognize', 'recognised': 'recognized',
        'characterise': 'characterize', 'characterised': 'characterized',
        'ionisation': 'ionization', 'organisation': 'organization',
        'sulphate': 'sulfate', 'sulphur': 'sulfur',
        'oestrogen': 'estrogen', 'foetus': 'fetus',
        'haemoglobin': 'hemoglobin', 'haemorrhage': 'hemorrhage',
        'leukaemia': 'leukemia', 'oedema': 'edema',
        'aetiology': 'etiology', 'anaemia': 'anemia',
        'defence': 'defense', 'offence': 'offense',
        'licence': 'license', 'practise': 'practice',
    }

    stop_words = {
        'the', 'and', 'for', 'are', 'but', 'not', 'you', 'all', 'can', 'her', 'was',
        'one', 'our', 'out', 'its', 'his', 'has', 'had', 'how', 'may', 'who', 'did',
        'with', 'from', 'that', 'this', 'these', 'their', 'into', 'when', 'then',
        'where', 'which', 'while', 'been', 'being', 'have', 'were', 'what', 'than',
        'will', 'would', 'could', 'should', 'about', 'after', 'before', 'also',
        'between', 'during', 'through', 'under', 'within', 'without', 'upon',
        'review', 'study', 'studies', 'analysis', 'research', 'present', 'against',
        'effect', 'effects', 'role', 'roles', 'evidence', 'results', 'conclusion',
        'various', 'different', 'several', 'other', 'some', 'many', 'both', 'such',
        'using', 'based', 'related', 'associated', 'mediated', 'induced', 'novel',
        'new', 'case', 'report', 'update', 'current', 'recent', 'future', 'first',
        'further', 'following', 'possible', 'potential', 'important', 'significant',
    }

    low_priority_words = {
        'changes', 'levels', 'content', 'activity', 'properties', 'synthesis',
        'formation', 'production', 'reaction', 'products', 'model', 'system',
        'acid', 'acids', 'cells', 'human', 'cell', 'data', 'type', 'types',
    }

    clean_title_phrases = clean_title.replace('in vitro', 'invitro').replace('in vivo', 'invivo')
    clean_title_phrases = clean_title_phrases.replace('In vitro', 'invitro').replace('In vivo', 'invivo')

    raw_words = re.findall(r'[A-Za-z0-9]+(?:-[A-Za-z0-9]+)*', clean_title_phrases)
    high_priority_words = []
    low_priority_list = []
    for w in raw_words:
        if w.lower() == 'invitro':
            high_priority_words.append('"in vitro"')
            continue
        if w.lower() == 'invivo':
            high_priority_words.append('"in vivo"')
            continue
        if '-' in w:
            parts = w.split('-')
            for part in parts:
                if len(part) > 2 and part.lower() not in stop_words and not part.isdigit():
                    if part.lower() in low_priority_words:
                        low_priority_list.append(part)
                    else:
                        high_priority_words.append(part)
        elif len(w) > 2 and w.lower() not in stop_words:
            if w.lower() in low_priority_words:
                low_priority_list.append(w)
            else:
                high_priority_words.append(w)

    high_priority_words.sort(key=lambda x: -len(x))
    words = high_priority_words + low_priority_list
    keywords = words[:MAX_KEYWORDS]

    if not keywords:
        _search_cache[search_title] = (None, None)
        return None, None, "no keywords"

    keywords_with_variants = []
    for kw in keywords:
        kw_lower = kw.lower()
        if kw_lower in spelling_variants:
            us_variant = spelling_variants[kw_lower]
            if us_variant != kw:
                keywords_with_variants.append(f"({kw} OR {us_variant})")
            else:
                keywords_with_variants.append(kw)
        else:
            converted = None
            if kw_lower.endswith('lling'):
                converted = kw[:-5] + 'ling'
            elif kw_lower.endswith('lled'):
                converted = kw[:-4] + 'led'
            elif kw_lower.endswith('isation'):
                converted = kw[:-7] + 'ization'
            elif kw_lower.endswith('ised'):
                converted = kw[:-4] + 'ized'
            elif kw_lower.endswith('ise'):
                converted = kw[:-3] + 'ize'
            elif kw_lower.endswith('yse'):
                converted = kw[:-3] + 'yze'
            elif kw_lower.endswith('ysed'):
                converted = kw[:-4] + 'yzed'
            elif kw_lower.endswith('our'):
                converted = kw[:-3] + 'or'

            if converted and converted != kw:
                keywords_with_variants.append(f"({kw} OR {converted})")
            else:
                keywords_with_variants.append(kw)

    queries = []

    if first_author:
        query = " AND ".join([f"{w}[Title]" for w in keywords_with_variants])
        query += f" AND {first_author}[Author]"
        queries.append((query, RETMAX_HIGH_CONFIDENCE, "all_kw+author"))

    query = " AND ".join([f"{w}[Title]" for w in keywords_with_variants])
    queries.append((query, RETMAX_HIGH_CONFIDENCE, "all_kw"))

    if len(keywords_with_variants) > 3:
        query = " AND ".join([f"{w}[Title]" for w in keywords_with_variants[:3]])
        queries.append((query, RETMAX_MEDIUM_CONFIDENCE, "3_kw"))

    if first_author and len(keywords_with_variants) >= 2:
        query = f"{first_author}[Author] AND " + " AND ".join([f"{w}[Title]" for w in keywords_with_variants[:2]])
        queries.append((query, RETMAX_MEDIUM_CONFIDENCE, "author+2_kw"))

    if len(keywords_with_variants) >= 2:
        query = " AND ".join([f"{w}[Title]" for w in keywords_with_variants[:2]])
        queries.append((query, RETMAX_DESPERATE, "2_kw"))

    if first_author and journal and year:
        query = f"{first_author}[Author] AND {journal}[Journal] AND {year}[Date - Publication]"
        queries.append((query, RETMAX_HIGH_CONFIDENCE, "author+journal+year"))

    if journal and year and keywords_with_variants:
        kw_part = " AND ".join([f"{w}[Title]" for w in keywords_with_variants[:2]])
        query = f"{kw_part} AND {journal}[Journal] AND {year}[Date - Publication]"
        queries.append((query, RETMAX_MEDIUM_CONFIDENCE, "kw+journal+year"))

    if first_author and year:
        query = f"{first_author}[Author] AND {year}[Date - Publication]"
        queries.append((query, RETMAX_DESPERATE, "author+year"))

    if journal and year:
        query = f"{journal}[Journal] AND {year}[Date - Publication]"
        queries.append((query, RETMAX_DESPERATE, "journal+year"))

    if journal and not year:
        query = f"{journal}[Journal]"
        queries.append((query, RETMAX_DESPERATE, "journal_only"))

    best_pmid = None
    best_doi = None
    best_ratio = 0
    search_start_time = time.time()

    for query_idx, (query_str, retmax, stage) in enumerate(queries):
        elapsed_time = time.time() - search_start_time
        if elapsed_time > SEARCH_TIMEOUT:
            if best_pmid and best_ratio >= FUZZY_THRESHOLD:
                _search_cache[search_title] = (best_pmid, best_doi)
                return best_pmid, best_doi, f"timeout_found_{best_ratio}%"
            else:
                _search_cache[search_title] = (None, None)
                return None, None, "timeout_not_found"

        try:
            handle = Entrez.esearch(db="pubmed", term=query_str, retmax=retmax)
            record = Entrez.read(handle)
            handle.close()
            time.sleep(REQUEST_DELAY)
        except Exception:
            continue

        for pmid in record.get("IdList", []):
            fetched_title, doi = fetch_article_data(pmid)
            if not fetched_title:
                continue
            normalized_fetched = normalize_title(fetched_title)
            ratio_full = fuzz.ratio(normalized_title, normalized_fetched)
            ratio_partial = fuzz.partial_ratio(normalized_title, normalized_fetched)
            ratio = max(ratio_full, ratio_partial - 5)
            if ratio > best_ratio:
                best_ratio = ratio
                best_pmid = pmid
                best_doi = doi

            if ratio >= EXACT_MATCH_THRESHOLD:
                _search_cache[search_title] = (best_pmid, best_doi)
                return best_pmid, best_doi, f"exact_{ratio}%_{stage}"

        if best_ratio >= GOOD_MATCH_THRESHOLD:
            break
        elif best_ratio >= FUZZY_THRESHOLD:
            if query_idx > 1:
                break

    if best_ratio >= FUZZY_THRESHOLD:
        _search_cache[search_title] = (best_pmid, best_doi)
        return best_pmid, best_doi, f"fuzzy_{best_ratio}%"
    else:
        _search_cache[search_title] = (None, None)
        return None, None, "not found"


def search_doi_by_reference(ref_text: str, email: str) -> Tuple[Optional[str], str]:
    """
    Ищет DOI по тексту библиографической ссылки.
    
    Args:
        ref_text: Текст библиографической ссылки
        email: Email для PubMed API
        
    Returns:
        (doi, status) - DOI и статус поиска
    """
    # Автоопределение стиля
    detected_style = detect_reference_style(ref_text)
    
    if detected_style:
        # Парсим ссылку согласно определённому стилю
        if detected_style == "ГОСТ 7.0.100-2018":
            title, authors = parse_gost_reference(ref_text)
        elif detected_style == "APA":
            title, authors = parse_apa_reference(ref_text)
        elif detected_style == "Vancouver":
            title, authors = parse_vancouver_reference(ref_text)
        elif detected_style == "Hybrid":
            title, authors = parse_hybrid_reference(ref_text)
        else:
            title, authors = None, []
    else:
        # Пробуем все парсеры по очереди
        title, authors = None, []
        for parser_func in [parse_vancouver_reference, parse_apa_reference, parse_gost_reference, parse_hybrid_reference]:
            t, a = parser_func(ref_text)
            if t:
                title, authors = t, a
                break
    
    if not title:
        return None, "parse_failed"
    
    # Извлекаем журнал и год для улучшения поиска
    journal = None
    year = None
    if detected_style:
        journal, year = extract_journal_and_year(ref_text, detected_style)
    
    # Ищем в PubMed
    pmid, doi, status = search_pubmed(title, authors, email, journal, year)
    
    return doi, status


def clear_cache():
    """Очищает кэш результатов поиска"""
    _search_cache.clear()
