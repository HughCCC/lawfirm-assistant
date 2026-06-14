"""Web search service for legal case discovery.

Provides a pluggable search backend abstraction for finding legal news articles
and case references. Default backend is Bing scraping (no API key required).
Additional backends (Bing API, Google CSE, etc.) can be added by implementing
the same protocol.

Data flow:
    question → build_search_queries() → _search_bing_scrape() → SearchResults
    SearchResults → _extract_case_numbers() → whitelist set
    SearchResults → format_search_results_for_prompt() → LLM context string
"""

import logging
import re
import time
from dataclasses import dataclass, field
from typing import Optional
from urllib.parse import urlparse, parse_qs

import httpx
from lxml import html

logger = logging.getLogger("legal-report")

# ─── Case Number Regex (same pattern as generator.py) ──────────────────

_CASE_NUMBER_RE = re.compile(
    r'（\d{4}）'                          # （年份）
    r'[一-龥]{1,4}\d{0,6}'               # 法院代字（1-4汉字 + 可选数字）
    r'(?:民|刑|行|赔|执|破|申|再|知|商|海|环|少|家)'  # 案件类型
    r'(?:初|终|再|监|抗|重|复|申|调|确|认|裁|命)'    # 审判程序
    r'\d+'                                 # 案件序号
    r'[号字]'                              # 号/字结尾
)

# ─── Data Classes ──────────────────────────────────────────────────────

@dataclass
class SearchResult:
    """A single web search result with extracted case numbers."""
    title: str
    url: str
    snippet: str
    case_numbers: list[str] = field(default_factory=list)


# ─── Query Construction ────────────────────────────────────────────────

# Legal-domain suffixes appended to user question keywords
_LEGAL_SUFFIXES = [
    "判决 案例 案号",
    "法院 裁判 新闻",
]

# Keywords that indicate a legal question worth searching for
_LEGAL_KEYWORDS = [
    '赌博', '赌场', '抽卡', '游戏', '交易', '变现', '合同',
    '侵权', '知识产权', '商标', '专利', '劳动', '仲裁', '刑事',
    '民事', '行政', '公司', '股权', '投资', '金融', '证券',
    '数据', '隐私', '网络', '电商', '房地产', '拆迁', '婚姻',
    '继承', '税务', '海关', '反垄断', '不正当竞争', '环保',
]


def build_search_queries(question: str, max_queries: int = 2) -> list[str]:
    """Build 2-3 targeted search queries from a legal question.

    Strategy:
    1. Extract key legal terms from the question
    2. Generate focused queries with legal-domain suffixes
    3. Limit to max_queries distinct queries

    Args:
        question: The user's legal question (may be truncated).
        max_queries: Maximum number of distinct queries to return.

    Returns:
        List of search query strings.
    """
    # Truncate to first 80 chars for keyword extraction
    short = question[:80]

    # Extract potential legal terms
    found_terms = []
    for kw in _LEGAL_KEYWORDS:
        if kw in short:
            found_terms.append(kw)

    queries = []

    if found_terms:
        # Query 1: extracted terms + legal case suffix
        core_terms = ' '.join(found_terms[:5])
        queries.append(f'{core_terms} {_LEGAL_SUFFIXES[0]}')
    else:
        # Fallback: use the question text directly
        queries.append(f'{short.strip()} {_LEGAL_SUFFIXES[0]}')

    # Query 2: broader legal search
    if len(queries) < max_queries:
        if found_terms:
            broader = ' '.join(found_terms[:3])
            queries.append(f'{broader} {_LEGAL_SUFFIXES[1]}')
        else:
            queries.append(f'{short.strip()} {_LEGAL_SUFFIXES[1]}')

    return queries[:max_queries]


# ─── Case Number Utilities ─────────────────────────────────────────────

def _extract_case_numbers(text: str) -> list[str]:
    """Extract Chinese case numbers from text using the standard regex.

    Args:
        text: Text to scan for case numbers.

    Returns:
        List of matched case number strings (deduplicated, order preserved).
    """
    seen = set()
    result = []
    for match in _CASE_NUMBER_RE.finditer(text):
        cn = match.group(0)
        if cn not in seen:
            seen.add(cn)
            result.append(cn)
    return result


def normalize_case_number(cn: str) -> str:
    """Normalize a case number for fuzzy whitelist comparison.

    Handles common formatting variations:
    - Half-width parentheses → full-width
    - Internal whitespace → removed
    - Leading/trailing whitespace → stripped

    Args:
        cn: Raw case number string.

    Returns:
        Normalized case number string suitable for set membership check.
    """
    cn = cn.strip()
    # Replace half-width parens with full-width
    cn = cn.replace('(', '（').replace(')', '）')
    # Remove all whitespace (spaces, tabs, newlines, etc.)
    cn = re.sub(r'\s+', '', cn)
    return cn


# ─── Bing Scraping Backend ─────────────────────────────────────────────

_BING_SEARCH_URL = "https://www.bing.com/search"
_BING_USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
)


def _search_bing_scrape(
    query: str,
    max_results: int = 8,
    timeout: int = 15,
) -> list[SearchResult]:
    """Search Bing and extract results via HTML scraping.

    Uses XPath to parse Bing's result page structure. Bing's HTML uses
    <li class="b_algo"> for organic results, with <h2><a> for titles,
    <p> for snippets, and the href attribute on <a> for URLs.

    Args:
        query: Search query string.
        max_results: Maximum number of results to return.
        timeout: HTTP request timeout in seconds.

    Returns:
        List of SearchResult objects (may be empty if parsing fails).

    Raises:
        httpx.HTTPError: On network/HTTP errors (caught by caller).
    """
    params = {
        'q': query,
        'setlang': 'zh-Hans',
        'cc': 'zh',  # Prefer Chinese content
    }
    headers = {'User-Agent': _BING_USER_AGENT}

    resp = httpx.get(
        _BING_SEARCH_URL,
        params=params,
        headers=headers,
        timeout=timeout,
        follow_redirects=True,
    )
    resp.raise_for_status()

    tree = html.fromstring(resp.text)

    # Bing organic results are in <li class="b_algo">
    result_elements = tree.xpath('//li[contains(@class, "b_algo")]')

    results: list[SearchResult] = []
    for elem in result_elements:
        if len(results) >= max_results:
            break

        # Extract title + URL from <h2><a>
        title_el = elem.xpath('.//h2/a')
        if not title_el:
            continue
        title = title_el[0].text_content().strip()
        href = title_el[0].get('href', '')

        # Skip empty or javascript links
        if not title or not href or href.startswith('javascript:'):
            continue

        # Extract snippet from first <p> inside the result
        snippet_el = elem.xpath('.//p')
        snippet = ""
        if snippet_el:
            snippet = snippet_el[0].text_content().strip()

        # Deduplicate by URL
        if any(r.url == href for r in results):
            continue

        # Extract case numbers from title + snippet
        case_numbers = _extract_case_numbers(title + ' ' + snippet)

        results.append(SearchResult(
            title=title,
            url=href,
            snippet=snippet,
            case_numbers=case_numbers,
        ))

    return results


# ─── Main Search Entry Point ───────────────────────────────────────────

def search_for_cases(
    question: str,
    max_results: int = 8,
    timeout: int = 15,
    max_queries: int = 2,
    query_delay: float = 2.0,
) -> tuple[list[SearchResult], set[str]]:
    """Search the web for legal cases related to the question.

    This is the main entry point called by the report generator. It:
    1. Builds targeted search queries from the question
    2. Searches Bing for each query
    3. Extracts case numbers from all results
    4. Builds a normalized whitelist for validation
    5. Returns both the results (for prompt injection) and the whitelist

    Graceful degradation: any exception during search is caught and logged,
    and empty results + empty whitelist are returned. The caller can
    distinguish "search worked but found nothing" from "search failed" via
    log level (info vs warning).

    Args:
        question: The user's legal question.
        max_results: Max total search results across all queries.
        timeout: HTTP request timeout per search.
        max_queries: Max distinct search queries to send.
        query_delay: Seconds delay between queries (rate limiting).

    Returns:
        Tuple of (search_results, case_whitelist_set).
        - search_results: list of SearchResult for prompt injection
        - case_whitelist_set: set of normalized case numbers for validation
    """
    queries = build_search_queries(question, max_queries)
    logger.info("Search: %d queries for question (len=%d): %s",
                len(queries), len(question), queries)

    all_results: list[SearchResult] = []
    seen_urls: set[str] = set()

    for qi, query in enumerate(queries):
        if qi > 0:
            time.sleep(query_delay)

        try:
            batch = _search_bing_scrape(query, max_results=max_results, timeout=timeout)
            new_count = 0
            for r in batch:
                if r.url not in seen_urls:
                    seen_urls.add(r.url)
                    all_results.append(r)
                    new_count += 1
            logger.info("Search query '%s': %d results (%d new, %d total)",
                        query[:60], len(batch), new_count, len(all_results))
        except Exception as e:
            logger.warning("Search query '%s' failed: %s", query[:60], e)
            # Continue with next query
            continue

        if len(all_results) >= max_results:
            break

    # Build case number whitelist from all results
    whitelist: set[str] = set()
    for r in all_results:
        for cn in r.case_numbers:
            whitelist.add(normalize_case_number(cn))

    logger.info("Search complete: %d results, %d unique case numbers in whitelist",
                len(all_results), len(whitelist))

    return all_results, whitelist


# ─── Prompt Formatting ─────────────────────────────────────────────────

def format_search_results_for_prompt(results: list[SearchResult]) -> str:
    """Format search results as a structured block for LLM context injection.

    Results with case numbers are listed first (most actionable for the LLM),
    followed by results without case numbers.

    The output format is designed to be compact but informative, giving the LLM
    enough context to decide which cases are relevant.

    Args:
        results: List of SearchResult objects.

    Returns:
        Formatted string ready for injection into the user prompt.
    """
    if not results:
        return ""

    # Split: results with case numbers first
    with_cases = [r for r in results if r.case_numbers]
    without_cases = [r for r in results if not r.case_numbers]

    lines = [
        "═══════════════════════════════════════════════════════════",
        "网络检索结果（法律案例参考）",
        "═══════════════════════════════════════════════════════════",
        "",
        "以下是通过网络检索找到的法律相关新闻/案例。你可以引用其中的案号（必须附带来源URL）。",
        "每个案号引用都需标注对应的来源链接。如果某案例与本案关联性不强，请如实说明。",
        '如果没有找到高度相关的案例，使用一般性描述（如"在类似纠纷的司法实践中…"）。',
        "",
    ]

    idx = 1
    if with_cases:
        lines.append("── 含案号的结果 ──")
        lines.append("")
        for r in with_cases:
            cn_str = "、".join(r.case_numbers)
            lines.append(f"[{idx}] 标题：{r.title}")
            lines.append(f"    案号：{cn_str}")
            lines.append(f"    摘要：{r.snippet[:200]}")
            lines.append(f"    来源：{r.url}")
            lines.append("")
            idx += 1

    if without_cases:
        lines.append("── 其他相关结果 ──")
        lines.append("")
        for r in without_cases:
            lines.append(f"[{idx}] 标题：{r.title}")
            lines.append(f"    摘要：{r.snippet[:200]}")
            lines.append(f"    来源：{r.url}")
            lines.append("")
            idx += 1

    lines.append("（以上搜索结果仅供参考，请自行判断案例相关性）")
    lines.append("")

    return "\n".join(lines)
