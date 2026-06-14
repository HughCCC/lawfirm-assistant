"""Report generator: system prompt, JSON validation, and retry logic."""

import json
import logging
import re
import sys
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date
from pathlib import Path
from typing import Optional

import httpx

# Import scan_json_content from the existing generate_docx module
SCRIPTS_DIR = (Path(__file__).resolve().parent.parent / ".law-research" / "scripts")
sys.path.insert(0, str(SCRIPTS_DIR))
from generate_docx import scan_json_content

from config import (
    MAX_RETRIES, WEBAPP_OUTPUT_DIR, MAX_OUTPUT_TOKENS, DEFAULT_TEMPERATURE,
    URL_VALIDATION_ENABLED, URL_VALIDATION_TIMEOUT, URL_VALIDATION_MAX_CONCURRENT,
    URL_VALIDATION_MAX_URLS, URL_SKIP_DOMAINS,
    SEARCH_ENABLED, SEARCH_MAX_RESULTS, SEARCH_TIMEOUT,
    SEARCH_MAX_QUERIES, SEARCH_QUERY_DELAY,
)
from search_service import normalize_case_number
from models import (
    GenerateRequest,
    Report,
    Section,
    TableContent,
    ValidationError,
)
from llm_service import create_provider, LLMResult

logger = logging.getLogger("legal-report")


# ─── System Prompt ────────────────────────────────────────────────────

def build_system_prompt() -> str:
    """Build the system prompt encoding the BIRACS framework, JSON schema,
    formatting rules, and forbidden patterns."""

    return """你是一位资深中国执业律师，正在撰写一份正式的法律评估意见书。

你必须**只输出一个合法的 JSON 对象**——不得有任何 Markdown 格式、不得有代码块包裹（```json）、
不得在 JSON 前后添加任何解释性文字。回复必须以 `{` 开头，以 `}` 结尾。

═══════════════════════════════════════════════════════════
JSON SCHEMA（必须严格遵循）
═══════════════════════════════════════════════════════════

JSON 对象有两个顶层字段：
- "title": string —— 报告标题，格式为"关于{事项}的法律评估意见"
- "sections": array —— section 对象数组

每个 section 对象包含：
- "type": string —— 取值为 "h1"、"h2"、"h3"、"h4"、"h5"、"body"、"table"、"source_item"
- "content": string 或 object（仅 type="table" 时为 object）

═══════════════════════════════════════════════════════════
CONTENT 内容规则
═══════════════════════════════════════════════════════════

【标题类型 h1-h5 的 content 格式】
- "h1": "一、章节标题"   （使用中文数字：一二三四五六七八九十）
- "h2": "（一）子标题"    （括号中文数字）
- "h3": "1、三级标题"     （阿拉伯数字 + 、）
- "h4": "（1）四级标题"   （括号阿拉伯数字）
- "h5": "① 五级标题"     （圈号数字）
- "source_item": "（一）来源名称"  （用于信息来源条目，括号中文数字）

【正文类型 body 的 content 格式】
- 纯中文文本。用 \\n\\n 分隔段落。支持 \\n 做段内换行。
- 不得包含任何 Markdown 语法。

【表格类型 table 的 content 格式】
必须是一个对象，包含以下字段：
{
  "caption": "表1：表格标题",
  "headers": ["列1", "列2", ...],
  "rows": [["单元格", "单元格", ...], ...],
  "note": "表格注释（可选，如数据来源说明）",
  "footnote_on_col": "需要脚注标记的列名（可选）",
  "footnote_urls": ["https://...", "https://...", ...]
}
说明：
- footnote_on_col 指定某列需要在文本后面追加上标脚注标记（¹²³…），
  对应的 URL 从 footnote_urls 数组按行依次读取（与 rows 一一对应）。
- footnote_urls 仅在指定了 footnote_on_col 时需要，其他情况下可省略。
- 如果某行没有对应的脚注 URL，该位置的 footnote_urls 设为空字符串 ""。

═══════════════════════════════════════════════════════════
报告结构（BIRACS 框架）— 必须严格按照此顺序生成 6 个 h1 章节
═══════════════════════════════════════════════════════════

1. 一、背景介绍 [h1]
   - 客户基本情况、业务场景描述、面临的核心问题
   - 后跟 [body] 正文

2. 二、评估结论 [h1]（结论前置原则——先给结论再分析）
   - 用加粗风格的核心判断句概括法律评估结果
   - 后跟 [body] 正文（2-3 段）

3. 三、法律分析 [h1]
   - 按子问题展开，每个子问题用 [h2] 开启
   - 每个子问题内部可嵌套 [h3][h4][h5] 层级
   - [h2] 标题格式："（一）子问题描述"
   - 每个 [h2] 后跟 [body] 正文分析
   - 引用法条时使用《》书名号
   - 法律分析需覆盖：法条文义解释、立法目的解释、类推适用分析、案例参考

4. 四、行业实践与司法案例 [h1]
   - [h2] "（一）司法案例" → 包含 [body] 案例分析（讨论法律适用趋势和裁判观点）
   - **案号引用规则：你可以引用用户提示中"网络检索结果"部分提供的真实案号（如果存在）。每个引用的案号必须附带其来源URL作为脚注链接。**
   - **绝对禁止编造用户提示中未提供的案号。如果没有搜索结果支持，使用一般性描述（如"在类似纠纷的司法实践中，法院通常认定…"）。**
   - [h2] "（二）行业对标" → [body] 行业做法分析
   - [h2] "（三）行政处罚参考" → [body] 行政处罚情况

5. 五、风险评级与合规建议 [h1]
   - [h2] "（一）风险评级" → 包含 [table] 风险评级表（columns：风险场景、风险等级、风险描述、对应建议）
   - 风险等级分为四档：低风险（绿色）、中风险（黄色）、高风险（橙色）、极度风险（红色）
   - [h2] "（二）合规建议" → [body] 总体建议介绍
   - 然后每个具体建议用 [h3] 标题 + [body] 展开，共 3-6 条

6. 六、信息来源 [h1]
   - [h2] "（一）法律法规" → 每项法规用 [h3] + [body] 列出
   - [h2] "（二）相关案例" → 列出你引用的案例。如果你在前文中没有引用任何确认真实的案例，此处直接说明"经检索，未发现与本案事实高度相似的公开判例，以上分析基于法律原则和一般司法实践"，不得编造案例来填充此部分
   - [h2] "（三）其他参考文件" → 每个文献用 [h3] + [body] 列出
   - 最后以 [body] 结尾，包含最后检索日期（使用用户提示中提供的当前日期，不得编造其他日期）和免责声明
   - 此部分**不得使用 table 类型**，必须使用 h2+h3+body 层级结构

═══════════════════════════════════════════════════════════
格式禁则（以下内容绝对禁止出现在任何字符串中）
═══════════════════════════════════════════════════════════

- Markdown 加粗：**文字** 或 __文字__
- Markdown 标题：# ## ### 等井号标题
- Markdown 斜体：*文字* 或 _文字_
- Markdown 列表：- 或 * 开头的无序列表、1. 开头的有序列表
- 行内代码：`代码`
- 代码块：``` 包裹
- 引用：> 开头的 blockquote
- 链接语法：[文字](url)
- 删除线：~~文字~~
- 分割线：--- === *** ___ ###
- 管道表格：| 组成的伪表格（请使用 table 类型代替！）
- Emoji 符号：⚠️ 📋 ❌ ✅ 🎯 💡 🔴 🟢 🟡 等任何 emoji
- AI 风格标签：【风险评估】【合规建议】【重要提示】【分析结论】【注意事项】等方括号装饰标签
- ASCII 装饰线：------ ====== ****** ###### 等
- 连续感叹号或问号：三个以上 !!! 或 ???
- 半角英文标点在中文语境中的滥用

═══════════════════════════════════════════════════════════
写作规范
═══════════════════════════════════════════════════════════

- 使用正式的法律文书中文，避免口语化表达
- 每条法律引用尽可能附带来源 URL
- 每条结论必须有法律依据支撑
- 引用案例时：如果用户提示中有网络检索结果提供的案号，可以引用（必须附带来源URL）；如果没有搜索结果支持，使用一般性司法实践趋势描述
- 法律/法规/规范性文件名称使用《》书名号
- 使用中文弯引号 " "（U+201C/U+201D），不使用 ASCII 直引号 " "
- 风险描述使用建设性语气，不制造恐慌，不在法律不确定处做绝对化判断
- 中文标点符号：， 。 ； ： 、 ？ ！ — …… · 《 》 【 】 （ ） " "
- 在法律分析中引用案例时：优先使用用户提示中网络检索结果提供的案号（附带来源URL）；如果没有搜索结果支持，使用一般性描述（如"在类似合同纠纷的司法实践中，法院通常认定…"）。绝对禁止编造搜索结果中不存在的案号。

═══════════════════════════════════════════════════════════
真实性要求（绝对禁则）
═══════════════════════════════════════════════════════════

1. 只能引用用户提示中"网络检索结果"部分提供的案号。不得编造搜索结果中不存在的案号。如果搜索结果中没有找到相关案号，使用一般性描述（如"在类似纠纷的司法实践中，法院通常认定…"）。不要尝试猜测或编造案号——你没有案例数据库，无法独立确认案号真实性。
2. 只引用真实、公开、可溯源的内容。每条法律依据和案例引用必须有可查证的来源。
3. 案例相关性红线：引用的案例必须同时满足"法律领域相同"和"事实情节高度相似"两个条件，严禁仅因涉及类似罪名或技术领域就引用不相关案例。
4. URL 要求：每个引用的案号必须附带其来源URL（来自用户提示中的网络检索结果）。只提供你确信存在且当前可访问的 URL。如果无法确定链接存在，删除该 URL，仅保留文字描述。绝对禁止使用 example.com、test.com 等占位域名，也禁止使用 example、test、demo、placeholder 等词汇作为 URL 组成部分。
5. 宁可少引用，不要编造。如果无法找到足够的相关真实案例，明确注明"经检索未发现高度相关的公开判例"，并只引用确实存在的案例。不要虚构案例来填充篇幅。
6. 对于无法核实的信息，必须在正文中明确标注"待核实"或"仅供参考"，不得以肯定语气表述未经证实的内容。

═══════════════════════════════════════════════════════════
重要提示
═══════════════════════════════════════════════════════════

1. 你的整个回复必须是一个合法的 JSON 对象。以 { 开头，以 } 结尾。
2. 不要用 ```json 代码块包裹 JSON。直接输出裸 JSON。
3. 不要在 JSON 前后添加任何说明文字。
4. JSON 中所有字符串值内部的引号必须使用中文弯引号 " "（或者被转义为 \\" 的普通引号）
5. body 段落之间使用 \\n\\n 分隔（在 JSON 字符串中写作 \\\\n\\\\n）
6. 信息来源部分使用 h2+h3+body 结构，不使用 table
7. 报告需要具备实质性内容，不要使用占位符或省略号
8. 表格的 rows 数组不能为空，至少要有实际案例/数据
9. 案号引用规则：只能引用用户提示中网络检索结果提供的案号（附带来源URL）。没有搜索结果支持的案例使用一般性描述。绝对禁止编造案号。"""


# ─── User Prompt ──────────────────────────────────────────────────────

def build_user_prompt(question: str, search_context: str = "") -> str:
    """Build the user prompt wrapping the legal question.

    Args:
        question: The user's legal question.
        search_context: Optional formatted search results from web search.
                        Injected between the question and the formatting instructions.
    """
    prompt = f"""请就以下法律问题，生成一份完整的 BIRACS 格式法律评估意见书 JSON。

用户问题：
{question}"""

    if search_context:
        prompt += f"""

{search_context}"""

    # Inject today's date so the LLM doesn't fabricate one
    today_str = date.today().strftime('%Y年%m月%d日')

    prompt += f"""

请严格按照系统提示中的 JSON Schema 和 BIRACS 框架输出。确保：
1. 6 个 h1 章节完整覆盖
2. 法律分析部分有实质性的法条解读和案例支撑
3. 信息来源部分使用 h2+h3+body 结构，列明你引用的法律、案例和参考文献
4. 不包含任何 Markdown 格式、emoji、AI 风格标签
5. 输出裸 JSON，不要用代码块包裹
6. 案例引用规则：可以使用上述网络检索结果中的真实案号（附带来源URL），不得编造检索结果中不存在的案号。URL 必须指向可访问的真实页面，禁止使用 example/test/demo/placeholder 占位符
7. 当前日期：{today_str}。报告中所有涉及日期的地方（如最后检索日期）必须使用此日期，不得编造其他日期。"""
    return prompt


# ─── Retry Prompt ─────────────────────────────────────────────────────

def build_retry_prompt(errors: list[ValidationError], original_question: str, search_context: str = "") -> str:
    """Build a retry prompt listing specific validation errors.

    Args:
        errors: List of validation errors from the previous attempt.
        original_question: The original user question.
        search_context: Optional formatted search results (re-injected on retry).
    """
    error_lines = []
    has_url_errors = False
    has_case_errors = False
    for i, err in enumerate(errors, 1):
        loc = f"（第{err.section_index}个section）" if err.section_index is not None else ""
        error_lines.append(f"{i}. {err.field}{loc}：{err.message}")
        if err.field == "url":
            has_url_errors = True
        if "案号" in err.message:
            has_case_errors = True

    error_text = "\n".join(error_lines)

    url_guidance = ""
    if has_url_errors:
        url_guidance = """
═══════════════════════════════════════════════════════════
URL 修正指引（必读）
═══════════════════════════════════════════════════════════
- 将不可达的 URL 替换为你确信存在的真实网页链接
- 如果无法确认链接可访问，直接删除该 URL，仅保留文字描述
- 绝对禁止使用 example.com、test.com 等占位域名的 URL
- 绝对禁止用案号替代 URL——你没有案例数据库，不知道真实案号
- 移除所有指向 test/demo/placeholder 等测试性地址的 URL"""

    case_guidance = ""
    if has_case_errors:
        if search_context:
            case_guidance = """
═══════════════════════════════════════════════════════════
案号修正指引（必读）
═══════════════════════════════════════════════════════════
- 检查你使用的案号是否全部来自用户提示中的网络检索结果
- 删除检索结果中不存在的案号，或用一般性描述替代
- 每个案号引用必须附带来源URL（检索结果中提供的链接）
- 如果检索结果中的案例与本案关联性不强，如实说明关联性限制
- 宁可少写案例，绝不编造案号"""
        else:
            case_guidance = """
═══════════════════════════════════════════════════════════
案号修正指引（必读）
═══════════════════════════════════════════════════════════
- 删除报告中所有具体案号（网络搜索未找到可验证的案例）
- 用一般性描述替代，如"在类似合同纠纷的司法实践中，法院通常认定…"
- 你没有案例数据库的访问权限——你生成的任何具体案号都是虚构的
- "相关案例"部分如果无法引用确认真实的案例，直接说明"经检索，未发现与本案事实高度相似的公开判例，以上分析基于法律原则和一般司法实践"
- 宁可少写案例，绝不编造案号"""

    # Build retry prompt body
    prompt_body = f"""你上一次输出的 JSON 存在以下错误，请修正后重新输出完整的 JSON：

{error_text}
{url_guidance}
{case_guidance}"""

    # Re-inject search context so LLM sees it again on retry
    if search_context:
        prompt_body += f"""

{search_context}"""

    prompt_body += f"""

原始用户问题：
{original_question}

请修正以上所有错误，重新生成完整的 JSON。记住：只输出裸 JSON，以 {{ 开头，以 }} 结尾。"""

    return prompt_body


# ─── JSON Cleaning ────────────────────────────────────────────────────

def clean_json_text(text: str) -> str:
    """Strip markdown code fences and extract the JSON object from LLM response."""
    text = text.strip()

    # Remove markdown code blocks
    if text.startswith("```"):
        # Find the first newline after opening fence
        nl = text.find("\n")
        if nl != -1:
            text = text[nl + 1:]
        # Remove closing fence
        if text.endswith("```"):
            text = text[:-3]
        text = text.strip()

    # Handle case where LLM wraps in ```json ... ```
    if text.startswith("json\n"):
        text = text[5:].strip()
        if text.endswith("```"):
            text = text[:-3].strip()

    # Find the outermost JSON object
    # Look for the first '{' and the matching last '}'
    start = text.find("{")
    end = text.rfind("}")

    if start == -1 or end == -1 or start >= end:
        return text  # Return as-is, validation will catch the error

    return text[start:end + 1]


# ─── URL Extraction & Validation ──────────────────────────────────────

# Regex: match http/https URLs in text
_URL_RE = re.compile(r'https?://[^\s一-鿿"\'。，；：、》）】]+')


def extract_urls_from_json(data, max_urls: int = URL_VALIDATION_MAX_URLS) -> list[str]:
    """Recursively extract all http/https URLs from a JSON-compatible structure.

    Filters out placeholder URLs containing example, test, demo, placeholder,
    localhost, or 127.0.0.1.

    Args:
        data: Parsed JSON (dict, list, str, or scalar).
        max_urls: Cap on total URLs returned.

    Returns:
        Deduplicated list of real (non-placeholder) URLs.
    """
    urls = set()

    # Forbidden markers in a URL → skip it
    _FORBIDDEN = ('example', 'test', 'demo', 'placeholder', 'localhost', '127.0.0.1', '0.0.0.0')

    def _walk(obj):
        if len(urls) >= max_urls:
            return
        if isinstance(obj, dict):
            for v in obj.values():
                _walk(v)
        elif isinstance(obj, list):
            for item in obj:
                _walk(item)
        elif isinstance(obj, str):
            for match in _URL_RE.finditer(obj):
                if len(urls) >= max_urls:
                    return
                url = match.group(0)
                # Strip trailing punctuation that's not part of the URL
                url = url.rstrip('.,;:!?)')
                lower = url.lower()
                if any(f in lower for f in _FORBIDDEN):
                    continue
                urls.add(url)

    _walk(data)
    return list(urls)[:max_urls]


def _check_single_url(url: str, timeout: int = URL_VALIDATION_TIMEOUT) -> Optional[str]:
    """Check a single URL for reachability.

    Strategy: HEAD first (fast), fall back to GET if HEAD returns 405/400.
    Follows redirects (max 5). SSL verification disabled for compatibility
    with government sites that may use self-signed certificates.

    Args:
        url: The URL to check.
        timeout: Request timeout in seconds.

    Returns:
        None if the URL is reachable (HTTP 2xx/3xx/4xx except 404/410/403).
        Error message string if unreachable or an exception occurs.
    """
    try:
        with httpx.Client(timeout=timeout, follow_redirects=True, max_redirects=5, verify=False) as client:
            # HEAD first
            try:
                resp = client.head(url)
                if resp.status_code in (405, 400):
                    # Method not allowed — fall back to GET
                    resp = client.get(url)
            except httpx.HTTPError:
                # HEAD failed — try GET
                resp = client.get(url)

            # 2xx = success
            if 200 <= resp.status_code < 300:
                return None
            # 3xx redirect (shouldn't happen with follow_redirects, but just in case)
            if 300 <= resp.status_code < 400:
                return None
            # 429 Too Many Requests — not an error (rate limiting)
            if resp.status_code == 429:
                return None
            # 4xx client errors — 404/410 = gone, 403 = forbidden
            if resp.status_code in (404, 410):
                return f"链接不可达（404/410 不存在）：{url}"
            if resp.status_code == 403:
                return f"链接可能需权限访问（403 Forbidden）：{url}"
            # Other 4xx / 5xx
            return f"链接返回 HTTP {resp.status_code}：{url}"

    except httpx.TimeoutException:
        return f"链接超时（{timeout}s）：{url}"
    except httpx.ConnectError:
        return f"链接无法连接（DNS/网络不可达）：{url}"
    except httpx.SSLError:
        return f"链接 SSL 证书错误：{url}"
    except Exception as e:
        return f"链接检查异常（{type(e).__name__}）：{url}"


def validate_urls(data) -> list[ValidationError]:
    """Validate URLs extracted from the report JSON for reachability.

    Runs concurrent HEAD/GET checks via ThreadPoolExecutor. Skips domains
    listed in URL_SKIP_DOMAINS (e.g., gov.cn, npc.gov.cn).

    Args:
        data: Parsed report JSON.

    Returns:
        List of ValidationError objects for unreachable URLs.
    """
    urls = extract_urls_from_json(data)
    if not urls:
        return []

    # Filter out skip domains
    check_urls = []
    skipped = 0
    for url in urls:
        try:
            from urllib.parse import urlparse
            domain = urlparse(url).netloc.lower()
        except Exception:
            domain = ""
        if any(domain.endswith(sd) for sd in URL_SKIP_DOMAINS):
            skipped += 1
            continue
        check_urls.append(url)

    if not check_urls:
        logger.info("URL validation: %d URLs extracted, %d skipped (gov.cn/etc), 0 to check",
                    len(urls), skipped)
        return []

    logger.info("URL validation: checking %d URLs (%d extracted, %d skipped, max_concurrent=%d)...",
                len(check_urls), len(urls), skipped, URL_VALIDATION_MAX_CONCURRENT)

    # Concurrent check
    errors = []
    with ThreadPoolExecutor(max_workers=URL_VALIDATION_MAX_CONCURRENT) as executor:
        future_map = {executor.submit(_check_single_url, url): url for url in check_urls}
        for future in as_completed(future_map):
            result = future.result()
            if result:
                url = future_map[future]
                errors.append(ValidationError(
                    field="url",
                    message=result,
                ))

    logger.info("URL validation done: %d/%d URLs failed", len(errors), len(check_urls))
    return errors


# ─── Case Number Validation ──────────────────────────────────────────

# Chinese case number format: （年份）法院代字案件类型审判程序序号号
# Court codes can include digits: 京0108, 粤03, 沪0115
# Examples: （2024）京0108民初1234号  （2023）最高法民终456号  （2022）粤03民终7890号
_CASE_NUMBER_RE = re.compile(
    r'（\d{4}）'                          # （年份）
    r'[一-龥]{1,4}\d{0,6}'               # 法院代字（1-4汉字 + 可选数字，如京0108）
    r'(?:民|刑|行|赔|执|破|申|再|知|商|海|环|少|家)'  # 案件类型
    r'(?:初|终|再|监|抗|重|复|申|调|确|认|裁|命)'    # 审判程序
    r'\d+'                                 # 案件序号
    r'[号字]'                              # 号/字结尾
)

# Pattern to catch case numbers with placeholder characters
# Matches text starting with （年份）that contains X/x/×/?/?/…/某 before 号
_CASE_PLACEHOLDER_RE = re.compile(
    r'（\d{4}）'                          # （年份）
    r'[一-龥]{1,4}\d{0,6}'               # 法院代字
    r'(?:民|刑|行|赔|执|破|申|再|知|商|海|环|少|家)?'  # 案件类型（可选）
    r'(?:初|终|再|监|抗|重|复|申|调|确|认|裁|命)?'    # 审判程序（可选）
    r'[^号字]{0,20}'                      # 中间内容
    r'(?:[Xx×\?？…\.]{1,4}|某{1,2})'     # 占位符字符（必现）
    r'[^号字]{0,5}'                       # 其余内容
    r'[号字]?'                             # 可选的号/字
)


def scan_case_numbers(data, whitelist: Optional[set[str]] = None) -> list[ValidationError]:
    """Scan report JSON for fabricated or malformed case numbers.

    Three checks:
    1. Placeholder case numbers: contains X/x/×/?/?/…/某 placeholder characters
       → These are definitely fabricated. Error.
    2. Text that looks like it contains case references but doesn't match
       the standard Chinese case number format → Warning.
    3. Any text matching the standard Chinese case number format → checked
       against the whitelist. If the normalized case number is in the
       whitelist, it passes (came from web search results). If not in the
       whitelist (or whitelist is None/empty), it's flagged as an error
       since the LLM has no independent case database.

    Args:
        data: Parsed report JSON.
        whitelist: Set of normalized case numbers that are allowed
                   (extracted from web search results). None or empty
                   means no case numbers are allowed.

    Returns:
        List of ValidationError objects.
    """
    if whitelist is None:
        whitelist = set()

    errors = []

    def _collect_texts(obj, texts: list[tuple[str, str]]):
        """Recursively collect (value, key) string tuples from JSON."""
        if isinstance(obj, str):
            texts.append((obj, ""))
        elif isinstance(obj, dict):
            for k, v in obj.items():
                if isinstance(v, str):
                    texts.append((v, k))
                else:
                    _collect_texts(v, texts)
        elif isinstance(obj, list):
            for item in obj:
                _collect_texts(item, texts)

    texts = []
    _collect_texts(data, texts)

    for text, key_hint in texts:
        # Check 1: Placeholder case numbers
        placeholder_matches = _CASE_PLACEHOLDER_RE.findall(text)
        for match in placeholder_matches:
            snippet = match[:60] if len(match) > 60 else match
            errors.append(ValidationError(
                field="content",
                message=f"案号含占位符（疑似杜撰）：{snippet}（提示：你没有案例数据库，禁止编造案号。请使用一般性描述替代）",
            ))

        # Check 2: Loose scan for text sections that claim to list cases
        # but don't have properly formatted case numbers.
        # Only check in sections whose key suggests case content.
        if key_hint.lower() in ('content', '标题', 'rows', 'caption'):
            # Look for patterns like "（20XX）...号" that don't match the standard format
            # This catches fabricated case numbers with wrong format
            loose_pattern = re.findall(r'（\d{4}）[^）]{3,30}号', text)
            for candidate in loose_pattern:
                if not _CASE_NUMBER_RE.search(candidate):
                    snippet = candidate[:60] if len(candidate) > 60 else candidate
                    errors.append(ValidationError(
                        field="content",
                        message=f"案号格式异常（不符合中国案号标准格式）：{snippet}",
                    ))

        # Check 3: Flag case numbers NOT in the whitelist.
        # Case numbers found in web search results are allowed (they come
        # from verifiable news sources). All others are flagged as errors
        # since the LLM has no independent case database.
        case_number_matches = _CASE_NUMBER_RE.findall(text)
        for match in case_number_matches:
            normalized = normalize_case_number(match)
            if normalized in whitelist:
                # Case number was found in search results → allowed
                logger.info("Case number allowed (in whitelist): %s", match[:40])
                continue
            # Not in whitelist → flag as error
            snippet = match[:60] if len(match) > 60 else match
            if whitelist:
                msg = (f'案号未在搜索结果中找到：{snippet}'
                       f'（只能引用用户提示中网络检索结果提供的案号。'
                       f'如果搜索结果中没有相关案号，请使用一般性描述，如"在类似纠纷的司法实践中，法院通常认定…"）')
            else:
                msg = (f'禁止输出具体案号：{snippet}'
                       f'（网络搜索未找到可验证的案例。'
                       f'请删除案号，改用一般性描述，如"在类似纠纷的司法实践中，法院通常认定…"）')
            errors.append(ValidationError(
                field="content",
                message=msg,
            ))

    return errors


# ─── Structural Validation ────────────────────────────────────────────

# Expected h1 title keywords (in order, fuzzy matched)
EXPECTED_H1_KEYWORDS = [
    ["背景"],       # 一、背景介绍
    ["结论"],       # 二、评估结论
    ["分析"],       # 三、法律分析
    ["案例", "实践", "司法"],  # 四、行业实践与司法案例
    ["风险", "合规", "建议"],  # 五、风险评级与合规建议
    ["信息", "来源"],         # 六、信息来源
]


def validate_report_structure(data: dict) -> list[ValidationError]:
    """Validate the report JSON structure beyond what Pydantic catches.

    Checks:
    - title is present and non-empty
    - 6 h1 sections in approximately correct order
    - table sections have valid content structure
    - No empty body/heading content
    """
    errors = []

    # Check title
    title = data.get("title", "")
    if not title or len(title.strip()) < 4:
        errors.append(ValidationError(
            field="title",
            message="报告标题为空或过短（至少4个字符）",
        ))

    sections = data.get("sections", [])
    if not sections:
        errors.append(ValidationError(field="sections", message="sections 数组为空"))
        return errors

    # Collect h1 sections
    h1_indices = []
    for i, sec in enumerate(sections):
        if sec.get("type") == "h1":
            h1_indices.append(i)

    # Check h1 count
    if len(h1_indices) < 6:
        errors.append(ValidationError(
            field="sections",
            message=f"h1 章节数量不足：期望 6 个，实际 {len(h1_indices)} 个",
        ))
    elif len(h1_indices) > 6:
        errors.append(ValidationError(
            field="sections",
            message=f"h1 章节数量过多：期望 6 个，实际 {len(h1_indices)} 个",
        ))

    # Check h1 title keywords (fuzzy)
    for idx, h1_pos in enumerate(h1_indices[:6]):
        content = sections[h1_pos].get("content", "")
        if isinstance(content, str):
            keywords = EXPECTED_H1_KEYWORDS[idx] if idx < len(EXPECTED_H1_KEYWORDS) else []
            if keywords:
                matched = any(kw in content for kw in keywords)
                if not matched:
                    content_snip = content[:30]
                    errors.append(ValidationError(
                        field=f"sections[{h1_pos}].content",
                        message=f"h1 标题「{content_snip}…」可能不符合预期（期望包含关键词：{'/'.join(keywords)}）",
                        section_index=h1_pos,
                    ))

    # Check each section has non-empty content
    for i, sec in enumerate(sections):
        sec_type = sec.get("type", "")
        content = sec.get("content", "")

        if sec_type == "table":
            if not isinstance(content, dict):
                errors.append(ValidationError(
                    field=f"sections[{i}].content",
                    message="table 类型的 content 必须是对象",
                    section_index=i,
                ))
            else:
                if not content.get("headers"):
                    errors.append(ValidationError(
                        field=f"sections[{i}].content.headers",
                        message="表格缺少 headers",
                        section_index=i,
                    ))
                if not content.get("rows"):
                    errors.append(ValidationError(
                        field=f"sections[{i}].content.rows",
                        message="表格 rows 为空",
                        section_index=i,
                    ))
        elif sec_type in ("h1", "h2", "h3", "h4", "h5", "body", "source_item"):
            if not isinstance(content, str) or not content.strip():
                errors.append(ValidationError(
                    field=f"sections[{i}].content",
                    message=f"{sec_type} 类型的 content 为空",
                    section_index=i,
                ))
        elif sec_type:
            errors.append(ValidationError(
                field=f"sections[{i}].type",
                message=f"未知的 section 类型：{sec_type}",
                section_index=i,
            ))

    # Check 信息来源 section (last h1) doesn't use table type
    if len(h1_indices) >= 6:
        last_h1 = h1_indices[5]  # 六、信息来源 should be the 6th h1
        for i in range(last_h1 + 1, len(sections)):
            if sections[i].get("type") == "table":
                errors.append(ValidationError(
                    field=f"sections[{i}].type",
                    message="信息来源部分不得使用 table 类型，请使用 h2+h3+body 结构",
                    section_index=i,
                ))
                break

    return errors


def validate_report(raw_text: str, case_whitelist: Optional[set[str]] = None) -> tuple[Optional[dict], list[ValidationError]]:
    """Full validation pipeline: clean → parse → Pydantic → structural → content scan.

    Args:
        raw_text: Raw LLM response text.
        case_whitelist: Optional set of normalized case numbers allowed
                        (from web search results). Forwarded to scan_case_numbers().

    Returns:
        (parsed_data, errors) — parsed_data is None if parsing failed
    """
    errors = []

    # Step 1: Clean the response
    cleaned = clean_json_text(raw_text)

    # Step 2: Parse JSON
    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError as e:
        errors.append(ValidationError(
            field="json",
            message=f"JSON 解析错误：{e.msg}（位置：第{e.lineno}行，第{e.colno}列）",
        ))
        return None, errors

    # Step 3: Pydantic structural validation
    try:
        Report.model_validate(data)
    except Exception as e:
        # Extract key info from Pydantic error
        error_str = str(e)
        # Truncate very long Pydantic errors
        if len(error_str) > 500:
            error_str = error_str[:500] + "..."
        errors.append(ValidationError(
            field="report",
            message=f"结构校验失败：{error_str}",
        ))
        # Continue with further checks even if Pydantic fails
        # (we want to collect all issues for the retry prompt)

    # Step 4: Custom structural checks
    struct_errors = validate_report_structure(data)
    errors.extend(struct_errors)

    # Step 5: Content scan (forbidden patterns)
    scan_violations = scan_json_content(data)
    for v in scan_violations:
        errors.append(ValidationError(
            field="content",
            message=f"违禁内容：{v['rule']}（示例：{v['matches'][0][:50] if v['matches'] else ''}）",
        ))

    # Step 5.5: Case number scan (fabricated / placeholder detection)
    # Runs unconditionally — catches fabricated case numbers even if earlier
    # steps have errors (case number fabrication is a content quality issue,
    # not a structural one). Whitelist from web search results is passed
    # through to allow case numbers from verifiable sources.
    case_errors = scan_case_numbers(data, whitelist=case_whitelist)
    errors.extend(case_errors)

    # Step 6: URL reachability validation
    if URL_VALIDATION_ENABLED and not errors:
        # Only run URL checks if earlier steps passed (no point checking URLs
        # if the JSON is structurally broken — LLM should fix structure first)
        url_errors = validate_urls(data)
        errors.extend(url_errors)

    return data, errors


# ─── Main Generation Orchestration ────────────────────────────────────

def generate_report(request: GenerateRequest) -> tuple[Optional[dict], str, int, str]:
    """Generate a report by calling the LLM, validating, and retrying.

    Args:
        request: The generation request with provider, api_key, question.

    Returns:
        (report_data, json_path, retries_used, error_message)
        - report_data is the parsed JSON dict, or None on failure
        - json_path is the path to the saved JSON file
        - retries_used is how many retries were attempted
        - error_message is empty on success
    """
    provider = create_provider(
        provider_type=request.provider,
        api_key=request.api_key,
        model=request.model,
        base_url=request.base_url,
    )

    # ─── Step 0: Web Search (runs once before the retry loop) ─────────
    search_context = ""
    case_whitelist: set = set()
    if SEARCH_ENABLED:
        try:
            from search_service import search_for_cases, format_search_results_for_prompt
            search_results, case_whitelist = search_for_cases(
                request.question,
                max_results=SEARCH_MAX_RESULTS,
                timeout=SEARCH_TIMEOUT,
                max_queries=SEARCH_MAX_QUERIES,
                query_delay=SEARCH_QUERY_DELAY,
            )
            if search_results:
                search_context = format_search_results_for_prompt(search_results)
                logger.info(
                    "Search: %d results, %d case numbers whitelisted for question (len=%d)",
                    len(search_results), len(case_whitelist), len(request.question),
                )
            else:
                logger.info("Search returned no results; proceeding without case data")
        except Exception as e:
            logger.warning("Search failed (continuing without case data): %s", e)
            # Graceful degradation: proceed without search results

    # ─── Step 1: Build prompts ────────────────────────────────────────
    system_prompt = build_system_prompt()
    user_prompt = build_user_prompt(request.question, search_context)

    max_tokens = MAX_OUTPUT_TOKENS.get(request.provider, 16000)

    # Generate with retries
    current_user_prompt = user_prompt
    report_data = None
    errors = []
    retries = 0

    for attempt in range(MAX_RETRIES + 1):
        retries = attempt

        try:
            result = provider.generate(
                system_prompt=system_prompt,
                user_message=current_user_prompt,
                temperature=DEFAULT_TEMPERATURE,
                max_tokens=max_tokens,
            )
        except Exception as e:
            error_msg = f"LLM API 调用失败：{str(e)}"
            return None, "", retries, error_msg

        report_data, errors = validate_report(result.text, case_whitelist)

        if not errors:
            break  # Success!

        # Build retry prompt for next attempt
        if attempt < MAX_RETRIES:
            current_user_prompt = build_retry_prompt(errors, request.question, search_context)

    if errors:
        # Format errors for display
        error_details = []
        for e in errors[:10]:  # Max 10 errors in message
            error_details.append(f"• {e.field}: {e.message}")
        error_message = f"经过 {retries + 1} 次尝试后仍存在 {len(errors)} 个问题：\n" + "\n".join(error_details)
        return None, "", retries, error_message

    # Save successful JSON
    report_id = uuid.uuid4().hex[:12]
    json_path = WEBAPP_OUTPUT_DIR / f"report_{report_id}.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(report_data, f, ensure_ascii=False, indent=2)

    return report_data, str(json_path), retries, ""
