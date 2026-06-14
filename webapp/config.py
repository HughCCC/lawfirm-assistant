"""Configuration constants for the legal report web application."""

from pathlib import Path

# Project root
PROJECT_ROOT = Path(__file__).parent.parent.resolve()

# Existing scripts and output
LAW_RESEARCH_DIR = PROJECT_ROOT / ".law-research"
SCRIPTS_DIR = LAW_RESEARCH_DIR / "scripts"
LAW_OUTPUT_DIR = LAW_RESEARCH_DIR / "output"

# Web app output
WEBAPP_OUTPUT_DIR = Path(__file__).parent / "output"
WEBAPP_OUTPUT_DIR.mkdir(exist_ok=True)

# Retry configuration
MAX_RETRIES = 2

# LLM temperature — low for consistency
DEFAULT_TEMPERATURE = 0.3

# Max tokens per provider
MAX_OUTPUT_TOKENS = {
    "openai": 16000,
    "anthropic": 32000,
    "deepseek": 8000,
    "custom": 16000,
}

# Default models per provider
DEFAULT_MODELS = {
    "openai": "gpt-4o",
    "anthropic": "claude-sonnet-4-20250514",
    "deepseek": "deepseek-chat",
    "custom": "",
}

# DeepSeek base URL
DEEPSEEK_BASE_URL = "https://api.deepseek.com/v1"

# Cover template
COVER_TEMPLATE_PATH = PROJECT_ROOT / "封面.docx"

# URL validation
URL_VALIDATION_ENABLED = True           # Master switch for URL reachability checks
URL_VALIDATION_TIMEOUT = 10             # Seconds per URL HEAD/GET request
URL_VALIDATION_MAX_CONCURRENT = 5       # Max concurrent URL checks
URL_VALIDATION_MAX_URLS = 50            # Cap total URLs checked per report
URL_SKIP_DOMAINS = [                    # Domains assumed stable (skip HEAD check)
    "gov.cn",
    "npc.gov.cn",
]

# Web search for legal case discovery
SEARCH_ENABLED = True            # Master switch; set False to disable all web search
SEARCH_MAX_RESULTS = 8           # Max total search results across all queries
SEARCH_TIMEOUT = 15              # Seconds per HTTP request
SEARCH_MAX_QUERIES = 2           # Max distinct search queries to send
SEARCH_QUERY_DELAY = 2.0         # Seconds delay between queries (rate limiting)

# File cleanup: keep last N reports
MAX_OUTPUT_FILES = 50
