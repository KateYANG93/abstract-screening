"""
Conservative abstract screening for TMX/LMX and team-level performance meta-analysis.

This script is for conservative abstract screening only. It should not be treated as
final inclusion/exclusion. Ambiguous records are retained for manual/full-text review
to avoid false negative exclusion.

Dependencies:
    pip install pandas openpyxl
"""

from __future__ import annotations

import argparse
import re
import sys
import unicodedata
from pathlib import Path

import pandas as pd
from openpyxl.utils import get_column_letter

# =============================================================================
# Configuration
# =============================================================================

DEFAULT_INPUT = Path(r"C:\Lingnan University\Mini meta\dedup\dedup_meta_analysis_output.xlsx")
DEFAULT_SHEET = "unique_records"
DEFAULT_OUTPUT = Path(__file__).resolve().parent / "abstract_screening_output.xlsx"

MIN_ABSTRACT_CHARS = 80
MIN_TITLE_CHARS = 10

SCREENING_TEXT_FIELDS = [
    "title",
    "abstract",
    "academic journal",
    "publication_type",
    "document_type",
    "source_type",
    "genre",
    "keywords",
]

DECISION_ORDER = {
    "include_likely": 0,
    "maybe_manual_review": 1,
    "exclude_likely": 2,
}

CONFIDENCE_ORDER = {"high": 0, "medium": 1, "low": 2}

UNICODE_DASHES = "\u2010\u2011\u2012\u2013\u2014\u2015\u2212"
UNICODE_APOSTROPHES = "'\u2018\u2019\u201b\u2032"
UNICODE_CURLY_QUOTES = "\u201c\u201d\u201e\u201f"

# =============================================================================
# Keyword dictionaries
# =============================================================================

DIRECT_TMX_LMX_PHRASES = [
    "tmx",
    "team member exchange",
    "team-member exchange",
    "teammember exchange",
    "team member exchange quality",
    "team-member exchange quality",
    "tmxd",
    "team member exchange differentiation",
    "team-member exchange differentiation",
    "lmx",
    "leader member exchange",
    "leader-member exchange",
    "leadermember exchange",
    "leader member exchange quality",
    "leader-member exchange quality",
    "lmxd",
    "leader member exchange differentiation",
    "leader-member exchange differentiation",
    "leader member exchange differentiation",
]

EXCHANGE_EQUIVALENT_PHRASES = [
    "exchange quality between leader and member",
    "exchange quality between leaders and members",
    "quality of leader member relationship",
    "quality of leader-member relationship",
    "quality of the leader member relationship",
    "exchange relationship between supervisors and subordinates",
    "exchange relationship between supervisor and subordinate",
    "exchange quality among team members",
    "team member exchange quality",
]

TMX_LMX_PHRASES = list(dict.fromkeys(DIRECT_TMX_LMX_PHRASES + EXCHANGE_EQUIVALENT_PHRASES))

TEAM_CONTEXT_PHRASES = [
    "team",
    "teams",
    "work group",
    "workgroup",
    "work groups",
    "project team",
    "r d team",
    "sales team",
    "service team",
    "group",
    "groups",
    "unit",
    "units",
    "branch",
    "branches",
    "store",
    "stores",
    "outlet",
    "outlets",
    "restaurant",
    "restaurants",
    "department",
    "departments",
    "hotel",
    "hotels",
    "hospital unit",
    "supervisor subordinate",
    "supervisor-subordinate",
    "employee",
    "employees",
    "organization",
    "organizations",
    "workplace",
    "firm employees",
    "organizational",
    "work context",
]

ELIGIBLE_TEAM_PERFORMANCE_PHRASES = [
    "team performance",
    "group performance",
    "team effectiveness",
    "group effectiveness",
    "unit performance",
    "unit effectiveness",
    "branch performance",
    "branch effectiveness",
    "store performance",
    "store effectiveness",
    "restaurant performance",
    "restaurant effectiveness",
    "outlet performance",
    "outlet effectiveness",
    "service performance",
    "team productivity",
    "group productivity",
    "unit productivity",
    "store sales",
    "branch sales",
    "team output",
    "unit output",
    "objective team performance",
    "objective sales",
    "objective output",
    "objective sales output",
    "supervisor rated team performance",
    "supervisor-rated team performance",
    "leader rated team performance",
    "leader-rated team performance",
    "member rated team performance",
    "member-rated team performance",
    "member rated group performance",
    "member-rated group performance",
    "customer rated team performance",
    "customer-rated team performance",
    "customer rated group performance",
    "customer-rated group performance",
]

PERFORMANCE_EFFECTIVENESS_OUTCOME_PHRASES = [
    "performance",
    "performances",
    "effectiveness",
    "effective performance",
    "productivity",
    "output",
    "outputs",
    "sales",
    "objective sales",
    "objective output",
    "service performance",
    "service effectiveness",
    "task performance",
    "rated performance",
    "goal attainment",
    "goal achievement",
    "team effectiveness",
    "group effectiveness",
    "unit effectiveness",
    "branch effectiveness",
    "store effectiveness",
    "restaurant effectiveness",
    "outlet effectiveness",
]

INNOVATION_CREATIVITY_PHRASES = [
    "innovation",
    "innovative performance",
    "team innovation",
    "team innovative performance",
    "creativity",
    "team creativity",
    "creative performance",
    "innovation performance",
    "innovative behavior",
    "innovative behaviour",
    "innovative work behavior",
    "innovative work behaviour",
    "employee creativity",
    "employee innovative behavior",
    "employee innovative behaviour",
]

INDIVIDUAL_NON_ELIGIBLE_PERFORMANCE_PHRASES = [
    "employee performance",
    "individual performance",
    "job performance",
    "leader performance",
    "managerial performance",
    "manager performance",
    "female leader performance",
    "faculty performance",
    "organizational performance",
    "firm performance",
    "company performance",
    "university performance",
]

# Legacy alias used in audit display
PERFORMANCE_PHRASES = ELIGIBLE_TEAM_PERFORMANCE_PHRASES

TEAM_LEVEL_PHRASES = [
    "team level",
    "group level",
    "unit level",
    "store level",
    "branch level",
    "aggregated",
    "aggregation",
    "team mean",
    "group mean",
    "unit mean",
    "shared",
    "collective",
    "within team agreement",
    "within-team agreement",
    "rwg",
    "icc",
    "icc 1",
    "icc 2",
    "icc1",
    "icc2",
    "multilevel",
    "team average",
    "group average",
    "aggregated member rated",
    "aggregated member-rated",
    "team level lmx",
    "team level tmx",
    "group level lmx",
    "group level tmx",
    "between team",
    "between-team",
    "cross level",
    "cross-level",
]

QUANTITATIVE_PHRASES = [
    "empirical",
    "quantitative",
    "survey",
    "data",
    "sample",
    "respondents",
    "regression",
    "correlation",
    "multilevel",
    "hierarchical linear modeling",
    "hierarchical linear model",
    "structural equation modeling",
    "structural equation model",
    "sem",
    "hlm",
    "analysis",
    "model",
    "tested",
    "hypothesis",
    "hypotheses",
    "results",
    "findings",
    "examined",
    "study",
    "studies",
]

WRONG_CONTEXT_PHRASES = [
    "student sample",
    "student samples",
    "undergraduate students",
    "undergraduate student",
    "mba students",
    "mba student",
    "classroom",
    "classroom exercise",
    "classroom only",
    "laboratory team",
    "laboratory teams",
    "lab team",
    "lab teams",
    "experimental lab",
    "simulation team",
    "simulation teams",
    "simulated team",
    "simulated teams",
    "student team",
    "student teams",
    "undergraduate team",
    "undergraduate teams",
    "mba team",
    "mba teams",
]

NON_PERFORMANCE_OUTCOME_PHRASES = [
    "job satisfaction",
    "employee satisfaction",
    "satisfaction",
    "organizational citizenship",
    "organizational citizenship behavior",
    "organizational citizenship behaviours",
    "ocb",
    "ocbs",
    "trust",
    "cohesion",
    "commitment",
    "thriving",
    "expedient behavior",
    "expedient behaviour",
    "creativity",
    "voice",
    "citizenship",
    "psychological safety",
    "well being",
    "wellbeing",
    "engagement",
    "burnout",
    "turnover intention",
    "intention to leave",
]

GENERAL_LEADERSHIP_CONSTRUCT_PHRASES = [
    "servant leadership",
    "transformational leadership",
    "transactional leadership",
    "ethical leadership",
    "authentic leadership",
    "inclusive leadership",
    "empowering leadership",
    "paternalistic leadership",
    "shared leadership",
    "charismatic leadership",
    "leadership style",
    "leadership styles",
    "female leadership",
    "leader behavior",
    "leader behaviour",
    "leadership behavior",
    "leadership behaviour",
    "leadership effectiveness",
    "leadership practices",
    "leadership development",
    "general leadership",
    "leadership quality",
]

OTHER_NON_EXCHANGE_CONSTRUCT_PHRASES = [
    "teamwork",
    "team climate",
    "group climate",
    "organizational climate",
    "organizational support",
    "perceived organizational support",
    "team support",
    "social support",
    "psychological safety",
    "team cohesion",
    "group cohesion",
]

OTHER_LEADERSHIP_PHRASES = GENERAL_LEADERSHIP_CONSTRUCT_PHRASES + OTHER_NON_EXCHANGE_CONSTRUCT_PHRASES

INDIVIDUAL_LEVEL_FOCUS_PHRASES = [
    "individual performance",
    "employee performance",
    "job performance",
    "leader performance",
    "manager performance",
    "managerial performance",
    "female leader performance",
    "faculty performance",
    "innovative performance of female leaders",
    "employee innovative performance",
    "individual innovative performance",
    "performance of individuals",
    "individual leaders",
    "female leaders",
    "faculty members",
    "faculty member",
    "individual employees",
    "individual employee",
    "heads of department",
    "head of department",
    "deans",
    "coordinators",
    "directors",
    "individual level",
    "employee level",
    "personal performance",
]

WEAK_REVIEW_BACKGROUND_PHRASES = [
    "literature review",
    "research review",
]

STRONG_REVIEW_META_PHRASES = [
    "systematic review",
    "narrative review",
    "scoping review",
    "review article",
    "integrative review",
    "bibliometric review",
    "meta analysis",
    "meta-analysis",
    "meta analyses",
    "meta-analyses",
    "meta analytic",
    "meta-analytic",
    "meta analytical",
    "meta-analytical",
    "meta analytical correlation matrix",
    "meta-analytical correlation matrix",
    "meta analytic correlation matrix",
]

OTHER_INELIGIBLE_PUBLICATION_PHRASES = [
    "conceptual paper",
    "conceptual article",
    "theoretical paper",
    "theoretical article",
    "qualitative only",
    "qualitative study only",
    "editorial",
    "commentary",
    "practitioner article",
    "practitioner paper",
    "trade magazine",
    "magazine article",
    "newspaper",
    "book review",
    "conference abstract only",
    "meeting abstract only",
]

EMPIRICAL_STUDY_EVIDENCE_PHRASES = [
    "empirical study",
    "empirical research",
    "survey",
    "survey data",
    "questionnaire",
    "data collected",
    "two wave data",
    "two-wave data",
    "sample",
    "sample of",
    "respondents",
    "employees",
    "subordinates",
    "supervisors",
    "teams",
    "companies",
    "firms",
    "regression",
    "hierarchical multiple regression",
    "structural equation modeling",
    "structural equation model",
    "structural equation",
    "sem",
    "hlm",
    "tested the model",
    "tested the hypothesis",
    "tested the hypotheses",
    "field study",
    "cross sectional",
    "cross-sectional",
    "self reported questionnaire",
    "self-reported questionnaire",
]

INELIGIBLE_PUBLICATION_PHRASES = (
    OTHER_INELIGIBLE_PUBLICATION_PHRASES
    + WEAK_REVIEW_BACKGROUND_PHRASES
    + STRONG_REVIEW_META_PHRASES
)

CONFERENCE_METADATA_PHRASES = [
    "conference paper",
    "conference proceeding",
    "conference proceedings",
    "conference abstract",
    "meeting abstract",
    "inproceedings",
    "proceedings",
]

INDIVIDUAL_LEVEL_ONLY_PHRASES = [
    "individual level only",
    "purely individual level",
    "only individual level",
    "individual level analysis only",
    "firm level only",
    "only firm level",
    "organization level only",
    "only organizational level",
    "no team level",
    "not team level",
]

QUALITATIVE_ONLY_PHRASES = [
    "qualitative only",
    "purely qualitative",
    "conceptual only",
    "purely conceptual",
    "review only",
    "theoretical only",
]

ENGLISH_FUNCTION_WORDS = frozenset({
    "the", "and", "of", "in", "to", "for", "with", "on", "by", "from", "as",
    "is", "are", "was", "were", "be", "this", "that", "which", "among", "between",
    "a", "an", "or", "at", "into", "through", "during", "including", "against",
})

NON_ENGLISH_FUNCTION_WORDS = frozenset({
    "der", "die", "das", "den", "dem", "des", "ein", "eine", "einer", "eines", "einem",
    "und", "von", "mit", "nach", "uber", "zum", "zur",
    "ist", "sind", "wird", "werden", "nicht", "auch", "diese", "dieser", "dieses",
    "durch", "wurde", "wurden", "kann", "noch", "schon", "sowie",
    "les", "une", "dans", "pour", "avec", "sont", "aux", "cette", "ces", "mais",
    "chez", "dont", "leur", "leurs", "etait", "etaient", "entre", "etre",
    "como", "pero", "sus", "han", "ser", "sobre", "hay", "fueron", "segun",
    "nao", "sao", "tem", "pelo", "pela", "seu", "sua",
    "het", "zijn", "wordt", "naar",
    # Turkish (ASCII-folded forms common in corrupted exports)
    "uzerine", "uzerindeki", "arastirma", "arastirmasi", "degisken", "etkisi",
    "etkilesimi", "bir", "icin", "calismasi", "yeni", "bagimsiz",
})

SCREENING_COLUMNS = [
    "screening_text",
    "screen_decision",
    "screen_confidence",
    "primary_reason",
    "inclusion_signals",
    "exclusion_signals",
    "direct_tmx_lmx_title_abstract_signal",
    "performance_signal",
    "team_context_signal",
    "team_level_signal",
    "quantitative_signal",
    "wrong_context_signal",
    "wrong_publication_type_signal",
]


# =============================================================================
# Text utilities
# =============================================================================

def is_blank(value) -> bool:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return True
    return str(value).strip() == "" or str(value).strip().lower() in {"nan", "none"}


def ascii_fold(text: str) -> str:
    normalized = unicodedata.normalize("NFKD", text)
    return normalized.encode("ascii", "ignore").decode("ascii")


def preprocess_unicode_text(text: str) -> str:
    result = text
    for char in UNICODE_DASHES:
        result = result.replace(char, " ")
    for char in UNICODE_APOSTROPHES:
        result = result.replace(char, "")
    for char in UNICODE_CURLY_QUOTES:
        result = result.replace(char, " ")
    return result


def normalize_text(text) -> str:
    """Lowercase, remove punctuation, normalize Unicode, collapse spaces."""
    if is_blank(text):
        return ""
    value = str(text)
    value = re.sub(r"<[^>]+>", " ", value)
    value = value.replace("{", "").replace("}", "")
    value = preprocess_unicode_text(value)
    value = ascii_fold(value)
    value = value.lower()
    value = value.replace("&", " and ")
    value = re.sub(r"[^a-z0-9\s]", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def join_screening_text(row: pd.Series) -> str:
    parts: list[str] = []
    for field in SCREENING_TEXT_FIELDS:
        if field in row.index and not is_blank(row.get(field)):
            parts.append(str(row.get(field)))
    return " ".join(parts)


def match_phrases(text: str, phrases: list[str]) -> list[str]:
    """Return matched phrases (longest first to avoid duplicate sub-matches)."""
    if not text:
        return []
    matched: list[str] = []
    for phrase in sorted(phrases, key=len, reverse=True):
        if phrase in text and phrase not in matched:
            # Avoid counting a phrase if a longer overlapping phrase already matched
            if not any(phrase in existing for existing in matched):
                matched.append(phrase)
    return matched


def has_any_phrase(text: str, phrases: list[str]) -> bool:
    return bool(match_phrases(text, phrases))


def format_signal_list(signals: list[str]) -> str:
    return "; ".join(sorted(set(signals))) if signals else ""


def classify_performance_signals(text: str) -> dict:
    """Split performance mentions into eligible team/unit outcomes vs individual-level outcomes."""
    eligible = match_phrases(text, ELIGIBLE_TEAM_PERFORMANCE_PHRASES)
    individual = match_phrases(text, INDIVIDUAL_NON_ELIGIBLE_PERFORMANCE_PHRASES)
    innovation_creativity = match_phrases(text, INNOVATION_CREATIVITY_PHRASES)

    eligible = list(dict.fromkeys(eligible))
    individual = list(dict.fromkeys(individual))
    innovation_creativity = list(dict.fromkeys(innovation_creativity))

    innovation_creativity_only = bool(innovation_creativity) and not bool(eligible)

    return {
        "eligible_performance": eligible,
        "individual_performance": individual,
        "innovation_creativity": innovation_creativity,
        "has_eligible_performance": bool(eligible),
        "has_individual_performance": bool(individual),
        "innovation_creativity_only": innovation_creativity_only,
    }


def join_title_abstract_text(row: pd.Series) -> str:
    parts: list[str] = []
    for field in ("title", "abstract"):
        if field in row.index and not is_blank(row.get(field)):
            parts.append(str(row.get(field)))
    return " ".join(parts)


def get_publication_type_metadata_text(row: pd.Series) -> str:
    parts: list[str] = []
    for field in ("publication_type", "document_type", "source_type", "genre"):
        if field in row.index and not is_blank(row.get(field)):
            parts.append(str(row.get(field)))
    return normalize_text(" ".join(parts))


def get_journal_text(row: pd.Series) -> str:
    if "academic journal" not in row.index or is_blank(row.get("academic journal")):
        return ""
    return normalize_text(str(row.get("academic journal")))


def get_title_abstract_text(row: pd.Series) -> str:
    return normalize_text(join_title_abstract_text(row))


def has_empirical_study_evidence(text: str) -> bool:
    """Return True when text contains clear empirical-study indicators."""
    normalized = normalize_text(text)
    if not normalized:
        return False
    if match_phrases(normalized, EMPIRICAL_STUDY_EVIDENCE_PHRASES):
        return True
    if re.search(r"\bn\s*=\s*\d+", normalized):
        return True
    if re.search(
        r"\b\d+\s+(employees|subordinates|supervisors|respondents|participants|"
        r"teams|companies|firms|organizations|organisations)\b",
        normalized,
    ):
        return True
    return False


def _append_unique_matches(target: list[str], phrases: list[str]) -> None:
    for phrase in phrases:
        if phrase not in target:
            target.append(phrase)


def detect_ineligible_publication_type(row: pd.Series) -> list[str]:
    """
    Context-aware ineligible publication detection.

    Journal names are not treated as article-type evidence. Weak review-background
    phrases are ignored when empirical study evidence is present in title/abstract.
    """
    title_abstract_text = get_title_abstract_text(row)
    title_text = normalize_text(row.get("title")) if not is_blank(row.get("title")) else ""
    abstract_text = normalize_text(row.get("abstract")) if not is_blank(row.get("abstract")) else ""
    metadata_text = get_publication_type_metadata_text(row)
    has_empirical = has_empirical_study_evidence(title_abstract_text)

    matches: list[str] = []

    for text in (metadata_text, title_abstract_text):
        if text:
            _append_unique_matches(matches, match_phrases(text, OTHER_INELIGIBLE_PUBLICATION_PHRASES))

    if metadata_text:
        _append_unique_matches(matches, match_phrases(metadata_text, STRONG_REVIEW_META_PHRASES))

    if title_text:
        _append_unique_matches(matches, match_phrases(title_text, STRONG_REVIEW_META_PHRASES))

    if abstract_text:
        abstract_strong = match_phrases(abstract_text, STRONG_REVIEW_META_PHRASES)
        title_strong = set(match_phrases(title_text, STRONG_REVIEW_META_PHRASES)) if title_text else set()
        metadata_strong = set(match_phrases(metadata_text, STRONG_REVIEW_META_PHRASES)) if metadata_text else set()
        for phrase in abstract_strong:
            if phrase in title_strong or phrase in metadata_strong:
                continue
            if has_empirical:
                continue
            _append_unique_matches(matches, [phrase])

    weak_matches: list[str] = []
    if title_text:
        _append_unique_matches(weak_matches, match_phrases(title_text, WEAK_REVIEW_BACKGROUND_PHRASES))
    if abstract_text:
        _append_unique_matches(weak_matches, match_phrases(abstract_text, WEAK_REVIEW_BACKGROUND_PHRASES))
    if weak_matches and not has_empirical:
        _append_unique_matches(matches, weak_matches)

    return matches


def has_direct_tmx_lmx_signal(title, abstract) -> bool:
    """
    Return True only when title or abstract contains a direct TMX/LMX-family term
    or a narrowly defined exchange-quality equivalent phrase.
    """
    title_abstract_text = normalize_text(
        " ".join(part for part in (title, abstract) if not is_blank(part))
    )
    if not title_abstract_text:
        return False
    if match_phrases(title_abstract_text, DIRECT_TMX_LMX_PHRASES):
        return True
    return bool(match_phrases(title_abstract_text, EXCHANGE_EQUIVALENT_PHRASES))


# =============================================================================
# Language checks
# =============================================================================

def looks_english_text(text: str, *, min_alpha_chars: int = 12) -> bool:
    """Conservative English detection; ambiguous text is treated as English."""
    if is_blank(text):
        return True

    raw_folded = ascii_fold(preprocess_unicode_text(str(text)))
    alpha_chars = [char for char in raw_folded if char.isalpha()]
    if not alpha_chars:
        return True

    ascii_alpha = sum(1 for char in alpha_chars if char.isascii())
    ascii_ratio = ascii_alpha / len(alpha_chars)
    if ascii_ratio < 0.75:
        return False

    normalized = normalize_text(text)
    tokens = normalized.split()
    if not tokens:
        return True

    english_func = sum(1 for token in tokens if token in ENGLISH_FUNCTION_WORDS)
    non_english_func = sum(1 for token in tokens if token in NON_ENGLISH_FUNCTION_WORDS)
    non_english_share = non_english_func / len(tokens)

    if non_english_func >= 3 and english_func < non_english_func:
        return False
    if non_english_func >= 2 and english_func <= 1:
        return False
    if non_english_func >= 1 and english_func == 0 and non_english_share >= 0.12:
        return False

    if english_func >= 1:
        return True
    if non_english_func == 0 and ascii_ratio >= 0.92 and len(alpha_chars) >= 5:
        return True
    if ascii_ratio >= 0.92 and len(alpha_chars) >= min_alpha_chars:
        return True
    if ascii_ratio >= 0.85 and len(tokens) >= 3:
        return True
    return ascii_alpha >= min_alpha_chars and ascii_ratio >= 0.80


def is_clearly_non_english_title(row: pd.Series) -> bool:
    """Return True only when the title is clearly non-English."""
    title = str(row.get("title") or "")

    if is_blank(title):
        return False

    if "\ufffd" in title:
        normalized = normalize_text(title)
        non_english_tokens = sum(1 for token in normalized.split() if token in NON_ENGLISH_FUNCTION_WORDS)
        english_tokens = sum(1 for token in normalized.split() if token in ENGLISH_FUNCTION_WORDS)
        if non_english_tokens >= 2 and non_english_tokens > english_tokens:
            return True

    alpha = [char for char in title if char.isalpha()]
    if alpha:
        non_ascii = sum(1 for char in alpha if not char.isascii())
        if non_ascii / len(alpha) >= 0.20:
            return True

    return not looks_english_text(title, min_alpha_chars=8)


# =============================================================================
# Signal extraction
# =============================================================================

def extract_signals(row: pd.Series, screening_text: str) -> dict:
    title_text = normalize_text(row.get("title"))
    abstract_text = normalize_text(row.get("abstract"))
    title_abstract_text = get_title_abstract_text(row)
    metadata_text = get_publication_type_metadata_text(row)

    direct_tmx_lmx = match_phrases(title_abstract_text, DIRECT_TMX_LMX_PHRASES)
    direct_exchange_equiv = match_phrases(title_abstract_text, EXCHANGE_EQUIVALENT_PHRASES)
    has_direct_tmx = bool(direct_tmx_lmx or direct_exchange_equiv)

    team_context = match_phrases(screening_text, TEAM_CONTEXT_PHRASES)
    performance_info = classify_performance_signals(screening_text)
    eligible_performance = performance_info["eligible_performance"]
    individual_performance = performance_info["individual_performance"]
    innovation_creativity = performance_info["innovation_creativity"]
    team_level = match_phrases(screening_text, TEAM_LEVEL_PHRASES)
    quantitative = match_phrases(screening_text, QUANTITATIVE_PHRASES)
    wrong_context = match_phrases(screening_text, WRONG_CONTEXT_PHRASES)
    non_performance = match_phrases(screening_text, NON_PERFORMANCE_OUTCOME_PHRASES)
    outcome_related = match_phrases(screening_text, PERFORMANCE_EFFECTIVENESS_OUTCOME_PHRASES)
    other_leadership = match_phrases(screening_text, GENERAL_LEADERSHIP_CONSTRUCT_PHRASES)
    other_non_exchange = match_phrases(screening_text, OTHER_NON_EXCHANGE_CONSTRUCT_PHRASES)
    individual_level_focus = match_phrases(screening_text, INDIVIDUAL_LEVEL_FOCUS_PHRASES)
    ineligible_pub = detect_ineligible_publication_type(row)
    conference_meta = match_phrases(metadata_text, CONFERENCE_METADATA_PHRASES)
    individual_only = match_phrases(screening_text, INDIVIDUAL_LEVEL_ONLY_PHRASES)
    qualitative_only = match_phrases(screening_text, QUALITATIVE_ONLY_PHRASES)

    title_len = len(title_text)
    abstract_len = len(abstract_text)
    insufficient_text = (
        title_len < MIN_TITLE_CHARS
        or (is_blank(row.get("abstract")) and title_len < 25)
        or (not is_blank(row.get("abstract")) and abstract_len < MIN_ABSTRACT_CHARS)
    )

    performance_only_non_perf = bool(non_performance) and not performance_info["has_eligible_performance"]
    has_outcome_related = bool(
        outcome_related
        or eligible_performance
        or individual_performance
        or innovation_creativity
    )

    return {
        "direct_tmx_lmx": direct_tmx_lmx + direct_exchange_equiv,
        "has_direct_tmx_lmx": has_direct_tmx,
        "team_context": team_context,
        "eligible_performance": eligible_performance,
        "individual_performance": individual_performance,
        "innovation_creativity": innovation_creativity,
        "innovation_creativity_only": performance_info["innovation_creativity_only"],
        "performance": eligible_performance,
        "team_level": team_level,
        "quantitative": quantitative,
        "wrong_context": wrong_context,
        "non_performance": non_performance,
        "outcome_related": outcome_related,
        "has_outcome_related": has_outcome_related,
        "other_leadership": other_leadership,
        "other_non_exchange": other_non_exchange,
        "individual_level_focus": individual_level_focus,
        "ineligible_pub": ineligible_pub,
        "conference_meta": conference_meta,
        "individual_only": individual_only,
        "qualitative_only": qualitative_only,
        "insufficient_text": insufficient_text,
        "performance_only_non_perf": performance_only_non_perf,
        "has_eligible_performance": performance_info["has_eligible_performance"],
        "has_individual_performance": performance_info["has_individual_performance"],
        "title_text": title_text,
        "abstract_text": abstract_text,
        "metadata_text": metadata_text,
        "title_abstract_text": title_abstract_text,
    }


# =============================================================================
# Classification
# =============================================================================

def classify_record(row: pd.Series) -> dict:
    screening_text = join_screening_text(row)
    normalized_text = normalize_text(screening_text)
    signals = extract_signals(row, normalized_text)

    inclusion_signals: list[str] = []
    exclusion_signals: list[str] = []

    if signals["direct_tmx_lmx"]:
        inclusion_signals.append("direct_tmx_lmx:" + ",".join(signals["direct_tmx_lmx"][:5]))
    if signals["team_context"]:
        inclusion_signals.append("team_context:" + ",".join(signals["team_context"][:5]))
    if signals["eligible_performance"]:
        inclusion_signals.append(
            "performance:" + ",".join(signals["eligible_performance"][:5])
        )
    if signals["innovation_creativity"]:
        exclusion_signals.append(
            "innovation_creativity:" + ",".join(signals["innovation_creativity"][:5])
        )
    if signals["team_level"]:
        inclusion_signals.append("team_level:" + ",".join(signals["team_level"][:5]))
    if signals["quantitative"]:
        inclusion_signals.append("quantitative:" + ",".join(signals["quantitative"][:5]))

    if signals["wrong_context"]:
        exclusion_signals.append("wrong_context:" + ",".join(signals["wrong_context"][:5]))
    if signals["ineligible_pub"]:
        exclusion_signals.append("ineligible_pub:" + ",".join(signals["ineligible_pub"][:5]))
    if signals["conference_meta"]:
        exclusion_signals.append("conference_metadata:" + ",".join(signals["conference_meta"][:5]))
    if signals["individual_only"]:
        exclusion_signals.append("individual_level_only:" + ",".join(signals["individual_only"][:5]))
    if signals["qualitative_only"]:
        exclusion_signals.append("qualitative_only:" + ",".join(signals["qualitative_only"][:5]))

    has_direct_tmx = signals["has_direct_tmx_lmx"]
    has_team_context = bool(signals["team_context"])
    has_eligible_performance = signals["has_eligible_performance"]
    has_innovation_creativity_only = signals["innovation_creativity_only"]
    has_team_level = bool(signals["team_level"])
    has_quantitative = bool(signals["quantitative"])
    has_wrong_context = bool(signals["wrong_context"])
    has_ineligible_pub = bool(signals["ineligible_pub"]) or bool(signals["conference_meta"])
    has_individual_only = bool(signals["individual_only"])
    has_qualitative_only = bool(signals["qualitative_only"])

    if is_clearly_non_english_title(row):
        return _build_result(
            screening_text, "exclude_likely", "high", "non_english",
            inclusion_signals, exclusion_signals, signals,
        )

    if signals["insufficient_text"]:
        return _build_result(
            screening_text, "maybe_manual_review", "low", "insufficient_abstract_information",
            inclusion_signals, exclusion_signals, signals,
        )

    if has_qualitative_only and not has_quantitative:
        return _build_result(
            screening_text, "exclude_likely", "high", "ineligible_publication_type",
            inclusion_signals, exclusion_signals, signals,
        )
    if has_ineligible_pub:
        return _build_result(
            screening_text, "exclude_likely", "high", "ineligible_publication_type",
            inclusion_signals, exclusion_signals, signals,
        )

    if not has_direct_tmx:
        return _build_result(
            screening_text, "exclude_likely", "high", "no_direct_tmx_lmx_in_title_or_abstract",
            inclusion_signals, exclusion_signals, signals,
        )

    if has_innovation_creativity_only:
        return _build_result(
            screening_text, "exclude_likely", "high", "no_eligible_performance_effectiveness_outcome",
            inclusion_signals, exclusion_signals, signals,
        )

    if has_individual_only and not has_eligible_performance:
        return _build_result(
            screening_text, "exclude_likely", "high", "wrong_level_individual_only",
            inclusion_signals, exclusion_signals, signals,
        )

    if signals["performance_only_non_perf"] and not signals["has_outcome_related"]:
        return _build_result(
            screening_text, "exclude_likely", "high", "no_eligible_performance_effectiveness_outcome",
            inclusion_signals, exclusion_signals, signals,
        )

    if signals["has_individual_performance"] and not has_eligible_performance:
        return _build_result(
            screening_text, "exclude_likely", "high", "no_eligible_performance_effectiveness_outcome",
            inclusion_signals, exclusion_signals, signals,
        )

    if has_wrong_context and not has_team_context:
        return _build_result(
            screening_text, "exclude_likely", "high", "wrong_context_student_or_lab",
            inclusion_signals, exclusion_signals, signals,
        )
    if has_wrong_context and has_team_context:
        return _build_result(
            screening_text, "maybe_manual_review", "medium", "mixed_student_and_organizational_context",
            inclusion_signals, exclusion_signals, signals,
        )

    if has_eligible_performance:
        if has_team_level:
            confidence = "high" if has_quantitative else "medium"
            return _build_result(
                screening_text, "include_likely", confidence,
                "eligible_tmx_lmx_performance_team_context_and_level",
                inclusion_signals, exclusion_signals, signals,
            )
        return _build_result(
            screening_text, "maybe_manual_review", "medium",
            "eligible_construct_and_outcome_but_level_unclear",
            inclusion_signals, exclusion_signals, signals,
        )

    if not signals["has_outcome_related"]:
        return _build_result(
            screening_text,
            "exclude_likely",
            "high",
            "no_eligible_performance_effectiveness_outcome",
            inclusion_signals,
            exclusion_signals,
            signals,
        )

    return _build_result(
        screening_text,
        "maybe_manual_review",
        "medium",
        "outcome_related_but_team_unit_level_unclear",
        inclusion_signals,
        exclusion_signals,
        signals,
    )


def _build_result(
    screening_text: str,
    decision: str,
    confidence: str,
    primary_reason: str,
    inclusion_signals: list[str],
    exclusion_signals: list[str],
    signals: dict,
) -> dict:
    return {
        "screening_text": screening_text,
        "screen_decision": decision,
        "screen_confidence": confidence,
        "primary_reason": primary_reason,
        "inclusion_signals": format_signal_list(inclusion_signals),
        "exclusion_signals": format_signal_list(exclusion_signals),
        "direct_tmx_lmx_title_abstract_signal": format_signal_list(signals["direct_tmx_lmx"]),
        "performance_signal": format_signal_list(signals["performance"]),
        "team_context_signal": format_signal_list(signals["team_context"]),
        "team_level_signal": format_signal_list(signals["team_level"]),
        "quantitative_signal": format_signal_list(signals["quantitative"]),
        "wrong_context_signal": format_signal_list(signals["wrong_context"]),
        "wrong_publication_type_signal": format_signal_list(
            signals["ineligible_pub"] + signals["conference_meta"]
        ),
    }


def screen_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    screened = df.copy()
    results = screened.apply(classify_record, axis=1, result_type="expand")
    for col in SCREENING_COLUMNS:
        screened[col] = results[col]
    return screened


# =============================================================================
# Summary and audit
# =============================================================================

def build_screening_summary(df: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, str | int]] = []

    def add_row(
        metric: str,
        count: int,
        *,
        screen_decision: str = "",
        primary_reason: str = "",
        note: str = "",
    ) -> None:
        rows.append({
            "metric": metric,
            "screen_decision": screen_decision,
            "primary_reason": primary_reason,
            "count": count,
            "note": note,
        })

    total = len(df)
    add_row("total_input_records", total, note="All input records")

    decision_totals: list[int] = []
    for decision in ("include_likely", "maybe_manual_review", "exclude_likely"):
        decision_count = int((df["screen_decision"] == decision).sum())
        decision_totals.append(decision_count)
        add_row(
            f"count_{decision}",
            decision_count,
            screen_decision=decision,
            note=f"Total {decision}",
        )

        subset = df[df["screen_decision"] == decision]
        reason_counts = subset["primary_reason"].value_counts(dropna=False).sort_index()
        reason_sum = 0
        for reason, count in reason_counts.items():
            reason_label = str(reason)
            reason_sum += int(count)
            add_row(
                f"count_{decision}__primary_reason:{reason_label}",
                int(count),
                screen_decision=decision,
                primary_reason=reason_label,
                note="Component of decision total",
            )
        add_row(
            f"check_{decision}_primary_reason_sum",
            reason_sum,
            screen_decision=decision,
            note=f"Should equal count_{decision}",
        )

    add_row(
        "check_decision_counts_sum_to_total_input_records",
        sum(decision_totals),
        note="Should equal total_input_records",
    )

    return pd.DataFrame(rows, columns=["metric", "screen_decision", "primary_reason", "count", "note"])


def build_rules_audit() -> pd.DataFrame:
    sections = [
        ("overview", (
            "Abstract screening for TMX/LMX and team/group/unit-level performance. "
            "Records must contain a direct TMX/LMX-family term in the title or abstract, "
            "or a narrowly defined exchange-quality equivalent phrase, to be retained. "
            "The script does not apply publication date filtering; date range eligibility is "
            "assumed to have been handled during database downloading/export. "
            "Unpublished/working-paper/manuscript detection is not applied during abstract screening "
            "to avoid false exclusions from phrases such as 'working paper' or 'manuscript' in abstracts."
        )),
        ("direct_tmx_lmx_rule", (
            "Revised rule: records are excluded at abstract screening if no direct TMX/LMX-family "
            "term appears in the title or abstract. Equivalent terms are retained only when the "
            "title/abstract clearly defines or measures the construct as TMX/LMX exchange quality. "
            "General leadership, support, trust, cohesion, teamwork, or relationship terms are not sufficient."
        )),
        ("innovation_exclusion_rule", (
            "Revised rule: innovation, innovative performance, creative performance, creativity, "
            "and innovative behavior are not treated as eligible team performance/effectiveness outcomes. "
            "Records with only innovation/creativity outcomes are excluded."
        )),
        ("review_meta_exclusion_rule", (
            "Review/meta-analysis exclusions are context-aware. Journal names are not treated as article-type "
            "evidence, so journals such as Management Research Review do not trigger review exclusion by "
            "themselves. Weak background phrases such as literature review or research review do not trigger "
            "exclusion when the title/abstract clearly reports empirical data or statistical analysis. Strong "
            "review/meta-analysis article types in publication metadata or title are still excluded; strong "
            "phrases found only in the abstract are ignored when empirical evidence is present."
        )),
        ("outcome_exclusion_rule", (
            "Records with TMX/LMX but no performance/effectiveness-related outcome signal are excluded. "
            "Records with only non-performance outcomes (e.g., satisfaction, trust, cohesion, OCB, commitment, "
            "thriving, voice, psychological safety, engagement, burnout, turnover intention) and no eligible "
            "team/group/unit performance/effectiveness outcome are excluded."
        )),
        ("individual_performance_exclusion_rule", (
            "Records with only individual-level or firm/organization-level performance (e.g., employee performance, "
            "job performance, leader performance, firm performance, organizational performance) are excluded unless "
            "a credible team/group/unit/store/branch/outlet performance/effectiveness outcome is also indicated."
        )),
        ("vague_outcome_manual_review_rule", (
            "Records with vague performance/effectiveness wording but unclear team/group/unit level are retained "
            "for maybe_manual_review. Absence of raw correlation, sample size, or extractable effect size in the "
            "abstract is not a reason to exclude during abstract screening; those are full-text coding issues."
        )),
        ("summary_tab_rule", (
            "The screening_summary tab groups primary-reason subtotals under each screen_decision total, "
            "with check rows verifying that subtotals sum to their decision total and that decision totals "
            "sum to total_input_records. Aggregate signal-count metrics are not included in the summary."
        )),
        ("literature_review_rule", (
            "Weak review-background phrases (literature review, research review) are evaluated only in title "
            "and abstract, never from journal names. They are ignored when empirical-study evidence is present."
        )),
        ("language_rule", (
            "Only a clearly non-English title triggers exclusion (primary_reason = non_english). "
            "A non-English abstract alone is not an exclusion criterion when the title is English or ambiguous."
        )),
        ("decisions", "include_likely | maybe_manual_review | exclude_likely"),
        ("rule_priority_1", "clearly non-English title -> exclude_likely"),
        ("rule_priority_2", "insufficient title/abstract information -> maybe_manual_review"),
        ("rule_priority_3", "ineligible_publication_type -> exclude_likely"),
        ("rule_priority_4", "no_direct_tmx_lmx_in_title_or_abstract -> exclude_likely"),
        ("rule_priority_5", "innovation/creativity-only outcome -> exclude_likely"),
        ("rule_priority_6", "non-performance outcome only -> exclude_likely"),
        ("rule_priority_7", "individual/firm performance without team/unit outcome -> exclude_likely"),
        ("rule_priority_8", "TMX/LMX + eligible performance + clear team/unit level -> include_likely"),
        ("rule_priority_9", "TMX/LMX + eligible performance but level unclear -> maybe_manual_review"),
        ("rule_priority_10", "TMX/LMX but no performance/effectiveness-related outcome -> exclude_likely"),
        ("rule_priority_11", "TMX/LMX + vague outcome wording but team/unit level unclear -> maybe_manual_review"),
        ("direct_tmx_lmx_terms", "; ".join(DIRECT_TMX_LMX_PHRASES)),
        ("exchange_equivalent_terms", "; ".join(EXCHANGE_EQUIVALENT_PHRASES)),
        ("eligible_performance_terms", "; ".join(ELIGIBLE_TEAM_PERFORMANCE_PHRASES)),
        ("outcome_related_terms", "; ".join(PERFORMANCE_EFFECTIVENESS_OUTCOME_PHRASES)),
        ("innovation_creativity_exclusion_terms", "; ".join(INNOVATION_CREATIVITY_PHRASES)),
        ("team_context_terms", "; ".join(TEAM_CONTEXT_PHRASES[:30]) + "; ..."),
        ("team_level_terms", "; ".join(TEAM_LEVEL_PHRASES)),
        ("ineligible_publication_terms", (
            "strong: " + "; ".join(STRONG_REVIEW_META_PHRASES)
            + " | weak: " + "; ".join(WEAK_REVIEW_BACKGROUND_PHRASES)
            + " | other: " + "; ".join(OTHER_INELIGIBLE_PUBLICATION_PHRASES)
            + " | conference: " + "; ".join(CONFERENCE_METADATA_PHRASES)
        )),
        ("overlap_note", "Suspected duplicate/overlapping samples are not removed here; assess during full-text coding."),
        ("text_normalization", "Lowercase, Unicode dash/apostrophe normalization, ASCII fold, punctuation removal, space collapse."),
        ("screening_text_fields", "; ".join(SCREENING_TEXT_FIELDS)),
        ("focal_construct_search_scope", "Direct TMX/LMX detection uses title and abstract only."),
    ]
    return pd.DataFrame(sections, columns=["item", "description"])


# =============================================================================
# I/O
# =============================================================================

def load_input(path: Path, sheet: str) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")
    try:
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    except ValueError as exc:
        raise ValueError(f"Sheet '{sheet}' not found in {path}") from exc
    if df.empty:
        raise ValueError(f"Sheet '{sheet}' in {path} contains no records.")
    return df


def sort_screened_records(df: pd.DataFrame) -> pd.DataFrame:
    sorted_df = df.copy()
    sorted_df["_decision_order"] = sorted_df["screen_decision"].map(DECISION_ORDER).fillna(99)
    sorted_df["_confidence_order"] = sorted_df["screen_confidence"].map(CONFIDENCE_ORDER).fillna(99)
    title_col = "title" if "title" in sorted_df.columns else sorted_df.columns[0]
    sorted_df = sorted_df.sort_values(
        by=["_decision_order", "_confidence_order", title_col],
        ascending=[True, True, True],
        kind="stable",
    )
    return sorted_df.drop(columns=["_decision_order", "_confidence_order"])


def auto_fit_worksheet(worksheet, max_width: int = 50) -> None:
    for column_cells in worksheet.columns:
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))
        worksheet.column_dimensions[column_letter].width = min(max_length + 2, max_width)


def format_worksheet(worksheet) -> None:
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    auto_fit_worksheet(worksheet)


def write_output(
    screened: pd.DataFrame,
    summary: pd.DataFrame,
    rules_audit: pd.DataFrame,
    output_path: Path,
) -> None:
    screened_sorted = sort_screened_records(screened)
    if "screening_text" in screened_sorted.columns:
        screened_sorted = screened_sorted.drop(columns=["screening_text"])

    sheets = {
        "screened_records": screened_sorted,
        "screening_summary": summary,
        "rules_audit": rules_audit,
    }

    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, frame in sheets.items():
            safe_name = sheet_name[:31]
            frame.to_excel(writer, sheet_name=safe_name, index=False)
            format_worksheet(writer.sheets[safe_name])


# =============================================================================
# Classification tests
# =============================================================================

TEST_CASE_PAKISTANI_FEMALE_LEADERS = (
    "Leadership plays a significant role in the performance of individuals and organizations. "
    "This paper investigates the impact of leadership styles on the innovative performance of "
    "female leaders in Pakistani Universities using a survey approach. The investigation revealed "
    "that most female leaders practice the transactional leadership style. This study also discovered "
    "a moderate positive relationship between both leadership styles, namely transactional and "
    "transformational, and innovative performance."
)

TEST_CASE_LMX_TEAM = (
    "We examined leader-member exchange at the team level using aggregated member ratings. "
    "Team-level LMX was positively related to team performance in a sample of organizational work teams."
)

TEST_CASE_TRANSFORMATIONAL_TEAM = (
    "This study tests how transformational leadership influences team performance in organizational teams. "
    "Results from a survey of employees show a positive association between transformational leadership "
    "and team performance."
)

TEST_CASE_LMX_INDIVIDUAL_JOB = (
    "Leader-member exchange quality was positively associated with employee job performance in a "
    "field study of employees and their supervisors."
)

TEST_CASE_LMX_INDIVIDUAL_ONLY = (
    "Leader-member exchange quality was examined at the individual level only and was positively "
    "associated with employee job performance. The analysis found no team-level or unit-level effects."
)


TEST_CASE_LMX_TEAM_INNOVATION_ONLY = (
    "This study examined leader-member exchange within organizational teams. "
    "Results showed that LMX was positively related to team innovation and employee creativity."
)

TEST_CASE_LMX_TEAM_PERFORMANCE = (
    "We examined leader-member exchange and team performance in a sample of organizational work teams."
)

TEST_CASE_META_ANALYSIS_VARIANT = (
    "A meta-analytical correlation matrix was constructed from prior studies. "
    "Meta-analyses revealed consistent associations involving leader-member-exchange."
)

TEST_CASE_LMX_TRUST_ONLY = (
    "This study examines leader-member exchange and trust among work teams in organizational settings."
)

TEST_CASE_LMX_TEAM_PERFORMANCE_LEVEL_UNCLEAR = (
    "Leader-member exchange was related to team performance in organizations."
)

TEST_CASE_LMX_TEAM_LEVEL_PERFORMANCE = (
    "Team-level LMX was positively related to team performance in organizational work teams."
)

TEST_CASE_EMPIRICAL_LITERATURE_REVIEW = (
    "Purpose The purpose of this paper is to investigate the impact of native Chinese R&D team directors' "
    "differential leadership on team performance... A literature review on differential leadership and team "
    "conflict provided the model and hypothesis. Two-wave data collected from 103 directors and 344 "
    "subordinates from 103 R&D teams of high-tech enterprises from China's Pearl River Delta Area were used "
    "as empirical study samples. Hierarchical multiple regression analysis was conducted to test the model "
    "and hypothesis."
)

TEST_CASE_TRUE_LITERATURE_REVIEW = (
    "This article provides a literature review of differential leadership and organizational conflict. "
    "The paper synthesizes prior theory and proposes directions for future research without new observations."
)

TEST_CASE_ENGLISH_TITLE_NON_ENGLISH_ABSTRACT = (
    "Der Zweck dieser Studie ist es, den Zusammenhang zwischen Führungsqualität und Teamleistung "
    "in organisationalen Arbeitsgruppen zu untersuchen. Die Daten wurden aus einer Umfrage von "
    "Mitarbeitern und Vorgesetzten in mehreren Unternehmen erhoben."
)

TEST_CASE_NON_ENGLISH_TITLE = (
    "Diese Studie untersucht die Auswirkungen von Führungsqualität auf die Teamleistung in "
    "organisationalen Arbeitsgruppen mit Umfragedaten von Mitarbeitern und Vorgesetzten."
)

TEST_CASE_MANAGEMENT_RESEARCH_REVIEW_EMPIRICAL = (
    "Purpose The purpose of this study is to describe and explain the relationship between perceived "
    "social-organizational climate (PSOC), organizational citizenship behaviors (OCB) of other employees "
    "and innovative workplace behaviors (IWB) initiated and performed by employees. The mediating role of "
    "person-organization fit (P-O Fit) is tested within the relationship of PSOC, OCB and IWB. "
    "Design/methodology/approach The study was conducted anonymously on a group of 246 employees from "
    "76 companies operating in Poland. Structural equation modeling (SEM) was used in the process of "
    "statistical analysis."
)

TEST_CASE_MRR_LMX_TEAM = (
    "We examined leader-member exchange at the team level using aggregated member ratings in "
    "organizational work teams. Team-level LMX was positively related to team performance."
)


def _test_row(title: str, abstract: str, **extra) -> pd.Series:
    row = {
        "title": title,
        "abstract": abstract,
        "publicationDate": 2020,
        "year_clean": 2020,
    }
    row.update(extra)
    return pd.Series(row)


def run_classification_tests() -> bool:
    """Run built-in classification checks for key edge cases."""
    checks_passed = True

    def assert_classification(
        name: str,
        row: pd.Series,
        *,
        expected_decision: str | set[str] | None = None,
        expected_primary: str | None = None,
        forbidden_decisions: list[str] | None = None,
        forbidden_primary: str | None = None,
    ) -> None:
        nonlocal checks_passed
        result = classify_record(row)
        decision = result["screen_decision"]
        primary = result["primary_reason"]

        if expected_decision is not None:
            if isinstance(expected_decision, str):
                expected_decisions = {expected_decision}
            else:
                expected_decisions = set(expected_decision)

            if decision not in expected_decisions:
                print(f"  [FAIL] {name}: decision={decision!r}, expected one of {sorted(expected_decisions)!r}")
                checks_passed = False

        if expected_primary and primary != expected_primary:
            print(f"  [FAIL] {name}: primary_reason={primary!r}, expected {expected_primary!r}")
            checks_passed = False

        if forbidden_decisions and decision in forbidden_decisions:
            print(f"  [FAIL] {name}: decision {decision!r} should not be {forbidden_decisions!r}")
            checks_passed = False

        if forbidden_primary and primary == forbidden_primary:
            print(f"  [FAIL] {name}: primary_reason={primary!r} should not be {forbidden_primary!r}")
            checks_passed = False

    assert_classification(
        "Pakistani female leaders / transformational-transactional / innovative performance",
        _test_row("Leadership and innovative performance in universities", TEST_CASE_PAKISTANI_FEMALE_LEADERS),
        expected_decision="exclude_likely",
        expected_primary="no_direct_tmx_lmx_in_title_or_abstract",
    )

    assert_classification(
        "LMX with team innovation only (no team performance/effectiveness)",
        _test_row("LMX and team innovation", TEST_CASE_LMX_TEAM_INNOVATION_ONLY),
        expected_decision="exclude_likely",
        expected_primary="no_eligible_performance_effectiveness_outcome",
    )

    assert_classification(
        "Leader-member exchange and team performance",
        _test_row("Leader-member exchange and team performance", TEST_CASE_LMX_TEAM_PERFORMANCE),
        expected_decision={"include_likely", "maybe_manual_review"},
        forbidden_decisions=["exclude_likely"],
    )

    assert_classification(
        "Transformational leadership and team performance without LMX/TMX",
        _test_row("Transformational leadership and team performance", TEST_CASE_TRANSFORMATIONAL_TEAM),
        expected_decision="exclude_likely",
        expected_primary="no_direct_tmx_lmx_in_title_or_abstract",
    )

    assert_classification(
        "LMX team-level with team performance",
        _test_row("Team-level LMX and team performance", TEST_CASE_LMX_TEAM),
        expected_decision={"include_likely", "maybe_manual_review"},
        forbidden_decisions=["exclude_likely"],
    )

    assert_classification(
        "LMX with employee job performance (individual-level only)",
        _test_row("LMX and employee job performance", TEST_CASE_LMX_INDIVIDUAL_JOB),
        expected_decision="exclude_likely",
        expected_primary="no_eligible_performance_effectiveness_outcome",
    )

    assert_classification(
        "LMX with clearly individual-level only outcome",
        _test_row("LMX and individual job performance only", TEST_CASE_LMX_INDIVIDUAL_ONLY),
        expected_decision="exclude_likely",
        expected_primary="wrong_level_individual_only",
    )

    assert_classification(
        "Meta-analysis variant should be excluded",
        _test_row("Meta-analytical LMX synthesis", TEST_CASE_META_ANALYSIS_VARIANT),
        expected_decision="exclude_likely",
        expected_primary="ineligible_publication_type",
    )

    assert_classification(
        "LMX present but no performance/effectiveness outcome",
        _test_row("LMX and trust in teams", TEST_CASE_LMX_TRUST_ONLY),
        expected_decision="exclude_likely",
        expected_primary="no_eligible_performance_effectiveness_outcome",
    )

    assert_classification(
        "LMX + team performance but level unclear",
        _test_row("LMX and team performance", TEST_CASE_LMX_TEAM_PERFORMANCE_LEVEL_UNCLEAR),
        expected_decision={"include_likely", "maybe_manual_review"},
        forbidden_decisions=["exclude_likely"],
    )

    assert_classification(
        "LMX + team-level LMX + team performance",
        _test_row("Team-level LMX and team performance", TEST_CASE_LMX_TEAM_LEVEL_PERFORMANCE),
        expected_decision="include_likely",
    )

    assert_classification(
        "Empirical article from Management Research Review journal",
        _test_row(
            "Social-organizational climate and innovative workplace behaviors",
            TEST_CASE_MANAGEMENT_RESEARCH_REVIEW_EMPIRICAL,
            **{"academic journal": "Management Research Review"},
        ),
        expected_decision="exclude_likely",
        expected_primary="no_direct_tmx_lmx_in_title_or_abstract",
        forbidden_primary="ineligible_publication_type",
    )

    assert_classification(
        "Management Research Review journal with LMX and team performance",
        _test_row(
            "Team-level leader-member exchange and team performance",
            TEST_CASE_MRR_LMX_TEAM,
            **{"academic journal": "Management Research Review"},
        ),
        expected_decision={"include_likely", "maybe_manual_review"},
        forbidden_primary="ineligible_publication_type",
    )

    assert_classification(
        "Empirical article mentioning literature review should not be ineligible_publication_type",
        _test_row(
            "Differential leadership and team performance in R&D teams",
            TEST_CASE_EMPIRICAL_LITERATURE_REVIEW,
        ),
        expected_decision="exclude_likely",
        expected_primary="no_direct_tmx_lmx_in_title_or_abstract",
        forbidden_primary="ineligible_publication_type",
    )

    assert_classification(
        "True literature review without empirical study evidence",
        _test_row(
            "A literature review of differential leadership and team conflict",
            TEST_CASE_TRUE_LITERATURE_REVIEW,
        ),
        expected_decision="exclude_likely",
        expected_primary="ineligible_publication_type",
    )

    assert_classification(
        "Systematic review should still be excluded",
        _test_row(
            "A systematic review of leader-member exchange and team performance",
            "This systematic review synthesizes prior studies on leader-member exchange and team performance.",
        ),
        expected_decision="exclude_likely",
        expected_primary="ineligible_publication_type",
    )

    assert_classification(
        "English title with non-English abstract should not be excluded as non_english",
        _test_row(
            "Leader-member exchange and team performance",
            TEST_CASE_ENGLISH_TITLE_NON_ENGLISH_ABSTRACT,
        ),
        forbidden_primary="non_english",
    )

    assert_classification(
        "Clearly non-English title should be excluded as non_english",
        _test_row(
            "领导成员交换与团队绩效的关系研究",
            TEST_CASE_NON_ENGLISH_TITLE,
        ),
        expected_decision="exclude_likely",
        expected_primary="non_english",
    )

    if checks_passed:
        print("  [OK] All classification tests passed")
    return checks_passed


# =============================================================================
# CLI
# =============================================================================

def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Conservative abstract screening for TMX/LMX and team-level performance meta-analysis.",
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=DEFAULT_INPUT,
        help=f"Input Excel file (default: {DEFAULT_INPUT})",
    )
    parser.add_argument(
        "--sheet",
        default=DEFAULT_SHEET,
        help=f"Input sheet name (default: {DEFAULT_SHEET})",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help=f"Output Excel file (default: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--run-tests",
        action="store_true",
        help="Run built-in classification tests and exit",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    if args.run_tests:
        print("Running classification tests...")
        return 0 if run_classification_tests() else 1

    print("Conservative abstract screening")
    print(f"Input:  {args.input}")
    print(f"Sheet:  {args.sheet}")
    print(f"Output: {args.output}")
    print()

    try:
        df = load_input(args.input, args.sheet)
        print(f"Loaded {len(df)} records from '{args.sheet}'.")

        screened = screen_dataframe(df)
        summary = build_screening_summary(screened)
        rules_audit = build_rules_audit()
        write_output(screened, summary, rules_audit, args.output)

        print()
        print("Screening complete.")
        print(summary.to_string(index=False))
        print()
        print(f"[OK] Output written to: {args.output}")
        return 0

    except FileNotFoundError as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 1
    except ValueError as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"[ERROR] Unexpected failure: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
