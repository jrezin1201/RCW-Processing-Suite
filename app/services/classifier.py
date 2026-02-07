"""
Signal-based task classification service.

This classifier uses stable signals (tokens, patterns, markers) to classify
painting tasks into canonical categories. It's designed to handle varying
naming conventions used by different painters.

Categories:
- EXT PRIME: Exterior prime/prep work
- EXTERIOR: General exterior work
- EXTERIOR UA: Exterior with UA designation
- INTERIOR: Interior work
- ROLL WALLS FINAL: Roll walls/ceiling work
- TOUCH UP: Touch-up and punch work
- Q4 REVERSAL: Q4 accounting reversals (config-driven)
- UNMAPPED: Unclassified tasks (requires review)
"""

import os
import re
import logging
from typing import Optional, Dict, List, Any, Tuple
from datetime import datetime
from dataclasses import dataclass, field, asdict
import yaml

from app.models.schemas import ParsedRow

logger = logging.getLogger(__name__)

# Cache for loaded config
_config_cache: Optional[Dict[str, Any]] = None


@dataclass
class TaskSignals:
    """Extracted signals from a task string."""
    # Location markers
    is_exterior: bool = False
    is_interior: bool = False

    # Designation markers
    is_ua: bool = False
    is_op: bool = False
    is_ls: bool = False

    # Pattern markers
    has_job_code: bool = False
    job_code: Optional[str] = None

    # Keyword markers
    keyword_prime: bool = False
    keyword_touchup: bool = False
    keyword_rollwalls: bool = False

    # Date info
    raw_date: Optional[str] = None
    date_month: Optional[int] = None

    # Audit trail
    matched_patterns: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for serialization."""
        return asdict(self)


@dataclass
class ClassificationResult:
    """Result of classifying a task."""
    category: str
    rule_fired: str
    signals: TaskSignals
    normalized_text: str
    original_text: str
    debug_info: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for serialization."""
        return {
            "category": self.category,
            "rule_fired": self.rule_fired,
            "signals": self.signals.to_dict(),
            "normalized_text": self.normalized_text,
            "original_text": self.original_text,
            "debug_info": self.debug_info
        }


def load_config() -> Dict[str, Any]:
    """
    Load classification config from YAML file.

    Returns:
        Configuration dictionary
    """
    global _config_cache

    if _config_cache is not None:
        return _config_cache

    config_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "data",
        "classification_config.yaml"
    )

    try:
        with open(config_path, "r") as f:
            _config_cache = yaml.safe_load(f)
    except FileNotFoundError:
        logger.warning(f"Config file not found: {config_path}, using defaults")
        _config_cache = get_default_config()

    return _config_cache


def get_default_config() -> Dict[str, Any]:
    """Return default configuration if YAML file not found."""
    return {
        "prime_keywords": [
            "PRIME", "PRIMER", "UNDERCOAT", "SEAL", "SEALER",
            "SAND", "SANDING", "BLOCK", "BLOCKOUT", "CAULK",
            "PATCH", "PREP", "FASCIA"
        ],
        "touchup_keywords": [
            "TOUCH UP", "TOUCHUP", "TOUCH-UP", "PUNCH",
            "FINAL TOUCH", "AFTER CARPET"
        ],
        "rollwalls_keywords": [
            "ROLL WALL", "ROLL WALLS", "ROLLED WALL", "ROLLED WALLS",
            "ROLLING WALL", "ROLLING WALLS", "WALL ROLL", "WALLS ROLL",
            "CEILING", "CEILINGS"
        ],
        "q4_reversal": {
            "enabled": False,
            "months": [10, 11, 12]
        }
    }


def reload_config() -> Dict[str, Any]:
    """Force reload of configuration."""
    global _config_cache
    _config_cache = None
    return load_config()


def normalize_task_text(text: str) -> str:
    """
    Normalize task text for consistent signal extraction.

    - Converts to uppercase
    - Replaces multiple spaces with one
    - Trims whitespace
    - Normalizes separators (_, -, —, /) to spaces
    - Preserves parentheses/brackets for marker extraction

    Args:
        text: Raw task text

    Returns:
        Normalized text

    Example:
        "Painting - Interior / Prep & Enamel (437205) (INT) [LS]"
        -> "PAINTING   INTERIOR   PREP & ENAMEL (437205) (INT) [LS]"
    """
    if not text:
        return ""

    # Convert to uppercase
    normalized = text.upper()

    # Normalize separators to spaces (but keep the text readable)
    # Replace common separators with spaces
    normalized = re.sub(r'[-_—/]', ' ', normalized)

    # Replace multiple spaces with single space
    normalized = re.sub(r'\s+', ' ', normalized)

    # Trim
    normalized = normalized.strip()

    return normalized


def parse_date(date_value: Any) -> Tuple[Optional[str], Optional[int]]:
    """
    Parse a date value and extract the month.

    Args:
        date_value: Date as string, datetime, or None

    Returns:
        Tuple of (formatted date string YYYY-MM-DD, month number 1-12)
    """
    if date_value is None:
        return None, None

    if isinstance(date_value, datetime):
        return date_value.strftime("%Y-%m-%d"), date_value.month

    # Try to parse string date
    date_str = str(date_value).strip()

    # Common date formats
    formats = [
        "%Y-%m-%d",      # 2026-03-20
        "%m/%d/%Y",      # 03/20/2026
        "%m/%d/%y",      # 03/20/26
        "%d/%m/%Y",      # 20/03/2026
        "%Y-%m-%d %H:%M:%S",  # 2026-03-20 00:00:00
    ]

    for fmt in formats:
        try:
            dt = datetime.strptime(date_str, fmt)
            return dt.strftime("%Y-%m-%d"), dt.month
        except ValueError:
            continue

    return None, None


def parse_signals(task: str, date_value: Any = None) -> TaskSignals:
    """
    Extract stable signals from a task string.

    Signals extracted:
    - is_exterior: (EXT) marker or standalone EXT token
    - is_interior: (INT) marker or standalone INT token
    - is_ua: [UA] marker or standalone UA token
    - is_op: [OP] marker
    - is_ls: [LS] marker
    - has_job_code: 5+ digit code in parentheses
    - keyword_prime: Prime/prep related keywords
    - keyword_touchup: Touch-up related keywords
    - keyword_rollwalls: Roll walls/ceiling keywords

    Args:
        task: Task text (will be normalized internally)
        date_value: Optional date for date-based rules

    Returns:
        TaskSignals dataclass with extracted signals
    """
    signals = TaskSignals()
    normalized = normalize_task_text(task)

    if not normalized:
        return signals

    config = load_config()

    # Parse date
    signals.raw_date, signals.date_month = parse_date(date_value)

    # Check for location markers - parenthesized markers are most reliable
    if "(EXT)" in normalized:
        signals.is_exterior = True
        signals.matched_patterns.append("(EXT) marker")
    elif re.search(r'\bEXT\b', normalized) and not re.search(r'\bEXTERIOR\b', normalized):
        # Standalone EXT (word boundary) but not part of EXTERIOR
        signals.is_exterior = True
        signals.matched_patterns.append("EXT token")
    elif re.search(r'\bEXTERIOR\b', normalized):
        signals.is_exterior = True
        signals.matched_patterns.append("EXTERIOR keyword")

    if "(INT)" in normalized:
        signals.is_interior = True
        signals.matched_patterns.append("(INT) marker")
    elif re.search(r'\bINT\b', normalized) and not re.search(r'\bINTERIOR\b', normalized):
        signals.is_interior = True
        signals.matched_patterns.append("INT token")
    elif re.search(r'\bINTERIOR\b', normalized):
        signals.is_interior = True
        signals.matched_patterns.append("INTERIOR keyword")

    # Check for designation markers
    if "[UA]" in normalized:
        signals.is_ua = True
        signals.matched_patterns.append("[UA] marker")
    elif re.search(r'\bUA\b', normalized):
        signals.is_ua = True
        signals.matched_patterns.append("UA token")

    if "[OP]" in normalized:
        signals.is_op = True
        signals.matched_patterns.append("[OP] marker")

    if "[LS]" in normalized:
        signals.is_ls = True
        signals.matched_patterns.append("[LS] marker")

    # Check for job code pattern (5+ digit number in parentheses)
    job_code_match = re.search(r'\((\d{5,})\)', normalized)
    if job_code_match:
        signals.has_job_code = True
        signals.job_code = job_code_match.group(1)
        signals.matched_patterns.append(f"job_code:{signals.job_code}")

    # Check for prime/prep keywords (word boundary match)
    prime_keywords = config.get("prime_keywords", [])
    for keyword in prime_keywords:
        pattern = r'\b' + re.escape(keyword.upper()) + r'\b'
        if re.search(pattern, normalized):
            signals.keyword_prime = True
            signals.matched_patterns.append(f"prime_keyword:{keyword}")
            break

    # Check for touch-up keywords
    touchup_keywords = config.get("touchup_keywords", [])
    for keyword in touchup_keywords:
        if keyword.upper() in normalized:
            signals.keyword_touchup = True
            signals.matched_patterns.append(f"touchup_keyword:{keyword}")
            break

    # Check for roll walls keywords
    # Special handling: only treat as rollwalls if interior-like OR "ROLL" near "WALL"
    rollwalls_keywords = config.get("rollwalls_keywords", [])
    for keyword in rollwalls_keywords:
        if keyword.upper() in normalized:
            # Additional check: if it has ROLL and WALL within proximity
            if "ROLL" in keyword.upper() and "WALL" in keyword.upper():
                signals.keyword_rollwalls = True
                signals.matched_patterns.append(f"rollwalls_keyword:{keyword}")
                break
            elif "CEILING" in keyword.upper():
                signals.keyword_rollwalls = True
                signals.matched_patterns.append(f"rollwalls_keyword:{keyword}")
                break
            elif signals.is_interior or not signals.is_exterior:
                # Only apply generic rollwalls keywords if interior or unspecified
                signals.keyword_rollwalls = True
                signals.matched_patterns.append(f"rollwalls_keyword:{keyword}")
                break

    # Additional check: ROLL near WALL (within ~30 chars)
    roll_match = re.search(r'ROLL\w*\s+\w*\s*WALL', normalized)
    if roll_match and not signals.keyword_rollwalls:
        signals.keyword_rollwalls = True
        signals.matched_patterns.append("roll_near_wall")

    return signals


def classify_task(
    signals: TaskSignals,
    normalized_text: str,
    original_text: str
) -> ClassificationResult:
    """
    Classify a task based on extracted signals.

    Rules are applied in this order (first match wins):
    A) Q4 REVERSAL - config-driven, date-based
    B) EXT PRIME - exterior + prime keywords
    C) EXTERIOR UA - exterior + UA designation
    D) EXTERIOR - exterior location
    E) TOUCH UP - touch-up keywords
    F) ROLL WALLS FINAL - interior/neutral + roll walls keywords
    G) INTERIOR - interior location
    H) UNMAPPED - fallback

    Args:
        signals: Extracted TaskSignals
        normalized_text: Normalized task text
        original_text: Original task text

    Returns:
        ClassificationResult with category and debug info
    """
    config = load_config()

    def make_result(category: str, rule: str, **extra_debug) -> ClassificationResult:
        return ClassificationResult(
            category=category,
            rule_fired=rule,
            signals=signals,
            normalized_text=normalized_text,
            original_text=original_text,
            debug_info=extra_debug
        )

    # Rule A: Q4 REVERSAL
    q4_config = config.get("q4_reversal", {})
    if q4_config.get("enabled", False):
        q4_months = q4_config.get("months", [10, 11, 12])
        marker_regex = q4_config.get("marker_regex")

        if signals.date_month and signals.date_month in q4_months:
            # Check for marker if regex specified
            if marker_regex:
                if re.search(marker_regex, normalized_text, re.IGNORECASE):
                    return make_result("Q4 REVERSAL", "rule_a_q4_reversal_with_marker")
            else:
                # If no marker regex, we cannot reliably classify as Q4 REVERSAL
                # This will fall through to other rules
                pass

    # Rule B: EXT PRIME
    if signals.is_exterior and signals.keyword_prime:
        return make_result("EXT PRIME", "rule_b_ext_prime")

    # Rule C: EXTERIOR UA
    if signals.is_exterior and signals.is_ua:
        return make_result("EXTERIOR UA", "rule_c_exterior_ua")

    # Rule D: EXTERIOR
    if signals.is_exterior:
        return make_result("EXTERIOR", "rule_d_exterior")

    # Rule E: TOUCH UP
    if signals.keyword_touchup:
        return make_result("TOUCH UP", "rule_e_touch_up")

    # Rule F: ROLL WALLS FINAL
    # Apply if interior OR not exterior (neutral) AND has rollwalls keywords
    if (signals.is_interior or not signals.is_exterior) and signals.keyword_rollwalls:
        return make_result("ROLL WALLS FINAL", "rule_f_roll_walls_final")

    # Rule G: INTERIOR
    if signals.is_interior:
        return make_result("INTERIOR", "rule_g_interior")

    # Rule H: UNMAPPED (fallback)
    unmapped_reason = "no_matching_rule"
    if not signals.is_exterior and not signals.is_interior:
        unmapped_reason = "no_location_marker"

    return make_result(
        "UNMAPPED",
        "rule_h_unmapped",
        unmapped_reason=unmapped_reason
    )


def classify_row(row: ParsedRow) -> str:
    """
    Classify a parsed row into a bucket.

    This is the main entry point that maintains backward compatibility
    with the existing codebase.

    Args:
        row: The parsed row to classify

    Returns:
        Category string
    """
    if not row.task_text and not row.task_text_raw:
        return "UNMAPPED"

    # Use raw task text if available for more complete signal extraction
    task_text = row.task_text_raw or row.task_text

    # Extract signals
    signals = parse_signals(task_text, row.task_start_date)

    # Normalize text
    normalized = normalize_task_text(task_text)

    # Classify
    result = classify_task(signals, normalized, task_text)

    return result.category


def classify_row_with_details(row: ParsedRow) -> ClassificationResult:
    """
    Classify a parsed row and return full details.

    Use this when you need the complete audit trail.

    Args:
        row: The parsed row to classify

    Returns:
        ClassificationResult with full details
    """
    if not row.task_text and not row.task_text_raw:
        return ClassificationResult(
            category="UNMAPPED",
            rule_fired="no_task_text",
            signals=TaskSignals(),
            normalized_text="",
            original_text="",
            debug_info={"reason": "empty_task_text"}
        )

    # Use raw task text if available
    task_text = row.task_text_raw or row.task_text

    # Extract signals
    signals = parse_signals(task_text, row.task_start_date)

    # Normalize text
    normalized = normalize_task_text(task_text)

    # Classify
    return classify_task(signals, normalized, task_text)


def classify_rows(rows: List[ParsedRow]) -> Dict[str, List[ParsedRow]]:
    """
    Classify multiple rows into buckets.

    Args:
        rows: List of parsed rows

    Returns:
        Dictionary mapping category names to lists of rows
    """
    buckets: Dict[str, List[ParsedRow]] = {}

    for row in rows:
        category = classify_row(row)
        if category not in buckets:
            buckets[category] = []
        buckets[category].append(row)

    return buckets


def get_classification_summary(rows: List[ParsedRow]) -> Dict[str, Any]:
    """
    Generate a classification summary report.

    Args:
        rows: List of parsed rows

    Returns:
        Summary dictionary with counts and unmapped details
    """
    counts: Dict[str, int] = {}
    unmapped_details: List[Dict[str, Any]] = []

    for row in rows:
        result = classify_row_with_details(row)

        # Count categories
        counts[result.category] = counts.get(result.category, 0) + 1

        # Track unmapped details
        if result.category == "UNMAPPED":
            unmapped_details.append({
                "date": result.signals.raw_date,
                "original_task": result.original_text,
                "normalized_task": result.normalized_text,
                "signals": result.signals.to_dict(),
                "rule_fired": result.rule_fired,
                "debug_info": result.debug_info
            })

    # Sort unmapped by frequency (group by normalized text)
    unmapped_grouped: Dict[str, Dict[str, Any]] = {}
    for item in unmapped_details:
        key = item["normalized_task"]
        if key not in unmapped_grouped:
            unmapped_grouped[key] = {
                "normalized_task": key,
                "original_examples": [],
                "count": 0,
                "signals": item["signals"],
                "rule_fired": item["rule_fired"]
            }
        unmapped_grouped[key]["count"] += 1
        if len(unmapped_grouped[key]["original_examples"]) < 3:
            unmapped_grouped[key]["original_examples"].append(item["original_task"])

    # Sort by count descending
    unmapped_sorted = sorted(
        unmapped_grouped.values(),
        key=lambda x: x["count"],
        reverse=True
    )

    return {
        "total_rows": len(rows),
        "counts_per_category": counts,
        "unmapped_count": counts.get("UNMAPPED", 0),
        "unmapped_details": unmapped_sorted[:20],  # Top 20
        "categories_seen": list(counts.keys())
    }
