"""
Category mapping service with template-first approach and auto-category creation.

Key principle: NEVER lose dollars. If a task doesn't map to an existing
template category, create a new category column automatically.
"""

import re
import logging
from typing import Optional, Dict, List, Any, Tuple, Set
from dataclasses import dataclass, field

logger = logging.getLogger(__name__)


@dataclass
class TaskSignals:
    """Extracted signals from a task string."""
    is_ext: bool = False
    is_int: bool = False
    is_ua: bool = False
    is_op: bool = False
    is_ls: bool = False

    keyword_undercoat: bool = False
    keyword_prime: bool = False
    keyword_touchup: bool = False
    keyword_rollwalls: bool = False
    keyword_baseshoe: bool = False

    matched_keywords: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "is_ext": self.is_ext,
            "is_int": self.is_int,
            "is_ua": self.is_ua,
            "is_op": self.is_op,
            "is_ls": self.is_ls,
            "keyword_undercoat": self.keyword_undercoat,
            "keyword_prime": self.keyword_prime,
            "keyword_touchup": self.keyword_touchup,
            "keyword_rollwalls": self.keyword_rollwalls,
            "keyword_baseshoe": self.keyword_baseshoe,
            "matched_keywords": self.matched_keywords,
        }


@dataclass
class MappingResult:
    """Result of category mapping."""
    category_display: str
    reason: str
    is_new_category: bool = False
    signals: TaskSignals = None
    scope_fragment: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "category_display": self.category_display,
            "reason": self.reason,
            "is_new_category": self.is_new_category,
            "signals": self.signals.to_dict() if self.signals else {},
            "scope_fragment": self.scope_fragment,
        }


def canonical(s: str) -> str:
    """
    Convert string to canonical form for matching.
    - Uppercase
    - Collapse spaces
    - Strip punctuation from ends
    """
    if not s:
        return ""
    result = s.upper().strip()
    # Remove leading/trailing punctuation
    result = re.sub(r'^[^\w]+|[^\w]+$', '', result)
    # Collapse whitespace
    result = re.sub(r'\s+', ' ', result)
    return result


def normalize_task_text(text: str) -> str:
    """
    Normalize task text for signal extraction.

    - Uppercase
    - Replace separators (-, /, _, —) with spaces (keep &)
    - Collapse whitespace
    - Keep parentheses/brackets for marker extraction
    """
    if not text:
        return ""

    result = text.upper()
    # Replace separators with spaces (but NOT &)
    result = re.sub(r'[-/_—]', ' ', result)
    # Collapse whitespace
    result = re.sub(r'\s+', ' ', result)
    return result.strip()


def extract_scope_fragment(task_text: str) -> str:
    """
    Extract a clean scope fragment from task text for creating new column names.

    - Remove leading date if present
    - Remove leading "PAINTING" or "PAINTING -" prefix
    - Strip bracket codes [LS], [OP], [UA]
    - Strip numeric job codes in parentheses (437225)
    - Strip trailing (INT)/(EXT) markers

    Returns a cleaned human-readable scope fragment.
    """
    if not task_text:
        return ""

    text = task_text.strip()

    # Remove leading date patterns (e.g., "2026-03-20" or "03/20/2026")
    text = re.sub(r'^\d{4}[-/]\d{2}[-/]\d{2}\s*', '', text)
    text = re.sub(r'^\d{2}[-/]\d{2}[-/]\d{4}\s*', '', text)

    # Remove "PAINTING" or "PAINTING -" prefix (case-insensitive)
    text = re.sub(r'^PAINTING\s*[-–—]?\s*', '', text, flags=re.IGNORECASE)

    # Strip bracket codes [LS], [OP], [UA], [578700 - 34749538-000], etc.
    text = re.sub(r'\[[\w\s\-]+\]', '', text)

    # Strip numeric job codes in parentheses (5+ digits)
    text = re.sub(r'\(\d{5,}\)', '', text)

    # Strip (INT), (EXT) markers
    text = re.sub(r'\(INT\)', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\(EXT\)', '', text, flags=re.IGNORECASE)

    # Clean up remaining artifacts
    text = re.sub(r'\s*[-–—]\s*$', '', text)  # Trailing dashes
    text = re.sub(r'\s+', ' ', text)  # Collapse whitespace

    return text.strip()


def parse_signals(task_text: str) -> TaskSignals:
    """
    Extract stable signals from normalized task text.
    """
    signals = TaskSignals()

    if not task_text:
        return signals

    normalized = normalize_task_text(task_text)

    # Location markers
    if "(EXT)" in normalized:
        signals.is_ext = True
        signals.matched_keywords.append("(EXT)")
    elif re.search(r'\bEXT\b', normalized) and "EXTERIOR" not in normalized:
        signals.is_ext = True
        signals.matched_keywords.append("EXT token")
    elif "EXTERIOR" in normalized:
        signals.is_ext = True
        signals.matched_keywords.append("EXTERIOR")

    if "(INT)" in normalized:
        signals.is_int = True
        signals.matched_keywords.append("(INT)")
    elif re.search(r'\bINT\b', normalized) and "INTERIOR" not in normalized:
        signals.is_int = True
        signals.matched_keywords.append("INT token")
    elif "INTERIOR" in normalized:
        signals.is_int = True
        signals.matched_keywords.append("INTERIOR")

    # Designation markers
    if "[UA]" in normalized or re.search(r'\bUA\b', normalized):
        signals.is_ua = True
        signals.matched_keywords.append("UA")

    if "[OP]" in normalized:
        signals.is_op = True
        signals.matched_keywords.append("[OP]")

    if "[LS]" in normalized:
        signals.is_ls = True
        signals.matched_keywords.append("[LS]")

    # Keyword: UNDERCOAT
    undercoat_keywords = ["UNDERCOAT", "FIRST COAT"]
    for kw in undercoat_keywords:
        if kw in normalized:
            signals.keyword_undercoat = True
            signals.matched_keywords.append(f"undercoat:{kw}")
            break

    # Keyword: PRIME
    prime_keywords = [
        "PRIME", "PRIMER", "PRIMING", "SEAL", "SEALER",
        "SAND", "BLOCK", "BLOCKOUT", "PREP", "CAULK", "PATCH", "FASCIA"
    ]
    for kw in prime_keywords:
        if re.search(r'\b' + re.escape(kw) + r'\b', normalized):
            signals.keyword_prime = True
            signals.matched_keywords.append(f"prime:{kw}")
            break

    # Keyword: TOUCHUP
    touchup_keywords = ["TOUCH UP", "TOUCHUP", "TOUCH-UP", "PUNCH", "AFTER CARPET"]
    for kw in touchup_keywords:
        if kw.replace("-", " ") in normalized or kw in normalized:
            signals.keyword_touchup = True
            signals.matched_keywords.append(f"touchup:{kw}")
            break

    # Keyword: ROLLWALLS - ROLL near WALL/WALLS/CEILING
    if re.search(r'ROLL\w*\s+(?:\w+\s+){0,10}(?:WALL|WALLS|CEILING)', normalized):
        signals.keyword_rollwalls = True
        signals.matched_keywords.append("rollwalls:ROLL near WALL/CEILING")
    elif "ROLL WALL" in normalized or "ROLLED WALL" in normalized:
        signals.keyword_rollwalls = True
        signals.matched_keywords.append("rollwalls:ROLL WALL")

    # Keyword: BASESHOE
    baseshoe_keywords = ["BASE SHOE", "BASEBOARD", "SHOE MOLD", "SHOE MOULD"]
    for kw in baseshoe_keywords:
        if kw in normalized:
            signals.keyword_baseshoe = True
            signals.matched_keywords.append(f"baseshoe:{kw}")
            break

    return signals


def map_category(
    task_text: str,
    signals: TaskSignals,
    template_canon_to_display: Dict[str, str]
) -> Tuple[Optional[str], str]:
    """
    Map task to existing template category (template-first approach).

    Returns:
        (category_display, reason) - category_display is None if no match
    """
    # Helper to check if canonical category exists in template
    def template_has(canon_key: str) -> Optional[str]:
        """Return display name if canonical key exists in template."""
        return template_canon_to_display.get(canon_key)

    # 1) BASE SHOE
    if signals.keyword_baseshoe:
        if signals.is_ua:
            # UA variant — check for "BASE SHOE UA" in template, else auto-create
            display = template_has("BASE SHOE UA")
            if display:
                return (display, "matched_baseshoe_ua")
            return (None, "unmapped_baseshoe_ua")
        display = template_has("BASE SHOE")
        if display:
            return (display, "matched_baseshoe")

    # 2) UNDERCOAT
    if signals.keyword_undercoat:
        if signals.is_ua:
            display = template_has("UNDERCOAT UA")
            if display:
                return (display, "matched_undercoat_ua")
            return (None, "unmapped_undercoat_ua")
        display = template_has("UNDERCOAT")
        if display:
            return (display, "matched_undercoat")

    # 3) TOUCH UP
    if signals.keyword_touchup:
        if signals.is_ua:
            display = template_has("TOUCH UP UA")
            if display:
                return (display, "matched_touchup_ua")
            return (None, "unmapped_touchup_ua")
        display = template_has("TOUCH UP")
        if display:
            return (display, "matched_touchup")

    # 4) ROLL WALLS FINAL
    if signals.keyword_rollwalls:
        if signals.is_ua:
            display = template_has("ROLL WALLS FINAL UA")
            if display:
                return (display, "matched_rollwalls_ua")
            return (None, "unmapped_rollwalls_ua")
        display = template_has("ROLL WALLS FINAL")
        if display:
            return (display, "matched_rollwalls")

    # 5) EXT PRIME
    if signals.is_ext and signals.keyword_prime:
        if signals.is_ua:
            display = template_has("EXT PRIME UA")
            if display:
                return (display, "matched_ext_prime_ua")
            return (None, "unmapped_ext_prime_ua")
        display = template_has("EXT PRIME")
        if display:
            return (display, "matched_ext_prime")

    # Compute scope once for rules 6-8
    scope = extract_scope_fragment(task_text)
    # Strip parenthetical content for comparison (e.g. "(Flooring Orders)")
    scope_core = re.sub(r'\([^)]*\)', '', scope).strip().upper()

    # 6) EXTERIOR UA — only for generic exterior UA tasks
    # Distinctive scopes like "Spray Overhang (EXT) [UA]" should auto-create
    if signals.is_ext and signals.is_ua:
        scope_mentions_ext = (
            scope_core in {"", "EXTERIOR", "EXT", "EXTERIOR PAINTING"}
            or re.search(r'\bEXTERIOR\b', scope_core)
            or re.search(r'\bEXT\b', scope_core)
        )
        if scope_mentions_ext:
            display = template_has("EXTERIOR UA")
            if display:
                return (display, "matched_exterior_ua")

    # 7) EXTERIOR — only for non-UA generic exterior tasks
    if signals.is_ext and not signals.is_ua:
        scope_mentions_ext = (
            scope_core in {"", "EXTERIOR", "EXT", "EXTERIOR PAINTING"}
            or re.search(r'\bEXTERIOR\b', scope_core)
            or re.search(r'\bEXT\b', scope_core)
        )
        if scope_mentions_ext:
            display = template_has("EXTERIOR")
            if display:
                return (display, "matched_exterior")

    # 8a) INTERIOR UA — only for generic interior UA tasks
    if signals.is_int and signals.is_ua:
        scope_mentions_int = (
            scope_core in {"", "INTERIOR", "INT", "INTERIOR PAINTING"}
            or re.search(r'\bINTERIOR\b', scope_core)
            or re.search(r'\bINT\b', scope_core)
        )
        if scope_mentions_int:
            display = template_has("INTERIOR UA")
            if display:
                return (display, "matched_interior_ua")
            return (None, "unmapped_interior_ua")

    # 8b) INTERIOR — only for non-UA generic interior tasks
    if signals.is_int and not signals.is_ua:
        scope_mentions_int = (
            scope_core in {"", "INTERIOR", "INT", "INTERIOR PAINTING"}
            or re.search(r'\bINTERIOR\b', scope_core)
            or re.search(r'\bINT\b', scope_core)
        )
        if scope_mentions_int:
            display = template_has("INTERIOR")
            if display:
                return (display, "matched_interior")

    # No template match — auto-creation will handle this
    return (None, "unmapped_template")


def compute_base_category_name(task_text: str, signals: TaskSignals) -> str:
    """
    Compute the base category name for a task (without uniqueness suffix).

    This is used to check if a matching category already exists before
    creating a new one.
    """
    fragment = extract_scope_fragment(task_text)

    if not fragment:
        fragment = "MISC"

    name = fragment.upper()
    name = re.sub(r'\s+', ' ', name).strip()

    # Add disambiguation based on signals
    has_ext_prefix = name.startswith("EXT") or "EXTERIOR" in name
    has_int_prefix = name.startswith("INT") or "INTERIOR" in name

    if signals.is_ext and not has_ext_prefix and not has_int_prefix:
        name = "EXT " + name
    elif signals.is_int and not has_int_prefix and not has_ext_prefix:
        name = "INT " + name

    # Limit length to avoid ugly headers — truncate BEFORE adding UA
    # so the UA suffix is never lost
    max_len = 35
    if signals.is_ua and "UA" not in name:
        # Reserve space for " UA" suffix
        truncate_limit = max_len - 3  # len(" UA") == 3
    else:
        truncate_limit = max_len

    if len(name) > truncate_limit:
        name = name[:truncate_limit].rsplit(' ', 1)[0]  # Cut at word boundary
        name = name.rstrip(' /-&')  # Clean trailing separators

    # Add UA suffix AFTER truncation so it's never lost
    if signals.is_ua and "UA" not in name:
        name = name + " UA"

    return name


def create_category_name(
    task_text: str,
    signals: TaskSignals,
    existing_canonicals: Set[str]
) -> str:
    """
    Create a new unique category name from task text for auto-category creation.
    """
    name = compute_base_category_name(task_text, signals)

    # Ensure uniqueness
    base_name = name
    counter = 1
    while canonical(name) in existing_canonicals:
        counter += 1
        name = f"{base_name} {counter}"

    return name


def _strip_prefix(h: str) -> str:
    """Strip INT/EXT/INTERIOR/EXTERIOR prefix from header for matching."""
    for prefix in ['INTERIOR ', 'EXTERIOR ', 'INT ', 'EXT ']:
        if h.startswith(prefix):
            h = h[len(prefix):].strip()
            break
    # Strip leading punctuation left behind (e.g. "/ PREP & ENAMEL")
    h = h.lstrip('/ -&')
    return h


def organize_headers(headers: List[str]) -> List[str]:
    """
    Organize headers so UA variants are placed next to their base category.

    Uses multi-strategy matching:
    1. Exact match on normalized names (strip prefix + UA suffix)
    2. Normalized substring match (higher priority)
    3. Raw substring match (fallback)
    """
    ua_headers = [h for h in headers if h.upper().strip().endswith(' UA')]
    non_ua_headers = [h for h in headers if not h.upper().strip().endswith(' UA')]

    # Match each UA header to its best non-UA base
    ua_to_base: Dict[str, str] = {}

    for ua in ua_headers:
        ua_upper = ua.upper().strip()
        ua_raw = ua_upper[:-3].strip()  # Strip " UA"
        ua_norm = _strip_prefix(ua_raw)

        best_match = None
        best_score = 0

        for base in non_ua_headers:
            base_upper = base.upper().strip()
            base_norm = _strip_prefix(base_upper)

            # Strategy 1: Exact normalized match
            if ua_norm == base_norm:
                best_match = base
                break

            # Strategy 2: Normalized substring (boosted score)
            if base_norm in ua_norm and len(base_norm) + 100 > best_score:
                best_match = base
                best_score = len(base_norm) + 100
            elif ua_norm in base_norm and len(ua_norm) + 100 > best_score:
                best_match = base
                best_score = len(ua_norm) + 100

            # Strategy 3: Raw substring (lower priority fallback)
            if base_upper in ua_raw and len(base_upper) > best_score and best_score < 100:
                best_match = base
                best_score = len(base_upper)

        if best_match:
            ua_to_base[ua] = best_match

    # Build result: each non-UA header followed by its UA partner(s)
    result = []
    placed_ua: Set[str] = set()

    for header in non_ua_headers:
        result.append(header)
        for ua in ua_headers:
            if ua not in placed_ua and ua_to_base.get(ua) == header:
                result.append(ua)
                placed_ua.add(ua)

    # Append any unmatched UA headers at the end
    for ua in ua_headers:
        if ua not in placed_ua:
            result.append(ua)

    return result


class CategoryMapper:
    """
    Manages category mapping with template-first approach and auto-creation.
    """

    def __init__(self, template_headers: List[str] = None):
        """
        Initialize with template headers.

        Args:
            template_headers: List of category headers from template
                             (in order, between PLAN and TOTAL)
        """
        # Default canonical categories if no template provided
        self.default_categories = [
            "EXT PRIME", "EXTERIOR", "EXTERIOR UA", "INTERIOR",
            "BASE SHOE", "ROLL WALLS FINAL", "TOUCH UP", "Q4 REVERSAL"
        ]

        if template_headers:
            self.category_headers = list(template_headers)
        else:
            self.category_headers = list(self.default_categories)

        # Build canonical -> display mapping
        self.template_canon_to_display: Dict[str, str] = {}
        for header in self.category_headers:
            self.template_canon_to_display[canonical(header)] = header

        # Track created categories
        self.created_categories: List[Dict[str, Any]] = []
        self._created_canonicals: Set[str] = set()

    def get_all_canonicals(self) -> Set[str]:
        """Get all canonical category names (template + created)."""
        return set(self.template_canon_to_display.keys()) | self._created_canonicals

    def map_task(self, task_text: str) -> MappingResult:
        """
        Map a task to a category.

        If no template category matches, auto-creates a new category.
        """
        # Parse signals
        signals = parse_signals(task_text)

        # Try template-first mapping
        category_display, reason = map_category(
            task_text, signals, self.template_canon_to_display
        )

        if category_display:
            return MappingResult(
                category_display=category_display,
                reason=reason,
                is_new_category=False,
                signals=signals,
                scope_fragment=extract_scope_fragment(task_text)
            )

        # Check if the base name already exists (reuse previously auto-created)
        base_name = compute_base_category_name(task_text, signals)
        base_canon = canonical(base_name)

        if base_canon in self.template_canon_to_display:
            return MappingResult(
                category_display=self.template_canon_to_display[base_canon],
                reason="reused_created_category",
                is_new_category=False,
                signals=signals,
                scope_fragment=extract_scope_fragment(task_text)
            )

        # Truly new category — create with uniqueness suffix if needed
        new_name = create_category_name(
            task_text, signals, self.get_all_canonicals()
        )
        new_canon = canonical(new_name)

        # Register the new category
        self._created_canonicals.add(new_canon)
        self.category_headers.append(new_name)
        self.template_canon_to_display[new_canon] = new_name

        # Record for QA
        self.created_categories.append({
            "header": new_name,
            "example_tasks": [task_text],
            "reason": "auto_created",
            "signals": signals.to_dict()
        })

        return MappingResult(
            category_display=new_name,
            reason="auto_created",
            is_new_category=True,
            signals=signals,
            scope_fragment=extract_scope_fragment(task_text)
        )

    def add_example_to_created_category(self, category: str, task_text: str):
        """Add example task to a created category (for QA reporting)."""
        canon = canonical(category)
        for cat in self.created_categories:
            if canonical(cat["header"]) == canon:
                if len(cat["example_tasks"]) < 3:
                    cat["example_tasks"].append(task_text)
                break

    def get_category_headers(self) -> List[str]:
        """Get all category headers (template + created) in order."""
        return self.category_headers

    def get_created_categories_report(self) -> List[Dict[str, Any]]:
        """Get report of auto-created categories."""
        return self.created_categories
