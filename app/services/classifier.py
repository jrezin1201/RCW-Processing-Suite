"""Task classification service using YAML-defined rules."""
import os
from typing import Optional, Dict, List, Any
import yaml

from app.models.schemas import ParsedRow


# Cache for loaded rules
_rules_cache: Optional[List[Dict[str, Any]]] = None


def load_rules() -> List[Dict[str, Any]]:
    """
    Load classification rules from YAML file.

    Returns:
        List of classification rules
    """
    global _rules_cache

    if _rules_cache is not None:
        return _rules_cache

    # Get the path to the YAML file
    yaml_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "data",
        "mapping_rules.yaml"
    )

    with open(yaml_path, "r") as f:
        data = yaml.safe_load(f)

    _rules_cache = data.get("rules", [])
    return _rules_cache


def classify_row(row: ParsedRow) -> str:
    """
    Classify a parsed row into a bucket based on rules.

    Args:
        row: The parsed row to classify

    Returns:
        Bucket name or "UNMAPPED"
    """
    if not row.task_text:
        return "UNMAPPED"

    # Normalize task text for matching
    task_lower = row.task_text.lower()

    # Load and apply rules
    rules = load_rules()

    for rule in rules:
        if matches_rule(task_lower, rule):
            return rule["bucket"]

    return "UNMAPPED"


def matches_rule(text: str, rule: Dict[str, Any]) -> bool:
    """
    Check if text matches a classification rule.

    Args:
        text: The normalized text to check
        rule: The rule dictionary

    Returns:
        True if the text matches the rule
    """
    # Check all_contains - all terms must be present
    if "all_contains" in rule:
        for term in rule["all_contains"]:
            if term.lower() not in text:
                return False

    # Check any_contains - at least one term must be present
    if "any_contains" in rule:
        found_any = False
        for term in rule["any_contains"]:
            if term.lower() in text:
                found_any = True
                break
        if not found_any:
            return False

    # Check none_contains - none of these terms should be present
    if "none_contains" in rule:
        for term in rule["none_contains"]:
            if term.lower() in text:
                return False

    # If we've passed all checks, it's a match
    return True


def classify_rows(rows: List[ParsedRow]) -> Dict[str, List[ParsedRow]]:
    """
    Classify multiple rows into buckets.

    Args:
        rows: List of parsed rows

    Returns:
        Dictionary mapping bucket names to lists of rows
    """
    buckets: Dict[str, List[ParsedRow]] = {}

    for row in rows:
        bucket = classify_row(row)
        if bucket not in buckets:
            buckets[bucket] = []
        buckets[bucket].append(row)

    return buckets