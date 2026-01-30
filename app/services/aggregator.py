"""Data aggregation service for summarizing classified rows."""
from typing import List, Dict, Tuple, Any
from collections import defaultdict, Counter
import re

from app.models.schemas import ParsedRow, SummaryRow, QAReport, QAMeta
from app.services.classifier import classify_row


def clean_lot_number(lot: str) -> str:
    """
    Clean lot number by removing leading zeros and trailing slashes.

    Examples:
        "0044/" -> "44"
        "0143/" -> "143"
        "101" -> "101"
    """
    if not lot:
        return ""

    # Remove trailing slashes and spaces
    lot = lot.rstrip('/ ')

    # Remove leading zeros but keep at least one digit
    lot = lot.lstrip('0') or '0'

    return lot


def combine_plan_elevation(plan: str, elevation: str) -> str:
    """
    Combine plan and elevation into a single string.

    Examples:
        plan="2", elevation="B" -> "2B"
        plan="3", elevation="A" -> "3A"
        plan="1", elevation=None -> "1"
    """
    if not plan:
        plan = ""

    # Clean up the plan (remove extra spaces)
    plan = str(plan).strip()

    # If elevation exists and is not empty, append it
    if elevation and str(elevation).strip():
        elevation = str(elevation).strip()
        # Don't add elevation if it's already part of the plan
        if not plan.endswith(elevation):
            plan = f"{plan}{elevation}"

    return plan


def aggregate_data(
    rows: List[ParsedRow],
    qa_meta: QAMeta
) -> Tuple[List[SummaryRow], QAReport]:
    """
    Aggregate classified rows by lot/plan and create summary.

    Args:
        rows: List of parsed rows
        qa_meta: QA metadata from parsing

    Returns:
        Tuple of (summary rows, QA report)
    """
    # Group by (lot_block, plan) and bucket
    aggregated_data: Dict[Tuple[str, str], Dict[str, float]] = defaultdict(
        lambda: {
            "EXT PRIME": 0.0,
            "EXTERE": 0.0,
            "EXTERIOR UA": 0.0,
            "INTERIOR": 0.0,
        }
    )

    # Track the order in which lot/plan combinations first appear
    appearance_order: Dict[Tuple[str, str], int] = {}

    # Track statistics
    counts_per_bucket: Dict[str, int] = defaultdict(int)
    unmapped_tasks: List[str] = []

    # Classify and aggregate rows
    for row in rows:
        bucket = classify_row(row)
        counts_per_bucket[bucket] += 1

        if bucket == "UNMAPPED":
            if row.task_text:
                unmapped_tasks.append(row.task_text)
            continue

        # Get the amount to aggregate (prefer total, fallback to subtotal)
        amount = row.total if row.total is not None else row.subtotal
        if amount is None:
            amount = 0.0

        # Clean lot number and combine plan with elevation
        cleaned_lot = clean_lot_number(row.lot_block or "")
        combined_plan = combine_plan_elevation(row.plan or "", row.elevation or "")

        # Create group key with cleaned values
        group_key = (cleaned_lot, combined_plan)

        # Track the order of first appearance
        if group_key not in appearance_order:
            appearance_order[group_key] = len(appearance_order)

        # Add to the appropriate bucket
        if bucket in aggregated_data[group_key]:
            aggregated_data[group_key][bucket] += amount

    # Create summary rows
    summary_rows = []
    suspicious_totals = []

    for (lot_block, plan), buckets in aggregated_data.items():
        total = sum(buckets.values())

        # Check for suspicious totals
        if total < 0:
            suspicious_totals.append({
                "lot_block": lot_block,
                "plan": plan,
                "total": total,
                "reason": "Negative total"
            })
        elif total > 100000:
            suspicious_totals.append({
                "lot_block": lot_block,
                "plan": plan,
                "total": total,
                "reason": "Unusually high total (> $100k)"
            })

        summary_row = SummaryRow(
            lot_block=lot_block,
            plan=plan,
            ext_prime=buckets["EXT PRIME"],
            extere=buckets["EXTERE"],
            exterior_ua=buckets["EXTERIOR UA"],
            interior=buckets["INTERIOR"],
            total=total
        )
        summary_rows.append(summary_row)

    # Sort summary rows by their original appearance order in the input
    summary_rows.sort(key=lambda x: appearance_order.get((x.lot_block, x.plan), float('inf')))

    # Get top unmapped examples
    unmapped_counter = Counter(unmapped_tasks)
    unmapped_examples = [
        {"task_text": task, "count": count}
        for task, count in unmapped_counter.most_common(30)
    ]

    # Create QA report
    qa_report = QAReport(
        counts_per_bucket=dict(counts_per_bucket),
        unmapped_examples=unmapped_examples,
        suspicious_totals=suspicious_totals,
        parse_meta=qa_meta
    )

    return summary_rows, qa_report