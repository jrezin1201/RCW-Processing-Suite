"""Data aggregation service for summarizing classified rows."""
from typing import List, Dict, Tuple, Any
from collections import defaultdict, Counter
import re
import logging

from app.models.schemas import ParsedRow, SummaryRow, QAReport, QAMeta
from app.services.category_mapper import CategoryMapper, MappingResult, organize_headers

logger = logging.getLogger(__name__)


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
    qa_meta: QAMeta,
    template_headers: List[str] = None
) -> Tuple[List[Dict[str, Any]], QAReport, List[str]]:
    """
    Aggregate classified rows by lot/plan and create summary.

    Uses template-first category mapping. If a task doesn't map to any
    template category, creates a new category automatically.
    NO DOLLARS ARE LOST.

    Args:
        rows: List of parsed rows
        qa_meta: QA metadata from parsing
        template_headers: Optional list of category headers from template

    Returns:
        Tuple of (summary rows as dicts, QA report, final category headers)
    """
    # Initialize category mapper with template headers
    mapper = CategoryMapper(template_headers)

    # Group by (lot_block, plan) and category
    # Using dict of dicts for flexible category columns
    aggregated_data: Dict[Tuple[str, str], Dict[str, float]] = defaultdict(
        lambda: defaultdict(float)
    )

    # Track the order in which lot/plan combinations first appear
    appearance_order: Dict[Tuple[str, str], int] = {}

    # Track statistics
    counts_per_category: Dict[str, int] = defaultdict(int)
    mapping_details: List[Dict[str, Any]] = []
    unmapped_tasks: List[str] = []  # For backward compatibility

    # Process each row
    for row in rows:
        # Get task text (prefer raw for full signal extraction)
        task_text = row.task_text_raw or row.task_text

        if not task_text:
            counts_per_category["UNMAPPED"] += 1
            unmapped_tasks.append("(empty task)")
            continue

        # Map the task to a category
        result: MappingResult = mapper.map_task(task_text)
        category = result.category_display

        counts_per_category[category] += 1

        # Track mapping details for QA
        mapping_details.append({
            "lot": row.lot_block,
            "plan": row.plan,
            "task_text": task_text,
            "category": category,
            "reason": result.reason,
            "is_new_category": result.is_new_category,
            "signals": result.signals.to_dict() if result.signals else {}
        })

        # If this is a created category, add example
        if result.is_new_category:
            mapper.add_example_to_created_category(category, task_text)

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

        # Add to the category
        aggregated_data[group_key][category] += amount

    # Get final category headers — only include categories that actually have data
    active_headers = [h for h in mapper.get_category_headers() if counts_per_category.get(h, 0) > 0]
    # Organize so UA variants are adjacent to their base category
    final_headers = organize_headers(active_headers)

    # Create summary rows as dicts (flexible columns)
    summary_rows = []
    suspicious_totals = []

    for (lot_block, plan), categories in aggregated_data.items():
        total = sum(categories.values())

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

        # Create row dict with all category columns
        row_dict = {
            "lot_block": lot_block,
            "plan": plan,
        }

        # Add all category values
        for header in final_headers:
            row_dict[header] = categories.get(header, 0.0)

        row_dict["total"] = total

        summary_rows.append(row_dict)

    # Sort summary rows by their original appearance order in the input
    summary_rows.sort(key=lambda x: appearance_order.get(
        (x["lot_block"], x["plan"]), float('inf')
    ))

    # Get top unmapped examples (now includes auto-created categories info)
    unmapped_counter = Counter(unmapped_tasks)
    unmapped_examples = [
        {"task_text": task, "count": count}
        for task, count in unmapped_counter.most_common(30)
    ]

    # Add created categories to unmapped_examples for visibility
    created_report = mapper.get_created_categories_report()
    if created_report:
        for cat in created_report:
            unmapped_examples.append({
                "task_text": f"[AUTO-CREATED] {cat['header']}",
                "count": counts_per_category.get(cat['header'], 0),
                "examples": cat.get("example_tasks", [])
            })

    # Create QA report
    qa_report = QAReport(
        counts_per_bucket=dict(counts_per_category),
        unmapped_examples=unmapped_examples,
        suspicious_totals=suspicious_totals,
        parse_meta=qa_meta
    )

    # Log summary
    logger.info(f"Aggregation complete: {len(summary_rows)} lot/plan combinations")
    logger.info(f"Categories used: {list(counts_per_category.keys())}")
    if created_report:
        logger.info(f"Auto-created {len(created_report)} new categories:")
        for cat in created_report:
            logger.info(f"  - {cat['header']}: {counts_per_category.get(cat['header'], 0)} rows")

    return summary_rows, qa_report, final_headers


def aggregate_data_legacy(
    rows: List[ParsedRow],
    qa_meta: QAMeta
) -> Tuple[List[SummaryRow], QAReport]:
    """
    Legacy aggregation for backward compatibility.

    Uses fixed category columns. Prefer aggregate_data() for new code.
    """
    from app.services.classifier import classify_row

    # Fixed categories for legacy mode
    aggregated_data: Dict[Tuple[str, str], Dict[str, float]] = defaultdict(
        lambda: {
            "EXT PRIME": 0.0,
            "EXTERIOR": 0.0,
            "EXTERIOR UA": 0.0,
            "INTERIOR": 0.0,
            "ROLL WALLS FINAL": 0.0,
            "TOUCH UP": 0.0,
            "Q4 REVERSAL": 0.0,
        }
    )

    appearance_order: Dict[Tuple[str, str], int] = {}
    counts_per_bucket: Dict[str, int] = defaultdict(int)
    unmapped_tasks: List[str] = []

    for row in rows:
        bucket = classify_row(row)
        counts_per_bucket[bucket] += 1

        if bucket == "UNMAPPED":
            if row.task_text:
                unmapped_tasks.append(row.task_text)
            continue

        amount = row.total if row.total is not None else row.subtotal
        if amount is None:
            amount = 0.0

        cleaned_lot = clean_lot_number(row.lot_block or "")
        combined_plan = combine_plan_elevation(row.plan or "", row.elevation or "")
        group_key = (cleaned_lot, combined_plan)

        if group_key not in appearance_order:
            appearance_order[group_key] = len(appearance_order)

        if bucket in aggregated_data[group_key]:
            aggregated_data[group_key][bucket] += amount

    summary_rows = []
    suspicious_totals = []

    for (lot_block, plan), buckets in aggregated_data.items():
        total = sum(buckets.values())

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
            exterior=buckets["EXTERIOR"],
            exterior_ua=buckets["EXTERIOR UA"],
            interior=buckets["INTERIOR"],
            roll_walls_final=buckets["ROLL WALLS FINAL"],
            touch_up=buckets["TOUCH UP"],
            q4_reversal=buckets["Q4 REVERSAL"],
            total=total
        )
        summary_rows.append(summary_row)

    summary_rows.sort(key=lambda x: appearance_order.get(
        (x.lot_block, x.plan), float('inf')
    ))

    unmapped_counter = Counter(unmapped_tasks)
    unmapped_examples = [
        {"task_text": task, "count": count}
        for task, count in unmapped_counter.most_common(30)
    ]

    qa_report = QAReport(
        counts_per_bucket=dict(counts_per_bucket),
        unmapped_examples=unmapped_examples,
        suspicious_totals=suspicious_totals,
        parse_meta=qa_meta
    )

    return summary_rows, qa_report
