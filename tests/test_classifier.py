"""Tests for the task classification service."""
import pytest
from datetime import datetime
from app.models.schemas import ParsedRow
from app.services.classifier import classify_row, matches_rule


class TestTaskClassification:
    """Test suite for task classification."""

    def test_ext_prime_classification(self):
        """Test that exterior prime tasks are classified correctly."""
        row = ParsedRow(
            lot_block="101",
            plan="A",
            task_text="Painting - Exterior Stucco Prime",
            total=1500.0
        )
        assert classify_row(row) == "EXT PRIME"

        # Alternative format
        row2 = ParsedRow(
            lot_block="102",
            plan="B",
            task_text="EXTERIOR PRIME COATING",
            total=1200.0
        )
        assert classify_row(row2) == "EXT PRIME"

    def test_exterior_ua_classification(self):
        """Test that exterior UA tasks are classified correctly."""
        row = ParsedRow(
            lot_block="201",
            plan="C",
            task_text="Painting - Exterior [UA]",
            total=2000.0
        )
        assert classify_row(row) == "EXTERIOR UA"

        # Alternative format without brackets
        row2 = ParsedRow(
            lot_block="202",
            plan="D",
            task_text="Exterior UA Coating Application",
            total=1800.0
        )
        assert classify_row(row2) == "EXTERIOR UA"

    def test_interior_classification(self):
        """Test that interior tasks are classified correctly."""
        row = ParsedRow(
            lot_block="301",
            plan="E",
            task_text="Painting - Interior Walls",
            total=3000.0
        )
        assert classify_row(row) == "INTERIOR"

        # Different interior task
        row2 = ParsedRow(
            lot_block="302",
            plan="F",
            task_text="INTERIOR TRIM WORK",
            total=2500.0
        )
        assert classify_row(row2) == "INTERIOR"

    def test_extere_classification(self):
        """Test that general exterior tasks are classified correctly."""
        row = ParsedRow(
            lot_block="401",
            plan="G",
            task_text="Painting - Exterior",
            total=1000.0
        )
        assert classify_row(row) == "EXTERE"

        # Exterior but not prime or UA
        row2 = ParsedRow(
            lot_block="402",
            plan="H",
            task_text="Exterior General Coating",
            total=1100.0
        )
        assert classify_row(row2) == "EXTERE"

    def test_unmapped_classification(self):
        """Test that unmappable tasks are classified as UNMAPPED."""
        row = ParsedRow(
            lot_block="501",
            plan="I",
            task_text="Random Task Description",
            total=500.0
        )
        assert classify_row(row) == "UNMAPPED"

        # Another unmapped example
        row2 = ParsedRow(
            lot_block="502",
            plan="J",
            task_text="Miscellaneous Work",
            total=750.0
        )
        assert classify_row(row2) == "UNMAPPED"

    def test_empty_task_text(self):
        """Test that rows with empty task text are classified as UNMAPPED."""
        row = ParsedRow(
            lot_block="601",
            plan="K",
            task_text=None,
            total=100.0
        )
        assert classify_row(row) == "UNMAPPED"

        row2 = ParsedRow(
            lot_block="602",
            plan="L",
            task_text="",
            total=200.0
        )
        assert classify_row(row2) == "UNMAPPED"

    def test_case_insensitive_matching(self):
        """Test that classification is case-insensitive."""
        row_lower = ParsedRow(
            lot_block="701",
            plan="M",
            task_text="painting - exterior prime",
            total=1500.0
        )
        assert classify_row(row_lower) == "EXT PRIME"

        row_upper = ParsedRow(
            lot_block="702",
            plan="N",
            task_text="PAINTING - EXTERIOR PRIME",
            total=1500.0
        )
        assert classify_row(row_upper) == "EXT PRIME"

        row_mixed = ParsedRow(
            lot_block="703",
            plan="O",
            task_text="PaInTiNg - ExTeRiOr PrImE",
            total=1500.0
        )
        assert classify_row(row_mixed) == "EXT PRIME"

    def test_painting_specific_rules(self):
        """Test the painting-specific classification rules."""
        # Painting ext prime
        row1 = ParsedRow(
            lot_block="801",
            plan="P",
            task_text="Painting Ext Prime Coat",
            total=1600.0
        )
        assert classify_row(row1) == "EXT PRIME"

        # Painting ext UA
        row2 = ParsedRow(
            lot_block="802",
            plan="Q",
            task_text="Painting Ext UA Application",
            total=1700.0
        )
        assert classify_row(row2) == "EXTERIOR UA"

        # Painting int
        row3 = ParsedRow(
            lot_block="803",
            plan="R",
            task_text="Painting Int Walls",
            total=1800.0
        )
        assert classify_row(row3) == "INTERIOR"

        # Painting ext (general)
        row4 = ParsedRow(
            lot_block="804",
            plan="S",
            task_text="Painting Ext Finish",
            total=1900.0
        )
        assert classify_row(row4) == "EXTERE"

    def test_rule_priority(self):
        """Test that rules are applied in the correct priority order."""
        # Should match EXT PRIME before EXTERE (more specific rule first)
        row = ParsedRow(
            lot_block="901",
            plan="T",
            task_text="Exterior Prime Application",
            total=2100.0
        )
        assert classify_row(row) == "EXT PRIME"

        # Should match EXTERIOR UA before EXTERE
        row2 = ParsedRow(
            lot_block="902",
            plan="U",
            task_text="Exterior [UA] Coating",
            total=2200.0
        )
        assert classify_row(row2) == "EXTERIOR UA"


class TestMatchesRule:
    """Test suite for the rule matching logic."""

    def test_all_contains(self):
        """Test all_contains rule matching."""
        rule = {
            "bucket": "TEST",
            "all_contains": ["exterior", "prime"]
        }
        assert matches_rule("exterior prime coating", rule) == True
        assert matches_rule("prime exterior work", rule) == True
        assert matches_rule("exterior work", rule) == False  # Missing "prime"
        assert matches_rule("prime coating", rule) == False  # Missing "exterior"

    def test_any_contains(self):
        """Test any_contains rule matching."""
        rule = {
            "bucket": "TEST",
            "any_contains": ["interior", "inside"]
        }
        assert matches_rule("interior walls", rule) == True
        assert matches_rule("inside painting", rule) == True
        assert matches_rule("interior inside work", rule) == True
        assert matches_rule("exterior walls", rule) == False

    def test_none_contains(self):
        """Test none_contains rule matching."""
        rule = {
            "bucket": "TEST",
            "all_contains": ["exterior"],
            "none_contains": ["prime", "ua"]
        }
        assert matches_rule("exterior coating", rule) == True
        assert matches_rule("exterior prime", rule) == False  # Contains "prime"
        assert matches_rule("exterior ua work", rule) == False  # Contains "ua"

    def test_combined_rules(self):
        """Test combinations of rule types."""
        rule = {
            "bucket": "TEST",
            "all_contains": ["painting", "exterior"],
            "none_contains": ["prime", "[ua]"],
            "any_contains": ["coat", "finish", "application"]
        }
        assert matches_rule("painting exterior coat", rule) == True
        assert matches_rule("painting exterior finish", rule) == True
        assert matches_rule("painting exterior", rule) == False  # Missing any_contains
        assert matches_rule("painting exterior prime coat", rule) == False  # Contains "prime"
        assert matches_rule("painting interior coat", rule) == False  # Missing "exterior"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])