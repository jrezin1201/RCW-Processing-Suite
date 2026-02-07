"""
Unit tests for the signal-based task classifier.

Tests cover:
- Normalization
- Signal extraction
- Classification rules
- Edge cases
"""

import pytest
from datetime import datetime

from app.services.classifier import (
    normalize_task_text,
    parse_signals,
    classify_task,
    classify_row,
    classify_row_with_details,
    TaskSignals,
    ClassificationResult,
    reload_config,
)
from app.models.schemas import ParsedRow


class TestNormalizeTaskText:
    """Tests for the normalize_task_text function."""

    def test_basic_normalization(self):
        """Test basic text normalization."""
        result = normalize_task_text("Painting - Interior / Prep & Enamel")
        assert result == "PAINTING INTERIOR PREP & ENAMEL"

    def test_preserves_markers(self):
        """Test that parentheses and brackets are preserved."""
        result = normalize_task_text("Painting - Exterior (437225) (EXT) [LS]")
        assert "(437225)" in result
        assert "(EXT)" in result
        assert "[LS]" in result

    def test_handles_empty_string(self):
        """Test handling of empty string."""
        assert normalize_task_text("") == ""

    def test_handles_none(self):
        """Test handling of None (should not raise)."""
        assert normalize_task_text(None) == ""

    def test_removes_multiple_spaces(self):
        """Test that multiple spaces are collapsed."""
        result = normalize_task_text("Painting   -   Interior")
        assert "  " not in result

    def test_normalizes_separators(self):
        """Test that various separators are normalized to spaces."""
        result = normalize_task_text("Paint_task-name—test/value")
        # All separators should become spaces
        assert "_" not in result
        assert "-" not in result
        assert "—" not in result
        assert "/" not in result


class TestParseSignals:
    """Tests for the parse_signals function."""

    def test_detects_ext_marker(self):
        """Test detection of (EXT) marker."""
        signals = parse_signals("Painting - Exterior (437225) (EXT) [LS]")
        assert signals.is_exterior is True
        assert "(EXT) marker" in signals.matched_patterns

    def test_detects_int_marker(self):
        """Test detection of (INT) marker."""
        signals = parse_signals("Painting - Interior (437205) (INT) [LS]")
        assert signals.is_interior is True
        assert "(INT) marker" in signals.matched_patterns

    def test_detects_ua_marker(self):
        """Test detection of [UA] marker."""
        signals = parse_signals("Painting - Exterior (437225) (EXT) [UA]")
        assert signals.is_ua is True
        assert "[UA] marker" in signals.matched_patterns

    def test_detects_op_marker(self):
        """Test detection of [OP] marker."""
        signals = parse_signals("Painting - Exterior (437225) (EXT) [LS] [OP]")
        assert signals.is_op is True

    def test_detects_ls_marker(self):
        """Test detection of [LS] marker."""
        signals = parse_signals("Painting - Exterior (437225) (EXT) [LS]")
        assert signals.is_ls is True

    def test_detects_job_code(self):
        """Test detection of job code pattern."""
        signals = parse_signals("Painting - Exterior (437225) (EXT)")
        assert signals.has_job_code is True
        assert signals.job_code == "437225"

    def test_detects_prime_keywords(self):
        """Test detection of prime/prep keywords."""
        # Test PRIME
        signals = parse_signals("Painting - Exterior Prime/Fascia (437220) (EXT)")
        assert signals.keyword_prime is True

        # Test FASCIA
        signals = parse_signals("Painting - Exterior/Fascia (437220) (EXT)")
        assert signals.keyword_prime is True

        # Test PREP
        signals = parse_signals("Painting - Prep & Enamel (437205) (INT)")
        assert signals.keyword_prime is True

    def test_detects_touchup_keywords(self):
        """Test detection of touch-up keywords."""
        # Test TOUCH UP
        signals = parse_signals("Painting - Touch Up (437230) (INT)")
        assert signals.keyword_touchup is True

        # Test AFTER CARPET
        signals = parse_signals("Painting - After Carpet Touch-Up (INT)")
        assert signals.keyword_touchup is True

        # Test PUNCH
        signals = parse_signals("Painting - Punch List (INT)")
        assert signals.keyword_touchup is True

    def test_detects_rollwalls_keywords(self):
        """Test detection of roll walls keywords."""
        signals = parse_signals("Painting - Roll Walls (437240) (INT)")
        assert signals.keyword_rollwalls is True

    def test_parses_date(self):
        """Test date parsing."""
        # Test datetime object
        dt = datetime(2026, 3, 20)
        signals = parse_signals("Painting - Exterior (EXT)", dt)
        assert signals.raw_date == "2026-03-20"
        assert signals.date_month == 3

        # Test string date
        signals = parse_signals("Painting - Exterior (EXT)", "03/20/2026")
        assert signals.raw_date == "2026-03-20"
        assert signals.date_month == 3

    def test_exterior_keyword_detection(self):
        """Test detection of EXTERIOR keyword (not just marker)."""
        signals = parse_signals("Painting - Exterior Work")
        assert signals.is_exterior is True
        assert "EXTERIOR keyword" in signals.matched_patterns


class TestClassifyTask:
    """Tests for the classify_task function."""

    def test_ext_prime_classification(self):
        """Test EXT PRIME classification."""
        task = "Painting - Exterior Prime/Fascia (437220) (EXT) [LS] [578700 - 34749538-000] [OP]"
        signals = parse_signals(task)
        normalized = normalize_task_text(task)
        result = classify_task(signals, normalized, task)

        assert result.category == "EXT PRIME"
        assert result.rule_fired == "rule_b_ext_prime"

    def test_exterior_ua_classification(self):
        """Test EXTERIOR UA classification."""
        task = "Painting - Exterior (437225) (EXT) [578700 - 43468818-000] [UA]"
        signals = parse_signals(task)
        normalized = normalize_task_text(task)
        result = classify_task(signals, normalized, task)

        assert result.category == "EXTERIOR UA"
        assert result.rule_fired == "rule_c_exterior_ua"

    def test_exterior_classification(self):
        """Test EXTERIOR classification (with [OP])."""
        task = "Painting - Exterior (437225) (EXT) [LS] [578700 - 34749546-000] [OP]"
        signals = parse_signals(task)
        normalized = normalize_task_text(task)
        result = classify_task(signals, normalized, task)

        assert result.category == "EXTERIOR"
        assert result.rule_fired == "rule_d_exterior"

    def test_interior_classification(self):
        """Test INTERIOR classification."""
        task = "Painting - Kitchen, Bath & Lids (INT) [LS]"
        signals = parse_signals(task)
        normalized = normalize_task_text(task)
        result = classify_task(signals, normalized, task)

        assert result.category == "INTERIOR"
        assert result.rule_fired == "rule_g_interior"

    def test_interior_prep_enamel(self):
        """Test INTERIOR classification for Prep & Enamel tasks."""
        task = "Painting - Interior / Prep & Enamel (437205) (INT) [LS] [578700 - 34749520-000] [OP]"
        signals = parse_signals(task)
        normalized = normalize_task_text(task)
        result = classify_task(signals, normalized, task)

        # This has (INT) so should be INTERIOR, even though it has PREP keyword
        # PREP keyword is for prime, but rule order puts EXT PRIME first which requires is_exterior
        assert result.category == "INTERIOR"

    def test_touch_up_classification(self):
        """Test TOUCH UP classification."""
        task = "Painting - After Carpet Touch-Up (INT) [LS]"
        signals = parse_signals(task)
        normalized = normalize_task_text(task)
        result = classify_task(signals, normalized, task)

        assert result.category == "TOUCH UP"
        assert result.rule_fired == "rule_e_touch_up"

    def test_roll_walls_classification(self):
        """Test ROLL WALLS FINAL classification."""
        task = "Painting - Roll Walls (437240) (INT) [LS]"
        signals = parse_signals(task)
        normalized = normalize_task_text(task)
        result = classify_task(signals, normalized, task)

        assert result.category == "ROLL WALLS FINAL"
        assert result.rule_fired == "rule_f_roll_walls_final"

    def test_unmapped_fallback(self):
        """Test UNMAPPED fallback for unrecognized tasks."""
        task = "Random Task Description"
        signals = parse_signals(task)
        normalized = normalize_task_text(task)
        result = classify_task(signals, normalized, task)

        assert result.category == "UNMAPPED"
        assert result.rule_fired == "rule_h_unmapped"

    def test_rule_order_ext_prime_before_exterior(self):
        """Test that EXT PRIME takes precedence over EXTERIOR."""
        # Task with both exterior and prime should be EXT PRIME
        task = "Painting - Exterior Prime (EXT)"
        signals = parse_signals(task)
        normalized = normalize_task_text(task)
        result = classify_task(signals, normalized, task)

        assert result.category == "EXT PRIME"

    def test_rule_order_exterior_ua_before_exterior(self):
        """Test that EXTERIOR UA takes precedence over EXTERIOR."""
        task = "Painting - Exterior (EXT) [UA]"
        signals = parse_signals(task)
        normalized = normalize_task_text(task)
        result = classify_task(signals, normalized, task)

        assert result.category == "EXTERIOR UA"


class TestClassifyRow:
    """Tests for the classify_row function with ParsedRow objects."""

    def test_classify_row_basic(self):
        """Test basic row classification."""
        row = ParsedRow(
            lot_block="0044/",
            plan="2",
            task_text="Exterior Prime/Fascia",
            task_text_raw="Painting - Exterior Prime/Fascia (437220) (EXT) [LS]",
            task_start_date=datetime(2026, 3, 13)
        )
        result = classify_row(row)
        assert result == "EXT PRIME"

    def test_classify_row_uses_raw_text(self):
        """Test that raw task text is preferred for classification."""
        row = ParsedRow(
            lot_block="0044/",
            plan="2",
            task_text="Exterior",  # Normalized version
            task_text_raw="Painting - Exterior (437225) (EXT) [UA]",  # Full version with signals
            task_start_date=datetime(2026, 3, 20)
        )
        result = classify_row(row)
        # Should use raw text and detect [UA]
        assert result == "EXTERIOR UA"

    def test_classify_row_empty_task(self):
        """Test classification of row with empty task."""
        row = ParsedRow(
            lot_block="0044/",
            plan="2",
            task_text=None,
            task_text_raw=None
        )
        result = classify_row(row)
        assert result == "UNMAPPED"

    def test_classify_row_with_details(self):
        """Test classify_row_with_details returns full information."""
        row = ParsedRow(
            lot_block="0044/",
            plan="2",
            task_text_raw="Painting - Kitchen, Bath & Lids (INT) [LS]",
            task_start_date=datetime(2026, 3, 30)
        )
        result = classify_row_with_details(row)

        assert isinstance(result, ClassificationResult)
        assert result.category == "INTERIOR"
        assert result.signals.is_interior is True
        assert result.signals.raw_date == "2026-03-30"


class TestRealWorldSamples:
    """Tests using real-world task samples from the data."""

    @pytest.fixture(autouse=True)
    def setup(self):
        """Reload config before each test."""
        reload_config()

    def test_sample_exterior_prime_fascia(self):
        """Test: Painting - Exterior Prime/Fascia (437220) (EXT) [LS] [578700 - 34749538-000] [OP]"""
        task = "Painting - Exterior Prime/Fascia (437220) (EXT) [LS] [578700 - 34749538-000] [OP]"
        row = ParsedRow(task_text_raw=task)
        assert classify_row(row) == "EXT PRIME"

    def test_sample_exterior_ua(self):
        """Test: Painting - Exterior (437225) (EXT) [578700 - 43468818-000] [UA]"""
        task = "Painting - Exterior (437225) (EXT) [578700 - 43468818-000] [UA]"
        row = ParsedRow(task_text_raw=task)
        assert classify_row(row) == "EXTERIOR UA"

    def test_sample_exterior_op(self):
        """Test: Painting - Exterior (437225) (EXT) [LS] [578700 - 34749546-000] [OP]"""
        task = "Painting - Exterior (437225) (EXT) [LS] [578700 - 34749546-000] [OP]"
        row = ParsedRow(task_text_raw=task)
        assert classify_row(row) == "EXTERIOR"

    def test_sample_interior_kitchen_bath(self):
        """Test: Painting - Kitchen, Bath & Lids (INT) [LS]"""
        task = "Painting - Kitchen, Bath & Lids (INT) [LS]"
        row = ParsedRow(task_text_raw=task)
        assert classify_row(row) == "INTERIOR"

    def test_sample_interior_prep_enamel(self):
        """Test: Painting - Interior / Prep & Enamel (437205) (INT) [LS] [578700 - 34749520-000] [OP]"""
        task = "Painting - Interior / Prep & Enamel (437205) (INT) [LS] [578700 - 34749520-000] [OP]"
        row = ParsedRow(task_text_raw=task)
        assert classify_row(row) == "INTERIOR"

    def test_sample_base_shoe(self):
        """Test: Painting - Base Shoe (Flooring Orders) (455240) (INT) [LS]"""
        task = "Painting - Base Shoe (Flooring Orders) (455240) (INT) [LS]"
        row = ParsedRow(task_text_raw=task)
        assert classify_row(row) == "INTERIOR"

    def test_sample_exterior_prep_enamel(self):
        """Test: Painting - Exterior / Prep & Enamel (437205) (INT) [LS]"""
        # Note: This has (INT) marker despite "Exterior" in name
        task = "Painting - Exterior / Prep & Enamel (437205) (INT) [LS]"
        row = ParsedRow(task_text_raw=task)
        # (INT) marker should take precedence, making it INTERIOR
        result = classify_row(row)
        assert result == "INTERIOR"

    def test_sample_touch_up(self):
        """Test: Painting - Touch Up (437230) (INT) [LS]"""
        task = "Painting - Touch Up (437230) (INT) [LS]"
        row = ParsedRow(task_text_raw=task)
        assert classify_row(row) == "TOUCH UP"

    def test_sample_after_carpet_touchup(self):
        """Test: Painting - After Carpet Touch-Up (INT)"""
        task = "Painting - After Carpet Touch-Up (INT)"
        row = ParsedRow(task_text_raw=task)
        assert classify_row(row) == "TOUCH UP"


class TestEdgeCases:
    """Tests for edge cases and unusual inputs."""

    def test_case_insensitivity(self):
        """Test that classification is case-insensitive."""
        tasks = [
            "painting - exterior prime (ext)",
            "PAINTING - EXTERIOR PRIME (EXT)",
            "Painting - Exterior Prime (EXT)",
        ]
        for task in tasks:
            row = ParsedRow(task_text_raw=task)
            assert classify_row(row) == "EXT PRIME"

    def test_extra_whitespace(self):
        """Test handling of extra whitespace."""
        task = "Painting  -  Exterior   Prime   (EXT)"
        row = ParsedRow(task_text_raw=task)
        assert classify_row(row) == "EXT PRIME"

    def test_no_markers(self):
        """Test task with no standard markers."""
        task = "Painting - General Work"
        row = ParsedRow(task_text_raw=task)
        assert classify_row(row) == "UNMAPPED"

    def test_conflicting_markers(self):
        """Test task with potentially conflicting markers."""
        # Has both EXTERIOR in name but (INT) marker - marker should win
        task = "Painting - Exterior Style Interior (INT)"
        signals = parse_signals(task)
        # (INT) marker is most reliable
        assert signals.is_interior is True

    def test_audit_trail(self):
        """Test that audit trail is properly recorded."""
        task = "Painting - Exterior Prime/Fascia (437220) (EXT) [LS] [OP]"
        row = ParsedRow(task_text_raw=task)
        result = classify_row_with_details(row)

        # Check signals have matched patterns
        assert len(result.signals.matched_patterns) > 0
        assert "(EXT) marker" in result.signals.matched_patterns
        assert "[LS] marker" in result.signals.matched_patterns
        assert "[OP] marker" in result.signals.matched_patterns

        # Check result has debug info
        assert result.rule_fired == "rule_b_ext_prime"
