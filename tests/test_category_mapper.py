"""
Tests for the category mapper with template-first approach and auto-creation.

Key test scenarios:
- Known mappings go to existing template columns
- Unmapped tasks create new columns automatically
- Repeated unmapped tasks reuse the same new column
- No dollars are lost
"""

import pytest
from app.services.category_mapper import (
    normalize_task_text,
    extract_scope_fragment,
    parse_signals,
    map_category,
    create_category_name,
    canonical,
    CategoryMapper,
    TaskSignals,
)


class TestCanonical:
    """Tests for the canonical() function."""

    def test_basic(self):
        assert canonical("EXT PRIME") == "EXT PRIME"
        assert canonical("ext prime") == "EXT PRIME"
        assert canonical("  EXT  PRIME  ") == "EXT PRIME"

    def test_strips_punctuation(self):
        assert canonical("--EXT PRIME--") == "EXT PRIME"
        assert canonical("...INTERIOR...") == "INTERIOR"


class TestNormalizeTaskText:
    """Tests for normalize_task_text()."""

    def test_basic(self):
        result = normalize_task_text("Painting - Interior / Prep & Enamel")
        assert result == "PAINTING INTERIOR PREP & ENAMEL"

    def test_preserves_markers(self):
        result = normalize_task_text("Painting - Exterior (437225) (EXT) [LS]")
        assert "(EXT)" in result
        assert "(437225)" in result
        assert "[LS]" in result

    def test_preserves_ampersand(self):
        result = normalize_task_text("Kitchen & Bath")
        assert "&" in result


class TestExtractScopeFragment:
    """Tests for extract_scope_fragment()."""

    def test_removes_painting_prefix(self):
        result = extract_scope_fragment("Painting - Interior Walls (INT)")
        assert "PAINTING" not in result.upper()
        assert "Interior Walls" in result or "INTERIOR WALLS" in result.upper()

    def test_removes_job_codes(self):
        result = extract_scope_fragment("Painting - Exterior (437225) (EXT)")
        assert "437225" not in result

    def test_removes_bracket_codes(self):
        result = extract_scope_fragment("Painting - Exterior [LS] [578700 - 34749538-000] [OP]")
        assert "[LS]" not in result
        assert "[OP]" not in result

    def test_removes_int_ext_markers(self):
        result = extract_scope_fragment("Painting - Interior (INT)")
        assert "(INT)" not in result
        result = extract_scope_fragment("Painting - Exterior (EXT)")
        assert "(EXT)" not in result

    def test_handles_complex_task(self):
        task = "Painting - Exterior Prime/Fascia (437220) (EXT) [LS] [578700 - 34749538-000] [OP]"
        result = extract_scope_fragment(task)
        # Should contain the scope name without codes
        assert "437220" not in result
        assert "[LS]" not in result
        assert "[OP]" not in result


class TestParseSignals:
    """Tests for parse_signals()."""

    def test_ext_marker(self):
        signals = parse_signals("Painting - Exterior (437225) (EXT)")
        assert signals.is_ext is True
        assert signals.is_int is False

    def test_int_marker(self):
        signals = parse_signals("Painting - Interior (437205) (INT)")
        assert signals.is_int is True
        assert signals.is_ext is False

    def test_ua_marker(self):
        signals = parse_signals("Painting - Exterior (EXT) [UA]")
        assert signals.is_ua is True

    def test_prime_keyword(self):
        signals = parse_signals("Painting - Exterior Prime/Fascia (EXT)")
        assert signals.keyword_prime is True

    def test_touchup_keyword(self):
        signals = parse_signals("Painting - Touch Up (INT)")
        assert signals.keyword_touchup is True

        signals = parse_signals("Painting - After Carpet Touch-Up")
        assert signals.keyword_touchup is True

    def test_rollwalls_keyword(self):
        signals = parse_signals("Painting - Roll Walls (INT)")
        assert signals.keyword_rollwalls is True

    def test_baseshoe_keyword(self):
        signals = parse_signals("Painting - Base Shoe (INT)")
        assert signals.keyword_baseshoe is True

    def test_undercoat_keyword(self):
        signals = parse_signals("Painting - Undercoat (INT)")
        assert signals.keyword_undercoat is True


class TestMapCategory:
    """Tests for map_category()."""

    def test_maps_to_ext_prime(self):
        template = {"EXT PRIME": "EXT PRIME", "EXTERIOR": "EXTERIOR"}
        signals = TaskSignals(is_ext=True, keyword_prime=True)
        category, reason = map_category("Exterior Prime", signals, template)
        assert category == "EXT PRIME"

    def test_maps_to_exterior_ua(self):
        template = {"EXTERIOR UA": "EXTERIOR UA", "EXTERIOR": "EXTERIOR"}
        signals = TaskSignals(is_ext=True, is_ua=True)
        category, reason = map_category("Exterior [UA]", signals, template)
        assert category == "EXTERIOR UA"

    def test_maps_to_interior(self):
        template = {"INTERIOR": "INTERIOR"}
        signals = TaskSignals(is_int=True)
        category, reason = map_category("Interior (INT)", signals, template)
        assert category == "INTERIOR"

    def test_returns_none_for_unmapped(self):
        template = {"EXTERIOR": "EXTERIOR"}
        signals = TaskSignals()  # No markers
        category, reason = map_category("Random Task", signals, template)
        assert category is None
        assert reason == "unmapped_template"


class TestCreateCategoryName:
    """Tests for create_category_name()."""

    def test_creates_name_from_fragment(self):
        signals = TaskSignals(is_ext=True)
        name = create_category_name(
            "Painting - Rolling Roof (EXT) [LS]",
            signals,
            set()
        )
        assert "ROLLING ROOF" in name.upper()
        assert "EXT" in name.upper()

    def test_adds_int_prefix(self):
        signals = TaskSignals(is_int=True)
        name = create_category_name(
            "Painting - Backyard Playground (INT) [LS]",
            signals,
            set()
        )
        assert "INT" in name.upper()

    def test_adds_ua_suffix(self):
        signals = TaskSignals(is_ext=True, is_ua=True)
        name = create_category_name(
            "Painting - Custom Work (EXT) [UA]",
            signals,
            set()
        )
        assert "UA" in name.upper()

    def test_deduplicates(self):
        signals = TaskSignals(is_ext=True)
        existing = {"EXT ROLLING ROOF"}
        name = create_category_name(
            "Painting - Rolling Roof (EXT)",
            signals,
            existing
        )
        # Should get a different name since EXT ROLLING ROOF exists
        assert canonical(name) != "EXT ROLLING ROOF"


class TestCategoryMapper:
    """Tests for the CategoryMapper class."""

    def test_maps_known_categories(self):
        mapper = CategoryMapper()
        result = mapper.map_task("Painting - Exterior Prime/Fascia (437220) (EXT) [LS]")
        assert result.category_display == "EXT PRIME"
        assert result.is_new_category is False

    def test_maps_exterior_ua(self):
        mapper = CategoryMapper()
        result = mapper.map_task("Painting - Exterior (437225) (EXT) [UA]")
        assert result.category_display == "EXTERIOR UA"

    def test_maps_interior(self):
        mapper = CategoryMapper()
        # Generic interior scope maps to INTERIOR
        result = mapper.map_task("Painting - Interior / Prep & Enamel (437205) (INT) [LS]")
        assert result.category_display == "INTERIOR"

    def test_distinctive_interior_auto_creates(self):
        mapper = CategoryMapper()
        # Distinctive scope auto-creates its own column
        result = mapper.map_task("Painting - Kitchen, Bath & Lids (INT) [LS]")
        assert result.is_new_category is True
        assert "KITCHEN" in result.category_display.upper()

    def test_auto_creates_category(self):
        mapper = CategoryMapper()
        result = mapper.map_task("Painting - Rolling Roof (EXT) [LS]")
        # Should auto-create since "Rolling Roof" doesn't match any template category
        assert result.is_new_category is True
        assert "ROLLING ROOF" in result.category_display.upper()

    def test_reuses_created_category(self):
        mapper = CategoryMapper()
        # First occurrence
        result1 = mapper.map_task("Painting - Rolling Roof (EXT) [LS]")
        assert result1.is_new_category is True
        cat1 = result1.category_display

        # Second occurrence - should reuse
        result2 = mapper.map_task("Painting - Rolling Roof (EXT) [OP]")
        assert result2.category_display == cat1
        # Second time should not be marked as new
        assert result2.is_new_category is False or result2.reason == "reused_created_category"

    def test_tracks_created_categories(self):
        mapper = CategoryMapper()
        mapper.map_task("Painting - Rolling Roof (EXT) [LS]")
        mapper.map_task("Painting - Custom Work (INT) [LS]")

        report = mapper.get_created_categories_report()
        assert len(report) >= 2

    def test_no_dollars_lost(self):
        """Every task should map to SOME category."""
        mapper = CategoryMapper()

        test_tasks = [
            "Painting - Exterior Prime/Fascia (EXT)",
            "Painting - Interior (INT)",
            "Painting - Rolling Roof (EXT)",  # Unmapped
            "Painting - Custom Special Work",  # No markers
            "Random Task Description",  # Completely unknown
        ]

        for task in test_tasks:
            result = mapper.map_task(task)
            assert result.category_display is not None
            assert len(result.category_display) > 0


class TestRealWorldTasks:
    """Tests with real-world task examples."""

    def test_sample_tasks(self):
        mapper = CategoryMapper()

        # Known mappings
        test_cases = [
            ("Painting - Exterior Prime/Fascia (437220) (EXT) [LS] [578700 - 34749538-000] [OP]", "EXT PRIME"),
            ("Painting - Exterior (437225) (EXT) [578700 - 43468818-000] [UA]", "EXTERIOR UA"),
            ("Painting - Exterior (437225) (EXT) [LS] [578700 - 34749546-000] [OP]", "EXTERIOR"),
            ("Painting - Interior / Prep & Enamel (437205) (INT) [LS]", "INTERIOR"),
        ]

        for task, expected in test_cases:
            result = mapper.map_task(task)
            assert result.category_display == expected, f"Task: {task[:50]}..."

    def test_touch_up_maps_correctly(self):
        mapper = CategoryMapper()
        result = mapper.map_task("Painting - Touch Up (437230) (INT) [LS]")
        assert result.category_display == "TOUCH UP"

    def test_roll_walls_maps_correctly(self):
        mapper = CategoryMapper()
        result = mapper.map_task("Painting - Roll Walls (437240) (INT) [LS]")
        assert result.category_display == "ROLL WALLS FINAL"
