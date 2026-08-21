from pathlib import Path


XML_LAST = Path(__file__).parents[1] / "commands" / "smart_post_dialog" / "xml_last.cps"


def _function(source, name):
    start = source.index(f"function {name}(")
    opening = source.index("{", start)
    depth = 0
    for index in range(opening, len(source)):
        if source[index] == "{":
            depth += 1
        elif source[index] == "}":
            depth -= 1
            if depth == 0:
                return source[start : index + 1]
    raise AssertionError(f"Unterminated function: {name}")


def test_current_xml_post_preserves_real_tool_type():
    source = XML_LAST.read_text(encoding="utf-8")
    on_section = _function(source, "onSection")

    assert "getToolTypeName(tool.type)" in on_section
    assert "type='unspecified'" not in on_section


def test_source_section_type_is_logged_without_inferring_from_strategy():
    source = XML_LAST.read_text(encoding="utf-8")
    on_section = _function(source, "onSection")

    assert "currentSection.getType()" in on_section
    assert "currentSection.type=" in on_section
    assert "operation:isTurningStrategy" not in on_section


def test_cycle_parameters_remain_float_for_xml_importer_contract():
    source = XML_LAST.read_text(encoding="utf-8")
    on_cycle = _function(source, "onCycle")

    assert 'var type = "float";' in on_cycle
    assert "value % 1" not in on_cycle
    assert "<group id='" in on_cycle


def test_cycles_are_grouped_instead_of_expanded():
    source = XML_LAST.read_text(encoding="utf-8")

    assert "expandCyclePoint" not in _function(source, "onCyclePoint")
    assert "<linear to='" in _function(source, "onCyclePoint")
    assert "</group>" in _function(source, "onCycleEnd")


def test_parameter_values_support_strings_numbers_arrays_and_vectors():
    source = XML_LAST.read_text(encoding="utf-8")
    make_value = _function(source, "makeValue")

    assert 'typeof value == "string"' in make_value
    assert 'typeof value == "number"' in make_value
    assert "value instanceof Array" in make_value
    assert "value.x" in make_value
    assert "value.y" in make_value
    assert "value.z" in make_value
