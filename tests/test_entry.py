from pathlib import Path
from types import SimpleNamespace

import pytest

from commands.smart_post_dialog import entry


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("true", True),
        (" YES ", True),
        ("on", True),
        ("false", False),
        ("0", False),
        (False, False),
        (1, True),
    ],
)
def test_as_bool_handles_persisted_values(value, expected):
    assert entry.as_bool(value) is expected


def test_personal_xml_post_defaults_to_current_and_keeps_legacy_option():
    assert Path(entry.get_personal_xml_post(False)).name == "xml_last.cps"
    assert Path(entry.get_personal_xml_post("false")).name == "xml_last.cps"
    assert Path(entry.get_personal_xml_post(True)).name == "xml.cps"


@pytest.mark.parametrize(
    ("unit", "option_name"),
    [(0, "InchesOutput"), (1, "MillimetersOutput")],
)
def test_process_operations_maps_ui_unit_to_fusion_enum(
    monkeypatch, tmp_path, unit, option_name
):
    created = []

    class FakePostProcessInput:
        @staticmethod
        def create(program_name, post_processor, output_folder, output_units):
            value = SimpleNamespace(
                program_name=program_name,
                post_processor=post_processor,
                output_folder=output_folder,
                output_units=output_units,
                isOpenInEditor=True,
                postProperties=None,
            )
            created.append(value)
            return value

    class FakeNamedValues:
        @staticmethod
        def create():
            return SimpleNamespace(add=lambda *_args: None)

    class FakeCam:
        @staticmethod
        def postProcess(_operation, post_input):
            output = Path(post_input.output_folder) / f"{post_input.program_name}.xml"
            output.write_text("<nc></nc>", encoding="utf-8")
            return True

    monkeypatch.setattr(entry.adsk.cam, "PostProcessInput", FakePostProcessInput)
    monkeypatch.setattr(entry.adsk.core, "NamedValues", FakeNamedValues)
    monkeypatch.setattr(entry.futil, "log", lambda *_args, **_kwargs: None)

    result = entry.process_operations(
        FakeCam(),
        [SimpleNamespace(name="Drill")],
        "1001",
        "xml_last.cps",
        str(tmp_path),
        unit,
        {},
    )

    expected = getattr(entry.adsk.cam.PostOutputUnitOptions, option_name)
    assert created[0].output_units == expected
    assert created[0].isOpenInEditor is False
    assert result == [entry.normalize_path(str(tmp_path / "1001_1.xml"))]


def test_process_operations_rejects_unresolved_document_unit(tmp_path):
    with pytest.raises(ValueError, match="Unsupported resolved output unit"):
        entry.process_operations(
            SimpleNamespace(), [], "1001", "xml_last.cps", str(tmp_path), 2, {}
        )


def test_merge_xml_files_keeps_first_header_and_appends_later_sections(
    monkeypatch, tmp_path
):
    first = tmp_path / "first.xml"
    second = tmp_path / "second.xml"
    output = tmp_path / "merged.xml"
    first.write_text(
        "<nc>\n<parameter name='program' value='1001'/>\n"
        "<tool number='1'/><section><rapid to='1 2 3'/></section>\n</nc>",
        encoding="utf-8",
    )
    second.write_text(
        "<nc>\n<parameter name='program' value='ignored'/>\n"
        "<parameter name='areBothSpindlesGrabbed' value='0'/>\n"
        "<tool number='2'/><section><linear to='4 5 6'/></section>\n</nc>",
        encoding="utf-8",
    )
    monkeypatch.setattr(entry.futil, "log", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(
        entry,
        "show_message",
        lambda message, *_args: pytest.fail(f"Unexpected merge error: {message}"),
    )

    assert entry.merge_xml_files([str(first), str(second)], str(output)) is True

    merged = output.read_text(encoding="utf-8")
    assert merged.count("<nc>") == 1
    assert merged.count("</nc>") == 1
    assert "number='1'" in merged
    assert "number='2'" in merged
    assert "areBothSpindlesGrabbed" in merged
    assert "value='ignored'" not in merged
    assert not first.exists()
    assert not second.exists()
