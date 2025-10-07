import textwrap
import xml.etree.ElementTree as ET

from logic.sc_xml_parser import (
    MAX_SMARTCAT_INDEX_GAP,
    SMARTCAT_NS,
    _worksheet_to_rows,
    ROW_NAMES,
    parse_smartcat_report,
)
from logic.trados_xml_parser import parse_reports


def _build_worksheet(index: int) -> ET.Element:
    xml = f"""
    <Worksheet xmlns='{SMARTCAT_NS}' xmlns:ss='{SMARTCAT_NS}'>
        <Table>
            <Row>
                <Cell ss:Index='{index}'>
                    <Data>Value</Data>
                </Cell>
            </Row>
        </Table>
    </Worksheet>
    """
    return ET.fromstring(xml)


def test_worksheet_to_rows_caps_large_index_gap():
    worksheet = _build_worksheet(1048576)

    rows = _worksheet_to_rows(worksheet)

    assert len(rows) == 1
    row = rows[0]

    assert len(row) <= MAX_SMARTCAT_INDEX_GAP + 1
    assert row[-1] == "Value"


def _write_smartcat_report(tmp_path) -> str:
    content = textwrap.dedent(
        f"""
        <?xml version="1.0" encoding="utf-8"?>
        <Workbook xmlns="{SMARTCAT_NS}" xmlns:ss="{SMARTCAT_NS}">
            <Worksheet ss:Name="Statistics [DE]">
                <Table>
                    <Row>
                        <Cell><Data>Match type</Data></Cell>
                        <Cell ss:Index="1048576"><Data>Words</Data></Cell>
                    </Row>
                    <Row>
                        <Cell><Data>New</Data></Cell>
                        <Cell ss:Index="1048576"><Data>12</Data></Cell>
                    </Row>
                </Table>
            </Worksheet>
        </Workbook>
        """
    ).strip()

    path = tmp_path / "smartcat_large_index.xml"
    path.write_text(content, encoding="utf-8")
    return str(path)


def _write_trados_report(tmp_path) -> str:
    content = textwrap.dedent(
        """
        <?xml version="1.0" encoding="utf-8"?>
        <task>
            <taskInfo>
                <language name="Russian" lcid="ru-RU" />
            </taskInfo>
            <file name="example.docx">
                <analyse>
                    <new words="10" />
                    <fuzzy min="75" max="84" words="5" />
                    <fuzzy min="95" max="99" words="3" />
                    <exact words="7" />
                </analyse>
            </file>
        </task>
        """
    ).strip()

    path = tmp_path / "Analyze Files en-US_ru-RU.xml"
    path.write_text(content, encoding="utf-8")
    return str(path)


def test_parse_smartcat_report_handles_large_index(tmp_path):
    report_path = _write_smartcat_report(tmp_path)

    results, warnings, processed, pair_key = parse_smartcat_report(report_path, "words")

    assert warnings == []
    assert processed is True
    assert pair_key in results
    values = results[pair_key]
    assert values[ROW_NAMES[0]] == 12.0
    assert sum(values[name] for name in ROW_NAMES[1:]) == 0.0


def test_parse_reports_handles_trados_and_smartcat(tmp_path):
    smartcat_path = _write_smartcat_report(tmp_path)
    trados_path = _write_trados_report(tmp_path)

    results, warnings, sources_map = parse_reports(
        [smartcat_path, trados_path], unit="Words"
    )

    assert warnings == []
    assert len(results) == 2

    smartcat_entry = None
    trados_entry = None

    for pair_key, values in results.items():
        sources = sources_map[pair_key]
        if "smartcat_large_index.xml" in sources:
            smartcat_entry = (pair_key, values)
        if "Analyze Files en-US_ru-RU.xml" in sources:
            trados_entry = (pair_key, values)

    assert smartcat_entry is not None
    assert trados_entry is not None
    assert smartcat_entry[0] != trados_entry[0]

    smartcat_values = smartcat_entry[1]
    assert smartcat_values[ROW_NAMES[0]] == 12.0
    assert sum(smartcat_values[name] for name in ROW_NAMES[1:]) == 0.0

    trados_values = trados_entry[1]
    assert trados_values[ROW_NAMES[0]] == 10.0
    assert trados_values[ROW_NAMES[1]] == 5.0
    assert trados_values[ROW_NAMES[2]] == 3.0
    assert trados_values[ROW_NAMES[3]] == 7.0
