import xml.etree.ElementTree as ET

from logic.sc_xml_parser import (
    MAX_SMARTCAT_INDEX_GAP,
    SMARTCAT_NS,
    _worksheet_to_rows,
)


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
