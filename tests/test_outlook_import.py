import datetime as dt

from logic.outlook_import import OutlookMessage, map_message_to_project_info


HTML_TABLE = """
<table>
<tr><td>Название клиента</td><td>ООО Клиент</td></tr>
<tr><td>Контрагент Logrus IT (с НДС или нет, если с НДС, укажите размер НДС)</td><td>ООО "Логрус ИТ"</td></tr>
<tr><td>Валюта расчетов</td><td>доллары США</td></tr>
<tr><td>Контактное лицо со стороны клиента</td><td>Jane Smith, jane.smith@example.com</td></tr>
<tr><td>Email</td><td>Jane Smith &lt;jane.smith@example.com&gt;</td></tr>
</table>
"""


def make_message(html=HTML_TABLE):
    return OutlookMessage(
        subject="Коммерческое предложение [Project Phoenix]",
        sender_name="Ivan Manager",
        sender_email="manager@example.com",
        sent_at=dt.datetime(2023, 11, 5, 10, 15),
        body="",
        html_body=html,
    )


def test_map_message_extracts_project_fields():
    result = map_message_to_project_info(make_message())

    data = result.data
    assert data.project_name == "Коммерческое предложение"
    assert data.client_name == "ООО Клиент"
    assert data.legal_entity == 'ООО "Логрус ИТ"'
    assert data.currency_code == "USD"
    assert data.contact_name == "Jane Smith"
    assert data.contact_email == "jane.smith@example.com"
    assert data.email == "jane.smith@example.com"
    assert result.missing_fields == []
    assert result.warnings == []
    assert result.sender_name == "Ivan Manager"
    assert result.sender_email == "manager@example.com"
    assert result.sent_at == "2023-11-05 10:15"


def test_map_message_handles_missing_table():
    message = make_message(html="<div>no table here</div>")
    result = map_message_to_project_info(message)

    assert "Таблица в письме не найдена" in result.warnings
    assert "Название проекта" not in result.missing_fields  # project from subject
    assert "Название клиента" in result.missing_fields
    assert result.data.currency_code is None


def test_map_message_handles_byte_html_body():
    html_bytes = HTML_TABLE.encode("utf-8")
    message = make_message(html=html_bytes)

    result = map_message_to_project_info(message)

    assert result.data.client_name == "ООО Клиент"
    assert result.warnings == []


def test_map_message_normalizes_logrus_usa_entity():
    html = HTML_TABLE.replace('ООО "Логрус ИТ"', 'Logrus IT USA')
    result = map_message_to_project_info(make_message(html=html))

    assert result.data.legal_entity == "Logrus IT"
    assert "Юрлицо" not in result.missing_fields


def test_map_message_strips_bracket_tags_from_subject():
    message = make_message()
    message.subject = "[Tag] Новая тема [Internal]"

    result = map_message_to_project_info(message)

    assert result.data.project_name == "Новая тема"


def test_map_message_falls_back_to_plain_text_rows():
    plain_text = """
Название клиента
ООО Клиент

Контрагент Logrus IT (с НДС или нет, если с НДС, укажите размер НДС)
ООО "Логрус ИТ"

Валюта расчетов:
доллары США

Контактное лицо со стороны клиента:
Jane Smith
jane.smith@example.com

Email
Jane Smith <jane.smith@example.com>
"""

    message = OutlookMessage(
        subject="Коммерческое предложение",
        sender_name=None,
        sender_email=None,
        sent_at=None,
        body=plain_text,
        html_body="",
    )

    result = map_message_to_project_info(message)

    data = result.data
    assert data.client_name == "ООО Клиент"
    assert data.legal_entity == 'ООО "Логрус ИТ"'
    assert data.currency_code == "USD"
    assert data.contact_email == "jane.smith@example.com"
    assert data.email == "jane.smith@example.com"
