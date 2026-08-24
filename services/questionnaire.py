from zipfile import ZipFile
from lxml import etree


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
}


def read_docx_xml(file_path):
    with ZipFile(file_path) as docx:
        xml_data = docx.read("word/document.xml")

    return etree.fromstring(xml_data)


def get_element_text(element):
    paragraphs = element.xpath(".//w:p", namespaces=NS)

    result = []

    for paragraph in paragraphs:
        texts = paragraph.xpath(".//w:t/text()", namespaces=NS)

        text = "".join(texts).strip()

        if text:
            result.append(text)

    return "\n".join(result).strip()

def _find_questionnaire_table(tables):
    anchor = "Полное наименование организации"

    for table in tables:
        if anchor in get_element_text(table):
            return table

    return None


def get_questionnaire_rows(root):
    tables = root.xpath("//w:tbl", namespaces=NS)

    table = _find_questionnaire_table(tables)

    if table is None:
        raise ValueError("Таблица опросного листа не найдена")

    rows = []

    for row in table.xpath("./w:tr", namespaces=NS):
        cells = row.xpath("./w:tc | .//w:sdtContent/w:tc", namespaces=NS)

        values = []

        for cell in cells:
            values.append(get_element_text(cell))

        rows.append(values)

    return rows


def normalize_list(value):
    if not value:
        return []

    value = value.strip()

    if value.lower() == "нет":
        return []

    result = []

    for line in value.splitlines():

        line = line.strip()

        if not line:
            continue

        if line.startswith("-"):
            line = line[1:].strip()

        if line:
            result.append(line)

    return result


def normalize_field_name(value):
    return " ".join(value.split()).strip()


def parse_questionnaire(file_path):
    root = read_docx_xml(file_path)

    rows = get_questionnaire_rows(root)

    fields = {}

    for row in rows:

        if len(row) < 2:
            continue

        field_name = normalize_field_name(row[0])

        if len(row) >= 3:
            customer_value = row[2].strip()
        else:
            customer_value = row[1].strip()

        fields[field_name] = customer_value

    data = {
        "company_name": fields.get(
            "Полное наименование организации",
            ""
        ),

        "inn": fields.get(
            "ИНН организации",
            ""
        ),

        "legal_address": fields.get(
            "Юридический адрес организации",
            ""
        ),

        "ot_responsible_name": fields.get(
            "ФИО ответственного по охране труда",
            ""
        ),

        "ot_responsible_position": fields.get(
            "Должность ответственного по охране труда",
            ""
        ),

        "director_name": fields.get(
            "ФИО руководителя организации (или сотрудника кто будет утверждать документы)",
            ""
        ),

        "director_position": fields.get(
            "Должность руководителя организации (или сотрудника кто будет утверждать документы)",
            ""
        ),

        "object_address": fields.get(
            "Адрес объекта (в одном опросном листе можно указать только один объект)",
            ""
        ),

        "object_name": fields.get(
            "Наименование объекта (в одном опросном листе можно указать только один объект)",
            ""
        ),

        "email": fields.get(
            "Адрес электронной почты для отправки документов в электронной форме Заказчику",
            ""
        ),

        "employees_count": fields.get(
            "Сколько человек работает в организации?",
            ""
        ),

        "positions": normalize_list(
            fields.get(
                "Список должностей (обязательно приложить штатное расписание в электронном виде)",
                ""
            )
        ),

        "regular_works": normalize_list(
            fields.get(
                "Работы без повышенной опасности, выполняемые в организации",
                ""
            )
        ),

        "high_risk_works": normalize_list(
            fields.get(
                "Работы повышенной опасности, выполняемые в организации (выбрать из списка, если такие имеются):",
                ""
            )
        )
    }

    return data