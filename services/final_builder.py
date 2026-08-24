from __future__ import annotations

from datetime import datetime
from pathlib import Path
import json
import os
import shutil
import tempfile
import zipfile
import re

from lxml import etree

try:
    import win32com.client
except ImportError as exc:
    raise RuntimeError(
        "Не установлен pywin32. Выполни: pip install pywin32"
    ) from exc

from services.instruction_naming import instruction_registry_name


SOURCE_DOCM = Path("templates/ИСХОДНЫЙ_ФАЙЛ.docm")
OUTPUT_DIR = Path("output")

COUNTER_FILE = Path(
    "runtime/counters/document_counter.json"
)

WD_COLLAPSE_END = 0
WD_PAGE_BREAK = 7
WD_DO_NOT_SAVE_CHANGES = 0
WD_FORMAT_XML_DOCUMENT_MACRO_ENABLED = 13
WD_REPLACE_ALL = 2


# ============================================================
# ОБЩИЕ ФУНКЦИИ
# ============================================================

def _clean(value) -> str:
    return " ".join(
        str(value or "")
        .replace("\x07", "")
        .split()
    ).strip()


def _cell_text(cell) -> str:
    return (
        str(cell.Range.Text or "")
        .replace("\r", "")
        .replace("\x07", "")
        .strip()
    )


def _set_cell_text(cell, value) -> None:
    cell.Range.Text = _clean(value)


def _add_years(
    date_value: datetime,
    years: int,
) -> datetime:
    try:
        return date_value.replace(
            year=date_value.year + years
        )
    except ValueError:
        return date_value.replace(
            month=2,
            day=28,
            year=date_value.year + years,
        )


# ============================================================
# НОМЕР B1 / B2 / B3...
# ============================================================

def _next_daily_number() -> str:
    today = datetime.now().strftime(
        "%Y-%m-%d"
    )

    COUNTER_FILE.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    state = {
        "date": today,
        "counter": 0,
    }

    if COUNTER_FILE.exists():
        try:
            loaded = json.loads(
                COUNTER_FILE.read_text(
                    encoding="utf-8"
                )
            )

            if loaded.get("date") == today:
                state = loaded

        except Exception:
            pass

    state["date"] = today
    state["counter"] = int(
        state.get("counter", 0)
    ) + 1

    fd, temp_name = tempfile.mkstemp(
        suffix=".json",
        dir=str(
            COUNTER_FILE.parent
        ),
    )
    os.close(fd)

    temp_path = Path(temp_name)

    try:
        temp_path.write_text(
            json.dumps(
                state,
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )

        os.replace(
            temp_path,
            COUNTER_FILE,
        )

    finally:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except Exception:
                pass

    return f"B{state['counter']}"


# ============================================================
# СЛУЖЕБНАЯ ТАБЛИЦА
# ============================================================

def _find_service_table(document):
    for table in document.Tables:
        if table.Rows.Count < 3:
            continue

        try:
            first_row = " ".join(
                _cell_text(
                    table.Cell(
                        1,
                        col,
                    )
                )
                for col in range(
                    1,
                    min(
                        table.Columns.Count,
                        5,
                    ) + 1,
                )
            ).lower()
        except Exception:
            continue

        if (
            "информация от заказчика"
            in first_row
            and "автозамена"
            in first_row
        ):
            return table

    raise RuntimeError(
        "Не найдена служебная таблица автозамены."
    )


def _build_replacement_values(
    questionnaire: dict,
) -> tuple[dict, str]:
    """
    Здесь остаются только общие теги документа.

    Autozam_instr_XXX больше не нужны для реестра:
    Python заполняет реестр напрямую.
    """
    now = datetime.now()

    batch_number = (
        _next_daily_number()
    )

    values = {
        "name_full": questionnaire.get(
            "company_name",
            "",
        ),
        "inndoc": questionnaire.get(
            "inn",
            "",
        ),
        "uridicheskiy_adress": questionnaire.get(
            "legal_address",
            "",
        ),
        "n_ame_otvetstvenniy_ot_full": questionnaire.get(
            "ot_responsible_name",
            "",
        ),
        "dolznost_first_ot": questionnaire.get(
            "ot_responsible_position",
            "",
        ),
        "podpisant_ot": questionnaire.get(
            "director_name",
            "",
        ),
        "dolzn_shortcompname": questionnaire.get(
            "director_position",
            "",
        ),
        "address_object": questionnaire.get(
            "object_address",
            "",
        ),
        "obj_name": questionnaire.get(
            "object_name",
            "",
        ),
        "date_doc": now.strftime(
            "%d.%m.%Y"
        ),
        "date_final": _add_years(
            now,
            5,
        ).strftime(
            "%d.%m.%Y"
        ),
        "NMBP": batch_number,
    }

    return values, batch_number


def _fill_service_table(
    document,
    questionnaire: dict,
) -> str:
    table = _find_service_table(
        document
    )

    values, batch_number = (
        _build_replacement_values(
            questionnaire
        )
    )

    for row_index in range(
        3,
        table.Rows.Count + 1,
    ):
        if table.Columns.Count < 5:
            break

        tag_raw = _cell_text(
            table.Cell(
                row_index,
                5,
            )
        )

        if not tag_raw:
            continue

        tag = (
            tag_raw
            .replace(" ", "")
            .strip("/")
        )

        if tag not in values:
            continue

        _set_cell_text(
            table.Cell(
                row_index,
                4,
            ),
            values[tag],
        )

    return batch_number


# ============================================================
# ДИНАМИЧЕСКИЙ РЕЕСТР ИОТ
# ============================================================

def _find_registry_table(document):
    """
    Реестр ИОТ:
    2 колонки
    № | Наименование инструкции
    """
    for table in document.Tables:

        if table.Columns.Count != 2:
            continue

        try:
            header = (
                _cell_text(table.Cell(1, 1))
                + " "
                + _cell_text(table.Cell(1, 2))
            ).lower()
        except Exception:
            continue

        if (
            "№" in _cell_text(table.Cell(1, 1))
            and "наименование инструкции" in header
        ):
            return table

    raise RuntimeError("Не найден реестр ИОТ.")


def _get_reserved_rows(table) -> list[dict]:
    """
    Возвращает резервные строки реестра и номера их тегов.
    """
    result = []

    for row_index in range(
        2,
        table.Rows.Count + 1,
    ):
        try:
            tag = _cell_text(
                table.Cell(
                    row_index,
                    2,
                )
            )
        except Exception:
            continue

        if not tag.startswith(
            "Autozam_instr_"
        ):
            continue

        try:
            tag_number = int(
                tag.rsplit(
                    "_",
                    1,
                )[1]
            )
        except Exception:
            continue

        result.append({
            "row_index": row_index,
        })

    return result

def _find_toc_table(document):
    """
    Оглавление:
    3 колонки
    № | Наименование документа | Тип документа
    """
    for table in document.Tables:
        if table.Columns.Count != 3:
            continue
        try:
            header = (
                _cell_text(table.Cell(1, 1))
                + " "
                + _cell_text(table.Cell(1, 2))
                + " "
                + _cell_text(table.Cell(1, 3))
            ).lower()
        except Exception:
            continue
        if (
            "№" in _cell_text(table.Cell(1, 1))
            and "наименование документа" in header
        ):
            return table
    raise RuntimeError("Не найдено оглавление.")
def _fill_toc_registry(
    document,
    instruction_items: list[dict],
) -> int:
    """
    Дописывает в конец «Оглавления» строки для всех инструкций,
    продолжая нумерацию с последней имеющейся строки.
    """
    table = _find_toc_table(document)
    last_row = table.Rows.Count
    try:
        last_number = int(
            _cell_text(table.Cell(last_row, 1))
        )
    except Exception:
        last_number = last_row - 1
    added = 0
    for item in instruction_items:
        last_number += 1
        name = instruction_registry_name(
            target_name=item["name"],
            target_kind=item["kind"],
        )
        new_row = table.Rows.Add()
        _set_cell_text(new_row.Cells(1), str(last_number))
        _set_cell_text(new_row.Cells(2), name)
        _set_cell_text(new_row.Cells(3), "Инструкция")
        added += 1
    return added

def _fill_dynamic_registry(
    document,
    instruction_items: list[dict],
) -> dict:
    """
    Реестр теперь полностью формирует Python.

    Если нужно:
      5 ИОТ   -> будет 5 строк;
      13 ИОТ  -> будет 13 строк;
      55 ИОТ  -> будет 55 строк;
      120 ИОТ -> будет 120 строк.

    Имена инструкций пишутся прямо в реестр.
    Autozam_instr_XXX в реестре после сборки не остаются.
    """
    table = _find_registry_table(
        document
    )

    reserved_rows = _get_reserved_rows(
        table
    )

    if not reserved_rows:
        raise RuntimeError(
            "В реестре не найдены резервные строки Autozam_instr_XXX."
        )

    existing_capacity = len(
        reserved_rows
    )

    first_row_index = reserved_rows[0]["row_index"]

    try:
        number_offset = int(
            _cell_text(
                table.Cell(first_row_index - 1, 1)
            )
        )
    except Exception:
        number_offset = 0

    instruction_count = len(
        instruction_items
    )

    # 1. Заполняем существующие резервные строки.
    reusable_count = min(
        instruction_count,
        existing_capacity,
    )

    for index in range(
        reusable_count
    ):
        row_index = reserved_rows[
            index
        ][
            "row_index"
        ]

        item = instruction_items[
            index
        ]

        name = instruction_registry_name(
            target_name=item["name"],
            target_kind=item["kind"],
        )

        _set_cell_text(
            table.Cell(
                row_index,
                1,
            ),
            str(number_offset + index + 1),
        )

        _set_cell_text(
            table.Cell(
                row_index,
                2,
            ),
            name,
        )

    # 2. Если ИОТ больше, чем резервных строк,
    # добавляем новые строки в конец этой же таблицы.
    added = 0

    if instruction_count > existing_capacity:

        for index in range(
                existing_capacity,
                instruction_count,
        ):
            new_row = table.Rows.Add()

            item = instruction_items[index]

            name = instruction_registry_name(
                target_name=item["name"],
                target_kind=item["kind"],
            )

            _set_cell_text(
                new_row.Cells(1),
                str(number_offset + index + 1)
            )

            _set_cell_text(
                new_row.Cells(2),
                name,
            )

            added += 1

    # 3. Если ИОТ меньше, чем резервных строк,
    # удаляем ненужный хвост снизу вверх.
    deleted = 0

    if instruction_count < existing_capacity:
        rows_to_delete = reserved_rows[
            instruction_count:
        ]

        for item in reversed(
            rows_to_delete
        ):
            table.Rows(
                item[
                    "row_index"
                ]
            ).Delete()

            deleted += 1

    return {
        "instruction_count": instruction_count,
        "template_capacity": existing_capacity,
        "rows_reused": reusable_count,
        "rows_added": added,
        "rows_deleted": deleted,
    }

# ============================================================
# ОТСТУПЫ ПОСЛЕ ЗАГОЛОВКОВ РАЗДЕЛОВ
# ============================================================

_SECTION_HEADING_PATTERN = re.compile(r"^\d+\.\s+\S")

def _fix_instruction_title_font(document) -> int:
    """
    У заголовка "ИНСТРУКЦИЯ" и следующей за ним строки с
    названием инструкции размер шрифта иногда не проставлен
    явно и наследуется от стиля (11pt). Принудительно ставим 10pt.
    """
    fixed = 0
    fixing = False

    for paragraph in document.Paragraphs:
        text = _clean(paragraph.Range.Text)

        if fixing:
            paragraph.Range.Font.Size = 10
            fixed += 1

            if text:
                fixing = False
            continue

        if text == "ИНСТРУКЦИЯ":
            paragraph.Range.Font.Size = 10
            fixed += 1
            fixing = True

    return fixed

_DOCUMENT_NUMBER_PATTERN = re.compile(r"^\d+/B\d+$")

def _replace_generic_placeholders(
    document,
    questionnaire: dict,
    batch_number: str,
) -> dict:
    """
    Отдельно сгенерированные через ИИ ИОТ (см.
    instruction_builder.py) намеренно оставляют общие теги
    "obj_name", "address_object" и "/NMBP" нетронутыми — чтобы
    сохранённый в библиотеку файл оставался пригоден для любого
    другого клиента. Реальные значения подставляются здесь, при
    сборке итогового документа конкретного клиента.
    """
    replacements = {
        "obj_name": _clean(
            questionnaire.get("object_name", "")
        ),
        "address_object": _clean(
            questionnaire.get("object_address", "")
        ),
        "/NMBP": "/" + batch_number,
    }

    for tag, value in replacements.items():
        find = document.Content.Find
        find.ClearFormatting()
        find.Text = tag
        find.Replacement.ClearFormatting()
        find.Replacement.Text = value
        find.Execute(Replace=WD_REPLACE_ALL)

    return replacements

def _fix_document_number_font(document) -> int:
    """
    В таблице "Номер документа* / Дата составления" размер
    шрифта у значений иногда наследуется от стиля (11pt).
    Ставим 10pt прямо на ячейки таблицы, а не по тексту абзаца.
    """
    fixed = 0

    for table in document.Tables:
        if table.Columns.Count != 2 or table.Rows.Count != 2:
            continue

        try:
            header = (
                _cell_text(table.Cell(1, 1))
                + " "
                + _cell_text(table.Cell(1, 2))
            ).lower()
        except Exception:
            continue

        if "номер документа" not in header:
            continue

        table.Cell(2, 1).Range.Font.Size = 10
        table.Cell(2, 2).Range.Font.Size = 10
        fixed += 1

    return fixed

def _apply_heading_spacing(document) -> int:
    updated = 0

    for paragraph in document.Paragraphs:
        text = _clean(paragraph.Range.Text)

        if _SECTION_HEADING_PATTERN.match(text):
            paragraph.SpaceAfter = 12
            updated += 1

    return updated

# ============================================================
# ВСТАВКА ИОТ
# ============================================================

def _append_instruction(
    document,
    instruction_path: Path,
    add_page_break: bool = True,
) -> None:
    rng = document.Content

    rng.Collapse(
        WD_COLLAPSE_END
    )

    WD_CHARACTER = 1

    rng.MoveEnd(
        Unit=WD_CHARACTER,
        Count=-1,
    )

    rng.Collapse(
        WD_COLLAPSE_END
    )

    if add_page_break:
        rng.InsertBreak(
            WD_PAGE_BREAK
        )

        rng.Collapse(
            WD_COLLAPSE_END
        )

    rng.InsertFile(
        str(
            instruction_path.resolve()
        )
    )


# ============================================================
# МИНИМАЛЬНАЯ OOXML-ЧИСТКА
# ============================================================

W_NS = (
    "http://schemas.openxmlformats.org/"
    "wordprocessingml/2006/main"
)

NS = {
    "w": W_NS
}


def _paragraph_visible_text(element) -> str:
    return "".join(
        node.text or ""
        for node in element.xpath(
            ".//w:t",
            namespaces=NS,
        )
    ).strip()


def _paragraph_has_visual_content(element) -> bool:
    return bool(
        element.xpath(
            ".//w:drawing | .//w:pict | .//w:object",
            namespaces=NS,
        )
    )

def _paragraph_has_break(element) -> bool:
    if element.xpath(
        "./w:pPr/w:sectPr",
        namespaces=NS,
    ):
        return True

    if element.xpath(
        ".//w:br[@w:type='page']",
        namespaces=NS,
    ):
        return True

    return False

def _neutralize_redundant_section_breaks(body) -> int:
    """
    Разрыв раздела, который вставленный файл инструкции принёс
    сам, больше не нужен: мы и так добавляем свой разрыв страницы
    перед каждой инструкцией. Иногда такой разрыв тянет за собой
    свою ссылку на колонтитул, а смена колонтитула сама по себе
    заставляет Word начать новую страницу — поэтому просто менять
    тип раздела недостаточно, убираем разрыв раздела целиком.
    """
    children = list(body)
    fixed = 0

    for i, child in enumerate(children):
        if etree.QName(child).localname != "p":
            continue

        p_pr = child.find("w:pPr", namespaces=NS)

        if p_pr is None:
            continue

        sect_pr = p_pr.find("w:sectPr", namespaces=NS)

        if sect_pr is None:
            continue

        if sect_pr.find("w:type", namespaces=NS) is not None:
            continue

        j = i + 1
        redundant = False

        while j < len(children):
            nxt = children[j]

            if etree.QName(nxt).localname != "p":
                break

            if (
                _paragraph_visible_text(nxt)
                or _paragraph_has_visual_content(nxt)
            ):
                break

            if nxt.findall(
                './/w:br[@w:type="page"]',
                namespaces=NS,
            ):
                redundant = True
                break

            j += 1

        if redundant:
            p_pr.remove(sect_pr)
            fixed += 1

            for k in range(i + 1, j):
                body.remove(children[k])

    return fixed

def _is_empty_break_paragraph(element) -> bool:
    paragraph_tag = (
        "{"
        + W_NS
        + "}p"
    )

    if element.tag != paragraph_tag:
        return False

    if _paragraph_visible_text(
        element
    ):
        return False

    if _paragraph_has_visual_content(
        element
    ):
        return False

    return _paragraph_has_break(
        element
    )


def _cleanup_document_xml(
    file_path: Path,
) -> dict:
    archive_items = []

    with zipfile.ZipFile(
        file_path,
        "r",
    ) as archive:
        document_xml = archive.read(
            "word/document.xml"
        )

        for item in archive.infolist():
            archive_items.append(
                (
                    item,
                    archive.read(
                        item.filename
                    ),
                )
            )

    root = etree.fromstring(
        document_xml
    )

    body = root.find(
        "w:body",
        namespaces=NS,
    )

    if body is None:
        return {
            "leading_removed": 0,
            "duplicate_breaks_removed": 0,
            "trailing_removed": 0,
            "redundant_breaks_fixed": 0,
        }

    redundant_breaks_fixed = _neutralize_redundant_section_breaks(body)


    leading_removed = 0
    duplicate_breaks_removed = 0
    trailing_removed = 0

    while (
        len(body) > 0
        and _is_empty_break_paragraph(
            body[0]
        )
    ):
        body.remove(
            body[0]
        )
        leading_removed += 1

    index = 1

    while index < len(body):
        previous = body[
            index - 1
        ]

        current = body[
            index
        ]

        if (
            _is_empty_break_paragraph(
                previous
            )
            and _is_empty_break_paragraph(
                current
            )
        ):
            body.remove(
                current
            )
            duplicate_breaks_removed += 1
            continue

        index += 1

    while len(body) > 0:
        last = body[
            len(body) - 1
        ]

        if _is_empty_break_paragraph(
            last
        ):
            body.remove(
                last
            )
            trailing_removed += 1
            continue

        sect_tag = (
            "{"
            + W_NS
            + "}sectPr"
        )

        if (
            last.tag == sect_tag
            and len(body) >= 2
            and _is_empty_break_paragraph(
                body[
                    len(body) - 2
                ]
            )
        ):
            body.remove(
                body[
                    len(body) - 2
                ]
            )
            trailing_removed += 1
            continue

        break

    changed = (
        leading_removed
        + duplicate_breaks_removed
        + trailing_removed
        + redundant_breaks_fixed
    )

    if changed == 0:
        return {
            "leading_removed": 0,
            "duplicate_breaks_removed": 0,
            "trailing_removed": 0,
            "redundant_breaks_fixed": redundant_breaks_fixed
        }



    new_xml = etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )

    fd, temp_name = tempfile.mkstemp(
        suffix=file_path.suffix,
        dir=str(
            file_path.parent
        ),
    )
    os.close(fd)

    temp_path = Path(
        temp_name
    )

    try:
        with zipfile.ZipFile(
            temp_path,
            "w",
        ) as output_archive:
            for item, data in archive_items:
                if (
                    item.filename
                    == "word/document.xml"
                ):
                    data = new_xml

                output_archive.writestr(
                    item,
                    data,
                )

        os.replace(
            temp_path,
            file_path,
        )

    finally:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except Exception:
                pass

    return {
        "leading_removed": leading_removed,
        "duplicate_breaks_removed": (
            duplicate_breaks_removed
        ),
        "trailing_removed": trailing_removed,
        "redundant_breaks_fixed": redundant_breaks_fixed
    }


# ============================================================
# ОСНОВНАЯ СБОРКА
# ============================================================

def build_final_document(
    questionnaire: dict,
    instruction_items: list[dict],
    source_docm=SOURCE_DOCM,
    output_dir=OUTPUT_DIR,
    visible_word=False,
) -> dict:
    source_docm = Path(
        source_docm
    )

    output_dir = Path(
        output_dir
    )

    if not source_docm.exists():
        raise FileNotFoundError(
            f"Не найден исходный DOCM: {source_docm}"
        )

    if not instruction_items:
        raise RuntimeError(
            "Список ИОТ для сборки пуст."
        )

    output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    work_path = (
        output_dir
        / f"_work_{timestamp}.docm"
    )

    output_path = (
        output_dir
        / f"ИТОГОВЫЙ_КОМПЛЕКТ_{timestamp}.docm"
    )

    shutil.copy2(
        source_docm,
        work_path,
    )

    word = (
        win32com.client.DispatchEx(
            "Word.Application"
        )
    )

    word.Visible = visible_word
    word.DisplayAlerts = 0

    document = None

    try:
        document = (
            word.Documents.Open(
                str(
                    work_path.resolve()
                )
            )
        )

        print()
        print(
            "=== FINAL BUILDER ==="
        )

        batch_number = (
            _fill_service_table(
                document,
                questionnaire,
            )
        )

        print(
            "Номер комплекта:",
            batch_number,
        )

        registry_result = (
            _fill_dynamic_registry(
                document,
                instruction_items,
            )
        )

        print(
            "Динамический реестр:",
            registry_result,
        )

        toc_added = _fill_toc_registry(
            document,
            instruction_items,
        )

        print(
            "Оглавление дополнено строк:",
            toc_added,
        )

        print(
            "ИОТ для вставки:",
            len(instruction_items),
        )

        for index, item in enumerate(
            instruction_items,
            start=1,
        ):
            path = Path(
                item["path"]
            )

            if not path.exists():
                raise FileNotFoundError(
                    f"Не найдена ИОТ: {path}"
                )

            print(
                f"Вставляю {index}/"
                f"{len(instruction_items)}: "
                f"{path.name}"
            )

            _append_instruction(
                document,
                path,
            )

        _apply_heading_spacing(document)

        title_font_fixed = _fix_instruction_title_font(document)

        print(
            "Заголовков ИНСТРУКЦИЯ поправлен шрифт:",
            title_font_fixed,
        )

        doc_number_font_fixed = _fix_document_number_font(document)

        print(
            "Номеров документа поправлен шрифт:",
            doc_number_font_fixed,
        )

        placeholders_replaced = _replace_generic_placeholders(
            document,
            questionnaire,
            batch_number,
        )

        print(
            "Общие плейсхолдеры заменены:",
            placeholders_replaced,
        )

        document.SaveAs2(

            str(
                output_path.resolve()
            ),
            FileFormat=(
                WD_FORMAT_XML_DOCUMENT_MACRO_ENABLED
            ),
        )

        document.Close(
            SaveChanges=(
                WD_DO_NOT_SAVE_CHANGES
            )
        )

        document = None

        try:
            word.Quit()
        except Exception:
            pass

        word = None

        cleanup = (
            _cleanup_document_xml(
                output_path
            )
        )

        try:
            work_path.unlink(
                missing_ok=True
            )
        except Exception:
            pass

        print(
            "OOXML cleanup:",
            cleanup,
        )

        print(
            "Итоговый DOCM:",
            output_path,
        )

        return {
            "output_path": output_path,
            "instruction_count": len(
                instruction_items
            ),
            "batch_number": batch_number,
            "registry": registry_result,
            "cleanup": cleanup,
        }

    finally:
        if document is not None:
            try:
                document.Close(
                    SaveChanges=(
                        WD_DO_NOT_SAVE_CHANGES
                    )
                )
            except Exception:
                pass

        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
