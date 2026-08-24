from __future__ import annotations

from pathlib import Path

from services.questionnaire import parse_questionnaire
from services.instruction_batch import (
    collect_instruction_targets,
    generate_missing_instruction,
)
from services.instruction_library import find_instruction
from services.final_builder import build_final_document


DEFAULT_QUESTIONNAIRE = Path(
    "input/ОПРОСНЫЙ_ЛИСТ.docx"
)

DEFAULT_SOURCE_DOCM = Path(
    "templates/ИСХОДНЫЙ_ФАЙЛ.docm"
)

DEFAULT_OUTPUT_DIR = Path(
    "output"
)


def _clean(value: str) -> str:
    return " ".join(
        str(value or "").split()
    ).strip()


def _lower_first(value: str) -> str:
    value = _clean(value)

    if not value:
        return value

    return value[:1].lower() + value[1:]


def _normalize_list(values) -> list[str]:
    result = []
    seen = set()

    for value in values or []:
        value = _lower_first(
            value
        )

        if not value:
            continue

        key = (
            value.lower()
            .replace("ё", "е")
        )

        if key in seen:
            continue

        seen.add(
            key
        )

        result.append(
            value
        )

    return result


def _normalize_questionnaire(
    questionnaire: dict,
) -> dict:
    """
    Нормализуются только названия целей из опросника.
    Реквизиты организации не меняются.
    """
    data = dict(
        questionnaire
    )

    data["positions"] = _normalize_list(
        questionnaire.get(
            "positions",
            [],
        )
    )

    data["regular_works"] = _normalize_list(
        questionnaire.get(
            "regular_works",
            [],
        )
    )

    data["high_risk_works"] = _normalize_list(
        questionnaire.get(
            "high_risk_works",
            [],
        )
    )

    return data


def _scan_library(
    targets: list[dict],
) -> tuple[list[dict], list[dict]]:
    found = []
    missing = []

    for target in targets:
        file_path = find_instruction(
            target["name"]
        )

        item = {
            "name": target["name"],
            "kind": target["kind"],
            "kind_label": target.get(
                "kind_label",
                "",
            ),
            "path": (
                Path(file_path)
                if file_path
                else None
            ),
        }

        if file_path:
            found.append(
                item
            )
        else:
            missing.append(
                item
            )

    return found, missing


def _generate_missing(
    missing: list[dict],
    questionnaire: dict,
) -> list[dict]:
    generated = []

    if not missing:
        print(
            "Новых ИОТ создавать не требуется."
        )
        return generated

    print()
    print(
        "=== ГЕНЕРАЦИЯ ОТСУТСТВУЮЩИХ ИОТ ==="
    )

    for index, target in enumerate(
        missing,
        start=1,
    ):
        print()
        print(
            f">>> {index}/{len(missing)}: "
            f"{target['name']}"
        )

        result = (
            generate_missing_instruction(
                target,
                questionnaire,
            )
        )

        if result.get(
            "generated"
        ):
            generated.append({
                "name": target["name"],
                "kind": target["kind"],
                "kind_label": target.get(
                    "kind_label",
                    "",
                ),
                "path": Path(
                    result[
                        "library_path"
                    ]
                ),
            })

    return generated


def _build_instruction_items(
    targets: list[dict],
    resolved: dict,
) -> list[dict]:
    """
    Собирает пути ИОТ по уже известным результатам поиска и
    генерации (found + generated), не сканируя библиотеку заново —
    свежесохранённый по сети файл может быть ещё не виден при
    повторном скане каталога.
    """
    result = []

    for target in targets:
        key = (
            target["name"],
            target["kind"],
        )

        path = resolved.get(key)

        if path is None:
            raise RuntimeError(
                "Не найдена ИОТ перед сборкой: "
                + target["name"]
            )

        result.append({
            "name": target["name"],
            "kind": target["kind"],
            "path": Path(path),
        })

    return result


def run_full_pipeline(
    questionnaire_path=DEFAULT_QUESTIONNAIRE,
    source_docm=DEFAULT_SOURCE_DOCM,
    output_dir=DEFAULT_OUTPUT_DIR,
    visible_word=False,
) -> dict:
    """
    Полный рабочий сценарий:

    1. Читает опросный лист.
    2. Формирует независимые цели:
       должности и виды работ.
    3. Ищет каждую ИОТ в библиотеке.
    4. Генерирует только отсутствующие.
    5. Найденные ИОТ НЕ изменяет.
    6. Собирает итоговый DOCM.
    7. ReplaceTags автоматически НЕ запускает.
    """
    questionnaire_path = Path(
        questionnaire_path
    )

    source_docm = Path(
        source_docm
    )

    output_dir = Path(
        output_dir
    )

    if not questionnaire_path.exists():
        raise FileNotFoundError(
            "Не найден опросный лист: "
            f"{questionnaire_path}"
        )

    if not source_docm.exists():
        raise FileNotFoundError(
            "Не найден исходный DOCM: "
            f"{source_docm}"
        )

    print()
    print(
        "========================================"
    )
    print(
        "      PIPELINE ИОТ"
    )
    print(
        "========================================"
    )

    questionnaire = (
        parse_questionnaire(
            questionnaire_path
        )
    )

    questionnaire = (
        _normalize_questionnaire(
            questionnaire
        )
    )

    targets = (
        collect_instruction_targets(
            questionnaire
        )
    )

    print(
        "Организация:",
        questionnaire.get(
            "company_name",
            "",
        )
    )

    print(
        "Должностей:",
        len(
            questionnaire.get(
                "positions",
                [],
            )
        )
    )

    print(
        "Обычных видов работ:",
        len(
            questionnaire.get(
                "regular_works",
                [],
            )
        )
    )

    print(
        "Работ повышенной опасности:",
        len(
            questionnaire.get(
                "high_risk_works",
                [],
            )
        )
    )

    print(
        "Всего ИОТ:",
        len(targets)
    )

    found, missing = (
        _scan_library(
            targets
        )
    )

    print(
        "Найдено в библиотеке:",
        len(found)
    )

    print(
        "Отсутствует:",
        len(missing)
    )

    generated = (
        _generate_missing(
            missing,
            questionnaire,
        )
    )

    resolved = {
        (item["name"], item["kind"]): item["path"]
        for item in found + generated
    }

    instruction_items = (
        _build_instruction_items(
            targets,
            resolved,
        )
    )

    print()
    print(
        "=== СБОРКА DOCM ==="
    )

    final_result = (
        build_final_document(
            questionnaire=questionnaire,
            instruction_items=instruction_items,
            source_docm=source_docm,
            output_dir=output_dir,
            visible_word=visible_word,
        )
    )

    print()
    print(
        "========================================"
    )
    print(
        "         PIPELINE ЗАВЕРШЁН"
    )
    print(
        "========================================"
    )

    print(
        "ИОТ в комплекте:",
        len(instruction_items)
    )

    print(
        "Новых ИОТ:",
        len(generated)
    )

    print(
        "Итоговый файл:",
        final_result[
            "output_path"
        ]
    )

    return {
        "questionnaire": questionnaire,
        "targets": targets,
        "found_before_generation": found,
        "generated": generated,
        "instruction_items": instruction_items,
        "final": final_result,
        "output_path": final_result[
            "output_path"
        ],
    }
