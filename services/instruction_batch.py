from __future__ import annotations

from pathlib import Path
import shutil

from services.instruction_library import (
    find_instruction,
    get_library_dir,
)
from services.instruction_generator import generate_instruction
from services.instruction_builder import build_instruction_document
from services.instruction_naming import canonical_library_filename


LIBRARY_DIR = get_library_dir()

TEMP_OUTPUT_DIR = Path(
    "output/generated"
)


def _clean(value) -> str:
    return " ".join(
        str(value or "").split()
    ).strip()


def _deduplicate(values) -> list[str]:
    result = []
    seen = set()

    for value in values or []:
        value = _clean(value)

        if not value:
            continue

        key = (
            value.lower()
            .replace("ё", "е")
        )

        if key in seen:
            continue

        seen.add(key)
        result.append(value)

    return result


def collect_instruction_targets(
    questionnaire: dict,
) -> list[dict]:
    positions = _deduplicate(
        questionnaire.get(
            "positions",
            [],
        )
    )

    regular_works = _deduplicate(
        questionnaire.get(
            "regular_works",
            [],
        )
    )

    high_risk_works = _deduplicate(
        questionnaire.get(
            "high_risk_works",
            [],
        )
    )

    targets = []

    for name in positions:
        targets.append({
            "name": name,
            "kind": "position",
            "kind_label": "должность",
        })

    for name in regular_works:
        targets.append({
            "name": name,
            "kind": "regular_work",
            "kind_label": "вид работ",
        })

    for name in high_risk_works:
        targets.append({
            "name": name,
            "kind": "high_risk_work",
            "kind_label": (
                "работа повышенной опасности"
            ),
        })

    return targets


def build_generation_context(
    questionnaire: dict,
    target: dict,
) -> dict:
    context = dict(
        questionnaire
    )

    context["source"] = (
        "questionnaire"
    )

    context["target_name"] = (
        target["name"]
    )

    context["target_kind"] = (
        target["kind"]
    )

    if target["kind"] in {
        "regular_work",
        "high_risk_work",
    }:
        context["position_works"] = [
            target["name"]
        ]
    else:
        context["position_works"] = []

    context.setdefault(
        "shift_work",
        None,
    )

    context.setdefault(
        "dangerous_substances",
        None,
    )

    context.setdefault(
        "dangerous_waste",
        None,
    )

    context.setdefault(
        "professional_risks",
        [],
    )

    context.setdefault(
        "ppe",
        [],
    )

    return context


def _temporary_output_path(
    target: dict,
) -> Path:
    TEMP_OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    filename = (
        canonical_library_filename(
            target_name=target["name"],
            target_kind=target["kind"],
        )
    )

    return (
        TEMP_OUTPUT_DIR
        / filename
    )


def generate_missing_instruction(
    target: dict,
    questionnaire: dict,
) -> dict:
    name = _clean(
        target["name"]
    )

    kind = target[
        "kind"
    ]

    existing = find_instruction(
        name,
        target_kind=kind,
    )

    if existing:
        return {
            "name": name,
            "kind": kind,
            "library_path": Path(
                existing
            ),
            "generated": False,
        }

    print()
    print(
        "=" * 60
    )

    print(
        "Генерирую ИОТ:",
        name,
    )

    print(
        "=" * 60
    )

    context = build_generation_context(
        questionnaire,
        target,
    )

    instruction_text = (
        generate_instruction(
            name,
            context=context,
        )
    )

    LIBRARY_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    library_path = (
        LIBRARY_DIR
        / canonical_library_filename(
            target_name=name,
            target_kind=kind,
        )
    )

    temp_path = _temporary_output_path(
        target
    )

    build_instruction_document(
        target_name=name,
        target_kind=kind,
        instruction_text=instruction_text,
        output_path=temp_path,
    )

    if not temp_path.exists():
        raise FileNotFoundError(
            "instruction_builder не создал файл: "
            f"{temp_path}"
        )

    shutil.copy2(
        temp_path,
        library_path,
    )

    temp_path.unlink()

    if not library_path.exists():
        raise RuntimeError(
            "ИОТ скопирована в библиотеку, "
            "но файл не найден по ожидаемому пути: "
            f"{library_path}"
        )

    print(
        "Сохранено в библиотеку:",
        library_path,
    )

    return {
        "name": name,
        "kind": kind,
        "library_path": library_path,
        "generated": True,
    }
