from __future__ import annotations

import os
import re
import shutil
from pathlib import Path

import pymorphy3
from dotenv import load_dotenv

from services.instruction_naming import instruction_registry_name


load_dotenv()

LIBRARY_DIR = Path(
    os.getenv(
        "IOT_LIBRARY_PATH",
        "library",
    )
)

morph = pymorphy3.MorphAnalyzer()

PREFIXES = (
    "иот для ",
    "иот при ",
    "иот по ",
    "иот ",
)


def get_library_dir() -> Path:
    """
    Возвращает папку библиотеки ИОТ.
    Путь берётся из IOT_LIBRARY_PATH в .env.
    """
    return LIBRARY_DIR


def ensure_library_dir() -> Path:
    """
    Создаёт библиотеку, если её ещё нет.
    """
    library_dir = get_library_dir()

    library_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    return library_dir


def _clean_text(
    value: str,
) -> str:
    value = str(
        value or ""
    ).strip()

    value = value.replace(
        "ё",
        "е",
    )

    value = re.sub(
        r"\s+",
        " ",
        value,
    )

    return value


def _remove_file_extension(
    value: str,
) -> str:
    return re.sub(
        r"\.(docx|docm)$",
        "",
        value,
        flags=re.IGNORECASE,
    )


def _remove_iot_prefix(
    value: str,
) -> str:
    normalized = value.strip()
    lowered = normalized.lower()

    for prefix in PREFIXES:
        if lowered.startswith(
            prefix
        ):
            return normalized[
                len(prefix):
            ].strip()

    return normalized


def _normalize_token(
    token: str,
) -> str:
    token = token.strip(
        " \t\r\n.,;:()[]{}«»\"'"
    )

    if not token:
        return ""

    parsed = morph.parse(
        token
    )

    if not parsed:
        return token.lower()

    return (
        parsed[0]
        .normal_form
        .lower()
    )


def normalize_instruction_name(
    value: str,
) -> str:
    """
    Нормализация для сравнения по смысловому имени.

    Например:
      электрик
      электрика
      ИОТ для электрика.docx

    -> сопоставимое значение.
    """
    value = _clean_text(
        value
    )

    value = _remove_file_extension(
        value
    )

    value = _remove_iot_prefix(
        value
    )

    value = value.lower()

    words = re.findall(
        r"[а-яa-z0-9-]+",
        value,
        flags=re.IGNORECASE,
    )

    normalized_words = []

    for word in words:
        normalized = _normalize_token(
            word
        )

        if normalized:
            normalized_words.append(
                normalized
            )

    return " ".join(
        normalized_words
    )


def _detect_file_kind(
    file_path: Path,
) -> str | None:
    """
    Определяет тип ИОТ по имени файла.

    ИОТ для ... -> position
    ИОТ при ... / ИОТ по ... -> work

    Для старых файлов без префикса тип определить нельзя.
    """
    name = _remove_file_extension(
        file_path.name
    ).strip().lower()

    if name.startswith(
        "иот для "
    ):
        return "position"

    if (
        name.startswith("иот при ")
        or name.startswith("иот по ")
    ):
        return "work"

    return None


def _kind_group(
    target_kind: str | None,
) -> str | None:
    if target_kind == "position":
        return "position"

    if target_kind in {
        "regular_work",
        "high_risk_work",
    }:
        return "work"

    return None


def list_instruction_files() -> list[Path]:
    """
    Возвращает DOCX/DOCM из библиотеки.
    """
    library_dir = ensure_library_dir()

    files = []

    for extension in (
        "*.docx",
        "*.docm",
    ):
        files.extend(
            library_dir.glob(
                extension
            )
        )

    return sorted(
        files,
        key=lambda path: (
            path.name.lower()
        ),
    )


def find_instruction(
    instruction_name: str,
    target_kind: str | None = None,
) -> Path | None:
    """
    Ищет ИОТ в библиотеке.

    Если передан target_kind:
      position
      regular_work
      high_risk_work

    поиск дополнительно учитывает тип ИОТ,
    чтобы должность и вид работ с одинаковым названием
    не схлопывались.
    """
    target_normalized = (
        normalize_instruction_name(
            instruction_name
        )
    )

    if not target_normalized:
        return None

    requested_group = _kind_group(
        target_kind
    )

    exact_candidates = []
    legacy_candidates = []

    for file_path in list_instruction_files():
        file_normalized = (
            normalize_instruction_name(
                file_path.name
            )
        )

        if (
            file_normalized
            != target_normalized
        ):
            continue

        if requested_group is None:
            exact_candidates.append(
                file_path
            )
            continue

        file_group = _detect_file_kind(
            file_path
        )

        if file_group == requested_group:
            exact_candidates.append(
                file_path
            )
        elif file_group is None:
            legacy_candidates.append(
                file_path
            )

    if exact_candidates:
        return exact_candidates[0]

    # Совместимость со старой библиотекой:
    # если есть только старый файл без понятного префикса,
    # используем его.
    if legacy_candidates:
        return legacy_candidates[0]

    return None


def instruction_exists(
    instruction_name: str,
    target_kind: str | None = None,
) -> bool:
    return (
        find_instruction(
            instruction_name,
            target_kind=target_kind,
        )
        is not None
    )


def save_instruction_to_library(
    source_path: str | Path,
    target_name: str | None = None,
    target_kind: str | None = None,
    destination_filename: str | None = None,
) -> Path:
    """
    Копирует готовую ИОТ в библиотеку.

    Предпочтительный вариант:
        target_name + target_kind

    Тогда имя файла строится через единый naming-модуль.
    """
    source_path = Path(
        source_path
    )

    if not source_path.exists():
        raise FileNotFoundError(
            "Не найден исходный файл ИОТ: "
            f"{source_path}"
        )

    library_dir = ensure_library_dir()

    if destination_filename:
        destination_path = (
            library_dir
            / destination_filename
        )

    elif (
        target_name
        and target_kind
    ):
        destination_path = (
            library_dir
            / (
                instruction_registry_name(
                    target_name,
                    target_kind,
                )
                + ".docx"
            )
        )

    else:
        destination_path = (
            library_dir
            / source_path.name
        )

    shutil.copy2(
        source_path,
        destination_path,
    )

    return destination_path


def print_library_info() -> None:
    library_dir = get_library_dir()

    print(
        "Библиотека ИОТ:",
        library_dir,
    )

    print(
        "Существует:",
        library_dir.exists(),
    )

    print(
        "Файлов ИОТ:",
        len(
            list_instruction_files()
        ),
    )


if __name__ == "__main__":
    print_library_info()
