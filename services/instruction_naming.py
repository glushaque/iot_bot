from __future__ import annotations

import re
import pymorphy3


morph = pymorphy3.MorphAnalyzer()

WORK_KINDS = {
    "regular_work",
    "high_risk_work",
}

PREPOSITIONS = {
    "по", "при", "на", "в", "во", "с", "со",
    "к", "ко", "из", "от", "до", "для", "без",
    "над", "под", "между",
}


def clean_name(value: str) -> str:
    return " ".join(str(value or "").split()).strip()


def lower_first(value: str) -> str:
    value = clean_name(value)

    if not value:
        return value

    return value[:1].lower() + value[1:]


def _normalized_word(word: str) -> str:
    return (
        word.lower()
        .strip(".,;:()[]{}«»\"'")
        .replace("ё", "е")
    )


def _restore_case(source: str, value: str) -> str:
    if source[:1].isupper():
        return value[:1].upper() + value[1:]

    return value


def position_genitive(name: str) -> str:
    """
    Родительный падеж должности.

    генеральный директор -> генерального директора
    заместитель генерального директора ->
        заместителя генерального директора
    менеджер по продажам -> менеджера по продажам
    """
    words = clean_name(name).split()

    result = []
    stop_inflecting = False

    for word in words:
        normalized = _normalized_word(word)

        if normalized in PREPOSITIONS:
            stop_inflecting = True
            result.append(word)
            continue

        if stop_inflecting:
            result.append(word)
            continue

        parses = morph.parse(normalized)

        if not parses:
            result.append(word)
            continue

        inflected = parses[0].inflect({"gent"})

        if inflected:
            result.append(
                _restore_case(word, inflected.word)
            )
        else:
            result.append(word)

    return lower_first(" ".join(result))


def _detect_work_plural(words: list[str]) -> bool:
    """
    Определяет число ГЛАВНОЙ части названия работы.

    Ключевой принцип:
    - "уборка территории" -> смотрим на первое существительное
      "уборка" и получаем единственное число;
    - "разгрузочные работы" -> прилагательное до существительного
      даёт уверенный признак множественного числа;
    - зависимое слово "территории" не должно заставлять всю фразу
      ошибочно считаться множественным числом.
    """
    plural_adjective_before_noun = False

    for word in words:
        normalized = _normalized_word(word)

        if normalized in PREPOSITIONS:
            break

        parses = morph.parse(normalized)

        if not parses:
            continue

        # До первого существительного собираем признак
        # согласованного прилагательного во множественном числе.
        for parse in parses:
            if (
                ("ADJF" in parse.tag or "PRTF" in parse.tag)
                and "plur" in parse.tag
            ):
                plural_adjective_before_noun = True
                break

        noun_parses = [
            parse
            for parse in parses
            if "NOUN" in parse.tag
        ]

        if not noun_parses:
            continue

        # Первое существительное считаем главным.
        if plural_adjective_before_noun:
            return True

        # Берём наиболее вероятный разбор первого существительного.
        best_noun = noun_parses[0]

        if "plur" in best_noun.tag:
            return True

        return False

    return False


def work_prepositional(name: str) -> str:
    """
    Предложный падеж названия работы.

    уборка территории -> уборке территории
    работа с оргтехникой -> работе с оргтехникой
    работа на персональном компьютере ->
        работе на персональном компьютере
    разгрузочные работы -> разгрузочных работах
    погрузочно-разгрузочные работы ->
        погрузочно-разгрузочных работах
    """
    words = clean_name(name).split()

    if not words:
        return ""

    phrase_is_plural = _detect_work_plural(words)

    result = []

    for index, word in enumerate(words):
        normalized = _normalized_word(word)

        if normalized in PREPOSITIONS:
            result.extend(words[index:])
            break

        parses = morph.parse(normalized)

        if not parses:
            result.append(word)
            continue

        selected = None

        if phrase_is_plural:
            # Для множественного числа выбираем разбор,
            # который действительно является plural.
            for parse in parses:
                if "plur" in parse.tag:
                    selected = parse
                    break

        if selected is None:
            selected = parses[0]

        if phrase_is_plural:
            if "ADJF" in selected.tag or "PRTF" in selected.tag:
                inflected = selected.inflect(
                    {"loct", "plur"}
                )

                if inflected:
                    result.append(
                        _restore_case(
                            word,
                            inflected.word,
                        )
                    )
                else:
                    result.append(word)

                continue

            if "NOUN" in selected.tag:
                inflected = selected.inflect(
                    {"loct", "plur"}
                )

                if inflected:
                    result.append(
                        _restore_case(
                            word,
                            inflected.word,
                        )
                    )
                else:
                    result.append(word)

                # После главного существительного зависимую часть
                # оставляем без изменения.
                result.extend(words[index + 1:])
                break

            result.append(word)
            continue

        # Единственное число:
        # склоняем первое существительное и прекращаем склонение
        # зависимой части.
        if "NOUN" in selected.tag:
            inflected = selected.inflect({"loct"})

            if inflected:
                result.append(
                    _restore_case(
                        word,
                        inflected.word,
                    )
                )
            else:
                result.append(word)

            result.extend(words[index + 1:])
            break

        # Если перед главным существительным стоит прилагательное,
        # склоняем его в предложный падеж единственного числа.
        if "ADJF" in selected.tag or "PRTF" in selected.tag:
            inflected = selected.inflect({"loct"})

            if inflected:
                result.append(
                    _restore_case(
                        word,
                        inflected.word,
                    )
                )
            else:
                result.append(word)

            continue

        result.append(word)

    return lower_first(" ".join(result))


def instruction_registry_name(
    target_name: str,
    target_kind: str,
) -> str:
    target_name = clean_name(target_name)

    if target_kind == "position":
        return (
            "ИОТ для "
            + position_genitive(target_name)
        )

    if target_kind in WORK_KINDS:
        return (
            "ИОТ при "
            + work_prepositional(target_name)
        )

    raise ValueError(
        f"Неизвестный тип ИОТ: {target_kind}"
    )


def instruction_title(
    target_name: str,
    target_kind: str,
    object_name: str = "obj_name",
    object_address: str = "address_object",
) -> str:
    object_name = clean_name(object_name) or "obj_name"
    object_address = (
        clean_name(object_address)
        or "address_object"
    )

    if target_kind == "position":
        core = (
            "по охране труда для "
            + position_genitive(target_name)
        )

    elif target_kind in WORK_KINDS:
        core = (
            "по охране труда при "
            + work_prepositional(target_name)
        )

    else:
        raise ValueError(
            f"Неизвестный тип ИОТ: {target_kind}"
        )

    return (
        f"{core} на объекте {object_name}, "
        f"расположенном по адресу: {object_address}"
    )


def canonical_library_filename(
    target_name: str,
    target_kind: str,
) -> str:
    filename = instruction_registry_name(
        target_name,
        target_kind,
    )

    for char in '<>:"/\\|?*':
        filename = filename.replace(char, "_")

    return filename.rstrip(". ") + ".docx"
