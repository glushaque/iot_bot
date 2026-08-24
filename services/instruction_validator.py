import re


EXPECTED_COUNTS = {
    1: 17,
    2: 10,
    3: 14,
    4: 13,
    5: 11
}


def extract_points(text):
    pattern = r"(?m)^(\d+)\.(\d+)\.\s*(.+?)(?=^\d+\.\d+\.|\Z)"

    matches = re.findall(
        pattern,
        text,
        flags=re.MULTILINE | re.DOTALL
    )

    points = []

    for section, number, content in matches:
        points.append({
            "section": int(section),
            "number": int(number),
            "content": content.strip()
        })

    return points


def validate_instruction(text):
    points = extract_points(text)

    errors = []
    section_counts = {}

    for section_number in EXPECTED_COUNTS:
        section_points = [
            point
            for point in points
            if point["section"] == section_number
        ]

        section_counts[section_number] = len(section_points)

        expected = EXPECTED_COUNTS[section_number]

        if len(section_points) != expected:
            errors.append(
                f"Раздел {section_number}: "
                f"ожидалось {expected}, "
                f"получено {len(section_points)}"
            )

    short_points = []

    for point in points:
        content_length = len(point["content"])

        if content_length < 200:
            short_points.append({
                "section": point["section"],
                "number": point["number"],
                "length": content_length
            })

            errors.append(
                f'Пункт {point["section"]}.{point["number"]}: '
                f"только {content_length} символов"
            )

    return {
        "valid": len(errors) == 0,
        "errors": errors,
        "section_counts": section_counts,
        "short_points": short_points,
        "points_total": len(points)
    }