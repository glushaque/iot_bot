from __future__ import annotations

import json
import os
import sys
from pathlib import Path


from services.pipeline import run_full_pipeline
from services.cost_tracker import get_session_summary


def main():
    if len(sys.argv) != 3:
        raise RuntimeError(
            "Использование: python telegram_worker.py "
            "<questionnaire.docx> <output_dir>"
        )

    questionnaire_path = Path(sys.argv[1])
    output_dir = Path(sys.argv[2])

    result = run_full_pipeline(
        questionnaire_path=questionnaire_path,
        source_docm="templates/ИСХОДНЫЙ_ФАЙЛ.docm",
        output_dir=output_dir,
        visible_word=False,
    )

    session_cost, session_calls = get_session_summary()

    result_json = output_dir / "telegram_result.json"
    result_json.write_text(
        json.dumps(
            {
                "output_path": str(result["output_path"]),
                "generated": result["generated"],
                "session_cost_rub": round(session_cost, 2),
                "session_calls": session_calls,
            },
            ensure_ascii=False,
            default=str,
        ),
        encoding="utf-8",
    )

    print("\nTELEGRAM_RESULT:", result["output_path"])
    print(f"Потрачено на эту сборку: {session_cost:.2f} ₽ ({session_calls} запросов к ИИ)")


if __name__ == "__main__":
    main()