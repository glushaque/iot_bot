import os

from dotenv import load_dotenv

load_dotenv()

from services.pipeline import run_full_pipeline


if __name__ == "__main__":
    result = run_full_pipeline(
        questionnaire_path="input/ОПРОСНЫЙ_ЛИСТ.docx",
        source_docm="templates/ИСХОДНЫЙ_ФАЙЛ.docm",
        output_dir="output",
        visible_word=False,
    )

    print()
    print(
        "Результат:",
        result["output_path"]
    )