from __future__ import annotations

import asyncio
import json
import os
import shutil
import sys
from pathlib import Path

from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.types import (
    FSInputFile,
    KeyboardButton,
    Message,
    ReplyKeyboardMarkup,
)
from dotenv import load_dotenv

from services.questionnaire import parse_questionnaire


load_dotenv()

BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()

if not BOT_TOKEN:
    raise RuntimeError("В .env не задан TELEGRAM_BOT_TOKEN")

BASE_DIR = Path(__file__).resolve().parent
SESSIONS_DIR = BASE_DIR / "runtime" / "sessions"
WORKER_PATH = BASE_DIR / "telegram_worker.py"

SESSIONS_DIR.mkdir(parents=True, exist_ok=True)

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

generation_lock = asyncio.Lock()
sessions: dict[int, dict] = {}

MENU = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="📄 Загрузить опросный лист")],
        [
            KeyboardButton(text="📋 Посмотреть список"),
            KeyboardButton(text="🔄 Очистить"),
        ],
    ],
    resize_keyboard=True,
)


def _new_session(user_id: int) -> dict:
    session_dir = SESSIONS_DIR / str(user_id)
    session_dir.mkdir(parents=True, exist_ok=True)

    return {
        "dir": session_dir,
        "questionnaire_path": None,
        "last_result": None,
    }


def _session(user_id: int) -> dict:
    if user_id not in sessions:
        sessions[user_id] = _new_session(user_id)
    return sessions[user_id]


def _lower_first(value: str) -> str:
    value = " ".join(str(value or "").split()).strip()
    if not value:
        return value
    return value[:1].lower() + value[1:]


def _is_docx(message: Message) -> bool:
    if not message.document:
        return False

    return (
        (message.document.file_name or "")
        .lower()
        .endswith(".docx")
    )


def _format_questionnaire_list(questionnaire: dict) -> str:
    positions = [
        _lower_first(value)
        for value in questionnaire.get("positions", [])
        if str(value or "").strip()
    ]

    regular_works = [
        _lower_first(value)
        for value in questionnaire.get("regular_works", [])
        if str(value or "").strip()
    ]

    high_risk_works = [
        _lower_first(value)
        for value in questionnaire.get("high_risk_works", [])
        if str(value or "").strip()
    ]

    lines = [f"Должности: {len(positions)}"]

    for index, value in enumerate(positions, start=1):
        lines.append(f"{index}. {value}")

    lines.append("")
    lines.append(f"Виды работ: {len(regular_works)}")

    for index, value in enumerate(regular_works, start=1):
        lines.append(f"{index}. {value}")

    if high_risk_works:
        lines.append("")
        lines.append(
            "Работы повышенной опасности: "
            f"{len(high_risk_works)}"
        )
        for index, value in enumerate(high_risk_works, start=1):
            lines.append(f"{index}. {value}")

    return "\n".join(lines)


async def _run_worker(
    questionnaire_path: Path,
    output_dir: Path,
) -> tuple[int, str]:
    process = await asyncio.create_subprocess_exec(
        sys.executable,
        str(WORKER_PATH),
        str(questionnaire_path),
        str(output_dir),
        cwd=str(BASE_DIR),
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.STDOUT,
    )

    stdout, _ = await process.communicate()

    return (
        process.returncode,
        stdout.decode("utf-8", errors="replace"),
    )


def _read_worker_result(output_dir: Path) -> dict:
    path = output_dir / "telegram_result.json"

    if not path.exists():
        raise RuntimeError("telegram_result.json не создан.")

    return json.loads(path.read_text(encoding="utf-8"))


def _format_generated(generated: list[dict]) -> str:
    if not generated:
        return (
            "Новых ИОТ не создавалось: все необходимые "
            "инструкции уже были в библиотеке."
        )

    lines = ["Новые ИОТ, которых не было в библиотеке:"]

    for index, item in enumerate(generated, start=1):
        kind = item.get("kind", "")
        prefix = "Должность" if kind == "position" else "Вид работ"
        lines.append(f"{index}. {prefix}: {item['name']}")

    return "\n".join(lines)


@dp.message(Command("start"))
async def command_start(message: Message):
    _session(message.from_user.id)

    await message.answer(
        "Загрузите заполненный опросный лист DOCX.\n\n"
        "Должности и виды работ бот автоматически возьмёт "
        "из опросного листа.",
        reply_markup=MENU,
    )


@dp.message(F.text == "📄 Загрузить опросный лист")
async def ask_questionnaire(message: Message):
    await message.answer(
        "Отправьте заполненный ОПРОСНЫЙ_ЛИСТ.docx.",
        reply_markup=MENU,
    )


@dp.message(F.text == "📋 Посмотреть список")
async def show_list(message: Message):
    session = _session(message.from_user.id)
    questionnaire_path = session.get("questionnaire_path")

    if not questionnaire_path:
        await message.answer(
            "Сначала загрузите опросный лист.",
            reply_markup=MENU,
        )
        return

    try:
        questionnaire = parse_questionnaire(questionnaire_path)
    except Exception as error:
        await message.answer(
            "Не удалось прочитать опросный лист.\n"
            f"Ошибка: {error}",
            reply_markup=MENU,
        )
        return

    await message.answer(
        _format_questionnaire_list(questionnaire),
        reply_markup=MENU,
    )


@dp.message(F.text == "🔄 Очистить")
async def clear_session(message: Message):
    user_id = message.from_user.id
    old = sessions.pop(user_id, None)

    if old:
        try:
            shutil.rmtree(old["dir"], ignore_errors=True)
        except Exception:
            pass

    sessions[user_id] = _new_session(user_id)

    await message.answer(
        "Текущий опросный лист и результаты очищены.",
        reply_markup=MENU,
    )


@dp.message(F.document)
async def handle_document(message: Message):
    session = _session(message.from_user.id)

    if not _is_docx(message):
        await message.answer(
            "Нужен файл в формате .docx.",
            reply_markup=MENU,
        )
        return

    questionnaire_path = session["dir"] / "ОПРОСНЫЙ_ЛИСТ.docx"

    telegram_file = await bot.get_file(message.document.file_id)
    await bot.download_file(
        telegram_file.file_path,
        destination=questionnaire_path,
    )

    try:
        questionnaire = parse_questionnaire(questionnaire_path)
    except Exception as error:
        questionnaire_path.unlink(missing_ok=True)
        await message.answer(
            "Не удалось разобрать опросный лист.\n"
            f"Ошибка: {error}",
            reply_markup=MENU,
        )
        return

    session["questionnaire_path"] = questionnaire_path

    await message.answer(
        "Опросный лист принят.\n\n"
        + _format_questionnaire_list(questionnaire)
        + "\n\nНачинаю формирование комплекта.",
        reply_markup=MENU,
    )

    output_dir = session["dir"] / "output"

    if output_dir.exists():
        shutil.rmtree(output_dir, ignore_errors=True)

    output_dir.mkdir(parents=True, exist_ok=True)

    status = await message.answer(
        "Проверяю библиотеку ИОТ. Отсутствующие "
        "инструкции будут созданы автоматически."
    )

    try:
        async with generation_lock:
            return_code, log = await _run_worker(
                questionnaire_path,
                output_dir,
            )

        if return_code != 0:
            print("\n=== TELEGRAM WORKER ERROR ===\n" + log)
            await status.edit_text(
                "Не удалось сформировать комплект. "
                "Подробности записаны в консоль."
            )
            return

        worker_result = _read_worker_result(output_dir)

        candidates = sorted(
            output_dir.glob("ИТОГОВЫЙ_КОМПЛЕКТ_*.docm"),
            key=lambda path: path.stat().st_mtime,
            reverse=True,
        )

        if not candidates:
            raise RuntimeError("Итоговый DOCM не найден.")

        result_file = candidates[0]
        session["last_result"] = worker_result

        print("\n=== TELEGRAM WORKER OK ===\n" + log)

        await status.edit_text("Комплект сформирован.")

        await message.answer_document(
            document=FSInputFile(result_file),
            caption=(
                "DOCM готов для ручной проверки.\n\n"
                "Проверьте данные. После проверки запустите "
                "ReplaceTags в Word."
            ),
        )

        await message.answer(
            _format_generated(worker_result.get("generated", [])),
            reply_markup=MENU,
        )

        await status.delete()

    except Exception as error:
        print(
            "\n=== TELEGRAM ERROR ===\n"
            f"{type(error).__name__}: {error}\n"
        )
        await status.edit_text(
            "При формировании произошла ошибка. "
            "Подробности записаны в консоль."
        )


@dp.message(F.text)
async def handle_other_text(message: Message):
    await message.answer(
        "Используйте кнопки меню.",
        reply_markup=MENU,
    )


async def main():
    print("Telegram-бот запущен.")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
