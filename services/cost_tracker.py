import contextvars
import html
import json
import os
from datetime import date

PRICE_PER_1K_INPUT_RUB = 1.20
PRICE_PER_1K_OUTPUT_RUB = 1.20

_COST_FILE = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "runtime",
    "ai_cost.json",
)

_current_position = contextvars.ContextVar("current_position", default="")
_position_stats = contextvars.ContextVar("position_stats", default=None)

_session_cost = 0.0
_session_calls = 0


def set_current_position(name):
    position_token = _current_position.set(name)
    stats_token = _position_stats.set({
        "api_calls": 0,
        "input_tokens": 0,
        "output_tokens": 0,
        "cost": 0.0,
    })
    return (position_token, stats_token)


def reset_current_position(tokens):
    position_token, stats_token = tokens
    _finalize_position()
    _current_position.reset(position_token)
    _position_stats.reset(stats_token)


def _load():
    if not os.path.exists(_COST_FILE):
        return {"total_rub": 0.0, "history": []}
    try:
        with open(_COST_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return {"total_rub": 0.0, "history": []}


def _save(data):
    os.makedirs(os.path.dirname(_COST_FILE), exist_ok=True)
    with open(_COST_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)


def _finalize_position():
    stats = _position_stats.get()

    if not stats or stats["api_calls"] == 0:
        return

    position = _current_position.get() or "—"

    data = _load()
    data["total_rub"] = data.get("total_rub", 0.0) + stats["cost"]
    data.setdefault("history", []).append({
        "date": date.today().isoformat(),
        "position": position,
        "api_calls": stats["api_calls"],
        "input_tokens": stats["input_tokens"],
        "output_tokens": stats["output_tokens"],
        "cost_rub": round(stats["cost"], 2),
    })
    data["history"] = data["history"][-500:]
    _save(data)


def log_generation_cost(response):
    global _session_cost, _session_calls

    usage = getattr(response, "usage", None)
    if usage is None:
        return 0.0

    input_tokens = usage.prompt_tokens or 0
    output_tokens = usage.completion_tokens or 0

    cost = (
        input_tokens / 1000 * PRICE_PER_1K_INPUT_RUB
        + output_tokens / 1000 * PRICE_PER_1K_OUTPUT_RUB
    )

    stats = _position_stats.get()
    if stats is not None:
        stats["api_calls"] += 1
        stats["input_tokens"] += input_tokens
        stats["output_tokens"] += output_tokens
        stats["cost"] += cost

    _session_cost += cost
    _session_calls += 1

    print(
        f"Генерация ({_current_position.get() or '—'}): "
        f"{input_tokens} вх. + {output_tokens} исх. токенов ≈ {cost:.2f} ₽."
    )

    return cost


def get_session_summary():
    return _session_cost, _session_calls


def get_stats_text():
    data = _load()
    total = data.get("total_rub", 0.0)
    history = data.get("history", [])

    today = date.today().isoformat()
    today_items = [h for h in history if h["date"] == today]
    month_prefix = today[:7]
    month_items = [h for h in history if h["date"].startswith(month_prefix)]

    lines = [
        "Расходы на ИИ-генерацию инструкций:",
        f"Сегодня: {sum(h['cost_rub'] for h in today_items):.2f} ₽ ({len(today_items)} генераций)",
        f"За месяц: {sum(h['cost_rub'] for h in month_items):.2f} ₽ ({len(month_items)} генераций)",
        f"Всего: {total:.2f} ₽ ({len(history)} генераций)",
    ]

    if history:
        lines.append("")
        lines.append("Последние генерации:")
        lines.append("")

        table_lines = [
            f"{'API':>3}  {'ВХОД':>7}  {'ВЫХОД':>7}  {'ЦЕНА':>8}  ПОЗИЦИЯ"
        ]

        for h in history[-8:][::-1]:
            api = h.get("api_calls", "-")
            input_t = h.get("input_tokens", "-")
            output_t = h.get("output_tokens", "-")
            cost = h.get("cost_rub", 0)
            position = h.get("position", "—")

            table_lines.append(
                f"{str(api):>3}  {str(input_t):>7}  {str(output_t):>7}  "
                f"{cost:>7.2f}₽  {position}"
            )

        table_text = html.escape("\n".join(table_lines))
        lines.append(f"<pre>{table_text}</pre>")

    return "\n".join(lines)