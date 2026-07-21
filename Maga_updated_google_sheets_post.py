import sys
from datetime import datetime
from io import BytesIO, StringIO
from pathlib import Path

import pandas as pd
import requests


MIPT_URL = (
    "https://priem.mipt.ru/applications_v2/"
    "bWFzdGVyL1JhZGlvdGVraG5pa2EgaSBrb21weXV0ZXJueWUgdGVraG5vbG9naWlfQnl1ZHpoZXQuaHRtbA=="
)

GOOGLE_SPREADSHEET_ID = "19ksmM7HZ8TXO85FDsob48NxsHz9Ue9jY68JsO0tCHc0"
GOOGLE_SHEET_NAME = "Расшифровки"
GOOGLE_SHEET_CSV_URL = (
    f"https://docs.google.com/spreadsheets/d/{GOOGLE_SPREADSHEET_ID}/"
    "gviz/tq"
)

APPS_SCRIPT_URL = (
    "https://script.google.com/macros/s/"
    "AKfycby9Iba_CW7CLVUZhBrCj9z03hB11Ooaomgq6XL3AJzDhQrEB-as8Je2zDWO6vKMJaBz3g/"
    "exec"
)

# Должен совпадать со значением SECRET_KEY в Apps Script.
# Оставьте пустым только если в Apps Script проверка ключа не используется.
APPS_SCRIPT_SECRET = "ABOBA_SECRET_KEY"

EXCEL_FILE = Path("mipt_scores.xlsx")
SHEET_NAME = "Выгрузки"

REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 Chrome/150.0 Safari/537.36"
    )
}


def normalize_column_name(column: object) -> str:
    """Приводит обычные и многоуровневые заголовки pandas к одной строке."""
    if isinstance(column, tuple):
        parts = [
            str(part).strip()
            for part in column
            if str(part).strip().lower() != "nan"
        ]
        return " ".join(parts)

    return str(column).strip()


def find_column(columns, *variants: str):
    """Ищет столбец по точному названию или по вхождению текста."""
    normalized_variants = [variant.casefold().strip() for variant in variants]

    for column in columns:
        name = str(column).casefold().strip()
        if name in normalized_variants:
            return column

    for column in columns:
        name = str(column).casefold().strip()
        if any(variant in name for variant in normalized_variants):
            return column

    return None


def normalize_code(value: object) -> str:
    """Нормализует код и убирает окончание .0, появившееся после Excel/CSV."""
    if value is None or pd.isna(value):
        return ""

    code = str(value).strip()
    if code.endswith(".0") and code[:-2].isdigit():
        code = code[:-2]

    return code


def convert_score(value):
    """Преобразует значение таблицы в int/float либо None."""
    number = pd.to_numeric(value, errors="coerce")

    if pd.isna(number):
        return None
    if float(number).is_integer():
        return int(number)
    return float(number)


def get_people_from_google_sheet() -> list[dict[str, str]]:
    """Загружает имена и коды из листа «Расшифровки» Google Таблицы."""
    response = requests.get(
        GOOGLE_SHEET_CSV_URL,
        params={
            "tqx": "out:csv",
            "sheet": GOOGLE_SHEET_NAME,
        },
        headers=REQUEST_HEADERS,
        timeout=30,
    )
    response.raise_for_status()

    # При закрытой таблице Google может вернуть HTML страницы входа вместо CSV.
    content_type = response.headers.get("Content-Type", "").casefold()
    if "text/html" in content_type:
        raise RuntimeError(
            "Google Таблица недоступна без авторизации. "
            "Откройте доступ: «Все, у кого есть ссылка — Читатель»."
        )

    # BytesIO исключает искажение русских символов из-за неверной кодировки requests.
    table = pd.read_csv(BytesIO(response.content), dtype=str)
    table.columns = [normalize_column_name(column) for column in table.columns]

    code_column = find_column(table.columns, "Уникальный код")
    name_column = find_column(table.columns, "Фамилия, Имя")

    if code_column is None:
        raise ValueError(
            f"В листе «{GOOGLE_SHEET_NAME}» отсутствует столбец "
            "«Уникальный код»."
        )

    people: list[dict[str, str]] = []
    seen_codes: set[str] = set()

    for _, row in table.iterrows():
        code = normalize_code(row[code_column])

        # Пропускаем пустые значения, ошибки формул и повторяющиеся коды.
        if not code or code.startswith("#"):
            continue
        if code in seen_codes:
            continue

        name = ""
        if name_column is not None and not pd.isna(row[name_column]):
            name = str(row[name_column]).strip()

        people.append({
            "Фамилия, Имя": name,
            "Уникальный код": code,
        })
        seen_codes.add(code)

    if not people:
        raise ValueError(
            f"В листе «{GOOGLE_SHEET_NAME}» не найдено ни одного корректного кода."
        )

    return people


def get_scores_for_people(people: list[dict[str, str]]) -> list[dict]:
    """Находит результаты МФТИ по кодам людей из Google Таблицы."""
    response = requests.get(MIPT_URL, headers=REQUEST_HEADERS, timeout=30)
    response.raise_for_status()
    response.encoding = response.apparent_encoding

    tables = pd.read_html(StringIO(response.text))
    requested_codes = {person["Уникальный код"] for person in people}
    found_results: dict[str, dict] = {}

    for table in tables:
        table.columns = [normalize_column_name(column) for column in table.columns]

        code_column = find_column(table.columns, "Уникальный код")
        position_column = find_column(table.columns, "№")
        total_column = find_column(table.columns, "Сумма баллов")
        exam_column = find_column(table.columns, "Сумма баллов по предметам")
        id_column = find_column(
            table.columns,
            "Сумма баллов за инд.дост.(конкурсные)",
            "Сумма баллов за индивидуальные достижения",
        )

        if code_column is None or total_column is None:
            continue

        table[code_column] = table[code_column].map(normalize_code)
        matching_rows = table[table[code_column].isin(requested_codes)]

        for _, row in matching_rows.iterrows():
            code = row[code_column]
            found_results[code] = {
                "Позиция": convert_score(row[position_column]) if position_column else None,
                "Сумма баллов": convert_score(row[total_column]),
                "Баллы за экзамен": convert_score(row[exam_column]) if exam_column else None,
                "Баллы за ИД": convert_score(row[id_column]) if id_column else None,
            }

    checked_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    results: list[dict] = []

    for person in people:
        code = person["Уникальный код"]
        found = found_results.get(code)

        if found is None:
            found = {
                "Позиция": None,
                "Сумма баллов": None,
                "Баллы за экзамен": None,
                "Баллы за ИД": None,
            }
            status = "Код не найден"
        elif found["Сумма баллов"] is None:
            status = "Баллы не заполнены"
        else:
            status = "Найден"

        results.append(
            {
                "Дата и время проверки": checked_at,
                "Уникальный код": code,
                "Позиция": found["Позиция"],
                "Сумма баллов": found["Сумма баллов"],
                "Баллы за экзамен": found["Баллы за экзамен"],
                "Баллы за ИД": found["Баллы за ИД"],
                "Статус": status,
                "Источник": MIPT_URL,
            }
        )

    return results



def upload_results_to_google_sheet(results: list[dict]) -> int:
    """Отправляет результаты POST-запросом в Apps Script."""
    payload = {
        "key": APPS_SCRIPT_SECRET,
        "results": results,
    }

    response = requests.post(
        APPS_SCRIPT_URL,
        json=payload,
        headers=REQUEST_HEADERS,
        timeout=60,
        allow_redirects=True,
    )
    response.raise_for_status()

    try:
        data = response.json()
    except requests.JSONDecodeError as error:
        preview = response.text[:500].replace("\n", " ")
        raise RuntimeError(
            "Apps Script вернул не JSON. "
            f"Первые символы ответа: {preview}"
        ) from error

    if not data.get("ok"):
        raise RuntimeError(
            f'Apps Script отклонил данные: {data.get("error", "неизвестная ошибка")}'
        )

    return int(data.get("written", 0))


def save_results_to_excel(
    results: list[dict],
    filename: Path = EXCEL_FILE,
) -> None:
    """Сохраняет актуальную выгрузку в Excel."""
    data = pd.DataFrame(results)

    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        data.to_excel(writer, sheet_name=SHEET_NAME, index=False)

        worksheet = writer.sheets[SHEET_NAME]
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        # Структура столбцов полностью совпадает с исходным скриптом:
        # A — дата проверки, B — уникальный код, C — позиция и т. д.
        widths = {
            "A": 22,
            "B": 18,
            "C": 12,
            "D": 16,
            "E": 18,
            "F": 16,
            "G": 22,
            "H": 80,
        }

        # Коды сохраняются как текст, чтобы Excel не менял их
        # и не удалял возможные ведущие нули.
        for cell in worksheet["B"][1:]:
            if cell.value is not None:
                cell.value = str(cell.value)
                cell.number_format = "@"

        for column, width in widths.items():
            worksheet.column_dimensions[column].width = width


def main() -> None:
    try:
        print(f"Загружаю коды из листа «{GOOGLE_SHEET_NAME}» Google Таблицы...")
        people = get_people_from_google_sheet()
        print(f"Загружено записей: {len(people)}")

        results = get_scores_for_people(people)

        for result in results:
            position = result["Позиция"] if result["Позиция"] is not None else "—"
            total = (
                result["Сумма баллов"]
                if result["Сумма баллов"] is not None
                else "—"
            )
            exam = (
                result["Баллы за экзамен"]
                if result["Баллы за экзамен"] is not None
                else "—"
            )
            individual = (
                result["Баллы за ИД"]
                if result["Баллы за ИД"] is not None
                else "—"
            )

            print(
                f'{result["Уникальный код"]}: '
                f"позиция {position}, всего {total}, экзамен {exam}, "
                f'ИД {individual} — {result["Статус"]}'
            )

        save_results_to_excel(results)
        print(f"\nАктуальная выгрузка сохранена: {EXCEL_FILE.resolve()}")

        print("Отправляю результаты в Google Таблицу через Apps Script...")
        written = upload_results_to_google_sheet(results)
        print(f"В Google Таблицу записано строк: {written}")

    except requests.RequestException as error:
        print(f"Ошибка загрузки данных: {error}")
        sys.exit(1)
    except Exception as error:
        print(f"Ошибка: {error}")
        sys.exit(2)


if __name__ == "__main__":
    main()
