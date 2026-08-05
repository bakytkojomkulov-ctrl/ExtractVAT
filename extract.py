import re
import sys
import traceback
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tkinter import Tk, filedialog, messagebox

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# ============================================================
# СЛУЖЕБНЫЕ ФУНКЦИИ
# ============================================================

def application_directory() -> Path:
    """
    Возвращает папку, где находится EXE-файл
    или исходный Python-файл.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    return Path(__file__).resolve().parent


def clean_spaces(value: str) -> str:
    """
    Убирает повторяющиеся пробелы и переносы строк.
    """
    if not value:
        return ""

    value = value.replace("\u00a0", " ")
    value = value.replace("\r", " ")
    value = value.replace("\n", " ")
    value = re.sub(r"\s+", " ", value)

    return value.strip()


def normalize_text_for_search(text: str) -> str:
    """
    Делает текст удобным для поиска регулярными выражениями.
    """
    text = text.replace("\u00a0", " ")
    text = text.replace("–", "-")
    text = text.replace("—", "-")
    text = re.sub(r"[ \t]+", " ", text)

    return text


def normalize_number(value: str) -> str:
    """
    Приводит денежное значение к виду 12345.67.
    """
    value = value.strip()
    value = value.replace("\u00a0", "")
    value = value.replace(" ", "")
    value = value.replace(",", ".")

    return value


def parse_decimal(value: str):
    """
    Преобразует найденную строку в Decimal.
    """
    normalized = normalize_number(value)

    try:
        return Decimal(normalized)
    except InvalidOperation:
        return None


def normalize_date(day: str, month: str, year: str) -> str:
    """
    Возвращает дату в формате ДД.ММ.ГГГГ.
    """
    try:
        day_number = int(day)
        month_number = int(month)
        year_number = int(year)

        if not 1 <= day_number <= 31:
            return ""

        if not 1 <= month_number <= 12:
            return ""

        if not 2000 <= year_number <= 2100:
            return ""

        return f"{day_number:02d}.{month_number:02d}.{year_number:04d}"

    except ValueError:
        return ""


def convert_date_string(value: str) -> str:
    """
    Преобразует 01-08-2026, 01/08/2026 или 01.08.2026
    в формат 01.08.2026.
    """
    if not value:
        return ""

    match = re.search(
        r"\b(\d{1,2})[./\-](\d{1,2})[./\-](\d{4})\b",
        value
    )

    if not match:
        return ""

    return normalize_date(
        match.group(1),
        match.group(2),
        match.group(3)
    )


# ============================================================
# ЧТЕНИЕ PDF
# ============================================================

def extract_pdf_text(pdf_path: Path) -> tuple[str, str]:
    """
    Извлекает текст PDF двумя способами.

    Обычный текст используется для большинства полей.
    Layout-текст помогает сохранить расположение колонок.
    """
    ordinary_parts = []
    layout_parts = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            ordinary_text = page.extract_text(
                x_tolerance=2,
                y_tolerance=3
            )

            layout_text = page.extract_text(
                x_tolerance=2,
                y_tolerance=3,
                layout=True
            )

            if ordinary_text:
                ordinary_parts.append(ordinary_text)

            if layout_text:
                layout_parts.append(layout_text)

    return (
        "\n".join(ordinary_parts),
        "\n".join(layout_parts)
    )


# ============================================================
# ИЗВЛЕЧЕНИЕ ПОЛЕЙ
# ============================================================

def extract_invoice_number(text: str) -> str:
    """
    Извлекает поле 102 — номер счета-фактуры.
    """
    patterns = [
        r"102\s*Номер\s*:\s*([0-9]{4,}-[0-9]{3}-[0-9]{5,})",
        r"\b(000\d{3,}-\d{3}-\d{5,})\b"
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)

        if match:
            return match.group(1).strip()

    return ""


def extract_supplier_inn(text: str) -> str:
    """
    Извлекает поле 201 — ИНН поставщика.
    """
    match = re.search(
        r"201\s+Поставщик\s+ИНН\s*:\s*(\d{10,16})",
        text,
        re.IGNORECASE
    )

    if match:
        return match.group(1)

    return ""


def extract_buyer_inn(text: str) -> str:
    """
    Извлекает поле 301 — ИНН покупателя.
    """
    match = re.search(
        r"301\s+Покупатель\s+ИНН\s*:\s*(\d{10,16})",
        text,
        re.IGNORECASE
    )

    if match:
        return match.group(1)

    return ""


def simplify_organization_name(name: str) -> str:
    """
    Сокращает распространённые организационно-правовые формы.
    """
    name = clean_spaces(name)

    replacements = [
        (
            r"^Общество\s+с\s+ограниченной\s+ответственностью\s*",
            "ООО "
        ),
        (
            r"^Открытое\s+акционерное\s+общество\s*",
            "ОАО "
        ),
        (
            r"^Закрытое\s+акционерное\s+общество\s*",
            "ЗАО "
        ),
        (
            r"^Индивидуальный\s+предприниматель\s*",
            "ИП "
        )
    ]

    for pattern, replacement in replacements:
        name = re.sub(
            pattern,
            replacement,
            name,
            flags=re.IGNORECASE
        )

    name = re.sub(r"\s+", " ", name).strip()

    return name


def extract_supplier_name(text: str) -> str:
    """
    Извлекает наименование поставщика.

    Используется участок между строкой
    «Ф.И.О. ИП/Наименование организации»
    и следующими служебными кодами формы.
    """
    normalized = normalize_text_for_search(text)

    start_patterns = [
        r"Ф\.?\s*И\.?\s*О\.?\s*ИП\s*/\s*Наименование\s+организации\s*:",
        r"Ф\.?\s*И\.?\s*О\.?\s+ИП\s*/\s*Наименование\s+организации\s*:"
    ]

    start_match = None

    for pattern in start_patterns:
        start_match = re.search(
            pattern,
            normalized,
            re.IGNORECASE
        )

        if start_match:
            break

    if not start_match:
        return ""

    fragment = normalized[start_match.end():start_match.end() + 500]

    stop_patterns = [
        r"\s+202\b",
        r"\s+302\b",
        r"\s+203\b",
        r"\s+Филиал\s+поставщика",
        r"\s+Ф\.?\s*И\.?\s*О\.?\s*ИП\s*/"
    ]

    end_position = len(fragment)

    for pattern in stop_patterns:
        stop_match = re.search(
            pattern,
            fragment,
            re.IGNORECASE
        )

        if stop_match:
            end_position = min(end_position, stop_match.start())

    supplier = fragment[:end_position]
    supplier = clean_spaces(supplier)

    # Удаляем случайно захваченные служебные цифры.
    supplier = re.sub(r"^\d+\s*", "", supplier)
    supplier = re.sub(r"\s+\d+$", "", supplier)

    return simplify_organization_name(supplier)


def extract_issue_date(text: str) -> str:
    """
    Извлекает дату оформления.

    В некоторых PDF дата распознаётся так:
    0 3 0 8 2 0 2 6

    Поэтому анализируется участок между
    «Дата оформления» и разделом реквизитов.
    """
    normalized = normalize_text_for_search(text)

    marker = re.search(
        r"103\s+Дата\s+оформления",
        normalized,
        re.IGNORECASE
    )

    if marker:
        fragment = normalized[marker.end():marker.end() + 250]

        # Сначала ищем обычную дату.
        normal_date = re.search(
            r"\b(\d{1,2})[./\-](\d{1,2})[./\-](\d{4})\b",
            fragment
        )

        if normal_date:
            return normalize_date(
                normal_date.group(1),
                normal_date.group(2),
                normal_date.group(3)
            )

        # Затем собираем раздельно распознанные цифры.
        before_supplier = re.split(
            r"201\s+Поставщик",
            fragment,
            maxsplit=1,
            flags=re.IGNORECASE
        )[0]

        digits = re.findall(r"\d", before_supplier)

        if len(digits) >= 8:
            date_digits = "".join(digits[:8])

            date_value = normalize_date(
                date_digits[0:2],
                date_digits[2:4],
                date_digits[4:8]
            )

            if date_value:
                return date_value

    return ""


def extract_delivery_date(text: str) -> str:
    """
    Извлекает поле «Дата поставки».
    """
    match = re.search(
        r"Дата\s+поставки\s*:\s*"
        r"(\d{1,2}[./\-]\d{1,2}[./\-]\d{4})",
        text,
        re.IGNORECASE
    )

    if match:
        return convert_date_string(match.group(1))

    return ""


def extract_contract(text: str) -> str:
    """
    Извлекает номер договора.

    Сначала анализирует участок рядом с полем договора,
    затем выполняет резервный поиск по всему документу.
    """
    normalized = normalize_text_for_search(text)

    marker = re.search(
        r"Договор\s*\(\s*контракт\s*\)",
        normalized,
        re.IGNORECASE
    )

    candidates = []

    if marker:
        start = max(0, marker.start() - 700)
        end = min(len(normalized), marker.end() + 700)
        fragment = normalized[start:end]

        candidates.extend(
            re.findall(
                r"№\s*[A-Za-zА-Яа-яЁё]?[A-Za-zА-Яа-яЁё0-9._\-]*"
                r"/[A-Za-zА-Яа-яЁё0-9._\-]+"
                r"/\d{2,4}",
                fragment
            )
        )

    if not candidates:
        candidates.extend(
            re.findall(
                r"№\s*[A-Za-zА-Яа-яЁё]?[A-Za-zА-Яа-яЁё0-9._\-]*"
                r"/[A-Za-zА-Яа-яЁё0-9._\-]+"
                r"/\d{2,4}",
                normalized
            )
        )

    for candidate in candidates:
        candidate = clean_spaces(candidate)
        candidate = candidate.replace("№ ", "№")

        # Номер счета-фактуры имеет другой формат и сюда не подходит.
        if candidate:
            return candidate

    return ""


def extract_vat_from_total(text: str):
    """
    Извлекает сумму НДС из строки:
    «Итого по счету-фактуре».

    После названия строки обычно идут:
    1. Стоимость без налогов
    2. Сумма НДС
    3. Сумма НсП
    4. Общая стоимость

    Поэтому берётся второе денежное значение.
    """
    normalized = normalize_text_for_search(text)

    marker = re.search(
        r"Итого\s+по\s+счету\s*-\s*фактуре\s*:",
        normalized,
        re.IGNORECASE
    )

    if not marker:
        marker = re.search(
            r"Итого\s+по\s+счетуфактуре\s*:",
            normalized,
            re.IGNORECASE
        )

    if not marker:
        return None

    fragment = normalized[marker.end():marker.end() + 350]

    numbers = re.findall(
        r"(?<!\d)"
        r"\d[\d \u00a0]*[.,]\d{2,5}"
        r"(?!\d)",
        fragment
    )

    parsed = []

    for number in numbers:
        decimal_value = parse_decimal(number)

        if decimal_value is not None:
            parsed.append(decimal_value)

    if len(parsed) >= 2:
        return parsed[1]

    return None


def process_pdf(pdf_path: Path) -> dict:
    """
    Обрабатывает один PDF и возвращает одну строку Excel.
    """
    ordinary_text, layout_text = extract_pdf_text(pdf_path)

    if not ordinary_text.strip() and not layout_text.strip():
        raise ValueError("PDF не содержит извлекаемого текста")

    combined_text = ordinary_text + "\n" + layout_text

    return {
        "file": pdf_path.name,
        "invoice_number": extract_invoice_number(combined_text),
        "issue_date": extract_issue_date(ordinary_text),
        "supplier": extract_supplier_name(ordinary_text),
        "supplier_inn": extract_supplier_inn(combined_text),
        "buyer_inn": extract_buyer_inn(combined_text),
        "delivery_date": extract_delivery_date(combined_text),
        "contract": extract_contract(ordinary_text),
        "vat": extract_vat_from_total(combined_text)
    }


# ============================================================
# СОЗДАНИЕ EXCEL
# ============================================================

def save_excel(rows: list[dict], errors: list[dict], output_path: Path):
    """
    Создаёт Result.xlsx.
    """
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Счета-фактуры"

    headers = [
        "Файл",
        "Номер СФ",
        "Дата",
        "Поставщик",
        "ИНН поставщика",
        "Покупатель ИНН",
        "Дата поставки",
        "Договор",
        "Сумма НДС"
    ]

    sheet.append(headers)

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="D9EAF7"
    )

    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    for row in rows:
        sheet.append([
            row["file"],
            row["invoice_number"],
            row["issue_date"],
            row["supplier"],
            row["supplier_inn"],
            row["buyer_inn"],
            row["delivery_date"],
            row["contract"],
            float(row["vat"]) if row["vat"] is not None else None
        ])

    # Сохраняем ИНН и номер СФ как текст,
    # чтобы Excel не удалял ведущие нули.
    for row_number in range(2, sheet.max_row + 1):
        sheet.cell(row=row_number, column=2).number_format = "@"
        sheet.cell(row=row_number, column=5).number_format = "@"
        sheet.cell(row=row_number, column=6).number_format = "@"
        sheet.cell(row=row_number, column=9).number_format = '#,##0.00'

    widths = {
        1: 38,
        2: 27,
        3: 14,
        4: 42,
        5: 20,
        6: 20,
        7: 16,
        8: 22,
        9: 18
    }

    for column_number, width in widths.items():
        column_letter = get_column_letter(column_number)
        sheet.column_dimensions[column_letter].width = width

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.row_dimensions[1].height = 28

    for row in sheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(
                vertical="center",
                wrap_text=True
            )

    # Отдельный лист для PDF, которые не удалось открыть.
    if errors:
        error_sheet = workbook.create_sheet("Ошибки")
        error_sheet.append(["Файл", "Ошибка"])

        for cell in error_sheet[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(
                fill_type="solid",
                fgColor="F4CCCC"
            )

        for error in errors:
            error_sheet.append([
                error["file"],
                error["error"]
            ])

        error_sheet.column_dimensions["A"].width = 45
        error_sheet.column_dimensions["B"].width = 80
        error_sheet.freeze_panes = "A2"

    workbook.save(output_path)


# ============================================================
# ИНТЕРФЕЙС
# ============================================================

def choose_folder(default_folder: Path):
    """
    Открывает окно выбора папки.
    """
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    selected = filedialog.askdirectory(
        title="Выберите папку с PDF-счетами-фактурами",
        initialdir=str(default_folder)
    )

    root.destroy()

    if not selected:
        return None

    return Path(selected)


def show_message(title: str, text: str, error: bool = False):
    """
    Показывает информационное окно.
    """
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    if error:
        messagebox.showerror(title, text)
    else:
        messagebox.showinfo(title, text)

    root.destroy()


# ============================================================
# ЗАПУСК
# ============================================================

def main():
    app_dir = application_directory()
    selected_folder = choose_folder(app_dir)

    if selected_folder is None:
        return

    # Обрабатываются PDF только в выбранной папке.
    pdf_files = sorted(
        selected_folder.glob("*.pdf"),
        key=lambda path: path.name.lower()
    )

    if not pdf_files:
        show_message(
            "PDF не найдены",
            "В выбранной папке нет PDF-файлов.",
            error=True
        )
        return

    rows = []
    errors = []

    for pdf_path in pdf_files:
        try:
            result = process_pdf(pdf_path)
            rows.append(result)

        except Exception as exc:
            errors.append({
                "file": pdf_path.name,
                "error": str(exc)
            })

    output_path = selected_folder / "Result.xlsx"
    save_excel(rows, errors, output_path)

    found_vat = sum(
        1 for row in rows
        if row["vat"] is not None
    )

    complete_rows = sum(
        1 for row in rows
        if all([
            row["invoice_number"],
            row["supplier_inn"],
            row["buyer_inn"]
        ])
    )

    show_message(
        "Обработка завершена",
        f"Обработано PDF: {len(pdf_files)}\n"
        f"Создано строк: {len(rows)}\n"
        f"Основные реквизиты найдены: {complete_rows}\n"
        f"Сумма НДС найдена: {found_vat}\n"
        f"Ошибок открытия: {len(errors)}\n\n"
        f"Результат сохранён:\n{output_path}"
    )


if __name__ == "__main__":
    try:
        main()

    except Exception:
        error_path = application_directory() / "Extract_error.txt"

        error_path.write_text(
            traceback.format_exc(),
            encoding="utf-8"
        )

        show_message(
            "Ошибка программы",
            "Произошла непредвиденная ошибка.\n\n"
            f"Подробности сохранены в файле:\n{error_path}",
            error=True
        )