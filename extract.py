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


def application_directory() -> Path:
    """Папка, где находится EXE или исходный Python-файл."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    return Path(__file__).resolve().parent


def normalize_number(value: str) -> str:
    """Приводит сумму к стандартному виду: 12345.67."""
    value = value.strip()
    value = value.replace("\u00a0", "")
    value = value.replace(" ", "")
    value = value.replace(",", ".")

    return value


def parse_decimal(value: str):
    """Преобразует найденное значение в число Excel."""
    normalized = normalize_number(value)

    try:
        return Decimal(normalized)
    except InvalidOperation:
        return None


def extract_text(pdf_path: Path) -> str:
    """Извлекает текст со всех страниц PDF."""
    parts = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text(
                x_tolerance=2,
                y_tolerance=3
            )

            if page_text:
                parts.append(page_text)

    return "\n".join(parts)


def extract_vat_from_total_line(text: str):
    """
    Ищет строку:
    Итого по счету-фактуре:
    стоимость без НДС | сумма НДС | сумма НсП | общая стоимость

    Сумма НДС — второе числовое значение после названия строки.
    """
    normalized_text = text.replace("\u00a0", " ")

    pattern = re.compile(
        r"Итого\s+по\s+счету[\s\-–—]*фактуре\s*:"
        r"\s*"
        r"([\d\s]+[.,]\d{2,5})"
        r"\s+"
        r"([\d\s]+[.,]\d{2,5})"
        r"\s+"
        r"([\d\s]+[.,]\d{2,5})"
        r"\s+"
        r"([\d\s]+[.,]\d{2,5})",
        re.IGNORECASE
    )

    match = pattern.search(normalized_text)

    if match:
        return parse_decimal(match.group(2)), "Найдено по итоговой строке"

    return None, "Итоговая строка не распознана"


def extract_vat_fallback(text: str):
    """
    Резервный поиск. Анализирует фрагмент возле строки
    «Итого по счету-фактуре».
    """
    normalized_text = re.sub(r"[ \t]+", " ", text)

    marker = re.search(
        r"Итого\s+по\s+счету[\s\-–—]*фактуре\s*:",
        normalized_text,
        re.IGNORECASE
    )

    if not marker:
        return None, "Строка «Итого по счету-фактуре» отсутствует"

    fragment = normalized_text[marker.end():marker.end() + 250]

    numbers = re.findall(
        r"\d[\d \u00a0]*[.,]\d{2,5}",
        fragment
    )

    parsed_numbers = []

    for number in numbers:
        parsed = parse_decimal(number)

        if parsed is not None:
            parsed_numbers.append(parsed)

    if len(parsed_numbers) >= 2:
        return parsed_numbers[1], "Найдено резервным способом"

    return None, "После итоговой строки найдено недостаточно чисел"


def process_pdf(pdf_path: Path):
    """Обрабатывает один PDF."""
    text = extract_text(pdf_path)

    if not text.strip():
        return None, "PDF не содержит извлекаемого текста"

    vat, status = extract_vat_from_total_line(text)

    if vat is not None:
        return vat, status

    return extract_vat_fallback(text)


def save_excel(rows, output_path: Path):
    """Сохраняет результаты в Result.xlsx."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "НДС"

    headers = [
        "№",
        "Файл",
        "Сумма НДС",
        "Статус"
    ]

    sheet.append(headers)

    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(
            fill_type="solid",
            fgColor="D9EAF7"
        )
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    for index, row in enumerate(rows, start=1):
        sheet.append([
            index,
            row["file"],
            float(row["vat"]) if row["vat"] is not None else None,
            row["status"]
        ])

    for cell in sheet["C"][1:]:
        cell.number_format = '#,##0.00'

    widths = {
        1: 8,
        2: 45,
        3: 18,
        4: 38
    }

    for column_number, width in widths.items():
        column_letter = get_column_letter(column_number)
        sheet.column_dimensions[column_letter].width = width

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    workbook.save(output_path)


def choose_folder(default_folder: Path) -> Path | None:
    """Позволяет выбрать папку с PDF."""
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
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    if error:
        messagebox.showerror(title, text)
    else:
        messagebox.showinfo(title, text)

    root.destroy()


def main():
    app_dir = application_directory()
    selected_folder = choose_folder(app_dir)

    if selected_folder is None:
        return

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

    for pdf_path in pdf_files:
        try:
            vat, status = process_pdf(pdf_path)

            rows.append({
                "file": pdf_path.name,
                "vat": vat,
                "status": status
            })

        except Exception as exc:
            rows.append({
                "file": pdf_path.name,
                "vat": None,
                "status": f"Ошибка: {exc}"
            })

    output_path = selected_folder / "Result.xlsx"
    save_excel(rows, output_path)

    successful = sum(
        1 for row in rows
        if row["vat"] is not None
    )

    show_message(
        "Обработка завершена",
        f"Обработано PDF: {len(rows)}\n"
        f"Сумма НДС найдена: {successful}\n"
        f"Не распознано: {len(rows) - successful}\n\n"
        f"Результат:\n{output_path}"
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
            "Произошла непредвиденная ошибка.\n"
            f"Подробности сохранены в:\n{error_path}",
            error=True
        )
