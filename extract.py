import re
import sys
import traceback
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tkinter import Tk, filedialog, messagebox

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    PatternFill,
    Side
)
from openpyxl.utils import get_column_letter


# ============================================================
# ОБЩИЕ ФУНКЦИИ
# ============================================================

def application_directory() -> Path:
    """
    Возвращает папку, где находится Extract.exe
    или исходный файл extract.py.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    return Path(__file__).resolve().parent


def clean_spaces(value: str) -> str:
    """
    Убирает лишние пробелы, неразрывные пробелы
    и переносы строк.
    """
    if not value:
        return ""

    value = value.replace("\u00a0", " ")
    value = value.replace("\r", " ")
    value = value.replace("\n", " ")
    value = re.sub(r"\s+", " ", value)

    return value.strip()


def normalize_text(text: str) -> str:
    """
    Нормализует текст для поиска.
    """
    if not text:
        return ""

    text = text.replace("\u00a0", " ")
    text = text.replace("–", "-")
    text = text.replace("—", "-")
    text = text.replace("−", "-")
    text = re.sub(r"[ \t]+", " ", text)

    return text


def digits_only(value: str) -> str:
    """
    Оставляет в строке только цифры.
    """
    return re.sub(r"\D", "", value or "")


def normalize_money(value: str) -> str:
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
    Преобразует строку в Decimal.
    """
    normalized = normalize_money(value)

    try:
        return Decimal(normalized)
    except (InvalidOperation, ValueError):
        return None


def normalize_date(day: str, month: str, year: str) -> str:
    """
    Проверяет дату и возвращает её в формате ДД.ММ.ГГГГ.
    """
    try:
        date_value = datetime(
            int(year),
            int(month),
            int(day)
        )

        return date_value.strftime("%d.%m.%Y")

    except (ValueError, TypeError):
        return ""


def convert_date_string(value: str) -> str:
    """
    Преобразует:
    03-08-2026
    03/08/2026
    03.08.2026

    в 03.08.2026.
    """
    match = re.search(
        r"\b(\d{1,2})[./\-](\d{1,2})[./\-](\d{4})\b",
        value or ""
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
    Извлекает текст двумя способами:

    ordinary_text — обычный текст;
    layout_text — текст с приблизительным сохранением расположения.
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
# НОМЕР СЧЕТА-ФАКТУРЫ
# ============================================================

def extract_invoice_number(text: str) -> str:
    """
    Извлекает номер счета-фактуры из поля 102.
    """
    normalized = normalize_text(text)

    patterns = [
        r"102\s*Номер\s*:\s*([0-9]{4,}-[0-9]{3}-[0-9]{5,})",
        r"\b([0-9]{7}-[0-9]{3}-[0-9]{8})\b",
        r"\b([0-9]{4,}-[0-9]{3}-[0-9]{5,})\b"
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            normalized,
            re.IGNORECASE
        )

        if match:
            return match.group(1).strip()

    return ""


# ============================================================
# ИНН ПОСТАВЩИКА — УСИЛЕННАЯ ПРОВЕРКА
# ============================================================

def is_valid_supplier_inn(value: str) -> bool:
    """
    Проверяет формат ИНН поставщика.

    В обрабатываемых ЭСФ ИНН обычно состоит из 14 цифр.
    Допускается диапазон 10–16 цифр для совместимости,
    но приоритет всегда отдаётся 14-значному значению.
    """
    inn = digits_only(value)

    if not inn:
        return False

    if len(inn) < 10 or len(inn) > 16:
        return False

    if inn == "0" * len(inn):
        return False

    return True


def extract_supplier_inn_strict(text: str) -> str:
    """
    Основной способ.

    Ищет цифры непосредственно после:
    201 Поставщик ИНН:
    """
    normalized = normalize_text(text)

    patterns = [
        # Нормальное извлечение:
        # 201 Поставщик ИНН: 03004202410293
        r"201\s+Поставщик\s+ИНН\s*:\s*(\d{10,16})",

        # Между цифрами могли появиться пробелы:
        r"201\s+Поставщик\s+ИНН\s*:\s*((?:\d[\s]*){10,16})",

        # Код 201 может оказаться на соседней позиции:
        r"\b201\b.{0,80}?Поставщик\s+ИНН\s*:\s*((?:\d[\s]*){10,16})"
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            normalized,
            re.IGNORECASE | re.DOTALL
        )

        if match:
            candidate = digits_only(match.group(1))

            if is_valid_supplier_inn(candidate):
                return candidate

    return ""


def extract_supplier_inn_from_section(text: str) -> str:
    """
    Резервный способ.

    Анализирует участок между заголовком раздела реквизитов
    и полем покупателя. Это позволяет не перепутать ИНН
    поставщика с ИНН покупателя.
    """
    normalized = normalize_text(text)

    start_match = re.search(
        r"Раздел\s*1.*?Реквизиты\s+поставщика\s+и\s+покупателя",
        normalized,
        re.IGNORECASE | re.DOTALL
    )

    start_position = start_match.end() if start_match else 0

    supplier_marker = re.search(
        r"201\s+Поставщик\s+ИНН",
        normalized[start_position:],
        re.IGNORECASE
    )

    if not supplier_marker:
        return ""

    supplier_start = (
        start_position
        + supplier_marker.start()
    )

    buyer_marker = re.search(
        r"301\s+Покупатель\s+ИНН",
        normalized[supplier_start:],
        re.IGNORECASE
    )

    if buyer_marker:
        supplier_end = (
            supplier_start
            + buyer_marker.start()
        )
    else:
        supplier_end = supplier_start + 300

    fragment = normalized[
        supplier_start:supplier_end
    ]

    candidates = re.findall(
        r"(?<!\d)\d{10,16}(?!\d)",
        fragment
    )

    valid_candidates = [
        digits_only(candidate)
        for candidate in candidates
        if is_valid_supplier_inn(candidate)
    ]

    # Сначала предпочитаем 14-значный ИНН.
    for candidate in valid_candidates:
        if len(candidate) == 14:
            return candidate

    if valid_candidates:
        return valid_candidates[0]

    # Резервный поиск для цифр с пробелами.
    spaced_candidates = re.findall(
        r"(?<!\d)(?:\d[\s]*){10,16}(?!\d)",
        fragment
    )

    for candidate in spaced_candidates:
        candidate_digits = digits_only(candidate)

        if (
            is_valid_supplier_inn(candidate_digits)
            and len(candidate_digits) == 14
        ):
            return candidate_digits

    return ""


def extract_supplier_inn_by_coordinates(
    pdf_path: Path
) -> str:
    """
    Дополнительный резервный способ.

    Использует слова первой страницы и ищет значение,
    расположенное рядом с надписью «Поставщик ИНН».

    Это помогает, если обычное извлечение текста
    смешало колонки поставщика и покупателя.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return ""

            page = pdf.pages[0]

            words = page.extract_words(
                x_tolerance=2,
                y_tolerance=3,
                keep_blank_chars=False,
                use_text_flow=False
            )

            if not words:
                return ""

            # Ищем строку, где присутствуют слова
            # "Поставщик" и "ИНН".
            for index, word in enumerate(words):
                word_text = word.get("text", "").lower()

                if "поставщик" not in word_text:
                    continue

                current_top = float(word.get("top", 0))
                current_bottom = float(word.get("bottom", 0))

                nearby_words = []

                for candidate in words:
                    candidate_top = float(
                        candidate.get("top", 0)
                    )
                    candidate_bottom = float(
                        candidate.get("bottom", 0)
                    )

                    same_line = (
                        abs(candidate_top - current_top) <= 8
                        or abs(
                            candidate_bottom
                            - current_bottom
                        ) <= 8
                    )

                    # Берём значения справа от слова
                    # "Поставщик", но только в левой части
                    # формы, чтобы не захватить покупателя.
                    if (
                        same_line
                        and float(candidate.get("x0", 0))
                        >= float(word.get("x0", 0))
                        and float(candidate.get("x0", 0))
                        < page.width * 0.58
                    ):
                        nearby_words.append(candidate)

                nearby_words.sort(
                    key=lambda item: float(
                        item.get("x0", 0)
                    )
                )

                line_text = " ".join(
                    item.get("text", "")
                    for item in nearby_words
                )

                candidates = re.findall(
                    r"\d{10,16}",
                    digits_only_with_separators(line_text)
                )

                for candidate in candidates:
                    if (
                        is_valid_supplier_inn(candidate)
                        and len(candidate) == 14
                    ):
                        return candidate

    except Exception:
        return ""

    return ""


def digits_only_with_separators(value: str) -> str:
    """
    Удаляет буквы, но сохраняет группы цифр разделёнными пробелом.
    """
    value = re.sub(r"[^\d]+", " ", value or "")
    value = re.sub(r"\s+", " ", value)

    return value.strip()


def extract_supplier_inn(
    ordinary_text: str,
    layout_text: str,
    pdf_path: Path
) -> tuple[str, str]:
    """
    Последовательно применяет несколько способов определения ИНН.

    Возвращает:
    (ИНН, способ определения)
    """
    combined_text = (
        ordinary_text
        + "\n"
        + layout_text
    )

    inn = extract_supplier_inn_strict(combined_text)

    if inn:
        return inn, "ИНН найден по полю 201"

    inn = extract_supplier_inn_from_section(
        ordinary_text
    )

    if inn:
        return inn, "ИНН найден в разделе поставщика"

    inn = extract_supplier_inn_from_section(
        layout_text
    )

    if inn:
        return inn, "ИНН найден в layout-тексте"

    inn = extract_supplier_inn_by_coordinates(
        pdf_path
    )

    if inn:
        return inn, "ИНН найден по координатам"

    return "", "ИНН поставщика не найден"


# ============================================================
# НАИМЕНОВАНИЕ ПОСТАВЩИКА
# ============================================================

def simplify_organization_name(name: str) -> str:
    """
    Преобразует полную организационно-правовую форму
    в привычное сокращение.
    """
    name = clean_spaces(name)

    replacements = [
        (
            r"^Общество\s+с\s+ограниченной\s+ответственностью\s*",
            'ООО '
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

    return clean_spaces(name)


def extract_supplier_name(text: str) -> str:
    """
    Извлекает наименование поставщика из левой части формы.
    """
    normalized = normalize_text(text)

    patterns = [
        (
            r"201\s+Поставщик\s+ИНН\s*:\s*"
            r"(?:\d[\s]*){10,16}"
            r".{0,250}?"
            r"Ф\.?\s*И\.?\s*О\.?\s*ИП\s*/\s*"
            r"Наименование\s+организации\s*:\s*"
            r"(.+?)"
            r"(?=\s+202\b|\s+302\b|\s+203\b|"
            r"\s+Филиал\s+поставщика)"
        ),
        (
            r"Ф\.?\s*И\.?\s*О\.?\s*ИП\s*/\s*"
            r"Наименование\s+организации\s*:\s*"
            r"(.+?)"
            r"(?=\s+202\b|\s+302\b|\s+203\b|"
            r"\s+Филиал\s+поставщика)"
        )
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            normalized,
            re.IGNORECASE | re.DOTALL
        )

        if match:
            supplier = clean_spaces(match.group(1))

            # Если в результат попала метка правой колонки,
            # обрезаем её.
            supplier = re.split(
                r"\b302\b|Ф\.?\s*И\.?\s*О\.?\s*ИП\s*/",
                supplier,
                maxsplit=1,
                flags=re.IGNORECASE
            )[0]

            supplier = clean_spaces(supplier)

            if supplier:
                return simplify_organization_name(
                    supplier
                )

    return ""


# ============================================================
# КОД НАЛОГОВОГО ОРГАНА
# ============================================================

def extract_tax_office_code(text: str) -> str:
    """
    Извлекает код налогового органа поставщика.

    Пример:
    Код и наименование налогового органа:
    004 - УГНС по Первомайскому району
    """
    normalized = normalize_text(text)

    patterns = [
        (
            r"Код\s+и\s+наименование\s+налогового\s+органа\s*:"
            r"\s*(\d{3,4})\s*-"
        ),
        (
            r"206\s+Код\s+и\s+наименование\s+налогового\s+органа"
            r"\s*:\s*(\d{3,4})"
        ),
        (
            r"налогового\s+органа\s*:\s*(\d{3,4})\s*-"
        )
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            normalized,
            re.IGNORECASE | re.DOTALL
        )

        if match:
            return match.group(1).zfill(3)

    return ""


# ============================================================
# ДАТА ОФОРМЛЕНИЯ
# ============================================================

def extract_issue_date(text: str) -> str:
    """
    Извлекает дату оформления из поля 103.

    Обрабатывает как нормальную дату,
    так и отдельно распознанные цифры:
    0 3 0 8 2 0 2 6
    """
    normalized = normalize_text(text)

    marker = re.search(
        r"103\s+Дата\s+оформления",
        normalized,
        re.IGNORECASE
    )

    if marker:
        fragment = normalized[
            marker.end():marker.end() + 350
        ]

        normal_match = re.search(
            r"\b(\d{1,2})[./\-](\d{1,2})[./\-](\d{4})\b",
            fragment
        )

        if normal_match:
            return normalize_date(
                normal_match.group(1),
                normal_match.group(2),
                normal_match.group(3)
            )

        # Ограничиваем участок началом поля поставщика.
        fragment = re.split(
            r"201\s+Поставщик",
            fragment,
            maxsplit=1,
            flags=re.IGNORECASE
        )[0]

        digits = re.findall(r"\d", fragment)

        # Проверяем все последовательности из восьми цифр,
        # а не только первые восемь.
        for index in range(
            max(0, len(digits) - 7)
        ):
            date_digits = "".join(
                digits[index:index + 8]
            )

            date_value = normalize_date(
                date_digits[0:2],
                date_digits[2:4],
                date_digits[4:8]
            )

            if date_value:
                return date_value

    # Резерв: первая подходящая дата рядом
    # с верхней частью документа.
    top_fragment = normalized[:1000]

    date_candidates = re.findall(
        r"\b(\d{1,2})[./\-](\d{1,2})[./\-](\d{4})\b",
        top_fragment
    )

    for day, month, year in date_candidates:
        date_value = normalize_date(
            day,
            month,
            year
        )

        if date_value:
            return date_value

    return ""


# ============================================================
# СУММА НДС
# ============================================================

def extract_vat_from_total(text: str):
    """
    Из строки «Итого по счету-фактуре»
    извлекает сумму НДС.

    Структура итоговой строки:
    1 — стоимость без НДС;
    2 — сумма НДС;
    3 — сумма НсП;
    4 — общая стоимость.
    """
    normalized = normalize_text(text)

    marker_patterns = [
        r"Итого\s+по\s+счету\s*-\s*фактуре\s*:",
        r"Итого\s+по\s+счету-фактуре\s*:",
        r"Итого\s+по\s+счетуфактуре\s*:"
    ]

    marker = None

    for pattern in marker_patterns:
        marker = re.search(
            pattern,
            normalized,
            re.IGNORECASE
        )

        if marker:
            break

    if not marker:
        return None

    fragment = normalized[
        marker.end():marker.end() + 400
    ]

    numbers = re.findall(
        r"(?<!\d)"
        r"\d[\d \u00a0]*[.,]\d{2,5}"
        r"(?!\d)",
        fragment
    )

    parsed_numbers = []

    for number in numbers:
        decimal_value = parse_decimal(number)

        if decimal_value is not None:
            parsed_numbers.append(decimal_value)

    if len(parsed_numbers) >= 2:
        return parsed_numbers[1]

    return None


# ============================================================
# ОБРАБОТКА ОДНОГО PDF
# ============================================================

def process_pdf(pdf_path: Path) -> dict:
    """
    Обрабатывает один PDF.
    """
    ordinary_text, layout_text = extract_pdf_text(
        pdf_path
    )

    if (
        not ordinary_text.strip()
        and not layout_text.strip()
    ):
        raise ValueError(
            "PDF не содержит извлекаемого текста"
        )

    combined_text = (
        ordinary_text
        + "\n"
        + layout_text
    )

    supplier_inn, inn_status = extract_supplier_inn(
        ordinary_text,
        layout_text,
        pdf_path
    )

    invoice_number = extract_invoice_number(
        combined_text
    )

    supplier = extract_supplier_name(
        ordinary_text
    )

    if not supplier:
        supplier = extract_supplier_name(
            layout_text
        )

    tax_office_code = extract_tax_office_code(
        ordinary_text
    )

    if not tax_office_code:
        tax_office_code = extract_tax_office_code(
            layout_text
        )

    issue_date = extract_issue_date(
        ordinary_text
    )

    if not issue_date:
        issue_date = extract_issue_date(
            layout_text
        )

    vat = extract_vat_from_total(
        combined_text
    )

    missing_fields = []

    if not supplier:
        missing_fields.append(
            "наименование поставщика"
        )

    if not supplier_inn:
        missing_fields.append(
            "ИНН поставщика"
        )

    if not tax_office_code:
        missing_fields.append(
            "код налогового органа"
        )

    if not invoice_number:
        missing_fields.append(
            "номер счета-фактуры"
        )

    if not issue_date:
        missing_fields.append(
            "дата"
        )

    if vat is None:
        missing_fields.append(
            "сумма НДС"
        )

    if missing_fields:
        validation_status = (
            "Не найдено: "
            + ", ".join(missing_fields)
            + f". {inn_status}"
        )
    else:
        validation_status = (
            "Все поля найдены. "
            + inn_status
        )

    return {
        "source_file": pdf_path.name,
        "supplier": supplier,
        "supplier_inn": supplier_inn,
        "tax_office_code": tax_office_code,
        "invoice_number": invoice_number,
        "issue_date": issue_date,
        "vat": vat,
        "validation_status": validation_status,
        "is_complete": len(missing_fields) == 0
    }


# ============================================================
# EXCEL
# ============================================================

def create_thin_border() -> Border:
    """
    Создаёт тонкие границы ячеек.
    """
    side = Side(
        style="thin",
        color="000000"
    )

    return Border(
        left=side,
        right=side,
        top=side,
        bottom=side
    )


def save_excel(
    rows: list[dict],
    errors: list[dict],
    output_path: Path
):
    """
    Создаёт Result.xlsx по структуре VAAT.
    """
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "НДС"

    # --------------------------------------------------------
    # ЗАГОЛОВОК
    # --------------------------------------------------------

    sheet.merge_cells("A1:A2")
    sheet.merge_cells("B1:B2")
    sheet.merge_cells("C1:C2")
    sheet.merge_cells("D1:D2")
    sheet.merge_cells("E1:G1")

    sheet["A1"] = "№"
    sheet["B1"] = (
        "Наименование поставщика "
        "товаров (работ, услуг)"
    )
    sheet["C1"] = "ИНН"
    sheet["D1"] = "Код налогового органа"
    sheet["E1"] = "Счет-фактура"

    sheet["E2"] = "№"
    sheet["F2"] = "дата"
    sheet["G2"] = "сумма НДС (сом.)"

    # Третья строка — номера граф.
    for column_number in range(1, 8):
        sheet.cell(
            row=3,
            column=column_number,
            value=column_number
        )

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="D9EAF7"
    )

    warning_fill = PatternFill(
        fill_type="solid",
        fgColor="FFF2CC"
    )

    error_fill = PatternFill(
        fill_type="solid",
        fgColor="F4CCCC"
    )

    border = create_thin_border()

    for row_cells in sheet.iter_rows(
        min_row=1,
        max_row=3,
        min_col=1,
        max_col=7
    ):
        for cell in row_cells:
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.border = border
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

    # --------------------------------------------------------
    # ДАННЫЕ
    # --------------------------------------------------------

    for index, row in enumerate(rows, start=1):
        excel_row = sheet.max_row + 1

        sheet.append([
            index,
            row["supplier"],
            row["supplier_inn"],
            row["tax_office_code"],
            row["invoice_number"],
            row["issue_date"],
            (
                float(row["vat"])
                if row["vat"] is not None
                else None
            )
        ])

        for column_number in range(1, 8):
            cell = sheet.cell(
                row=excel_row,
                column=column_number
            )

            cell.border = border
            cell.alignment = Alignment(
                vertical="center",
                wrap_text=True
            )

        sheet.cell(
            row=excel_row,
            column=1
        ).alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        # Текстовый формат обязателен:
        # ведущие нули ИНН и кодов сохраняются.
        sheet.cell(
            row=excel_row,
            column=3
        ).number_format = "@"

        sheet.cell(
            row=excel_row,
            column=4
        ).number_format = "@"

        sheet.cell(
            row=excel_row,
            column=5
        ).number_format = "@"

        sheet.cell(
            row=excel_row,
            column=6
        ).number_format = "@"

        sheet.cell(
            row=excel_row,
            column=7
        ).number_format = '#,##0.00'

        # Если ИНН отсутствует — выделяем его красным.
        if not row["supplier_inn"]:
            sheet.cell(
                row=excel_row,
                column=3
            ).fill = error_fill

            sheet.cell(
                row=excel_row,
                column=3
            ).value = "НЕ НАЙДЕН"

        elif len(row["supplier_inn"]) != 14:
            # Допускаем нестандартный ИНН,
            # но выделяем для ручной проверки.
            sheet.cell(
                row=excel_row,
                column=3
            ).fill = warning_fill

        # Незавершённые строки выделяем жёлтым.
        if not row["is_complete"]:
            for column_number in range(1, 8):
                cell = sheet.cell(
                    row=excel_row,
                    column=column_number
                )

                if cell.fill.fill_type is None:
                    cell.fill = warning_fill

    # --------------------------------------------------------
    # ИТОГО
    # --------------------------------------------------------

    total_row = sheet.max_row + 1

    sheet.merge_cells(
        start_row=total_row,
        start_column=1,
        end_row=total_row,
        end_column=6
    )

    sheet.cell(
        row=total_row,
        column=1,
        value="Итого"
    )

    sheet.cell(
        row=total_row,
        column=7,
        value=f"=SUM(G4:G{total_row - 1})"
    )

    for column_number in range(1, 8):
        cell = sheet.cell(
            row=total_row,
            column=column_number
        )

        cell.font = Font(bold=True)
        cell.border = border
        cell.fill = header_fill
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    sheet.cell(
        row=total_row,
        column=7
    ).number_format = '#,##0.00'

    # --------------------------------------------------------
    # РАЗМЕРЫ И ПЕЧАТЬ
    # --------------------------------------------------------

    widths = {
        "A": 7,
        "B": 48,
        "C": 20,
        "D": 22,
        "E": 28,
        "F": 15,
        "G": 20
    }

    for column_letter, width in widths.items():
        sheet.column_dimensions[
            column_letter
        ].width = width

    sheet.row_dimensions[1].height = 38
    sheet.row_dimensions[2].height = 27
    sheet.row_dimensions[3].height = 22

    sheet.freeze_panes = "A4"
    sheet.auto_filter.ref = (
        f"A3:G{total_row - 1}"
    )

    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0

    sheet.print_title_rows = "1:3"
    sheet.sheet_view.showGridLines = False

    # --------------------------------------------------------
    # ЛИСТ ПРОВЕРКИ
    # --------------------------------------------------------

    check_sheet = workbook.create_sheet(
        "Проверка"
    )

    check_headers = [
        "Файл PDF",
        "Наименование поставщика",
        "ИНН поставщика",
        "Код налогового органа",
        "Номер счета-фактуры",
        "Дата",
        "Сумма НДС",
        "Результат проверки"
    ]

    check_sheet.append(check_headers)

    for cell in check_sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

    for row in rows:
        check_sheet.append([
            row["source_file"],
            row["supplier"],
            row["supplier_inn"],
            row["tax_office_code"],
            row["invoice_number"],
            row["issue_date"],
            (
                float(row["vat"])
                if row["vat"] is not None
                else None
            ),
            row["validation_status"]
        ])

        current_row = check_sheet.max_row

        for column_number in range(1, 9):
            cell = check_sheet.cell(
                row=current_row,
                column=column_number
            )

            cell.border = border
            cell.alignment = Alignment(
                vertical="center",
                wrap_text=True
            )

            if not row["is_complete"]:
                cell.fill = warning_fill

        check_sheet.cell(
            row=current_row,
            column=3
        ).number_format = "@"

        check_sheet.cell(
            row=current_row,
            column=4
        ).number_format = "@"

        check_sheet.cell(
            row=current_row,
            column=5
        ).number_format = "@"

        check_sheet.cell(
            row=current_row,
            column=7
        ).number_format = '#,##0.00'

    check_widths = {
        "A": 42,
        "B": 40,
        "C": 20,
        "D": 22,
        "E": 28,
        "F": 15,
        "G": 18,
        "H": 65
    }

    for column_letter, width in check_widths.items():
        check_sheet.column_dimensions[
            column_letter
        ].width = width

    check_sheet.freeze_panes = "A2"
    check_sheet.auto_filter.ref = (
        f"A1:H{check_sheet.max_row}"
    )

    # --------------------------------------------------------
    # ОШИБКИ ОТКРЫТИЯ PDF
    # --------------------------------------------------------

    if errors:
        error_sheet = workbook.create_sheet(
            "Ошибки"
        )

        error_sheet.append([
            "Файл PDF",
            "Ошибка"
        ])

        for cell in error_sheet[1]:
            cell.font = Font(bold=True)
            cell.fill = error_fill
            cell.border = border

        for error in errors:
            error_sheet.append([
                error["file"],
                error["error"]
            ])

            current_row = error_sheet.max_row

            for column_number in range(1, 3):
                error_sheet.cell(
                    row=current_row,
                    column=column_number
                ).border = border

        error_sheet.column_dimensions[
            "A"
        ].width = 45

        error_sheet.column_dimensions[
            "B"
        ].width = 90

        error_sheet.freeze_panes = "A2"

    workbook.save(output_path)


# ============================================================
# ИНТЕРФЕЙС
# ============================================================

def choose_folder(default_folder: Path):
    """
    Открывает окно выбора папки с PDF.
    """
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    selected = filedialog.askdirectory(
        title=(
            "Выберите папку "
            "с PDF-счетами-фактурами"
        ),
        initialdir=str(default_folder)
    )

    root.destroy()

    if not selected:
        return None

    return Path(selected)


def show_message(
    title: str,
    text: str,
    error: bool = False
):
    """
    Показывает сообщение пользователю.
    """
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    if error:
        messagebox.showerror(
            title,
            text
        )
    else:
        messagebox.showinfo(
            title,
            text
        )

    root.destroy()


# ============================================================
# ЗАПУСК
# ============================================================

def main():
    app_dir = application_directory()

    selected_folder = choose_folder(
        app_dir
    )

    if selected_folder is None:
        return

    # Обрабатываются PDF в выбранной папке.
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
            result = process_pdf(
                pdf_path
            )

            rows.append(result)

        except Exception as exc:
            errors.append({
                "file": pdf_path.name,
                "error": str(exc)
            })

    output_path = (
        selected_folder
        / "Result.xlsx"
    )

    try:
        save_excel(
            rows,
            errors,
            output_path
        )

    except PermissionError:
        show_message(
            "Не удалось сохранить Excel",
            "Закройте файл Result.xlsx в Excel "
            "и запустите программу повторно.",
            error=True
        )
        return

    complete_count = sum(
        1 for row in rows
        if row["is_complete"]
    )

    missing_inn_count = sum(
        1 for row in rows
        if not row["supplier_inn"]
    )

    vat_count = sum(
        1 for row in rows
        if row["vat"] is not None
    )

    message = (
        f"Обработано PDF: {len(pdf_files)}\n"
        f"Успешно открыто: {len(rows)}\n"
        f"Все поля найдены: {complete_count}\n"
        f"ИНН поставщика не найден: "
        f"{missing_inn_count}\n"
        f"Сумма НДС найдена: {vat_count}\n"
        f"Ошибок открытия PDF: {len(errors)}\n\n"
        f"Результат сохранён:\n"
        f"{output_path}\n\n"
        "Проверьте лист «Проверка»."
    )

    show_message(
        "Обработка завершена",
        message
    )


if __name__ == "__main__":
    try:
        main()

    except Exception:
        error_path = (
            application_directory()
            / "Extract_error.txt"
        )

        error_path.write_text(
            traceback.format_exc(),
            encoding="utf-8"
        )

        show_message(
            "Ошибка программы",
            "Произошла непредвиденная ошибка.\n\n"
            "Подробности сохранены в файле:\n"
            f"{error_path}",
            error=True
        )