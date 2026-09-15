import os
import re
import sys
import traceback
from pathlib import Path
from decimal import Decimal, InvalidOperation

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================
# НАСТРОЙКИ
# ============================================================

OUTPUT_FILE = "Result.xlsx"

# Возможные подписи полей в PDF
SUPPLIER_MARKERS = [
    "наименование организации",
    "наименование организации:",
    "ф.и.о. ип/наименование организации",
    "ф.и.о. ип / наименование организации",
    "фио ип/наименование организации",
]

SUPPLIER_INN_MARKERS = [
    "филиал поставщика инн",
    "поставщик инн",
    "инн поставщика",
]

BUYER_INN_MARKERS = [
    "покупатель инн",
    "инн покупателя",
]

DELIVERY_DATE_MARKERS = [
    "дата поставки",
]

TAX_CODE_MARKERS = [
    "код налогового органа",
]


# ============================================================
# ОБЩИЕ ФУНКЦИИ
# ============================================================

def clean_spaces(value):
    """Убирает лишние пробелы и переводы строк."""
    if value is None:
        return ""

    value = str(value)
    value = value.replace("\xa0", " ")
    value = value.replace("\r", " ")
    value = value.replace("\n", " ")

    return re.sub(r"\s+", " ", value).strip()


def digits_only(value):
    """Оставляет только цифры."""
    if value is None:
        return ""

    return re.sub(r"\D", "", str(value))


def normalize_text(value):
    """Текст для поиска."""
    value = clean_spaces(value).lower()

    # Ё -> Е
    value = value.replace("ё", "е")

    return value


def get_word_text(word):
    return clean_spaces(
        word.get("text", "")
    )


def get_word_x0(word):
    try:
        return float(word.get("x0", 0))
    except Exception:
        return 0.0


def get_word_x1(word):
    try:
        return float(word.get("x1", 0))
    except Exception:
        return 0.0


def get_word_top(word):
    try:
        return float(word.get("top", 0))
    except Exception:
        return 0.0


def get_word_bottom(word):
    try:
        return float(word.get("bottom", 0))
    except Exception:
        return 0.0


def group_words_into_lines(words, tolerance=4):
    """
    Группирует слова PDF по горизонтальным строкам.
    """
    if not words:
        return []

    words = sorted(
        words,
        key=lambda w: (
            get_word_top(w),
            get_word_x0(w)
        )
    )

    lines = []

    for word in words:
        top = get_word_top(word)

        target = None

        for line in lines:
            if abs(line["top"] - top) <= tolerance:
                target = line
                break

        if target is None:
            target = {
                "top": top,
                "words": []
            }
            lines.append(target)

        target["words"].append(word)

    result = []

    for line in sorted(lines, key=lambda x: x["top"]):
        result.append(
            sorted(
                line["words"],
                key=get_word_x0
            )
        )

    return result


def line_to_text(line):
    return clean_spaces(
        " ".join(
            get_word_text(word)
            for word in line
            if get_word_text(word)
        )
    )


def is_probably_field_code(text):
    """
    Проверяет, является ли строка номером поля:
    202, 203, 204, 301, 302 и т.п.
    """
    compact = digits_only(text)

    return compact in {
        "201", "202", "203", "204", "205", "206",
        "301", "302", "303", "304", "305", "306",
        "307", "308", "309"
    }


# ============================================================
# PAYMENT VOUCHER
# ============================================================

def extract_payment_voucher(filename):
    """
    Извлекает Payment Voucher# из имени PDF.

    Примеры:

        53112605-INV.pdf
            -> 53112605

        53112618-INV.pdf
            -> 53112618

        53112633-INV.pdf
            -> 53112633

    Номер берётся из имени файла, а не из PDF.
    """

    if not filename:
        return ""

    name = os.path.basename(filename)

    name = re.sub(
        r"\.pdf$",
        "",
        name,
        flags=re.IGNORECASE
    )

    # Основной формат:
    # 53112605-INV
    match = re.match(
        r"^(\d+)-INV$",
        name,
        flags=re.IGNORECASE
    )

    if match:
        return match.group(1)

    # Любой формат:
    # 53112605-...
    match = re.match(
        r"^(\d+)-",
        name
    )

    if match:
        return match.group(1)

    # Если имя только цифры
    match = re.match(
        r"^(\d+)$",
        name
    )

    if match:
        return match.group(1)

    # Последняя попытка:
    # найти первую последовательность >= 5 цифр
    match = re.search(
        r"\d{5,}",
        name
    )

    if match:
        return match.group(0)

    return ""


# ============================================================
# ИНН
# ============================================================

def normalize_inn_candidate(value):
    """
    Нормализует ИНН.

    Поддерживает:
    00406200910056
    0040 6200 9100 56
    0 0 4 0 6 2 0 0 9 1 0 0 5 6
    [0][0][4][0]...
    """
    if not value:
        return ""

    # Убираем квадратные/круглые поля
    value = re.sub(
        r"[\[\]\(\)\{\}<>]",
        " ",
        value
    )

    digits = re.sub(
        r"\D",
        "",
        value
    )

    # Кыргызский ИНН в рассматриваемых документах
    # обычно 14 цифр.
    if len(digits) == 14:
        return digits

    return ""


def extract_14_digit_sequences(text):
    """
    Находит обычные 14-значные ИНН.
    """
    if not text:
        return []

    result = []

    # Вариант:
    # 00406200910056
    for match in re.finditer(
        r"(?<!\d)\d{14}(?!\d)",
        text
    ):
        value = match.group(0)

        if value not in result:
            result.append(value)

    return result


def extract_spaced_inn(text):
    """
    Распознаёт ИНН, записанный отдельными цифрами.

    Примеры:

    [0] [0] [4] [0] [6] [2] ...

    0 0 4 0 6 2 0 0 9 1 0 0 5 6

    0|0|4|0|6|2|0|0|9|1|0|0|5|6
    """

    if not text:
        return []

    # Сначала заменяем поля на пробелы.
    normalized = re.sub(
        r"[\[\]\(\)\{\}<>|:_;]",
        " ",
        text
    )

    # Все одиночные цифры
    tokens = re.findall(
        r"(?<!\d)\d(?!\d)",
        normalized
    )

    result = []

    # Ищем последовательности из 14 отдельных цифр.
    for i in range(
        0,
        len(tokens) - 13
    ):
        candidate = "".join(
            tokens[i:i + 14]
        )

        if len(candidate) == 14:
            if candidate not in result:
                result.append(candidate)

    return result


def extract_inn_from_text(text):
    """
    Общий поиск ИНН.
    """

    candidates = extract_14_digit_sequences(text)

    if candidates:
        return candidates[0]

    candidates = extract_spaced_inn(text)

    if candidates:
        return candidates[0]

    return ""


# ============================================================
# НАЗВАНИЕ ПОСТАВЩИКА
# ============================================================

def simplify_organization_name(value):
    """
    Приводит длинные организационные формы
    к удобному виду.

    Например:

    Общество с ограниченной ответственностью Альфа
        ->
    ООО Альфа

    Закрытое акционерное общество Альфа
        ->
    ЗАО Альфа
    """

    value = clean_spaces(value)

    replacements = [
        (
            r"^Общество\s+с\s+ограниченной\s+ответственностью\s*",
            "ООО "
        ),
        (
            r"^Закрытое\s+акционерное\s+общество\s*",
            "ЗАО "
        ),
        (
            r"^Открытое\s+акционерное\s+общество\s*",
            "ОАО "
        ),
        (
            r"^Муниципальное\s+предприятие\s*",
            "МП "
        ),
        (
            r"^Государственное\s+предприятие\s*",
            "ГП "
        ),
        (
            r"^Индивидуальный\s+предприниматель\s*",
            "ИП "
        ),
    ]

    for pattern, replacement in replacements:
        value = re.sub(
            pattern,
            replacement,
            value,
            flags=re.IGNORECASE
        )

    return clean_spaces(value)


def find_supplier_column_right_edge(words, page_width):
    """
    Определяет границу левой колонки поставщика.

    В документах справа обычно находятся поля 301/302/303...
    """

    possible_edges = []

    for word in words:
        text = get_word_text(word)

        compact = digits_only(text)

        if compact not in {
            "301",
            "302",
            "303",
            "304",
            "305",
            "306",
        }:
            continue

        x0 = get_word_x0(word)

        if x0 > page_width * 0.40:
            possible_edges.append(x0)

    if possible_edges:
        return min(possible_edges) - 2

    return page_width * 0.55


def extract_supplier_name_from_coordinates(page_info):
    """
    Извлекает название поставщика из координат PDF.

    Важное отличие:
    НЕ начинает читать после номера поля 202.

    Он читает всю левую колонку, поэтому
    вторую строку названия тоже получает.
    """

    if not page_info:
        return ""

    words = page_info.get("words", [])

    if not words:
        return ""

    page_width = page_info.get(
        "width",
        595
    )

    right_edge = find_supplier_column_right_edge(
        words,
        page_width
    )

    supplier_words = []

    for word in words:
        x0 = get_word_x0(word)
        x1 = get_word_x1(word)

        center = (x0 + x1) / 2

        if center < right_edge:
            supplier_words.append(word)

    lines = group_words_into_lines(
        supplier_words,
        tolerance=4
    )

    collecting = False
    parts = []

    for line in lines:

        text = line_to_text(line)

        if not text:
            continue

        lower = normalize_text(text)

        if not collecting:

            if (
                "наименование организации" in lower
                or
                (
                    "наименование" in lower
                    and "организац" in lower
                )
            ):

                collecting = True

                # Убираем номер поля 202
                text = re.sub(
                    r"^\s*2\s*0\s*2\s*",
                    "",
                    text
                )

                # Находим двоеточие после названия поля
                match = re.search(
                    r"наименование\s+организац(?:ии|ия)"
                    r"\s*:?\s*(.*)$",
                    normalize_text(text),
                    flags=re.IGNORECASE
                )

                if match:

                    original_match = re.search(
                        r"Наименование\s+организац(?:ии|ия)"
                        r"\s*:?\s*(.*)$",
                        text,
                        flags=re.IGNORECASE
                    )

                    if original_match:
                        first_part = clean_spaces(
                            original_match.group(1)
                        )
                    else:
                        first_part = clean_spaces(
                            match.group(1)
                        )

                    if first_part:
                        # Не допускаем код 302
                        first_part = re.sub(
                            r"\s+302\s*$",
                            "",
                            first_part
                        )

                        if first_part:
                            parts.append(
                                first_part
                            )

                continue

        # ====================================================
        # ОСТАНОВКА
        # ====================================================

        stop_markers = [
            "филиал поставщика",
            "наименование филиала",
            "адрес (юридич",
            "адрес юридического",
            "код и наименование налогового органа",
        ]

        if any(
            marker in lower
            for marker in stop_markers
        ):
            break

        compact = digits_only(text)

        if compact in {
            "203",
            "204",
            "205",
            "206",
        }:
            break

        # Убираем служебный номер поля
        text = re.sub(
            r"^\s*2\s*0\s*[3-9]\s*",
            "",
            text
        )

        # Убираем 301/302...
        text = re.sub(
            r"\s+30[1-9]\s*$",
            "",
            text
        )

        text = clean_spaces(text)

        if text:
            parts.append(text)

    result = clean_spaces(
        " ".join(parts)
    )

    result = re.sub(
        r"^\s*202\s*",
        "",
        result
    )

    result = re.sub(
        r"\s+30[1-9]\s*$",
        "",
        result
    )

    result = clean_spaces(result)

    if not result:
        return ""

    return simplify_organization_name(result)


def extract_supplier_name_from_text(text):
    """
    Резервный способ извлечения названия
    из обычного текста PDF.
    """

    if not text:
        return ""

    lines = text.splitlines()

    collecting = False
    parts = []

    for original_line in lines:

        line = clean_spaces(original_line)

        if not line:
            continue

        lower = normalize_text(line)

        if not collecting:

            if (
                "наименование организации" in lower
                or
                (
                    "наименование" in lower
                    and "организац" in lower
                )
            ):

                collecting = True

                # Не захватываем покупателя
                line = re.split(
                    r"\b302\b",
                    line,
                    maxsplit=1
                )[0]

                match = re.search(
                    r"Наименование\s+организац(?:ии|ия)"
                    r"\s*:?\s*(.*)$",
                    line,
                    flags=re.IGNORECASE
                )

                if match:
                    part = clean_spaces(
                        match.group(1)
                    )

                    if part:
                        parts.append(part)

                continue

        # Остановка
        if (
            "филиал поставщика" in lower
            or
            "наименование филиала" in lower
            or
            re.search(
                r"^\s*203\b",
                line
            )
        ):
            break

        # Отрезаем покупателя
        line = re.split(
            r"\b302\b",
            line,
            maxsplit=1
        )[0]

        # Отрезаем код поля справа
        line = re.sub(
            r"\s+30[1-9]\s*$",
            "",
            line
        )

        line = clean_spaces(line)

        if line:
            parts.append(line)

    result = clean_spaces(
        " ".join(parts)
    )

    return simplify_organization_name(
        result
    )


def extract_supplier_name(
    ordinary_text,
    layout_text,
    page_info
):
    """
    Сначала координаты, потом текстовые методы.
    """

    result = extract_supplier_name_from_coordinates(
        page_info
    )

    if result:
        return result

    result = extract_supplier_name_from_text(
        ordinary_text
    )

    if result:
        return result

    return extract_supplier_name_from_text(
        layout_text
    )


# ============================================================
# ПОИСК ИНН ПОСТАВЩИКА
# ============================================================

def extract_supplier_inn_from_coordinates(page_info):
    """
    Ищет ИНН поставщика возле поля
    «Филиал поставщика ИНН».
    """

    if not page_info:
        return ""

    words = page_info.get("words", [])

    if not words:
        return ""

    lines = group_words_into_lines(
        words,
        tolerance=4
    )

    collecting = False

    for line in lines:

        text = line_to_text(line)

        if not text:
            continue

        lower = normalize_text(text)

        if (
            "филиал поставщика" in lower
            and "инн" in lower
        ) or (
            "инн поставщика" in lower
        ):

            collecting = True

            # ИНН может быть в этой же строке
            inn = extract_inn_from_text(
                text
            )

            if inn:
                return inn

            continue

        if collecting:

            # Берём несколько следующих строк,
            # чтобы поймать отдельные квадраты.
            inn = extract_inn_from_text(
                text
            )

            if inn:
                return inn

    return ""


def extract_supplier_inn(
    ordinary_text,
    layout_text,
    page_info
):
    """
    Несколько способов поиска ИНН.
    """

    result = extract_supplier_inn_from_coordinates(
        page_info
    )

    if result:
        return result

    texts = [
        ordinary_text,
        layout_text
    ]

    for text in texts:

        lines = text.splitlines()

        for i, line in enumerate(lines):

            lower = normalize_text(line)

            if (
                "инн поставщика" in lower
                or
                (
                    "филиал поставщика" in lower
                    and "инн" in lower
                )
            ):

                # Проверяем эту и следующие 3 строки
                block = "\n".join(
                    lines[i:i + 4]
                )

                inn = extract_inn_from_text(
                    block
                )

                if inn:
                    return inn

    return ""


# ============================================================
# ПОКУПАТЕЛЬ ИНН
# ============================================================

def extract_buyer_inn_from_coordinates(page_info):
    """
    Ищет ИНН покупателя в правой колонке.
    """

    if not page_info:
        return ""

    words = page_info.get("words", [])

    if not words:
        return ""

    lines = group_words_into_lines(
        words,
        tolerance=4
    )

    for index, line in enumerate(lines):

        text = line_to_text(line)

        lower = normalize_text(text)

        if (
            "покупатель" in lower
            and "инн" in lower
        ) or (
            lower.strip() == "инн"
        ):

            # Проверяем текущую строку
            inn = extract_inn_from_text(
                text
            )

            if inn:
                return inn

            # И следующие строки
            block_parts = [
                text
            ]

            for next_line in lines[
                index + 1:index + 4
            ]:
                block_parts.append(
                    line_to_text(next_line)
                )

            inn = extract_inn_from_text(
                "\n".join(block_parts)
            )

            if inn:
                return inn

    return ""


def extract_buyer_inn(
    ordinary_text,
    layout_text,
    page_info
):
    result = extract_buyer_inn_from_coordinates(
        page_info
    )

    if result:
        return result

    for text in [
        ordinary_text,
        layout_text
    ]:

        lines = text.splitlines()

        for i, line in enumerate(lines):

            lower = normalize_text(line)

            if (
                "покупатель" in lower
                and "инн" in lower
            ):

                block = "\n".join(
                    lines[i:i + 4]
                )

                inn = extract_inn_from_text(
                    block
                )

                if inn:
                    return inn

    return ""


# ============================================================
# КОД НАЛОГОВОГО ОРГАНА
# ============================================================

def extract_tax_code(
    ordinary_text,
    layout_text,
    page_info
):
    """
    Ищет код налогового органа.
    """

    texts = [
        ordinary_text,
        layout_text
    ]

    for text in texts:

        lines = text.splitlines()

        for i, line in enumerate(lines):

            lower = normalize_text(line)

            if (
                "код налогового органа" in lower
                or
                (
                    "налогового органа" in lower
                    and "код" in lower
                )
            ):

                # Сначала числа в этой строке
                nums = re.findall(
                    r"\b\d{2,6}\b",
                    line
                )

                # Исключаем типичные номера полей
                nums = [
                    n for n in nums
                    if n not in {
                        "202",
                        "203",
                        "204",
                        "301",
                        "302",
                        "303",
                    }
                ]

                if nums:
                    return nums[-1]

                # Следующая строка
                if i + 1 < len(lines):

                    nums = re.findall(
                        r"\b\d{2,6}\b",
                        lines[i + 1]
                    )

                    if nums:
                        return nums[0]

    # Попытка по координатам
    if page_info:

        words = page_info.get(
            "words",
            []
        )

        lines = group_words_into_lines(
            words,
            tolerance=4
        )

        for i, line in enumerate(lines):

            text = line_to_text(line)

            lower = normalize_text(text)

            if (
                "код налогового органа" in lower
            ):

                nums = re.findall(
                    r"\b\d{2,6}\b",
                    text
                )

                nums = [
                    n for n in nums
                    if n not in {
                        "202",
                        "203",
                        "301",
                        "302",
                    }
                ]

                if nums:
                    return nums[-1]

    return ""


# ============================================================
# ДАТА ПОСТАВКИ
# ============================================================

def normalize_date(value):
    """
    Приводит даты к DD.MM.YYYY.
    """

    if not value:
        return ""

    value = clean_spaces(value)

    # DD.MM.YYYY
    match = re.search(
        r"\b(\d{1,2})[./-](\d{1,2})[./-](\d{4})\b",
        value
    )

    if match:

        day = int(match.group(1))
        month = int(match.group(2))
        year = int(match.group(3))

        return (
            f"{day:02d}."
            f"{month:02d}."
            f"{year:04d}"
        )

    # YYYY-MM-DD
    match = re.search(
        r"\b(\d{4})-(\d{1,2})-(\d{1,2})\b",
        value
    )

    if match:

        year = int(match.group(1))
        month = int(match.group(2))
        day = int(match.group(3))

        return (
            f"{day:02d}."
            f"{month:02d}."
            f"{year:04d}"
        )

    return ""


def extract_delivery_date(
    ordinary_text,
    layout_text
):
    """
    Ищет дату поставки.
    """

    for text in [
        ordinary_text,
        layout_text
    ]:

        lines = text.splitlines()

        for i, line in enumerate(lines):

            lower = normalize_text(line)

            if "дата поставки" in lower:

                date = normalize_date(
                    line
                )

                if date:
                    return date

                # Следующие строки
                for next_line in lines[
                    i + 1:i + 3
                ]:

                    date = normalize_date(
                        next_line
                    )

                    if date:
                        return date

    return ""


# ============================================================
# НОМЕР СЧЁТ-ФАКТУРЫ
# ============================================================

def extract_invoice_number(
    ordinary_text,
    layout_text
):
    """
    Ищет номер счёт-фактуры.

    Ориентируется на длинные номера,
    содержащие структуру вроде:

    0002026-004-01482118
    """

    texts = [
        ordinary_text,
        layout_text
    ]

    for text in texts:

        # Основной формат
        matches = re.findall(
            r"\b\d{5,}-\d{3}-\d{6,}\b",
            text
        )

        if matches:
            return matches[0]

        # Более общий формат
        matches = re.findall(
            r"\b\d{5,}-\d{2,5}-\d{5,}\b",
            text
        )

        if matches:
            return matches[0]

    return ""


# ============================================================
# ДАТА СЧЁТ-ФАКТУРЫ
# ============================================================

def extract_invoice_date(
    ordinary_text,
    layout_text
):
    """
    Ищет дату счёт-фактуры.
    """

    for text in [
        ordinary_text,
        layout_text
    ]:

        lines = text.splitlines()

        for i, line in enumerate(lines):

            lower = normalize_text(line)

            if (
                "дата" in lower
                and (
                    "счет-фактур" in lower
                    or
                    "счёт-фактур" in lower
                )
            ):

                date = normalize_date(
                    line
                )

                if date:
                    return date

                for next_line in lines[
                    i + 1:i + 3
                ]:

                    date = normalize_date(
                        next_line
                    )

                    if date:
                        return date

    # Если подпись не распознана,
    # пытаемся взять дату рядом с номером СФ.
    invoice_number = extract_invoice_number(
        ordinary_text,
        layout_text
    )

    if invoice_number:

        for text in [
            ordinary_text,
            layout_text
        ]:

            pos = text.find(
                invoice_number
            )

            if pos >= 0:

                fragment = text[
                    max(0, pos - 200):
                    pos + 300
                ]

                date = normalize_date(
                    fragment
                )

                if date:
                    return date

    return ""


# ============================================================
# СУММА НДС
# ============================================================

def parse_money(value):
    """
    Преобразует сумму в Decimal.
    """

    if value is None:
        return None

    value = str(value)

    value = value.replace(
        "\xa0",
        ""
    )

    value = value.replace(
        " ",
        ""
    )

    value = value.replace(
        ",",
        "."
    )

    value = re.sub(
        r"[^\d.-]",
        "",
        value
    )

    if not value:
        return None

    try:
        return Decimal(value)
    except InvalidOperation:
        return None


def format_money(value):
    """
    Excel получит числовое значение,
    а не текст.
    """

    if value is None:
        return 0

    try:
        return float(value)
    except Exception:
        return 0


def extract_vat_amount(
    ordinary_text,
    layout_text
):
    """
    Ищет именно сумму НДС.

    НЕ берёт общую сумму.
    """

    texts = [
        ordinary_text,
        layout_text
    ]

    patterns = [
        # Сумма НДС
        r"сумма\s+ндс\s*[:\-]?\s*([0-9][0-9\s.,]*)",

        # НДС
        r"\bндс\b\s*[:\-]?\s*([0-9][0-9\s.,]*)",

        # сумма налога на добавленную стоимость
        r"налог\s+на\s+добавленную\s+стоимость"
        r"\s*[:\-]?\s*([0-9][0-9\s.,]*)",
    ]

    for text in texts:

        for pattern in patterns:

            matches = re.findall(
                pattern,
                text,
                flags=re.IGNORECASE
            )

            for match in matches:

                amount = parse_money(
                    match
                )

                if amount is not None:
                    return amount

    # ========================================================
    # Попытка найти НДС в табличной части.
    # ========================================================

    for text in texts:

        lines = text.splitlines()

        for i, line in enumerate(lines):

            lower = normalize_text(line)

            if "ндс" not in lower:
                continue

            # Пробуем число в той же строке.
            numbers = re.findall(
                r"\d[\d\s]*[.,]\d{1,2}",
                line
            )

            if numbers:

                # Обычно последнее число в строке
                amount = parse_money(
                    numbers[-1]
                )

                if amount is not None:
                    return amount

            # Следующие строки
            for next_line in lines[
                i + 1:i + 3
            ]:

                numbers = re.findall(
                    r"\d[\d\s]*[.,]\d{1,2}",
                    next_line
                )

                if numbers:

                    amount = parse_money(
                        numbers[-1]
                    )

                    if amount is not None:
                        return amount

    return Decimal("0.00")


# ============================================================
# ОБРАБОТКА ОДНОГО PDF
# ============================================================

def process_pdf(pdf_path):
    """
    Полностью обрабатывает один PDF.
    """

    filename = pdf_path.name

    payment_voucher = extract_payment_voucher(
        filename
    )

    ordinary_pages = []
    layout_pages = []
    first_page_info = None

    with pdfplumber.open(
        str(pdf_path)
    ) as pdf:

        for page_number, page in enumerate(
            pdf.pages
        ):

            ordinary = page.extract_text(
                x_tolerance=2,
                y_tolerance=3
            ) or ""

            layout = page.extract_text(
                layout=True,
                x_tolerance=2,
                y_tolerance=3
            ) or ""

            words = page.extract_words(
                x_tolerance=2,
                y_tolerance=3,
                keep_blank_chars=False
            )

            if first_page_info is None:
                first_page_info = {
                    "words": words,
                    "width": page.width,
                    "height": page.height
                }

            ordinary_pages.append(
                ordinary
            )

            layout_pages.append(
                layout
            )

    ordinary_text = "\n".join(
        ordinary_pages
    )

    layout_text = "\n".join(
        layout_pages
    )

    supplier_name = extract_supplier_name(
        ordinary_text,
        layout_text,
        first_page_info
    )

    supplier_inn = extract_supplier_inn(
        ordinary_text,
        layout_text,
        first_page_info
    )

    tax_code = extract_tax_code(
        ordinary_text,
        layout_text,
        first_page_info
    )

    buyer_inn = extract_buyer_inn(
        ordinary_text,
        layout_text,
        first_page_info
    )

    delivery_date = extract_delivery_date(
        ordinary_text,
        layout_text
    )

    invoice_number = extract_invoice_number(
        ordinary_text,
        layout_text
    )

    invoice_date = extract_invoice_date(
        ordinary_text,
        layout_text
    )

    vat_amount = extract_vat_amount(
        ordinary_text,
        layout_text
    )

    # Проверка обязательного ИНН поставщика
    if not supplier_inn:
        result = "ОШИБКА: не найден ИНН поставщика"
    else:
        result = "OK"

    return {
        "Файл": filename,
        "Payment Voucher#": payment_voucher,
        "Номер СФ": invoice_number,
        "Дата": invoice_date,
        "Поставщик": supplier_name,
        "ИНН поставщика": supplier_inn,
        "Код налогового органа": tax_code,
        "Покупатель ИНН": buyer_inn,
        "Дата поставки": delivery_date,
        "Сумма НДС": format_money(
            vat_amount
        ),
        "Результат": result,
    }


# ============================================================
# EXCEL
# ============================================================

def create_excel(results, output_path):
    """
    Создаёт Result.xlsx.
    """

    wb = Workbook()

    ws = wb.active
    ws.title = "Result"

    headers = [
        "Файл",
        "Payment Voucher#",
        "Номер СФ",
        "Дата",
        "Поставщик",
        "ИНН поставщика",
        "Код налогового органа",
        "Покупатель ИНН",
        "Дата поставки",
        "Сумма НДС",
        "Результат",
    ]

    # Заголовки
    for col, header in enumerate(
        headers,
        start=1
    ):

        cell = ws.cell(
            row=1,
            column=col,
            value=header
        )

        cell.font = Font(
            bold=True
        )

        cell.fill = PatternFill(
            "solid",
            fgColor="D9EAF7"
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    # Данные
    for row_number, item in enumerate(
        results,
        start=2
    ):

        values = [
            item["Файл"],
            item["Payment Voucher#"],
            item["Номер СФ"],
            item["Дата"],
            item["Поставщик"],
            item["ИНН поставщика"],
            item["Код налогового органа"],
            item["Покупатель ИНН"],
            item["Дата поставки"],
            item["Сумма НДС"],
            item["Результат"],
        ]

        for col, value in enumerate(
            values,
            start=1
        ):

            cell = ws.cell(
                row=row_number,
                column=col,
                value=value
            )

            cell.alignment = Alignment(
                vertical="center"
            )

    # Формат суммы НДС
    vat_column = headers.index(
        "Сумма НДС"
    ) + 1

    for row in range(
        2,
        ws.max_row + 1
    ):

        ws.cell(
            row=row,
            column=vat_column
        ).number_format = '#,##0.00'

    # Границы
    thin = Side(
        style="thin"
    )

    for row in ws.iter_rows():
        for cell in row:
            cell.border = Border(
                left=thin,
                right=thin,
                top=thin,
                bottom=thin
            )

    # Ширина колонок
    widths = {
        "A": 34,
        "B": 20,
        "C": 30,
        "D": 14,
        "E": 45,
        "F": 20,
        "G": 24,
        "H": 20,
        "I": 18,
        "J": 18,
        "K": 35,
    }

    for column, width in widths.items():
        ws.column_dimensions[
            column
        ].width = width

    ws.freeze_panes = "A2"

    # Автофильтр
    ws.auto_filter.ref = ws.dimensions

    # ========================================================
    # Отдельный лист "Ошибки"
    # ========================================================

    errors = [
        item
        for item in results
        if item["Результат"] != "OK"
    ]

    if errors:

        ws_errors = wb.create_sheet(
            "Ошибки"
        )

        error_headers = [
            "Файл",
            "Payment Voucher#",
            "Ошибка"
        ]

        for col, header in enumerate(
            error_headers,
            start=1
        ):

            cell = ws_errors.cell(
                row=1,
                column=col,
                value=header
            )

            cell.font = Font(
                bold=True
            )

        for row_number, item in enumerate(
            errors,
            start=2
        ):

            ws_errors.cell(
                row=row_number,
                column=1,
                value=item["Файл"]
            )

            ws_errors.cell(
                row=row_number,
                column=2,
                value=item[
                    "Payment Voucher#"
                ]
            )

            ws_errors.cell(
                row=row_number,
                column=3,
                value=item["Результат"]
            )

        ws_errors.column_dimensions[
            "A"
        ].width = 40

        ws_errors.column_dimensions[
            "B"
        ].width = 25

        ws_errors.column_dimensions[
            "C"
        ].width = 50

    wb.save(
        output_path
    )


# ============================================================
# ПРОГРЕСС
# ============================================================

def print_progress(
    current,
    total,
    filename
):
    """
    Показывает прогресс в консоли.
    """

    if total <= 0:
        return

    percent = (
        current / total
    ) * 100

    bar_length = 30

    filled = int(
        bar_length * current / total
    )

    bar = (
        "#" * filled
        +
        "-" * (
            bar_length - filled
        )
    )

    print(
        f"\r[{bar}] "
        f"{percent:6.2f}% "
        f"{current}/{total} "
        f"{filename[:45]:45}",
        end="",
        flush=True
    )


# ============================================================
# MAIN
# ============================================================

def main():
    """
    Основная функция.
    """

    # ========================================================
    # Определяем папку запуска
    # ========================================================

    if getattr(
        sys,
        "frozen",
        False
    ):

        base_dir = Path(
            sys.executable
        ).resolve().parent

    else:

        base_dir = Path(
            __file__
        ).resolve().parent

    # ========================================================
    # Ищем PDF
    # ========================================================

    pdf_files = sorted(
        base_dir.glob(
            "*.pdf"
        )
    )

    if not pdf_files:

        print()
        print(
            "ОШИБКА: PDF-файлы не найдены."
        )
        print()
        print(
            "Положите Extract.exe "
            "в папку с PDF-файлами."
        )

        input(
            "\nНажмите Enter для выхода..."
        )

        return

    total = len(
        pdf_files
    )

    print()
    print(
        "=========================================="
    )
    print(
        "        EXTRACT VAT"
    )
    print(
        "=========================================="
    )
    print()
    print(
        f"Найдено PDF: {total}"
    )
    print()

    results = []

    for index, pdf_path in enumerate(
        pdf_files,
        start=1
    ):

        try:

            result = process_pdf(
                pdf_path
            )

            results.append(
                result
            )

        except Exception as error:

            results.append({
                "Файл": pdf_path.name,
                "Payment Voucher#":
                    extract_payment_voucher(
                        pdf_path.name
                    ),
                "Номер СФ": "",
                "Дата": "",
                "Поставщик": "",
                "ИНН поставщика": "",
                "Код налогового органа": "",
                "Покупатель ИНН": "",
                "Дата поставки": "",
                "Сумма НДС": 0,
                "Результат":
                    "ОШИБКА: " + str(error),
            })

        print_progress(
            index,
            total,
            pdf_path.name
        )

    print()
    print()

    # ========================================================
    # Создаём Result.xlsx
    # ========================================================

    output_path = base_dir / OUTPUT_FILE

    try:

        create_excel(
            results,
            output_path
        )

    except Exception as error:

        print()
        print(
            "ОШИБКА при создании Excel:"
        )
        print(
            str(error)
        )

        traceback.print_exc()

        input(
            "\nНажмите Enter для выхода..."
        )

        return

    # ========================================================
    # Статистика
    # ========================================================

    successful = sum(
        1
        for item in results
        if item["Результат"] == "OK"
    )

    errors = total - successful

    print(
        "=========================================="
    )

    print(
        f"Обработано: {total}"
    )

    print(
        f"Успешно:    {successful}"
    )

    print(
        f"Ошибок:     {errors}"
    )

    print()

    print(
        f"Result.xlsx создан:"
    )

    print(
        str(output_path)
    )

    print(
        "=========================================="
    )

    input(
        "\nНажмите Enter для выхода..."
    )


if __name__ == "__main__":
    main()
