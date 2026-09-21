#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Parse RFC RES tariff indexation PDFs (2021–2026) → Excel workbooks."""

from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
SOURCES = ROOT / "sources"
OUT_DIR = ROOT / "excel"

YEARS = list(range(2021, 2027))

SOURCE_URLS = {
    2026: "https://rfc.kz/upload/iblock/ec4/cipxy5puv9glcnpxrwph6f2n7atteain.pdf",
    2025: "https://rfc.kz/upload/iblock/880/wqf5gtulgsmjb87drenj95ecmr0u2pw7.pdf",
    2024: "https://rfc.kz/upload/iblock/06d/tl7t8z3i5w2qcy3brmjctayltaxpqhfc.pdf",
    2023: "https://rfc.kz/upload/iblock/e0f/wupeui7d9m27rk6ag1yblnh54lk6g6u8.pdf",
    2022: "https://rfc.kz/upload/iblock/e47/dt6xg7hp4rj50zwennfbzjwnkmoe0mor.pdf",
    2021: "https://rfc.kz/upload/iblock/41e/4ha629bcmdvynfrg03b0c2veowpjcnhe.pdf",
}

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
THIN = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
TITLE_FONT = Font(bold=True, size=13, color="1F4E79", name="Calibri")
NUM_FMT = "0.00"


def clean(s: str | None) -> str:
    if not s:
        return ""
    s = s.replace("\n", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def parse_num(s: str | None) -> float | None:
    if s is None:
        return None
    s = clean(s).replace(" ", "").replace("\xa0", "")
    if not s or s in {"-", "—", "–"}:
        return None
    s = s.replace(",", ".")
    # keep trailing zeros like 20.1200
    if not re.fullmatch(r"-?\d+(\.\d+)?", s):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def norm_formula(s: str) -> str:
    s = clean(s)
    s = s.replace("Тt", "Tt").replace("Тt+1", "Tt+1")
    return s


def classify_header(row: list) -> str | None:
    """Return 'fixed', 'auction', or None if not a header row."""
    blob = " ".join(clean(c) for c in row).lower()
    if not blob:
        return None
    if "№п/п" in blob or "аукционн" in blob:
        return "auction"
    if "утвержден" in blob or ("тип" in blob and "виэ" in blob and "формула" in blob):
        return "fixed"
    if "тип" in blob and "ед" in blob and "формула" in blob:
        return "fixed"
    return None


def looks_like_data_row(row: list) -> bool:
    vals = [clean(c) for c in row if clean(c)]
    if not vals:
        return False
    # auction: starts with number
    if re.fullmatch(r"\d+", vals[0]):
        return True
    # fixed: type name + unit or number somewhere
    joined = " ".join(vals).lower()
    if any(
        k in joined
        for k in (
            "гидро",
            "ветр",
            "солнеч",
            "фотоэлектр",
            "био",
            "kaz pv",
            "каскад",
            "малые",
        )
    ):
        return True
    return False


def extract_year_ipc_from_header(cell: str) -> tuple[int | None, float | None]:
    """Only treat cells that are indexation-year columns (1 октября / ИПЦ)."""
    cell = clean(cell)
    # Ignore legal references like «ППРК от 12 июня 2014 года №645»
    if re.search(r"ППРК|постановлен|№\s*\d{3}", cell, re.I) and "октября" not in cell.lower():
        return None, None
    if "октября" not in cell.lower() and "ИПЦ" not in cell:
        return None, None
    ym = re.search(r"(\d{4})\s*г", cell)
    year = int(ym.group(1)) if ym else None
    if year is not None and not (2014 <= year <= 2035):
        year = None
    im = re.search(r"ИПЦ\s*[-–—]?\s*(\d{2,3}[,.]\d)\s*%", cell, re.I)
    ipc = float(im.group(1).replace(",", ".")) if im else None
    return year, ipc


def parse_year_columns(header: list) -> list[tuple[int, int, float | None]]:
    """Return list of (col_index, year, ipc_percent_or_None)."""
    out = []
    for i, cell in enumerate(header):
        year, ipc = extract_year_ipc_from_header(cell or "")
        if year:
            out.append((i, year, ipc))
    return out


def cohort_from_years(years: list[int]) -> int | None:
    return min(years) if years else None


def extract_params(text: str, doc_year: int) -> dict:
    params: dict = {
        "doc_year": doc_year,
        "ipc_by_year": {},
        "usd_current": None,
        "usd_avg": None,
        "cny_current": None,
        "cny_prev": None,
        "source_url": SOURCE_URLS.get(doc_year, ""),
        "notes": [],
    }

    # IPC from headers / footnotes like "ИПЦ - 112,9%"
    for m in re.finditer(
        r"(?:на 1 октября\s*)?(\d{4})\s*г\.?\s*\(ИПЦ\s*[-–—]?\s*(\d{2,3}[,.]\d)\s*%\)",
        text,
        re.I,
    ):
        y = int(m.group(1))
        v = float(m.group(2).replace(",", "."))
        params["ipc_by_year"][y] = v

    # Also standalone "ИПЦ - 12,9%" meaning growth; convert if looks like growth
    m = re.search(
        r"ИПЦ\s*[-–—]\s*индекс[^\n]{0,200}?[-–—]\s*(\d{1,2}[,.]\d)\s*%",
        text,
        re.I | re.S,
    )
    if m:
        growth = float(m.group(1).replace(",", "."))
        params["notes"].append(f"ИПЦ прирост (из сноски): {growth}%")

    # USD
    m = re.search(
        r"USDt\+1[^\n]{0,200}?[-–—]\s*(\d{2,3}[,.]\d{2})\s*тенге",
        text,
        re.I | re.S,
    )
    if m:
        params["usd_current"] = float(m.group(1).replace(",", "."))
    m = re.search(
        r"USDt\s*[-–—][^\n]{0,200}?[-–—]\s*(\d{2,3}[,.]\d{2})\s*тенге",
        text,
        re.I | re.S,
    )
    if m:
        params["usd_avg"] = float(m.group(1).replace(",", "."))
    # fallback: "X тенге за 1 доллар" occurrences
    usd_vals = [
        float(x.replace(",", "."))
        for x in re.findall(r"(\d{2,3}[,.]\d{2})\s*тенге за 1 доллар", text, re.I)
    ]
    if usd_vals:
        if params["usd_current"] is None and len(usd_vals) >= 1:
            params["usd_current"] = usd_vals[0]
        if params["usd_avg"] is None and len(usd_vals) >= 2:
            params["usd_avg"] = usd_vals[1]

    # CNY
    cny_vals = [
        float(x.replace(",", "."))
        for x in re.findall(r"(\d{2,3}[,.]\d{2})\s*тенге за 1 китайск", text, re.I)
    ]
    if cny_vals:
        params["cny_current"] = cny_vals[0]
        if len(cny_vals) >= 2:
            params["cny_prev"] = cny_vals[1]

    return params


def is_fixed_row(row: list, kind_hint: str) -> bool:
    return kind_hint == "fixed" and looks_like_data_row(row)


def parse_fixed_row(
    row: list, year_cols: list[tuple[int, int, float | None]], cohort: int | None
) -> dict | None:
    # Expected: type, unit, approved, formula, years...
    # Sometimes approved/formula missing (Kaz PV special formula blank)
    cells = [clean(c) for c in row]
    if not cells or not any(cells):
        return None

    # Find unit column
    unit_idx = next((i for i, c in enumerate(cells) if "тенге" in c.lower()), None)
    if unit_idx is None:
        # continuation without unit? skip unless type-like
        if not looks_like_data_row(row):
            return None
        unit_idx = 1 if len(cells) > 1 else 0

    type_name = " ".join(cells[:unit_idx]).strip() if unit_idx > 0 else cells[0]
    if not type_name:
        return None

    unit = cells[unit_idx] if unit_idx < len(cells) else "тенге/кВтч"
    approved = None
    formula = ""
    # after unit: approved number, then formula
    rest_start = unit_idx + 1
    if rest_start < len(cells):
        approved = parse_num(cells[rest_start])
    # formula cell: contains Tt or ИПЦ or валюта or empty
    formula_idx = None
    for i in range(rest_start + 1, len(cells)):
        c = cells[i]
        if re.search(r"Tt\+1|Тt\+1|ИПЦ|валюта|USD", c, re.I):
            formula_idx = i
            formula = norm_formula(c)
            break
        if c == "" and approved is not None:
            # blank formula (Kaz PV often)
            formula_idx = i
            formula = ""
            break

    by_year: dict[int, float] = {}
    for col_i, year, _ipc in year_cols:
        if col_i < len(cells):
            val = parse_num(cells[col_i])
            if val is not None:
                by_year[year] = val

    # If year_cols empty (continuation without header), collect trailing numbers
    if not year_cols:
        nums = []
        for i, c in enumerate(cells):
            if i <= (formula_idx or rest_start):
                continue
            n = parse_num(c)
            if n is not None:
                nums.append(n)
        # store under synthetic keys later if needed
        if nums and approved is not None:
            by_year = {}  # will be filled by caller with known years
            return {
                "type": type_name,
                "unit": unit,
                "approved": approved,
                "formula": formula,
                "cohort_year": cohort,
                "by_year": by_year,
                "_orphan_nums": nums,
            }

    if approved is None and not by_year:
        return None

    return {
        "type": type_name,
        "unit": unit,
        "approved": approved,
        "formula": formula,
        "cohort_year": cohort,
        "by_year": by_year,
    }


def parse_auction_row(
    row: list, year_cols: list[tuple[int, int, float | None]], cohort: int | None
) -> dict | None:
    cells = [clean(c) for c in row]
    if not cells or not any(cells):
        return None

    # Find №
    num = None
    start = 0
    if cells and re.fullmatch(r"\d+", cells[0]):
        num = int(cells[0])
        start = 1

    unit_idx = next((i for i, c in enumerate(cells[start:], start) if "тенге" in c.lower()), None)
    if unit_idx is None:
        return None

    type_name = " ".join(cells[start:unit_idx]).strip()
    unit = cells[unit_idx]
    auction_price = None
    formula = ""
    rest = unit_idx + 1
    if rest < len(cells):
        auction_price = parse_num(cells[rest])

    formula_idx = None
    for i in range(rest + 1, len(cells)):
        c = cells[i]
        if re.search(r"Tt\+1|Тt\+1|ИПЦ|валюта|USD", c, re.I):
            formula_idx = i
            formula = norm_formula(c)
            break

    # Detect construction-indexation column name via header later; here price is auction_price
    by_year: dict[int, float] = {}
    for col_i, year, _ipc in year_cols:
        if col_i < len(cells):
            val = parse_num(cells[col_i])
            if val is not None:
                by_year[year] = val

    if not type_name or (auction_price is None and not by_year):
        return None

    return {
        "num": num,
        "type": type_name,
        "unit": unit,
        "auction_price": auction_price,
        "formula": formula,
        "cohort_year": cohort,
        "by_year": by_year,
        "price_label": "Аукционная цена",
    }


def section_cohort_from_text(text_before: str) -> int | None:
    """Find cohort start year from nearest section title above the table."""
    titles = list(
        re.finditer(
            r"Расчет индексации[^\n]{0,200}?в\s+(\d{4})(?:\s*[-–—]\s*(\d{4}))?",
            text_before,
            re.I,
        )
    )
    if not titles:
        # "Для первого применения ... в 2026 году"
        m = list(
            re.finditer(
                r"первого применения[^\n]{0,80}?в\s+(\d{4})",
                text_before,
                re.I,
            )
        )
        if m:
            return int(m[-1].group(1))
        return None
    last = titles[-1]
    return int(last.group(1))


def parse_pdf(path: Path, doc_year: int) -> dict:
    fixed_rows: list[dict] = []
    auction_rows: list[dict] = []
    full_text_parts: list[str] = []

    with pdfplumber.open(path) as pdf:
        current_kind: str | None = None
        current_year_cols: list[tuple[int, int, float | None]] = []
        current_cohort: int | None = None
        current_price_label = "Аукционная цена"
        text_so_far = ""

        for page in pdf.pages:
            page_text = page.extract_text() or ""
            full_text_parts.append(page_text)
            tables = page.extract_tables() or []

            # Approximate vertical positions: process tables in order
            for table in tables:
                if not table:
                    continue
                header = table[0]
                kind = classify_header(header)

                if kind:
                    current_kind = kind
                    current_year_cols = parse_year_columns(header)
                    years = [y for _, y, _ in current_year_cols]
                    # Cohort = earliest 1-Oct year column in this table
                    current_cohort = cohort_from_years(years)
                    # For «первого применения» tables with a single latest-year column,
                    # keep that year as cohort (first indexation calendar year).
                    hdr_blob = " ".join(clean(c) for c in header)
                    if "единовременн" in hdr_blob.lower() or "строительств" in hdr_blob.lower():
                        current_price_label = (
                            "Аукционная цена с учетом единовременной индексации "
                            "на период строительства"
                        )
                    else:
                        current_price_label = "Аукционная цена"
                    data_rows = table[1:]
                else:
                    # continuation of previous table
                    data_rows = table
                    if current_kind is None:
                        # guess
                        blob = " ".join(clean(c) for row in table[:2] for c in row).lower()
                        if re.match(r"^\d+", clean(table[0][0] or "")) or "аукцион" in blob:
                            current_kind = "auction"
                        else:
                            current_kind = "fixed"

                for row in data_rows:
                    if not looks_like_data_row(row):
                        continue
                    if current_kind == "fixed":
                        parsed = parse_fixed_row(row, current_year_cols, current_cohort)
                        if not parsed:
                            continue
                        # orphan nums: assign to year_cols years if lengths match
                        if "_orphan_nums" in parsed:
                            nums = parsed.pop("_orphan_nums")
                            years = [y for _, y, _ in current_year_cols]
                            if years and len(nums) == len(years):
                                parsed["by_year"] = dict(zip(years, nums))
                            elif years and len(nums) == len(years) + 1:
                                # first num may be approved already set
                                parsed["by_year"] = dict(zip(years, nums[-len(years) :]))
                            else:
                                # try match trailing
                                if years and len(nums) >= len(years):
                                    parsed["by_year"] = dict(
                                        zip(years, nums[-len(years) :])
                                    )
                        if parsed.get("by_year") or parsed.get("approved") is not None:
                            fixed_rows.append(parsed)
                    else:
                        parsed = parse_auction_row(row, current_year_cols, current_cohort)
                        if parsed:
                            parsed["price_label"] = current_price_label
                            auction_rows.append(parsed)

            text_so_far += "\n" + page_text

    full_text = "\n".join(full_text_parts)
    params = extract_params(full_text, doc_year)

    # Enrich ipc_by_year from table headers already collected
    for row in fixed_rows + auction_rows:
        for y in row.get("by_year", {}):
            params["ipc_by_year"].setdefault(y, None)

    # Re-scan headers for ipc
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                if not table:
                    continue
                for col_i, year, ipc in parse_year_columns(table[0]):
                    if ipc is not None:
                        params["ipc_by_year"][year] = ipc

    return {
        "doc_year": doc_year,
        "fixed": fixed_rows,
        "auction": auction_rows,
        "params": params,
        "path": str(path),
    }


def latest_value(by_year: dict[int, float], doc_year: int) -> tuple[int | None, float | None]:
    if not by_year:
        return None, None
    # Document "на YYYY год" → values as of 1 Oct (YYYY-1)
    target = doc_year - 1
    if target in by_year:
        return target, by_year[target]
    y = max(by_year)
    return y, by_year[y]


def verify_ipc_chain(rows: list[dict], ipc_by_year: dict[int, float | None]) -> list[str]:
    """Return soft warnings for pure-IPC rows where T*ipc != next."""
    warnings = []
    for r in rows:
        formula = (r.get("formula") or "").replace(" ", "")
        if "ИПЦ" not in formula or "USD" in formula or "валюта" in formula:
            continue
        if "*ИПЦ" not in formula and "ИПЦ" not in formula:
            continue
        # skip mixed formulas
        if "0,3" in formula or "0.3" in formula or "валюта" in formula.lower():
            continue
        years = sorted(r["by_year"])
        base = r.get("approved") if "approved" in r else r.get("auction_price")
        prev = base
        label = r.get("type", "?")[:40]
        for y in years:
            ipc = ipc_by_year.get(y)
            if prev is None or ipc is None:
                prev = r["by_year"][y]
                continue
            expected = round(prev * (ipc / 100.0) - 1e-12, 2)  # floor-ish to 2 decimals
            # RFC rounds down to tiyn (2 decimals)
            expected_floor = int(prev * (ipc / 100.0) * 100) / 100.0
            actual = r["by_year"][y]
            if abs(actual - expected_floor) > 0.02 and abs(actual - expected) > 0.02:
                warnings.append(
                    f"{label} @ {y}: got {actual}, expected~{expected_floor} "
                    f"(prev={prev}, ipc={ipc})"
                )
            prev = actual
    return warnings


def style_header(ws, headers: list[str]) -> None:
    for col, h in enumerate(headers, 1):
        cell = ws.cell(1, col, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = THIN
    ws.row_dimensions[1].height = 36
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"


def autosize(ws, widths: dict[str, float] | None = None) -> None:
    if widths:
        for letter, w in widths.items():
            ws.column_dimensions[letter].width = w
        return
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        maxlen = 0
        for cell in col:
            if cell.value is not None:
                maxlen = max(maxlen, min(len(str(cell.value)), 60))
        ws.column_dimensions[letter].width = max(10, maxlen + 2)


def all_years_from_rows(rows: list[dict]) -> list[int]:
    ys = set()
    for r in rows:
        ys.update(r.get("by_year", {}).keys())
    return sorted(ys)


def write_year_workbook(data: dict, out_path: Path) -> None:
    doc_year = data["doc_year"]
    fixed = data["fixed"]
    auction = data["auction"]
    params = data["params"]

    wb = Workbook()

    # --- Фикс_тарифы ---
    ws = wb.active
    ws.title = "Фикс_тарифы"
    year_cols = all_years_from_rows(fixed)
    headers = [
        "Тип ВИЭ",
        "Ед. изм",
        "Утверждённый тариф",
        "Год первой индексации (когорта)",
        "Формула",
    ] + [f"На 1 октября {y}" for y in year_cols] + [
        f"Итого на 1 октября {doc_year - 1}",
    ]
    style_header(ws, headers)
    for i, r in enumerate(fixed, 2):
        ly, lv = latest_value(r["by_year"], doc_year)
        vals = [
            r["type"],
            r["unit"],
            r.get("approved"),
            r.get("cohort_year"),
            r.get("formula") or "",
        ]
        for y in year_cols:
            vals.append(r["by_year"].get(y))
        vals.append(lv)
        for col, v in enumerate(vals, 1):
            cell = ws.cell(i, col, v)
            cell.border = THIN
            cell.alignment = Alignment(vertical="center", wrap_text=col == 1)
            if isinstance(v, float):
                cell.number_format = NUM_FMT
    autosize(
        ws,
        {
            "A": 55,
            "B": 12,
            "C": 16,
            "D": 18,
            "E": 40,
            **{get_column_letter(6 + i): 14 for i in range(len(year_cols) + 1)},
        },
    )

    # --- Аукционные_цены ---
    ws2 = wb.create_sheet("Аукционные_цены")
    ay = all_years_from_rows(auction)
    headers2 = [
        "№",
        "Тип ВИЭ",
        "Ед. изм",
        "Базовая цена",
        "Тип базовой цены",
        "Год первой индексации (когорта)",
        "Формула",
    ] + [f"На 1 октября {y}" for y in ay] + [
        f"Итого на 1 октября {doc_year - 1}",
    ]
    style_header(ws2, headers2)
    for i, r in enumerate(auction, 2):
        ly, lv = latest_value(r["by_year"], doc_year)
        vals = [
            r.get("num"),
            r["type"],
            r["unit"],
            r.get("auction_price"),
            r.get("price_label") or "Аукционная цена",
            r.get("cohort_year"),
            r.get("formula") or "",
        ]
        for y in ay:
            vals.append(r["by_year"].get(y))
        vals.append(lv)
        for col, v in enumerate(vals, 1):
            cell = ws2.cell(i, col, v)
            cell.border = THIN
            cell.alignment = Alignment(vertical="center", wrap_text=col in {2, 5, 7})
            if isinstance(v, float):
                cell.number_format = NUM_FMT
    autosize(
        ws2,
        {
            "A": 6,
            "B": 28,
            "C": 12,
            "D": 14,
            "E": 40,
            "F": 18,
            "G": 45,
            **{get_column_letter(8 + i): 14 for i in range(len(ay) + 1)},
        },
    )

    # --- Параметры ---
    ws3 = wb.create_sheet("Параметры")
    ws3["A1"] = f"Параметры индексации — документ РФЦ на {doc_year} год"
    ws3["A1"].font = TITLE_FONT
    ws3["A2"] = "Источник PDF"
    ws3["B2"] = params.get("source_url") or data["path"]
    ws3["A3"] = "Локальный файл"
    ws3["B3"] = data["path"]
    ws3["A5"] = "Год (1 октября)"
    ws3["B5"] = "ИПЦ, %"
    ws3["A5"].font = Font(bold=True)
    ws3["B5"].font = Font(bold=True)
    row = 6
    for y in sorted(k for k, v in params["ipc_by_year"].items() if k is not None and v is not None):
        ws3.cell(row, 1, y)
        val = params["ipc_by_year"][y]
        cell = ws3.cell(row, 2, val)
        if isinstance(val, float):
            cell.number_format = "0.0"
        row += 1
    row += 1
    ws3.cell(row, 1, "USD текущий (USDt+1), тенге").font = Font(bold=True)
    ws3.cell(row, 2, params.get("usd_current"))
    row += 1
    ws3.cell(row, 1, "USD средний (USDt), тенге").font = Font(bold=True)
    ws3.cell(row, 2, params.get("usd_avg"))
    row += 1
    ws3.cell(row, 1, "CNY текущий (валютаt+1), тенге").font = Font(bold=True)
    ws3.cell(row, 2, params.get("cny_current"))
    row += 1
    ws3.cell(row, 1, "CNY предыдущий (валютаt), тенге").font = Font(bold=True)
    ws3.cell(row, 2, params.get("cny_prev"))
    row += 2
    ws3.cell(row, 1, "Примечание").font = Font(bold=True)
    row += 1
    ws3.cell(
        row,
        1,
        "Заголовок документа «на YYYY год» соответствует индексации на 1 октября (YYYY−1).",
    )
    for note in params.get("notes") or []:
        row += 1
        ws3.cell(row, 1, note)
    ws3.column_dimensions["A"].width = 45
    ws3.column_dimensions["B"].width = 70

    # --- Сводка ---
    ws4 = wb.create_sheet("Сводка")
    ws4["A1"] = f"Актуальные значения на 1 октября {doc_year - 1} (документ на {doc_year} год)"
    ws4["A1"].font = TITLE_FONT
    headers4 = [
        "Категория",
        "№",
        "Тип ВИЭ",
        "Когорта (год 1-й индексации)",
        "Базовый тариф/цена",
        "Формула",
        f"Значение на 1.10.{doc_year - 1}",
    ]
    for col, h in enumerate(headers4, 1):
        cell = ws4.cell(3, col, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = THIN
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws4.row_dimensions[3].height = 32
    r_i = 4
    for r in fixed:
        _, lv = latest_value(r["by_year"], doc_year)
        vals = [
            "Фиксированный тариф",
            None,
            r["type"],
            r.get("cohort_year"),
            r.get("approved"),
            r.get("formula") or "",
            lv,
        ]
        for col, v in enumerate(vals, 1):
            cell = ws4.cell(r_i, col, v)
            cell.border = THIN
            if isinstance(v, float):
                cell.number_format = NUM_FMT
        r_i += 1
    for r in auction:
        _, lv = latest_value(r["by_year"], doc_year)
        vals = [
            "Аукционная цена",
            r.get("num"),
            r["type"],
            r.get("cohort_year"),
            r.get("auction_price"),
            r.get("formula") or "",
            lv,
        ]
        for col, v in enumerate(vals, 1):
            cell = ws4.cell(r_i, col, v)
            cell.border = THIN
            if isinstance(v, float):
                cell.number_format = NUM_FMT
        r_i += 1
    autosize(
        ws4,
        {"A": 22, "B": 6, "C": 55, "D": 16, "E": 14, "F": 45, "G": 16},
    )

    wb.save(out_path)


def write_summary(all_data: list[dict], out_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Фикс_все_годы"
    headers = [
        "Год документа",
        "Тип ВИЭ",
        "Когорта",
        "Утверждённый тариф",
        "Формула",
        "Год значения",
        "Значение на 1 октября",
    ]
    style_header(ws, headers)
    row = 2
    for data in all_data:
        for r in data["fixed"]:
            ly, lv = latest_value(r["by_year"], data["doc_year"])
            vals = [
                data["doc_year"],
                r["type"],
                r.get("cohort_year"),
                r.get("approved"),
                r.get("formula") or "",
                ly,
                lv,
            ]
            for col, v in enumerate(vals, 1):
                cell = ws.cell(row, col, v)
                cell.border = THIN
                if isinstance(v, float):
                    cell.number_format = NUM_FMT
            row += 1
    autosize(ws, {"A": 14, "B": 55, "C": 10, "D": 14, "E": 40, "F": 12, "G": 16})

    ws2 = wb.create_sheet("Аукцион_все_годы")
    headers2 = [
        "Год документа",
        "№",
        "Тип ВИЭ",
        "Когорта",
        "Базовая цена",
        "Тип базовой цены",
        "Формула",
        "Год значения",
        "Значение на 1 октября",
    ]
    style_header(ws2, headers2)
    row = 2
    for data in all_data:
        for r in data["auction"]:
            ly, lv = latest_value(r["by_year"], data["doc_year"])
            vals = [
                data["doc_year"],
                r.get("num"),
                r["type"],
                r.get("cohort_year"),
                r.get("auction_price"),
                r.get("price_label") or "",
                r.get("formula") or "",
                ly,
                lv,
            ]
            for col, v in enumerate(vals, 1):
                cell = ws2.cell(row, col, v)
                cell.border = THIN
                if isinstance(v, float):
                    cell.number_format = NUM_FMT
            row += 1
    autosize(
        ws2,
        {"A": 14, "B": 6, "C": 28, "D": 10, "E": 12, "F": 40, "G": 40, "H": 12, "I": 16},
    )

    ws3 = wb.create_sheet("Параметры_по_годам")
    headers3 = [
        "Год документа",
        "USD текущий",
        "USD средний",
        "CNY текущий",
        "CNY предыдущий",
        "Источник",
    ]
    style_header(ws3, headers3)
    # Also add IPC columns dynamically
    all_ipc_years = sorted(
        {
            y
            for d in all_data
            for y in d["params"]["ipc_by_year"]
            if d["params"]["ipc_by_year"][y] is not None
        }
    )
    # rebuild with ipc
    wb.remove(ws3)
    ws3 = wb.create_sheet("Параметры_по_годам")
    headers3 = (
        ["Год документа"]
        + [f"ИПЦ {y}" for y in all_ipc_years]
        + ["USD текущий", "USD средний", "CNY текущий", "CNY предыдущий", "Источник"]
    )
    style_header(ws3, headers3)
    for i, data in enumerate(all_data, 2):
        p = data["params"]
        vals = [data["doc_year"]]
        for y in all_ipc_years:
            vals.append(p["ipc_by_year"].get(y))
        vals += [
            p.get("usd_current"),
            p.get("usd_avg"),
            p.get("cny_current"),
            p.get("cny_prev"),
            p.get("source_url"),
        ]
        for col, v in enumerate(vals, 1):
            cell = ws3.cell(i, col, v)
            cell.border = THIN
    autosize(ws3)

    wb.save(out_path)


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    all_data = []
    for year in YEARS:
        path = SOURCES / f"indexation_{year}.pdf"
        if not path.exists():
            raise FileNotFoundError(path)
        print(f"Parsing {path.name}...")
        data = parse_pdf(path, year)
        warns = verify_ipc_chain(data["fixed"], data["params"]["ipc_by_year"])
        warns += verify_ipc_chain(data["auction"], data["params"]["ipc_by_year"])
        print(
            f"  fixed={len(data['fixed'])} auction={len(data['auction'])} "
            f"ipc_years={sorted(y for y,v in data['params']['ipc_by_year'].items() if v)}"
        )
        if warns:
            print(f"  IPC check warnings: {len(warns)}")
            for w in warns[:5]:
                print(f"    - {w}")
        out = OUT_DIR / f"Индексация_тарифов_ВИЭ_{year}.xlsx"
        write_year_workbook(data, out)
        print(f"  -> {out.name}")
        all_data.append(data)

    summary = OUT_DIR / "Индексация_тарифов_ВИЭ_сводка_2021-2026.xlsx"
    write_summary(all_data, summary)
    print(f"Summary -> {summary.name}")

    # Spot-check known 2026 values
    d2026 = next(d for d in all_data if d["doc_year"] == 2026)
    expo = [r for r in d2026["fixed"] if r.get("approved") == 59.7]
    if expo:
        r = expo[0]
        print(
            f"Spot-check EXPO: approved={r.get('approved')} "
            f"2025={r['by_year'].get(2025)} cohort={r.get('cohort_year')}"
        )
    wind = [
        r
        for r in d2026["fixed"]
        if r.get("approved") == 22.68 and r.get("cohort_year") == 2024
    ]
    if wind:
        r = wind[0]
        print(
            f"Spot-check Wind cohort 2024: 2024={r['by_year'].get(2024)} "
            f"2025={r['by_year'].get(2025)} (expect 24.56 / 27.72)"
        )
    print("Done.")


if __name__ == "__main__":
    main()
