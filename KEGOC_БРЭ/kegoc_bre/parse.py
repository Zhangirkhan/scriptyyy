"""Parse KEGOC BRE Excel workbooks and normalize into DataFrames/CSV."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd

SHEET_ALIASES = {
    "zona": (
        "зона север-юг",
        "зона запад",
        "клиринговая цена север-юг",
        "клиринговая цена запад",
        "север-юг",
        "запад",
    ),
    "rf": ("рф",),
    "zayavki": ("заявки на участие в бал-ии", "заявки"),
    "dohody": ("доходы затраты рц брэ", "доходы затраты"),
}

# Map older / alternate column labels onto the 2024+ schema where possible.
COLUMN_ALIASES: Dict[str, str] = {
    # zayavki 2023
    "объем на понижение": "Фактический дисбаланс по заявке на участие в балансировании Получение",
    "объем на повышение": "Фактический дисбаланс по заявке на участие в балансировании Поставка",
    "цена": "Цена по заявке на участие в балансировании",
    "доход": "Стоимость по заявке на участие в балансировании Доходы \"+\"",
    "затраты": "Стоимость по заявке на участие в балансировании Затраты \"-\"",
    # dohody 2023
    "доходы/затраты от арчм": "Доходы и затраты РЦ БРЭ от АРЧМ, тенге",
    "доходы/затраты от энергопередающих организаций": "Доходы  (затраты) РЦ БРЭ от  энергопередающих организаций, тенге",
    "доходы от срп": "Доходы РЦ БРЭ от субъектов СРП, тенге",
    "операционные затраты": "Доход РЦ БРЭ от деятельности на БРЭ, тенге",
    "итого доходы / затраты рц брэ": "Затраты (доходы), тенге",
    # zona 2023 shorter income/expense labels
    "сумма доходов от продажи отрицательных дисбалансов в час на повышение, тг": (
        "Сумма доходов субъектов без предельного тарифа от продажи отрицательных дисбалансов"
    ),
    "сумма расходов от покупки положительных дисбалансов в час на понижение, тг": (
        "Сумма расходов субъектов без предельного тарифа от покупки положительных дисбалансов"
    ),
}

META_COLS = ["zone", "year", "month", "source_file", "sheet_raw", "sheet_kind"]


def classify_sheet(name: str) -> Optional[str]:
    n = re.sub(r"\s+", " ", name.strip().lower())
    for kind, aliases in SHEET_ALIASES.items():
        for alias in aliases:
            if alias in n or n == alias:
                return kind
    # fuzzy: zona sheets often contain "зона" or "клиринговая"
    if "зона" in n or "клиринговая" in n:
        return "zona"
    if n.startswith("рф") or n == "рф":
        return "rf"
    if "заявк" in n:
        return "zayavki"
    if "доход" in n or "рц" in n:
        return "dohody"
    return None


def normalize_header(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).replace("\xa0", " ")
    text = re.sub(r"[\r\n]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def canonicalize_columns(columns: Iterable[str]) -> List[str]:
    result = []
    for col in columns:
        key = col.lower().strip()
        # Prefer mapped alias when present
        mapped = COLUMN_ALIASES.get(key)
        if mapped:
            result.append(mapped)
            continue
        result.append(col)
    return result


def parse_date(value) -> Optional[str]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    if hasattr(value, "strftime"):
        try:
            return value.strftime("%Y-%m-%d")
        except Exception:
            pass
    # Excel serial dates sometimes come as numbers via openpyxl with data_only
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            # Excel epoch; only treat large serials as dates
            if value > 20000:
                dt = pd.to_datetime(value, unit="D", origin="1899-12-30")
                return dt.strftime("%Y-%m-%d")
        except Exception:
            pass
    text = str(value).strip()
    if not text:
        return None
    for fmt in (
        "%d.%m.%Y",
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%d/%m/%Y",
        "%Y/%m/%d",
    ):
        try:
            return pd.to_datetime(text, format=fmt).strftime("%Y-%m-%d")
        except (ValueError, TypeError):
            continue
    try:
        dayfirst = "." in text or "/" in text
        return pd.to_datetime(text, dayfirst=dayfirst).strftime("%Y-%m-%d")
    except Exception:
        return text

def parse_hour(value) -> Optional[int]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        h = int(value)
        if 0 <= h <= 23:
            return h
        return None
    text = str(value).strip().lower().replace(" ", "")
    if not text:
        return None
    # "0-1", "23-24", "0–1"
    m = re.match(r"^(\d{1,2})[-–—](\d{1,2})$", text)
    if m:
        return int(m.group(1)) % 24
    m = re.match(r"^(\d{1,2})$", text)
    if m:
        return int(m.group(1)) % 24
    return None


def _find_date_hour_cols(columns: List[str]) -> Tuple[Optional[str], Optional[str]]:
    date_col = hour_col = None
    for c in columns:
        cl = c.lower()
        if date_col is None and cl in ("дата", "date"):
            date_col = c
        if hour_col is None and cl in ("час", "hour"):
            hour_col = c
    return date_col, hour_col


def read_workbook(path: Path) -> Dict[str, pd.DataFrame]:
    xl = pd.ExcelFile(path, engine="openpyxl")
    sheets: Dict[str, pd.DataFrame] = {}
    for name in xl.sheet_names:
        kind = classify_sheet(name)
        if kind is None:
            continue
        df = pd.read_excel(xl, sheet_name=name, header=0, dtype=object)
        df.columns = [normalize_header(c) for c in df.columns]
        df = df.dropna(how="all").dropna(axis=1, how="all")
        df.columns = canonicalize_columns(list(df.columns))
        seen: Dict[str, int] = {}
        new_cols = []
        for c in df.columns:
            if c not in seen:
                seen[c] = 0
                new_cols.append(c)
            else:
                seen[c] += 1
                new_cols.append(f"{c}_{seen[c]}")
        df.columns = new_cols
        df.attrs["sheet_raw"] = name
        sheets[kind] = df
    return sheets


def enrich_frame(
    df: pd.DataFrame,
    *,
    zone: str,
    year: int,
    month: int,
    source_file: str,
    sheet_kind: str,
) -> pd.DataFrame:
    out = df.copy()
    date_col, hour_col = _find_date_hour_cols(list(out.columns))

    if date_col:
        out["date"] = out[date_col].map(parse_date)
    else:
        out["date"] = None

    if hour_col:
        out["hour"] = out[hour_col].map(parse_hour)
    else:
        out["hour"] = None

    out["zone"] = zone
    out["year"] = year
    out["month"] = month
    out["source_file"] = source_file
    out["sheet_raw"] = df.attrs.get("sheet_raw", sheet_kind)
    out["sheet_kind"] = sheet_kind

    # Prefer normalized date/hour at the front
    front = META_COLS + ["date", "hour"]
    # Avoid duplicating original Дата/Час at front; keep them in body
    rest = [c for c in out.columns if c not in front]
    out = out[front + rest]
    return out


def parse_file(
    path: Path,
    *,
    zone: str,
    year: int,
    month: int,
) -> Dict[str, pd.DataFrame]:
    sheets = read_workbook(path)
    result = {}
    for kind, df in sheets.items():
        result[kind] = enrich_frame(
            df,
            zone=zone,
            year=year,
            month=month,
            source_file=str(path.name),
            sheet_kind=kind,
        )
    return result


def build_processed(
    file_records: List[dict],
    data_dir: Path,
) -> Dict[str, Dict[str, pd.DataFrame]]:
    """
    Returns nested dict: zone_or_all -> sheet_kind -> DataFrame
    """
    by_zone: Dict[str, Dict[str, List[pd.DataFrame]]] = {}
    all_buckets: Dict[str, List[pd.DataFrame]] = {
        "zona": [],
        "rf": [],
        "zayavki": [],
        "dohody": [],
    }

    for rec in file_records:
        path = Path(rec.get("abs_path") or data_dir / rec["path"])
        if not path.exists():
            continue
        parsed = parse_file(
            path, zone=rec["zone"], year=rec["year"], month=rec["month"]
        )
        by_zone.setdefault(rec["zone"], {k: [] for k in all_buckets})
        for kind, df in parsed.items():
            by_zone[rec["zone"]][kind].append(df)
            all_buckets[kind].append(df)
            # annotate catalog with sheet names
            sheets = rec.setdefault("sheets", [])
            if kind not in sheets:
                sheets.append(kind)

    def concat(frames: List[pd.DataFrame]) -> pd.DataFrame:
        frames = [f for f in frames if f is not None and not f.empty]
        if not frames:
            return pd.DataFrame()
        out = pd.concat(frames, ignore_index=True, sort=False)
        if "date" in out.columns and "hour" in out.columns:
            out = out.sort_values(
                by=[c for c in ("zone", "date", "hour") if c in out.columns],
                kind="mergesort",
            )
        return out.reset_index(drop=True)

    result: Dict[str, Dict[str, pd.DataFrame]] = {}
    for zone, kinds in by_zone.items():
        result[zone] = {k: concat(v) for k, v in kinds.items()}
    result["all"] = {k: concat(v) for k, v in all_buckets.items()}
    return result


def write_processed(
    processed: Dict[str, Dict[str, pd.DataFrame]],
    data_dir: Path,
) -> List[Path]:
    written = []
    for zone, kinds in processed.items():
        out_dir = data_dir / "processed" / zone
        out_dir.mkdir(parents=True, exist_ok=True)
        for kind, df in kinds.items():
            dest = out_dir / f"{kind}.csv"
            df.to_csv(dest, index=False, encoding="utf-8-sig")
            written.append(dest)
    return written
