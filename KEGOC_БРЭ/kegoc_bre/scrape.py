"""Parse KEGOC balancing-market accordion pages for monthly Excel links."""

from __future__ import annotations

import re
from dataclasses import asdict, dataclass
from typing import Iterable, List, Optional
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

ZONE_PAGES = {
    "sever-yug": "https://www.kegoc.kz/ru/electric-power/balancing-electricity-market/sever-yug/",
    "zapad": "https://www.kegoc.kz/ru/electric-power/balancing-electricity-market/zapad/",
}

MONTH_NAMES_RU = {
    "январь": 1,
    "февраль": 2,
    "март": 3,
    "апрель": 4,
    "май": 5,
    "июнь": 6,
    "июль": 7,
    "август": 8,
    "сентябрь": 9,
    "октябрь": 10,
    "ноябрь": 11,
    "декабрь": 12,
}

USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)

YEAR_RE = re.compile(r"(\d{4})")
# "1 Январь 2026 года СЮ" / "7 Июль 2023 года запад"
MONTH_LINK_RE = re.compile(
    r"(?P<ord>\d+)\s+(?P<month>[А-Яа-яA-Za-z]+)\s+(?P<year>\d{4})",
    re.UNICODE,
)


@dataclass
class MonthEntry:
    zone: str
    year: int
    month: int
    title: str
    url: str

    def to_dict(self) -> dict:
        return asdict(self)


def fetch_html(url: str, session: Optional[requests.Session] = None) -> str:
    sess = session or requests.Session()
    resp = sess.get(url, headers={"User-Agent": USER_AGENT}, timeout=60)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def _parse_month_from_title(title: str) -> Optional[tuple]:
    m = MONTH_LINK_RE.search(title.replace("\xa0", " ").strip())
    if not m:
        return None
    month_raw = m.group("month").lower()
    month = MONTH_NAMES_RU.get(month_raw)
    if month is None:
        return None
    return int(m.group("year")), month


def parse_accordion(html: str, zone: str, base_url: str) -> List[MonthEntry]:
    soup = BeautifulSoup(html, "lxml")
    entries: List[MonthEntry] = []
    seen = set()

    for container in soup.select(".accordeon .acc-container"):
        year_hint = None
        title_el = container.select_one(".title .name")
        if title_el:
            ym = YEAR_RE.search(title_el.get_text(" ", strip=True))
            if ym:
                year_hint = int(ym.group(1))

        for a in container.select("a.corporative-docs, a[href*='.xlsx']"):
            href = a.get("href") or ""
            if ".xlsx" not in href.lower():
                continue
            title = a.select_one(".name")
            text = (title.get_text(" ", strip=True) if title else a.get_text(" ", strip=True))
            text = re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()
            parsed = _parse_month_from_title(text)
            if not parsed:
                continue
            year, month = parsed
            if year_hint and year != year_hint:
                year = year_hint
            url = urljoin(base_url, href)
            key = (zone, year, month, url)
            if key in seen:
                continue
            seen.add(key)
            entries.append(
                MonthEntry(zone=zone, year=year, month=month, title=text, url=url)
            )

    # Fallback: any xlsx links on the page
    if not entries:
        for a in soup.select("a[href*='.xlsx']"):
            href = a.get("href") or ""
            text = re.sub(r"\s+", " ", a.get_text(" ", strip=True).replace("\xa0", " "))
            parsed = _parse_month_from_title(text)
            if not parsed:
                continue
            year, month = parsed
            url = urljoin(base_url, href)
            key = (zone, year, month, url)
            if key in seen:
                continue
            seen.add(key)
            entries.append(
                MonthEntry(zone=zone, year=year, month=month, title=text, url=url)
            )

    entries.sort(key=lambda e: (e.zone, e.year, e.month))
    return entries


def scrape_zone(
    zone: str, session: Optional[requests.Session] = None
) -> List[MonthEntry]:
    if zone not in ZONE_PAGES:
        raise ValueError(f"Unknown zone: {zone}. Expected one of {list(ZONE_PAGES)}")
    url = ZONE_PAGES[zone]
    html = fetch_html(url, session=session)
    return parse_accordion(html, zone=zone, base_url=url)


def scrape_zones(
    zones: Iterable[str], session: Optional[requests.Session] = None
) -> List[MonthEntry]:
    sess = session or requests.Session()
    all_entries: List[MonthEntry] = []
    for zone in zones:
        all_entries.extend(scrape_zone(zone, session=sess))
    return all_entries


def filter_entries(
    entries: Iterable[MonthEntry],
    *,
    year: Optional[int] = None,
    month: Optional[int] = None,
    date_from: Optional[tuple] = None,
    date_to: Optional[tuple] = None,
) -> List[MonthEntry]:
    result = []
    for e in entries:
        if year is not None and e.year != year:
            continue
        if month is not None and e.month != month:
            continue
        if date_from is not None and (e.year, e.month) < date_from:
            continue
        if date_to is not None and (e.year, e.month) > date_to:
            continue
        result.append(e)
    return result
