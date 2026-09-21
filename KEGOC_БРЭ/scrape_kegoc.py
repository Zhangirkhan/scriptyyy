#!/usr/bin/env python3
"""CLI: scrape, download, and assemble KEGOC BRE monthly Excel into CSV."""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

import requests

from kegoc_bre.download import download_entries
from kegoc_bre.parse import build_processed, write_processed
from kegoc_bre.scrape import ZONE_PAGES, filter_entries, scrape_zones


def parse_ym(value: str):
    try:
        dt = datetime.strptime(value, "%Y-%m")
        return dt.year, dt.month
    except ValueError as exc:
        raise argparse.ArgumentTypeError(
            f"Expected YYYY-MM, got {value!r}"
        ) from exc


def build_argparser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description=(
            "Download KEGOC balancing market (БРЭ) monthly Excel files "
            "for Север-Юг and Запад and assemble unified CSV tables."
        )
    )
    p.add_argument(
        "--data-dir",
        type=Path,
        default=Path("data"),
        help="Output root (default: ./data)",
    )
    p.add_argument(
        "--zone",
        choices=list(ZONE_PAGES) + ["all"],
        default="all",
        help="Zone to scrape (default: all)",
    )
    p.add_argument("--year", type=int, default=None, help="Filter by year")
    p.add_argument(
        "--month",
        type=int,
        default=None,
        choices=range(1, 13),
        metavar="{1..12}",
        help="Filter by month",
    )
    p.add_argument(
        "--from",
        dest="date_from",
        type=parse_ym,
        default=None,
        help="Inclusive start YYYY-MM",
    )
    p.add_argument(
        "--to",
        dest="date_to",
        type=parse_ym,
        default=None,
        help="Inclusive end YYYY-MM",
    )
    p.add_argument(
        "--skip-download",
        action="store_true",
        help="Only rebuild CSV from already downloaded xlsx files",
    )
    p.add_argument(
        "--force",
        action="store_true",
        help="Re-download even if local file exists",
    )
    p.add_argument(
        "--delay",
        type=float,
        default=0.4,
        help="Pause between downloads in seconds (default: 0.4)",
    )
    return p


def discover_local(data_dir: Path, zones) -> list:
    """Build MonthEntry-like records from data/raw when --skip-download."""
    records = []
    for zone in zones:
        root = data_dir / "raw" / zone
        if not root.exists():
            continue
        for year_dir in sorted(root.iterdir()):
            if not year_dir.is_dir() or not year_dir.name.isdigit():
                continue
            year = int(year_dir.name)
            for f in sorted(year_dir.glob("*.xlsx")):
                try:
                    month = int(f.stem)
                except ValueError:
                    continue
                records.append(
                    {
                        "zone": zone,
                        "year": year,
                        "month": month,
                        "title": f"{year}-{month:02d}",
                        "url": "",
                        "path": str(f.relative_to(data_dir)),
                        "abs_path": str(f),
                        "size": f.stat().st_size,
                        "skipped": True,
                    }
                )
    return records


def main(argv=None) -> int:
    args = build_argparser().parse_args(argv)
    data_dir: Path = args.data_dir
    data_dir.mkdir(parents=True, exist_ok=True)

    zones = list(ZONE_PAGES) if args.zone == "all" else [args.zone]
    session = requests.Session()

    if args.skip_download:
        catalog = discover_local(data_dir, zones)
        catalog = [
            r
            for r in catalog
            if (args.year is None or r["year"] == args.year)
            and (args.month is None or r["month"] == args.month)
            and (args.date_from is None or (r["year"], r["month"]) >= args.date_from)
            and (args.date_to is None or (r["year"], r["month"]) <= args.date_to)
        ]
        print(f"Found {len(catalog)} local xlsx file(s)")
    else:
        print(f"Scraping pages for: {', '.join(zones)}")
        entries = scrape_zones(zones, session=session)
        entries = filter_entries(
            entries,
            year=args.year,
            month=args.month,
            date_from=args.date_from,
            date_to=args.date_to,
        )
        print(f"Found {len(entries)} monthly file(s) on site")
        if not entries:
            print("Nothing to download.", file=sys.stderr)
            return 1
        catalog = download_entries(
            entries,
            data_dir,
            session=session,
            force=args.force,
            delay=args.delay,
        )
        downloaded = sum(1 for r in catalog if not r["skipped"])
        skipped = sum(1 for r in catalog if r["skipped"])
        print(f"Downloaded {downloaded}, skipped existing {skipped}")

    print("Parsing Excel and writing CSV...")
    processed = build_processed(catalog, data_dir)
    written = write_processed(processed, data_dir)

    catalog_path = data_dir / "catalog.json"
    catalog_path.write_text(
        json.dumps(catalog, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(f"Catalog: {catalog_path}")
    for path in written:
        # row count from file would be slow; report path only
        print(f"  wrote {path}")

    for zone, kinds in processed.items():
        for kind, df in kinds.items():
            print(f"  {zone}/{kind}: {len(df)} rows")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
