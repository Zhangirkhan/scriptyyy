"""Download monthly Excel files from KEGOC."""

from __future__ import annotations

import time
from pathlib import Path
from typing import List, Optional

import requests

from .scrape import USER_AGENT, MonthEntry


def raw_path(data_dir: Path, entry: MonthEntry) -> Path:
    return data_dir / "raw" / entry.zone / f"{entry.year:04d}" / f"{entry.month:02d}.xlsx"


def download_file(
    url: str,
    dest: Path,
    session: Optional[requests.Session] = None,
    *,
    force: bool = False,
) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists() and dest.stat().st_size > 0 and not force:
        return dest

    sess = session or requests.Session()
    resp = sess.get(url, headers={"User-Agent": USER_AGENT}, timeout=120)
    resp.raise_for_status()
    tmp = dest.with_suffix(dest.suffix + ".part")
    tmp.write_bytes(resp.content)
    tmp.replace(dest)
    return dest


def download_entries(
    entries: List[MonthEntry],
    data_dir: Path,
    session: Optional[requests.Session] = None,
    *,
    force: bool = False,
    delay: float = 0.4,
) -> List[dict]:
    sess = session or requests.Session()
    results = []
    for i, entry in enumerate(entries):
        dest = raw_path(data_dir, entry)
        skipped = dest.exists() and dest.stat().st_size > 0 and not force
        if not skipped:
            download_file(entry.url, dest, session=sess, force=force)
            if delay and i < len(entries) - 1:
                time.sleep(delay)
        results.append(
            {
                **entry.to_dict(),
                "path": str(dest.relative_to(data_dir)),
                "abs_path": str(dest),
                "size": dest.stat().st_size if dest.exists() else 0,
                "skipped": skipped,
            }
        )
    return results
