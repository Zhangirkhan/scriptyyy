#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Export RFC VIE list with coordinates from OSM / GPPD / Photon geocoding."""

import csv
import html
import json
import re
import time
import urllib.parse
import urllib.request
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
OUT = ROOT / "ВИЭ_Казахстан_координаты.xlsx"

TYPE_MAP = {
    "Солнечная электростанция": "solar",
    "Ветровая электростанция": "wind",
    "Гидроэлектростанция": "hydro",
    "Биоэлектростанция": "biogas",
}
TYPE_COLOR = {
    "Солнечная электростанция": "F9C74F",
    "Ветровая электростанция": "4D96FF",
    "Гидроэлектростанция": "2EC4B6",
    "Биоэлектростанция": "8AC926",
}
TYPE_ORDER = {
    "Ветровая электростанция": 1,
    "Солнечная электростанция": 2,
    "Гидроэлектростанция": 3,
    "Биоэлектростанция": 4,
}

# Region approximate centers only used as LAST resort labels, never as plant coords alone.
REGION_HINT = {
    "KZ-AKM": (51.92, 69.41),
    "KZ-AKT": (50.28, 57.17),
    "KZ-ALM": (43.87, 77.07),
    "KZ-ATY": (47.11, 51.92),
    "KZ-KAR": (49.80, 73.09),
    "KZ-KUS": (53.21, 63.62),
    "KZ-KZY": (44.85, 65.50),
    "KZ-MAN": (43.65, 51.20),
    "KZ-PAV": (52.29, 76.97),
    "KZ-SEV": (54.87, 69.15),
    "KZ-VOS": (49.95, 82.63),
    "KZ-YUZ": (43.30, 68.25),
    "KZ-ZAP": (51.23, 51.37),
    "KZ-ZHA": (42.90, 71.37),
    "KZ-10": (48.95, 80.40),
    "KZ-33": (45.00, 78.40),
    "KZ-62": (47.80, 67.71),
}

# Extra place aliases with verified OSM/Nominatim coords (not region centers).
PLACE_ALIASES = {
    "капшагай": (43.85517, 77.06146, "Конаев / Капшагай"),
    "капчагай": (43.85517, 77.06146, "Конаев / Капшагай"),
    "конаев": (43.85517, 77.06146, "Конаев"),
    "сарыбулак": (43.90610, 77.67549, "Сарыбұлақ (у Капшагая)"),  # Nominatim near Kapshagay
    "шенгельды": (43.70, 77.55, "Шенгельды (прибл. у Капшагая)"),  # approximate area near Kapshagay corridor
    "текели": (44.8330, 78.7640, "Текели"),
    "тараз": (42.8999, 71.3660, "Тараз"),
    "аральск": (46.8030, 61.6690, "Аральск"),
    "шаульдер": (42.7660, 68.3330, "Шаульдер"),
    "аксукент": (42.4160, 69.8330, "Аксукент"),
    "тургень": (43.4000, 77.6000, "Тургень"),
    "иссык": (43.3500, 77.4660, "Есик / Иссык"),
    "иссыкская": (43.3500, 77.4660, "Есик / Иссык"),
    "саумалколь": (53.2910, 68.1040, "Саумалколь (СКО)"),
    "форт-шевченко": (44.5170, 50.2620, "Форт-Шевченко"),
    "чарск": (49.52299, 81.22668, "Салкынтобе / район Чарска"),
    "шар": (49.52299, 81.22668, "Салкынтобе / район Чарска"),
    "актогай": (46.95220, 79.67543, "Актогай"),
    "жезказган": (47.80434, 67.71457, "Жезказган"),
    "байконур": (45.6160, 63.3160, "Байконур"),
    "уштобе": (45.2517, 77.9811, "Уштобе"),
    "жарма": (48.78231, 80.87832, "Жарма"),
    "бурабай": (53.08299, 70.30860, "Бурабай"),
    "какпак": (42.79919, 79.89968, "Какпак"),
    "турген": (43.4000, 77.6000, "Тургень"),
}

def load_settlements(path=None):
    path = Path(path) if path else DATA / "kz_settlements.json"
    try:
        raw = json.load(open(path, encoding="utf-8"))
    except Exception:
        return {}
    # map lower name -> list
    out = {}
    for k, v in raw.items():
        out.setdefault(k.lower(), []).append(v)
        out.setdefault(v["name"].lower(), []).append(v)
    # inject aliases
    for k, (lat, lon, name) in PLACE_ALIASES.items():
        out.setdefault(k, []).insert(0, {"lat": lat, "lon": lon, "name": name, "place": "alias", "osm": ""})
    return out

SETTLEMENTS = None  # lazy-loaded in build_rows

def nearest_settlement(name, region_code=None, region_name=None, settlements=None):
    if not name:
        return None
    settlements = settlements if settlements is not None else (SETTLEMENTS or {})
    key = name.lower().strip()
    # normalize
    aliases = {
        "капчагай": "капшагай",
        "г.капшагай": "капшагай",
        "есик": "иссык",
    }
    key = aliases.get(key, key)
    cands = list(settlements.get(key) or [])
    # also try without dashes
    if not cands and "-" in key:
        cands = list(settlements.get(key.replace("-", "")) or [])
    if not cands:
        # substring search
        for sk, arr in settlements.items():
            if key in sk or sk in key:
                cands.extend(arr)
    if not cands:
        return None
    # prefer by region hint distance
    hint = REGION_HINT.get(region_code) if region_code else None
    if hint:
        cands = sorted(cands, key=lambda c: (c["lat"] - hint[0]) ** 2 + (c["lon"] - hint[1]) ** 2)
    c = cands[0]
    return {"lat": float(c["lat"]), "lon": float(c["lon"]), "label": c.get("name") or name, "q": name}

# Explicit RFC title substring -> OSM/GPPD match keys (lat, lon, source, display, osm_url)
MANUAL = {
    "сэс нура": (50.7938388, 71.4529118, "OpenStreetMap", "«Нұра» күн электр станциясы", "https://www.openstreetmap.org/way/"),
    'kb enterprises': (50.7938388, 71.4529118, "OpenStreetMap", "СЭС Нура", ""),
    "ерейментау": (51.5898798, 73.0819106, "OpenStreetMap", "ВЭС Ерейментау", ""),
    "первая ветровая": (51.5898798, 73.0819106, "OpenStreetMap", "ВЭС Ерейментау", ""),
    "бадамша": (50.4880271, 58.2610052, "OpenStreetMap", "Badamsha 1 Wind Farm", ""),
    "ses saran": (49.8093123, 72.8719067, "OpenStreetMap / 2ГИС", "СЭС Сарань", "https://2gis.kz/karaganda/geo/70030076277188324"),
    "сэс сарань": (49.8093123, 72.8719067, "OpenStreetMap / 2ГИС", "СЭС Сарань", "https://2gis.kz/karaganda/geo/70030076277188324"),
    "бурное": (42.7142788, 70.821661, "OpenStreetMap / WRI", "СЭС Бурное", ""),
    "гульшат": (46.6377378, 74.323205, "OpenStreetMap / WRI", "СЭС Гульшат", ""),
    "агадырь": (48.2870707, 72.9092378, "OpenStreetMap", "СЭС Акадыр", ""),
    "жанатас": (43.4798303, 69.7133918, "OpenStreetMap", "Жанатасская ВЭС", ""),
    "кордай": (43.3350061, 74.9658276, "OpenStreetMap", "ВЭС Кордай", ""),
    "ветроинвест": (43.3350061, 74.9658276, "OpenStreetMap", "ВЭС Кордай", ""),
    "vista international": (43.3350061, 74.9658276, "OpenStreetMap", "ВЭС Кордай", ""),
    "eneverse kunkuat": (43.9677441, 77.1141341, "OpenStreetMap", "Қапшағай күн электр станциясы", ""),
    "тургусун": (49.9494722, 84.0499924, "OpenStreetMap", "Тургусунская ГЭС", ""),
    "кентау": (43.4721097, 68.5280927, "OpenStreetMap", "СЭС Кентау", ""),
    "m-kat green": (43.5890222, 73.705709, "OpenStreetMap", "Shu Solar Plant", ""),
    "форт-шевченко": (44.456464, 50.304649, "OpenStreetMap", "ВЭС Форт-Шевченко", ""),
    "редкометал": (44.456464, 50.304649, "OpenStreetMap", "ВЭС Форт-Шевченко", ""),
    "шокпарская ветровая": (43.513047, 69.7574286, "OpenStreetMap", "Шокпар жел электр станциясы", ""),
    "номад солар": (45.0972783, 64.6910113, "OpenStreetMap", "Nomad Solar Plant", ""),
    "жангиз солар": (49.2255853, 81.187849, "OpenStreetMap / WRI", "Zhalgyz Tobe Solar Plant", ""),
    "лепсы-2": (45.693203, 80.3067061, "OpenStreetMap", "ГЭС Лепсы-2", ""),
    "коринская гэс-2": (44.9142155, 78.8571097, "OpenStreetMap", 'ГЭС "Кора-3"', ""),
    "коринская гэс": (44.86981, 78.7945792, "OpenStreetMap", "Коринская ГЭС", ""),
    "верхне-баскан": (45.3530189, 80.1947866, "OpenStreetMap", "Verkhne Baskanskaya HPP-1", ""),
    "иссыкская гэс-3": (43.2776473, 77.4950589, "OpenStreetMap", "Issyk HPP-3", ""),
    "чижинская": (44.8652475, 78.8144334, "OpenStreetMap", "Chazhinskaya HPP-2", ""),
    "каскад каратал": (44.8724531, 78.737272, "OpenStreetMap", "Каратальская ГЭС", ""),
    "ыбырай": (53.2744194, 63.8310992, "OpenStreetMap", "Ybyrai wind farm", ""),
    "аршалын": (51.2714564, 71.8929675, "OpenStreetMap", "Arshalynsky wind farm", ""),
    "борей энерго": (51.2714564, 71.8929675, "OpenStreetMap", "Arshalynsky wind farm", ""),
    "нурлы": (43.6442452, 78.5371576, "OpenStreetMap", "Nurly wind farm", ""),
    "задарья": (42.4074722, 68.7599274, "OpenStreetMap", "Задарья-1", ""),
    "сайрам": (42.2231577, 70.2074446, "OpenStreetMap", "Sayramsu", ""),
    "новоникольское": (54.5259118, 68.6675039, "OpenStreetMap", "Petropavlovsk Wind Farm", ""),
    "зенченко": (54.5259118, 68.6675039, "OpenStreetMap", "Petropavlovsk Wind Farm", ""),
    "исатай": (46.9718375, 50.5603701, "OpenStreetMap", "ВЭС Тайман", ""),
    "ветроэнерготехнолог": (46.9718375, 50.5603701, "OpenStreetMap", "ВЭС Тайман", ""),
}


def norm(s: str) -> str:
    s = html.unescape(s or "").lower()
    s = s.replace("«", '"').replace("»", '"').replace("“", '"').replace("”", '"')
    s = re.sub(r"[^\w\s\-а-яёәіңғүұқөһa-z0-9]", " ", s, flags=re.I)
    return re.sub(r"\s+", " ", s).strip()


def tokens(s: str):
    stop = {
        "тоо", "ао", "в", "на", "г", "п", "с", "р", "реке", "области", "обл", "район",
        "районе", "мвт", "бывш", "бывшие", "возле", "поселке", "селе", "город", "города",
        "компания", "llp", "llc", "jsc", "and", "the", "energy", "green", "электростанция",
        "ветровая", "солнечная", "гидроэлектростанция", "корп", "corp",
    }
    return [
        t
        for t in re.findall(r"[a-zа-яёәіңғүұқөһ0-9\-]{3,}", norm(s), flags=re.I)
        if t not in stop
    ]


def parse_cap(s):
    try:
        return float(str(s).replace(",", "."))
    except Exception:
        return None


PLACE_RE = re.compile(
    r"(?:в|возле|близ|около)\s+(?:г\.|п\.|с\.|пос\.|селе|поселке|городе)?\s*"
    r"([A-Za-zА-Яа-яЁёӘәІіҢңҒғҮүҰұҚқӨөҺһ][A-Za-zА-Яа-яЁёӘәІіҢңҒғҮүҰұҚқӨөҺһ\-\s]{2,35})|"
    r"(?:г\.|п\.|с\.|пос\.)\s*"
    r"([A-Za-zА-Яа-яЁёӘәІіҢңҒғҮүҰұҚқӨөҺһ][A-Za-zА-Яа-яЁёӘәІіҢңҒғҮүҰұҚқӨөҺһ\-]{2,30})",
    re.I,
)


def extract_places(title: str):
    out = []
    for m in PLACE_RE.finditer(title):
        g = next(x for x in m.groups() if x)
        g = re.split(r"\s+(?:в|на|район|области|,|\()", g)[0].strip(" .,;")
        if len(g) >= 3 and g.lower() not in {x.lower() for x in out}:
            out.append(g)
    for pat in [r"СЭС\s+([^\(,]+)", r"ВЭС\s+([^\(,]+)", r"ГЭС\s+([^\(,]+)"]:
        m = re.search(pat, title, re.I)
        if m:
            g = m.group(1).strip(" .,;")
            g = re.split(r"\s{2,}|\s+в\s+", g)[0].strip()
            if len(g) >= 3 and g.lower() not in {x.lower() for x in out}:
                out.append(g)
    return out[:3]


def load_osm(path=None):
    path = Path(path) if path else DATA / "osm_kz_res.json"
    raw = json.load(open(path, encoding="utf-8"))
    plants = []
    for e in raw["elements"]:
        t = e.get("tags", {})
        lat = e.get("lat") or (e.get("center") or {}).get("lat")
        lon = e.get("lon") or (e.get("center") or {}).get("lon")
        if lat is None:
            continue
        src = (t.get("plant:source") or "").split(";")[0].strip().lower()
        if src == "biomass":
            src = "biogas"
        cap = t.get("plant:output:electricity") or ""
        m = re.search(r"([\d.,]+)\s*(MW|МВт|kW|кВт|MWp)?", cap, re.I)
        cap_mw = None
        if m:
            val = float(m.group(1).replace(",", "."))
            unit = (m.group(2) or "MW").lower()
            if unit in ("kw", "квт"):
                val /= 1000.0
            cap_mw = val
        names = " ".join(
            filter(
                None,
                [
                    t.get("name"),
                    t.get("name:ru"),
                    t.get("name:en"),
                    t.get("name:kk"),
                    t.get("operator"),
                    t.get("short_name"),
                ],
            )
        )
        plants.append(
            {
                "id": f"{e['type']}/{e['id']}",
                "names": names,
                "toks": set(tokens(names)),
                "src": src,
                "cap": cap_mw,
                "lat": float(lat),
                "lon": float(lon),
                "display": t.get("name")
                or t.get("name:ru")
                or t.get("name:en")
                or t.get("name:kk")
                or f"OSM {e['type']}/{e['id']}",
                "origin": "OpenStreetMap",
                "osm_url": f"https://www.openstreetmap.org/{e['type']}/{e['id']}",
            }
        )
    return plants


def load_gppd(path=None):
    path = Path(path) if path else DATA / "gppd_kaz_res.csv"
    out = []
    with open(path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if row.get("country") != "KAZ":
                continue
            fuel = (row.get("primary_fuel") or "").lower()
            if fuel not in ("solar", "wind", "hydro", "biomass", "biogas"):
                continue
            cap = float(row["capacity_mw"]) if row.get("capacity_mw") else None
            if fuel == "hydro" and cap and cap > 100:
                continue
            out.append(
                {
                    "id": row["gppd_idnr"],
                    "names": row.get("name") or "",
                    "toks": set(tokens(row.get("name") or "")),
                    "src": "biogas" if fuel == "biomass" else fuel,
                    "cap": cap,
                    "lat": float(row["latitude"]),
                    "lon": float(row["longitude"]),
                    "display": row.get("name") or "",
                    "origin": "WRI Global Power Plant Database",
                    "osm_url": row.get("url") or "",
                }
            )
    return out


def score_match(item, plant):
    want = TYPE_MAP.get(item["type"])
    if want and plant["src"] and want != plant["src"]:
        return -1
    itoks = set(tokens(item["title"] + " " + item.get("name_raw", "")))
    overlap = len(itoks & plant["toks"])
    icap = parse_cap(item["power_mw"])
    cap_score = 0
    if icap and plant["cap"]:
        diff = abs(icap - plant["cap"]) / max(icap, plant["cap"])
        if diff <= 0.05:
            cap_score = 4
        elif diff <= 0.15:
            cap_score = 2
        elif diff <= 0.35:
            cap_score = 1
        elif diff > 0.55:
            return -1
    bonus = 0
    nt = norm(item["title"])
    pn = norm(plant["names"])
    for tok in [
        "ерейментау", "бурное", "сарань", "saran", "капшагай", "капчагай", "нура",
        "жанатас", "кордай", "qorday", "бадамша", "badamsha", "гульшат", "агадырь",
        "акадыр", "шокпар", "шелек", "тургусун", "коринск", "лепсы", "кентау",
        "сарыбулак", "нурлы", "шенгельды", "форт", "zhalgyz", "nomad", "ybyrai",
        "ыбырай", "тайман", "аршалын", "arshalyn", "шу", "shu solar", "задарья",
        "баскан", "иссык", "каратал", "сайрам",
    ]:
        if tok in nt and tok in pn:
            bonus += 5
    return overlap * 2 + cap_score + bonus


geo_cache = {}


def photon(query, region=None):
    key = f"{query}|{region}"
    if key in geo_cache:
        return geo_cache[key]
    q = f"{query}, {region}, Kazakhstan" if region else f"{query}, Kazakhstan"
    # keep query ASCII-friendly length
    q = q[:120]
    url = "https://photon.komoot.io/api/?" + urllib.parse.urlencode({"q": q, "limit": 3})
    req = urllib.request.Request(url, headers={"User-Agent": "vie-rfc-export/1.0"})
    res = None
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            data = json.load(r)
        for feat in data.get("features") or []:
            props = feat.get("properties") or {}
            cc = (props.get("countrycode") or "").upper()
            country = (props.get("country") or "").lower()
            lon, lat = feat["geometry"]["coordinates"]
            # accept KZ by countrycode, name variants, or Kazakhstan bbox
            in_kz_bbox = 40.5 <= float(lat) <= 55.5 and 46.4 <= float(lon) <= 87.5
            ok = cc == "KZ" or "казах" in country or "qazaq" in country or "kazakh" in country or in_kz_bbox
            if not ok:
                continue
            label = ", ".join(
                filter(
                    None,
                    [
                        props.get("name"),
                        props.get("city"),
                        props.get("county"),
                        props.get("state"),
                        props.get("country"),
                    ],
                )
            )
            res = {"lat": float(lat), "lon": float(lon), "label": label or q, "q": q}
            break
    except Exception as e:
        print("photon err", q, e, flush=True)
        res = None
    geo_cache[key] = res
    time.sleep(0.12)
    return res


def manual_lookup(title: str):
    nt = norm(title)
    # fix typo key
    for k, v in MANUAL.items():
        if v is None:
            continue
        if k in nt:
            return v
    if "первая ветровая" in nt or "первaя ветровая" in nt:
        return MANUAL["ерейментау"]
    return None


def build_rows(rfc, candidates):
    global SETTLEMENTS
    SETTLEMENTS = load_settlements()
    used = set()
    results = []
    for item in rfc:
        row = {
            **item,
            "lat": None,
            "lon": None,
            "coord_source": "",
            "matched_name": "",
            "match_score": 0,
            "confidence": "",
            "place_hint": "",
            "maps_url": "",
            "dgis_url": "",
            "osm_url": "",
        }

        # 1) manual known plants (from OSM/GPPD verified coords)
        man = manual_lookup(item["title"])
        if man:
            lat, lon, src, name, link = man
            row.update(
                lat=lat,
                lon=lon,
                coord_source=src,
                matched_name=name,
                confidence="высокая",
                match_score=20,
                osm_url=link or "",
            )
        else:
            # 2) fuzzy OSM/GPPD
            best = None
            best_s = -1
            for p in candidates:
                if p["id"] in used:
                    continue
                s = score_match(item, p)
                if s > best_s:
                    best_s = s
                    best = p
            if best and best_s >= 6:
                row.update(
                    lat=best["lat"],
                    lon=best["lon"],
                    coord_source=best["origin"],
                    matched_name=best["display"],
                    match_score=best_s,
                    confidence="высокая" if best_s >= 9 else "средняя",
                    osm_url=best.get("osm_url", ""),
                )
                used.add(best["id"])
            else:
                places = extract_places(item["title"])
                row["place_hint"] = ", ".join(places)
                found = None
                queries = []
                for p in places:
                    queries.append(p)
                for pat in [
                    r"в\s+(?:г\.|п\.|с\.|пос\.)\s*([^\s,\(]{3,30})",
                    r"селе\s+([^\s,\(]{3,30})",
                    r"поселке\s+([^\s,\(]{3,30})",
                    r"(?:СЭС|ВЭС|ГЭС)\s+([^\s,\(]{3,30})",
                ]:
                    m2 = re.search(pat, item["title"], re.I)
                    if m2:
                        queries.append(m2.group(1).strip(" ."))
                # known keyword places in title — only if look like location context
                nt = norm(item["title"])
                for alias in PLACE_ALIASES:
                    if alias in nt and re.search(
                        rf"(?:г\.|п\.|с\.|пос\.|селе|поселке|городе|возле|близ|реке|р\.)\s*[^\n]{{0,25}}{re.escape(alias)}|"
                        rf"(?:в|на)\s+[^\n]{{0,15}}{re.escape(alias)}|"
                        rf"(?:сэс|вэс|гэс)\s+{re.escape(alias)}|"
                        rf"{re.escape(alias)}\s*(?:сэс|вэс|гэс)",
                        nt,
                        re.I,
                    ):
                        queries.append(alias)
                # districts
                for dist, place in [
                    ("жарминск", "жарма"),
                    ("тупкаран", "форт-шевченко"),
                    ("каракиян", "форт-шевченко"),
                    ("отырар", "шаульдер"),
                    ("бурабай", "бурабай"),
                    ("burabai", "бурабай"),
                ]:
                    if dist in nt:
                        queries.append(place)
                seen = set()
                uniq = []
                for q in queries:
                    ql = q.lower().strip(" .")
                    if not ql or ql in seen or len(ql) < 3:
                        continue
                    if ql in {"район", "области", "мощностью", "тоо", "жарминском", "тупкаранском", "каракиянском", "отырарском"}:
                        continue
                    seen.add(ql)
                    uniq.append(q.strip(" ."))
                for q in uniq:
                    found = nearest_settlement(
                        q, item.get("region_code"), item.get("region"), SETTLEMENTS
                    )
                    if found:
                        break
                if found:
                    row.update(
                        lat=found["lat"],
                        lon=found["lon"],
                        coord_source=f"Населённый пункт (OSM): {found['q']}",
                        matched_name=found["label"],
                        confidence="приблизительно",
                    )

        if row["lat"] is not None:
            row["maps_url"] = f"https://www.google.com/maps?q={row['lat']},{row['lon']}"
            row["dgis_url"] = f"https://2gis.kz/geo/{row['lon']}%2C{row['lat']}"
        results.append(row)
    return results


def write_excel(results):
    wb = Workbook()
    ws = wb.active
    ws.title = "Объекты ВИЭ"
    headers = [
        "№",
        "Название объекта",
        "Тип ВИЭ",
        "Мощность, МВт",
        "Область",
        "Широта",
        "Долгота",
        "Координаты (lat, lon)",
        "Точность",
        "Источник координат",
        "Сопоставлено с",
        "Подсказка места",
        "Google Maps",
        "2ГИС",
        "OpenStreetMap / ссылка",
        "ID РФЦ",
    ]
    ws.append(headers)

    header_fill = PatternFill("solid", "1B4332")
    header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    thin = Border(
        left=Side(style="thin", color="D8E2DC"),
        right=Side(style="thin", color="D8E2DC"),
        top=Side(style="thin", color="D8E2DC"),
        bottom=Side(style="thin", color="D8E2DC"),
    )
    wrap = Alignment(wrap_text=True, vertical="center")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for col, h in enumerate(headers, 1):
        cell = ws.cell(1, col, h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    results_sorted = sorted(
        results, key=lambda r: (TYPE_ORDER.get(r["type"], 9), r["region"], r["title"])
    )

    for i, r in enumerate(results_sorted, 1):
        coords = f"{r['lat']:.6f}, {r['lon']:.6f}" if r["lat"] is not None else ""
        ws.append(
            [
                i,
                r["title"],
                r["type"],
                parse_cap(r["power_mw"]),
                r["region"],
                round(r["lat"], 6) if r["lat"] is not None else "",
                round(r["lon"], 6) if r["lon"] is not None else "",
                coords,
                r["confidence"],
                r["coord_source"],
                r["matched_name"],
                r["place_hint"],
                "",
                "",
                "",
                r["id"],
            ]
        )
        type_cell = ws.cell(i + 1, 3)
        type_cell.fill = PatternFill("solid", TYPE_COLOR.get(r["type"], "FFFFFF"))
        for col in range(1, 17):
            c = ws.cell(i + 1, col)
            c.border = thin
            c.alignment = wrap if col in (2, 10, 11, 12) else center
            c.font = Font(name="Calibri", size=10)
        if r["maps_url"]:
            cell = ws.cell(i + 1, 13, "Открыть")
            cell.hyperlink = r["maps_url"]
            cell.font = Font(name="Calibri", color="0563C1", underline="single", size=10)
        if r["dgis_url"]:
            cell = ws.cell(i + 1, 14, "Открыть")
            cell.hyperlink = r["dgis_url"]
            cell.font = Font(name="Calibri", color="0563C1", underline="single", size=10)
        if r["osm_url"]:
            cell = ws.cell(i + 1, 15, "Ссылка")
            cell.hyperlink = r["osm_url"]
            cell.font = Font(name="Calibri", color="0563C1", underline="single", size=10)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:P{len(results_sorted)+1}"
    ws.row_dimensions[1].height = 32
    widths = [5, 48, 24, 12, 26, 12, 12, 22, 16, 40, 28, 22, 12, 10, 14, 12]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Summary
    ws2 = wb.create_sheet("Сводка", 0)
    ws2["A1"] = "Объекты ВИЭ Казахстана — координаты"
    ws2["A1"].font = Font(name="Calibri", bold=True, size=16, color="1B4332")
    ws2["A2"] = "Список: https://rfc.kz/ru/res-sector/map/  (API /api/locality/?lang=ru)"
    ws2["A3"] = (
        "Координаты: OpenStreetMap / OpenInfraMap, WRI GPPD, геокодинг Photon (OSM), ссылки 2ГИС"
    )
    ws2["A5"] = "Всего объектов"
    ws2["B5"] = len(results_sorted)
    ws2["A6"] = "С координатами"
    ws2["B6"] = sum(1 for r in results_sorted if r["lat"] is not None)
    ws2["A7"] = "Без координат"
    ws2["B7"] = sum(1 for r in results_sorted if r["lat"] is None)
    ws2["A8"] = "Суммарная мощность, МВт"
    ws2["B8"] = round(sum(parse_cap(r["power_mw"]) or 0 for r in results_sorted), 3)
    ws2["A9"] = "Точность: высокая"
    ws2["B9"] = sum(1 for r in results_sorted if r["confidence"] == "высокая")
    ws2["A10"] = "Точность: средняя"
    ws2["B10"] = sum(1 for r in results_sorted if r["confidence"] == "средняя")
    ws2["A11"] = "Точность: приблизительно"
    ws2["B11"] = sum(1 for r in results_sorted if r["confidence"] == "приблизительно")

    ws2["A13"] = "По типу"
    ws2["A13"].font = Font(bold=True)
    for col, h in enumerate(["Тип", "Кол-во", "МВт", "С координатами"], 1):
        cell = ws2.cell(14, col, h)
        cell.fill = header_fill
        cell.font = header_font

    agg = defaultdict(lambda: [0, 0.0, 0])
    for r in results_sorted:
        a = agg[r["type"]]
        a[0] += 1
        a[1] += parse_cap(r["power_mw"]) or 0
        a[2] += 1 if r["lat"] is not None else 0
    rowi = 15
    for t, (n, mw, c) in sorted(agg.items(), key=lambda x: -x[1][1]):
        ws2.cell(rowi, 1, t).fill = PatternFill("solid", TYPE_COLOR.get(t, "FFFFFF"))
        ws2.cell(rowi, 2, n)
        ws2.cell(rowi, 3, round(mw, 3))
        ws2.cell(rowi, 4, c)
        rowi += 1

    rowi += 1
    ws2.cell(rowi, 1, "По области").font = Font(bold=True)
    rowi += 1
    for col, h in enumerate(["Область", "Кол-во", "МВт", "С координатами"], 1):
        cell = ws2.cell(rowi, col, h)
        cell.fill = header_fill
        cell.font = header_font
    rowi += 1
    agg2 = defaultdict(lambda: [0, 0.0, 0])
    for r in results_sorted:
        a = agg2[r["region"]]
        a[0] += 1
        a[1] += parse_cap(r["power_mw"]) or 0
        a[2] += 1 if r["lat"] is not None else 0
    for reg, (n, mw, c) in sorted(agg2.items(), key=lambda x: -x[1][1]):
        ws2.cell(rowi, 1, reg)
        ws2.cell(rowi, 2, n)
        ws2.cell(rowi, 3, round(mw, 3))
        ws2.cell(rowi, 4, c)
        rowi += 1

    ws2.column_dimensions["A"].width = 36
    ws2.column_dimensions["B"].width = 14
    ws2.column_dimensions["C"].width = 14
    ws2.column_dimensions["D"].width = 16

    ws3 = wb.create_sheet("Легенда")
    ws3["A1"] = "Как читать файл"
    ws3["A1"].font = Font(bold=True, size=14, color="1B4332")
    notes = [
        "1. Список объектов — с официальной карты РФЦ (rfc.kz), 129 объектов.",
        "2. На карте РФЦ координат отдельных станций нет (только центры областей) — поэтому координаты собраны извне.",
        "3. Источники координат: OpenStreetMap / OpenInfraMap, WRI Global Power Plant Database, геокодинг Photon (данные OSM), ссылки 2ГИС.",
        "4. «высокая» — координаты самой станции (контур/точка объекта).",
        "5. «средняя» — уверенное сопоставление по названию/мощности с OSM.",
        "6. «приблизительно» — координаты населённого пункта из названия объекта (не точная точка станции).",
        "7. Колонки Google Maps и 2ГИС — кликабельные ссылки на точку.",
        "8. Цвета типа: синий ВЭС, жёлтый СЭС, бирюзовый ГЭС, зелёный БиоЭС.",
        "9. Дата выгрузки: 2026-09-21. Крупные объекты лучше сверять на https://openinframap.org и 2gis.kz.",
    ]
    for i, n in enumerate(notes, 3):
        ws3.cell(i, 1, n)
    ws3.column_dimensions["A"].width = 130

    wb.save(OUT)
    return results_sorted


def main():
    rfc = json.load(open(DATA / "vie_list.json", encoding="utf-8"))
    plants = load_osm()
    gppd = load_gppd()
    candidates = plants + gppd
    print(f"RFC={len(rfc)} OSM={len(plants)} GPPD={len(gppd)}", flush=True)
    results = build_rows(rfc, candidates)
    print("confidence", Counter(r["confidence"] or "нет" for r in results), flush=True)
    print("with coords", sum(1 for r in results if r["lat"] is not None), "/", len(results), flush=True)
    sorted_rows = write_excel(results)
    json.dump(
        sorted_rows,
        open(DATA / "vie_with_coords.json", "w", encoding="utf-8"),
        ensure_ascii=False,
        indent=2,
    )
    missing = [r["title"] for r in sorted_rows if r["lat"] is None]
    print("SAVED", OUT, flush=True)
    print("missing", len(missing), missing[:25], flush=True)


if __name__ == "__main__":
    main()
