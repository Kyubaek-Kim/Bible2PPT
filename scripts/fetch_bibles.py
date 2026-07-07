"""Reproducibly download free/public-domain Bible translations.

Each translation is fetched from a public source and written to
``data/bibles/<CODE>.json`` in the app's canonical schema::

    {
      "meta": { "code", "name", "language", "abbr", "versification",
                "source", "license" },
      "books": { "<BookId>": { "<chapter>": { "<verse>": "<text>" } } }
    }

**Verse text is stored exactly as received** — no normalisation, trimming or
reordering — per the project's text-integrity rule. Book numbers from the source
(1..66) map to canonical book ids via ``data/canon.json`` order.

Primary source: getbible.net v2 API (aggregates public-domain texts). Copyright
editions (e.g. 개역개정) are intentionally excluded.

Usage:
    python scripts/fetch_bibles.py                # all configured translations
    python scripts/fetch_bibles.py KRV KJV        # a subset (by our code)
    python scripts/fetch_bibles.py --versification # (re)build data/versification/kjv.json
"""
from __future__ import annotations

import json
import sys
import time
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
BIBLES = DATA / "bibles"
VERSIFICATION = DATA / "versification"

GETBIBLE = "https://api.getbible.net/v2/{trans}/{book}.json"

# our code -> source config. ``license`` records the distribution status of the
# underlying text as published by the source; verify before redistribution.
TRANSLATIONS: list[dict] = [
    {"code": "KRV", "getbible": "korean", "name": "개역한글", "language": "ko",
     "abbr": "개역", "versification": "kjv",
     "license": "Public domain (Korean Revised Version, 1961); via getbible.net"},
    {"code": "KJV", "getbible": "kjv", "name": "King James Version", "language": "en",
     "abbr": "KJV", "versification": "kjv", "license": "Public domain"},
    {"code": "ASV", "getbible": "asv", "name": "American Standard Version", "language": "en",
     "abbr": "ASV", "versification": "kjv", "license": "Public domain"},
    {"code": "WEB", "getbible": "web", "name": "World English Bible", "language": "en",
     "abbr": "WEB", "versification": "kjv", "license": "Public domain"},
    {"code": "YLT", "getbible": "ylt", "name": "Young's Literal Translation", "language": "en",
     "abbr": "YLT", "versification": "kjv", "license": "Public domain"},
    {"code": "TR", "getbible": "textusreceptus", "name": "Textus Receptus (NT)",
     "language": "grc", "abbr": "TR", "versification": "kjv",
     "license": "Public domain (Scrivener 1894 / Stephanus)"},
    {"code": "WH", "getbible": "westcotthort", "name": "Westcott-Hort (NT)",
     "language": "grc", "abbr": "WH", "versification": "kjv",
     "license": "Public domain"},
    {"code": "LXX", "getbible": "lxx", "name": "Septuagint (LXX, OT)", "language": "grc",
     "abbr": "LXX", "versification": "lxx", "license": "Public domain"},
    {"code": "WLC", "getbible": "codex", "name": "Westminster Leningrad Codex (OT)",
     "language": "hbo", "abbr": "WLC", "versification": "mt",
     "license": "Public domain / CC (Westminster Leningrad Codex)"},
    {"code": "ALEPPO", "getbible": "aleppo", "name": "Aleppo Codex (OT)", "language": "hbo",
     "abbr": "Aleppo", "versification": "mt", "license": "Public domain"},
    {"code": "VULGATE", "getbible": "vulgate", "name": "Clementine Vulgate", "language": "la",
     "abbr": "Vulg", "versification": "vulgate", "license": "Public domain"},
]

# Requested but not on getbible.net; add a dedicated adapter to enable.
PENDING = {
    "SBLGNT": "https://github.com/LogosBible/SBLGNT (CC BY 4.0) — needs adapter",
    "Nestle1904": "https://github.com/biblicalhumanities/Nestle1904 — needs adapter",
}


def _canon_order() -> dict[int, str]:
    canon = json.loads((DATA / "canon.json").read_text(encoding="utf-8"))
    return {entry["order"]: entry["id"] for entry in canon}


def _fetch_getbible_book(trans: str, book_nr: int, session: requests.Session):
    url = GETBIBLE.format(trans=trans, book=book_nr)
    resp = session.get(url, timeout=60)
    if resp.status_code == 404:
        return None
    resp.raise_for_status()
    return resp.json()


def fetch_translation(cfg: dict, order_to_id: dict[int, str], session: requests.Session) -> Path:
    books: dict[str, dict[str, dict[str, str]]] = {}
    present = 0
    for nr in range(1, 67):
        data = _fetch_getbible_book(cfg["getbible"], nr, session)
        if not data or "chapters" not in data:
            continue
        book_id = order_to_id[nr]
        book_out: dict[str, dict[str, str]] = {}
        for chap in data["chapters"]:
            cnum = str(chap["chapter"])
            verses = {str(v["verse"]): v["text"] for v in chap.get("verses", [])}
            if verses:
                book_out[cnum] = verses
        if book_out:
            books[book_id] = book_out
            present += 1
        time.sleep(0.05)

    payload = {
        "meta": {
            "code": cfg["code"],
            "name": cfg["name"],
            "language": cfg["language"],
            "abbr": cfg["abbr"],
            "versification": cfg["versification"],
            "source": GETBIBLE.format(trans=cfg["getbible"], book="{1-66}"),
            "license": cfg["license"],
        },
        "books": books,
    }
    BIBLES.mkdir(parents=True, exist_ok=True)
    out = BIBLES / f"{cfg['code']}.json"
    out.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    print(f"  {cfg['code']:8s} {present:2d} books -> {out.name}")
    return out


def build_versification_from_kjv() -> None:
    """Derive the canonical last-verse table from the fetched KJV translation."""
    kjv_path = BIBLES / "KJV.json"
    if not kjv_path.exists():
        print("KJV.json not found; fetch KJV first.")
        return
    data = json.loads(kjv_path.read_text(encoding="utf-8"))
    table: dict[str, dict[str, int]] = {}
    for book_id, chapters in data["books"].items():
        table[book_id] = {c: max(int(v) for v in verses) for c, verses in chapters.items()}
    VERSIFICATION.mkdir(parents=True, exist_ok=True)
    out = VERSIFICATION / "kjv.json"
    out.write_text(json.dumps(table, ensure_ascii=False, indent=0), encoding="utf-8")
    print(f"wrote {out}")


def main(argv: list[str]) -> None:
    if "--versification" in argv:
        build_versification_from_kjv()
        return
    wanted = [a for a in argv if not a.startswith("-")]
    order_to_id = _canon_order()
    session = requests.Session()
    session.headers.update({"User-Agent": "Bible2PPT/1.0 fetch_bibles"})
    print("Fetching translations...")
    for cfg in TRANSLATIONS:
        if wanted and cfg["code"] not in wanted:
            continue
        try:
            fetch_translation(cfg, order_to_id, session)
        except Exception as exc:  # noqa: BLE001
            print(f"  {cfg['code']}: FAILED ({exc})")
    build_versification_from_kjv()
    if PENDING:
        print("\nPending (need a dedicated adapter):")
        for code, note in PENDING.items():
            print(f"  {code}: {note}")


if __name__ == "__main__":
    main(sys.argv[1:])
