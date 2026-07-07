"""Generate ``data/canon.json``: the canonical 66-book table.

The canonical **book id** (OSIS-style, e.g. ``Gen``) is the language-independent
key used everywhere in the app. Per-language display names live in
``data/i18n/*.json``; this file only pins the ids, ordering, testament and a
couple of built-in (ko/en) names/abbreviations so the app can bootstrap even if
an i18n file is missing.

Run:  python scripts/build_canon.py
"""
from __future__ import annotations

import json
from pathlib import Path

# (osis_id, testament, ko_name, ko_abbr, en_name, en_abbr)
# ``ko_abbr`` matches the single/short keys used by the legacy data files so the
# importer/migration can line data up without guessing.
BOOKS: list[tuple[str, str, str, str, str, str]] = [
    ("Gen", "OT", "창세기", "창", "Genesis", "Gen"),
    ("Exod", "OT", "출애굽기", "출", "Exodus", "Exod"),
    ("Lev", "OT", "레위기", "레", "Leviticus", "Lev"),
    ("Num", "OT", "민수기", "민", "Numbers", "Num"),
    ("Deut", "OT", "신명기", "신", "Deuteronomy", "Deut"),
    ("Josh", "OT", "여호수아", "수", "Joshua", "Josh"),
    ("Judg", "OT", "사사기", "삿", "Judges", "Judg"),
    ("Ruth", "OT", "룻기", "룻", "Ruth", "Ruth"),
    ("1Sam", "OT", "사무엘상", "삼상", "1 Samuel", "1Sam"),
    ("2Sam", "OT", "사무엘하", "삼하", "2 Samuel", "2Sam"),
    ("1Kgs", "OT", "열왕기상", "왕상", "1 Kings", "1Kgs"),
    ("2Kgs", "OT", "열왕기하", "왕하", "2 Kings", "2Kgs"),
    ("1Chr", "OT", "역대상", "대상", "1 Chronicles", "1Chr"),
    ("2Chr", "OT", "역대하", "대하", "2 Chronicles", "2Chr"),
    ("Ezra", "OT", "에스라", "스", "Ezra", "Ezra"),
    ("Neh", "OT", "느헤미야", "느", "Nehemiah", "Neh"),
    ("Esth", "OT", "에스더", "에", "Esther", "Esth"),
    ("Job", "OT", "욥기", "욥", "Job", "Job"),
    ("Ps", "OT", "시편", "시", "Psalms", "Ps"),
    ("Prov", "OT", "잠언", "잠", "Proverbs", "Prov"),
    ("Eccl", "OT", "전도서", "전", "Ecclesiastes", "Eccl"),
    ("Song", "OT", "아가", "아", "Song of Songs", "Song"),
    ("Isa", "OT", "이사야", "사", "Isaiah", "Isa"),
    ("Jer", "OT", "예레미야", "렘", "Jeremiah", "Jer"),
    ("Lam", "OT", "예레미야애가", "애", "Lamentations", "Lam"),
    ("Ezek", "OT", "에스겔", "겔", "Ezekiel", "Ezek"),
    ("Dan", "OT", "다니엘", "단", "Daniel", "Dan"),
    ("Hos", "OT", "호세아", "호", "Hosea", "Hos"),
    ("Joel", "OT", "요엘", "욜", "Joel", "Joel"),
    ("Amos", "OT", "아모스", "암", "Amos", "Amos"),
    ("Obad", "OT", "오바댜", "옵", "Obadiah", "Obad"),
    ("Jonah", "OT", "요나", "욘", "Jonah", "Jonah"),
    ("Mic", "OT", "미가", "미", "Micah", "Mic"),
    ("Nah", "OT", "나훔", "나", "Nahum", "Nah"),
    ("Hab", "OT", "하박국", "합", "Habakkuk", "Hab"),
    ("Zeph", "OT", "스바냐", "습", "Zephaniah", "Zeph"),
    ("Hag", "OT", "학개", "학", "Haggai", "Hag"),
    ("Zech", "OT", "스가랴", "슥", "Zechariah", "Zech"),
    ("Mal", "OT", "말라기", "말", "Malachi", "Mal"),
    ("Matt", "NT", "마태복음", "마", "Matthew", "Matt"),
    ("Mark", "NT", "마가복음", "막", "Mark", "Mark"),
    ("Luke", "NT", "누가복음", "눅", "Luke", "Luke"),
    ("John", "NT", "요한복음", "요", "John", "John"),
    ("Acts", "NT", "사도행전", "행", "Acts", "Acts"),
    ("Rom", "NT", "로마서", "롬", "Romans", "Rom"),
    ("1Cor", "NT", "고린도전서", "고전", "1 Corinthians", "1Cor"),
    ("2Cor", "NT", "고린도후서", "고후", "2 Corinthians", "2Cor"),
    ("Gal", "NT", "갈라디아서", "갈", "Galatians", "Gal"),
    ("Eph", "NT", "에베소서", "엡", "Ephesians", "Eph"),
    ("Phil", "NT", "빌립보서", "빌", "Philippians", "Phil"),
    ("Col", "NT", "골로새서", "골", "Colossians", "Col"),
    ("1Thess", "NT", "데살로니가전서", "살전", "1 Thessalonians", "1Thess"),
    ("2Thess", "NT", "데살로니가후서", "살후", "2 Thessalonians", "2Thess"),
    ("1Tim", "NT", "디모데전서", "딤전", "1 Timothy", "1Tim"),
    ("2Tim", "NT", "디모데후서", "딤후", "2 Timothy", "2Tim"),
    ("Titus", "NT", "디도서", "딛", "Titus", "Titus"),
    ("Phlm", "NT", "빌레몬서", "몬", "Philemon", "Phlm"),
    ("Heb", "NT", "히브리서", "히", "Hebrews", "Heb"),
    ("Jas", "NT", "야고보서", "약", "James", "Jas"),
    ("1Pet", "NT", "베드로전서", "벧전", "1 Peter", "1Pet"),
    ("2Pet", "NT", "베드로후서", "벧후", "2 Peter", "2Pet"),
    ("1John", "NT", "요한일서", "요일", "1 John", "1John"),
    ("2John", "NT", "요한이서", "요이", "2 John", "2John"),
    ("3John", "NT", "요한삼서", "요삼", "3 John", "3John"),
    ("Jude", "NT", "유다서", "유", "Jude", "Jude"),
    ("Rev", "NT", "요한계시록", "계", "Revelation", "Rev"),
]


def build() -> list[dict]:
    return [
        {
            "id": osis,
            "order": i + 1,
            "testament": testament,
            "ko": ko,
            "ko_abbr": ko_abbr,
            "en": en,
            "en_abbr": en_abbr,
        }
        for i, (osis, testament, ko, ko_abbr, en, en_abbr) in enumerate(BOOKS)
    ]


def main() -> None:
    out = Path(__file__).resolve().parent.parent / "data" / "canon.json"
    out.write_text(
        json.dumps(build(), ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(f"wrote {out} ({len(BOOKS)} books)")


if __name__ == "__main__":
    main()
