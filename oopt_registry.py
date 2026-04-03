"""
Load ООПТ territory lists from Word .doc files (OOPT_territories/) and match
(лесничество, участковое лесничество, квартал) to a short label for tables.

On macOS, plain text is extracted with ``textutil``. Other platforms can set
OOPT_DOC_TEXT_CMD or pre-convert .doc to .txt.
"""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import unicodedata
from dataclasses import dataclass
from functools import lru_cache
from typing import Iterable, Tuple


@dataclass(frozen=True)
class OOPTRule:
    """One ООПТ footprint: участковое лесничество + set of кварталов in a big forest zone."""

    big_zone: str
    uchastkovoe: str
    kvartals: frozenset[int]
    label: str
    title: str


def _normalize_dashes(s: str) -> str:
    return (
        s.replace("\u2013", "-")
        .replace("\u2014", "-")
        .replace("\u2212", "-")
    )


def normalize_zone_name(name: str | None) -> str:
    if name is None or (isinstance(name, float) and str(name) == "nan"):
        return ""
    # Standard cleanup
    s = str(name).strip().lower().replace("ё", "е")
    s = unicodedata.normalize("NFKC", s)
    s = re.sub(r"\s+", " ", s)
    # Remove common case endings for Russian adjectives in forest/district names
    # e.g. "Приозерское", "Приозерского", "Приозерскoe" -> "приозерск"
    s = re.sub(r"(ское|ского|екое|кого|ое|ая|ий|ый|ов|ка)$", "", s)
    return s


def forest_zone_from_filename(basename: str) -> str:
    """``Площади_Приозерское_ООПТ.doc`` -> ``Приозерское``."""
    stem = os.path.splitext(basename)[0]
    stem = re.sub(r"_[a-zA-Z0-9]{6,}$", "", stem)
    m = re.match(r"Площади_(.+)_ООПТ", stem)
    if not m:
        return stem
    return m.group(1).replace("_", "-")


def read_doc_plaintext(doc_path: str) -> str:
    custom = os.environ.get("OOPT_DOC_TEXT_CMD")
    if custom:
        out = subprocess.check_output(custom.format(path=doc_path), shell=True)
        return out.decode("utf-8", errors="replace")

    if sys.platform == "darwin":
        out = subprocess.check_output(
            ["textutil", "-convert", "txt", "-stdout", doc_path],
            stderr=subprocess.DEVNULL,
        )
        return out.decode("utf-8", errors="replace")

    raise RuntimeError(
        "Cannot read .doc files: use macOS, set OOPT_DOC_TEXT_CMD, or convert .doc to .txt."
    )


def expand_kv_spec(fragment: str) -> set[int]:
    """Expand strings like ``72,73,82-163``, ``кв. 63,64``, ``60 (ч)`` into integers."""
    s = _normalize_dashes(fragment)
    s = re.sub(r"\(\s*ч\s*\)", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\bкв\.?\s*", " ", s, flags=re.IGNORECASE)
    s = s.replace("–", "-")
    nums: set[int] = set()
    for raw in re.split(r"[;,]", s):
        part = raw.strip()
        if not part:
            continue
        part = re.sub(r"[^\d\-]+", " ", part)
        part = part.strip()
        if not part:
            continue
        m = re.match(r"^(\d+)\s*-\s*(\d+)$", part)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if a <= b:
                nums.update(range(a, b + 1))
            continue
        for m in re.finditer(r"\d+", part):
            nums.add(int(m.group(0)))
    return nums


_TYPE_ABBR = (
    (re.compile(r"государственный природный комплексный заказник", re.I), "ГПЗК"),
    (re.compile(r"государственный природный заказник", re.I), "ГПЗ"),
    (re.compile(r"комплексный памятник природы", re.I), "КПП"),
    (re.compile(r"памятник природы", re.I), "ПП"),
    (re.compile(r"заповедник", re.I), "ЗП"),
)


def _abbr_for_block(block: str) -> str:
    for rx, abbr in _TYPE_ABBR:
        if rx.search(block):
            return abbr
    return "ООПТ"


def _clean_title(title: str) -> str:
    t = title.strip()
    t = re.sub(r"\s*\([^)]*плановый[^)]*\)\s*", " ", t, flags=re.I)
    t = re.sub(r"\s+", " ", t).strip()
    return t


_UCH_LINE = re.compile(
    r"^([А-Яа-яЁё][А-Яа-яЁё\-]{0,50}[сС]кое)[:\.\-\s]*(?:-\s*(\d+)ч?)?(?:\s+кв\.?\s*(.*))?\s*$",
    re.IGNORECASE
)


def _parse_oopt_blocks(full_text: str, big_zone: str) -> list[OOPTRule]:
    text = _normalize_dashes(full_text)
    lines = [ln.rstrip() for ln in text.splitlines()]
    results: list[OOPTRule] = []

    current_title = None
    current_uch = None
    current_block_lines: list[str] = []

    # Regex to catch:
    # 1. Quoted names: "«Name»" or "1. «Name»"
    # 2. Numbered names: "12. Ржевский лесопарк"
    header_rx = re.compile(
        r"^(?:\s*)((?:\d+[\.\)]\s*)?[«\"].*?[»\"]|(?:\d+[\.\)]\s*)[А-Яа-яЁё][А-Яа-яЁё\-\s]{5,100})(?:\s*)$"
    )

    for i, line in enumerate(lines):
        raw = line.strip()
        if not raw:
            continue

        # 1. Detect New OOPT Header
        hm = header_rx.match(raw)
        # Avoid matching categorical headers or district headers as OOPT names
        is_metadata = re.match(r"(?:Государственный|Постановление|Памятник природы|заказник|Площадь|Участковое)", raw, re.I)
        
        if hm and not is_metadata:
            # Finish current block
            current_title = _clean_title(hm.group(1).strip('«»"" 0123456789.'))
            current_uch = None
            current_block_lines = [raw]
            continue

        if not current_title:
            continue

        # 2. Detect District (Uchastkovoe)
        um = _UCH_LINE.match(raw)
        if um:
            current_uch = um.group(1).strip()
            # Handle "Кавголовское кв. 63,64" (same line)
            if um.group(3):
                all_kvs = expand_kv_spec(um.group(3).strip())
                if all_kvs:
                    abbr = _abbr_for_block("\n".join(current_block_lines + [raw]))
                    results.append(
                        OOPTRule(
                            big_zone=big_zone,
                            uchastkovoe=current_uch,
                            kvartals=frozenset(all_kvs),
                            label=f"{abbr} {current_title}".strip(),
                            title=current_title,
                        )
                    )
            continue

        # 3. Detect Kvartals in subsequent lines if we have a district or just numbers
        # If we have a district, any numbers/kvartal line is a candidate
        if current_uch:
            all_kvs = expand_kv_spec(raw)
            # Must look like a kvartal line: "кв. 63,64" or a list of numbers
            if all_kvs and (re.search(r'кв\.?|части?', raw, re.I) or re.search(r'[\d,\s\-]{3,}', raw)):
                abbr = _abbr_for_block("\n".join(current_block_lines + [raw]))
                results.append(
                    OOPTRule(
                        big_zone=big_zone,
                        uchastkovoe=current_uch,
                        kvartals=frozenset(all_kvs),
                        label=f"{abbr} {current_title}".strip(),
                        title=current_title,
                    )
                )

        current_block_lines.append(raw)

    return results


def _load_supplement_rules(oopt_dir: str) -> list[OOPTRule]:
    path = os.path.join(oopt_dir, "oopt_supplements.json")
    if not os.path.isfile(path):
        return []
    with open(path, encoding="utf-8") as f:
        raw = json.load(f)
    out: list[OOPTRule] = []
    for item in raw:
        kvs = frozenset(int(x) for x in item["kvartals"])
        out.append(
            OOPTRule(
                big_zone=str(item["big_zone"]),
                uchastkovoe=str(item["uchastkovoe"]),
                kvartals=kvs,
                label=str(item["label"]),
                title=str(item.get("title", item["label"])),
            )
        )
    return out


def load_rules_from_dir(oopt_dir: str) -> list[OOPTRule]:
    all_rules: list[OOPTRule] = []
    if not os.path.isdir(oopt_dir):
        return all_rules

    all_rules.extend(_load_supplement_rules(oopt_dir))

    for name in sorted(os.listdir(oopt_dir)):
        if not name.lower().endswith(".doc") and not name.lower().endswith(".txt"):
            continue
        path = os.path.join(oopt_dir, name)
        if name.lower().endswith(".txt"):
            with open(path, encoding="utf-8", errors="replace") as f:
                text = f.read()
        else:
            try:
                text = read_doc_plaintext(path)
            except (RuntimeError, subprocess.CalledProcessError, FileNotFoundError):
                continue
        zone = forest_zone_from_filename(name)
        all_rules.extend(_parse_oopt_blocks(text, zone))
    return all_rules


def _big_zone_aliases(big: str) -> set[str]:
    b = normalize_zone_name(big)
    if not b:
        return set()
    out = {b, b.replace("-", " "), b.replace(" ", "-")}
    return {x for x in out if x}


def match_oopt(
    rules: Iterable[OOPTRule],
    lesnichestvo: str | None,
    uchastkovoe: str | None,
    kvartal: int | None,
) -> str | None:
    """
    Return table label (e.g. ``ГПЗ Моторное-Заостровье``) or None.
    If several rules match, picks the narrowest kv set; ties join with ``; ``.
    """
    if kvartal is None or uchastkovoe is None:
        return None
    uch_n = normalize_zone_name(uchastkovoe)
    if not uch_n:
        return None

    big_aliases = _big_zone_aliases(lesnichestvo) if lesnichestvo else set()

    candidates: list[OOPTRule] = []
    for r in rules:
        if kvartal not in r.kvartals:
            continue
        if big_aliases and normalize_zone_name(r.big_zone) not in big_aliases:
            continue
        candidates.append(r)

    if not candidates:
        return None

    # Scoring: 
    # 1. Exact district match (uch_n)
    # 2. Same big_zone
    def match_score(r: OOPTRule) -> int:
        score = 0
        if normalize_zone_name(r.uchastkovoe) == uch_n:
            score += 10
        return score

    best_score = max(match_score(r) for r in candidates)
    if best_score == 0 and uch_n:
        return None
    
    best_candidates = [r for r in candidates if match_score(r) == best_score]
    
    # narrowest kv set among best scores
    best_size = min(len(r.kvartals) for r in best_candidates)
    finalists = [r for r in best_candidates if len(r.kvartals) == best_size]

    def key(r: OOPTRule) -> tuple[int, str]:
        return (len(r.kvartals), r.label)

    best_size = min(len(r.kvartals) for r in candidates)
    best = [r for r in candidates if len(r.kvartals) == best_size]
    labels = sorted({r.label for r in best})
    return "; ".join(labels) if labels else None


def parse_kv_cell(cell: str | None) -> tuple[int | None, int | None]:
    """
    ``26(12)`` -> квартал 26, выдел 12 (table convention in sample xlsx).
    Also accepts ``кв. 26 выд. 12``.
    """
    if cell is None:
        return None, None
    s = str(cell).strip()
    if not s or s.lower() == "nan":
        return None, None
    m = re.search(r"кв\.?\s*(\d+).*?выд\.?\s*(\d+)", s, re.I)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.match(r"^(\d+)\s*\(\s*(\d+)\s*\)\s*$", s)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.match(r"^(\d+)\s*$", s)
    if m:
        return int(m.group(1)), None
    return None, None


@lru_cache(maxsize=1)
def get_cached_rules(oopt_dir: str) -> Tuple[OOPTRule, ...]:
    return tuple(load_rules_from_dir(oopt_dir))
