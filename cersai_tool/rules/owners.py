"""Property owners: splitting owner lists, comparing names, deciding the
owners of a collateral and matching them to its borrowers.

Name comparison ignores case, punctuation, extra spaces and titles (Mr, Mrs,
Smt, Shri, Dr, M/s...). It also treats "PVT LTD" and "PRIVATE LIMITED" as
the same, and lets an initial match a whole name ("G" matches "GOPAL").
"""
import re

from rapidfuzz import fuzz

from .. import settings
from .text import as_text, straighten

TITLES = r"(mr|mrs|ms|miss|smt|shri|sri|shrimati|sh|dr|m/s|messrs|kum|km|late)"
COMPANY_WORDS = re.compile(r"^(company|co|associates|enterprises?|industries|traders?|sons|brothers|bros|"
                           r"pvt|private|ltd|limited|llp|trust|group|agencies|agency|corporation)\b", re.I)


def split_owner_list(text) -> list:
    """'Mr. A & Mrs. B' -> ['A', 'B'];  'A, B and C' -> ['A', 'B', 'C'].
    A piece that starts with a company word is joined back to the one before
    it, so 'RAMPAL PAREEK AND COMPANY' stays one owner."""
    text = straighten(as_text(text)).strip(" .")
    if not text:
        return []
    pieces = [p.strip(" .") for p in re.split(r"\s*,\s*|\s*&\s*|\s+and\s+", text, flags=re.I)]
    owners = []
    for piece in pieces:
        if not piece:
            continue
        if owners and COMPANY_WORDS.match(piece):
            owners[-1] = f"{owners[-1]} and {piece}"
        else:
            owners.append(piece)
    return [strip_titles(o) for o in owners if strip_titles(o)]


def strip_titles(name: str) -> str:
    name = as_text(name)
    while True:
        new = re.sub(rf"^\s*{TITLES}\.?\s+", "", name, flags=re.I)
        new = re.sub(rf"^\s*{TITLES}\.(?=\S)", "", new, flags=re.I)
        if new == name:
            return name.strip(" .")
        name = new


def normalise_name(name: str) -> str:
    text = strip_titles(straighten(as_text(name))).upper()
    text = re.sub(r"\bPVT\b", "PRIVATE", text)
    text = re.sub(r"\bLTD\b", "LIMITED", text)
    text = re.sub(r"\bCO\b", "COMPANY", text)
    text = re.sub(r"[^A-Z0-9 ]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _word_cover(short_tokens, long_tokens):
    """Matches each word of the shorter name to a different word of the longer
    one: the same word, an initial ("G" for "GOPAL") or a near-identical
    spelling ("BANU" for "BHANU"). Returns (all covered?, exact matches)."""
    exact = 0
    pool = list(long_tokens)
    for token in short_tokens:
        if token in pool:
            pool.remove(token)
            exact += 1
            continue
        if len(token) == 1:
            match = next((w for w in pool if w.startswith(token)), None)
        else:
            match = next((w for w in pool if len(w) > 1 and fuzz.ratio(token, w) >= 80), None)
        if match is None:
            return False, exact
        pool.remove(match)
    return True, exact


def name_score(a: str, b: str) -> int:
    """0-100 similarity between two person or company names.
    100 = same after tidying; 95 = same apart from initials; 85 = probably the
    same (spelling differences: someone should confirm); otherwise a fuzzy
    similarity score."""
    na, nb = normalise_name(a), normalise_name(b)
    if not na or not nb:
        return 0
    if na == nb or na.replace(" ", "") == nb.replace(" ", ""):
        return 100
    ta, tb = na.split(), nb.split()
    short, long_ = (ta, tb) if len(ta) <= len(tb) else (tb, ta)
    fuzzy = int(round(max(fuzz.token_set_ratio(na, nb), fuzz.ratio(na.replace(" ", ""), nb.replace(" ", "")))))
    covered, exact = _word_cover(short, long_)
    if covered and (exact >= 2 or (exact >= 1 and len(short) == 1)):
        return max(95, fuzzy if fuzzy < 100 else 95)
    if covered:
        return max(85, min(fuzzy, 89))
    return fuzzy


def best_match(name: str, candidates: list):
    """(index, score) of the most similar candidate, or (None, 0)."""
    best, best_score = None, 0
    for i, cand in enumerate(candidates):
        score = name_score(name, cand)
        if score > best_score:
            best, best_score = i, score
    return best, best_score


def lists_agree(lms_owners: list, pdf_owners: list) -> bool:
    """True when both lists name the same people (same number, each LMS owner
    matching a different title-report owner)."""
    if len(lms_owners) != len(pdf_owners) or not lms_owners:
        return False
    pool = list(pdf_owners)
    for owner in lms_owners:
        index, score = best_match(owner, pool)
        if index is None or score < settings.SAME_PERSON_SCORE:
            return False
        pool.pop(index)
    return True


def unique_names(names) -> list:
    """Distinct names, keeping the first spelling of each person."""
    out = []
    for name in names:
        name = strip_titles(as_text(name))
        if name and all(name_score(name, o) < 100 for o in out):
            out.append(name)
    return out
