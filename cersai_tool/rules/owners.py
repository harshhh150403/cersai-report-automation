"""Property owners: splitting owner lists, comparing names, deciding the
owners of a collateral and matching them to its borrowers.

Name comparison ignores case, punctuation, extra spaces and titles (Mr, Mrs,
Smt, Shri, Dr, M/s...). It also treats "PVT LTD" and "PRIVATE LIMITED" as
the same, lets an initial match a whole name ("G" matches "GOPAL"), and
allows common spelling variants of a word ("LAXMI" / "LAKSHMI").

Every word of the shorter name must match a word of the longer one. Relatives
share the surname and often the father's or husband's name, so two names that
differ in a given name ("SUNANDA RAMCHANDRA SAWANT", "RAMCHANDRA GANPAT
SAWANT") are different people however similar they look overall.
"""
import re
from itertools import permutations

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


# words written in several ways -> one spelling
SAME_WORDS = {"MOHD": "MOHAMMED", "MD": "MOHAMMED", "MOHAMMAD": "MOHAMMED", "MOHAMAD": "MOHAMMED",
              "MOHAMED": "MOHAMMED", "MUHAMMAD": "MOHAMMED", "MUHAMMED": "MOHAMMED", "PVT": "PRIVATE",
              "LTD": "LIMITED", "CO": "COMPANY"}
# spelling variants of Indian names (applied to single words before a fuzzy comparison)
_SOUNDS = [("KSH", "X"), ("PH", "F"), ("BH", "B"), ("DH", "D"), ("GH", "G"), ("JH", "J"), ("KH", "K"),
           ("TH", "T"), ("SH", "S"), ("EE", "I"), ("OO", "U"), ("W", "V"), ("Z", "J")]


def normalise_name(name: str) -> str:
    text = strip_titles(straighten(as_text(name))).upper()
    text = re.sub(r"[^A-Z0-9 ]+", " ", text)
    return " ".join(SAME_WORDS.get(word, word) for word in text.split())


def _sounds(word: str) -> str:
    for a, b in _SOUNDS:
        word = word.replace(a, b)
    word = re.sub(r"(.)\1+", r"\1", word)       # doubled letters
    return re.sub(r"Y$", "I", word)


def _base(word: str) -> str:
    """The word without a -BHAI / -BEN / -AMMA ending ("RAGINIBEN" -> "RAGINI")."""
    stripped = re.sub(r"(BHAI|BEN|AMMA)$", "", word)
    return stripped if len(stripped) >= 3 else word


def _close(x: str, y: str) -> bool:
    """Two words (of 3+ letters) that are spellings of the same name: near-identical
    ("SRINIVASRAO" / "SRINIVASARAO"), the same when spoken ("LAXMI" / "LAKSHMI"),
    the same apart from -BHAI / -BEN, or one the start of the other
    ("NUTALA" / "NUTALAPATI", "AHMED" / "AHMEDALI")."""
    if fuzz.ratio(x, y) >= 80 or fuzz.ratio(_sounds(x), _sounds(y)) >= 80:
        return True
    bx, by = _base(x), _base(y)
    if (bx, by) != (x, y) and (bx == by or fuzz.ratio(_sounds(bx), _sounds(by)) >= 85):
        return True
    short, long_ = sorted((bx, by), key=len)
    return len(short) >= 4 and long_.startswith(short)


def _same_word(x: str, y: str) -> bool:
    """The same name written differently: the same when spoken ("LAXMI" /
    "LAKSHMI", "JAGJEET" / "JAGJIT") or apart from a -BHAI / -BEN ending
    ("LATABEN" / "LATA")."""
    return _sounds(x) == _sounds(y) or (_base(x), _base(y)) != (x, y) and _sounds(_base(x)) == _sounds(_base(y))


def _pair_words(a: list, b: list):
    """Pairs the words of two names. Returns (left_a, left_b, exact, initials,
    close, strong): the words of each name with no counterpart; how many pairs
    are the same word (also written differently or as two words: "ANILKUMAR" /
    "ANIL KUMAR"), an initial for a word ("G" / "GOPAL", "CH" / "CHANDU"), or
    only a close spelling (someone should confirm); and how many pairs join two
    real words rather than an initial."""
    orig_a, orig_b = list(a), list(b)
    a, b = list(a), list(b)
    exact = initials = close = 0
    strong = 0              # pairs of real words (3+ letters), not just initials
    for w in list(a):                                   # the same word
        if w in b:
            a.remove(w)
            b.remove(w)
            exact += 1
            strong += len(w) > 2
    for w in list(a):                                   # the same word, written differently
        match = next((o for o in b if len(w) > 2 and len(o) > 2 and _same_word(w, o)), None)
        if match:
            a.remove(w)
            b.remove(match)
            exact += 1
            strong += 1
    for side, rest in ((a, b), (b, a)):                 # one word written as two
        for w in [x for x in side if len(x) > 2]:
            for x, y in permutations(rest, 2):
                joined = x + y
                if joined == w or _same_word(joined, w):
                    kind = "exact"
                elif len(w) >= 8 and fuzz.ratio(_sounds(joined), _sounds(w)) >= 90:
                    kind = "close"
                else:
                    continue
                side.remove(w)
                rest.remove(x)
                rest.remove(y)
                exact += kind == "exact"
                close += kind == "close"
                strong += 1
                break
    for side, rest in ((a, b), (b, a)):                 # an initial for a word
        for w in [x for x in side if len(x) <= 2]:
            match = next((o for o in rest if len(o) > len(w) and o.startswith(w)), None)
            if match:
                side.remove(w)
                rest.remove(match)
                initials += 1
            elif len(w) == 2:                           # a word cut short ("BU" for "ABU")
                match = next((o for o in rest if len(o) > 2 and o.endswith(w)), None)
                if match:
                    side.remove(w)
                    rest.remove(match)
                    close += 1
    for w in [x for x in a if len(x) > 2]:              # close spellings
        match = next((o for o in b if len(o) > 2 and _close(w, o)), None)
        if match:
            a.remove(w)
            b.remove(match)
            close += 1
            strong += 1
    # a word repeated by mistake in a name ("RAMAVADH RAMAVADH") is not a difference
    a = [w for w in a if orig_a.count(w) < 2]
    b = [w for w in b if orig_b.count(w) < 2]
    return a, b, exact, initials, close, strong


def name_score(a: str, b: str) -> int:
    """0-100 similarity between two person or company names.
      100    the same after tidying;
      95     the same apart from initials or extra words, with at least two
             words exactly alike ("RAMESH K PATIL" / "RAMESH KRISHNA PATIL");
      85-89  probably the same (spelling differences, or too few words alike):
             someone should confirm;
      <= 60  different: each name has a word the other lacks, i.e. a different
             given name ("SUNANDA RAMCHANDRA SAWANT" / "RAMCHANDRA GANPAT SAWANT")."""
    na, nb = normalise_name(a), normalise_name(b)
    if not na or not nb:
        return 0
    if na == nb or na.replace(" ", "") == nb.replace(" ", ""):
        return 100
    left_a, left_b, exact, initials, close, strong = _pair_words(na.split(), nb.split())
    fuzzy = int(round(max(fuzz.token_set_ratio(na, nb), fuzz.ratio(na.replace(" ", ""), nb.replace(" ", "")))))
    different = settings.PROBABLE_PERSON_SCORE - 15
    # each name has a word the other lacks, or only initials correspond
    # ("SUDEEP JAISWAL" / "J S G BUILDERS"): different
    if (left_a and left_b) or strong == 0:
        return min(fuzzy, different)
    if close == 0 and exact >= 2:
        return 95
    return max(85, min(fuzzy, 89))


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
