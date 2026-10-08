"""Property owners: splitting owner lists, comparing names, deciding the
owners of a collateral and matching them to its borrowers.

Name comparison ignores case, punctuation, extra spaces and titles (Mr, Mrs,
Smt, Shri, Dr, M/s...). It also treats "PVT LTD" and "PRIVATE LIMITED" as
the same, lets an initial match a whole name ("G" matches "GOPAL"), and
allows common spelling variants of a word ("LAXMI" / "LAKSHMI").

Relatives share the surname and often the father's or husband's name, so:
  - two names that each have a word the other lacks ("SUNANDA RAMCHANDRA
    SAWANT", "RAMCHANDRA GANPAT SAWANT") are different people, however similar
    they look overall;
  - when only one name has extra words, the given names must correspond: a
    wife's name usually contains her husband's given name and surname
    ("DEEPA RAVIRAJ HIREMATH" / "RAVIRAJ HIREMATH"), a missing middle name is
    fine ("ROHAN JOSHI" / "ROHAN MAHESH JOSHI");
  - the owners of one property are linked to borrowers one to one, best match
    first (assign_owners): owners listed separately are different people.
"""
import re
from itertools import permutations

from rapidfuzz import fuzz

from .. import settings
from .text import as_text, straighten

TITLES = r"(mr|mrs|ms|miss|smt|shri|sri|shrimati|sh|dr|m/s|messrs|kum|km|late|sardar|sardarni)"
COMPANY_WORDS = re.compile(r"^(company|co|associates|enterprises?|industries|traders?|sons|brothers|bros|"
                           r"pvt|private|ltd|limited|llp|trust|group|agencies|agency|corporation)\b", re.I)
# words that make a name a business ("JAMSON MACHINE AND TOOLS" is one owner)
BUSINESS_WORDS = re.compile(r"\b(company|co|associates|enterprises?|industries|traders?|trading|sons|brothers|bros|"
                            r"pvt|private|ltd|limited|llp|trust|group|agencies|agency|corporation|firm|infra\w*|"
                            r"constructions?|builders?|developers?|estates?|properties|realty|tools|paints|contracts?|"
                            r"centre|center|motors|textiles?|fabrics|foods|exports?|imports?|logistics|solutions|"
                            r"services|technologies|tech|systems|works|stores|marketing|hotels?|hospital|clinic|"
                            r"school|society|foundation|udhyog|udyog)\b", re.I)
# "KAMALA D/O HIMMATMAL", "ASHISH SO DOULATRAM": the words after S/O, D/O or W/O
# name a relative (father or husband), not the owner
RELATION = r"(?:[SDW]\s*/\s*O|SO|DO|WO|(?:SON|DAUGHTER|WIFE)\s+OF)[\s.]+(?=\S)"


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
        elif owners and re.match(RELATION, piece, re.I):     # "B. RAMESH, S/O B. KRISHNAPPA"
            owners[-1] = f"{owners[-1]} {piece}"
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
    """The name as compared: capitals, no titles or punctuation, and without a
    relative's name after S/O, D/O or W/O."""
    text = strip_titles(straighten(as_text(name))).upper()
    text = re.sub(rf"\s+{RELATION}.*$", "", text)
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
    """Pairs the words of two names. Returns (left_a, left_b, counts, pairs):
      left_a / left_b  the words of each name with no counterpart;
      counts           {"exact", "initials", "close", "strong"}: how many pairs are
                       the same word (also written differently or as two words:
                       "ANILKUMAR" / "ANIL KUMAR"), an initial for a word ("G" /
                       "GOPAL", "CH" / "CHANDU"), only a close spelling (someone
                       should confirm), and how many join two real words (3+
                       letters) rather than an initial;
      pairs            [(positions in a, positions in b)] of the paired words."""
    free_a, free_b = list(range(len(a))), list(range(len(b)))
    pairs = []
    counts = {"exact": 0, "initials": 0, "close": 0, "strong": 0}

    def take(pos_a, pos_b, kind, strong):
        for i in pos_a:
            free_a.remove(i)
        for j in pos_b:
            free_b.remove(j)
        pairs.append((set(pos_a), set(pos_b)))
        counts[kind] += 1
        counts["strong"] += strong

    for i in list(free_a):                              # the same word
        j = next((j for j in free_b if b[j] == a[i]), None)
        if j is not None:
            take([i], [j], "exact", len(a[i]) > 2)
    for i in list(free_a):                              # the same word, written differently
        j = next((j for j in free_b if len(a[i]) > 2 and len(b[j]) > 2 and _same_word(a[i], b[j])), None)
        if j is not None:
            take([i], [j], "exact", True)
    for words, other, own_free, other_free, a_side in ((a, b, free_a, free_b, True), (b, a, free_b, free_a, False)):
        for i in list(own_free):                        # one word written as two
            if i not in own_free or len(words[i]) <= 2:
                continue
            for x, y in permutations(list(other_free), 2):
                joined = other[x] + other[y]
                if joined == words[i] or _same_word(joined, words[i]):
                    kind = "exact"
                elif len(words[i]) >= 8 and fuzz.ratio(_sounds(joined), _sounds(words[i])) >= 90:
                    kind = "close"
                else:
                    continue
                if a_side:
                    take([i], [x, y], kind, True)
                else:
                    take([x, y], [i], kind, True)
                break
    for words, other, own_free, other_free, a_side in ((a, b, free_a, free_b, True), (b, a, free_b, free_a, False)):
        for i in list(own_free):                        # an initial for a word
            w = words[i]
            if i not in own_free or len(w) > 2:
                continue
            j = next((j for j in other_free if len(other[j]) > len(w) and other[j].startswith(w)), None)
            kind = "initials"
            if j is None and len(w) == 2:               # a word cut short ("BU" for "ABU")
                j = next((j for j in other_free if len(other[j]) > 2 and other[j].endswith(w)), None)
                kind = "close"
            if j is not None:
                if a_side:
                    take([i], [j], kind, False)
                else:
                    take([j], [i], kind, False)
    for i in list(free_a):                              # close spellings
        j = next((j for j in free_b if len(a[i]) > 2 and len(b[j]) > 2 and _close(a[i], b[j])), None)
        if j is not None:
            take([i], [j], "close", True)
    # a word repeated by mistake in a name ("RAMAVADH RAMAVADH") is not a difference
    left_a = [a[i] for i in free_a if a.count(a[i]) < 2]
    left_b = [b[j] for j in free_b if b.count(b[j]) < 2]
    return left_a, left_b, counts, pairs


def _given_names_correspond(pairs, words_a: list, words_b: list):
    """When one name has words the other lacks: do the two given names pair up?
    The surname is the paired word at the end of the names (last in "RAVIRAJ
    HIREMATH", first in "SHAH RAGINI DHIRAJ"); the given name is the first word
    of the rest. A wife's name usually contains her husband's given name and
    surname ("DEEPA RAVIRAJ HIREMATH" / "RAVIRAJ HIREMATH"): there the given
    names (DEEPA, RAVIRAJ) do not pair up. Returns True, False, or None when
    the order of the names cannot be told."""
    len_a, len_b = len(words_a), len(words_b)
    # A name that is one word, or whose words all pair with one word of the
    # other ("GANGA DEVI" / "GANGADEVI"), is a given name: it must pair with
    # the other name's given name (normally its first word), not with its last
    # word ("SOMASHEKAR" / "GEETHA SOMASHEKAR" is her husband's name).
    for side, own_len, other_words in ((0, len_a, words_b), (1, len_b, words_a)):
        whole = next((pair for pair in pairs if len(pair[side]) == own_len), None)
        if whole is not None:
            other_positions = whole[1 - side]
            if _given_position(0, other_words, pairs, 1 - side) in other_positions:
                return True
            if len(other_words) - 1 in other_positions:
                return False
            return None

    def ends(positions, length):
        return ({"first"} if 0 in positions else set()) | ({"last"} if length - 1 in positions else set())
    found = {(ea, eb) for pa, pb in pairs for ea in ends(pa, len_a) for eb in ends(pb, len_b)}
    if ("last", "last") in found:
        order = ("last", "last")
    elif ("first", "last") in found and ("last", "first") in found:
        return None
    elif ("first", "last") in found:
        order = ("first", "last")
    elif ("last", "first") in found:
        order = ("last", "first")
    else:
        return True if ("first", "first") in found else None
    given_a = _given_position(0 if order[0] == "last" else 1, words_a, pairs, 0)
    given_b = _given_position(0 if order[1] == "last" else 1, words_b, pairs, 1)
    if given_a >= len_a or given_b >= len_b:
        return True
    return any(given_a in pa and given_b in pb for pa, pb in pairs)


# words written before a given name that one record may leave out
_NAME_PREFIXES = {"MOHAMMED"}


def _given_position(start: int, words: list, pairs, side: int) -> int:
    """The position of the given name, from 'start' on, skipping an initial or
    a "MOHAMMED" that the other name does not have ("M GURU RAGHAVENDRA" ->
    GURU, "MOHAMMED SHAKIL KHAN" -> SHAKIL): a leading initial is usually the
    father's."""
    paired = {i for pair in pairs for i in pair[side]}
    position = start
    while position < len(words) - 1 and position not in paired and \
            (len(words[position]) <= 2 or words[position] in _NAME_PREFIXES):
        position += 1
    return position


def compare_names(a: str, b: str):
    """(score, kind): name_score() and why. kind is
      "same"       the same person (score 90+);
      "spelling"   probably the same person: close spellings, or too few words
                   alike ("BANU K" / "BHANU KATAM", "VIJAY" / "VIJAY RAMSWARUP");
      "unclear"    probably the same, but the word order cannot be told
                   ("POOJABEN ASHWINBHAI CHAUHAN" / "CHAUHAN POOJA");
      "relative"   similar, but the given names differ: probably a relative
                   ("GEETHA SOMASHEKAR" / "SOMASHEKAR", her husband);
      "different"  different people."""
    na, nb = normalise_name(a), normalise_name(b)
    if not na or not nb:
        return 0, "different"
    if na == nb or na.replace(" ", "") == nb.replace(" ", ""):
        return 100, "same"
    words_a, words_b = na.split(), nb.split()
    left_a, left_b, counts, pairs = _pair_words(words_a, words_b)
    fuzzy = int(round(max(fuzz.token_set_ratio(na, nb), fuzz.ratio(na.replace(" ", ""), nb.replace(" ", "")))))
    if (left_a and left_b) or counts["strong"] == 0:
        return min(fuzzy, settings.PROBABLE_PERSON_SCORE - 15), "different"
    probable = max(85, min(fuzzy, 89))
    if left_a or left_b:
        corresponds = _given_names_correspond(pairs, words_a, words_b)
        if corresponds is False:
            return probable, "relative"
        if corresponds is None:
            return probable, "unclear"
    if counts["close"] == 0 and counts["exact"] >= 2:
        return 95, "same"
    return probable, "spelling"


def name_score(a: str, b: str) -> int:
    """0-100 similarity between two person or company names.
      100    the same after tidying;
      95     the same apart from initials or extra words, with at least two
             words exactly alike and the given names corresponding
             ("RAMESH K PATIL" / "RAMESH KRISHNA PATIL", "ROHAN JOSHI" / "ROHAN
             MAHESH JOSHI");
      85-89  probably the same (spelling differences, too few words alike, or
             extra words where the given names do not clearly correspond, as in
             "DEEPA RAVIRAJ HIREMATH" / "RAVIRAJ HIREMATH", a wife and her
             husband): someone should confirm;
      <= 60  different: each name has a word the other lacks, i.e. a different
             given name ("SUNANDA RAMCHANDRA SAWANT" / "RAMCHANDRA GANPAT SAWANT"),
             or only initials correspond."""
    return compare_names(a, b)[0]


def assign_owners(owners: list, names: list) -> list:
    """Decides, for the owners of one property, which borrower each one is.
    Each owner is linked to one borrower at most and each borrower to one owner
    at most, best matches first (owners listed separately are different
    people, so a wife "DEEPA RAVIRAJ HIREMATH" cannot take the place of her
    husband "RAVIRAJ HIREMATH" when he is an owner too). Names of the kinds
    "same", "spelling" and "unclear" are linked; a "relative" is not.
    Returns, in owner order, (borrower index or None, score, kind, nearest
    borrower index). The nearest borrower is the one the decision can be
    corrected against in the required-details file (Y or N typed for that
    borrower): the linked borrower, or for a relative the most similar borrower
    that no other owner has, again one to one. Other unlinked owners are
    "different" and have no nearest borrower."""
    compared = sorted(((*compare_names(owner, name), oi, ni) for oi, owner in enumerate(owners)
                       for ni, name in enumerate(names)), key=lambda t: (-t[0], t[2], t[3]))
    compared = [t for t in compared if t[0] >= settings.PROBABLE_PERSON_SCORE]
    linked, near, used = {}, {}, set()
    for score, kind, oi, ni in compared:
        if kind != "relative" and oi not in linked and ni not in used:
            linked[oi] = (ni, score, kind)
            used.add(ni)
    for score, kind, oi, ni in compared:
        if oi not in linked and oi not in near and ni not in used:
            near[oi] = (ni, score, kind)
            used.add(ni)
    out = []
    for oi, owner in enumerate(owners):
        if oi in linked:
            ni, score, kind = linked[oi]
            out.append((ni, score, kind, ni))
        elif oi in near:
            ni, score, kind = near[oi]
            out.append((None, score, kind, ni))
        else:
            score = max((name_score(owner, name) for name in names), default=0)
            out.append((None, min(score, settings.PROBABLE_PERSON_SCORE - 1), "different", None))
    return out


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
    """Distinct names, keeping the first spelling of each person ("KAMALA D/O
    HIMMATMAL" and "KAMALA W/O BAJRANG LAL" stay two people)."""
    out = []
    for name in names:
        name = strip_titles(as_text(name))
        if name and all(name_score(name, o) < 100 or _relatives_differ(name, o) for o in out):
            out.append(name)
    return out


def _relatives_differ(a: str, b: str) -> bool:
    """Both names give a relative after S/O, D/O or W/O, and not the same one."""
    found = [re.search(rf"\s+{RELATION}(.*)$", strip_titles(straighten(as_text(n))).upper()) for n in (a, b)]
    return all(found) and name_score(found[0].group(1), found[1].group(1)) < 100
