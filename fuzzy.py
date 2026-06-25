"""Typo-tolerant lookup helpers (stdlib only).

Given a mistyped name, fetch plausible Salesforce candidates and rank them by
string similarity so a near-miss can be auto-corrected (high confidence) or
suggested (medium confidence). Never silently picks a weak match.
"""

import difflib
import re


def _norm(s):
    return re.sub(r"[^a-z0-9 ]", "", (s or "").lower()).strip()


# Generic company-suffix words. When the ONLY difference between a typed name and a
# candidate is one or more of these, it's almost certainly the same lender written
# short ("Asgore" -> "Asgore Partners"), so it's safe to auto-accept (when unique).
CORP_SUFFIXES = {
    "partners", "capital", "group", "management", "advisors", "advisers", "realty",
    "bank", "financial", "investments", "investment", "holdings", "company", "co",
    "corp", "corporation", "llc", "lp", "llp", "inc", "fund", "funds", "ventures",
    "equity", "properties", "mortgage", "trust", "asset", "assets", "securities",
    "real", "estate", "credit", "lending", "loan", "finance", "national", "global",
}


def is_clean_partial(typed, name):
    """True if `typed` is `name` with only company-suffix words removed
    (e.g. 'Asgore' vs 'Asgore Partners'), or an exact match."""
    nt, nc = _norm(typed), _norm(name)
    if nt == nc:
        return True
    if nc.startswith(nt + " "):
        rest = nc[len(nt):].split()
        return bool(rest) and all(w in CORP_SUFFIXES for w in rest)
    return False


def _tokens(name):
    return [t for t in _norm(name).split() if len(t) >= 3]


def candidate_query(sf, obj, typed, limit=120):
    """Fetch records whose Name shares a word or leading letters with `typed`.
    Casts a wide net (substring + 3-char prefixes) so transposition/double-letter
    typos near the start are still caught; difflib then ranks the results."""
    toks = _tokens(typed) or [_norm(typed)[:4]]
    clauses = []
    for t in toks:
        t = t.replace("'", "")
        clauses.append(f"Name LIKE '%{t}%'")     # substring (catches mid-word typos)
        clauses.append(f"Name LIKE '{t[:3]}%'")  # 3-char prefix (catches leading typos)
    head = _norm(typed)[:3].replace("'", "")
    if head:
        clauses.append(f"Name LIKE '{head}%'")
    where = " OR ".join(dict.fromkeys(clauses))
    soql = f"SELECT Id, Name FROM {obj} WHERE {where} LIMIT {limit}"
    try:
        return [(r["Name"], r["Id"]) for r in sf.query_all(soql)["records"]]
    except Exception:
        return []


def _score(nt, nc):
    """Similarity of typed (nt) vs candidate (nc), both normalized. Combines
    edit-distance (typos) with prefix/subset signals (abbreviations like
    'asgore' -> 'asgore partners')."""
    if nt == nc:
        return 1.0
    ratio = difflib.SequenceMatcher(None, nt, nc).ratio()
    t_tok, c_tok = nt.split(), nc.split()
    boost = 0.0
    if nc.startswith(nt + " "):                       # 'asgore' begins 'asgore partners'
        boost = 0.93
    elif t_tok and all(t in c_tok for t in t_tok):    # all typed words appear in candidate
        boost = 0.90
    elif len(nt) >= 4 and nt in nc:                   # typed is a substring of candidate
        boost = 0.85
    return round(max(ratio, boost), 3)


def _rank(typed, pairs):
    nt = _norm(typed)
    scored = [(name, rid, _score(nt, _norm(name))) for name, rid in pairs]
    scored.sort(key=lambda x: x[2], reverse=True)
    seen, out = set(), []
    for name, rid, score in scored:
        if rid not in seen:
            seen.add(rid)
            out.append((name, rid, score))
    return out


def closest_matches(sf, obj, typed, limit=5):
    """Return [(name, id, score 0..1), ...] best-first for a typed name."""
    return _rank(typed, candidate_query(sf, obj, typed))[:limit]


def closest_contacts_in_account(sf, typed, account_id, limit=5):
    """Rank contacts that belong to `account_id` against the typed name. Safe for
    typo-correcting a contact when we already know the right account."""
    if not account_id:
        return []
    try:
        rows = sf.query_all(f"SELECT Id, Name FROM Contact WHERE AccountId = '{account_id}'")["records"]
    except Exception:
        return []
    return _rank(typed, [(r["Name"], r["Id"]) for r in rows])[:limit]
