"""Backend API for the desktop app. Wraps the matching/upload logic from
import_lenders into stateful, JSON-friendly calls the UI can invoke.

Flow the UI drives:
  check_setup() -> save_settings() if needed
  load_excel(path) -> search_deals(q) -> analyze(deal, property)
  -> (answer each question) -> analyze() again to refresh -> upload()
"""

import io
import contextlib
import os

import import_lenders as il
from src.sf_client import connect, load_config, config_exists, save_config, get_object_api_name
from learned_store import load_learned, remember_account, remember_contact
from fuzzy import closest_matches

DEAL_FIELD = il.DEAL_FIELD


def _silent(fn, *a, **k):
    """Run a chatty function while swallowing its stdout prints."""
    with contextlib.redirect_stdout(io.StringIO()):
        return fn(*a, **k)


class Api:
    def __init__(self):
        self.sf = None
        self.desc = None
        self.api_name = None
        self.learned = load_learned()
        self.records, self.headers = [], []
        self.excel_name = ""
        self.deal_id = self.deal_name = None
        self.property_text = ""
        self.questions = {}      # qid -> question dict
        self.skips = set()       # qids the user chose to skip

    # ---- setup / connection -------------------------------------------------
    def check_setup(self):
        return {"configured": config_exists()}

    def save_settings(self, username, password, security_token, domain):
        try:
            save_config(username, password, security_token, domain or "login",
                        api_name="ascendix__DealSource__c")
            self.sf = None  # force reconnect with new creds
            self.connect()
            return {"ok": True, "object": self.api_name}
        except SystemExit as e:
            return {"ok": False, "error": str(e)}
        except Exception as e:  # noqa: BLE001
            return {"ok": False, "error": str(e)}

    def connect(self):
        if self.sf is not None:
            return
        cfg = load_config()
        self.sf = _silent(connect, cfg)
        self.api_name = get_object_api_name(cfg) or "ascendix__DealSource__c"
        self.desc = il.describe_object(self.sf, self.api_name)

    # ---- inputs -------------------------------------------------------------
    def load_excel(self, path, sheet=None):
        try:
            self.connect()
            headers, records = il.read_excel(path, sheet)
            _silent(il.validate_mapping, self.desc, headers)
            self.headers, self.records, self.excel_name = headers, records, os.path.basename(path)
            return {"ok": True, "rows": len(records), "file": self.excel_name,
                    "columns": [h for h in headers if h]}
        except SystemExit as e:
            return {"ok": False, "error": str(e)}
        except Exception as e:  # noqa: BLE001
            return {"ok": False, "error": str(e)}

    def search_deals(self, query):
        self.connect()
        target = {f["name"]: f for f in self.desc["fields"]}[DEAL_FIELD]["referenceTo"][0]
        q = (query or "").replace("'", "")
        if not q:
            return []
        rows = self.sf.query_all(
            f"SELECT Id, Name FROM {target} WHERE Name LIKE '%{q}%' ORDER BY Name LIMIT 25"
        )["records"]
        return [{"id": r["Id"], "name": r["Name"]} for r in rows]

    # ---- analysis -----------------------------------------------------------
    def analyze(self, deal_id, deal_name, property_text):
        self.connect()
        self.deal_id, self.deal_name = deal_id, deal_name
        self.property_text = property_text

        acct = _silent(il.build_account_index, self.sf, self.records, self.learned)
        account_index, real_name, acct_corr, acct_sugg, acct_unmatched = acct
        scoped, by_name = _silent(il.build_contact_indexes, self.sf, self.records)
        built, skipped, info = _silent(
            il.build_records, self.sf, self.records, self.desc, account_index, real_name,
            scoped, by_name, deal_id, property_text, False, self.learned)

        self._built, self._skipped = built, skipped
        # assemble questions (account ambiguous + contact ambiguous/first-name)
        questions = {}
        for typed, cands in acct_sugg.items():
            qid = f"account:{typed.lower()}"
            questions[qid] = {
                "id": qid, "kind": "account", "typed": typed, "row": None, "account_id": None,
                "candidates": [{"id": i, "name": n, "score": int(s * 100)} for n, i, s in cands],
            }
        for row, typed, cands in info["suggestions"]:
            qid = f"contact:{row}:{typed.lower()}"
            aid = account_index.get((self._row_account(row) or "").lower())
            questions[qid] = {
                "id": qid, "kind": "contact", "typed": typed, "row": row, "account_id": aid,
                "candidates": [{"id": i, "name": n, "score": int(s * 100)} for n, i, s in cands],
            }
        for row, typed, ids in info["ambiguous"]:
            qid = f"contact:{row}:{typed.lower()}"
            aid = account_index.get((self._row_account(row) or "").lower())
            cands = il._contact_candidates(self.sf, ids)
            questions[qid] = {
                "id": qid, "kind": "contact", "typed": typed, "row": row, "account_id": aid,
                "candidates": [{"id": c["Id"], "name": c["Name"], "score": None} for c in cands],
            }
        self.questions = questions
        pending = [q for qid, q in questions.items() if qid not in self.skips]

        auto = len(acct_corr) + len(info["corrections"])
        chosen = len([c for c in acct_corr if c[3] in ("learned", "chosen")]) + \
            len([c for c in info["corrections"] if c[4] in ("learned", "chosen")])
        blank = info["blank"] + len(info["unmatched"]) + len([1 for qid in self.questions if qid in self.skips])
        summary = {
            "rows": len(self.records), "ready": len(built), "skipped_rows": len(skipped),
            "auto_fixed": auto, "you_chose": chosen, "blank_contacts": blank,
            "unmatched_accounts": acct_unmatched,
        }
        preview = [{
            "name": rec.get(il.NAME_FIELD), "interest": rec.get("ascendix__Interest__c"),
            "account": bool(rec.get("ascendix__Account__c")),
            "contact": bool(rec.get("ascendix__Contact__c")),
        } for _, rec in built[:8]]
        return {"questions": pending, "summary": summary, "preview": preview,
                "deal": self.deal_name}

    def _row_account(self, row):
        idx = row - 2
        if 0 <= idx < len(self.records):
            return il.clean_text(self.records[idx].get(il.ACCOUNT_COLUMN))
        return None

    # ---- answering questions ------------------------------------------------
    def answer(self, qid, choice_id, remember=True):
        q = self.questions.get(qid)
        if not q:
            return {"ok": False, "error": "unknown question"}
        if choice_id is None:                      # skip -> leave blank, don't re-ask
            self.skips.add(qid)
            return {"ok": True}
        cand = next((c for c in q["candidates"] if c["id"] == choice_id), None)
        if not cand:
            return {"ok": False, "error": "unknown choice"}
        if q["kind"] == "account":
            self._learn_account(q["typed"], cand["name"], cand["id"], remember)
        else:
            self._learn_contact(q["account_id"], q["typed"], cand["name"], cand["id"], remember)
        return {"ok": True}

    def _learn_account(self, typed, name, rid, remember):
        if remember:
            remember_account(self.learned, typed, name, rid)
        else:
            self.learned["accounts"][typed.strip().lower()] = {"name": name, "id": rid}

    def _learn_contact(self, account_id, typed, name, rid, remember):
        if remember:
            remember_contact(self.learned, account_id, typed, name, rid)
        else:
            self.learned["contacts"][f"{account_id or '?'}::{typed.strip().lower()}"] = {"name": name, "id": rid}

    # ---- upload -------------------------------------------------------------
    def upload(self):
        if not getattr(self, "_built", None):
            return {"ok": False, "error": "nothing analyzed yet"}
        payload = [rec for _, rec in self._built]
        try:
            results = getattr(self.sf.bulk, self.api_name).insert(payload, batch_size=200)
            ok = sum(1 for r in results if r.get("success"))
            fails = [{"row": rn, "error": str(res.get("errors"))}
                     for (rn, _), res in zip(self._built, results) if not res.get("success")]
            return {"ok": True, "created": ok, "failed": len(fails), "errors": fails[:20],
                    "deal": self.deal_name}
        except Exception as e:  # noqa: BLE001
            return {"ok": False, "error": str(e)}
