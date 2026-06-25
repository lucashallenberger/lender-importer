"""Persistent memory of the ambiguous-match choices you make interactively.

When you pick the right Account or Contact for a name the tool couldn't resolve on
its own, that decision is saved to learned.json and applied automatically next time
- so you're never asked the same question twice.

Accounts are remembered by typed name. Contacts are remembered per ACCOUNT (keyed by
the account's Id) so "Tyler" -> Tyler Jackson at one lender never leaks to another.
"""

import json
import os

from app_paths import data_file

_PATH = data_file("learned.json")


def load_learned():
    if os.path.exists(_PATH):
        try:
            with open(_PATH) as fh:
                data = json.load(fh)
                data.setdefault("accounts", {})
                data.setdefault("contacts", {})
                return data
        except Exception:  # noqa: BLE001 - corrupt file shouldn't break a run
            pass
    return {"accounts": {}, "contacts": {}}


def save_learned(data):
    with open(_PATH, "w") as fh:
        json.dump(data, fh, indent=2, sort_keys=True)


def _ck(account_id, typed):
    return f"{account_id or '?'}::{(typed or '').strip().lower()}"


def remember_account(data, typed, name, rid):
    data["accounts"][(typed or "").strip().lower()] = {"name": name, "id": rid}
    save_learned(data)


def lookup_account(data, typed):
    return data["accounts"].get((typed or "").strip().lower())


def remember_contact(data, account_id, typed, name, rid):
    data["contacts"][_ck(account_id, typed)] = {"name": name, "id": rid}
    save_learned(data)


def lookup_contact(data, account_id, typed):
    return data["contacts"].get(_ck(account_id, typed))
