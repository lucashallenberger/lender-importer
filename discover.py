"""
Step 1: Discover the Salesforce object behind the form, and dump its fields.

Usage:
    ./.venv/bin/python discover.py

Behaviour:
  * If [object] api_name is blank in config.ini, this scans your org for the
    custom object that matches the form's distinctive fields (Loan Amount, LTV,
    Capital Type, Declined Reason, etc.) and prints the best candidates.
  * Once you paste the winning api_name into config.ini, run it again and it
    will dump every field (API name, label, type, picklist values, lookup
    targets) to the screen and to fields.txt.
"""

import os
import sys

from src.sf_client import connect, load_config, get_object_api_name

# Field labels that are distinctive to the form shown (used to auto-find the object).
SIGNATURE_LABELS = {
    "loan amount", "ltv", "ltc", "capital type", "declined reason",
    "interest rate", "dcr", "prepayment", "recourse", "spread",
    "amortization", "term (years)", "origination fee", "capital source",
}

# Keywords to pre-filter candidate objects so we don't describe the whole org.
NAME_HINTS = ("lender", "capital", "interest", "deal", "source", "marketing", "lead", "activity")


def find_object(sf):
    print("Scanning org for the object that matches the form...\n")
    glob = sf.describe()
    candidates = []
    for o in glob["sobjects"]:
        name = o["name"]
        label = (o["label"] or "").lower()
        if not o.get("createable"):
            continue
        # focus on custom objects, or standard objects whose label hints a match
        is_custom = name.endswith("__c")
        hinted = any(h in label or h in name.lower() for h in NAME_HINTS)
        if is_custom or hinted:
            candidates.append(name)

    scored = []
    for name in candidates:
        try:
            desc = getattr(sf, name).describe()
        except Exception:
            continue
        labels = {(f["label"] or "").lower() for f in desc["fields"]}
        hits = SIGNATURE_LABELS & labels
        if hits:
            scored.append((len(hits), name, desc["label"], sorted(hits)))

    scored.sort(reverse=True)
    if not scored:
        print("Could not auto-identify the object. You may need to paste its API")
        print("name into config.ini manually (ask your Salesforce admin).")
        return

    print("Best matches (most matching fields first):\n")
    for score, name, label, hits in scored[:8]:
        print(f"  {score:>2} matching fields   API name: {name}   (label: {label})")
        print(f"        matched: {', '.join(hits)}\n")
    print("-> Paste the API name of the right one into config.ini under [object] api_name,")
    print("   then run discover.py again to see all its fields.")


def dump_fields(sf, api_name):
    try:
        desc = getattr(sf, api_name).describe()
    except Exception as exc:  # noqa: BLE001
        sys.exit(f"ERROR: could not describe object '{api_name}': {exc}")

    lines = []
    lines.append(f"OBJECT: {api_name}  (label: {desc['label']})")
    lines.append(f"Total fields: {len(desc['fields'])}")
    lines.append("=" * 78)
    lines.append("")
    for f in desc["fields"]:
        if not (f["createable"] or f["updateable"]):
            continue  # skip read-only/system fields you can't write
        line = f"{f['label']!r}"
        lines.append(f"FIELD: {f['name']}")
        lines.append(f"    label:    {f['label']}")
        lines.append(f"    type:     {f['type']}" + (" (required)" if not f["nillable"] and f["createable"] else ""))
        if f["type"] in ("picklist", "multipicklist"):
            vals = [v["label"] for v in f["picklistValues"] if v["active"]]
            lines.append(f"    picklist: {vals}")
        if f["type"] == "reference":
            lines.append(f"    lookup -> {f['referenceTo']}")
        lines.append("")

    text = "\n".join(lines)
    print(text)

    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fields.txt")
    with open(out, "w") as fh:
        fh.write(text)
    print(f"\n(Also saved to {out})")


def main():
    cfg = load_config()
    sf = connect(cfg)
    print(f"Connected to Salesforce as {cfg['salesforce']['username']}\n")

    api_name = get_object_api_name(cfg)
    if api_name:
        dump_fields(sf, api_name)
    else:
        find_object(sf)


if __name__ == "__main__":
    main()
