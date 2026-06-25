"""One-off helper: list duplicate-name Accounts/Contacts that the import would hit,
with details so you can confirm the right record (or decide to match Contacts by email)."""

import openpyxl
from src.sf_client import connect, load_config, get_object_api_name

XLSX = "/Users/lucashallenberger/Downloads/251015 - Lender List - Ascendix.xlsx"


def clean(v):
    if v is None:
        return None
    s = str(v).replace("_x000D_", "").strip()
    return s or None


def main():
    wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    hdr = [clean(h) for h in rows[0]]
    data = [{hdr[i]: r[i] for i in range(len(hdr)) if hdr[i]} for r in rows[1:]]

    sf = connect(load_config())

    # ---- Accounts ----
    acct_names = sorted({clean(r.get("Account")) for r in data if clean(r.get("Account"))})
    quoted = ", ".join("'" + n.replace("'", "\\'") + "'" for n in acct_names)
    arows = sf.query_all(
        f"SELECT Id, Name, Type, BillingCity, BillingState FROM Account WHERE Name IN ({quoted})"
    )["records"]
    abyname = {}
    for a in arows:
        abyname.setdefault(a["Name"].strip().lower(), []).append(a)

    print("=" * 80)
    print("DUPLICATE ACCOUNTS (more than one Account with the same name)")
    print("=" * 80)
    found = False
    for name in acct_names:
        cands = abyname.get(name.lower(), [])
        if len(cands) > 1:
            found = True
            print(f"\n  '{name}'  ({len(cands)} records):")
            for a in cands:
                loc = ", ".join(x for x in [a.get("BillingCity"), a.get("BillingState")] if x)
                print(f"      Id {a['Id']} | Type: {a.get('Type')} | {loc or 'no location'}")
    if not found:
        print("  none")

    # ---- Contacts ----
    con_names = sorted({clean(r.get("Contact")) for r in data if clean(r.get("Contact"))})
    quoted = ", ".join("'" + n.replace("'", "\\'") + "'" for n in con_names)
    crows = sf.query_all(
        f"SELECT Id, Name, Email, Title, Account.Name FROM Contact WHERE Name IN ({quoted})"
    )["records"]
    cbyname = {}
    for c in crows:
        cbyname.setdefault(c["Name"].strip().lower(), []).append(c)

    # map excel email per contact name
    email_for = {}
    for r in data:
        cn = clean(r.get("Contact"))
        if cn:
            email_for.setdefault(cn.lower(), set()).add(clean(r.get("Email")))

    print("\n" + "=" * 80)
    print("DUPLICATE CONTACTS (more than one Contact with the same name)")
    print("Excel email is shown so you can see which SF contact matches by email.")
    print("=" * 80)
    found = False
    for name in con_names:
        cands = cbyname.get(name.lower(), [])
        if len(cands) > 1:
            found = True
            xl_emails = ", ".join(sorted(e for e in email_for.get(name.lower(), set()) if e)) or "(none)"
            print(f"\n  '{name}'   Excel email: {xl_emails}")
            xl_email_set = {e.lower() for e in email_for.get(name.lower(), set()) if e}
            for c in cands:
                acct = (c.get("Account") or {}).get("Name") if c.get("Account") else None
                star = "  <-- email match" if (c.get("Email") or "").lower() in xl_email_set else ""
                print(f"      Id {c['Id']} | {c.get('Email')} | {c.get('Title')} | Account: {acct}{star}")
    if not found:
        print("  none")


if __name__ == "__main__":
    main()
