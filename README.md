# Salesforce Autofill — bulk-create lender records from Excel

Reads a lender-list spreadsheet and creates the records in Salesforce in one shot,
instead of filling the web form by hand dozens of times a day.

## What it does
- Reads your Excel (`Name, Capital Source, Account, Contact, Email, Deal, Phone, Notes, Interest`)
- Translates the **Account / Contact / Deal / Capital Source** names into Salesforce record IDs automatically
- Validates the **Interest** picklist values
- Cleans up messy notes text
- Creates all the records via the Salesforce API
- Has a **dry-run preview** so nothing is written until you approve

---

## First-time setup (each computer)

1. **Install Python 3** (3.9+). On Mac it's usually already there: `python3 --version`.
2. Open Terminal in this folder and create the environment:
   ```
   python3 -m venv .venv
   ./.venv/bin/pip install -r requirements.txt
   ```
3. **Set up your login:** copy the example config and fill in your own Salesforce details:
   ```
   cp config.example.ini config.ini
   ```
   Open `config.ini` and enter your username, password, and security token.
   (Get a security token from Salesforce → your avatar → Settings → *Reset My Security Token*; it's emailed to you.)

> `config.ini` holds your password and is **never shared** (it's git-ignored). Each person fills in their own.

---

## Usage

**Step 1 — find the object (only needed once):**
```
./.venv/bin/python discover.py
```
This prints the best-matching Salesforce object. Paste its API name into `config.ini`
under `[object] api_name`, then run `discover.py` again to dump all the field names
into `fields.txt`. Use that to finalize `mapping.py` (already pre-filled with best guesses).

**Step 2 — preview, then upload.** The sheet needs columns: `Account, Contact, Notes, Interest`.
You provide the **deal** and the **property** (used in the record Name) at run time:
```
# preview only — writes nothing:
./.venv/bin/python import_lenders.py "list.xlsx" --deal "Sample - 1234 Main" --property "1234 Main" --dry-run

# real upload (after the dry-run looks right):
./.venv/bin/python import_lenders.py "list.xlsx" --deal "Sample - 1234 Main" --property "1234 Main"
```
- **Record Name** is built as `"<LenderName> - <Property>"`, where `LenderName` is the
  Account name with any leading number stripped (`3650 REIT` → `REIT`).
- **`--deal "name"`** searches deals (you confirm if several match); or use
  **`--deal-id <Id>`** to pin the exact deal with no ambiguity.
- **Contacts** are matched by name **scoped to their Account** (no email needed).
- **Declined** rows are included with Declined Reason = `Unknown`.
- If a contact match is ambiguous and you're in a real terminal, it asks you to pick;
  a piped/`--dry-run` run leaves it blank and lists it instead.

Always run `--dry-run` first to check the names, the contact-match summary, and any blanks.

### Typo tolerance (misspelled Account/Contact names)
If a typed name doesn't match a Salesforce record exactly, the tool compares it to
similar records:
- **Very close** (≥90% similar, e.g. `Goldmann Sachs` → `Goldman Sachs`): **auto-corrected**,
  and the correction is printed in the dry-run so you can veto it. Correcting the Account
  also fixes the record Name.
- **Somewhat close** (70–90%, e.g. `JP Morgon` → `JP Morgan`): **not** auto-applied; shown as
  a "did you mean?" suggestion to confirm.
- **Not close**: left blank and reported.

**Abbreviations / partial names** work too: typing `Asgore` matches `Asgore Partners`,
`Affinius` matches `Affinius Capital`. But if the short name fits **more than one**
account (e.g. `Goldman` → Goldman Sachs / Goldman Sachs Bank USA), it asks instead of
guessing. An exact account named just `Asgore` always wins over the partial.

Contact typos/abbreviations are resolved only **within the row's matched Account**, so a
near-miss name can't jump to a different company. This also covers **first-name-only**
contacts: `Tyler` resolves to `Tyler Jackson` if he's the only Tyler at that account; if
there are several, it asks you to pick the right person.

### Interactive prompts (real terminal runs)
When you run the real import in a terminal (not `--dry-run`), anything ambiguous —
an account that matched several records, a first name with multiple people — is shown as a
numbered menu so you can pick the right one (or `s` to skip / `q` to quit). A real run also
asks `Proceed to create N records? (y/n)` before writing. `--dry-run` never prompts; it
just lists what would be asked.

**Your picks are remembered.** Whenever you resolve an ambiguous Account/Contact in a
prompt, the choice is saved to `learned.json` and applied automatically next time - so the
same question is never asked twice. Contact picks are remembered per account, so a choice
for one lender never leaks to another. (Delete entries from `learned.json` to forget them.)

For recurring quirks/shorthand you want to hard-code, add them in `aliases.py`
(e.g. `ACCOUNT_ALIASES = {"bofa": "Bank of America"}`). The full resolution order is:
exact -> learned -> alias -> corporate-suffix abbreviation -> typo -> unique partial -> ask.

---

## Putting it on someone else's computer
Copy the whole `SalesforceAutofill` folder to their machine **except** `.venv/` and
`config.ini` (those are per-machine). Then they repeat the *First-time setup* above with
their own login. `mapping.py` and the scripts are shared as-is.

(Later we can package this as a double-clickable app so non-technical users don't touch
Terminal at all — ask when you're ready.)
```
