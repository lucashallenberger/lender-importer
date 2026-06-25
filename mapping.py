"""
Maps Excel columns -> Salesforce fields for ascendix__DealSource__c (label "Deal Source").

Field API names below were confirmed against the live org via discover.py (see fields.txt).
import_lenders.py also re-validates this mapping against the object at run time.

Each entry: "Excel Column Header": (sf_field_api_name, kind)
  kind = "text"     -> copied as-is (Notes are cleaned of carriage-return junk)
  kind = "picklist" -> value checked against the field's allowed picklist values
  kind = "lookup"   -> the Excel value is a NAME; the script finds the matching
                       record in Salesforce and stores its Id. The lookup target
                       object is read automatically from Salesforce.

NEW FORMAT (no Name / Email / Deal columns in the sheet):
  * The record Name is COMPUTED as  "<LenderName> - <Property>"  where LenderName
    is the Account name with any leading number stripped ("3650 REIT" -> "REIT")
    and Property is passed in at run time (--property).
  * The Deal is chosen at run time (--deal / --deal-id) and applied to every row.
  * Contacts are matched by name SCOPED TO the row's Account (no email anymore).
"""

# Excel columns expected in the new sheet: Account, Contact, Notes, Interest
MAPPING = {
    "Account":  ("ascendix__Account__c",  "lookup"),    # -> Account
    "Contact":  ("ascendix__Contact__c",  "lookup"),    # -> Contact (scoped to Account)
    "Interest": ("ascendix__Interest__c", "picklist"),
    "Notes":    ("Notes__c",              "text"),
}

# Deal is required on the object but is supplied via --deal/--deal-id, not a column.
REQUIRED_LOOKUPS = set()

# Required checkboxes on the object. Default them to False (= unchecked, matching the form).
DEFAULTS = {
    "ascendix__IsPrimary__c":   False,   # "Primary"
    "Target_FR__c":             False,   # "Target? (First Round)"
    "Target_Second_Round__c":   False,   # "Target? (Second Round)"
    "Communications_Closed__c": False,   # "Communications Closed?"
}

# --- Computed Name -----------------------------------------------------------
# Name = "<LenderName> - <Property>". LenderName = account name with leading
# numeric tokens removed ("3650 REIT" -> "REIT"; "A10 Capital" stays "A10 Capital").
NAME_FIELD = "Name"
NAME_JOINER = " - "

# --- Lookup matching ---------------------------------------------------------
CONTACT_COLUMN = "Contact"
ACCOUNT_COLUMN = "Account"
# When several Accounts share a name, prefer the one whose Type equals this.
ACCOUNT_PREFER_TYPE = "Capital Source"

# --- Typo tolerance ----------------------------------------------------------
# When a typed Account/Contact name doesn't match exactly, compare it to similar
# Salesforce records (0..1 similarity):
#   >= AUTO  -> treat as a typo and auto-correct (printed in the dry-run to veto)
#   >= SUGGEST (but < AUTO) -> don't guess; show "did you mean X?" for confirmation
FUZZY_AUTO_THRESHOLD = 0.90
FUZZY_SUGGEST_THRESHOLD = 0.70

# --- Interest / Declined handling -------------------------------------------
INTEREST_COLUMN = "Interest"
# Declined rows now imported; their required reason defaults to this picklist value.
DECLINED_INTEREST_VALUE = "Declined"
DECLINED_REASON_FIELD = "ascendix__DeclinedReason__c"
DECLINED_REASON_DEFAULT = "Unknown"
