"""
Out-of-Inventory (OOI) Rate — EN-US / EN-PK / Pipeline Variants / Reference
===========================================================================

Reads the evaluation workbook and reports, for each system, the proportion
of output phone tokens that fall outside the 79-phone Pakistani English
phonetic inventory.

Tokenisation is identical to per_from_sheet.py: greedy longest-match,
zero-width joiners stripped, spaces removed. Punctuation is excluded from
both numerator and denominator, since it is not a phone.

Usage
-----
    python ooi_evaluation.py --excel "PakE_Evaluation_Sentences.xlsx"
"""

import argparse
import unicodedata

import openpyxl

COL_NO, COL_ID, COL_SENT = 1, 2, 3
COL_ENUS, COL_ENPK, COL_PIPE_US, COL_PIPE_PK, COL_REF = 5, 6, 7, 8, 9

# --------------------------------------------------------------------------- #
# The 79-phone Pakistani English inventory (membership test)
# --------------------------------------------------------------------------- #

INVENTORY_79 = set(unicodedata.normalize("NFC", p) for p in """
p b k ɡ m n ŋ f v s z ʃ ʒ h tʃ dʒ r j l ɹ ɾ
ɪ ʊ ə ʌ ɛ æ ɔ ɒ ɑː iː uː ɜː eː oː ɔː ɐ ɨ ᵻ i ɚ ɝ aɪ ɑɪ aʊ ɑʊ ɔɪ
t̪ d̪ t̪ʰ d̪ʰ ɽ ɖ ʈ ʈʰ ɖʰ ɽʰ ʔ x ɣ pʰ bʰ kʰ ɡʰ tʃʰ dʒʰ mʰ nʰ lʰ rʰ jʰ
ɑ̃ː ẽː ĩː õː ũː ɔ̃ː æ̃ː ɪ̃ ʊ̃ ə̃ ʋ
""".split())

# Longest-first segmentation list: the inventory plus the symbols espeak
# emits that the inventory deliberately excludes, so multi-character units
# segment the same way they do during PER scoring.
SEGMENT_LIST = sorted(
    INVENTORY_79 | {"t", "d", "θ", "ð", "w", "q", "ˈ", "ˌ", "ː"},
    key=len, reverse=True,
)

PUNCTUATION = set(";:,.!?—…")


def tokenize_ipa(ipa_string, inventory=SEGMENT_LIST):
    """Greedy longest-match IPA tokenisation. Unknown chars kept as-is."""
    tokens, i = [], 0
    s = unicodedata.normalize("NFC", ipa_string or "").replace(" ", "").replace("‍", "")
    while i < len(s):
        for phone in inventory:
            if s.startswith(phone, i):
                tokens.append(phone)
                i += len(phone)
                break
        else:
            tokens.append(s[i])
            i += 1
    return tokens


def ooi_rate(ipa_string):
    """Return (n_out_of_inventory, n_phone_tokens, offending_symbols)."""
    tokens = [t for t in tokenize_ipa(ipa_string) if t not in PUNCTUATION]
    bad = [t for t in tokens if t not in INVENTORY_79]
    return len(bad), len(tokens), bad


def read_sheet(excel_path, sheet_name=None):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[0]]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        cell_a = str(row[COL_NO - 1]).strip() if row[COL_NO - 1] else ""
        sent_id = str(row[COL_ID - 1]).strip() if len(row) > 1 and row[COL_ID - 1] else ""
        if cell_a.lower().startswith("section") or not sent_id:
            continue

        def get(col):
            return str(row[col - 1]).strip() if len(row) >= col and row[col - 1] else ""

        rows.append({
            "id": sent_id,
            "enus": get(COL_ENUS), "enpk": get(COL_ENPK),
            "pipe_us": get(COL_PIPE_US), "pipe_pk": get(COL_PIPE_PK),
            "ref": get(COL_REF),
        })
    return rows


def evaluate(excel_path, sheet_name=None):
    rows = [r for r in read_sheet(excel_path, sheet_name) if r["ref"]]
    systems = [("ref", "REFERENCE"), ("enus", "EN-US"), ("enpk", "EN-PK"),
               ("pipe_us", "PIPE-US"), ("pipe_pk", "PIPE-PK")]

    print(f"Scoring {len(rows)} sentences against the 79-phone inventory.\n")
    print(f"{'System':<12}{'OOI':>7}{'Phones':>9}{'Rate':>8}   Offending symbols")
    print("-" * 78)

    for key, label in systems:
        n_bad = n_tok = 0
        counts = {}
        for r in rows:
            bad, tok, symbols = ooi_rate(r[key])
            n_bad += bad
            n_tok += tok
            for s in symbols:
                counts[s] = counts.get(s, 0) + 1
        detail = " ".join(f"{s}:{n}" for s, n in
                          sorted(counts.items(), key=lambda x: -x[1])) or "none"
        print(f"{label:<12}{n_bad:>7}{n_tok:>9}{100*n_bad/n_tok:>7.1f}%   {detail}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Out-of-inventory rate per system.")
    ap.add_argument("--excel", required=True, help="Evaluation workbook (.xlsx)")
    ap.add_argument("--sheet", default=None, help="Sheet name (defaults to first sheet)")
    args = ap.parse_args()
    evaluate(args.excel, args.sheet)
