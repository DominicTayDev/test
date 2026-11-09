import csv
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

DATA_DIR = sys.argv[1]
OUT = sys.argv[2]

FONT = "Arial"
F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_HDR = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=14, bold=True)
F_INPUT = Font(name=FONT, size=10, color="0000FF")
F_LINK = Font(name=FONT, size=10, color="008000")
F_EX = Font(name=FONT, size=10, italic=True, color="808080")
F_RED = Font(name=FONT, size=10, bold=True, color="C00000")
FILL_HDR = PatternFill("solid", fgColor="1F3864")
FILL_IN = PatternFill("solid", fgColor="FFFF00")
FILL_GREY = PatternFill("solid", fgColor="F2F2F2")
FILL_SEC = PatternFill("solid", fgColor="D9E1F2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
PCT = "0.00%;(0.00%);-"
SGD = '"S$"#,##0.00;("S$"#,##0.00);-'

wb = Workbook()

def hdr(ws, row, headers, widths=None):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = F_HDR
        c.fill = FILL_HDR
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = BORDER
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

# ---------------------------------------------------------------- Rate keys
# (key, scheme, product group, tiers included, card type, kind)
# kind: icpp = interchange++ row; max = conservative MAX of other rows; flat = single MDR
RATE_KEYS = [
    ("V-CR-CLASSIC", "Visa", "Consumer credit - Classic / Gold", "Visa Classic, Visa Gold", "Consumer credit", "icpp"),
    ("V-CR-PLAT", "Visa", "Consumer credit - Platinum", "Visa Platinum", "Consumer credit", "icpp"),
    ("V-CR-SIG", "Visa", "Consumer credit - Signature", "Visa Signature", "Consumer credit", "icpp"),
    ("V-CR-INF", "Visa", "Consumer credit - Infinite", "Visa Infinite, Visa Infinite Privilege, Visa Infinite Private", "Consumer credit", "icpp"),
    ("V-CR-UNK", "Visa", "Consumer credit - tier unknown", "Assumes the HIGHEST Visa consumer credit rate above (conservative)", "Consumer credit", "max:V-CR-CLASSIC:V-CR-INF"),
    ("V-DB-STD", "Visa", "Consumer debit - Classic", "Visa Debit / Visa Classic Debit", "Consumer debit", "icpp"),
    ("V-DB-PREM", "Visa", "Consumer debit - Platinum / premium", "Visa Platinum Debit, Visa Signature/Infinite Debit", "Consumer debit", "icpp"),
    ("V-DB-UNK", "Visa", "Consumer debit - tier unknown", "Assumes the HIGHER Visa debit rate above (conservative)", "Consumer debit", "max:V-DB-STD:V-DB-PREM"),
    ("V-PP", "Visa", "Prepaid", "Visa Prepaid", "Prepaid", "icpp"),
    ("V-BIZ-CR", "Visa", "Business credit", "Visa Business, Visa Business Platinum, Visa Signature Business", "Commercial credit", "icpp"),
    ("V-BIZ-DB", "Visa", "Business debit", "Visa Business Debit, Visa Business Platinum Debit", "Commercial debit", "icpp"),
    ("V-CORP", "Visa", "Corporate / Purchasing", "Visa Corporate, Visa Purchasing", "Commercial credit", "icpp"),
    ("M-CR-STD", "Mastercard", "Consumer credit - Standard / Titanium / Gold", "Standard, Titanium, Gold Mastercard", "Consumer credit", "icpp"),
    ("M-CR-PLAT", "Mastercard", "Consumer credit - Platinum", "Platinum Mastercard", "Consumer credit", "icpp"),
    ("M-CR-WORLD", "Mastercard", "Consumer credit - World", "World Mastercard (e.g. UOB PRVI Miles Mastercard)", "Consumer credit", "icpp"),
    ("M-CR-WSEL", "Mastercard", "Consumer credit - World Select", "World Select Mastercard (new tier, e.g. Citi PremierMiles)", "Consumer credit", "icpp"),
    ("M-CR-WE", "Mastercard", "Consumer credit - World Elite", "World Elite Mastercard", "Consumer credit", "icpp"),
    ("M-CR-WL", "Mastercard", "Consumer credit - World Legend", "World Legend Mastercard (new tier, e.g. Citi ULTIMA from 17 Nov 2026)", "Consumer credit", "icpp"),
    ("M-CR-UNK", "Mastercard", "Consumer credit - tier unknown", "Assumes the HIGHEST Mastercard consumer credit rate above (conservative)", "Consumer credit", "max:M-CR-STD:M-CR-WL"),
    ("M-DB-STD", "Mastercard", "Consumer debit - Standard", "Debit Mastercard", "Consumer debit", "icpp"),
    ("M-DB-PREM", "Mastercard", "Consumer debit - Platinum / World / World Elite", "Platinum, World, World Elite Debit Mastercard", "Consumer debit", "icpp"),
    ("M-DB-UNK", "Mastercard", "Consumer debit - tier unknown", "Assumes the HIGHER Mastercard debit rate above (conservative)", "Consumer debit", "max:M-DB-STD:M-DB-PREM"),
    ("M-PP", "Mastercard", "Prepaid", "Prepaid Mastercard", "Prepaid", "icpp"),
    ("M-BIZ-CR", "Mastercard", "Business credit", "Mastercard Business, Platinum Business", "Commercial credit", "icpp"),
    ("M-BIZ-PREM", "Mastercard", "Business credit - World / World Elite", "World Mastercard for Business, World Elite for Business", "Commercial credit", "icpp"),
    ("M-BIZ-DB", "Mastercard", "Business debit", "Business Debit Mastercard, World Business Debit", "Commercial debit", "icpp"),
    ("M-CORP", "Mastercard", "Corporate / Purchasing", "Mastercard Corporate, World Corporate, Purchasing", "Commercial credit", "icpp"),
    ("AMEX", "Amex", "American Express (all)", "Amex runs its own network: no interchange. Enter the all-in MDR Amex/your acquirer charges.", "All", "flat"),
    ("UPI", "UnionPay", "UnionPay (all)", "Enter the all-in MDR, or IC + scheme fee if your acquirer splits it.", "All", "flat"),
    ("NETS", "NETS", "NETS (domestic debit)", "Domestic NETS network: priced as a NETS MDR, not Visa/MC interchange.", "Consumer debit", "flat"),
    ("JCB", "JCB", "JCB (all)", "Enter the all-in MDR.", "All", "flat"),
    ("UNK", "Unknown", "Scheme / tier unknown - check BIN", "Do not price from this row: identify the card by BIN first.", "Unknown", "flat"),
]

# ---------------------------------------------------------------- mapping
def rate_key(scheme, tier, ctype):
    s = scheme.lower(); t = tier.lower(); c = ctype.lower()
    if s.startswith("amex"):
        return "AMEX", "Amex network"
    if s.startswith("unionpay"):
        return "UPI", "UnionPay network"
    if s.startswith("nets"):
        return "NETS", "NETS network"
    if s.startswith("jcb"):
        return "JCB", "JCB network"
    if "/" in s or s == "unknown":
        return "UNK", "Scheme not confirmed - check BIN"
    unknown = t.startswith("unknown")
    commercial = c.startswith("commercial")
    debit = "debit" in c or (not commercial and "debit" in t and "credit" not in c)
    prepaid = c.startswith("prepaid")
    if s == "visa":
        if prepaid: return "V-PP", "Card type"
        if commercial:
            if "corporate" in t or "purchasing" in t: return "V-CORP", "Stated tier"
            if "debit" in c: return "V-BIZ-DB", "Stated tier" if not unknown else "Card type (tier unknown)"
            return "V-BIZ-CR", "Stated tier" if not unknown else "Card type (tier unknown)"
        if debit:
            if unknown: return "V-DB-UNK", "Tier unknown - conservative (verify by BIN)"
            if any(k in t for k in ("platinum", "signature", "infinite")): return "V-DB-PREM", "Stated tier"
            return "V-DB-STD", "Stated tier"
        if unknown: return "V-CR-UNK", "Tier unknown - conservative (verify by BIN)"
        if " or " in t: return "V-CR-INF", "Tier varies by BIN - priced at the higher tier"
        if "infinite" in t: return "V-CR-INF", "Stated tier"
        if "signature" in t: return "V-CR-SIG", "Stated tier"
        if "platinum" in t: return "V-CR-PLAT", "Stated tier"
        return "V-CR-CLASSIC", "Stated tier"
    if s == "mastercard":
        if prepaid: return "M-PP", "Card type"
        if commercial:
            if "corporate" in t or "purchasing" in t: return "M-CORP", "Stated tier"
            if "debit" in c: return "M-BIZ-DB", "Stated tier" if not unknown else "Card type (tier unknown)"
            if "world" in t: return "M-BIZ-PREM", "Stated tier"
            return "M-BIZ-CR", "Stated tier" if not unknown else "Card type (tier unknown)"
        if debit:
            if unknown: return "M-DB-UNK", "Tier unknown - conservative (verify by BIN)"
            if any(k in t for k in ("platinum", "world", "elite")): return "M-DB-PREM", "Stated tier"
            return "M-DB-STD", "Stated tier"
        if unknown: return "M-CR-UNK", "Tier unknown - conservative (verify by BIN)"
        if "legend" in t: return "M-CR-WL", "Stated tier"
        if "elite" in t: return "M-CR-WE", "Stated tier"
        if "select" in t: return "M-CR-WSEL", "Stated tier"
        if "world" in t: return "M-CR-WORLD", "Stated tier"
        if "platinum" in t: return "M-CR-PLAT", "Stated tier"
        return "M-CR-STD", "Stated tier"
    return "UNK", "Unmapped"

def status(notes):
    n = notes.upper()
    if "DISCONTINUED" in n: return "Discontinued (old cards may still be in use)"
    if "SUSPENDED" in n: return "Applications suspended"
    if "BEING REPLACED" in n: return "Being replaced on renewal"
    if "CLOSED" in n: return "Closed to new applicants (existing cards in use)"
    if "PROBABLY DISCONTINUED" in n or "PROBABLY LEGACY" in n or "LIKELY LEGACY" in n or "PROBABLY CLOSED" in n or "PROBABLY REPLACED" in n:
        return "Possibly legacy"
    return "Active"

rows = []
for fn in ("catalog_big3.csv", "catalog_others.csv"):
    with open(f"{DATA_DIR}/{fn}", newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            rows.append(r)

bank_order = ["DBS", "POSB", "OCBC", "UOB", "Citibank Singapore", "Citibank N.A. (institutional)",
              "HSBC Singapore", "Standard Chartered Singapore", "Maybank Singapore", "CIMB Singapore",
              "Bank of China Singapore", "ICBC Singapore", "Trust Bank", "GXS Bank", "MariBank",
              "American Express (direct)", "YouTrip", "Revolut SG", "Wise SG", "Instarem", "Chocolate Finance"]
rows.sort(key=lambda r: (bank_order.index(r["Bank"]) if r["Bank"] in bank_order else 99))

# ================================================================ Sheet 1: Read Me
ws = wb.active
ws.title = "Read Me"
ws.column_dimensions["A"].width = 26
ws.column_dimensions["B"].width = 110
ws["A1"] = "Singapore Card Interchange++ Workbook"; ws["A1"].font = F_TITLE
ws["A2"] = "Prepared 9 Nov 2026 for Warely. Interchange++ = interchange (to the card-issuing bank) + scheme fee (Visa / Mastercard) + your markup."
ws["A2"].font = F_BASE
ws["A4"] = "STATUS"; ws["A4"].font = F_RED
ws["B4"] = ("RATES NOT YET FILLED. The rate source (your Google Doc) is private and could not be read (Google returned 401 sign-in "
            "required). Visa and Mastercard do not publish Singapore interchange tables, so no rate in this file has been "
            "estimated or invented. Every 'Pending' cell turns into a number once the yellow cells on 'Rate Table' are filled from your doc.")
ws["B4"].font = F_RED; ws["B4"].alignment = WRAP
ws.row_dimensions[4].height = 54
info = [
    ("Key point", "Interchange is set by the scheme per card TIER (e.g. World vs World Elite), card type (credit/debit/commercial), channel "
                  "(card-present vs online) and card origin (Singapore vs foreign). It is NOT set per bank. The bank/card list only tells "
                  "you which tier each card falls in."),
    ("Sheets", "Rate Table: fill the yellow cells from your acquirer / scheme rate document. Each row is a Rate Key (scheme + tier group)."),
    ("", "Card Catalog: every Singapore-issued card found (bank, card, scheme, tier), mapped to a Rate Key, with its IC++ pulled from Rate Table."),
    ("", "Merchant Pricing: enter a merchant's card mix to get their blended IC++ cost and compare it with the flat rate you plan to quote."),
    ("", "Card Data Capture: the per-transaction fields your platform must store to know each card's real interchange."),
    ("", "Gaps & Sources: what is unverified, recent tier changes, and how the catalog was built."),
    ("Colour legend", "Yellow fill + blue text = input cell you fill in.  Black text = formula, do not overwrite.  Green text = value pulled from another sheet."),
    ("", "Grey italic row on Rate Table = EXAMPLE FORMAT ONLY (dummy numbers, not real rates). It is not used by any formula."),
    ("Percentages", "Enter percentages as percentages (type 1.25% or 0.0125). Fixed fees are in S$ per transaction."),
    ("Confidence (catalog)", "Certain = the bank's own page/document states the tier.  Likely = reputable aggregator (MileLion, SingSaver, MoneySmart) "
                             "states it.  Guessing = inferred.  '(scheme)' = only the scheme is confirmed, tier unknown."),
    ("Unknown tiers", "Cards with an unknown tier are mapped to a '...-UNK' Rate Key that automatically uses the HIGHEST rate in that "
                      "scheme's group, so you never under-price. Replace with the real tier once you see the card's BIN / product ID."),
    ("Closed cards", "Cards closed to new applicants are kept: existing cards still transact at your merchants."),
    ("Not covered", "Foreign-issued cards are not listed individually: price them with the intra-regional / inter-regional columns on Rate Table."),
]
r = 6
for k, v in info:
    ws.cell(row=r, column=1, value=k).font = F_BOLD
    c = ws.cell(row=r, column=2, value=v); c.font = F_BASE; c.alignment = WRAP
    ws.row_dimensions[r].height = 30 if len(v) > 110 else 16
    r += 1

# ================================================================ Sheet 2: Rate Table
rt = wb.create_sheet("Rate Table")
rt["A1"] = "Rate Table - fill the yellow cells from your rate document"; rt["A1"].font = F_TITLE
rt["A2"] = "Default markup %"; rt["A2"].font = F_BOLD
rt["C2"] = None; rt["C2"].fill = FILL_IN; rt["C2"].font = F_INPUT; rt["C2"].number_format = PCT
rt["D2"] = "Default markup fixed (S$ per txn)"; rt["D2"].font = F_BOLD
rt["F2"] = None; rt["F2"].fill = FILL_IN; rt["F2"].font = F_INPUT; rt["F2"].number_format = SGD
rt["G2"] = "Used for every row unless you enter a row-level markup in columns N / O."; rt["G2"].font = F_BASE
rt["C2"].comment = Comment("Your margin on top of interchange + scheme fee. Input from Warely.", "Claude")

headers = ["Rate Key", "Scheme", "Product group", "Tiers included / how to fill",
           "Domestic card-present IC %", "Domestic card-present IC fixed (S$)",
           "Domestic online (CNP) IC %", "Domestic online (CNP) IC fixed (S$)",
           "Foreign card - Asia/Pacific (intra-regional) IC %", "Foreign card - outside AP (inter-regional) IC %",
           "Scheme fee % - domestic", "Scheme fee % - cross-border",
           "Row markup % (blank = default)", "Row markup fixed S$ (blank = default)",
           "IC++ % - domestic card-present", "IC++ % - domestic online",
           "IC++ % - foreign AP card", "IC++ % - foreign non-AP card",
           "IC++ fixed S$ per txn (card-present)", "IC++ fixed S$ per txn (online)",
           "Where in your doc (page / table)", "Notes"]
HROW = 6
hdr(rt, HROW, headers, [14, 11, 34, 46, 12, 12, 12, 12, 13, 13, 11, 11, 11, 11, 12, 12, 12, 12, 12, 12, 22, 30])
rt.row_dimensions[HROW].height = 66

# example row (row 5)
ex = ["EXAMPLE ONLY", "Mastercard", "Format example - dummy numbers, NOT real rates", "Not used by any formula",
      0.0150, 0.10, 0.0180, 0.10, 0.0200, 0.0250, 0.0010, 0.0060, None, None, None, None, None, None, None, None, "p.3, Table 2", ""]
for i, v in enumerate(ex, 1):
    c = rt.cell(row=5, column=i, value=v); c.font = F_EX; c.fill = FILL_GREY; c.border = BORDER
    if i in (5, 7, 9, 10, 11, 12): c.number_format = PCT
    if i in (6, 8): c.number_format = SGD
rt.cell(row=4, column=1, value="Expected input format (grey row is an example only):").font = F_EX

first = HROW + 1
key_row = {}
for idx, rk in enumerate(RATE_KEYS):
    key_row[rk[0]] = first + idx
last = first + len(RATE_KEYS) - 1

for idx, (key, scheme, group, tiers, ctype, kind) in enumerate(RATE_KEYS):
    r = first + idx
    for col, v in enumerate([key, scheme, group, tiers], 1):
        c = rt.cell(row=r, column=col, value=v); c.font = F_BOLD if col == 1 else F_BASE; c.border = BORDER
        c.alignment = Alignment(wrap_text=True, vertical="top")
    if kind.startswith("max:"):
        _, a, b = kind.split(":")
        ra, rb = key_row[a], key_row[b]
        for col in range(5, 13):
            L = get_column_letter(col)
            c = rt.cell(row=r, column=col, value=f'=IF(COUNT({L}{ra}:{L}{rb})=0,"",MAX({L}{ra}:{L}{rb}))')
            c.font = F_BASE; c.border = BORDER; c.fill = FILL_GREY
            c.number_format = SGD if col in (6, 8) else PCT
    else:
        for col in range(5, 13):
            c = rt.cell(row=r, column=col); c.fill = FILL_IN; c.font = F_INPUT; c.border = BORDER
            c.number_format = SGD if col in (6, 8) else PCT
    for col in (13, 14):
        c = rt.cell(row=r, column=col); c.fill = FILL_IN; c.font = F_INPUT; c.border = BORDER
        c.number_format = PCT if col == 13 else SGD
    mk = f'IF(M{r}="",$C$2,M{r})'
    mf = f'IF(N{r}="",$F$2,N{r})'
    if kind == "flat":
        # single all-in MDR: IC/scheme columns optional; markup or IC required
        f_cp = f'=IF(AND(E{r}="",M{r}=""),"Pending",SUM(E{r},K{r})+{mk})'
        f_cnp = f'=IF(AND(G{r}="",M{r}=""),"Pending",SUM(G{r},K{r})+{mk})'
        f_ap = f'=IF(AND(I{r}="",M{r}=""),"Pending",SUM(I{r},L{r})+{mk})'
        f_ia = f'=IF(AND(J{r}="",M{r}=""),"Pending",SUM(J{r},L{r})+{mk})'
    else:
        f_cp = f'=IF(OR(E{r}="",K{r}="",AND(M{r}="",$C$2="")),"Pending",E{r}+K{r}+{mk})'
        f_cnp = f'=IF(OR(G{r}="",K{r}="",AND(M{r}="",$C$2="")),"Pending",G{r}+K{r}+{mk})'
        f_ap = f'=IF(OR(I{r}="",L{r}="",AND(M{r}="",$C$2="")),"Pending",I{r}+L{r}+{mk})'
        f_ia = f'=IF(OR(J{r}="",L{r}="",AND(M{r}="",$C$2="")),"Pending",J{r}+L{r}+{mk})'
    f_fx_cp = f'=IF(ISTEXT(O{r}),"Pending",SUM(F{r})+{mf})'
    f_fx_cnp = f'=IF(ISTEXT(P{r}),"Pending",SUM(H{r})+{mf})'
    for col, f in zip(range(15, 21), (f_cp, f_cnp, f_ap, f_ia, f_fx_cp, f_fx_cnp)):
        c = rt.cell(row=r, column=col, value=f); c.font = F_BASE; c.border = BORDER
        c.number_format = SGD if col >= 19 else PCT
    for col in (21, 22):
        c = rt.cell(row=r, column=col); c.border = BORDER; c.alignment = WRAP
        if col == 21: c.fill = FILL_IN; c.font = F_INPUT
    if key == "M-CR-WORLD":
        rt.cell(row=r, column=22, value="UOB PRVI Miles Mastercard sits here (World Mastercard).").font = F_BASE
    if key in ("M-CR-WSEL", "M-CR-WL"):
        rt.cell(row=r, column=22, value="New Mastercard tier: your doc may not list it yet - ask your acquirer which rate applies.").font = F_BASE
    if key == "V-CR-INF":
        rt.cell(row=r, column=22, value="UOB is moving Visa Infinite holders to Visa Infinite Privilege / Private: confirm if your doc rates them separately.").font = F_BASE

rt.freeze_panes = rt.cell(row=first, column=5)
RT_LAST = last

# ================================================================ Sheet 3: Card Catalog
cc = wb.create_sheet("Card Catalog")
cc["A1"] = "Singapore-issued cards mapped to Rate Keys (domestic rates apply when these cards pay a Singapore merchant)"; cc["A1"].font = F_TITLE
cc["A2"] = f"{len(rows)} cards. Built from public bank pages and aggregators (search snippets), Nov 2026. Filter by Bank / Scheme / Rate Key with the header arrows."
cc["A2"].font = F_BASE
ch = ["#", "Bank", "Card", "Scheme", "Tier (as branded)", "Card type", "Status", "Rate Key", "How the Rate Key was set",
      "IC++ % domestic card-present", "IC++ % domestic online", "IC++ fixed S$ (card-present)", "Confidence", "Notes", "Source"]
CH = 4
hdr(cc, CH, ch, [5, 22, 44, 12, 34, 16, 26, 13, 30, 13, 13, 12, 14, 52, 50])
cc.row_dimensions[CH].height = 40
rng_key = f"'Rate Table'!$A${first}:$A${RT_LAST}"
def look(col_letter, r):
    return (f"=IFERROR(INDEX('Rate Table'!${col_letter}${first}:${col_letter}${RT_LAST},"
            f"MATCH($H{r},{rng_key},0)),\"Unmapped\")")
for i, rw in enumerate(rows, 1):
    r = CH + i
    key, basis = rate_key(rw["Scheme"], rw["Tier"], rw["CardType"])
    vals = [i, rw["Bank"], rw["Card"], rw["Scheme"], rw["Tier"], rw["CardType"], status(rw["Notes"]), key, basis]
    for col, v in enumerate(vals, 1):
        c = cc.cell(row=r, column=col, value=v); c.font = F_BASE; c.border = BORDER
        c.alignment = Alignment(vertical="top", wrap_text=col in (3, 5, 9))
    for col, L, fmt in ((10, "O", PCT), (11, "P", PCT), (12, "S", SGD)):
        c = cc.cell(row=r, column=col, value=look(L, r)); c.font = F_LINK; c.border = BORDER; c.number_format = fmt
    for col, v in ((13, rw["Confidence"]), (14, rw["Notes"]), (15, rw["SourceURL"])):
        c = cc.cell(row=r, column=col, value=v); c.font = F_BASE; c.border = BORDER
        c.alignment = Alignment(vertical="top", wrap_text=col == 14)
    if "PRVI Miles World Mastercard" in rw["Card"]:
        for col in range(1, 16):
            cc.cell(row=r, column=col).fill = PatternFill("solid", fgColor="E2EFDA")
CC_LAST = CH + len(rows)
cc.auto_filter.ref = f"A{CH}:O{CC_LAST}"
cc.freeze_panes = cc.cell(row=CH + 1, column=4)

# ================================================================ Sheet 4: Merchant Pricing
mp = wb.create_sheet("Merchant Pricing")
mp["A1"] = "Merchant Pricing - blended IC++ cost for one merchant's card mix"; mp["A1"].font = F_TITLE
mp.column_dimensions["A"].width = 34
for L, w in zip("BCDEFGH", (22, 14, 16, 16, 16, 18, 40)):
    mp.column_dimensions[L].width = w
inputs = [
    ("Merchant name", None, None),
    ("Monthly card volume (S$)", None, SGD),
    ("Average ticket (S$)", None, SGD),
    ("Channel", "Card present", None),
    ("Flat MDR you plan to quote", None, PCT),
]
for i, (lab, v, fmt) in enumerate(inputs, 3):
    mp.cell(row=i, column=1, value=lab).font = F_BOLD
    c = mp.cell(row=i, column=2, value=v); c.fill = FILL_IN; c.font = F_INPUT; c.border = BORDER
    if fmt: c.number_format = fmt
dv_ch = DataValidation(type="list", formula1='"Card present,Online"', allow_blank=False)
mp.add_data_validation(dv_ch); dv_ch.add("B6")
mp["C6"] = "Pick: Card present / Online"; mp["C6"].font = F_EX

MH = 10
mp.cell(row=MH - 1, column=1, value="Card mix (get this from the merchant's last 3 months of acquirer statements)").font = F_BOLD
hdr(mp, MH, ["Rate Key (pick)", "Card origin", "% of volume", "IC++ %", "IC++ fixed S$/txn", "Effective % incl. fixed", "Monthly cost (S$)", "Product group"])
mp.column_dimensions["A"].width = 34
dv_key = DataValidation(type="list", formula1=f"='Rate Table'!$A${first}:$A${RT_LAST}", allow_blank=True)
dv_org = DataValidation(type="list", formula1='"Singapore,Foreign - Asia/Pacific,Foreign - outside AP"', allow_blank=True)
mp.add_data_validation(dv_key); mp.add_data_validation(dv_org)
NMIX = 15
m1 = MH + 1; mN = MH + NMIX
for k in range(NMIX):
    r = m1 + k
    for col in (1, 2, 3):
        c = mp.cell(row=r, column=col); c.fill = FILL_IN; c.font = F_INPUT; c.border = BORDER
    mp.cell(row=r, column=3).number_format = PCT
    dv_key.add(f"A{r}"); dv_org.add(f"B{r}")
    idx = f"MATCH($A{r},'Rate Table'!$A${first}:$A${RT_LAST},0)"
    def col_(L):
        return f"INDEX('Rate Table'!${L}${first}:${L}${RT_LAST},{idx})"
    f_rate = (f'=IF($A{r}="","",IFERROR(IF($B{r}="Foreign - Asia/Pacific",{col_("Q")},'
              f'IF($B{r}="Foreign - outside AP",{col_("R")},IF($B$6="Online",{col_("P")},{col_("O")}))),"Unmapped"))')
    f_fix = f'=IF($A{r}="","",IFERROR(IF($B$6="Online",{col_("T")},{col_("S")}),0))'  # foreign cards use the domestic fixed fee
    f_eff = f'=IF(OR($A{r}="",ISTEXT(D{r})),D{r},D{r}+IF(N($B$5)>0,N(E{r})/$B$5,0))'
    f_cost = f'=IF(OR($A{r}="",ISTEXT(F{r})),"",F{r}*C{r}*N($B$4))'
    f_grp = f'=IF($A{r}="","",IFERROR({col_("C")},""))'
    for col, f, fmt in ((4, f_rate, PCT), (5, f_fix, SGD), (6, f_eff, PCT), (7, f_cost, SGD), (8, f_grp, None)):
        c = mp.cell(row=r, column=col, value=f); c.font = F_LINK if col in (4, 5, 8) else F_BASE; c.border = BORDER
        if fmt: c.number_format = fmt
tr = mN + 2
summary = [
    ("Mix entered (should be 100%)", f"=SUM(C{m1}:C{mN})", PCT),
    ("Blended IC++ cost %", f'=IF(SUMPRODUCT((C{m1}:C{mN}>0)*ISTEXT(F{m1}:F{mN}))>0,"Pending - fill Rate Table",IF(SUM(C{m1}:C{mN})=0,"",SUMPRODUCT(C{m1}:C{mN},F{m1}:F{mN})/SUM(C{m1}:C{mN})))', PCT),
    ("Blended monthly cost (S$)", f'=IF(ISTEXT(B{tr+1}),B{tr+1},IF(B{tr+1}="","",B{tr+1}*N(B4)))', SGD),
    ("Your margin at the flat MDR quoted (%)", f'=IF(OR(ISTEXT(B{tr+1}),B{tr+1}="",B7=""),"Needs rates + quote",B7-B{tr+1})', PCT),
    ("Your margin at the flat MDR quoted (S$/month)", f'=IF(ISTEXT(B{tr+3}),B{tr+3},B{tr+3}*N(B4))', SGD),
]
for i, (lab, f, fmt) in enumerate(summary):
    r = tr + i
    mp.cell(row=r, column=1, value=lab).font = F_BOLD
    c = mp.cell(row=r, column=2, value=f); c.font = F_BOLD; c.number_format = fmt; c.border = BORDER
mp.cell(row=tr + 6, column=1, value=("A negative margin means the flat rate you quote is below your blended cost for this merchant's card mix. "
                                      "Foreign and commercial cards usually drive this.")).font = F_EX

# ================================================================ Sheet 5: Card Data Capture
dc = wb.create_sheet("Card Data Capture")
dc["A1"] = "Fields to store per transaction to know each card's real interchange"; dc["A1"].font = F_TITLE
dc["A2"] = ("Layer 1 = estimate at payment time (from the card and the approval message). Layer 2 = actual cost (from the acquirer's "
            "settlement file). Layer 3 = compare them and flag downgrades. Field names differ by acquirer: confirm against their API / file spec.")
dc["A2"].font = F_BASE; dc["A2"].alignment = WRAP; dc.merge_cells("A2:F2"); dc.row_dimensions[2].height = 42
hdr(dc, 4, ["Layer", "Field", "Where it comes from", "Why it matters for interchange", "Confidence", "Notes"], [10, 34, 40, 52, 11, 46])
fields = [
    ("1", "BIN (first 8 digits) + last 4", "Card / terminal / gateway token response", "Identifies issuer, scheme, card range; key to look up tier", "Likely", "PCI DSS allows storing first 8 + last 4 for 16-digit cards. Never store the full card number."),
    ("1", "Scheme / network", "Gateway response or EMV AID", "Visa vs Mastercard vs Amex vs UnionPay vs NETS", "Certain", "NETS co-badged debit cards can route via NETS or Visa/MC: record which network was actually used."),
    ("1", "Issuer country", "BIN table / gateway response", "Domestic vs intra-AP vs inter-regional rate column", "Certain", "Foreign cards cost more: biggest pricing risk for tourist-facing merchants."),
    ("1", "Funding type (credit / debit / prepaid)", "BIN table / gateway response", "Debit rates are lower than credit", "Certain", ""),
    ("1", "Consumer vs commercial flag", "BIN table / product ID", "Commercial cards carry higher interchange", "Likely", ""),
    ("1", "Product ID / tier code", "Approval response: Mastercard 'Financial Network Code'; Visa 'Product ID' (card level result)", "Separates World vs World Elite, Signature vs Infinite", "Likely", "Ask your acquirer whether their API passes these through. If not, use the scheme BIN files via the acquirer."),
    ("1", "Entry mode", "Terminal: chip / contactless / magstripe / manual key-in / e-commerce", "Card-present vs online rate; manual key-in often downgrades", "Likely", ""),
    ("1", "3-D Secure result (online)", "Gateway / 3DS server", "Online without 3DS can land in a higher category", "Likely", ""),
    ("1", "Merchant category code (MCC)", "Merchant set-up at acquirer", "Some categories have special rates", "Likely", "F&B is typically MCC 5812 (restaurants) / 5814 (fast food)."),
    ("1", "Auth date/time and settlement (batch close) date/time", "POS / gateway", "Late settlement can downgrade the rate", "Likely", "Close batches daily."),
    ("2", "Actual interchange category", "Acquirer settlement file: Mastercard IRD (Interchange Rate Designator); Visa Fee Program Indicator", "The category the transaction actually landed in", "Likely", "Make a transaction-level file a contract requirement. Without it IC++ billing is an estimate."),
    ("2", "Actual interchange amount", "Acquirer settlement file", "True cost to bill IC++", "Likely", ""),
    ("2", "Scheme fees charged", "Acquirer settlement / billing file", "The first '+' in IC++", "Likely", ""),
    ("3", "Expected vs actual Rate Key", "Your platform (compare Layer 1 vs Layer 2)", "Flags downgrades and mis-mapped BINs", "Certain", "Use mismatches to update your BIN -> Rate Key table over time."),
]
for i, fr in enumerate(fields, 5):
    for col, v in enumerate(fr, 1):
        c = dc.cell(row=i, column=col, value=v); c.font = F_BASE; c.border = BORDER; c.alignment = WRAP

# ================================================================ Sheet 6: Gaps & Sources
gs = wb.create_sheet("Gaps & Sources")
gs.column_dimensions["A"].width = 30; gs.column_dimensions["B"].width = 120
gs["A1"] = "Gaps, recent tier changes and sources"; gs["A1"].font = F_TITLE
notes = [
    ("Rate source", "Your Google Doc (1-qfEHIGurq-RdDI-5vtsnTUphx-AFK8VPaNTu5_EfH8) could not be read: private (HTTP 401). No rates filled."),
    ("Public rates", "Visa Singapore's 'Rules + Interchange' page publishes no rate table (visa.com.sg/support/small-business/regulations-fees.html). "
                     "Mastercard publishes country tables for regulated markets such as Australia, not Singapore. Singapore does not cap interchange."),
    ("Catalog method", "Two research passes over public search results (bank pages, MileLion, SingSaver, MoneySmart, Wise, Aspire), Nov 2026. "
                       "Bank websites could not be opened directly, so tiers come from page titles/snippets. Check by BIN before billing."),
    ("Tier changes", "HSBC Revolution: Visa Platinum -> Visa Signature (1 Apr 2026). HSBC Premier Mastercard -> World Elite (1 Oct 2024)."),
    ("", "Citi PremierMiles -> World Select Mastercard (~Dec 2025). Citi ULTIMA -> World Legend Mastercard (17 Nov 2026)."),
    ("", "UOB is moving 300k+ Visa Infinite cardholders to Visa Infinite Privilege / Private (notices from Sep 2026): which card gets which tier is not published."),
    ("", "UOB PRVI Miles Visa can be Visa Signature or Visa Infinite depending on card number range; priced here at the higher tier."),
    ("", "UOB Preferred Platinum Visa renamed UOB Preferred Visa (10 Mar 2026), reported as Visa Signature."),
    ("Bank changes", "Citibank still issues consumer cards in Singapore (UOB bought Citi's consumer business in MY/TH/ID/VN, not SG). Citi SG commercial cards ended 30 Jun 2026."),
    ("", "Standard Chartered moved selected card customers to Trust Bank (from Sep/Oct 2026); exact products not confirmed. SC Smart card stopped new applications 20 Oct 2026."),
    ("", "Trust Freedom Card reportedly 'codes as debit' in some situations - check its BIN."),
    ("Tier unknown", "POSB Everyday, OCBC 90N Visa, OCBC INFINITY, OCBC NXT, OCBC Plus! Visa, KrisFlyer UOB credit, UOB EVOL, most UOB/OCBC/Citi debit cards, "
                     "Maybank Family & Friends / XL Cashback, CIMB AWSM / Founders, GXS FlexiCard, Mari cards, prepaid/fintech cards."),
    ("Not researched", "Bank of Singapore, DBS Private Bank, OCBC legacy cards, Aspire, Airwallex, Revolut/Wise Business, YouBiz, Atome Card, HeyMax, Singtel Dash."),
    ("Amex", "Amex is a three-party network: no Visa/MC-style interchange. Amex cards issued by DBS/UOB/OCBC are still priced on the AMEX row."),
    ("NETS", "Many debit cards are NETS co-badged (DBS Visa Debit, OCBC Debit, UOB One Debit, UOB Lady's Debit, UOB Visa Debit, OCBC Business Debit, likely HSBC Everyday Global). "
             "Whether a tap goes via NETS or Visa/MC depends on terminal routing - record it per transaction."),
    ("Per-card sources", "See the Source column on Card Catalog."),
]
for i, (k, v) in enumerate(notes, 3):
    gs.cell(row=i, column=1, value=k).font = F_BOLD
    c = gs.cell(row=i, column=2, value=v); c.font = F_BASE; c.alignment = WRAP
    gs.row_dimensions[i].height = 30 if len(v) > 120 else 16

# base font everywhere not already set
for sh in wb.worksheets:
    for row in sh.iter_rows():
        for c in row:
            if c.font is None or c.font.name != FONT:
                c.font = Font(name=FONT, size=c.font.size if c.font and c.font.size else 10,
                              bold=c.font.bold if c.font else False, italic=c.font.italic if c.font else False,
                              color=c.font.color if c.font else None)

wb.save(OUT)
print("rows", len(rows), "ratekeys", len(RATE_KEYS))
from collections import Counter
print(Counter(rate_key(r["Scheme"], r["Tier"], r["CardType"])[0] for r in rows))
