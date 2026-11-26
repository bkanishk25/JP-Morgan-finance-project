"""Build the JPMorgan Chase (JPM) bank financial model -> model/JPM_Financial_Model.xlsx

All calculations are live Excel formulas; Python only lays out inputs and formulas.
Run from the repository root:  python src/build_model.py
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L

FONT = "Arial"
BLUE, BLACK, GREEN, WHITE = "0000FF", "000000", "008000", "FFFFFF"
NAVY, LIGHT, YELLOW, GREY = "1F3864", "D9E1F2", "FFFF00", "F2F2F2"
USD = '$#,##0;($#,##0);"-"'
USD2 = '$#,##0.00;($#,##0.00);"-"'
NUM1 = '#,##0.0;(#,##0.0);"-"'
PCT = '0.0%;(0.0%);"-"'
PCT2 = '0.00%;(0.00%);"-"'
MULT = '0.00"x"'
thin = Side(style="thin", color="BFBFBF")

wb = Workbook()


def f(color=BLACK, bold=False, size=10, italic=False):
    return Font(name=FONT, color=color, bold=bold, size=size, italic=italic)


def setup(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[L(i)].width = w
    ws.sheet_view.showGridLines = False


def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = f(NAVY, True, 14)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = f("595959", italic=True, size=9)


def header(ws, row, labels, start_col=1):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i, value=lab)
        c.font = f(WHITE, True)
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center" if i else "left", vertical="center", wrap_text=True)


def section(ws, row, text, ncols):
    for col in range(1, ncols + 1):
        ws.cell(row=row, column=col).fill = PatternFill("solid", fgColor=LIGHT)
    ws.cell(row=row, column=1, value=text).font = f(NAVY, True)


def put(ws, ref, value, fmt=None, color=None, bold=False, fill=None, comment=None):
    c = ws[ref]
    c.value = value
    if color is None:
        if isinstance(value, str) and value.startswith("="):
            color = GREEN if "!" in value else BLACK
        elif isinstance(value, (int, float)):
            color = BLUE
        else:
            color = BLACK
    c.font = f(color, bold)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if comment:
        c.comment = Comment(comment, "Model")
    if isinstance(value, (int, float)) or (isinstance(value, str) and value.startswith("=")):
        c.alignment = Alignment(horizontal="right")
    return c


def label(ws, row, text, bold=False, indent=0, note=None, note_col=None):
    c = ws.cell(row=row, column=1, value=text)
    c.font = f(BLACK, bold)
    c.alignment = Alignment(indent=indent)
    if note and note_col:
        n = ws.cell(row=row, column=note_col, value=note)
        n.font = f("595959", italic=True, size=8)
        n.alignment = Alignment(wrap_text=False)


def topline(ws, row, c1, c2):
    for col in range(c1, c2 + 1):
        ws.cell(row=row, column=col).border = Border(top=Side(style="thin", color="000000"))


# =====================================================================
# HISTORICALS   B=FY2020 C..G=FY2021..FY2025  H=9M25  I=9M26 / 9-30-26  J=source
# =====================================================================
H = wb.active
H.title = "Historicals"
setup(H, [44, 12, 12, 12, 12, 12, 12, 12, 13, 70])
title(H, "JPMorgan Chase & Co. (NYSE: JPM) - Historical Financials ($ in millions, except per share)",
      "Source: SEC EDGAR XBRL company-facts API (10-K / 10-Q filings), JPM Annual Reports & 3Q26 earnings release. Blue = hard-coded reported data; black = formulas.")
header(H, 4, ["Fiscal year ending Dec-31", "FY2020", "FY2021", "FY2022", "FY2023", "FY2024", "FY2025", "9M 2025", "9M 2026 / 9-30-26", "Source / note"])
HC = {"2020": "B", "2021": "C", "2022": "D", "2023": "E", "2024": "F", "2025": "G", "9M25": "H", "9M26": "I"}
SRC = 10
hr = {}
r = 5


def hrow(key, text, data=None, fmt=USD, formula=None, cols=None, bold=False, note=None, indent=0):
    """data: dict col->value; formula: callable(col_letter, prev_col_letter)->str"""
    global r
    label(H, r, text, bold, indent, note, SRC)
    hr[key] = r
    if data:
        for col, v in data.items():
            put(H, f"{col}{r}", v, fmt, bold=bold)
    if formula:
        for col in cols:
            prev = chr(ord(col) - 1)
            put(H, f"{col}{r}", formula(col, prev), fmt, bold=bold)
    r += 1


def yrs(*vals, cols="CDEFGHI"):
    return {c: v for c, v in zip(cols, vals) if v is not None}


section(H, r, "Income statement", 10); r += 1
hrow("nii", "Net interest income", yrs(52311, 66710, 89267, 92583, 95443, 70448, 77913),
     note="XBRL us-gaap:InterestIncomeExpenseNet; 9M from 3Q26 10-Q")
hrow("nir", "Noninterest revenue", yrs(69338, 61985, 68837, 84973, 87004, 66201, 79442),
     note="XBRL us-gaap:NoninterestIncome. 2024 incl. $7.9B Visa share gain; 2Q26 incl. one-off gains")
hrow("rev", "Total net revenue", formula=lambda c, p: f"={c}{hr['nii']}+{c}{hr['nir']}", cols="CDEFGHI", bold=True,
     note="Ties to reported total net revenue (XBRL us-gaap:Revenues): 121,649 / 128,695 / 158,104 / 177,556 / 182,447")
topline(H, hr["rev"], 2, 9)
hrow("exp", "Noninterest expense", yrs(71343, 76140, 87172, 91797, 95640, 71657, 81461),
     note="XBRL us-gaap:NoninterestExpense")
hrow("ppnr", "Pre-provision net revenue (PPNR)", formula=lambda c, p: f"={c}{hr['rev']}-{c}{hr['exp']}", cols="CDEFGHI", bold=True)
hrow("prov", "Provision for credit losses", yrs(-9256, 6389, 9320, 10678, 14212, 9557, 7765),
     note="XBRL us-gaap:ProvisionForLoanLeaseAndOtherLosses. 2021 = reserve release; 2025 incl. $2.2B Apple Card reserve")
hrow("ebt", "Income before income taxes", formula=lambda c, p: f"={c}{hr['ppnr']}-{c}{hr['prov']}", cols="CDEFGHI")
hrow("tax", "Income tax expense", yrs(11228, 8490, 12060, 16610, 15547, 11412, 15175), note="XBRL us-gaap:IncomeTaxExpenseBenefit")
hrow("ni", "Net income", formula=lambda c, p: f"={c}{hr['ebt']}-{c}{hr['tax']}", cols="CDEFGHI", bold=True)
topline(H, hr["ni"], 2, 9)
hrow("ni_rep", "Memo: reported net income", yrs(48334, 37676, 49552, 58471, 57048, 44023, 52954), indent=1,
     note="XBRL us-gaap:NetIncomeLoss - used only as a tie-out check")
hrow("ni_chk", "Check: model NI - reported NI (should be 0)", formula=lambda c, p: f"={c}{hr['ni']}-{c}{hr['ni_rep']}", cols="CDEFGHI", indent=1)
hrow("nic", "Net income applicable to common", yrs(46503, 35892, 47760, 56868, 55681),
     note="XBRL us-gaap:NetIncomeLossAvailableToCommonStockholdersBasic")
hrow("pref", "Preferred dividends & other", formula=lambda c, p: f"={c}{hr['ni']}-{c}{hr['nic']}", cols="CDEFG", indent=1)
hrow("eps", "Diluted EPS ($)", yrs(15.36, 12.09, 16.23, 19.75, 20.02, 15.38, 19.20), fmt=USD2, note="XBRL us-gaap:EarningsPerShareDiluted")
hrow("shs", "Diluted weighted-average shares (mm)", yrs(3026.6, 2970.0, 2943.1, 2879.0, 2781.5, 2795.2, 2699.1), fmt=NUM1,
     note="XBRL us-gaap:WeightedAverageNumberOfDilutedSharesOutstanding (shown in millions)")
hrow("dps", "Dividends declared per common share ($)", yrs(3.80, 4.00, 4.10, 4.80, 5.80, 4.30, 4.65), fmt=USD2,
     note="XBRL us-gaap:CommonStockDividendsPerShareDeclared; 9M26 = $1.50 + $1.50 + $1.65 (3Q26 release)")
r += 1
section(H, r, "Balance sheet (period-end)", 10); r += 1
hrow("assets", "Total assets", yrs(3384757, 3743567, 3665743, 3875393, 4002814, 4424900, None, 5091545, cols="BCDEFGHI"),
     note="XBRL us-gaap:Assets; 9-30-26 from 3Q26 10-Q")
hrow("loans", "Loans (reported)", yrs(1012853, 1077714, 1135647, 1323706, 1347988, 1493429, None, 1584641, cols="BCDEFGHI"),
     note="JPM Annual Report 'Financial Highlights' (2022, 2023, 2025) and 3Q26 10-Q")
hrow("dep", "Deposits", yrs(2462303, 2340179, 2400688, 2406032, 2559320, None, 2780793), note="XBRL us-gaap:Deposits; 3Q26 10-Q")
hrow("te", "Total stockholders' equity", yrs(294127, 292332, 327878, 344758, 362438, None, 372368), note="XBRL us-gaap:StockholdersEquity")
hrow("ce", "Common stockholders' equity", yrs(249291, 259289, 264928, 300474, 324708, 342393, None, 351328, cols="BCDEFGHI"),
     note="JPM Annual Report Financial Highlights; 9-30-26 per 3Q26 10-Q")
hrow("pfd", "Preferred stock (implied)", formula=lambda c, p: f"={c}{hr['te']}-{c}{hr['ce']}", cols="CDEFGI", indent=1)
hrow("gw", "Goodwill", yrs(49248, 50315, 51662, 52634, 52565, 52731, None, 52717, cols="BCDEFGHI"), note="XBRL us-gaap:Goodwill")
hrow("so", "Common shares outstanding (mm)", {"I": 2655.6}, fmt=NUM1, note="3Q26 10-Q cover page: 2,655,609,509 shares")
hrow("bvps", "Book value per share ($)", {"I": 132.30}, fmt=USD2, note="3Q26 earnings release")
hrow("tbvps", "Tangible book value per share ($)", {"G": 107.56, "I": 112.64}, fmt=USD2, note="4Q25 and 3Q26 earnings releases (non-GAAP)")
hrow("intang", "Goodwill & intangibles, net of DTL (implied)", formula=lambda c, p: f"={c}{hr['ce']}-{c}{hr['tbvps']}*{c}{hr['so']}", cols="I",
     note="Calc: common equity - TBVPS x shares; held flat in the forecast")
r += 1
section(H, r, "Regulatory capital (Basel III Standardized)", 10); r += 1
hrow("cet1r", "CET1 capital ratio", yrs(0.131, 0.132, 0.150, 0.157, 0.146, None, 0.140), fmt=PCT,
     note="JPM Annual Reports (2022, 2025) and 3Q26 10-Q (13.9% estimate in release, 14.0% in 10-Q)")
hrow("cet1", "CET1 capital", {"G": 288000, "I": 304000}, note="4Q25 and 3Q26 releases (~$288B and ~$304B, rounded)")
hrow("rwa", "Risk-weighted assets (derived)", formula=lambda c, p: f"=IFERROR({c}{hr['cet1']}/{c}{hr['cet1r']},0)", cols="GI",
     note="Calc: CET1 capital / CET1 ratio (reported Std. RWA ~ $2.0T at YE25 and ~ $2.2T at 3Q26)")
hrow("dens", "RWA density (RWA / total assets)", formula=lambda c, p: f"=IFERROR({c}{hr['rwa']}/{c}{hr['assets']},0)", cols="GI", fmt=PCT)
r += 1
section(H, r, "Capital returned & cash flow", 10); r += 1
hrow("divpaid", "Common dividends (approx. = DPS x diluted shares)", formula=lambda c, p: f"={c}{hr['dps']}*{c}{hr['shs']}", cols="CDEFG")
hrow("bb", "Common share repurchases", yrs(18408, 3162, 9824, 18830, 31591), note="XBRL us-gaap:PaymentsForRepurchaseOfCommonStock")
hrow("ret", "Total capital returned to common shareholders", formula=lambda c, p: f"={c}{hr['divpaid']}+{c}{hr['bb']}", cols="CDEFG", bold=True)
hrow("payout", "Total payout ratio (% of NI to common)", formula=lambda c, p: f"=IFERROR({c}{hr['ret']}/{c}{hr['nic']},0)", cols="CDEFG", fmt=PCT)
hrow("ocf", "Memo: reported cash flow from operations", yrs(78084, 107119, 12974, -42012, -147782),
     note="XBRL us-gaap:NetCashProvidedByUsedInOperatingActivities. Swings with trading assets - NOT a usable valuation cash flow for a bank")
r += 1
section(H, r, "Key ratios (calculated)", 10); r += 1
hrow("k_revg", "Total net revenue growth", formula=lambda c, p: f"=IFERROR({c}{hr['rev']}/{p}{hr['rev']}-1,0)", cols="DEFG", fmt=PCT)
hrow("k_yld", "NII / average total assets", formula=lambda c, p: f"=IFERROR({c}{hr['nii']}/AVERAGE({p}{hr['assets']},{c}{hr['assets']}),0)", cols="CDEFG", fmt=PCT2)
put(H, f"I{hr['k_yld']}", f"=IFERROR(I{hr['nii']}*4/3/AVERAGE(G{hr['assets']},I{hr['assets']}),0)", PCT2)
hrow("k_nirg", "Noninterest revenue growth", formula=lambda c, p: f"=IFERROR({c}{hr['nir']}/{p}{hr['nir']}-1,0)", cols="DEFG", fmt=PCT)
hrow("k_eff", "Efficiency ratio (expense / revenue)", formula=lambda c, p: f"=IFERROR({c}{hr['exp']}/{c}{hr['rev']},0)", cols="CDEFGHI", fmt=PCT)
hrow("k_prov", "Provision / average loans", formula=lambda c, p: f"=IFERROR({c}{hr['prov']}/AVERAGE({p}{hr['loans']},{c}{hr['loans']}),0)", cols="CDEFG", fmt=PCT2)
put(H, f"I{hr['k_prov']}", f"=IFERROR(I{hr['prov']}*4/3/AVERAGE(G{hr['loans']},I{hr['loans']}),0)", PCT2)
hrow("k_tax", "Effective tax rate", formula=lambda c, p: f"=IFERROR({c}{hr['tax']}/{c}{hr['ebt']},0)", cols="CDEFGHI", fmt=PCT)
hrow("k_roe", "Return on common equity (ROE)", formula=lambda c, p: f"=IFERROR({c}{hr['nic']}/AVERAGE({p}{hr['ce']},{c}{hr['ce']}),0)", cols="CDEFG", fmt=PCT)
hrow("k_rotce", "ROTCE (approx., common equity less goodwill)",
     formula=lambda c, p: f"=IFERROR({c}{hr['nic']}/AVERAGE({p}{hr['ce']}-{p}{hr['gw']},{c}{hr['ce']}-{c}{hr['gw']}),0)", cols="CDEFG", fmt=PCT,
     note="Reported ROTCE also deducts other intangibles net of DTL; this approximation runs ~1pt lower")
hrow("k_ld", "Loans / deposits", formula=lambda c, p: f"=IFERROR({c}{hr['loans']}/{c}{hr['dep']},0)", cols="CDEFGI", fmt=PCT)
note_r = r + 1
H.cell(row=note_r, column=1, value="9M columns are not annualized except where noted (NII yield and provision rate x 4/3).").font = f("595959", italic=True, size=8)
H.freeze_panes = "B5"

# =====================================================================
# ASSUMPTIONS   driver columns D..H = FY2027..FY2031 (same letters as Forecast)
# =====================================================================
A = wb.create_sheet("Assumptions")
setup(A, [46, 16, 13, 11, 11, 11, 11, 11, 62])
title(A, "Assumptions & Scenario Drivers",
      "Blue = inputs; yellow = key levers. Change the scenario in C4 - the entire model, valuation and charts update.")
label(A, 4, "Active scenario (1 = Base, 2 = Bull, 3 = Bear)", True)
put(A, "C4", 1, "0", BLUE, True, YELLOW, "Scenario selector. 1 = Base, 2 = Bull, 3 = Bear")
label(A, 5, "Active scenario name")
put(A, "C5", '=CHOOSE($C$4,"Base","Bull","Bear")', None, BLACK, True)
ar = {}
r = 7
section(A, r, "A. 4Q26 estimate (bridges 9M-2026 actuals to FY2026E; same in all scenarios)", 9); r += 1
header(A, r, ["Item", "", "Value", "", "", "", "", "", "Rationale / source"]); r += 1


def ain(key, text, val, fmt, note, fill=None, formula=False):
    global r
    label(A, r, text, note=note, note_col=9)
    put(A, f"C{r}", val, fmt, BLACK if formula else BLUE, fill=fill)
    ar[key] = r
    r += 1


ain("q4nii", "4Q26E net interest income ($mm)", 28000, USD, "Mgmt. guided ~ $28B 4Q26 total NII on 3Q26 call (MarketBeat transcript summary)")
ain("q4nir", "4Q26E noninterest revenue ($mm)", 23500, USD, "~ 3Q26 reported run-rate ($23.1B) plus modest seasonal IB uplift; analyst estimate")
ain("q4exp", "4Q26E noninterest expense ($mm)", 27000, USD, "In line with 2Q26/3Q26 ($27.3B) - FY26 expense guide ~ $105B+; analyst estimate")
ain("q4prov", "4Q26E provision for credit losses ($mm)", 2800, USD, "3Q26 provision $2.74B ($2.2B NCOs + reserve build); analyst estimate")
ain("q4tax", "4Q26E effective tax rate", 0.22, PCT, "9M26 effective rate 22.3% (10-Q)")
ain("q4ebt", "4Q26E pre-tax income ($mm)", f"=C{ar['q4nii']}+C{ar['q4nir']}-C{ar['q4exp']}-C{ar['q4prov']}", USD, "Calculated", formula=True)
ain("q4ni", "4Q26E net income ($mm)", f"=C{ar['q4ebt']}*(1-C{ar['q4tax']})", USD, "Calculated", formula=True)
ain("q4dps", "4Q26E dividend per share ($)", 1.65, USD2, "Quarterly dividend raised to $1.65 in 3Q26 (3Q26 release)")
ain("q4bb", "4Q26E net share repurchases ($mm)", 3000, USD, "3Q26 net repurchases were only $664M; assume a modest step-up; analyst estimate")
ain("q4ag", "YE26 total assets growth vs. 9-30-26", 0.0, PCT, "Assume flat balance sheet into year-end")
ain("q4lg", "YE26 loan growth vs. 9-30-26", 0.0, PCT, "Assume flat")
ain("q4rg", "YE26 RWA growth vs. 9-30-26", 0.0, PCT, "Assume flat")
r += 1
section(A, r, "B. General assumptions (all scenarios)", 9); r += 1
ain("tax", "Effective tax rate, FY2027E+", 0.22, PCT, "FY25 21.4%, 9M26 22.3% (Historicals)")
ain("pref", "Preferred dividends ($mm per year)", 1400, USD, "FY25 preferred dividends & other $1,367mm (Historicals)")
ain("bbpx", "Buyback price appreciation (% per year)", 0.06, PCT, "Repurchase price grows from current price; roughly cost of equity less dividend yield")
ain("regmin", "Regulatory minimum CET1 ratio (approx.)", 0.115, PCT,
    "Approximation: 4.5% minimum + stress capital buffer + GSIB surcharge. Verify against latest 10-Q capital section")
r += 1
section(A, r, "C. Scenario drivers (FY2027E - FY2031E)", 9); r += 1
header(A, r, ["Driver / scenario", "", "Shock (+/-)", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "Rationale"]); r += 1

DRIVERS = [
    ("ag", "Total asset growth", PCT,
     [0.04] * 5, [0.06] * 5, [-0.02, 0.01, 0.02, 0.03, 0.03],
     "Base ~ nominal GDP; Bull = continued share gains; Bear = recession deleveraging in 2027"),
    ("lg", "Loan growth", PCT,
     [0.06, 0.06, 0.05, 0.05, 0.05], [0.08, 0.07, 0.07, 0.06, 0.06], [0.00, 0.02, 0.04, 0.04, 0.04],
     "FY23-25 CAGR ~6% (ex-First Republic) + Apple Card portfolio onboarding"),
    ("yld", "NII / average total assets", PCT2,
     [0.0212] * 5, [0.0218, 0.0220, 0.0222, 0.0222, 0.0222], [0.0195, 0.0192, 0.0198, 0.0202, 0.0205],
     "FY23-25 2.27-2.36%; 9M26 annualized ~2.2% as balance sheet grows in lower-yield Markets assets. Bear = Fed cuts compress NIM"),
    ("nirg", "Noninterest revenue growth (off 4Q26E run-rate)", PCT,
     [0.04, 0.05, 0.05, 0.045, 0.045], [0.08, 0.07, 0.06, 0.06, 0.05], [-0.08, 0.02, 0.05, 0.05, 0.04],
     "Fees track IB wallet, markets volumes and AUM. Bear = IB/markets downturn (2022-style -11%)"),
    ("eff", "Efficiency ratio (expense / revenue)", PCT,
     [0.545, 0.54, 0.535, 0.535, 0.535], [0.53, 0.525, 0.52, 0.52, 0.52], [0.58, 0.57, 0.56, 0.55, 0.55],
     "FY24 51.7%, FY25 52.4%; heavy tech/AI investment keeps ratio ~54%. Bear = negative operating leverage"),
    ("prv", "Provision / average loans", PCT2,
     [0.0070, 0.0075, 0.0075, 0.0075, 0.0075], [0.0055, 0.0060, 0.0060, 0.0065, 0.0065], [0.0160, 0.0130, 0.0095, 0.0085, 0.0080],
     "FY24 0.80%, FY25 1.00% (incl. Apple Card reserve). Bear ~ 2020 COVID provisioning (~1.7%)"),
    ("dens", "RWA density (RWA / total assets)", PCT,
     [0.426] * 5, [0.42] * 5, [0.445, 0.445, 0.44, 0.44, 0.44],
     "9-30-26 = 42.6% (Historicals). Bear = credit migration inflates risk weights"),
    ("tgt", "Target CET1 ratio (management buffer)", PCT,
     [0.135] * 5, [0.13] * 5, [0.14] * 5,
     "9-30-26 CET1 14.0%; excess above target is distributed. Bear = management holds a larger buffer"),
    ("pay", "Dividend payout ratio (% of NI to common)", PCT,
     [0.30] * 5, [0.28] * 5, [0.40] * 5,
     "FY25 dividends ~29% of NI to common; buybacks absorb the remaining excess capital"),
]
for key, text, fmt, base, bull, bear, why in DRIVERS:
    label(A, r, text, True)
    A.cell(row=r, column=9, value=why).font = f("595959", italic=True, size=8)
    r += 1
    rows = {}
    for name, vals in (("Base", base), ("Bull", bull), ("Bear", bear)):
        label(A, r, name, indent=2)
        for i, v in enumerate(vals):
            put(A, f"{L(4 + i)}{r}", v, fmt)
        rows[name] = r
        r += 1
    label(A, r, "Active (used in model)", True, indent=2)
    put(A, f"C{r}", 0, fmt, BLUE, fill=YELLOW, comment="Sensitivity shock added to every forecast year (0 = none). Used for tornado analysis.")
    for i in range(5):
        c = L(4 + i)
        put(A, f"{c}{r}", f"=CHOOSE($C$4,{c}{rows['Base']},{c}{rows['Bull']},{c}{rows['Bear']})+$C${r}", fmt, BLACK, True)
    ar[key] = r
    ar[key + "_shock"] = r
    r += 2
A.freeze_panes = "B7"

# =====================================================================
# FORECAST   B=FY2025A C=FY2026E D..H=FY2027E..FY2031E
# =====================================================================
F = wb.create_sheet("Forecast")
C = wb.create_sheet("Capital & Cash Flow")
setup(F, [46, 13, 13, 13, 13, 13, 13, 13, 60])
setup(C, [52, 13, 13, 13, 13, 13, 13, 13, 60])
title(F, "Income Statement Forecast ($ in millions, except per share)",
      "FY2026E = 9M-2026 actuals + 4Q26 estimate. FY2027E-FY2031E driven by the active scenario on Assumptions.")
YH = ["", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "Driver / note"]
FC = list("CDEFGH")
FY = list("DEFGH")
fr, cr = {}, {}
HS, AS, CS = "Historicals", "Assumptions", "'Capital & Cash Flow'"
hs = lambda k, col: f"{HS}!{col}{hr[k]}"
as_ = lambda k, col="$C": f"{AS}!{col}${ar[k]}" if col == "$C" else f"{AS}!{col}{ar[k]}"


def frow(ws, store, key, text, B=None, Cv=None, D=None, fmt=USD, bold=False, note=None, indent=0, line=False):
    """B, Cv: formula strings; D: callable(col, prevcol)->formula for D..H"""
    global r
    label(ws, r, text, bold, indent, note, 9)
    store[key] = r
    if B is not None:
        put(ws, f"B{r}", B, fmt, bold=bold)
    if Cv is not None:
        put(ws, f"C{r}", Cv, fmt, bold=bold)
    if D is not None:
        for col in FY:
            put(ws, f"{col}{r}", D(col, chr(ord(col) - 1)), fmt, bold=bold)
    if line:
        topline(ws, r, 2, 8)
    r += 1


ccap = lambda k: f"{CS}!{{}}{cr[k]}"  # placeholder (capital rows assigned later)

r = 4
header(F, r, YH); r += 1
section(F, r, "Balance sheet drivers", 9); r += 1
frow(F, fr, "assets", "Total assets (period-end)", f"={hs('assets','G')}", f"={hs('assets','I')}*(1+{as_('q4ag')})",
     lambda c, p: f"={p}{r}*(1+{AS}!{c}{ar['ag']})", note="Prior year x (1 + asset growth)")
frow(F, fr, "avga", "Average total assets", f"=AVERAGE({hs('assets','F')},{hs('assets','G')})", f"=AVERAGE(B{fr['assets']},C{fr['assets']})",
     lambda c, p: f"=AVERAGE({p}{fr['assets']},{c}{fr['assets']})")
frow(F, fr, "loans", "Loans (period-end)", f"={hs('loans','G')}", f"={hs('loans','I')}*(1+{as_('q4lg')})",
     lambda c, p: f"={p}{r}*(1+{AS}!{c}{ar['lg']})", note="Prior year x (1 + loan growth)")
frow(F, fr, "avgl", "Average loans", f"=AVERAGE({hs('loans','F')},{hs('loans','G')})", f"=AVERAGE(B{fr['loans']},C{fr['loans']})",
     lambda c, p: f"=AVERAGE({p}{fr['loans']},{c}{fr['loans']})")
r += 1
section(F, r, "Income statement", 9); r += 1
frow(F, fr, "nii", "Net interest income", f"={hs('nii','G')}", f"={hs('nii','I')}+{as_('q4nii')}",
     lambda c, p: f"={AS}!{c}{ar['yld']}*{c}{fr['avga']}", note="Average assets x NII yield")
frow(F, fr, "nirrr", "Noninterest revenue run-rate", None, f"={as_('q4nir')}*4",
     lambda c, p: f"={p}{r}*(1+{AS}!{c}{ar['nirg']})", indent=1,
     note="FY26E base = 4Q26E x 4, stripping 2Q26 one-off gains from the growth base")
frow(F, fr, "nir", "Noninterest revenue", f"={hs('nir','G')}", f"={hs('nir','I')}+{as_('q4nir')}",
     lambda c, p: f"={c}{fr['nirrr']}")
frow(F, fr, "rev", "Total net revenue", f"=B{fr['nii']}+B{fr['nir']}", f"=C{fr['nii']}+C{fr['nir']}",
     lambda c, p: f"={c}{fr['nii']}+{c}{fr['nir']}", bold=True, line=True)
frow(F, fr, "exp", "Noninterest expense", f"={hs('exp','G')}", f"={hs('exp','I')}+{as_('q4exp')}",
     lambda c, p: f"={AS}!{c}{ar['eff']}*{c}{fr['rev']}", note="Revenue x efficiency ratio")
frow(F, fr, "ppnr", "Pre-provision net revenue (PPNR)", f"=B{fr['rev']}-B{fr['exp']}", f"=C{fr['rev']}-C{fr['exp']}",
     lambda c, p: f"={c}{fr['rev']}-{c}{fr['exp']}", bold=True)
frow(F, fr, "prov", "Provision for credit losses", f"={hs('prov','G')}", f"={hs('prov','I')}+{as_('q4prov')}",
     lambda c, p: f"={AS}!{c}{ar['prv']}*{c}{fr['avgl']}", note="Average loans x provision rate")
frow(F, fr, "ebt", "Income before income taxes", f"=B{fr['ppnr']}-B{fr['prov']}", f"=C{fr['ppnr']}-C{fr['prov']}",
     lambda c, p: f"={c}{fr['ppnr']}-{c}{fr['prov']}")
frow(F, fr, "tax", "Income tax expense", f"={hs('tax','G')}", f"={hs('tax','I')}+{AS}!$C${ar['q4ebt']}*{as_('q4tax')}",
     lambda c, p: f"={c}{fr['ebt']}*{as_('tax')}")
frow(F, fr, "ni", "Net income", f"=B{fr['ebt']}-B{fr['tax']}", f"=C{fr['ebt']}-C{fr['tax']}",
     lambda c, p: f"={c}{fr['ebt']}-{c}{fr['tax']}", bold=True, line=True)
frow(F, fr, "pref", "Preferred dividends & other", f"={hs('pref','G')}", f"={as_('pref')}", lambda c, p: f"={as_('pref')}", indent=1)
frow(F, fr, "nic", "Net income applicable to common", f"=B{fr['ni']}-B{fr['pref']}", f"=C{fr['ni']}-C{fr['pref']}",
     lambda c, p: f"={c}{fr['ni']}-{c}{fr['pref']}", bold=True)
r += 1
section(F, r, "Per share", 9); r += 1
fr["shs"] = r; r += 1  # filled after capital sheet rows exist
fr["eps"] = r; r += 1
fr["dps"] = r; r += 1
fr["bvps"] = r; r += 1
fr["tbvps"] = r; r += 1
r += 1
section(F, r, "Key ratios", 9); r += 1
frow(F, fr, "k_revg", "Total net revenue growth", f"={hs('k_revg','G')}", f"=IFERROR(C{fr['rev']}/B{fr['rev']}-1,0)",
     lambda c, p: f"=IFERROR({c}{fr['rev']}/{p}{fr['rev']}-1,0)", PCT, note="FY27E growth looks flat because FY26 includes one-off 2Q26 gains")
frow(F, fr, "k_yld", "NII / average total assets", f"=IFERROR(B{fr['nii']}/B{fr['avga']},0)", f"=IFERROR(C{fr['nii']}/C{fr['avga']},0)",
     lambda c, p: f"=IFERROR({c}{fr['nii']}/{c}{fr['avga']},0)", PCT2)
frow(F, fr, "k_eff", "Efficiency ratio", f"=IFERROR(B{fr['exp']}/B{fr['rev']},0)", f"=IFERROR(C{fr['exp']}/C{fr['rev']},0)",
     lambda c, p: f"=IFERROR({c}{fr['exp']}/{c}{fr['rev']},0)", PCT)
frow(F, fr, "k_prov", "Provision / average loans", f"=IFERROR(B{fr['prov']}/B{fr['avgl']},0)", f"=IFERROR(C{fr['prov']}/C{fr['avgl']},0)",
     lambda c, p: f"=IFERROR({c}{fr['prov']}/{c}{fr['avgl']},0)", PCT2)
frow(F, fr, "k_tax", "Effective tax rate", f"=IFERROR(B{fr['tax']}/B{fr['ebt']},0)", f"=IFERROR(C{fr['tax']}/C{fr['ebt']},0)",
     lambda c, p: f"=IFERROR({c}{fr['tax']}/{c}{fr['ebt']},0)", PCT)
fr["roe"] = r; r += 1
fr["rotce"] = r; r += 1
fr["nimargin"] = r; r += 1
F.freeze_panes = "B5"

# =====================================================================
# CAPITAL & CASH FLOW
# =====================================================================
title(C, "Capital Plan & Cash Flows to Shareholders ($ in millions)",
      "For a bank, the cash available to equity holders is earnings minus the capital that must be retained to hold the target CET1 ratio (FCFE).")
r = 4
header(C, r, YH); r += 1
FS = "Forecast"
fs = lambda k, col: f"{FS}!{col}{fr[k]}"
section(C, r, "Regulatory capital roll-forward (FY2026E column = 4Q26 only, starting from 9-30-26)", 9); r += 1
frow(C, cr, "rwa", "Risk-weighted assets (period-end)", f"={hs('rwa','G')}", f"={hs('rwa','I')}*(1+{as_('q4rg')})",
     lambda c, p: f"={AS}!{c}{ar['dens']}*{FS}!{c}{fr['assets']}", note="Total assets x RWA density")
frow(C, cr, "cbop", "CET1 capital, beginning of period", None, f"={hs('cet1','I')}",
     lambda c, p: f"={p}{r + 4}")
frow(C, cr, "cni", "(+) Net income to common", None, f"={AS}!$C${ar['q4ni']}-{as_('pref')}/4",
     lambda c, p: f"={FS}!{c}{fr['nic']}", indent=1, note="FY26E column: 4Q26E only")
frow(C, cr, "cdiv", "(-) Common dividends", None, f"={as_('q4dps')}*{hs('so','I')}",
     lambda c, p: f"={AS}!{c}{ar['pay']}*{c}{cr['cni']}", indent=1, note="Payout ratio x NI to common")
cr["cbb"] = r; r += 1  # buybacks written after FCFE row known
frow(C, cr, "ceop", "CET1 capital, end of period", f"={hs('cet1','G')}", f"=C{cr['cbop']}+C{cr['cni']}-C{cr['cdiv']}-C{cr['cbb']}",
     lambda c, p: f"={c}{cr['cbop']}+{c}{cr['cni']}-{c}{cr['cdiv']}-{c}{cr['cbb']}", bold=True, line=True,
     note="Simplification: AOCI, deductions and intangibles held constant")
frow(C, cr, "cr", "CET1 ratio", f"=IFERROR(B{cr['ceop']}/B{cr['rwa']},0)", f"=IFERROR(C{cr['ceop']}/C{cr['rwa']},0)",
     lambda c, p: f"=IFERROR({c}{cr['ceop']}/{c}{cr['rwa']},0)", PCT, bold=True)
frow(C, cr, "tgt", "Target CET1 ratio", None, None, lambda c, p: f"={AS}!{c}{ar['tgt']}", PCT, indent=1)
frow(C, cr, "min", "Regulatory minimum CET1 ratio (approx.)", f"={as_('regmin')}", f"={as_('regmin')}", lambda c, p: f"={as_('regmin')}", PCT, indent=1)
frow(C, cr, "buf", "Buffer over regulatory minimum (bps)", f"=(B{cr['cr']}-B{cr['min']})*10000", f"=(C{cr['cr']}-C{cr['min']})*10000",
     lambda c, p: f"=({c}{cr['cr']}-{c}{cr['min']})*10000", "#,##0;(#,##0)", indent=1, note="Risk flag: < 100 bps means capital returns would be cut")
r += 1
section(C, r, "Free cash flow to equity (FCFE) & distributions", 9); r += 1
frow(C, cr, "req", "Required CET1 at target ratio", None, None, lambda c, p: f"={c}{cr['tgt']}*{c}{cr['rwa']}")
frow(C, cr, "fcfe", "FCFE = CET1 BoP + NI to common - required CET1", None, None,
     lambda c, p: f"={c}{cr['cbop']}+{c}{cr['cni']}-{c}{cr['req']}", bold=True,
     note="Capital generated above the target buffer - the cash that can be paid out")
frow(C, cr, "fdiv", "of which: common dividends", None, f"=C{cr['cdiv']}", lambda c, p: f"={c}{cr['cdiv']}", indent=1)
frow(C, cr, "fbb", "of which: share repurchases", None, f"=C{cr['cbb']}", lambda c, p: f"={c}{cr['cbb']}", indent=1)
frow(C, cr, "ret", "Total capital returned to common", f"={hs('ret','G')}", f"=C{cr['fdiv']}+C{cr['fbb']}",
     lambda c, p: f"={c}{cr['fdiv']}+{c}{cr['fbb']}", bold=True, line=True)
frow(C, cr, "payout", "Total payout (% of NI to common)", f"={hs('payout','G')}", None,
     lambda c, p: f"=IFERROR({c}{cr['ret']}/{c}{cr['cni']},0)", PCT, note="Can exceed 100% when excess capital is released")
# buybacks row
label(C, cr["cbb"], "(-) Share repurchases", indent=1, note="Calc: max(0, FCFE - dividends); FY26E = 4Q26E input", note_col=9)
put(C, f"C{cr['cbb']}", f"={as_('q4bb')}", USD)
for col in FY:
    put(C, f"{col}{cr['cbb']}", f"=MAX(0,{col}{cr['fcfe']}-{col}{cr['cdiv']})", USD)
r += 1
section(C, r, "Share count", 9); r += 1
frow(C, cr, "px", "Average repurchase price ($)", None, "=Valuation!$C$6",
     lambda c, p: f"={p}{r}*(1+{as_('bbpx')})", USD2, note="FY26E = current share price")
frow(C, cr, "sbop", "Shares outstanding, beginning (mm)", None, f"={hs('so','I')}", lambda c, p: f"={p}{r + 2}", NUM1)
frow(C, cr, "srep", "(-) Shares repurchased (mm)", None, f"=IFERROR(C{cr['cbb']}/C{cr['px']},0)",
     lambda c, p: f"=IFERROR({c}{cr['cbb']}/{c}{cr['px']},0)", NUM1, indent=1)
frow(C, cr, "seop", "Shares outstanding, end (mm)", None, f"=C{cr['sbop']}-C{cr['srep']}",
     lambda c, p: f"={c}{cr['sbop']}-{c}{cr['srep']}", NUM1, bold=True)
frow(C, cr, "savg", "Average shares (mm)", None, f"=AVERAGE(C{cr['sbop']},C{cr['seop']})",
     lambda c, p: f"=AVERAGE({c}{cr['sbop']},{c}{cr['seop']})", NUM1, note="Diluted ~ basic (dilution < 0.5%)")
r += 1
section(C, r, "Common equity & tangible book value", 9); r += 1
frow(C, cr, "ebop", "Common equity, beginning of period", None, f"={hs('ce','I')}", lambda c, p: f"={p}{r + 4}")
frow(C, cr, "eni", "(+) Net income to common", None, f"=C{cr['cni']}", lambda c, p: f"={c}{cr['cni']}", indent=1)
frow(C, cr, "ediv", "(-) Dividends", None, f"=C{cr['cdiv']}", lambda c, p: f"={c}{cr['cdiv']}", indent=1)
frow(C, cr, "ebb", "(-) Share repurchases", None, f"=C{cr['cbb']}", lambda c, p: f"={c}{cr['cbb']}", indent=1)
frow(C, cr, "eeop", "Common equity, end of period", f"={hs('ce','G')}", f"=C{cr['ebop']}+C{cr['eni']}-C{cr['ediv']}-C{cr['ebb']}",
     lambda c, p: f"={c}{cr['ebop']}+{c}{cr['eni']}-{c}{cr['ediv']}-{c}{cr['ebb']}", bold=True, line=True)
frow(C, cr, "intg", "Goodwill & intangibles, net of DTL", f"={hs('gw','G')}", f"={hs('intang','I')}", lambda c, p: f"={p}{r}",
     indent=1, note="Held flat (no acquisitions). FY25A uses goodwill only")
frow(C, cr, "tce", "Tangible common equity (TCE)", f"=B{cr['eeop']}-B{cr['intg']}", f"=C{cr['eeop']}-C{cr['intg']}",
     lambda c, p: f"={c}{cr['eeop']}-{c}{cr['intg']}")
frow(C, cr, "avgce", "Average common equity", None, f"=AVERAGE(B{cr['eeop']},C{cr['eeop']})",
     lambda c, p: f"=AVERAGE({p}{cr['eeop']},{c}{cr['eeop']})")
frow(C, cr, "avgtce", "Average TCE", None, f"=AVERAGE(B{cr['tce']},C{cr['tce']})", lambda c, p: f"=AVERAGE({p}{cr['tce']},{c}{cr['tce']})")
r += 1
nr = r
notes = [
    "Why not a standard FCF or operating cash flow? Deposits and debt are a bank's raw material, not its financing. Reported operating cash flow",
    "swings with trading balances (FY25: -$148B; FY22: +$107B - see Historicals), so it says nothing about value. Analysts instead value the",
    "capital a bank can distribute after holding the CET1 buffer regulators and management require: FCFE = earnings - retained capital.",
]
for i, t in enumerate(notes):
    C.cell(row=nr + i, column=1, value=t).font = f("595959", italic=True, size=9)
C.freeze_panes = "B5"

# fill forecast per-share & ratios now that capital rows exist
cs = lambda k, col: f"{CS}!{col}{cr[k]}"
r = fr["shs"]
frow(F, fr, "shs", "Diluted average shares (mm)", f"={hs('shs','G')}", f"=({hs('shs','I')}*3+{cs('savg','C')})/4",
     lambda c, p: f"={CS}!{c}{cr['savg']}", NUM1, note="FY26E = 9M weighted average x 3/4 + 4Q average x 1/4")
frow(F, fr, "eps", "Diluted EPS ($)", f"={hs('eps','G')}", f"=IFERROR(C{fr['nic']}/C{fr['shs']},0)",
     lambda c, p: f"=IFERROR({c}{fr['nic']}/{c}{fr['shs']},0)", USD2, bold=True)
frow(F, fr, "dps", "Dividends per share ($)", f"={hs('dps','G')}", f"={hs('dps','I')}+{as_('q4dps')}",
     lambda c, p: f"=IFERROR({CS}!{c}{cr['cdiv']}/{CS}!{c}{cr['savg']},0)", USD2)
frow(F, fr, "bvps", "Book value per share, period-end ($)", None, f"=IFERROR({cs('eeop','C')}/{cs('seop','C')},0)",
     lambda c, p: f"=IFERROR({CS}!{c}{cr['eeop']}/{CS}!{c}{cr['seop']},0)", USD2)
frow(F, fr, "tbvps", "Tangible book value per share, period-end ($)", f"={hs('tbvps','G')}", f"=IFERROR({cs('tce','C')}/{cs('seop','C')},0)",
     lambda c, p: f"=IFERROR({CS}!{c}{cr['tce']}/{CS}!{c}{cr['seop']},0)", USD2)
r = fr["roe"]
frow(F, fr, "roe", "Return on common equity (ROE)", f"={hs('k_roe','G')}", f"=IFERROR(C{fr['nic']}/{cs('avgce','C')},0)",
     lambda c, p: f"=IFERROR({c}{fr['nic']}/{CS}!{c}{cr['avgce']},0)", PCT)
frow(F, fr, "rotce", "Return on tangible common equity (ROTCE)", f"={hs('k_rotce','G')}", f"=IFERROR(C{fr['nic']}/{cs('avgtce','C')},0)",
     lambda c, p: f"=IFERROR({c}{fr['nic']}/{CS}!{c}{cr['avgtce']},0)", PCT, bold=True)
frow(F, fr, "nimargin", "Net margin (NI / revenue)", f"=IFERROR(B{fr['ni']}/B{fr['rev']},0)", f"=IFERROR(C{fr['ni']}/C{fr['rev']},0)",
     lambda c, p: f"=IFERROR({c}{fr['ni']}/{c}{fr['rev']},0)", PCT)

# =====================================================================
# VALUATION
# =====================================================================
V = wb.create_sheet("Valuation")
setup(V, [50, 14, 14, 13, 13, 13, 13, 13, 58])
title(V, "Valuation - FCFE DCF (excess-capital method), Residual Income and Justified P/TBV",
      "Equity-level valuation: for banks, discount cash to shareholders at the cost of equity (no WACC / enterprise value).")
vr = {}
r = 4
section(V, r, "Market & discount-rate inputs", 9); r += 1


def vin(key, text, val, fmt, note, color=None, fill=None):
    global r
    label(V, r, text, note=note, note_col=9)
    put(V, f"C{r}", val, fmt, color, fill=fill)
    vr[key] = r
    r += 1


vin("date", "Valuation date", "=DATE(2026,11,20)", "mmm d, yyyy", "Model date", BLUE)
vin("px", "Current share price ($)", 338.21, USD2, "JPM close on Nov 16, 2026 (MarketChameleon)", fill=YELLOW)
assert vr["px"] == 6, vr  # Capital sheet links to Valuation!$C$6
vin("rf", "Risk-free rate (10-yr UST)", 0.0512, PCT2, "10-yr Treasury 5.12% on Nov 16, 2026 (FRED via MarketXLS)")
vin("beta", "Equity beta", 1.02, "0.00", "5Y monthly beta vs. S&P 500, Yahoo Finance (user-provided, Nov 20, 2026)")
vin("erp", "Equity risk premium", 0.045, PCT2, "Implied US ERP ~4.0-4.5% (Damodaran-style estimate)")
vin("ke", "Cost of equity (CAPM)", f"=C{vr['rf']}+C{vr['beta']}*C{vr['erp']}", PCT2, "Calc: Rf + beta x ERP", BLACK)
vin("g", "Terminal growth rate", 0.04, PCT2, "Long-run nominal growth; must stay below cost of equity", fill=YELLOW)
r += 1
section(V, r, "1) FCFE DCF - distributable capital discounted at cost of equity", 9); r += 1
header(V, r, ["", "", "", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "Note"]); r += 1
vr["yr"] = r
label(V, r, "Fiscal year")
for i, col in enumerate(FY):
    put(V, f"{col}{r}", 2027 + i, "0")
r += 1
vr["fcfe"] = r
label(V, r, "FCFE (from Capital & Cash Flow)", True)
for col in FY:
    put(V, f"{col}{r}", f"={cs('fcfe', col)}", USD, bold=True)
r += 1
vr["t"] = r
label(V, r, "Discount period (years, mid-year convention)", note="Distributions are paid through the year", note_col=9)
for col in FY:
    put(V, f"{col}{r}", f"=YEARFRAC($C${vr['date']},DATE({col}{vr['yr']},6,30))", "0.00")
r += 1
vr["df"] = r
label(V, r, "Discount factor")
for col in FY:
    put(V, f"{col}{r}", f"=1/(1+$C${vr['ke']})^{col}{vr['t']}", "0.0000")
r += 1
vr["pv"] = r
label(V, r, "PV of FCFE")
for col in FY:
    put(V, f"{col}{r}", f"={col}{vr['fcfe']}*{col}{vr['df']}", USD)
r += 1
r += 1


def vcalc(key, text, formula, fmt, note=None, bold=False, fill=None):
    global r
    label(V, r, text, bold, note=note, note_col=9)
    put(V, f"C{r}", formula, fmt, bold=bold, fill=fill)
    vr[key] = r
    r += 1


vcalc("sumpv", "Sum of PV of FCFE, FY27E-FY31E", f"=SUM(D{vr['pv']}:H{vr['pv']})", USD)
vcalc("tfcfe", "Normalized terminal FCFE (FY2032E)", f"={cs('cni','H')}*(1+C{vr['g']})-C{vr['g']}*{cs('req','H')}", USD,
      "NI to common FY31 x (1+g) - g x required CET1 FY31: retention consistent with terminal growth")
vcalc("tv", "Terminal value at FY2031 (Gordon growth)", f"=C{vr['tfcfe']}/(C{vr['ke']}-C{vr['g']})", USD,
      "Normalized terminal FCFE / (ke - g)")
vcalc("ttv", "Terminal value discount period (years to Dec-31-2031)", f"=YEARFRAC(C{vr['date']},DATE(H{vr['yr']},12,31))", "0.00")
vcalc("pvtv", "PV of terminal value", f"=C{vr['tv']}/(1+C{vr['ke']})^C{vr['ttv']}", USD)
vcalc("eqv", "Equity value", f"=C{vr['sumpv']}+C{vr['pvtv']}", USD, bold=True)
vcalc("shs", "Shares outstanding at YE2026E (mm)", f"={cs('seop','C')}", NUM1)
vcalc("dcfpx", "FCFE DCF value per share ($)", f"=IFERROR(C{vr['eqv']}/C{vr['shs']},0)", USD2, bold=True, fill=LIGHT)
vcalc("tvpct", "Terminal value as % of equity value", f"=IFERROR(C{vr['pvtv']}/C{vr['eqv']},0)", PCT)
r += 1
section(V, r, "2) Residual income (excess return on book equity)", 9); r += 1
header(V, r, ["", "", "", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "Note"]); r += 1


def vline(key, text, fn, fmt, note=None, bold=False):
    global r
    label(V, r, text, bold, note=note, note_col=9)
    for col in FY:
        put(V, f"{col}{r}", fn(col), fmt, bold=bold)
    vr[key] = r
    r += 1


vline("ri_bv", "Common equity, beginning of year", lambda c: f"={cs('ebop', c)}", USD)
vline("ri_ni", "Net income to common", lambda c: f"={fs('nic', c)}", USD)
vline("ri_ch", "Equity charge (ke x beginning equity)", lambda c: f"=$C${vr['ke']}*{c}{vr['ri_bv']}", USD)
vline("ri", "Residual income", lambda c: f"={c}{vr['ri_ni']}-{c}{vr['ri_ch']}", USD, bold=True)
vline("ri_t", "Discount period (years to year-end)", lambda c: f"=YEARFRAC($C${vr['date']},DATE({c}{vr['yr']},12,31))", "0.00")
vline("ri_pv", "PV of residual income", lambda c: f"={c}{vr['ri']}/(1+$C${vr['ke']})^{c}{vr['ri_t']}", USD)
r += 1
vcalc("ri_bv0", "Common equity at YE2026E", f"={cs('eeop','C')}", USD)
vcalc("ri_sum", "Sum of PV of residual income", f"=SUM(D{vr['ri_pv']}:H{vr['ri_pv']})", USD)
vcalc("ri_tv", "PV of terminal residual income", f"=H{vr['ri']}*(1+C{vr['g']})/(C{vr['ke']}-C{vr['g']})/(1+C{vr['ke']})^H{vr['ri_t']}", USD)
vcalc("ri_eq", "Equity value", f"=C{vr['ri_bv0']}+C{vr['ri_sum']}+C{vr['ri_tv']}", USD, bold=True)
vcalc("ripx", "Residual income value per share ($)", f"=IFERROR(C{vr['ri_eq']}/C{vr['shs']},0)", USD2, bold=True, fill=LIGHT)
r += 1
section(V, r, "3) Justified price / tangible book (Gordon growth on ROTCE)", 9); r += 1
vcalc("rotce", "Forward ROTCE (FY2027E)", f"={fs('rotce','D')}", PCT)
vcalc("jptbv", "Justified P/TBV = (ROTCE - g) / (ke - g)", f"=(C{vr['rotce']}-C{vr['g']})/(C{vr['ke']}-C{vr['g']})", MULT)
vcalc("tbv0", "TBV per share at YE2026E ($)", f"={fs('tbvps','C')}", USD2)
vcalc("ptbvpx", "Justified P/TBV value per share ($)", f"=C{vr['jptbv']}*C{vr['tbv0']}", USD2, bold=True, fill=LIGHT)
vcalc("curptbv", "Memo: current P/TBV (on 9-30-26 TBVPS)", f"=C{vr['px']}/{hs('tbvps','I')}", MULT)
vcalc("impke", "Memo: market-implied cost of equity", f"=C{vr['g']}+(C{vr['rotce']}-C{vr['g']})/C{vr['curptbv']}", PCT2,
      "Cost of equity that justifies today's price given forward ROTCE")
r += 1
section(V, r, "Valuation summary", 9); r += 1
header(V, r, ["Method", "Weight", "Value / share", "vs. price", "", "", "", "", ""]); r += 1
vr["sum0"] = r
for text, key, w in (("FCFE DCF (excess capital)", "dcfpx", 0.5), ("Residual income", "ripx", 0.25), ("Justified P/TBV", "ptbvpx", 0.25)):
    label(V, r, text)
    put(V, f"B{r}", w, PCT, fill=YELLOW)
    put(V, f"C{r}", f"=C{vr[key]}", USD2)
    put(V, f"D{r}", f"=IFERROR(C{r}/$C${vr['px']}-1,0)", PCT)
    r += 1
label(V, r, "Blended value per share ($)", True)
put(V, f"B{r}", f"=SUM(B{vr['sum0']}:B{r - 1})", PCT, bold=True)
put(V, f"C{r}", f"=IFERROR(SUMPRODUCT(B{vr['sum0']}:B{r - 1},C{vr['sum0']}:C{r - 1})/B{r},0)", USD2, bold=True, fill=LIGHT)
put(V, f"D{r}", f"=IFERROR(C{r}/$C${vr['px']}-1,0)", PCT, bold=True)
vr["blend"] = r
topline(V, r, 1, 4)
r += 1
label(V, r, "Implied rating (>+15% Buy, <-15% Sell, else Hold)")
put(V, f"C{r}", f'=IF(D{vr["blend"]}>0.15,"BUY",IF(D{vr["blend"]}<-0.15,"SELL","HOLD"))', None, BLACK, True)
vr["rating"] = r
V.freeze_panes = "A4"

# =====================================================================
# SCENARIOS & SENSITIVITY (live tables + script-generated static tables)
# =====================================================================
S = wb.create_sheet("Scenarios & Sensitivity")
setup(S, [40, 12, 12, 12, 12, 12, 12, 12, 12, 50])
title(S, "Risk Assessment - Scenarios, Stress Test & Sensitivity",
      "Tables 1-2 are live formulas. Tables 3-6 are generated by re-running the full model under each scenario/shock (Python + LibreOffice).")
sr = {}
r = 4
section(S, r, "Table 1 (live): FCFE DCF value per share - cost of equity vs. terminal growth (active scenario)", 10); r += 1
gs = [0.03, 0.035, 0.04, 0.045, 0.05]
kes = [0.085, 0.09, 0.095, 0.10, 0.105, 0.11, 0.115, 0.12]
put(S, f"A{r}", "Cost of equity ↓ / Terminal growth →", color=BLACK, bold=True)
for j, g in enumerate(gs):
    put(S, f"{L(2 + j)}{r}", g, PCT, bold=True)
hdr = r
r += 1
VF = f"Valuation!$D${vr['fcfe']}:$H${vr['fcfe']}"
VT = f"Valuation!$D${vr['t']}:$H${vr['t']}"
for ke in kes:
    put(S, f"A{r}", ke, PCT, bold=True)
    S[f"A{r}"].alignment = Alignment(horizontal="right")
    for j in range(len(gs)):
        c = L(2 + j)
        fm = (f"=IF($A{r}<={c}${hdr},NA(),(SUMPRODUCT({VF},1/(1+$A{r})^{VT})"
              f"+({CS}!$H${cr['cni']}*(1+{c}${hdr})-{c}${hdr}*{CS}!$H${cr['req']})/($A{r}-{c}${hdr})/(1+$A{r})^Valuation!$C${vr['ttv']})/Valuation!$C${vr['shs']})")
        put(S, f"{c}{r}", fm, USD2)
    r += 1
S.cell(row=r, column=1, value="Current share price for comparison:").font = f("595959", italic=True, size=9)
put(S, f"C{r}", f"=Valuation!C{vr['px']}", USD2)
r += 2
section(S, r, "Table 2 (live): Justified P/TBV value per share - ROTCE vs. cost of equity", 10); r += 1
put(S, f"A{r}", "ROTCE ↓ / Cost of equity →", color=BLACK, bold=True)
kes2 = [0.085, 0.09, 0.095, 0.10, 0.105, 0.11, 0.115]
for j, k in enumerate(kes2):
    put(S, f"{L(2 + j)}{r}", k, PCT, bold=True)
hdr2 = r
r += 1
for ro in [0.14, 0.16, 0.18, 0.20, 0.22, 0.24]:
    put(S, f"A{r}", ro, PCT, bold=True)
    S[f"A{r}"].alignment = Alignment(horizontal="right")
    for j in range(len(kes2)):
        c = L(2 + j)
        put(S, f"{c}{r}", f"=({'$A'}{r}-Valuation!$C${vr['g']})/({c}${hdr2}-Valuation!$C${vr['g']})*Valuation!$C${vr['tbv0']}", USD2)
    r += 1
r += 1
sr["scen"] = r  # static tables appended by run_scenarios.py
S.freeze_panes = "A4"

# =====================================================================
# COVER + SOURCES
# =====================================================================
CV = wb.create_sheet("Cover", 0)
setup(CV, [44, 18, 16, 16, 16, 60])
CV["A1"] = "JPMorgan Chase & Co. (NYSE: JPM)"
CV["A1"].font = f(NAVY, True, 18)
CV["A2"] = "Bank Financial Model - Forecast, Capital & Cash Flows, Valuation and Risk Assessment"
CV["A2"].font = f(NAVY, True, 12)
CV["A3"] = "Prepared by Kanishk | MSBA candidate, USC Marshall | Model date: November 20, 2026 | $ in millions unless noted"
CV["A3"].font = f("595959", italic=True, size=9)
r = 5
section(CV, r, "Key outputs (live - follow the active scenario)", 6); r += 1


def kv(text, formula, fmt, bold=False, fill=None):
    global r
    label(CV, r, text, bold)
    put(CV, f"B{r}", formula, fmt, bold=bold, fill=fill)
    r += 1


kv("Active scenario", f"={AS}!C5", None, True, LIGHT)
kv("Current share price ($)", f"=Valuation!C{vr['px']}", USD2)
kv("FCFE DCF value per share ($)", f"=Valuation!C{vr['dcfpx']}", USD2)
kv("Residual income value per share ($)", f"=Valuation!C{vr['ripx']}", USD2)
kv("Justified P/TBV value per share ($)", f"=Valuation!C{vr['ptbvpx']}", USD2)
kv("Blended value per share ($)", f"=Valuation!C{vr['blend']}", USD2, True, LIGHT)
kv("Upside / (downside) vs. current price", f"=Valuation!D{vr['blend']}", PCT, True)
kv("Implied rating", f"=Valuation!C{vr['rating']}", None, True)
kv("Cost of equity", f"=Valuation!C{vr['ke']}", PCT2)
r += 1
header(CV, r, ["Forecast snapshot", "FY2025A", "FY2026E", "FY2027E", "FY2031E", ""]); r += 1
for text, key, fmt, sheet, store in (
    ("Total net revenue", "rev", USD, "Forecast", fr), ("Net income", "ni", USD, "Forecast", fr),
    ("Diluted EPS ($)", "eps", USD2, "Forecast", fr), ("ROTCE", "rotce", PCT, "Forecast", fr),
    ("CET1 ratio (period-end)", "cr", PCT, CS, cr), ("Capital returned to common (FY26E: 4Q only)", "ret", USD, CS, cr),
):
    label(CV, r, text)
    for col, src in zip("BCDE", "BCDH"):
        put(CV, f"{col}{r}", f"={sheet}!{src}{store[key]}", fmt)
    r += 1
r += 1
section(CV, r, "Model map", 6); r += 1
for tab, desc in [
    ("Historicals", "FY2020-FY2025 + 9M-2026 reported data from SEC filings, with a net-income tie-out check and key bank ratios"),
    ("Assumptions", "Scenario selector, 4Q26 bridge, and Base/Bull/Bear drivers per year (+ shock cells for sensitivity)"),
    ("Forecast", "Income statement FY2026E-FY2031E: NII, fee revenue, expenses, credit provisions, taxes, EPS, ROE/ROTCE"),
    ("Capital & Cash Flow", "CET1 capital roll-forward, RWA, FCFE (distributable capital), dividends, buybacks, share count, TBV"),
    ("Valuation", "FCFE DCF, residual income and justified P/TBV; blended value and implied rating"),
    ("Scenarios & Sensitivity", "Risk assessment: live sensitivity grids, scenario comparison, CET1 stress path, tornado chart, risk register"),
    ("Sources", "Every hard-coded number and where it came from"),
]:
    label(CV, r, tab, True)
    CV.cell(row=r, column=2, value=desc).font = f()
    r += 1
r += 1
section(CV, r, "Methodology & AI-assisted workflow", 6); r += 1
for t in [
    "1. Data: Claude (Anthropic AI) pulled JPM's reported financials directly from the SEC EDGAR XBRL API and earnings filings, then reconciled them",
    "   (revenue = NII + noninterest revenue; net income ties to reported to the dollar - see the check row on Historicals).",
    "2. Bank-specific design: revenue is modelled as balance sheet x yield (NII) plus fees; credit risk via provision / average loans;",
    "   cash flow is the capital in excess of a target CET1 ratio (FCFE), because operating cash flow is meaningless for a bank.",
    "3. Build: the workbook was generated programmatically in Python (openpyxl). Every forecast and valuation cell is a live Excel formula.",
    "4. Risk: Base / Bull / Bear scenarios (Bear ~ 2020-style credit shock + fee downturn + NIM compression), live sensitivity grids,",
    "   and a tornado analysis produced by re-running the full model with each driver shocked (automated with Python + LibreOffice).",
    "Colour code: blue = hard-coded input | black = formula | green = link to another sheet | yellow fill = key lever.",
    "For educational / portfolio purposes only - not investment advice.",
]:
    CV.cell(row=r, column=1, value=t).font = f(size=9)
    r += 1

SO = wb.create_sheet("Sources")
setup(SO, [46, 110])
title(SO, "Sources")
header(SO, 3, ["Data", "Source"])
srcs = [
    ("FY2019-FY2025 income statement, assets, deposits, equity, goodwill, EPS, shares, DPS, buybacks, operating cash flow",
     "SEC EDGAR XBRL company-concept API, CIK 0000019617, e.g. https://data.sec.gov/api/xbrl/companyconcept/CIK0000019617/us-gaap/InterestIncomeExpenseNet.json"),
    ("FY2020-FY2022 loans, common equity, CET1 ratios", "JPMorgan Chase 2022 Annual Report / 10-K Financial Highlights, https://www.sec.gov/Archives/edgar/data/19617/000001961723000286/jpmc2022ar.pdf"),
    ("FY2023-FY2025 loans, common equity, CET1 ratios", "JPMorgan Chase 2025 Annual Report Financial Highlights, https://www.jpmorganchase.com/content/dam/jpmc/jpmorgan-chase-and-co/investor-relations/documents/financial-highlights-2025.pdf"),
    ("YE2025 CET1 capital (~$288B), TBVPS $107.56", "4Q25 earnings release / call coverage (The Asian Banker; Barchart), January 2026"),
    ("9M-2026 income statement; 9-30-26 balance sheet, shares outstanding, CET1 14.0%",
     "JPM Form 10-Q for the quarter ended Sept 30, 2026, https://www.sec.gov/Archives/edgar/data/0000019617/000162828026069523/jpm-20260930.htm"),
    ("3Q26 CET1 capital ~$304B, BVPS $132.30, TBVPS $112.64, $1.65 dividend",
     "3Q26 earnings release (8-K Ex. 99.1), https://www.sec.gov/Archives/edgar/data/0000019617/000162828026065582/a3q26erfexhibit991narrative.htm"),
    ("4Q26 NII guidance (~$28B)", "3Q26 earnings call highlights, MarketBeat, Oct 13, 2026"),
    ("Share price $338.21 (Nov 16, 2026 close)", "MarketChameleon JPM summary, https://marketchameleon.com/Overview/JPM/Summary/"),
    ("10-yr Treasury 5.12% (Nov 16, 2026)", "MarketXLS / FRED, https://marketxls.com/indicators/10-year-treasury-yield"),
    ("Equity beta 1.02 (5Y monthly)", "Yahoo Finance, provided by user on Nov 20, 2026"),
    ("Equity risk premium, scenario drivers, 4Q26 estimates", "Analyst assumptions (documented on Assumptions and Valuation tabs)"),
]
for i, (a, b) in enumerate(srcs):
    SO.cell(row=4 + i, column=1, value=a).font = f()
    SO.cell(row=4 + i, column=2, value=b).font = f()
    SO.cell(row=4 + i, column=1).alignment = Alignment(wrap_text=True, vertical="top")
    SO.cell(row=4 + i, column=2).alignment = Alignment(wrap_text=True, vertical="top")

for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if c.font and c.font.name != FONT:
                c.font = Font(name=FONT, size=c.font.size or 10, bold=c.font.bold, italic=c.font.italic, color=c.font.color)
    ws.sheet_properties.tabColor = NAVY
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

import json
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(ROOT, "model")
os.makedirs(OUT_DIR, exist_ok=True)
# row map lets run_scenarios.py find each line item without hard-coded cell addresses
with open(os.path.join(OUT_DIR, "layout.json"), "w") as fh:
    json.dump({"ar": ar, "fr": fr, "cr": cr, "vr": vr, "sr": sr, "hr": hr}, fh)
wb.calculation.fullCalcOnLoad = True  # Excel computes every formula when the file opens
out = os.path.join(OUT_DIR, "JPM_Financial_Model.xlsx")
wb.save(out)
print("saved", out)
