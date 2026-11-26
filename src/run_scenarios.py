"""Re-run the full model under each scenario and each tornado shock, then write the
static risk-assessment tables and charts into the "Scenarios & Sensitivity" tab.

Each run copies the workbook, changes one input, recalculates it with LibreOffice
(see recalc.py) and reads the outputs back. Run after build_model.py:
    python src/run_scenarios.py
"""
import json, os, sys
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from recalc import recalc  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "model", "JPM_Financial_Model.xlsx")
TMP = os.path.join(ROOT, "model", "_scenario_run.xlsx")
R = json.load(open(os.path.join(ROOT, "model", "layout.json")))
ar, fr, cr, vr, sr = R["ar"], R["fr"], R["cr"], R["vr"], R["sr"]
FY = list("DEFGH")


def evaluate(changes):
    """changes: list of (sheet, cell, value). Returns dict of outputs."""
    wb = load_workbook(SRC)
    for sh, cell, val in changes:
        wb[sh][cell].value = val
    wb.save(TMP)
    recalc(TMP)
    v = load_workbook(TMP, data_only=True)
    Fo, Ca, Va = v["Forecast"], v["Capital & Cash Flow"], v["Valuation"]
    g = lambda ws, c, rr: ws[f"{c}{rr}"].value
    o = {
        "rev27": g(Fo, "D", fr["rev"]), "rev31": g(Fo, "H", fr["rev"]),
        "ni27": g(Fo, "D", fr["ni"]), "ni31": g(Fo, "H", fr["ni"]),
        "eps27": g(Fo, "D", fr["eps"]), "eps31": g(Fo, "H", fr["eps"]),
        "rotce": sum(g(Fo, c, fr["rotce"]) for c in FY) / 5,
        "prov": sum(g(Fo, c, fr["prov"]) for c in FY),
        "cet1min": min(g(Ca, c, cr["cr"]) for c in FY),
        "bufmin": min(g(Ca, c, cr["buf"]) for c in FY),
        "ret": sum(g(Ca, c, cr["ret"]) for c in FY),
        "dcf": g(Va, "C", vr["dcfpx"]), "ri": g(Va, "C", vr["ripx"]), "ptbv": g(Va, "C", vr["ptbvpx"]),
        "blend": g(Va, "C", vr["blend"]), "up": g(Va, "D", vr["blend"]),
        "cet1path": [g(Ca, c, cr["cr"]) for c in "CDEFGH"],
        "nipath": [g(Fo, c, fr["ni"]) for c in "BCDEFGH"],
        "regmin": g(Ca, "D", cr["min"]),
        "nic": [g(Ca, c, cr["cni"]) for c in FY], "ret_path": [g(Ca, c, cr["ret"]) for c in FY],
        "rwa": [g(Ca, c, cr["rwa"]) for c in FY], "cet1_26": g(Ca, "C", cr["ceop"]),
    }
    return o


scen = {}
for i, name in ((1, "Base"), (2, "Bull"), (3, "Bear")):
    scen[name] = evaluate([("Assumptions", "C4", i)])
    print(name, round(scen[name]["blend"], 2), round(scen[name]["cet1min"], 4))

base = scen["Base"]
rf0 = load_workbook(SRC)["Valuation"][f"C{vr['rf']}"].value
g0 = load_workbook(SRC)["Valuation"][f"C{vr['g']}"].value
shocks = [
    ("NII yield (+/- 10 bps)", "Assumptions", f"C{ar['yld']}", -0.0010, 0.0010, 0),
    ("Fee revenue growth (+/- 2 pts/yr)", "Assumptions", f"C{ar['nirg']}", -0.02, 0.02, 0),
    ("Efficiency ratio (+/- 2 pts)", "Assumptions", f"C{ar['eff']}", 0.02, -0.02, 0),
    ("Provision rate (+/- 25 bps of loans)", "Assumptions", f"C{ar['prv']}", 0.0025, -0.0025, 0),
    ("Asset growth (+/- 2 pts/yr)", "Assumptions", f"C{ar['ag']}", -0.02, 0.02, 0),
    ("Target CET1 ratio (+/- 50 bps)", "Assumptions", f"C{ar['tgt']}", 0.005, -0.005, 0),
    ("Cost of equity (+/- 50 bps)", "Valuation", f"C{vr['rf']}", rf0 + 0.005, rf0 - 0.005, None),
    ("Terminal growth (+/- 50 bps)", "Valuation", f"C{vr['g']}", g0 - 0.005, g0 + 0.005, None),
]
torn = []
for lab, sh, cell, lo, hi, _ in shocks:
    plo = evaluate([(sh, cell, lo)])["blend"]
    phi = evaluate([(sh, cell, hi)])["blend"]
    torn.append((lab, plo, phi))
    print(lab, round(plo, 1), round(phi, 1))
torn.sort(key=lambda t: abs(t[2] - t[1]), reverse=True)

# ------------------------------------------------------------------ write
wb = load_workbook(SRC)
S = wb["Scenarios & Sensitivity"]
FONT = "Arial"
NAVY, LIGHT = "1F3864", "D9E1F2"
STATIC = "404040"
USD = '$#,##0;($#,##0);"-"'
USD2 = '$#,##0.00;($#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'


def sec(row, text):
    for col in range(1, 11):
        S.cell(row=row, column=col).fill = PatternFill("solid", fgColor=LIGHT)
    S.cell(row=row, column=1, value=text).font = Font(name=FONT, bold=True, color=NAVY)


def hdr(row, labels):
    for i, t in enumerate(labels):
        c = S.cell(row=row, column=1 + i, value=t)
        c.font = Font(name=FONT, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="left" if i == 0 else "center", wrap_text=True)


def val(row, col, v, fmt, bold=False):
    c = S.cell(row=row, column=col, value=v)
    c.font = Font(name=FONT, color=STATIC, bold=bold)
    c.number_format = fmt
    c.alignment = Alignment(horizontal="right")


def text(row, col, t, bold=False, italic=False, size=10, color="000000"):
    c = S.cell(row=row, column=col, value=t)
    c.font = Font(name=FONT, bold=bold, italic=italic, size=size, color=color)
    return c


r = sr["scen"]
text(r, 1, "Tables 3-6 below are static snapshots (dark-grey numbers) produced by re-running the full model for each case; "
           "re-run run_scenarios.py after changing inputs.", italic=True, size=8, color="595959")
r += 1
sec(r, "Table 3 (static): Scenario comparison - FY2027E-FY2031E")
r += 1
hdr(r, ["Metric", "Base", "Bull", "Bear", "Bear vs. Base"])
r += 1
rows = [
    ("Total net revenue FY2027E ($mm)", "rev27", USD), ("Total net revenue FY2031E ($mm)", "rev31", USD),
    ("Net income FY2027E ($mm)", "ni27", USD), ("Net income FY2031E ($mm)", "ni31", USD),
    ("Diluted EPS FY2027E ($)", "eps27", USD2), ("Diluted EPS FY2031E ($)", "eps31", USD2),
    ("Average ROTCE FY27-31", "rotce", PCT), ("Cumulative credit provisions FY27-31 ($mm)", "prov", USD),
    ("Lowest CET1 ratio FY27-31", "cet1min", PCT), ("Lowest buffer over reg. minimum (bps)", "bufmin", "#,##0;(#,##0)"),
    ("Cumulative capital returned FY27-31 ($mm)", "ret", USD),
    ("FCFE DCF value / share ($)", "dcf", USD2), ("Residual income value / share ($)", "ri", USD2),
    ("Justified P/TBV value / share ($)", "ptbv", USD2), ("Blended value / share ($)", "blend", USD2),
    ("Upside / (downside) vs. current price", "up", PCT),
]
for labl, k, fmt in rows:
    text(r, 1, labl, bold=k in ("blend", "up"))
    for j, nm in enumerate(("Base", "Bull", "Bear")):
        val(r, 2 + j, scen[nm][k], fmt, bold=k in ("blend", "up"))
    if fmt in (PCT,) or k == "bufmin":
        val(r, 5, scen["Bear"][k] - base[k], fmt if k == "bufmin" else '+0.0%;-0.0%;"-"')
    else:
        val(r, 5, (scen["Bear"][k] / base[k] - 1) if base[k] else 0, '+0.0%;-0.0%;"-"')
    r += 1
text(r, 1, "Bear vs. Base: % change for $ values; percentage-point / bps change for ratios.", italic=True, size=8, color="595959")
r += 2

sec(r, "Scenario definitions")
r += 1
for nm, d in [
    ("Base", "Consensus-like path: ~4% balance-sheet growth, NII yield ~2.12%, fees +4-5%, efficiency ~54%, provisions ~0.75% of loans, CET1 target 13.5%."),
    ("Bull", "Strong economy & capital markets: 6% asset growth, yield ~2.2%, fees +5-8%, efficiency ~52%, benign credit (~0.6%), CET1 target 13.0%."),
    ("Bear", "2027 recession stress: balance sheet shrinks 2%, Fed cuts compress NII yield to ~1.92%, fees -8%, efficiency 58%, provisions spike to 1.6% of loans (~2020 COVID level), higher risk weights, CET1 target 14.0%."),
]:
    text(r, 1, nm, bold=True)
    text(r, 2, d, size=9)
    r += 1
r += 1

sec(r, "Table 4 (static): Capital stress path - CET1 ratio by scenario")
r += 1
hdr(r, ["Scenario", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"])
cet_hdr = r
r += 1
cet_first = r
# no-management-action stress: Bear earnings & RWA, but Base-case dividends + buybacks continue
bear = scen["Bear"]
cap = bear["cet1_26"]
nma = [bear["cet1path"][0]]
for ni, dist, rwa in zip(bear["nic"], base["ret_path"], bear["rwa"]):
    cap = cap + ni - dist
    nma.append(cap / rwa)
scen["Bear (no mgmt action)"] = {"cet1path": nma}
scen["Bear (no mgmt action)"]["min"] = min(nma[1:])
for nm in ("Base", "Bull", "Bear", "Bear (no mgmt action)"):
    text(r, 1, nm)
    for j, v in enumerate(scen[nm]["cet1path"]):
        val(r, 2 + j, v, PCT)
    r += 1
text(r, 1, "Regulatory minimum (approx.)")
for j in range(6):
    val(r, 2 + j, base["regmin"], PCT)
cet_last = r
r += 2
text(r - 1, 1, "Bear (managed): JPM cuts buybacks, CET1 holds at target. No mgmt action: Bear earnings & RWA but Base-case payouts continue.", italic=True, size=8, color="595959")
text(r, 1, "Capital returned to common by scenario ($mm)", bold=True)
r += 1
for nm in ("Base", "Bull", "Bear"):
    text(r, 1, nm)
    for j, v in enumerate(scen[nm]["ret_path"]):
        val(r, 3 + j, v, USD)
    r += 1
r += 1
ch = LineChart()
ch.title = "CET1 ratio by scenario vs. regulatory minimum"
ch.y_axis.title = "CET1 ratio"
ch.y_axis.number_format = "0%"
ch.height, ch.width = 7.5, 16
data = Reference(S, min_col=1, max_col=7, min_row=cet_first, max_row=cet_last)
ch.add_data(data, from_rows=True, titles_from_data=True)
ch.set_categories(Reference(S, min_col=2, max_col=7, min_row=cet_hdr))
colors = ["2A78D6", "EB6834", "1BAF7A", "1BAF7A", "8C8C8C"]
for s, colr in zip(ch.series, colors):
    s.graphicalProperties.line.solidFill = colr
    s.graphicalProperties.line.width = 25400
    s.smooth = False
ch.series[3].graphicalProperties.line.dashStyle = "sysDot"
ch.series[4].graphicalProperties.line.dashStyle = "dash"
ch.y_axis.scaling.min = 0.10
ch.roundedCorners = False
ch.legend.position = "b"
ch.legend.overlay = False
ch.title.overlay = False
ch.x_axis.delete = False
ch.y_axis.delete = False
ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="E1E0D9"))
S.add_chart(ch, f"A{r}")
r += 17

sec(r, "Table 5 (static): Tornado - blended value per share when each driver is shocked (Base scenario)")
r += 1
hdr(r, ["Driver (shock applied to every forecast year)", "Downside case", "Upside case", "Downside vs. base", "Upside vs. base", "Swing"])
S.row_dimensions[r].height = 30
t_hdr = r
r += 1
text(r, 1, "Base case blended value / share", bold=True)
val(r, 2, base["blend"], USD2, True)
r += 1
t_first = r
for lab, plo, phi in torn:
    text(r, 1, lab)
    val(r, 2, plo, USD2)
    val(r, 3, phi, USD2)
    val(r, 4, plo - base["blend"], '+$#,##0.00;-$#,##0.00;"-"')
    val(r, 5, phi - base["blend"], '+$#,##0.00;-$#,##0.00;"-"')
    val(r, 6, phi - plo, USD2)
    r += 1
t_last = r - 1
r += 3
bc = BarChart()
bc.type = "bar"
bc.grouping = "clustered"
bc.overlap = 100
bc.title = "Tornado: change in value per share vs. base ($)"
bc.height, bc.width = 11, 18
bc.add_data(Reference(S, min_col=4, max_col=5, min_row=t_first, max_row=t_last), titles_from_data=False)
# header row is not directly above data (base row in between) -> set titles manually
from openpyxl.chart.series import SeriesLabel
bc.series[0].tx = SeriesLabel(v="Downside shock")
bc.series[1].tx = SeriesLabel(v="Upside shock")
for s, colr in zip(bc.series, ["EB6834", "2A78D6"]):
    s.graphicalProperties.solidFill = colr
    s.graphicalProperties.line.solidFill = colr
bc.set_categories(Reference(S, min_col=1, min_row=t_first, max_row=t_last))
bc.y_axis.number_format = '$#,##0'
bc.x_axis.scaling.orientation = "maxMin"
bc.x_axis.tickLblPos = "low"
bc.roundedCorners = False
bc.legend.position = "b"
bc.legend.overlay = False
bc.title.overlay = False
bc.x_axis.delete = False
bc.y_axis.delete = False
bc.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="E1E0D9"))
S.add_chart(bc, f"A{r}")
r += 24

sec(r, "Table 6: Risk register - key risks and how the model captures them")
r += 1
hdr(r, ["Risk", "What could happen", "", "", "Model lever", "", "Scenario treatment", "", "", ""])
for a, b in ((2, 4), (5, 6), (7, 10)):
    S.merge_cells(start_row=r, start_column=a, end_row=r, end_column=b)
r += 1
for risk, what, lever, treat in [
    ("Interest-rate / NIM risk", "Fed cuts or deposit repricing compress net interest margin", "NII / avg assets", "Bear: 1.92-2.05% vs. 2.12% base"),
    ("Credit risk", "Recession lifts card & wholesale losses; Apple Card portfolio seasoning", "Provision / avg loans", "Bear: 1.6% of loans in 2027 (~2020 level)"),
    ("Capital-markets risk", "IB fee and trading revenue fall in a risk-off market", "Fee revenue growth", "Bear: -8% in 2027"),
    ("Operating leverage", "Tech / AI investment and wage inflation outpace revenue", "Efficiency ratio", "Bear: 58% vs. 54.5% base"),
    ("Regulatory capital", "Basel III endgame / GSIB surcharge raise required CET1 and RWA", "Target CET1, RWA density", "Bear: 14.0% target, 44.5% density"),
    ("Valuation / rates", "Higher Treasury yields raise cost of equity", "Rf, beta, ERP, terminal g", "Tables 1, 2 and 5"),
]:
    text(r, 1, risk, bold=True).alignment = Alignment(vertical="top")
    for col, endc, t in ((2, 4, what), (5, 6, lever), (7, 10, treat)):
        S.merge_cells(start_row=r, start_column=col, end_row=r, end_column=endc)
        text(r, col, t, size=9).alignment = Alignment(wrap_text=True, vertical="top")
    S.row_dimensions[r].height = 26
    r += 1

wb.calculation.fullCalcOnLoad = True
wb.save(SRC)
if os.path.exists(TMP):
    os.remove(TMP)
with open(os.path.join(ROOT, "model", "scenario_results.json"), "w") as fh:
    json.dump({"scen": scen, "torn": torn}, fh, indent=1, default=str)
print("written", SRC)
