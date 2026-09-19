"""Build the LSEG FY2025 corporate valuation model using Python only.

Dependency: XlsxWriter (bundled with the Codex Python runtime).
Run: python build_lseg_valuation_python_only.py
"""
from __future__ import annotations

from pathlib import Path
import math
from statistics import median
import xlsxwriter


OUT = Path(__file__).resolve().parent / "LSEG_Corporate_Valuation_2025_Python_Only.xlsx"
VALUATION_DATE = "2026-02-26"
PROXY_PUBLICATION_DATE = "2026-02-27"
FINANCIAL_YEAR_END = "2025-12-31"
REFERENCE_PRICE = 81.2660  # Purchases on 26 Feb; RNS published 27 Feb. Retrospective proxy, not close.
SHARE_COUNT = 505.332519   # prospective post-cancellation voting rights in the 27 Feb RNS, millions
NET_DEBT = 8175.0
RISK_FREE_RATE = 0.045
EQUITY_RISK_PREMIUM = 0.045
LEVERED_BETA = 0.90
PRE_TAX_COST_OF_DEBT = 0.045
TAX_RATE = 0.245
MARKET_CAP = REFERENCE_PRICE * SHARE_COUNT
DEBT_WEIGHT = NET_DEBT / (NET_DEBT + MARKET_CAP)
EQUITY_WEIGHT = MARKET_CAP / (NET_DEBT + MARKET_CAP)
COST_OF_EQUITY = RISK_FREE_RATE + LEVERED_BETA * EQUITY_RISK_PREMIUM
AFTER_TAX_COST_OF_DEBT = PRE_TAX_COST_OF_DEBT * (1 - TAX_RATE)
WACC = EQUITY_WEIGHT * COST_OF_EQUITY + DEBT_WEIGHT * AFTER_TAX_COST_OF_DEBT
CAPEX_RATES = [0.095, 0.090, 0.085, 0.080, 0.080]


def dcf_case(growth, target_margin, terminal_growth, wacc=WACC):
    revenue, margin = 8986.0, 4523.0 / 8986.0
    rows = []
    for year, g in enumerate(growth, 1):
        revenue *= 1 + g
        margin += (target_margin - 4523.0 / 8986.0) / 5
        ebitda = revenue * margin
        da = revenue * 0.113
        ebit = ebitda - da
        tax = ebit * TAX_RATE
        capex = revenue * CAPEX_RATES[year - 1]
        nwc = revenue * 0.015
        fcf = ebit - tax + da - capex - nwc
        discount = 1 / (1 + wacc) ** year
        rows.append((revenue, margin, ebitda, da, ebit, tax, capex, nwc, fcf, discount, fcf * discount))
    pv_forecast = sum(x[10] for x in rows)
    terminal = rows[-1][8] * (1 + terminal_growth) / (wacc - terminal_growth)
    enterprise = pv_forecast + terminal * rows[-1][9]
    equity = enterprise - NET_DEBT
    return rows, enterprise, equity / SHARE_COUNT


def build():
    wb = xlsxwriter.Workbook(OUT)
    wb.set_calc_mode("auto")
    wb.set_properties({
        "title": "LSEG Corporate Valuation 2025",
        "subject": "AI-Supported Valuation and Corporate Finance Decision-Making",
        "author": "David",
        "comments": "Python-only formula-driven valuation model",
    })

    navy, blue, light, pale = "#0B1F3A", "#1F4E78", "#D9EAF7", "#EEF4F8"
    yellow, green, red, white = "#FFF2CC", "#E2F0D9", "#FCE4D6", "#FFFFFF"
    base_font = {"font_name": "Arial", "font_size": 10}
    f = {
        "title": wb.add_format({**base_font, "bold": True, "font_size": 16, "font_color": white, "bg_color": navy, "valign": "vcenter"}),
        "meta": wb.add_format({**base_font, "italic": True, "font_color": blue, "bg_color": pale}),
        "section": wb.add_format({**base_font, "bold": True, "font_color": white, "bg_color": blue, "align": "left", "valign": "vcenter"}),
        "header": wb.add_format({**base_font, "bold": True, "bg_color": light, "bottom": 1, "text_wrap": True, "valign": "vcenter"}),
        "text": wb.add_format(base_font),
        "wrap": wb.add_format({**base_font, "text_wrap": True, "valign": "top"}),
        "input": wb.add_format({**base_font, "font_color": "#0000FF", "bg_color": yellow}),
        "input_pct": wb.add_format({**base_font, "font_color": "#0000FF", "bg_color": yellow, "num_format": "0.0%"}),
        "input_beta": wb.add_format({**base_font, "font_color": "#0000FF", "bg_color": yellow, "num_format": "0.00x"}),
        "formula": wb.add_format(base_font),
        "link": wb.add_format({**base_font, "font_color": "#008000"}),
        "num": wb.add_format({**base_font, "num_format": "#,##0;[Red](#,##0);-"}),
        "gbp": wb.add_format({**base_font, "num_format": "£#,##0;[Red](£#,##0);-"}),
        "ps": wb.add_format({**base_font, "num_format": "£0.00;[Red](£0.00);-"}),
        "pct": wb.add_format({**base_font, "num_format": "0.0%;[Red](0.0%);-"}),
        "multiple": wb.add_format({**base_font, "num_format": "0.0x;[Red](0.0x);-"}),
        "total": wb.add_format({**base_font, "bold": True, "top": 1, "num_format": "#,##0;[Red](#,##0);-"}),
        "total_ps": wb.add_format({**base_font, "bold": True, "top": 1, "num_format": "£0.00"}),
        "card": wb.add_format({**base_font, "bold": True, "bg_color": green, "border": 1, "border_color": blue}),
        "pass": wb.add_format({**base_font, "bold": True, "bg_color": green, "align": "center"}),
        "fail": wb.add_format({**base_font, "bold": True, "bg_color": red, "align": "center"}),
    }

    names = ["Cover", "Assumptions", "Historical", "DCF", "Comps", "Scenarios", "Sensitivities", "AI Review", "Critical Review", "Checks", "Sources"]
    ws = {n: wb.add_worksheet(n) for n in names}
    for s in ws.values():
        s.hide_gridlines(2)
        s.freeze_panes(4, 0)
        s.set_default_row(15)

    def title(s, text, last_col=7):
        s.merge_range(0, 0, 0, last_col, text, f["title"])
        s.set_row(0, 28)

    def section(s, row, text, last_col=7):
        s.merge_range(row, 0, row, last_col, text, f["section"])
        s.set_row(row, 20)

    # Shared assumptions and calculated values.
    cases = {
        "Downside": ([0.045, 0.045, 0.04, 0.04, 0.035], 0.515, 0.020),
        "Base": ([0.070, 0.065, 0.06, 0.06, 0.055], 0.540, 0.025),
        "Upside": ([0.085, 0.080, 0.075, 0.07, 0.065], 0.560, 0.030),
    }
    base_rows, base_ev, base_value = dcf_case(*cases["Base"])
    peer_median = 20.198509252912952
    comp_value = (peer_median * 3506 - NET_DEBT) / SHARE_COUNT
    dcf_weight, comps_weight = 0.70, 0.30
    blended = base_value * dcf_weight + comp_value * comps_weight

    # Cover.
    s = ws["Cover"]
    title(s, "AI-Supported Corporate Valuation: LSEG and the Limits of Automated Financial Analysis", 9)
    s.merge_range("A3:J3", "FY2025. Valuation reference: 26 February 2026. Price/share proxy disclosed 27 February. GBP millions except per share.", f["meta"])
    section(s, 4, "Valuation summary", 9)
    s.write_row("A6", ["Method", "Implied value / share", "Reference price", "Upside / (downside)"], f["header"])
    summary = [("DCF — principal method", "='DCF'!$B$36", base_value), ("Comparable companies — cross-check", "='Comps'!$B$21", comp_value), ("Weighted indication", "=B7*'Assumptions'!$B$40+B8*'Assumptions'!$B$41", blended)]
    for r, (label, formula, value) in enumerate(summary, 6):
        s.write(r, 0, label, f["text"])
        s.write_formula(r, 1, formula, f["ps"], value)
        s.write_formula(r, 2, "='Assumptions'!$B$8", f["ps"], REFERENCE_PRICE)
        s.write_formula(r, 3, f"=B{r+1}/C{r+1}-1", f["pct"], value / REFERENCE_PRICE - 1)
    s.write("A10", "David's judgement", f["card"])
    s.write_formula("B10", '=IF(D7>=15%,"POSITIVE / WATCH",IF(D7<=-15%,"NEGATIVE / AVOID","NEUTRAL / WATCH"))', f["card"], "POSITIVE / WATCH")
    s.write_formula("C10", "='Assumptions'!$B$8", f["card"], REFERENCE_PRICE)
    s.write_formula("D10", "=D7", f["card"], base_value / REFERENCE_PRICE - 1)
    section(s, 11, "Investment and transaction judgement", 9)
    judgement = ("Investment: the live judgement is a valuation signal rather than a direct trading recommendation. At ≥15% DCF upside it shows POSITIVE / WATCH; between -15% and +15% it shows NEUTRAL / WATCH; and at ≤-15% it shows NEGATIVE / AVOID. DCF is the principal method and comparables are a reasonableness check. "
                 "The conclusion is assumption-sensitive and remains subject to execution, leverage, forecast, dilution and terminal-value risk.\n\n"
                 "Transaction: ATTRACTIVE STRATEGIC ASSET, PRICE-DISCIPLINED. A control premium is justified only with credible, quantified synergies; "
                 "a highly levered acquisition would carry material integration and balance-sheet risk. This is a student case study, not investment advice.")
    s.merge_range("A13:J17", judgement, f["wrap"])
    section(s, 18, "How to use the model", 9)
    s.merge_range("A20:J23", "Select a case in Assumptions!B5 and edit yellow inputs. Checks tests calculation consistency, not source completeness or investment quality. This is a retrospective FY2025 study: results were released 26 February 2026 and the price/share proxy was disclosed 27 February. It is not a strict information-available-at-26-February backtest or a 2025-only information set. See Sources for input limitations.", f["wrap"])
    s.set_column("A:A", 35); s.set_column("B:D", 20); s.set_column("E:J", 16)

    # Assumptions.
    s = ws["Assumptions"]
    title(s, "Assumptions and Scenario Control")
    s.merge_range("A3:H3", "Blue-font cells are editable. Forecast assumptions are explicitly separated from reported facts.", f["meta"])
    section(s, 3, "Model control and market reference")
    controls = [("Scenario", "Base"), ("Valuation date", VALUATION_DATE), ("Currency / units", "GBP millions"), ("Reference price (£)", REFERENCE_PRICE), ("Voting shares used as proxy (m)", SHARE_COUNT), ("Operating net debt", NET_DEBT), ("Risk-free rate", RISK_FREE_RATE), ("Equity risk premium", EQUITY_RISK_PREMIUM)]
    for r, (lab, val) in enumerate(controls, 4):
        s.write(r, 0, lab, f["text"])
        form = f["input"] if r in (4, 10, 11) else (f["ps"] if r == 7 else f["num"] if r in (8, 9) else f["text"])
        s.write(r, 1, val, form)
    s.data_validation("B5", {"validate": "list", "source": ["Base", "Upside", "Downside"]})
    s.merge_range("D5:H8", "Price: £81.266 = 8,126.60p / 100. RNS published 27 Feb 2026 reports 34,592 shares bought on 26 Feb. Retained as a verified programme VWAP proxy, not the official close or all-market VWAP. Voting rights are the announced post-cancellation denominator, not diluted weighted-average shares. Source S3.", f["wrap"])
    s.merge_range("D10:H12", "The proxy publication is one day after the valuation reference. No claim is made that these disclosed figures were public on 26 Feb. Operating net debt is still the FY2025 balance; no intervening debt/cash roll-forward is modelled.", f["wrap"])
    section(s, 13, "WACC")
    wacc_items = [("Levered beta", LEVERED_BETA), ("Pre-tax cost of debt", PRE_TAX_COST_OF_DEBT), ("Tax rate", TAX_RATE)]
    for r, (lab, val) in enumerate(wacc_items, 14):
        s.write(r, 0, lab, f["text"]); s.write(r, 1, val, f["input_beta"] if r == 14 else f["input_pct"])
    formulas = [
        (17, "Market capitalisation", "=B8*B9", MARKET_CAP, f["gbp"]),
        (18, "Debt weight", "=B10/(B10+B18)", DEBT_WEIGHT, f["pct"]),
        (19, "Equity weight", "=B18/(B10+B18)", EQUITY_WEIGHT, f["pct"]),
        (20, "Cost of equity", "=B11+B15*B12", COST_OF_EQUITY, f["pct"]),
        (21, "After-tax cost of debt", "=B16*(1-B17)", AFTER_TAX_COST_OF_DEBT, f["pct"]),
        (22, "WACC", "=B20*B21+B19*B22", WACC, f["pct"]),
    ]
    for r, lab, formula, value, form in formulas:
        s.write(r, 0, lab, f["text"]); s.write_formula(r, 1, formula, form, value)
    section(s, 24, "Operating cases")
    s.write_row("A26", ["Case", "2026 growth", "2027 growth", "2028 growth", "2029 growth", "2030 growth", "2030 EBITDA margin", "Terminal growth"], f["header"])
    for r, (name, (growth, margin, tg)) in enumerate(cases.items(), 26):
        s.write(r, 0, name, f["text"])
        for c, value in enumerate(growth + [margin, tg], 1): s.write(r, c, value, f["input_pct"])
    s.write(29, 0, "Selected", f["total"])
    for c, col in enumerate("BCDEFGH", 1):
        value = cases["Base"][0][c-1] if c <= 5 else cases["Base"][1] if c == 6 else cases["Base"][2]
        s.write_formula(29, c, f'=INDEX({col}$27:{col}$29,MATCH($B$5,$A$27:$A$29,0))', f["pct"], value)
    section(s, 31, "Cash-flow drivers")
    s.write_row("A33", ["Driver", "Selected assumption", "Rationale"], f["header"])
    drivers = [("D&A as % revenue", .113, "FY2025 adjusted EBITDA less adjusted operating profit, divided by income"), ("Capex as % revenue", .095, "Explicit annual path: 9.5% in 2026, declining to 8.0% by 2029"), ("Change in NWC as % revenue", .015, "Normalised modelling assumption"), ("Cash tax rate", TAX_RATE, "2026 company guidance midpoint")]
    for r, (lab, val, note) in enumerate(drivers, 33):
        s.write(r, 0, lab, f["text"]); s.write(r, 1, val, f["input_pct"]); s.write(r, 2, note, f["wrap"])
    s.write_row("D33", ["2026E", "2027E", "2028E", "2029E", "2030E"], f["header"])
    for c, rate in enumerate(CAPEX_RATES, 3):
        s.write(34, c, rate, f["input_pct"])
    s.write_formula("B35", "=D35", f["pct"], CAPEX_RATES[0])
    s.write(37, 0, "Terminal growth", f["text"]); s.write_formula(37, 1, "=H30", f["pct"], .025); s.write(37, 2, "Selected case", f["text"])
    section(s, 38, "Method weighting and interpretation")
    s.write_row("A40", ["DCF weight", dcf_weight, "Principal method: cash-flow based and linked to operating assumptions"], f["text"])
    s.write("B40", dcf_weight, f["input_pct"])
    s.write_row("A41", ["Comparable-company weight", comps_weight, "Cross-check only because peer business mix and definitions differ"], f["text"])
    s.write("B41", comps_weight, f["input_pct"])
    s.write("A42", "Weight check", f["text"]); s.write_formula("B42", "=SUM(B40:B41)", f["pct"], 1.0)
    s.set_column("A:A", 31); s.set_column("B:B", 18); s.set_column("C:C", 48); s.set_column("D:H", 17)

    # Historical.
    s = ws["Historical"]
    title(s, "Historical Financials and Operating Review", 6)
    s.merge_range("A3:G3", "FY2025 preliminary results are the main source. Exact sections, reclassifications and legacy-input limitations are recorded in Sources.", f["meta"])
    section(s, 4, "Income and cash flow", 6)
    s.write_row("A6", ["GBP m unless stated", "2023A", "2024A", "2025A"], f["header"])
    hist = [("Total income excl. recoveries", 8061, 8494, 8986), ("Growth", None, None, None), ("Adjusted EBITDA", 3801, 4148, 4523), ("Adjusted EBITDA margin", None, None, None), ("Adjusted operating profit", 2909, 3165, 3506), ("D&A / impairment (derived)", None, None, None), ("Reported EBITDA", None, 3945, 4365), ("Change in working capital", None, -50, -419), ("Capex", None, -957, -919), ("Equity free cash flow", None, 2184, 2445), ("Operating net debt", None, 7178, 8175), ("Voting shares at FY2025 reporting (m)", None, None, 510.408075)]
    for r, row in enumerate(hist, 6): s.write_row(r, 0, row, f["num"])
    s.write_formula("C8", "=C7/B7-1", f["pct"], 8494/8061-1); s.write_formula("D8", "=D7/C7-1", f["pct"], 8986/8494-1)
    for col, value in zip("BCD", [3801/8061, 4148/8494, 4523/8986]): s.write_formula(f"{col}10", f"={col}9/{col}7", f["pct"], value)
    for col, value in zip("BCD", [892, 983, 1017]): s.write_formula(f"{col}12", f"={col}9-{col}11", f["num"], value)
    section(s, 19, "AI-supported document review", 6)
    s.write_row("A21", ["Finding", "Evidence", "Valuation implication", "Confidence", "Human check", "Source ID", "Treatment"], f["header"])
    review = [("Revenue quality", "ASV +5.9%; subscription-led data businesses", "Supports durable growth", "High", "Confirmed", "S1", "Base growth anchored to guidance"), ("Margin expansion", "2025 adjusted EBITDA margin 50.3%", "Operating leverage supports FCF", "High", "Confirmed", "S1", "Margin ramps by case"), ("Capital intensity", "2025 capex intensity 10.2%", "Near-term FCF burden eases", "High", "Confirmed", "S1", "Capex declines toward 8%"), ("Leverage", "Operating net debt £8.175bn", "Reduces equity value", "High", "Confirmed", "S1", "Full equity bridge deduction"), ("AI opportunity", "AI-ready data partnerships", "Potential upside, execution risk", "Medium", "Confirmed", "S1", "Upside case only"), ("Model risk", "Forecasts require assumptions", "Avoid false precision", "High", "Confirmed", "S10", "Scenarios and sensitivities")]
    for r, row in enumerate(review, 21): s.write_row(r, 0, row, f["wrap"])
    s.set_column("A:A", 30); s.set_column("B:B", 30); s.set_column("C:C", 32); s.set_column("D:F", 14); s.set_column("G:G", 28)

    # DCF.
    s = ws["DCF"]
    title(s, "Discounted Cash Flow Valuation", 6)
    s.merge_range("A3:G3", "Unlevered FCF | Gordon growth terminal value | Year-end discounting | GBP millions except per share", f["meta"])
    section(s, 4, "Forecast", 6)
    s.write_row("A6", ["GBP m unless stated", "2025A", "2026E", "2027E", "2028E", "2029E", "2030E"], f["header"])
    labels = ["Revenue growth", "Revenue", "EBITDA margin", "Adjusted EBITDA", "D&A / impairment", "EBIT", "Cash tax", "NOPAT", "D&A add-back", "Capex", "Change in NWC", "Unlevered FCF", "Discount period", "Discount factor", "PV of FCF"]
    for r, label in enumerate(labels, 6): s.write(r, 0, label, f["text"])
    actual = [8986/8494-1, 8986, 4523/8986, 4523, 1017, 3506, -859, 2647, 1017, -919, -419, 2326, 0, 1, 2326]
    for r, value in enumerate(actual, 6): s.write(r, 1, value, f["pct"] if r in (6,8) else f["num"])
    growth = cases["Base"][0]
    for i, data in enumerate(base_rows, 2):
        col = chr(65 + i)       # C:G forecast columns
        prev = chr(64 + i)      # B:F preceding-year columns
        assumption_col = chr(64 + i)  # B:F growth assumptions
        revenue, margin, ebitda, da, ebit, tax, capex, nwc, fcf, discount, pv = data
        values = [growth[i-2], revenue, margin, ebitda, da, ebit, -tax, ebit-tax, da, -capex, -nwc, fcf, i-1, discount, pv]
        capex_col = chr(68 + i - 2)  # D:H on Assumptions
        formulas = [f"='Assumptions'!{assumption_col}$30", f"={prev}8*(1+{col}7)", f"=$B$9+('Assumptions'!$G$30-$B$9)*({i-1}/5)", f"={col}8*{col}9", f"={col}8*'Assumptions'!$B$34", f"={col}10-{col}11", f"=-{col}12*'Assumptions'!$B$37", f"={col}12+{col}13", f"={col}11", f"=-{col}8*'Assumptions'!${capex_col}$35", f"=-{col}8*'Assumptions'!$B$36", f"=SUM({col}14:{col}17)", f"={i-1}", f"=1/(1+'Assumptions'!$B$23)^{i-1}", f"={col}18*{col}20"]
        for j, (formula, value) in enumerate(zip(formulas, values), 6): s.write_formula(j, i, formula, f["pct"] if j in (6,8) else f["num"], value)
    section(s, 23, "Terminal value and equity bridge", 6)
    terminal = base_rows[-1][8] * 1.025 / (WACC - .025)
    bridge = [("Terminal year FCF", "=G18", base_rows[-1][8], f["gbp"]), ("Terminal growth", "='Assumptions'!$B$38", .025, f["pct"]), ("WACC", "='Assumptions'!$B$23", WACC, f["pct"]), ("Terminal value", "=B25*(1+B26)/(B27-B26)", terminal, f["gbp"]), ("PV factor", "=G20", base_rows[-1][9], f["formula"]), ("PV of terminal value", "=B28*B29", terminal*base_rows[-1][9], f["gbp"]), ("PV of forecast FCF", "=SUM(C21:G21)", sum(x[10] for x in base_rows), f["gbp"]), ("Enterprise value", "=SUM(B30:B31)", base_ev, f["gbp"]), ("Less: operating net debt", "=-'Assumptions'!$B$10", -NET_DEBT, f["gbp"]), ("Equity value", "=SUM(B32:B33)", base_ev-NET_DEBT, f["gbp"]), ("Voting shares used as proxy (m)", "='Assumptions'!$B$9", SHARE_COUNT, f["num"]), ("DCF value per share", "=B34/B35", base_value, f["ps"]), ("Reference share price", "='Assumptions'!$B$8", REFERENCE_PRICE, f["ps"]), ("Upside / (downside)", "=B36/B37-1", base_value/REFERENCE_PRICE-1, f["pct"])]
    for r, (lab, formula, value, form) in enumerate(bridge, 24): s.write(r, 0, lab, f["text"]); s.write_formula(r, 1, formula, form, value)
    s.set_column("A:A", 30); s.set_column("B:G", 18)

    # Comps.
    s = ws["Comps"]
    title(s, "Comparable-Company Analysis", 10)
    s.merge_range("A3:K3", "Market prices fixed at 26 February 2026; financials are FY2025. EV/adjusted EBIT is used consistently. This is a reasonableness check, not the principal valuation.", f["meta"])
    section(s, 4, "Peer screen", 10)
    headers = ["Company", "Ticker", "Price", "Shares m", "Market cap LC m", "Net debt LC m", "Revenue LC m", "Adjusted EBIT LC m", "EV / Revenue", "EV / Adj. EBIT", "P / E"]
    s.write_row("A6", headers, f["header"])
    peers = [("LSEG", "LSEG.L", REFERENCE_PRICE, SHARE_COUNT, NET_DEBT, 8986, 3506, 19.32), ("Intercontinental Exchange", "ICE", 162.89, 575, 18763, 9931, 5959, 23.30), ("CME Group", "CME", 316.45, 361, -1200, 6520.6, 4524.8, 28.25), ("Nasdaq", "NDAQ", 88.59, 575, 8000, 5249, 2918, 25.46)]
    multiples = []
    for r, (name, ticker, price, shares, debt, revenue, profit, pe) in enumerate(peers, 6):
        market = price * shares; ev = market + debt; ev_rev, ev_profit = ev/revenue, ev/profit; multiples.append(ev_profit)
        s.write_row(r, 0, [name, ticker, price, shares], f["text"])
        s.write_formula(r, 4, f"=C{r+1}*D{r+1}", f["num"], market); s.write(r, 5, debt, f["num"]); s.write(r, 6, revenue, f["num"]); s.write(r, 7, profit, f["num"])
        s.write_formula(r, 8, f"=(E{r+1}+F{r+1})/G{r+1}", f["multiple"], ev_rev); s.write_formula(r, 9, f"=(E{r+1}+F{r+1})/H{r+1}", f["multiple"], ev_profit); s.write(r, 10, pe, f["multiple"])
    s.write(10, 0, "Median", f["total"]); s.write_formula(10, 9, "=MEDIAN(J8:J10)", f["multiple"], peer_median)
    # Exclude LSEG, exactly as the Excel range K8:K10 does.
    s.write_formula(10, 10, "=MEDIAN(K8:K10)", f["multiple"], median(p[-1] for p in peers[1:]))
    section(s, 13, "LSEG implied value", 10)
    comp_bridge = [("Selected peer median EV / adjusted EBIT", "=J11", peer_median, f["multiple"]), ("LSEG 2025 adjusted operating profit", "='Historical'!D11", 3506, f["gbp"]), ("Implied enterprise value", "=B15*B16", peer_median*3506, f["gbp"]), ("Less operating net debt", "=-'Assumptions'!$B$10", -NET_DEBT, f["gbp"]), ("Implied equity value", "=SUM(B17:B18)", peer_median*3506-NET_DEBT, f["gbp"]), ("Voting shares used as proxy (m)", "='Assumptions'!$B$9", SHARE_COUNT, f["num"]), ("Implied value / share", "=B19/B20", comp_value, f["total_ps"])]
    for r, (lab, formula, value, form) in enumerate(comp_bridge, 14): s.write(r, 0, lab, f["text"]); s.write_formula(r, 1, formula, form, value)
    s.merge_range("A23:K26", "EV/adjusted EBIT uses a common metric label, but company adjustments and business mix still differ. Some peer shares, net-debt figures and P/E inputs remain approximate or unreconciled legacy inputs (Sources). The P/E median cache is corrected to 25.46x, but P/E is descriptive and does not drive value. The £123.96 result is an indicative screening cross-check, not a fully harmonised target price.", f["wrap"])
    s.set_column("A:A", 39); s.set_column("B:B", 18); s.set_column("C:K", 15)

    # Scenarios.
    s = ws["Scenarios"]
    title(s, "Base, Upside and Downside Cases")
    s.merge_range("A3:H3", "Case values recalculate the underlying DCF using case-specific growth, margins and terminal growth.", f["meta"])
    section(s, 4, "Case valuation")
    s.write_row("A6", ["Case", "Revenue CAGR", "2030 revenue", "2030 margin", "2030 EBITDA", "Terminal growth", "Enterprise value", "Value / share"], f["header"])
    for r, (name, args) in enumerate(cases.items(), 6):
        rows, ev, value = dcf_case(*args); cagr=(rows[-1][0]/8986)**(1/5)-1
        ar = r + 21  # Assumptions rows 27:29
        rev_terms, fcf_terms = [], []
        for year in range(1, 6):
            end_col = chr(66 + year - 1)
            growth_product = "*".join(f"(1+'Assumptions'!${chr(66+k)}${ar})" for k in range(year))
            rev = f"'Historical'!$D$7*{growth_product}"
            margin = f"'Historical'!$D$10+('Assumptions'!$G${ar}-'Historical'!$D$10)*{year}/5"
            capex_col = chr(67 + year)  # D:H
            fcf = f"(({rev})*({margin})*(1-'Assumptions'!$B$37)+({rev})*'Assumptions'!$B$34*'Assumptions'!$B$37-({rev})*'Assumptions'!${capex_col}$35-({rev})*'Assumptions'!$B$36)"
            rev_terms.append(rev); fcf_terms.append(fcf)
        pv = "+".join(f"({term})/(1+'Assumptions'!$B$23)^{year}" for year, term in enumerate(fcf_terms, 1))
        ev_formula = f"={pv}+({fcf_terms[-1]})*(1+'Assumptions'!$H${ar})/('Assumptions'!$B$23-'Assumptions'!$H${ar})/(1+'Assumptions'!$B$23)^5"
        s.write(r, 0, name, f["text"])
        full_growth_product = "*".join(f"(1+'Assumptions'!${chr(66+k)}${ar})" for k in range(5))
        s.write_formula(r, 1, f"=({full_growth_product})^(1/5)-1", f["pct"], cagr)
        s.write_formula(r, 2, f"={rev_terms[-1]}", f["gbp"], rows[-1][0])
        s.write_formula(r, 3, f"='Assumptions'!$G${ar}", f["pct"], args[1])
        s.write_formula(r, 4, f"=C{r+1}*D{r+1}", f["gbp"], rows[-1][2])
        s.write_formula(r, 5, f"='Assumptions'!$H${ar}", f["pct"], args[2])
        s.write_formula(r, 6, ev_formula, f["gbp"], ev)
        s.write_formula(r, 7, f"=(G{r+1}-'Assumptions'!$B$10)/'Assumptions'!$B$9", f["ps"], value)
    s.set_column("A:H", 18)

    # Sensitivities.
    s = ws["Sensitivities"]
    title(s, "DCF Sensitivities")
    s.merge_range("A3:H3", "Sensitivity outputs independently recalculate the valuation mechanics.", f["meta"])

    def operating_sensitivity_formula(cagr_ref, margin_ref):
        """Return a live Excel formula that rebuilds all five annual FCFs and terminal value."""
        fcf_terms = []
        for year in range(1, 6):
            capex_col = chr(67 + year)  # D:H on Assumptions
            revenue = f"'Historical'!$D$7*(1+{cagr_ref})^{year}"
            margin = f"'Historical'!$D$10+({margin_ref}-'Historical'!$D$10)*{year}/5"
            fcf_terms.append(
                f"(({revenue})*({margin})*(1-'Assumptions'!$B$37)"
                f"+({revenue})*'Assumptions'!$B$34*'Assumptions'!$B$37"
                f"-({revenue})*'Assumptions'!${capex_col}$35"
                f"-({revenue})*'Assumptions'!$B$36)"
            )
        forecast_pv = "+".join(
            f"({term})/(1+'Assumptions'!$B$23)^{year}"
            for year, term in enumerate(fcf_terms, 1)
        )
        terminal_pv = (
            f"({fcf_terms[-1]})*(1+'Assumptions'!$B$38)"
            f"/('Assumptions'!$B$23-'Assumptions'!$B$38)"
            f"/(1+'Assumptions'!$B$23)^5"
        )
        return f"=(({forecast_pv})+({terminal_pv})-'Assumptions'!$B$10)/'Assumptions'!$B$9"

    section(s, 4, "Value per share: WACC vs terminal growth")
    tgs=[.015,.02,.025,.03,.035]; waccs=[.065,.07,.075,.08,.085]
    s.write(5,0,"WACC / TGR",f["header"])
    for c,tg in enumerate(tgs,1): s.write(5,c,tg,f["pct"])
    for r,wacc in enumerate(waccs,6):
        s.write(r,0,wacc,f["pct"])
        for c,tg in enumerate(tgs,1):
            _,_,value=dcf_case(*cases["Base"],wacc=wacc)
            # Recompute the terminal growth dimension while retaining the same forecast.
            rows,_,_=dcf_case(cases["Base"][0],cases["Base"][1],tg,wacc)
            pv=sum(x[10] for x in rows); terminal=rows[-1][8]*(1+tg)/(wacc-tg)*rows[-1][9]; value=(pv+terminal-NET_DEBT)/SHARE_COUNT
            formula=f"=(SUMPRODUCT('DCF'!$C$18:$G$18,1/(1+$A{r+1})^'DCF'!$C$19:$G$19)+'DCF'!$G$18*(1+{chr(65+c)}$6)/($A{r+1}-{chr(65+c)}$6)/(1+$A{r+1})^5-'Assumptions'!$B$10)/'Assumptions'!$B$9"
            s.write_formula(r,c,formula,f["ps"],value)
    section(s, 13, "Value per share: revenue CAGR vs 2030 EBITDA margin")
    margins=[.50,.52,.54,.56,.58]; cagrs=[.04,.05,.06,.07,.08]
    s.write(14,0,"CAGR / Margin",f["header"])
    for c,m in enumerate(margins,1): s.write(14,c,m,f["pct"])
    for r,cagr in enumerate(cagrs,15):
        s.write(r,0,cagr,f["pct"])
        for c,m in enumerate(margins,1):
            _,_,value=dcf_case([cagr]*5,m,.025)
            formula=operating_sensitivity_formula(f"$A{r+1}", f"{chr(65+c)}$15")
            s.write_formula(r,c,formula,f["ps"],value)
    section(s,22,"Sensitivity calculation audit")
    s.write_row("A24",["Selected grid point","Table output","Formula consistency","Difference"],f["header"])
    s.write("A25","6.0% CAGR / 54.0% margin",f["text"])
    audit_value = dcf_case([.06]*5,.54,.025)[2]
    s.write_formula("B25","=D18",f["ps"],audit_value)
    s.write_formula("C25",operating_sensitivity_formula("$A$18","D$15"),f["ps"],audit_value)
    s.write_formula("D25","=B25-C25",f["ps"],0)
    s.set_column("A:A",22); s.set_column("B:F",18)

    # AI review.
    s = ws["AI Review"]
    title(s, "AI-Supported Document Review and Comparable Screening")
    s.merge_range("A3:H3", "AI supports extraction, summarisation, peer screening and risk identification; Python automates the workbook. Human judgement remains responsible for verification, assumptions, peer selection and the conclusion.", f["meta"])
    section(s, 4, "Document review log")
    s.write_row("A6", ["Topic", "AI-supported extraction", "Source evidence", "Model impact", "Confidence", "Human verified", "Risk of error", "Mitigation"], f["header"])
    ai_rows=[("FY2025 performance","Income £8.986bn; adj EBITDA £4.523bn","LSEG FY2025 annual report","Historical base year","High","Yes","Definition mismatch","Use excl. recoveries consistently"),("Guidance","2026 growth 6.5-7.5%; margin +80-100bps","LSEG FY2025 results","Base forecast anchors","High","Yes","Guidance not guaranteed","Scenario ranges"),("Cash conversion","Capex £919m; equity FCF £2.445bn","LSEG annual report","Capex sanity check","High","Yes","FCFE vs UFCF","DCF explicitly uses UFCF"),("Capital structure","Operating net debt £8.175bn","LSEG annual report","Equity bridge","High","Yes","Regulatory cash treatment","Use company definition"),("Peer relevance","ICE, CME, Nasdaq","FY2025 company results","Comparable screen","Medium","Yes","Business-mix differences","Median and limitation note"),("AI theme","AI-ready data partnerships","LSEG FY2025 results","Upside narrative only","Medium","Yes","Double counting","No separate base premium")]
    for r,row in enumerate(ai_rows,6): s.write_row(r,0,row,f["wrap"])
    # Source review by the assistant is not evidence that David has personally verified it.
    s.write("F6", "David's review", f["header"])
    for r in range(6,12): s.write(r,5,"Pending David",f["wrap"])
    for r in (6,8,9): s.write(r,2,"LSEG FY2025 preliminary results (S1)",f["wrap"])
    s.set_column("A:A",22); s.set_column("B:D",28); s.set_column("E:F",15); s.set_column("G:H",25)

    # Critical review and assumption evidence.
    s = ws["Critical Review"]
    title(s, "Critical Review: What Automation Can and Cannot Decide", 7)
    s.merge_range("A3:H3", "This appendix responds directly to model-risk feedback and prevents mechanically correct outputs from being presented as automatically reliable conclusions.", f["meta"])
    section(s, 4, "Corrections made", 7)
    s.write_row("A6", ["Issue", "Earlier treatment", "Correction", "Why it matters"], f["header"])
    corrections = [
        ("Information timing", "31 Dec 2025 date with later results; proxy also disclosed after 26 Feb", "FY2025 retrospective study, market reference 26 Feb and proxy publication 27 Feb disclosed", "Does not claim a strict point-in-time backtest or a 2025-only information set"),
        ("Mixed peer metrics", "EBITDA and operating income combined", "EV/adjusted EBIT used consistently", "Removes a material denominator mismatch"),
        ("Static controls", "Scenario, checks and recommendation could become stale", "Live Excel formulas now drive all three", "Workbook responds when assumptions change"),
        ("Arbitrary blend", "Fixed 50/50 weighting", "DCF is principal method; editable 70/30 weighting", "Recognises lower confidence in peer comparability"),
        ("WACC cache mismatch", "Python used 7.8% while Excel derived about 7.695%", "One WACC calculation now drives Python caches and Excel formulas", "Saved values no longer change solely because Excel recalculates"),
        ("Operating sensitivity defect", "Scenario terminal FCF was combined with base forecast FCF", "Every grid cell now rebuilds all five scenario FCFs and terminal value", "Sensitivity outputs now reflect both changed operating drivers"),
        ("Capex timing", "Capex reached 8.0% in 2027", "Annual path is 9.5%, 9.0%, 8.5%, 8.0%, 8.0%", "Matches the stated objective of reaching about 8% by 2029"),
        ("Share definition", "Voting shares were labelled diluted shares", "Voting-share proxy is labelled consistently and dilution is disclosed", "Avoids overstating the precision of per-share valuation"),
    ]
    for r,row in enumerate(corrections,6): s.write_row(r,0,row,f["wrap"])
    section(s, 15, "Assumption evidence and limitations", 7)
    s.write_row("A17", ["Assumption", "Base", "Status", "Support", "Challenge / sensitivity"], f["header"])
    assumptions = [
        ("Beta", .90, "Analyst estimate", "Broad market-risk assumption; not a reported company fact", "Tested indirectly through WACC sensitivity"),
        ("Risk-free rate", .045, "Market assumption", "Rounded long-dated sterling rate near valuation date", "WACC sensitivity spans 6.5%-8.5%"),
        ("Equity-risk premium", .045, "Market assumption", "Rounded mature-market ERP", "Change ERP or use WACC sensitivity"),
        ("Pre-tax cost of debt", .045, "Analyst estimate", "Approximation for investment-grade LSEG", "Editable input"),
        ("NWC investment / revenue", .015, "Normalised estimate", "FY2025 working-capital outflow was unusually high", "Editable and explicitly labelled"),
        ("Capex intensity", .095, "Company-guided near term", "2026 guidance about 9.5%; 2029 ambition about 8%", "Downward path beyond 2026 is an estimate"),
        ("2030 EBITDA margin", .54, "Analyst estimate, above guidance bridge", "50.3% FY2025 + 0.9pp 2026 midpoint + 1.5pp cumulative 2027-29 = c.52.7% in 2029 (S2, guidance sections)", "54% needs c.1.3pp more in 2030. Assumes further efficiency; not company guidance. Grid tests 50%-58%."),
        ("Terminal growth", .025, "Analyst estimate", "Long-run nominal growth assumption", "1.5%-3.5% sensitivity range"),
    ]
    for r,(lab,val,status,support,challenge) in enumerate(assumptions,17):
        s.write(r,0,lab,f["text"]); s.write(r,1,val,f["pct"] if lab!="Beta" else f["input_beta"]); s.write_row(r,2,[status,support,challenge],f["wrap"])
    section(s, 27, "Human judgement conclusion", 7)
    s.merge_range("A29:H33", "The calculations can be reproduced automatically, but the conclusion cannot be delegated to automation. The DCF is treated as the principal method because it links value to explicit operating assumptions. The comparable result is retained only as a reasonableness check after metric normalisation. David's conclusion is therefore POSITIVE / WATCH rather than an unconditional buy recommendation: the apparent upside must compensate for execution, leverage, dilution, peer-comparability and terminal-value risk.", f["wrap"])
    section(s, 34, "Final review: 7 September 2026", 7)
    s.merge_range("A36:H39", "The selected-case check now matches the case name to Scenarios rather than always testing Base. The peer P/E cache now derives from the three peer inputs (25.46x). The programme price proxy is confirmed, with its 27-Feb publication date disclosed. Sources maps material inputs to sections and records reclassifications and unresolved provenance. These changes do not turn approximate inputs into verified facts.", f["wrap"])
    s.merge_range("A41:H44", "Margin interpretation: management's 2026 and 2027-29 guidance is on a constant-currency/underlying basis. Combining it with the reported FY2025 margin is only an illustrative bridge. The existing straight-line path adds about 73bp a year (2026 about 51.1%), so it is not an exact reproduction of 80-100bp 2026 guidance. A 54% endpoint assumes sustained efficiency after 2029; the 52.7% alternative should also be discussed. No separate AI revenue or margin premium is added.", f["wrap"])
    for r in range(6,14): s.set_row(r, 50)
    for r in range(17,25): s.set_row(r, 38)
    s.set_row(23, 66)
    s.set_column("A:A",24); s.set_column("B:B",20); s.set_column("C:C",22); s.set_column("D:E",46); s.set_column("F:H",14)

    # Checks.
    s=ws["Checks"]
    title(s,"Model Checks",6)
    s.merge_range("A3:G3","",f["pass"])
    s.write_formula("A3", '=\"MODEL STATUS: \"&IF(COUNTIF(F7:F20,\"FAIL\")=0,\"PASS\",\"FAIL\")', f["pass"], "MODEL STATUS: PASS")
    section(s,4,"Audit checks",6)
    s.write_row("A6",["Check","Actual","Expected","Difference","Tolerance","Status","Fix hint"],f["header"])
    checks = [
        ("DCF FCF tie", "='DCF'!G18", "=SUM('DCF'!G14:G17)", 1, "=IF(ABS(D7)<=E7,\"OK\",\"FAIL\")", "Review DCF rows 14-18", base_rows[-1][8], base_rows[-1][8]),
        ("EV bridge", "='DCF'!B34", "=SUM('DCF'!B32:B33)", 1, "=IF(ABS(D8)<=E8,\"OK\",\"FAIL\")", "Review DCF rows 30-34", base_ev-NET_DEBT, base_ev-NET_DEBT),
        ("WACC formula consistency", "='Assumptions'!B23", "='Assumptions'!B20*'Assumptions'!B21+'Assumptions'!B19*'Assumptions'!B22", .000001, "=IF(ABS(D9)<=E9,\"OK\",\"FAIL\")", "Review Assumptions WACC build", WACC, WACC),
        ("WACC > terminal growth", "='Assumptions'!B23", "='Assumptions'!B38", 0, "=IF(D10>E10,\"OK\",\"FAIL\")", "Adjust Assumptions B23 / B38", WACC, .025),
        ("DCF / selected scenario tie", "='DCF'!B36", "=INDEX('Scenarios'!$H$7:$H$9,MATCH('Assumptions'!$B$5,'Scenarios'!$A$7:$A$9,0))", .01, "=IF(ABS(D11)<=E11,\"OK\",\"FAIL\")", "Match Assumptions B5 to Scenarios A7:A9", base_value, base_value),
        ("Shares positive", "='Assumptions'!B9", "=0", 0, "=IF(D12>E12,\"OK\",\"FAIL\")", "Review Assumptions B9", SHARE_COUNT, 0),
        ("Net debt non-negative", "='Assumptions'!B10", "=0", 0, "=IF(D13>=E13,\"OK\",\"FAIL\")", "Review Assumptions B10", NET_DEBT, 0),
        ("Selected case valid", "=COUNTIF('Assumptions'!A27:A29,'Assumptions'!B5)", "=1", 0, "=IF(B14=C14,\"OK\",\"FAIL\")", "Select Base/Upside/Downside", 1, 1),
        ("Comps median positive", "='Comps'!J11", "=0", 0, "=IF(D15>E15,\"OK\",\"FAIL\")", "Review Comps inputs", peer_median, 0),
        ("DCF value positive", "='DCF'!B36", "=0", 0, "=IF(D16>E16,\"OK\",\"FAIL\")", "Review DCF and WACC", base_value, 0),
        ("Method weights total 100%", "=SUM('Assumptions'!B40:B41)", "=1", .0001, "=IF(ABS(D17)<=E17,\"OK\",\"FAIL\")", "Review Assumptions B40:B41", 1, 1),
        ("Operating sensitivity recalculation", "='Sensitivities'!B25", "='Sensitivities'!C25", .01, "=IF(ABS(D18)<=E18,\"OK\",\"FAIL\")", "Review Sensitivities calculation audit", dcf_case([.06]*5,.54,.025)[2], dcf_case([.06]*5,.54,.025)[2]),
        ("Capex reaches 8% in 2029", "=IF(AND('Assumptions'!D35=9.5%,'Assumptions'!E35=9.0%,'Assumptions'!F35=8.5%,'Assumptions'!G35=8.0%,'Assumptions'!H35=8.0%),1,0)", "=1", 0, "=IF(B19=C19,\"OK\",\"FAIL\")", "Review Assumptions D35:H35", 1, 1),
        ("Share-definition consistency", '=--(\'Assumptions\'!A9="Voting shares used as proxy (m)")', "=1", 0, "=IF(B20=C20,\"OK\",\"FAIL\")", "Use voting-share proxy label consistently", 1, 1),
    ]
    for r,(label,actual_f,expected_f,tol,status_f,note,actual_v,expected_v) in enumerate(checks,6):
        excel_row=r+1
        s.write(r,0,label,f["text"]); s.write_formula(r,1,actual_f,f["formula"],actual_v); s.write_formula(r,2,expected_f,f["formula"],expected_v)
        s.write_formula(r,3,f"=B{excel_row}-C{excel_row}",f["formula"],actual_v-expected_v); s.write(r,4,tol,f["formula"]); s.write_formula(r,5,status_f,f["formula"],"OK"); s.write(r,6,note,f["text"])
    s.conditional_format("F7:F20", {"type":"text", "criteria":"containing", "value":"OK", "format":f["pass"]})
    s.conditional_format("F7:F20", {"type":"text", "criteria":"containing", "value":"FAIL", "format":f["fail"]})
    s.set_column("A:A",30); s.set_column("B:F",16); s.set_column("G:G",35)
    s.merge_range("A23:G25", "PASS covers the listed arithmetic and label checks only. Source provenance, restatement scope, peer estimates and forecast credibility require separate review in Sources. The sensitivity consistency line repeats the valuation formula; independent recalculation testing is documented in the review guide.", f["wrap"])

    # Sources.
    s=ws["Sources"]
    title(s,"Sources and Audit Trail")
    s.merge_range("A3:H3","FY2025 study. Market reference 26 Feb 2026; programme VWAP/share disclosure 27 Feb. Material-input locators and restatement status below.",f["meta"])
    section(s,4,"Source log")
    s.write_row("A6",["ID","Item","Value / use","Units","Period / as-of","Source","URL","Notes"],f["header"])
    sources=[
        ("S1","LSEG FY2025 preliminary results PDF","Historical financials, cash flow and debt","GBP m","Released 2026-02-26","LSEG","https://www.lseg.com/content/dam/lseg/en_us/documents/investor-relations/financial-results/preliminary-results/rns/lseg-2025-preliminary-results-rns-26feb2026.pdf","Use release-date evidence; do not assume the annual report was already public"),
        ("S2","LSEG FY2025 preliminary results","Adjusted KPIs and initial 2026 guidance","GBP m","Released 2026-02-26","LSEG","https://www.lseg.com/en/media-centre/press-releases/2026/london-stock-exchange-group-plc-preliminary-results-for-the-year-ended-31-december-2025","Sets the information/valuation date"),
        ("S3","LSEG buyback RNS published 27 Feb","26-Feb purchases VWAP 8,126.60p; prospective voting rights 505.332519m","GBX / shares","Trade 2026-02-26; publication 2026-02-27","LSEG RNS","https://www.investegate.co.uk/announcement/rns/london-stock-exchange-group--lseg/transaction-in-own-shares/9449825","Ordinary Shares table; following-cancellation paragraph. Verified retrospective proxy, not close or strict same-day public information."),
        ("S4","ICE FY2025 results","Revenue, adjusted operating income, debt, cash and shares","USD m","FY2025","ICE","https://ir.theice.com/press/news-details/2026/Intercontinental-Exchange-Reports-Strong-Full-Year-2025-Results/","Primary financial source"),
        ("S5","CME FY2025 results","Revenue and adjusted operating income","USD m","FY2025","CME","https://investor.cmegroup.com/news-releases/news-release-details/cme-group-inc-reports-fourth-consecutive-year-record-annual","Adjusted operating income $4,524.8m"),
        ("S6","Nasdaq FY2025 results","Net revenue and adjusted operating income","USD m","FY2025","Nasdaq","https://ir.nasdaq.com/news-releases/news-release-details/nasdaq-reports-fourth-quarter-and-full-year-2025-results-annual","Primary financial source"),
        ("S7","ICE historical price","Close $162.89","USD/share","2026-02-26","Yahoo Finance / ChartExchange","https://chartexchange.com/symbol/nyse-ice/historical/","Date-aligned secondary market source"),
        ("S8","CME historical price","Close $316.45","USD/share","2026-02-26","ChartExchange","https://chartexchange.com/symbol/nasdaq-cme/historical/","Date-aligned secondary market source"),
        ("S9","Nasdaq historical price","Close $88.59","USD/share","2026-02-26","ChartExchange","https://chartexchange.com/symbol/nasdaq-ndaq/historical/","Date-aligned secondary market source"),
        ("S10","GBP/USD historical FX","1.3491 USD per GBP","FX","2026-02-26","ExchangeRates.org.uk","https://www.exchangerates.org.uk/historical/find-exchange-rate-history-for-26_02_2026","For date alignment; multiples use local-currency numerator and denominator"),
        ("S11","Valuation assumptions","Beta, WACC inputs, NWC, capex path, margin and TGR","Various","2026-02-26 valuation","Model author","N/A","Explicit estimates; see Critical Review")
    ]
    for r,row in enumerate(sources,6): s.write_row(r,0,row,f["wrap"])
    s.write_row("A18", ["S12", "FY2025 voting rights", "510.408075m excluding treasury", "shares m", "2025-12-31", "LSEG RNS", "https://www.investegate.co.uk/announcement/rns/london-stock-exchange-group--lseg/total-voting-rights/9330335", "Total Voting Rights: share-capital and voting-rights paragraphs"], f["wrap"])
    section(s,19,"Material inputs: source locator and restatement / estimation status")
    s.write_row("A21",["Input ID","Workbook location","Input / definition","Source ID","Page or section","Period / publication","Restatement / verification status","Treatment / limitation"],f["header"])
    input_audit = [
        ("I01","Historical C7:D11","2024/25 income excluding recoveries; adjusted EBITDA and operating profit","S1 / S2","S1 p.1 summary tables; S2 Reported and Adjusted tables","FY2024/25; published 26 Feb 2026","Matches displayed comparative figures; no restated label on these group totals","Adjusted is not synonymous with restated. Use the release's comparative basis."),
        ("I02","Historical B7:B12","2023 income 8,061; adj EBITDA 3,801; adj EBIT 2,909","Legacy input","Original FY2023 table/page not preserved","FY2023","Original source and restatement basis not independently resolved in this review","Context only; does not drive FY2025-based DCF. Do not cite as newly verified."),
        ("I03","Historical C13:D16","Reported EBITDA; working capital; capex; equity FCF","S1","p.20 Cash Flow, table and footnote 1; p.50 APM reconciliation","FY2024/25; 26 Feb 2026","2024 £12m reclassified from working capital to non-cash items","Uses reclassified 2024 working-capital outflow £50m. 2025 outflow £419m; equity FCF is not UFCF."),
        ("I04","Historical C17:D17; Assumptions B10","Operating net debt 7,178 / 8,175","S1","p.21 Net Debt / Leverage / Ratings; p.50 Net debt and operating net debt","31 Dec 2024/25","Matches current release; no restated label on debt table","Excludes lease liabilities and adds regulatory/operational cash. No 2026 debt roll-forward or full minority-interest bridge."),
        ("I05","Historical D18","510.408075m voting rights excluding treasury","S12","Total Voting Rights: share-capital and voting-rights paragraphs","31 Dec 2025","Matched to dated company announcement","Period-end voting rights, not diluted weighted-average shares."),
        ("I06","Assumptions B8; Comps C7","£81.266 programme purchase VWAP","S3","Ordinary Shares table: date of purchase and volume weighted average price","Purchases 26 Feb; public 27 Feb 2026","Verified 8,126.60p / 100; not a financial-statement restatement item","Retained proxy covers 34,592 bought shares. Not an official close or all-market VWAP; retrospective use disclosed."),
        ("I07","Assumptions B9; Comps D7","505.332519m voting-share proxy","S3","Paragraph beginning Following the cancellation","Published 27 Feb 2026","Verified announced post-cancellation denominator; not proof of completion time","Excludes treasury. No allowance for potential dilution; not a same-day public-information backtest."),
        ("I08","Assumptions B11:B17","Rf 4.5%, ERP 4.5%, beta 0.9, debt cost 4.5%, tax 24.5%","S11 / S2","Author WACC assumptions; S2 2026 guidance for tax only","Valuation assumptions","Market-risk inputs are rounded estimates, not verified historical observations","WACC uses operating net debt as a simplified weight. Debt tax rate and operating tax driver are separately editable."),
        ("I09","Assumptions B27:F29","Annual revenue growth by scenario","S2 / S11","2026 guidance; Medium-term guidance 2027-2029","2026-30 forecasts","2026 Base 7% is guidance midpoint; later yearly choices and other cases are estimates","Guidance is organic constant currency; model applies it to reported GBP income without a separate FX/acquisition bridge."),
        ("I10","Assumptions G27:G29; DCF C9:G9","2030 EBITDA margins 51.5%, 54%, 56%","S2 / S11","Financial highlights; 2026 guidance; Medium-term guidance 2027-2029","2025 anchor to 2030 forecast","54% is an analyst assumption. Approximate guidance bridge reaches 52.7% in 2029","Requires c.1.3pp additional 2030 expansion; flat model ramp gives c.73bp annually, not exact 2026 guidance."),
        ("I11","Assumptions B34","D&A / income 11.3%","S1 / S11","p.1 adjusted EBITDA less adjusted operating profit","FY2025 calibration","Derived and rounded: (4,523 - 3,506) / 8,986 = 11.3176%","Adjusted depreciation/amortisation proxy, not total statutory D&A; held flat in forecasts."),
        ("I12","Assumptions D35:H35","Capex / income 9.5%, 9%, 8.5%, 8%, 8%","S2 / S11","2026 guidance; Medium-term guidance 2027-2029","2026-30 forecast","2026 and 2029 endpoints guided; 2027/28 interpolation and 2030 flat rate estimated","Not annual company guidance; based on capex excluding sales commissions and SwapClear intangible payment (S1 p.50)."),
        ("I13","Assumptions B36:B38; H27:H29","NWC 1.5%; tax 24.5%; terminal growth 2%-3%","S1 / S2 / S11","S1 p.20 Cash Flow explanation; S2 2026 guidance; author TGR","2026-30 and terminal assumptions","NWC and TGR estimated; tax is 24%-25% guidance midpoint","NWC is annual investment as a share of revenue, not a balance-sheet NWC ratio. Cash tax timing simplified."),
        ("I14","Assumptions B40:B41","70% DCF / 30% comps","S11","Author method weighting","Valuation assumption","Judgement; not statistically estimated or source-reported","Principal conclusion is DCF-led. Blend inherits peer-data uncertainty."),
        ("I15","Comps G8:H8","ICE 9,931 revenue net of transaction costs; 5,959 adjusted operating income","S4","Adjusted Operating Income, Operating Margin and Operating Expense Reconciliation: FY2025 consolidated","FY2025; released 5 Feb 2026","Matched release; non-GAAP adjusted, not labelled restated in selected table","Company-specific exclusions remain; common adjusted EBIT label does not fully harmonise accounting."),
        ("I16","Comps D8:F8","ICE shares 575m; net debt 18,763m","S4","Adjusted Net Income / EPS: diluted weighted-average shares; Other Matters: cash and debt","FY2025 / 31 Dec 2025","575m is diluted weighted average; net debt uses rounded debt 19.6bn less cash 837m","Market cap therefore uses an annual share proxy, not verified 26-Feb spot shares. Debt precision limited by rounding."),
        ("I17","Comps G9:H9","CME revenue 6,520.6 and adjusted operating income 4,524.8","S5","Reconciliation of Adjusted Operating Income: FY2025 column","FY2025; released 4 Feb 2026","Matched selected non-GAAP table; no restated label in that table","Local USD numerator and denominator; company-specific adjustments."),
        ("I18","Comps D9:F9","CME shares 361m; net cash 1,200m","Legacy estimates / S5","Exact source line for both legacy capital inputs not established","Intended FY2025","Unreconciled approximations retained; not verified source facts","Comps estimate is provisional. No claim these are current point-in-time shares or exact unrestricted net cash."),
        ("I19","Comps G10:H10","Nasdaq net revenue 5,249 and non-GAAP operating income 2,918","S6","2025 highlights table; Reconciliation of U.S. GAAP to Non-GAAP Operating Income","FY2025; released 29 Jan 2026","Matched FY2025; 2024 non-GAAP revenue comparatives differ from GAAP in reconciliation","Use 2025 2,918 operating income, not the identically numbered GAAP operating-expense line."),
        ("I20","Comps D10:F10","Nasdaq shares 575m; net debt 8,000m","Legacy estimates / S6","Exact 575m/8,000m provenance not established","Intended FY2025","Approximate/unreconciled; release reports FY2025 diluted weighted-average shares 578.6m","Retained screening assumptions, not verified spot capital structure."),
        ("I21","Comps C8:C10","ICE 162.89; CME 316.45; Nasdaq 88.59 (USD)","S7-S9","Historical price tables: 26-Feb-2026 row","26 Feb 2026","Previously collected prices; archived rows not revalidated in this final review","Historical inputs frozen; not a live feed. No claim of renewed independent price verification."),
        ("I22","Comps K7:K11","P/E inputs 19.32 / 23.30 / 28.25 / 25.46; peer median 25.46x","Legacy estimates","K8:K10 sorted: 23.30, 25.46, 28.25","Legacy FY2025 screening inputs","Median arithmetic and cache verified; original P/E earnings basis not fully evidenced","P/E is descriptive only and excluded from valuation. Cache correction does not validate each input's basis."),
        ("I23","DCF B13:B21","Illustrative historical tax and UFCF bridge","S11","Historical adjusted EBIT and simplified tax; not reported cash-tax series","FY2025 illustration","Model-derived/rounded, not statutory historical UFCF","Forecast valuation uses C:G only. Do not present illustrative tax 859m as cash tax paid (S1 p.20 is 396m)."),
    ]
    for r,row in enumerate(input_audit,21):
        s.write_row(r,0,row,f["wrap"])
        s.set_row(r,86)
    for r in range(6,18): s.set_row(r,75)
    s.set_column("A:A",8); s.set_column("B:B",27); s.set_column("C:C",34); s.set_column("D:F",17); s.set_column("G:G",70); s.set_column("H:H",34)
    s.set_column("D:D",18); s.set_column("E:E",40); s.set_column("F:F",26)

    wb.close()
    print(OUT)


if __name__ == "__main__":
    build()
