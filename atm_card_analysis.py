"""
ATM, Acceptance Infrastructure & Card Statistics - September 2026 (RBI data)
=============================================================================
End-to-end analysis script.

HOW TO RUN (VS Code / terminal)
-------------------------------
1. pip install pandas numpy matplotlib openpyxl
2. Keep the Excel file inside the  data/  folder (or next to this script).
3. python atm_card_analysis.py
   (or: python atm_card_analysis.py "path/to/your_file.XLSX")

OUTPUT
------
A folder called  `output/`  with:
  - clean_bank_data.csv        -> cleaned, one row per bank, with all KPIs
  - analysis_summary.xlsx      -> every summary table on its own sheet
  - charts/*.png               -> 10 presentation-ready charts
  - key_insights.txt           -> auto-generated headline findings
"""

import sys
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# --------------------------------------------------------------------------
# 0. SETTINGS
# --------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent


def find_default_file() -> Path:
    """Look for the RBI Excel file in ./data first, then next to the script."""
    for folder in (SCRIPT_DIR / "data", SCRIPT_DIR):
        for f in sorted(folder.glob("*")):
            if f.suffix.lower() == ".xlsx" and not f.name.startswith(("~", "analysis_summary")):
                return f
    return SCRIPT_DIR / "data" / "RBI_ATM_file.xlsx"   # placeholder -> triggers friendly error


FILE_PATH = Path(sys.argv[1]) if len(sys.argv) > 1 else find_default_file()

OUT_DIR = SCRIPT_DIR / "output"
CHART_DIR = OUT_DIR / "charts"
CHART_DIR.mkdir(parents=True, exist_ok=True)

# Values in the file are in Rs'000.  1 crore = 10,000,000 = 10,000 x Rs'000
THOUSAND_TO_CRORE = 1 / 10_000

plt.rcParams.update({
    "figure.dpi": 110,
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.titlesize": 13,
    "axes.titleweight": "bold",
    "axes.labelsize": 10,
})
CAT_COLORS = {
    "Public Sector Banks": "#1f77b4",
    "Private Sector Banks": "#ff7f0e",
    "Foreign Banks": "#2ca02c",
    "Payment Banks": "#9467bd",
    "Small Finance Banks": "#d62728",
}


# --------------------------------------------------------------------------
# 1. LOAD + CLEAN
# --------------------------------------------------------------------------
# The sheet has a 7-row multi-line header, so we read it raw and name the
# columns ourselves (the numbers 1-26 in the sheet map to these names).
COLUMN_NAMES = [
    "atm_onsite", "atm_offsite", "pos_terminals", "micro_atms",
    "bharat_qr", "upi_qr", "credit_cards", "debit_cards",
    "cc_pos_vol", "cc_pos_val", "cc_online_vol", "cc_online_val",
    "cc_other_vol", "cc_other_val", "cc_atm_vol", "cc_atm_val",
    "dc_pos_vol", "dc_pos_val", "dc_online_vol", "dc_online_val",
    "dc_other_vol", "dc_other_val", "dc_atm_vol", "dc_atm_val",
    "dc_poscash_vol", "dc_poscash_val",
]


def load_and_clean(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None)

    # Find where bank data starts: first row whose 2nd column is a number (Sr. No.)
    records = []
    current_category = None
    for _, row in raw.iterrows():
        sr, name = row[1], row[2]

        # Section heading row (e.g. "Public Sector Banks"): text in col B, nothing in col C
        if isinstance(sr, str) and pd.isna(name):
            text = sr.strip()
            if text in CAT_COLORS:
                current_category = text
            continue

        # Real bank row: Sr. No. is numeric and bank name is text
        if pd.notna(sr) and isinstance(name, str) and str(sr).strip().isdigit():
            values = row.iloc[3:29].tolist()
            records.append([int(sr), name.strip(), current_category] + values)

    df = pd.DataFrame(records, columns=["sr_no", "bank", "category"] + COLUMN_NAMES)
    df[COLUMN_NAMES] = df[COLUMN_NAMES].apply(pd.to_numeric, errors="coerce").fillna(0)
    return df


# --------------------------------------------------------------------------
# 2. FEATURE ENGINEERING (the KPIs that answer business questions)
# --------------------------------------------------------------------------
def safe_div(a, b):
    """Divide, returning NaN instead of inf when the denominator is 0."""
    return np.where(b == 0, np.nan, a / b)


def add_kpis(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()

    # --- Infrastructure ---
    d["atm_total"] = d["atm_onsite"] + d["atm_offsite"]
    d["offsite_share_pct"] = safe_div(d["atm_offsite"], d["atm_total"]) * 100
    d["digital_qr_total"] = d["bharat_qr"] + d["upi_qr"]

    # --- Card base ---
    d["total_cards"] = d["credit_cards"] + d["debit_cards"]

    # --- Credit card spending (volume = no. of txns, value = Rs'000) ---
    d["cc_spend_vol"] = d["cc_pos_vol"] + d["cc_online_vol"] + d["cc_other_vol"]
    d["cc_spend_val"] = d["cc_pos_val"] + d["cc_online_val"] + d["cc_other_val"]
    d["cc_spend_cr"] = d["cc_spend_val"] * THOUSAND_TO_CRORE
    d["cc_online_share_pct"] = safe_div(d["cc_online_val"], d["cc_spend_val"]) * 100

    # --- Debit card spending ---
    d["dc_spend_vol"] = d["dc_pos_vol"] + d["dc_online_vol"] + d["dc_other_vol"]
    d["dc_spend_val"] = d["dc_pos_val"] + d["dc_online_val"] + d["dc_other_val"]
    d["dc_spend_cr"] = d["dc_spend_val"] * THOUSAND_TO_CRORE
    d["dc_online_share_pct"] = safe_div(d["dc_online_val"], d["dc_spend_val"]) * 100

    # --- Cash withdrawals ---
    d["atm_cash_val_cr"] = (d["dc_atm_val"] + d["cc_atm_val"]) * THOUSAND_TO_CRORE
    d["atm_cash_vol"] = d["dc_atm_vol"] + d["cc_atm_vol"]

    # --- Productivity / intensity ratios ---
    d["cards_per_atm"] = safe_div(d["total_cards"], d["atm_total"])
    d["dc_atm_txn_per_atm"] = safe_div(d["dc_atm_vol"], d["atm_total"])
    d["dc_atm_cash_per_atm_lakh"] = safe_div(d["dc_atm_val"] * 1000, d["atm_total"]) / 1e5
    d["pos_txn_per_terminal"] = safe_div(d["dc_pos_vol"] + d["cc_pos_vol"], d["pos_terminals"])

    # --- Per-card usage ---
    d["cc_spend_per_card_rs"] = safe_div(d["cc_spend_val"] * 1000, d["credit_cards"])
    d["dc_spend_per_card_rs"] = safe_div(d["dc_spend_val"] * 1000, d["debit_cards"])
    d["cc_txn_per_card"] = safe_div(d["cc_spend_vol"], d["credit_cards"])
    d["dc_txn_per_card"] = safe_div(d["dc_spend_vol"], d["debit_cards"])

    # --- Average ticket size (Rs per transaction) ---
    d["cc_avg_ticket_rs"] = safe_div(d["cc_spend_val"] * 1000, d["cc_spend_vol"])
    d["dc_avg_ticket_rs"] = safe_div(d["dc_spend_val"] * 1000, d["dc_spend_vol"])
    d["dc_atm_avg_withdrawal_rs"] = safe_div(d["dc_atm_val"] * 1000, d["dc_atm_vol"])

    # --- Cash dependence: of all money debit cards moved, how much came out as ATM cash? ---
    d["dc_cash_dependence_pct"] = safe_div(
        d["dc_atm_val"], d["dc_atm_val"] + d["dc_spend_val"]
    ) * 100

    # --- Is the bank a credit-card player or a debit-card player? ---
    d["credit_card_mix_pct"] = safe_div(d["credit_cards"], d["total_cards"]) * 100

    # --- Share of the entire system (market share) ---
    for col, new in [
        ("atm_total", "share_atm_pct"),
        ("total_cards", "share_cards_pct"),
        ("credit_cards", "share_cc_pct"),
        ("debit_cards", "share_dc_pct"),
        ("cc_spend_val", "share_cc_spend_pct"),
        ("dc_spend_val", "share_dc_spend_pct"),
        ("pos_terminals", "share_pos_pct"),
    ]:
        d[new] = d[col] / d[col].sum() * 100

    return d


# --------------------------------------------------------------------------
# 3. SUMMARY TABLES
# --------------------------------------------------------------------------
def category_summary(d: pd.DataFrame) -> pd.DataFrame:
    g = d.groupby("category").agg(
        banks=("bank", "count"),
        atms=("atm_total", "sum"),
        atms_offsite=("atm_offsite", "sum"),
        pos=("pos_terminals", "sum"),
        upi_qr=("upi_qr", "sum"),
        credit_cards=("credit_cards", "sum"),
        debit_cards=("debit_cards", "sum"),
        cc_spend_cr=("cc_spend_cr", "sum"),
        dc_spend_cr=("dc_spend_cr", "sum"),
        atm_cash_cr=("atm_cash_val_cr", "sum"),
        cc_vol=("cc_spend_vol", "sum"),
        dc_vol=("dc_spend_vol", "sum"),
        cc_online_val=("cc_online_val", "sum"),
        cc_val=("cc_spend_val", "sum"),
    )
    g["offsite_share_pct"] = g["atms_offsite"] / g["atms"] * 100
    g["cc_spend_per_card_rs"] = g["cc_spend_cr"] * 1e7 / g["credit_cards"]
    g["dc_spend_per_card_rs"] = g["dc_spend_cr"] * 1e7 / g["debit_cards"]
    g["cc_avg_ticket_rs"] = g["cc_spend_cr"] * 1e7 / g["cc_vol"]
    g["dc_avg_ticket_rs"] = g["dc_spend_cr"] * 1e7 / g["dc_vol"]
    g["cc_online_share_pct"] = g["cc_online_val"] / g["cc_val"] * 100
    g["cards_per_atm"] = (g["credit_cards"] + g["debit_cards"]) / g["atms"].replace(0, np.nan)
    g["share_cc_spend_pct"] = g["cc_spend_cr"] / g["cc_spend_cr"].sum() * 100
    g["share_dc_spend_pct"] = g["dc_spend_cr"] / g["dc_spend_cr"].sum() * 100
    g["share_atms_pct"] = g["atms"] / g["atms"].sum() * 100
    return g.drop(columns=["cc_vol", "dc_vol", "cc_online_val", "cc_val"]).reset_index()


def concentration(d: pd.DataFrame) -> pd.DataFrame:
    """How concentrated is the market? Top-N share + HHI for key metrics."""
    rows = []
    for label, col in [
        ("ATMs", "atm_total"),
        ("Credit cards", "credit_cards"),
        ("Debit cards", "debit_cards"),
        ("Credit card spend", "cc_spend_val"),
        ("Debit card spend", "dc_spend_val"),
        ("PoS terminals", "pos_terminals"),
        ("UPI QR codes", "upi_qr"),
    ]:
        share = d[col] / d[col].sum()
        s = share.sort_values(ascending=False)
        rows.append({
            "metric": label,
            "top1_bank": d.loc[d[col].idxmax(), "bank"],
            "top1_share_pct": s.iloc[0] * 100,
            "top3_share_pct": s.iloc[:3].sum() * 100,
            "top5_share_pct": s.iloc[:5].sum() * 100,
            "top10_share_pct": s.iloc[:10].sum() * 100,
            "HHI": (share ** 2).sum() * 10_000,   # <1500 competitive, >2500 concentrated
        })
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------
# 4. CHARTS
# --------------------------------------------------------------------------
def save(fig, name):
    fig.tight_layout()
    fig.savefig(CHART_DIR / name, bbox_inches="tight")
    plt.close(fig)


def short(name):
    n = (name.replace(" BANKING CORPORATION", "").replace(" LTD.", "").replace(" LTD", "")
             .replace(" SMALL FINANCE BANK", " SFB").replace(" PAYMENTS BANK", " PB")
             .replace(" BANK", ""))
    return n.title()


def make_charts(d: pd.DataFrame, cat: pd.DataFrame):
    colors = d["category"].map(CAT_COLORS)

    # 1. Top 10 banks by ATMs
    t = d.nlargest(10, "atm_total").iloc[::-1]
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.barh(t["bank"].map(short), t["atm_total"], color=t["category"].map(CAT_COLORS))
    ax.set_title("Top 10 banks by ATM/CRM network")
    ax.set_xlabel("ATMs & CRMs (on-site + off-site)")
    for i, v in enumerate(t["atm_total"]):
        ax.text(v, i, f" {v:,.0f}", va="center", fontsize=8)
    save(fig, "01_top10_atms.png")

    # 2. Category share: ATMs vs cards vs spend
    fig, ax = plt.subplots(figsize=(9, 5))
    share = pd.DataFrame({
        "ATMs": cat.set_index("category")["share_atms_pct"],
        "Credit card spend": cat.set_index("category")["share_cc_spend_pct"],
        "Debit card spend": cat.set_index("category")["share_dc_spend_pct"],
    })
    share.T.plot(kind="bar", stacked=True, ax=ax,
                 color=[CAT_COLORS[c] for c in share.index], width=0.6)
    ax.set_title("Who holds the market? Share by bank category")
    ax.set_ylabel("% share")
    ax.set_xticklabels(share.columns, rotation=0)
    ax.legend(bbox_to_anchor=(1.01, 1), loc="upper left", frameon=False)
    save(fig, "02_category_share.png")

    # 3. Credit card base vs spend (top 12 by credit cards)
    t = d.nlargest(12, "credit_cards").iloc[::-1]
    fig, ax = plt.subplots(1, 2, figsize=(12, 5), sharey=True)
    ax[0].barh(t["bank"].map(short), t["credit_cards"] / 1e6, color="#4c72b0")
    ax[0].set_title("Credit cards outstanding (millions)")
    ax[1].barh(t["bank"].map(short), t["cc_spend_cr"], color="#dd8452")
    ax[1].set_title("Credit card spend in month (Rs crore)")
    save(fig, "03_credit_cards_base_vs_spend.png")

    # 4. Spend per credit card (min 100k cards so tiny bases don't distort)
    t = d[d["credit_cards"] >= 100_000].nlargest(12, "cc_spend_per_card_rs").iloc[::-1]
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.barh(t["bank"].map(short), t["cc_spend_per_card_rs"], color=t["category"].map(CAT_COLORS))
    ax.set_title("Monthly spend per credit card (banks with 1 lakh+ cards)")
    ax.set_xlabel("Rs per card per month")
    save(fig, "04_cc_spend_per_card.png")

    # 5. Cash vs digital: debit cards (ATM cash vs POS+online)
    t = d.nlargest(12, "debit_cards").iloc[::-1]
    fig, ax = plt.subplots(figsize=(9, 5.5))
    ax.barh(t["bank"].map(short), t["dc_atm_val"] * THOUSAND_TO_CRORE, label="ATM cash withdrawal", color="#c44e52")
    ax.barh(t["bank"].map(short), t["dc_spend_cr"],
            left=t["dc_atm_val"] * THOUSAND_TO_CRORE, label="PoS + online + other spend", color="#55a868")
    ax.set_title("Debit cards: cash out vs spend (Rs crore)")
    ax.legend(frameon=False)
    save(fig, "05_debit_cash_vs_spend.png")

    # 6. Cash dependence %
    t = d[(d["debit_cards"] >= 1_000_000) & (d["atm_total"] > 0)].sort_values("dc_cash_dependence_pct").tail(15)
    fig, ax = plt.subplots(figsize=(8, 5.5))
    ax.barh(t["bank"].map(short), t["dc_cash_dependence_pct"], color=t["category"].map(CAT_COLORS))
    ax.set_title("Cash dependence of debit cards\n(% of debit value that was ATM cash; 10 lakh+ cards)")
    ax.set_xlabel("%")
    save(fig, "06_cash_dependence.png")

    # 7. E-commerce share of credit card spend by category
    fig, ax = plt.subplots(figsize=(8, 4.5))
    c = cat.sort_values("cc_online_share_pct")
    ax.barh(c["category"], c["cc_online_share_pct"], color=c["category"].map(CAT_COLORS))
    ax.set_title("Online (e-com) share of credit card spend")
    ax.set_xlabel("%")
    for i, v in enumerate(c["cc_online_share_pct"]):
        ax.text(v, i, f" {v:.1f}%", va="center", fontsize=9)
    save(fig, "07_cc_online_share.png")

    # 8. Scatter: cards per ATM vs spend per debit card
    t = d[(d["atm_total"] > 100) & (d["debit_cards"] > 500_000)]
    fig, ax = plt.subplots(figsize=(8.5, 6))
    ax.scatter(t["cards_per_atm"], t["dc_spend_per_card_rs"],
               s=t["debit_cards"] / 3e5, c=t["category"].map(CAT_COLORS), alpha=0.75, edgecolor="white")
    for _, r in t.nlargest(8, "debit_cards").iterrows():
        ax.annotate(short(r["bank"]), (r["cards_per_atm"], r["dc_spend_per_card_rs"]), fontsize=8,
                    xytext=(4, 4), textcoords="offset points")
    ax.set_xlabel("Cards per ATM (network load)")
    ax.set_ylabel("Debit spend per card (Rs/month)")
    ax.set_title("Network load vs card usage (bubble = debit card base)")
    save(fig, "08_load_vs_usage.png")

    # 9. Average ticket size: credit vs debit by category
    fig, ax = plt.subplots(figsize=(9, 4.8))
    x = np.arange(len(cat))
    ax.bar(x - 0.2, cat["cc_avg_ticket_rs"], 0.4, label="Credit card", color="#4c72b0")
    ax.bar(x + 0.2, cat["dc_avg_ticket_rs"], 0.4, label="Debit card", color="#dd8452")
    ax.set_xticks(x)
    ax.set_xticklabels(cat["category"], rotation=15)
    ax.set_title("Average transaction size (Rs)")
    ax.legend(frameon=False)
    save(fig, "09_avg_ticket.png")

    # 10. Pareto: cumulative share of credit card spend
    p = d.sort_values("cc_spend_val", ascending=False).reset_index(drop=True)
    p = p[p["cc_spend_val"] > 0]
    cum = p["cc_spend_val"].cumsum() / p["cc_spend_val"].sum() * 100
    fig, ax = plt.subplots(figsize=(9, 4.8))
    ax.bar(range(len(p)), p["cc_spend_val"] * THOUSAND_TO_CRORE, color="#8da0cb")
    ax2 = ax.twinx()
    ax2.plot(range(len(p)), cum, color="#c44e52", marker="o", ms=3)
    ax2.axhline(80, ls="--", color="grey", lw=0.8)
    ax2.set_ylim(0, 105)
    ax.set_xticks(range(len(p)))
    ax.set_xticklabels(p["bank"].map(short), rotation=90, fontsize=7)
    ax.set_ylabel("Credit card spend (Rs crore)")
    ax2.set_ylabel("Cumulative %")
    ax.set_title("Pareto: credit card spending is concentrated in a few banks")
    ax2.spines["right"].set_visible(True)
    save(fig, "10_pareto_cc_spend.png")


# --------------------------------------------------------------------------
# 5. AUTO INSIGHTS
# --------------------------------------------------------------------------
def build_insights(d, cat, conc) -> str:
    L = []
    tot = lambda c: d[c].sum()
    L.append("KEY INSIGHTS - ATM & CARD STATISTICS, SEPTEMBER 2026")
    L.append("=" * 60)
    L.append(f"Banks in dataset: {len(d)} | Active (have ATMs): {(d['atm_total'] > 0).sum()}")
    L.append("")
    L.append("SCALE")
    L.append(f"- Total ATMs/CRMs: {tot('atm_total'):,.0f}  (off-site share {tot('atm_offsite')/tot('atm_total')*100:.1f}%)")
    L.append(f"- PoS terminals: {tot('pos_terminals'):,.0f}")
    L.append(f"- UPI QR codes: {tot('upi_qr'):,.0f} | Bharat QR: {tot('bharat_qr'):,.0f}")
    L.append(f"- Credit cards: {tot('credit_cards')/1e6:,.1f} mn | Debit cards: {tot('debit_cards')/1e6:,.1f} mn")
    L.append(f"- Credit card spend: Rs {tot('cc_spend_cr'):,.0f} crore | Debit card spend: Rs {tot('dc_spend_cr'):,.0f} crore")
    L.append(f"- ATM cash withdrawn: Rs {tot('atm_cash_val_cr'):,.0f} crore")
    L.append("")
    L.append("CASH vs DIGITAL")
    dc_cash = tot("dc_atm_val") * THOUSAND_TO_CRORE
    L.append(f"- Debit cards: Rs {dc_cash:,.0f} cr taken out as cash vs Rs {tot('dc_spend_cr'):,.0f} cr spent "
             f"(cash = {dc_cash/(dc_cash+tot('dc_spend_cr'))*100:.1f}% of debit value)")
    L.append(f"- Credit cards: e-com is {tot('cc_online_val')/tot('cc_spend_val')*100:.1f}% of spend; "
             f"cash advances are only {tot('cc_atm_val')/(tot('cc_atm_val')+tot('cc_spend_val'))*100:.2f}%")
    L.append("")
    L.append("CATEGORY VIEW")
    for _, r in cat.iterrows():
        L.append(f"- {r['category']}: {r['share_atms_pct']:.1f}% of ATMs, "
                 f"{r['share_cc_spend_pct']:.1f}% of credit spend, {r['share_dc_spend_pct']:.1f}% of debit spend")
    L.append("")
    L.append("CONCENTRATION")
    for _, r in conc.iterrows():
        L.append(f"- {r['metric']}: top-1 = {r['top1_bank']} ({r['top1_share_pct']:.1f}%), "
                 f"top-5 = {r['top5_share_pct']:.1f}%, HHI = {r['HHI']:.0f}")
    L.append("")
    L.append("STANDOUTS")
    b = d[d["credit_cards"] >= 100_000]
    r = b.loc[b["cc_spend_per_card_rs"].idxmax()]
    L.append(f"- Highest spend per credit card (1L+ cards): {r['bank']} - Rs {r['cc_spend_per_card_rs']:,.0f}/card")
    b = d[(d["debit_cards"] >= 1_000_000) & (d["atm_total"] > 0)]  # banks that actually run ATMs
    r = b.loc[b["dc_cash_dependence_pct"].idxmax()]
    L.append(f"- Most cash-dependent debit base (10L+ cards): {r['bank']} - {r['dc_cash_dependence_pct']:.1f}%")
    r = b.loc[b["dc_cash_dependence_pct"].idxmin()]
    L.append(f"- Least cash-dependent debit base (10L+ cards): {r['bank']} - {r['dc_cash_dependence_pct']:.1f}%")
    b = d[d["atm_total"] >= 100]
    r = b.loc[b["cards_per_atm"].idxmax()]
    L.append(f"- Most stretched ATM network (cards per ATM, 100+ ATMs): {r['bank']} - {r['cards_per_atm']:,.0f}")
    zero = d[(d["atm_total"] == 0) & (d["total_cards"] > 0)]
    L.append(f"- Banks with cards but ZERO ATMs: {len(zero)} -> {', '.join(zero['bank'].map(short))}")
    return "\n".join(L)


# --------------------------------------------------------------------------
# 6. MAIN
# --------------------------------------------------------------------------
def main():
    if not FILE_PATH.exists():
        sys.exit(f"File not found: {FILE_PATH.resolve()}\n"
                 "Put the Excel file next to this script or pass its path as an argument.")

    print(f"Loading {FILE_PATH.name} ...")
    df = load_and_clean(FILE_PATH)
    print(f"  -> {len(df)} banks across {df['category'].nunique()} categories")

    # Sanity check against the 'Total' row of the sheet
    raw = pd.read_excel(FILE_PATH, header=None)
    tot_row = raw[raw[1].astype(str).str.strip() == "Total"]
    if not tot_row.empty:
        sheet_atm = tot_row.iloc[0, 3] + tot_row.iloc[0, 4]
        mine = df["atm_onsite"].sum() + df["atm_offsite"].sum()
        print(f"  -> Check: ATM total in sheet = {sheet_atm:,.0f} | mine = {mine:,.0f} | "
              f"{'OK' if abs(sheet_atm - mine) < 1 else 'MISMATCH'}")

    d = add_kpis(df)
    cat = category_summary(d)
    conc = concentration(d)

    d.to_csv(OUT_DIR / "clean_bank_data.csv", index=False)
    with pd.ExcelWriter(OUT_DIR / "analysis_summary.xlsx") as xw:
        d.to_excel(xw, sheet_name="Bank_level_KPIs", index=False)
        cat.to_excel(xw, sheet_name="Category_summary", index=False)
        conc.to_excel(xw, sheet_name="Concentration", index=False)
        d.nlargest(10, "atm_total")[["bank", "category", "atm_total", "share_atm_pct"]].to_excel(xw, sheet_name="Top10_ATMs", index=False)
        d.nlargest(10, "cc_spend_val")[["bank", "category", "cc_spend_cr", "share_cc_spend_pct"]].to_excel(xw, sheet_name="Top10_CC_spend", index=False)
        d.nlargest(10, "dc_spend_val")[["bank", "category", "dc_spend_cr", "share_dc_spend_pct"]].to_excel(xw, sheet_name="Top10_DC_spend", index=False)

    make_charts(d, cat)
    insights = build_insights(d, cat, conc)
    (OUT_DIR / "key_insights.txt").write_text(insights, encoding="utf-8")

    print("\n" + insights)
    print(f"\nDone. Everything saved in: {OUT_DIR.resolve()}")


if __name__ == "__main__":
    main()
