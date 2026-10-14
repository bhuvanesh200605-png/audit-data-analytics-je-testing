"""Audit Data Analytics: Journal Entry & Accounts Payable testing (free, offline).
Pipeline: 1 data -> 2 validate -> 3 materiality -> 4 tests -> 5 risk score -> 6 sample -> 7 workbook."""
import numpy as np, pandas as pd
from scipy.stats import chisquare
from openpyxl import load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill

rng = np.random.default_rng(42)
REVENUE = 50_000_000          # audit planning assumption (client revenue)
APPROVAL_LIMIT = 10_000       # client policy: payments >= this need manager approval

# ---------- 1. DATA (synthetic client ledger; replace with a real extract) ----------
n = 3000
amt = np.round(10 ** rng.uniform(1.5, 5.2, n), 2)                     # log-uniform => follows Benford
df = pd.DataFrame({
    "je_id": [f"JE{i:05d}" for i in range(n)],
    "posting_date": pd.to_datetime("2025-01-01") + pd.to_timedelta(rng.integers(0, 365, n), "D"),
    "vendor": rng.choice([f"V{i:03d}" for i in range(60)], n),
    "amount": amt,
    "created_by": rng.choice(["alice", "bob", "carol", "dan"], n),
})
df["approved_by"] = rng.choice(["mgr1", "mgr2"], n)
# inject known anomalies so every test has something to find
dup = df.sample(25, random_state=1).copy(); dup["je_id"] = [f"JE{n+i:05d}" for i in range(25)]
dup["posting_date"] += pd.Timedelta(days=2); df = pd.concat([df, dup], ignore_index=True)
idx = df.sample(40, random_state=2).index; df.loc[idx, "amount"] = rng.choice([5000, 10000, 20000, 50000], 40)  # round
idx = df.sample(30, random_state=3).index; df.loc[idx, "approved_by"] = df.loc[idx, "created_by"]               # SoD
split = pd.DataFrame({"je_id": [f"JE9{i:04d}" for i in range(12)], "posting_date": pd.Timestamp("2025-06-30"),
                      "vendor": "V007", "amount": 9_800.0, "created_by": "bob", "approved_by": "mgr1"})
df = pd.concat([df, split], ignore_index=True)
df.loc[df.sample(8, random_state=4).index, "posting_date"] = pd.NaT                                              # bad data
df["debit"] = df.amount; df["credit"] = df.amount
df.loc[df.sample(5, random_state=5).index, "credit"] -= 100                                                      # unbalanced

# ---------- 2. DATA VALIDATION (completeness & integrity before any testing) ----------
val = {
    "Total rows": len(df),
    "Missing posting date": int(df.posting_date.isna().sum()),
    "Duplicate JE IDs": int(df.je_id.duplicated().sum()),
    "Unbalanced entries (debit != credit)": int((df.debit.round(2) != df.credit.round(2)).sum()),
    "Non-positive amounts": int((df.amount <= 0).sum()),
    "Total value": round(df.amount.sum(), 2),
}
clean = df.dropna(subset=["posting_date"]).copy()    # exclude unusable rows, but report them

# ---------- 3. MATERIALITY (ISA 320 logic) ----------
OM = 0.005 * REVENUE; PM = 0.75 * OM; CT = 0.05 * OM
mat = {"Overall materiality (0.5% revenue)": OM, "Performance materiality (75%)": PM, "Clearly trivial (5%)": CT}

# ---------- 4. TESTS ----------
# T1 Benford first-digit
fd = clean.amount.astype(str).str.lstrip("0.").str[0].astype(int)
obs = fd.value_counts().reindex(range(1, 10), fill_value=0)
exp = np.log10(1 + 1 / np.arange(1, 10)) * len(fd)
chi, p = chisquare(obs, exp * obs.sum() / exp.sum())
mad = (np.abs(obs / obs.sum() - exp / exp.sum())).mean()
benford = pd.DataFrame({"Digit": range(1, 10), "Actual %": (obs / obs.sum()).values.round(4),
                        "Expected %": (exp / exp.sum()).round(4)})
# T2 Duplicate payments: same vendor + amount within 7 days
s = clean.sort_values(["vendor", "amount", "posting_date"]); g = s.groupby(["vendor", "amount"])
s["gap"] = g.posting_date.diff().dt.days
t_dup = s[(s.gap <= 7)][["je_id", "vendor", "amount", "posting_date", "gap"]]
# T3 Round-amount entries (>= 5,000 and multiple of 1,000)
t_round = clean[(clean.amount >= 5000) & (clean.amount % 1000 == 0)][["je_id", "vendor", "amount", "posting_date"]]
# T4 Weekend postings
t_wknd = clean[clean.posting_date.dt.dayofweek >= 5][["je_id", "vendor", "amount", "posting_date", "created_by"]]
# T5 Segregation of duties: creator == approver
t_sod = clean[clean.created_by == clean.approved_by][["je_id", "amount", "created_by", "approved_by"]]
# T6 Threshold splitting: same vendor/day, each below limit, together above it
d = clean[clean.amount < APPROVAL_LIMIT].groupby(["vendor", "posting_date"]).agg(count=("amount", "size"), total=("amount", "sum")).reset_index()
t_split = d[(d["count"] >= 2) & (d.total >= APPROVAL_LIMIT) & (d.total / d["count"] > 0.8 * APPROVAL_LIMIT)]

# ---------- 5. RISK SCORE (weighted flags) ----------
w = {"dup": 3, "round": 1, "wknd": 1, "sod": 3, "split": 3}
clean["score"] = (clean.je_id.isin(t_dup.je_id) * w["dup"] + clean.je_id.isin(t_round.je_id) * w["round"]
                  + clean.je_id.isin(t_wknd.je_id) * w["wknd"] + clean.je_id.isin(t_sod.je_id) * w["sod"]
                  + clean.set_index(["vendor", "posting_date"]).index.isin(t_split.set_index(["vendor", "posting_date"]).index) * w["split"])
clean["above_PM"] = clean.amount >= PM
risky = clean[clean.score > 0].sort_values(["score", "amount"], ascending=False)

# ---------- 6. SAMPLE (all high-risk + random remainder, documented seed) ----------
high = risky[risky.score >= 3]
rest = clean[~clean.je_id.isin(high.je_id)].sample(25, random_state=7)
sample = pd.concat([high.assign(basis="High risk (100%)"), rest.assign(basis="Random (seed 7)")])

# ---------- 7. WORKBOOK ----------
summary = pd.DataFrame([
    ("Benford chi-square p-value", round(p, 4), "p<0.05 => digits deviate from expectation (injected round amounts)"),
    ("Benford MAD", round(mad, 4), "<0.006 close, 0.006-0.012 acceptable, >0.012 nonconformity"),
    ("Possible duplicate payments", len(t_dup), "Confirm with invoices/bank"),
    ("Round-amount entries", len(t_round), "Inquire on business purpose"),
    ("Weekend postings", len(t_wknd), "Check authority/overrides"),
    ("SoD violations (creator = approver)", len(t_sod), "Control deficiency candidate"),
    ("Threshold-splitting groups", len(t_split), "Possible approval-limit circumvention"),
    ("Flagged entries above performance materiality", int(risky.above_PM.sum()), "Prioritise"),
], columns=["Test", "Result", "Interpretation"])
with pd.ExcelWriter("KPMG_style_JE_Testing.xlsx", engine="openpyxl") as xw:
    summary.to_excel(xw, sheet_name="Summary", index=False)
    pd.DataFrame(list(val.items()), columns=["Data validation", "Value"]).to_excel(xw, sheet_name="Data_Validation", index=False)
    pd.DataFrame(list(mat.items()), columns=["Materiality", "Amount"]).to_excel(xw, sheet_name="Materiality", index=False)
    benford.to_excel(xw, sheet_name="Benford", index=False)
    for name, t in [("Duplicates", t_dup), ("Round", t_round), ("Weekend", t_wknd), ("SoD", t_sod), ("Splitting", t_split)]:
        t.to_excel(xw, sheet_name=name, index=False)
    risky.head(200).to_excel(xw, sheet_name="Risk_Ranked", index=False)
    sample.to_excel(xw, sheet_name="Sample", index=False)
wb = load_workbook("KPMG_style_JE_Testing.xlsx")
for ws in wb:
    for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="00338D")
    for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 24
ws = wb["Benford"]; ch = LineChart(); ch.title = "Benford: actual vs expected"
ch.add_data(Reference(ws, min_col=2, max_col=3, min_row=1, max_row=10), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=10)); ws.add_chart(ch, "E2")
wb.save("KPMG_style_JE_Testing.xlsx")
print(summary.to_string(index=False)); print(val); print(mat)
