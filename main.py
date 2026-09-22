"""
RSV infection in children with cardiac disease hospitalised with SARI/ILI in Kenya
==================================================================================
Secondary analysis of KEMRI/CGHR hospital-based SARI/ILI surveillance data prepared for the 17th KASH Conference.

Objectives
  1. Prevalence of RSV among children with cardiac disease (including heart failure).
  2. Association between RSV and illness severity in this group.
  3. Outcomes (length of stay, in-hospital death, prior hospitalisation) and health-system
     resource use (bed-days, ICU/acute-care days) by RSV status.

Inputs   SARI_ILI_Dataset.dta (Stata), coded as in SARI_Refugee_Netbook Data Dictionary (July 2017)
Outputs  RSV_Cardiac_Tables.docx  - conference-ready Tables 1-4 and supplementary Table S1
         RSV_Cardiac_Tables.xlsx  - the same tables, one sheet per table, for editing/reuse
Run      python rsv_cardiac_analysis.py   (needs pandas, numpy, scipy, statsmodels, python-docx, openpyxl)

Analyst notes (read before interpreting)
  * The dataset has no dedicated heart-failure field. "Cardiac disease" is built from the chronic-cardiac
    comorbidity tick-boxes plus a keyword search of three free-text diagnosis fields. Keyword searches
    miss misspellings and pick up false positives, so check the flagged records by hand before publishing.
  * Readmission is NOT captured. "Hospitalised in the previous 12 months" is reported only as a
    descriptive proxy for prior health-care use. It must not be called readmission.
  * Sex coding is not given in the dictionary. We assume 1 = male (standard KEMRI coding). Sex is only
    used as a model covariate, so a wrong assumption does not change the RSV estimates.
  * Small numbers: 55 RSV-positive children with cardiac disease, 19 with heart failure. Adjusted models
    are fitted only when both RSV groups have at least 5 events (see MIN_EVENTS). Treat null findings
    as "no evidence of a difference", not as evidence of no difference.
  * ICU/ACU day counts are recorded for only a minority of ICU/ACU admissions, so they are reported
    as recorded and not imputed.
"""
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import statsmodels.formula.api as smf
from scipy.stats import fisher_exact, mannwhitneyu, chi2_contingency
from statsmodels.stats.proportion import proportion_confint
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm

warnings.filterwarnings("ignore")   # suppress repeated statsmodels warnings so the script can fail cleanly only when needed

# ----------------------------------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------------------------------
# Set the input dataset and output folder and define analysis rules used throughout the script.
DATA = Path("./Data/SARI_ILI_Dataset.dta")
OUT = Path("./outputs")
MISSING_CODES = [9999]          # missing-value sentinel used in this Stata extract
LOS_MAX_DAYS = 180              # treat unusually long stays as likely data-entry errors
MIN_EVENTS = 5                  # only run adjusted models when each RSV group has at least 5 events

HF_PATTERN = r"ccf|chf|heart ?failure|cardiac fail|congestive|cardiomyo|cardimyo|right sided heart"
CARDIAC_PATTERN = r"heart|cardi|rhd|valv"   # congenital/rheumatic/valvular etc.; "failure to thrive" never matches
FREE_TEXT = ["otherconditionspec", "otherexitdiag", "otherexit"]

# ----------------------------------------------------------------------------------------------------
# 1. Load and clean
# ----------------------------------------------------------------------------------------------------
# Read the Stata dataset and replace the sentinel missing code with actual missing values (NaN).
raw = pd.read_stata(DATA, convert_categoricals=False)
df = raw.replace(MISSING_CODES, np.nan)

# Keep track of how many records remain after each filtering step so the study flow is transparent.
flow = {"Records in surveillance dataset": len(df)}

# Build the cardiac disease and heart-failure definitions using both coded variables and free-text diagnosis fields.
text = df[FREE_TEXT].astype(str).agg(" ".join, axis=1).str.lower()
df["hf"] = text.str.contains(HF_PATTERN)
df["cardiac"] = (df.heartdisease.eq(1) | df.chroniccardiacexit.eq(1) | df.hf | text.str.contains(CARDIAC_PATTERN))

# Restrict the analysis dataset to children with a valid RSV RT-PCR result and create the main study variables.
t = df[df.rsvresult.notna()].copy()
flow["RSV RT-PCR result available"] = len(t)

t["rsv"] = t.rsvresult.astype(int)
t["agegrp"] = pd.cut(t.ageyears, [-1, 0.5, 1, 5, 200], right=False, labels=["<6 months", "6–11 months", "1–4 years", "≥5 years"])
# Collapse the oldest age groups for modeling to avoid unstable logistic regression when a cell is too sparse.
t["age_m"] = pd.cut(t.ageyears, [-1, 0.5, 1, 200], right=False, labels=["<6m", "6-11m", "1y+"])
t["male"] = t.gender.eq(1).astype(float).where(t.gender.notna())
t["knh"] = t.site.eq("Kenyatta NH").astype(int)      # most cardiac cases were enrolled at the national referral hospital
t["malnut"] = t.malnutrition

# Define severity markers; missing values remain missing rather than being treated as "no".
t["hypox"] = (t.oxygensat < 90).astype(float).where(t.oxygensat.notna())
t["icu_acu"] = t.icu
t["mv"] = t.mechventil.where(t.mechventil.isin([0, 1]))
t["died"] = t.outcome.eq(2).astype(float).where(t.outcome.isin([1, 2, 3, 4, 5]))   # 3-5 indicates not alive at discharge; these are treated as alive for the model
t["severe"] = t[["hypox", "icu_acu", "mv", "died"]].max(axis=1)                   # any severe marker = 1; all missing = NaN
t["prevhosp"] = t.hospitalized12mnths

# Calculate length of stay for children with a known end of admission and remove implausible outliers.
t["los"] = (t.dateoutcome - t.dateofadmission).dt.days.where(t.outcome.isin([1, 2]))
t.loc[~t.los.between(0, LOS_MAX_DAYS), "los"] = np.nan
t["crit_days"] = t[["icudays", "acudays"]].sum(axis=1, min_count=1)

# Keep only the children with cardiac disease for the main clinical analysis and summarise the flow.
c = t[t.cardiac].copy()
flow["Cardiac disease"] = len(c)
flow["  of whom documented heart failure"] = int(c.hf.sum())
flow["  of whom RSV-positive"] = int(c.rsv.sum())

# Split the cardiac population into RSV-positive and RSV-negative groups to compare them directly.
POS, NEG = c[c.rsv == 1], c[c.rsv == 0]

# ----------------------------------------------------------------------------------------------------
# 2. Formatting helpers
# ----------------------------------------------------------------------------------------------------
def n_pct(s):
    s = s.dropna(); return f"{int(s.sum())}/{len(s)} ({100 * s.mean():.1f})" if len(s) else "–"

def med_iqr(s):
    s = s.dropna(); return f"{s.median():.0f} ({s.quantile(.25):.0f}–{s.quantile(.75):.0f})" if len(s) else "–"

def fmt_p(p):
    return "–" if p is None or np.isnan(p) else ("<0.001" if p < 0.001 else f"{p:.3f}")

def p_binary(a, b):
    a, b = a.dropna(), b.dropna()
    return fisher_exact([[a.sum(), len(a) - a.sum()], [b.sum(), len(b) - b.sum()]])[1]

def or_ci(model, term="rsv"):
    lo, hi = np.exp(model.conf_int().loc[term]); return f"{np.exp(model.params[term]):.2f} ({lo:.2f}–{hi:.2f})"

def logit_or(d, y, covars):
    """Crude and adjusted odds ratio for RSV. Returns '–' when events are too sparse to model reliably."""
    d = d.dropna(subset=[y, "age_m", "male"]).copy()
    if (d.groupby("rsv")[y].sum() < MIN_EVENTS).any():
        return "–", "–", np.nan
    try:
        crude = smf.logit(f"{y} ~ rsv", d).fit(disp=0)
        adj = smf.logit(f"{y} ~ rsv + " + " + ".join(covars), d).fit(disp=0)
    except np.linalg.LinAlgError:          # separation despite the event rule: report unadjusted test instead
        return "–", "–", np.nan
    return or_ci(crude), or_ci(adj), adj.pvalues["rsv"]

ADJ = ["C(age_m)", "male", "knh"]

# ----------------------------------------------------------------------------------------------------
# 3. Tables
# ----------------------------------------------------------------------------------------------------
# Table 1 – baseline characteristics of children with cardiac disease, by RSV status
rows = [["Age, months, median (IQR)", med_iqr(POS.ageyears * 12), med_iqr(NEG.ageyears * 12),
         fmt_p(mannwhitneyu(POS.ageyears, NEG.ageyears).pvalue)]]
age_tab = pd.crosstab(c.agegrp, c.rsv)
rows.append(["Age group, n (%)", "", "", fmt_p(chi2_contingency(age_tab.loc[age_tab.sum(axis=1) > 0])[1])])
for g in c.agegrp.cat.categories:
    if (c.agegrp == g).any():
        rows.append([f"   {g}", n_pct(POS.agegrp.eq(g)), n_pct(NEG.agegrp.eq(g)), ""])
for lab, v in [("Male sex", "male"), ("Documented heart failure", "hf"), ("Malnutrition", "malnut"),
               ("Enrolled at Kenyatta National Hospital", "knh"), ("Hospitalised in previous 12 months†", "prevhosp")]:
    rows.append([lab, n_pct(POS[v].astype(float)), n_pct(NEG[v].astype(float)), fmt_p(p_binary(POS[v].astype(float), NEG[v].astype(float)))])
T1 = pd.DataFrame(rows, columns=["Characteristic", f"RSV-positive (n={len(POS)})", f"RSV-negative (n={len(NEG)})", "p-value"])

# Table 2 – RSV prevalence by subgroup
def prev_row(label, d):
    k, n = int(d.rsv.sum()), len(d); lo, hi = proportion_confint(k, n, method="wilson")
    return [label, n, k, f"{100 * k / n:.1f} ({100 * lo:.1f}–{100 * hi:.1f})"]

rows = [prev_row("All RSV-tested SARI/ILI patients", t), prev_row("Without cardiac disease", t[~t.cardiac]),
        prev_row("With cardiac disease", c)]
rows += [prev_row(f"   Age {g}", c[c.agegrp == g]) for g in c.agegrp.cat.categories if (c.agegrp == g).sum() >= 5]
rows += [prev_row("   Documented heart failure", c[c.hf])]
T2 = pd.DataFrame(rows, columns=["Group", "Tested, n", "RSV-positive, n", "Prevalence, % (95% CI)"])
p_card = fisher_exact(pd.crosstab(t.cardiac, t.rsv))[1]

# Table 3 – severity and outcomes by RSV status (children with cardiac disease)
rows = []
for lab, v in [("SpO₂ <90% at admission", "hypox"), ("Supplemental oxygen", "oxygen"), ("ICU / acute care unit admission", "icu_acu"),
               ("Mechanical ventilation", "mv"), ("Chest indrawing", "chestindrawing"), ("Composite severe outcome‡", "severe"),
               ("In-hospital death", "died")]:
    crude, adj, p = logit_or(c, v, ADJ)
    rows.append([lab, n_pct(POS[v]), n_pct(NEG[v]), crude, adj, fmt_p(p) if not np.isnan(p) else fmt_p(p_binary(POS[v], NEG[v])) + "§"])
rows.append(["Length of stay, days, median (IQR)", med_iqr(POS.los), med_iqr(NEG.los), "–", "–",
             fmt_p(mannwhitneyu(POS.los.dropna(), NEG.los.dropna()).pvalue) + "¶"])
T3 = pd.DataFrame(rows, columns=["Outcome", f"RSV-positive (n={len(POS)})", f"RSV-negative (n={len(NEG)})",
                                 "Crude OR (95% CI)", "Adjusted OR (95% CI)*", "p-value"])

# Table 4 – health-system resource use (links the evidence to planning and cost of prevention)
def resource(d):
    los = d.los.dropna(); cd = d.crit_days.dropna()
    return [len(d), int(los.sum()), f"{los.mean():.1f}", int(d.icu_acu.sum()), f"{int(cd.sum())} ({len(cd)})",
            int(d.died.sum())]
T4 = pd.DataFrame([["RSV-positive"] + resource(POS), ["RSV-negative"] + resource(NEG), ["All cardiac disease"] + resource(c)],
                  columns=["Group", "Admissions", "Total bed-days", "Mean bed-days per admission",
                           "ICU/ACU admissions", "Recorded ICU/ACU days (n with data)", "Deaths"])

# Table S1 – whole tested population: does cardiac disease modify the RSV–outcome association?
rows = []
for lab, y in [("Composite severe outcome", "severe"), ("In-hospital death", "died")]:
    d = t.dropna(subset=[y, "age_m", "male"]).copy(); d["card"] = d.cardiac.astype(int)
    m = smf.logit(f"{y} ~ rsv * card + C(age_m) + male + C(site)", d).fit(disp=0)
    main = smf.logit(f"{y} ~ rsv + card + C(age_m) + male + C(site)", d).fit(disp=0)
    rows.append([lab, int(m.nobs), n_pct(d[d.card == 1][y]), n_pct(d[d.card == 0][y]),
                 or_ci(main, "card"), or_ci(main, "rsv"), fmt_p(m.pvalues["rsv:card"])])
S1 = pd.DataFrame(rows, columns=["Outcome", "n", "Cardiac disease, n/N (%)", "No cardiac disease, n/N (%)",
                                 "aOR cardiac disease (95% CI)", "aOR RSV (95% CI)", "p (RSV × cardiac interaction)"])

TABLES = [
    ("Table 1. Baseline characteristics of RSV-tested children with cardiac disease hospitalised with SARI/ILI, by RSV status", T1,
     "Values are n/N (%) unless stated. p-values: Fisher's exact test (binary), χ² (age group), Mann–Whitney U (age). "
     "† Descriptive proxy for prior health-care use; readmission after discharge was not captured."),
    ("Table 2. RSV prevalence among hospitalised SARI/ILI patients, by cardiac disease status", T2,
     f"RSV detected by real-time RT-PCR on nasopharyngeal/oropharyngeal swabs. 95% CI by the Wilson method. "
     f"Cardiac vs non-cardiac: Fisher's exact p = {fmt_p(p_card)}. Age subgroups with <5 children are not shown."),
    ("Table 3. Severity and outcomes of children with cardiac disease, by RSV status", T3,
     "Values are n/N (%) unless stated. * Logistic regression adjusted for age group (<6, 6–11, ≥12 months), sex and enrolment site (Kenyatta NH vs other). "
     "‡ Any of SpO₂ <90%, ICU/ACU admission, mechanical ventilation or death. "
     f"§ Fisher's exact test; model not fitted because one group had <{MIN_EVENTS} events. ¶ Mann–Whitney U test."),
    ("Table 4. Hospital resource use among children with cardiac disease hospitalised with SARI/ILI", T4,
     "Bed-days = date of outcome minus date of admission, for children discharged alive or who died (stays of 0–180 days). "
     "ICU/ACU days were recorded for only a subset of ICU/ACU admissions, so the total is a minimum estimate."),
    ("Table S1. Cardiac disease as a risk factor and effect modifier among all RSV-tested SARI/ILI patients", S1,
     "Logistic regression adjusted for age group, sex and site. The interaction p-value tests whether the RSV–outcome "
     "association differs between children with and without cardiac disease."),
]

# ----------------------------------------------------------------------------------------------------
# 4. Export – Word (conference style: three-line tables, Times New Roman) and Excel
# ----------------------------------------------------------------------------------------------------
# Define helper functions to style the Word tables and add explanatory notes under each result table.
def set_border(cell, **edges):
    tcPr = cell._tc.get_or_add_tcPr(); b = OxmlElement("w:tcBorders")
    for edge, sz in edges.items():
        e = OxmlElement(f"w:{edge}"); e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(sz)); b.append(e)
    tcPr.append(b)

def para(doc_or_cell, text, size=9, bold=False, italic=False, align=None):
    p = doc_or_cell.add_paragraph(); r = p.add_run(text)
    r.font.size, r.bold, r.italic = Pt(size), bold, italic
    p.paragraph_format.space_after = Pt(2)
    if align: p.alignment = align
    return p

# Create the Word document and set it up in landscape format so each table fits cleanly on one page.
doc = Document()
sec = doc.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width, sec.page_height = sec.page_height, sec.page_width
for m in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
    setattr(sec, m, Cm(1.8))
style = doc.styles["Normal"]; style.font.name = "Times New Roman"; style.font.size = Pt(10)

# Add the title and a concise study flow summary to the report.
para(doc, "RSV infection in children with cardiac disease hospitalised with SARI in Kenya — results tables", 14, bold=True)
para(doc, "Study flow: " + "; ".join(f"{k.strip()} = {v:,}" for k, v in flow.items()), 9, italic=True)

# Write each analysis table as a full-page, nicely formatted Word table with consistent widths and notes.
for i, (title, tbl, note) in enumerate(TABLES):
    if i: doc.add_page_break()                     # each table gets its own page so it is easy to extract for slides/posters
    para(doc, title, 11, bold=True)
    wt = doc.add_table(rows=1, cols=tbl.shape[1]); wt.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths = [Cm(7.5)] + [Cm(18.5 / (tbl.shape[1] - 1))] * (tbl.shape[1] - 1)   # leave the first column wide enough for labels and keep the rest compact
    wt.autofit = False
    for j, col in enumerate(wt.columns): col.width = widths[j]

    def write(cell, j, text, bold=False):
        cell.text = ""; cell.width = widths[j]
        pp = cell.paragraphs[0]; pp.paragraph_format.space_after = Pt(1); pp.paragraph_format.space_before = Pt(1)
        r = pp.add_run(text); r.bold = bold; r.font.size = Pt(9)
        if j: pp.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for j, col in enumerate(tbl.columns):
        write(wt.rows[0].cells[j], j, str(col), bold=True)
        set_border(wt.rows[0].cells[j], top=12, bottom=6)
    for _, row in tbl.iterrows():
        cells = wt.add_row().cells
        for j, val in enumerate(row):
            write(cells[j], j, f"{val:,}" if isinstance(val, (int, np.integer)) else str(val))
    for cell in wt.rows[-1].cells:
        set_border(cell, bottom=12)
    para(doc, note, 8, italic=True)

doc.save(OUT / "RSV_Cardiac_Tables.docx")

# Export the same tables to Excel, with one sheet per table and a summary sheet for the study flow.
with pd.ExcelWriter(OUT / "RSV_Cardiac_Tables.xlsx", engine="openpyxl") as xw:
    for title, tbl, note in TABLES:
        sheet = title.split(".")[0].replace("Table ", "T")
        tbl.to_excel(xw, sheet_name=sheet, index=False, startrow=2)
        ws = xw.sheets[sheet]; ws["A1"] = title; ws.cell(row=len(tbl) + 5, column=1, value=note)
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 22
        ws.column_dimensions["A"].width = 42
    pd.Series(flow).rename("n").to_frame().to_excel(xw, sheet_name="Flow")

# ----------------------------------------------------------------------------------------------------
# 5. Key numbers for the abstract + sanity checks (fail loudly rather than publish a wrong number)
# ----------------------------------------------------------------------------------------------------
# These assertions are safety checks: if a key value is inconsistent, the script stops instead of producing misleading output.
assert c.rsv.isin([0, 1]).all() and len(POS) + len(NEG) == len(c)
assert T4.loc[2, "Admissions"] == len(c) and T4.loc[2, "Total bed-days"] == T4.loc[0, "Total bed-days"] + T4.loc[1, "Total bed-days"]
assert (t.severe.dropna().isin([0, 1])).all()

# Expand the console output so the tables are readable and print the final study flow plus every result table.
pd.set_option("display.width", 250, "display.max_columns", 20, "display.max_colwidth", 60)
print(pd.Series(flow).to_string())
for title, tbl, _ in TABLES:
    print(f"\n{title}\n{tbl.to_string(index=False)}")


if __name__ == "__main__":
    main()
