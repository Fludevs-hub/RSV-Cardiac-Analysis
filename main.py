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
# Suppress expected library warnings so the analysis output remains readable.
import warnings
# Build file paths in a platform-independent way.
from pathlib import Path

# Provide numeric operations, missing values, and array helpers.
import numpy as np
# Provide data-frame loading, cleaning, grouping, and export operations.
import pandas as pd
# Provide formula-based logistic regression models.
import statsmodels.formula.api as smf
# Provide the statistical tests used by the descriptive tables.
from scipy.stats import fisher_exact, mannwhitneyu, chi2_contingency
# Provide Wilson confidence intervals for prevalence estimates.
from statsmodels.stats.proportion import proportion_confint
# Build the conference-ready Word document.
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
# Build low-level Word XML borders for the three-line table style.
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
# Set Word font sizes and page margins.
from docx.shared import Pt, Cm

warnings.filterwarnings("ignore")   # statsmodels convergence chatter; failures are handled explicitly below

# ----------------------------------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------------------------------
# Input Stata file; change this path when using a different data location.
DATA = Path("./Data/SARI_ILI_Dataset.dta")
# Output directory; ensure it exists before running the script in a new environment.
OUT = Path("./outputs")
# Values interpreted as missing in this extract; extend this list only after checking the data dictionary.
MISSING_CODES = [9999]          # "don't know"/missing sentinel used in this extract
# Maximum plausible length of stay; higher values are treated as date-entry errors.
LOS_MAX_DAYS = 180              # stays beyond this are treated as date-entry errors
# Minimum events required in both RSV groups before fitting an adjusted model.
MIN_EVENTS = 5                  # minimum events per RSV group before fitting an adjusted model

# Regular expression for identifying heart-failure terms in free-text fields.
HF_PATTERN = r"ccf|chf|heart ?failure|cardiac fail|congestive|cardiomyo|cardimyo|right sided heart"
# Regular expression for broader cardiac-disease terms in free-text fields.
CARDIAC_PATTERN = r"heart|cardi|rhd|valv"   # congenital/rheumatic/valvular etc.; "failure to thrive" never matches
# Free-text columns searched for cardiac diagnoses.
FREE_TEXT = ["otherconditionspec", "otherexitdiag", "otherexit"]

# ----------------------------------------------------------------------------------------------------
# 1. Load and clean
# ----------------------------------------------------------------------------------------------------
# Load the raw Stata extract without converting coded categories to labels.
raw = pd.read_stata(DATA, convert_categoricals=False)
# Replace the extract's missing sentinel with pandas' standard missing value.
df = raw.replace(MISSING_CODES, np.nan)

# Track the number of records retained at each analysis stage.
flow = {"Records in surveillance dataset": len(df)}

# Cardiac disease / heart failure definitions (see analyst notes)
# Combine the selected free-text diagnosis fields into one searchable lowercase string per record.
text = df[FREE_TEXT].astype(str).agg(" ".join, axis=1).str.lower()
# Flag records containing a heart-failure keyword.
df["hf"] = text.str.contains(HF_PATTERN)
# Define cardiac disease from coded fields or the free-text keyword search.
df["cardiac"] = (df.heartdisease.eq(1) | df.chroniccardiacexit.eq(1) | df.hf | text.str.contains(CARDIAC_PATTERN))

# Analysis population = everyone with an RSV RT-PCR result (RSV testing began in 2024 in this extract).
# Keep this filter aligned with the study protocol when changing the analysis population.
t = df[df.rsvresult.notna()].copy()
# Record the size of the RSV-tested population for the flow summary.
flow["RSV RT-PCR result available"] = len(t)

# Convert the RSV result to the binary exposure variable used in all tables and models.
t["rsv"] = t.rsvresult.astype(int)
# Create descriptive age categories used in prevalence and baseline tables.
t["agegrp"] = pd.cut(t.ageyears, [-1, 0.5, 1, 5, 200], right=False, labels=["<6 months", "6–11 months", "1–4 years", "≥5 years"])
# Model age: ≥1 year collapsed because the ≥5 y cell is near-empty in the cardiac group (causes separation)
# Create the less granular age variable used by regression models.
t["age_m"] = pd.cut(t.ageyears, [-1, 0.5, 1, 200], right=False, labels=["<6m", "6-11m", "1y+"])
# Encode the assumed male sex code while preserving missing sex values.
t["male"] = t.gender.eq(1).astype(float).where(t.gender.notna())
# Mark observations enrolled at the national referral hospital.
t["knh"] = t.site.eq("Kenyatta NH").astype(int)      # most cardiac cases come from the national referral hospital
# Carry the source malnutrition indicator into the analysis data.
t["malnut"] = t.malnutrition

# Severity markers. Missing stays missing: do not code "not recorded" as "no".
t["hypox"] = (t.oxygensat < 90).astype(float).where(t.oxygensat.notna())
t["icu_acu"] = t.icu
t["mv"] = t.mechventil.where(t.mechventil.isin([0, 1]))
t["died"] = t.outcome.eq(2).astype(float).where(t.outcome.isin([1, 2, 3, 4, 5]))   # 3–5 = refused/absconded/referred: alive at exit
t["severe"] = t[["hypox", "icu_acu", "mv", "died"]].max(axis=1)                   # any marker = 1; all missing = NaN
t["prevhosp"] = t.hospitalized12mnths

# Length of stay: only for children discharged alive or who died (known end of the admission).
# Calculate the admission duration in calendar days.
t["los"] = (t.dateoutcome - t.dateofadmission).dt.days.where(t.outcome.isin([1, 2]))
# Remove negative and implausibly long stays before summarising resource use.
t.loc[~t.los.between(0, LOS_MAX_DAYS), "los"] = np.nan
# Sum recorded ICU and acute-care days while retaining missingness when both are absent.
t["crit_days"] = t[["icudays", "acudays"]].sum(axis=1, min_count=1)

# Restrict the main analysis to participants classified as having cardiac disease.
c = t[t.cardiac].copy()
flow["Cardiac disease"] = len(c)
flow["  of whom documented heart failure"] = int(c.hf.sum())
flow["  of whom RSV-positive"] = int(c.rsv.sum())
POS, NEG = c[c.rsv == 1], c[c.rsv == 0]

# ----------------------------------------------------------------------------------------------------
# 2. Formatting helpers
# ----------------------------------------------------------------------------------------------------
def n_pct(s):
    # Remove unavailable observations before calculating the percentage.
    s = s.dropna()
    # Return n/N (%) or an en dash when there are no usable observations.
    return f"{int(s.sum())}/{len(s)} ({100 * s.mean():.1f})" if len(s) else "–"

def med_iqr(s):
    # Exclude missing values before calculating the median and IQR.
    s = s.dropna()
    # Return the formatted median (IQR) or an en dash when there are no observations.
    return f"{s.median():.0f} ({s.quantile(.25):.0f}–{s.quantile(.75):.0f})" if len(s) else "–"

def fmt_p(p):
    # Use an en dash for unavailable p-values and a consistent display precision otherwise.
    return "–" if p is None or np.isnan(p) else ("<0.001" if p < 0.001 else f"{p:.3f}")

def p_binary(a, b):
    # Remove missing binary observations before building the two-by-two table.
    a, b = a.dropna(), b.dropna()
    # Return the two-sided Fisher exact-test p-value.
    return fisher_exact([[a.sum(), len(a) - a.sum()], [b.sum(), len(b) - b.sum()]])[1]

def or_ci(model, term="rsv"):
    # Convert model coefficients and confidence limits from log-odds to odds ratios.
    lo, hi = np.exp(model.conf_int().loc[term])
    # Return the odds ratio with its 95% confidence interval.
    return f"{np.exp(model.params[term]):.2f} ({lo:.2f}–{hi:.2f})"

def logit_or(d, y, covars):
    """Crude and adjusted odds ratio for RSV. Returns '–' when events are too sparse to model reliably."""
    # Keep observations complete for the outcome and core model covariates.
    d = d.dropna(subset=[y, "age_m", "male"]).copy()
    # Avoid unstable adjusted models when either RSV group has too few events.
    if (d.groupby("rsv")[y].sum() < MIN_EVENTS).any():
        return "–", "–", np.nan
    try:
        # Fit the unadjusted model for the crude odds ratio.
        crude = smf.logit(f"{y} ~ rsv", d).fit(disp=0)
        # Fit the adjusted model using the supplied covariates.
        adj = smf.logit(f"{y} ~ rsv + " + " + ".join(covars), d).fit(disp=0)
    except np.linalg.LinAlgError:          # separation despite the event rule: report unadjusted test instead
        # Return placeholders when numerical separation prevents estimation.
        return "–", "–", np.nan
    # Return formatted estimates and the adjusted RSV p-value.
    return or_ci(crude), or_ci(adj), adj.pvalues["rsv"]

ADJ = ["C(age_m)", "male", "knh"]

# ----------------------------------------------------------------------------------------------------
# 3. Tables
# ----------------------------------------------------------------------------------------------------
# Table 1 – baseline characteristics of children with cardiac disease, by RSV status
# Start Table 1 with continuous age and its non-parametric comparison.
rows = [["Age, months, median (IQR)", med_iqr(POS.ageyears * 12), med_iqr(NEG.ageyears * 12),
         fmt_p(mannwhitneyu(POS.ageyears, NEG.ageyears).pvalue)]]
# Build the age-group contingency table for the overall chi-square test.
age_tab = pd.crosstab(c.agegrp, c.rsv)
# Add the age-group test result before adding each individual category.
rows.append(["Age group, n (%)", "", "", fmt_p(chi2_contingency(age_tab.loc[age_tab.sum(axis=1) > 0])[1])])
# Add only age groups represented in the cardiac-disease population.
for g in c.agegrp.cat.categories:
    if (c.agegrp == g).any():
        rows.append([f"   {g}", n_pct(POS.agegrp.eq(g)), n_pct(NEG.agegrp.eq(g)), ""])
# Add binary baseline characteristics and their Fisher exact-test p-values.
for lab, v in [("Male sex", "male"), ("Documented heart failure", "hf"), ("Malnutrition", "malnut"),
               ("Enrolled at Kenyatta National Hospital", "knh"), ("Hospitalised in previous 12 months†", "prevhosp")]:
    rows.append([lab, n_pct(POS[v].astype(float)), n_pct(NEG[v].astype(float)), fmt_p(p_binary(POS[v].astype(float), NEG[v].astype(float)))])
# Convert the accumulated rows into the final Table 1 data frame.
T1 = pd.DataFrame(rows, columns=["Characteristic", f"RSV-positive (n={len(POS)})", f"RSV-negative (n={len(NEG)})", "p-value"])

# Table 2 – RSV prevalence by subgroup
# Build one prevalence row with a Wilson 95% confidence interval.
def prev_row(label, d):
    # Count tested participants and RSV-positive participants in the requested subgroup.
    k, n = int(d.rsv.sum()), len(d); lo, hi = proportion_confint(k, n, method="wilson")
    # Return the values in the column order used by Table 2.
    return [label, n, k, f"{100 * k / n:.1f} ({100 * lo:.1f}–{100 * hi:.1f})"]

# Add prevalence estimates for the full population and key cardiac subgroups.
rows = [prev_row("All RSV-tested SARI/ILI patients", t), prev_row("Without cardiac disease", t[~t.cardiac]),
        prev_row("With cardiac disease", c)]
# Include cardiac age subgroups only when at least five participants are available.
rows += [prev_row(f"   Age {g}", c[c.agegrp == g]) for g in c.agegrp.cat.categories if (c.agegrp == g).sum() >= 5]
# Add documented heart failure as a separate subgroup.
rows += [prev_row("   Documented heart failure", c[c.hf])]
# Convert the prevalence rows into the final Table 2 data frame.
T2 = pd.DataFrame(rows, columns=["Group", "Tested, n", "RSV-positive, n", "Prevalence, % (95% CI)"])
# Test the crude cardiac versus non-cardiac RSV prevalence difference.
p_card = fisher_exact(pd.crosstab(t.cardiac, t.rsv))[1]

# Table 3 – severity and outcomes by RSV status (children with cardiac disease)
# Start an empty list because Table 3 is built one outcome at a time.
rows = []
# Calculate group summaries and crude/adjusted associations for each binary outcome.
for lab, v in [("SpO₂ <90% at admission", "hypox"), ("Supplemental oxygen", "oxygen"), ("ICU / acute care unit admission", "icu_acu"),
               ("Mechanical ventilation", "mv"), ("Chest indrawing", "chestindrawing"), ("Composite severe outcome‡", "severe"),
               ("In-hospital death", "died")]:
    crude, adj, p = logit_or(c, v, ADJ)
    rows.append([lab, n_pct(POS[v]), n_pct(NEG[v]), crude, adj, fmt_p(p) if not np.isnan(p) else fmt_p(p_binary(POS[v], NEG[v])) + "§"])
# Add length of stay separately because it is continuous rather than binary.
rows.append(["Length of stay, days, median (IQR)", med_iqr(POS.los), med_iqr(NEG.los), "–", "–",
             fmt_p(mannwhitneyu(POS.los.dropna(), NEG.los.dropna()).pvalue) + "¶"])
# Convert the outcome rows into the final Table 3 data frame.
T3 = pd.DataFrame(rows, columns=["Outcome", f"RSV-positive (n={len(POS)})", f"RSV-negative (n={len(NEG)})",
                                 "Crude OR (95% CI)", "Adjusted OR (95% CI)*", "p-value"])

# Table 4 – health-system resource use (links the evidence to planning and cost of prevention)
# Summarise admissions, bed-days, critical-care use, and deaths for one group.
def resource(d):
    # Keep only valid length-of-stay observations for bed-day calculations.
    los = d.los.dropna(); cd = d.crit_days.dropna()
    # Return totals and denominators in the order expected by Table 4.
    return [len(d), int(los.sum()), f"{los.mean():.1f}", int(d.icu_acu.sum()), f"{int(cd.sum())} ({len(cd)})",
            int(d.died.sum())]
# Build Table 4 for RSV-positive, RSV-negative, and all cardiac-disease participants.
T4 = pd.DataFrame([["RSV-positive"] + resource(POS), ["RSV-negative"] + resource(NEG), ["All cardiac disease"] + resource(c)],
                  columns=["Group", "Admissions", "Total bed-days", "Mean bed-days per admission",
                           "ICU/ACU admissions", "Recorded ICU/ACU days (n with data)", "Deaths"])

# Table S1 – whole tested population: does cardiac disease modify the RSV–outcome association?
# Start the interaction table with one row per binary outcome.
rows = []
for lab, y in [("Composite severe outcome", "severe"), ("In-hospital death", "died")]:
    # Keep participants complete for the interaction model variables.
    d = t.dropna(subset=[y, "age_m", "male"]).copy(); d["card"] = d.cardiac.astype(int)
    # Fit the interaction model to test whether cardiac disease modifies the RSV association.
    m = smf.logit(f"{y} ~ rsv * card + C(age_m) + male + C(site)", d).fit(disp=0)
    # Fit the main-effects model for adjusted cardiac-disease and RSV odds ratios.
    main_effects = smf.logit(f"{y} ~ rsv + card + C(age_m) + male + C(site)", d).fit(disp=0)
    # Store sample size, group summaries, adjusted effects, and the interaction p-value.
    rows.append([lab, int(m.nobs), n_pct(d[d.card == 1][y]), n_pct(d[d.card == 0][y]),
                 or_ci(main_effects, "card"), or_ci(main_effects, "rsv"), fmt_p(m.pvalues["rsv:card"])])
# Convert interaction results into the supplementary Table S1 data frame.
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
def set_border(cell, **edges):
    # Create a Word table-border element for the supplied edge names and widths.
    tcPr = cell._tc.get_or_add_tcPr(); b = OxmlElement("w:tcBorders")
    for edge, sz in edges.items():
        # Add one border edge using Word's XML namespace and the requested size.
        e = OxmlElement(f"w:{edge}"); e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(sz)); b.append(e)
    # Attach the completed border definition to the table cell.
    tcPr.append(b)

def para(doc_or_cell, text, size=9, bold=False, italic=False, align=None):
    # Add a paragraph with consistent font, spacing, emphasis, and optional alignment.
    p = doc_or_cell.add_paragraph(); r = p.add_run(text)
    r.font.size, r.bold, r.italic = Pt(size), bold, italic
    p.paragraph_format.space_after = Pt(2)
    if align: p.alignment = align
    return p

# Create the Word document and configure it for landscape conference tables.
doc = Document()
sec = doc.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width, sec.page_height = sec.page_height, sec.page_width
for m in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
    setattr(sec, m, Cm(1.8))
style = doc.styles["Normal"]; style.font.name = "Times New Roman"; style.font.size = Pt(10)

para(doc, "RSV infection in children with cardiac disease hospitalised with SARI in Kenya — results tables", 14, bold=True)
para(doc, "Study flow: " + "; ".join(f"{k.strip()} = {v:,}" for k, v in flow.items()), 9, italic=True)

for i, (title, tbl, note) in enumerate(TABLES):
    # Give each table its own page so it can be reused in a poster or presentation.
    if i: doc.add_page_break()                     # one table per page: easy to lift onto a poster or slide
    para(doc, title, 11, bold=True)
    wt = doc.add_table(rows=1, cols=tbl.shape[1]); wt.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths = [Cm(7.5)] + [Cm(18.5 / (tbl.shape[1] - 1))] * (tbl.shape[1] - 1)   # 26 cm usable landscape width
    wt.autofit = False
    for j, col in enumerate(wt.columns): col.width = widths[j]

    # Write one formatted value into a Word table cell.
    def write(cell, j, text, bold=False):
        cell.text = ""; cell.width = widths[j]
        pp = cell.paragraphs[0]; pp.paragraph_format.space_after = Pt(1); pp.paragraph_format.space_before = Pt(1)
        r = pp.add_run(text); r.bold = bold; r.font.size = Pt(9)
        if j: pp.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Write column headings and apply the top/header borders.
    for j, col in enumerate(tbl.columns):
        write(wt.rows[0].cells[j], j, str(col), bold=True)
        set_border(wt.rows[0].cells[j], top=12, bottom=6)
    # Write each pandas row into a new Word table row.
    for _, row in tbl.iterrows():
        cells = wt.add_row().cells
        for j, val in enumerate(row):
            write(cells[j], j, f"{val:,}" if isinstance(val, (int, np.integer)) else str(val))
    # Apply the bottom border to the final row to complete the three-line style.
    for cell in wt.rows[-1].cells:
        set_border(cell, bottom=12)
    para(doc, note, 8, italic=True)

# Save the conference-ready Word tables.
doc.save(OUT / "RSV_Cardiac_Tables.docx")

# Create the editable Excel workbook with one worksheet per table.
with pd.ExcelWriter(OUT / "RSV_Cardiac_Tables.xlsx", engine="openpyxl") as xw:
    for title, tbl, note in TABLES:
        # Derive a short worksheet name from the table title.
        sheet = title.split(".")[0].replace("Table ", "T")
        # Write the table below its title and leave room for notes.
        tbl.to_excel(xw, sheet_name=sheet, index=False, startrow=2)
        # Add the title, note, and practical column widths for review.
        ws = xw.sheets[sheet]; ws["A1"] = title; ws.cell(row=len(tbl) + 5, column=1, value=note)
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 22
        ws.column_dimensions["A"].width = 42
    # Add the study-flow counts as a separate worksheet.
    pd.Series(flow).rename("n").to_frame().to_excel(xw, sheet_name="Flow")

# ----------------------------------------------------------------------------------------------------
# 5. Key numbers for the abstract + sanity checks (fail loudly rather than publish a wrong number)
# ----------------------------------------------------------------------------------------------------
# Confirm the exposure is binary and the RSV subgroups partition the cardiac population.
assert c.rsv.isin([0, 1]).all() and len(POS) + len(NEG) == len(c)
# Confirm Table 4 totals reconcile with the cardiac population and RSV subgroup totals.
assert T4.loc[2, "Admissions"] == len(c) and T4.loc[2, "Total bed-days"] == T4.loc[0, "Total bed-days"] + T4.loc[1, "Total bed-days"]
# Confirm the composite severity indicator is binary wherever it is observed.
assert (t.severe.dropna().isin([0, 1])).all()

# Set wide console display options so printed tables remain inspectable during review.
pd.set_option("display.width", 250, "display.max_columns", 20, "display.max_colwidth", 60)
# Print the study-flow counts for the abstract and audit trail.
print(pd.Series(flow).to_string())
# Print every exported table for a quick console review.
for title, tbl, _ in TABLES:
    print(f"\n{title}\n{tbl.to_string(index=False)}")

# The analysis runs top-to-bottom when this file is executed directly.
# Keep this guard as documentation for future refactoring into a main() function.
if __name__ == "__main__":
    pass
