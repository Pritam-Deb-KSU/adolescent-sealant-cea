"""
=============================================================================
Probabilistic sensitivity analysis (PSA) and cost-effectiveness
acceptability curves (CEAC) for the adolescent school sealant program model
=============================================================================

Companion to sealant_markov_model.py (the model behind Table 2 / Figures 2–3
of the manuscript).  Nothing in the model is changed here; this file only
(1) draws parameter sets from the distributions in Table 1,
(2) evaluates the model's EXPECTED outcomes for each draw, and
(3) summarises the resulting joint distribution of incremental cost and
    incremental effect.

Design choices (see docs/psa_methods.md for the manuscript text)
----------------------------------------------------------------
* Second-order (parameter) uncertainty only.  Each parameter draw is run
  through the expected-value (cohort) form of the Markov model
  (run_one_scenario_expected), so individual-level binomial noise — which
  averages out in a population and is not decision-relevant — does not
  inflate the PSA intervals (Briggs, Claxton & Sculpher 2006, ch. 4;
  ISPOR-SMDM Modeling Good Research Practices Task Force-6).
* Costs discounted at the manuscript's 3 %/yr (DISCOUNT_RATE of the model).
* Strategies compared for a cohort of 1,000 12-year-olds (4,000 1M + 4,000 2M):
      No sealants | Case 1 (seal 2M) | Case 2 (seal 1M + 2M) | Case 3 (Case 2 + reseal ages 12.5–14)
* Effect measure: averted decayed teeth by age 19 (fillings averted are
  reported too).  Willingness-to-pay (WTP, λ) is expressed in $ per averted
  averted decay; the cost of restoration per tooth (Table 1, ~$131) is a natural reference.
* CEACs: pairwise vs "No sealants" (P[λ·ΔE − ΔC > 0]) and the multi-way
  CEAC / acceptability frontier over the four mutually exclusive strategies.
* Implementation scenario: the Case 1 program delivered by a school-based
  sealant program (SBSP) to 1,000 enrolled 6th-graders, with
      – an uncertain consent rate p ~ Beta (mean 15 %: the Discussion reports a
        30 % median consent rate overall and "approximately only half that level
        for programs for 2M"), and
      – an additional program cost c_s per consenting adolescent (NOT part of
        the submitted model): e.g. a Medicaid screening/assessment claim
        (D0190/D0191 × 49.4 %, the Table 1 method) or per-adolescent
        set-up/logistics cost (Discussion).
  The consent rate scales impact and budget but not the cost per averted
  decay; c_s does change the ratio, so the CEAC is shown for several c_s.
  ►► The manuscript uses IMPL_SCREENING_COSTS = [0.0] (no extra cost input):
     the implementation CEAC then coincides with Case 1 and only the scale
     of impact/budget depends on the consent rate. Other values are for
     exploration only.

Outputs (written to ../results/)
--------------------------------
  psa_results.xlsx   parameter distributions, per-strategy summaries,
                     CEAC table, incremental (frontier) analysis, EVPI,
                     implementation scenario
  psa_draws.csv      one row per PSA draw (parameters + outcomes)
  figure_psa_ce_plane_ceac.png    CE plane + pairwise CEACs (main figure)
  figure_psa_ceac_multiway.png    multi-way CEAC with frontier (supplement)
  figure_psa_implementation.png   implementation-scenario CEACs (main/supplement)
"""
from __future__ import annotations

import csv
from pathlib import Path
from typing import Dict, List, Tuple

import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from sealant_markov_model import (
    AR_BASE, DISCOUNT_RATE, DV_PER_PERIOD, FILLING_COST_TRI, N_CYCLES, N_TOTAL_TEETH,
    PARAM_RANGE_PCT, R_TRI, SCENARIOS, SEALANT_COST_TRI, run_one_scenario_expected,
)

# =====================================================================
# SETTINGS
# =====================================================================
N_PSA        = 10_000
PSA_SEED     = 2026
OUT_DIR      = Path(__file__).resolve().parent.parent / "results"

WTP_GRID     = np.arange(0.0, 200.5, 0.5)          # $ per averted decayed tooth
WTP_REPORT   = [0, 25, 50, 75, 100, FILLING_COST_TRI[1], 150, 200]   # columns of the CEAC table

DV_RANGE_PCT = 0.20     # ±20 % triangular multiplier on the age-specific dental-visit probabilities
                        # (Table 1 gives no sampling SE for the MEPS estimate; ±20 % mirrors the DSA)

COHORT_CHILDREN        = 1_000
NATIONAL_6TH_GRADERS   = 1_700_000   # 6th-graders in the free/reduced-price lunch program (manuscript)

# ── Implementation scenario ───────────────────────────────────────────────────
IMPL_CASE               = "Case 1"       # program actually delivered by the SBSP ("Case 1" or "Case 3")
IMPL_PARTICIPATION_MEAN = 0.15           # Discussion: median consent 30% overall, "approximately only half that level for programs for 2M"
IMPL_PARTICIPATION_N    = 15.0           # Beta "effective sample size": Beta(mean·n, (1−mean)·n) = Beta(2.25, 12.75)
IMPL_SCREENING_COSTS    = [0.0]                     # manuscript: no extra cost input (CEAC = Case 1). e.g. [0, 10, 20, 30] to explore one
IMPL_SCREEN_PERIODS     = {"Case 1": [0], "Case 2": [0], "Case 3": [0, 2, 4]}   # visits at ages 12 (and 13, 14 for Case 3)

STRATEGIES = ["No sealants", "Case 1", "Case 2", "Case 3"]
COLORS     = {"Case 1": "#0072B2", "Case 2": "#D55E00", "Case 3": "#009E73", "No sealants": "#555555"}
LINESTYLES = {"Case 1": "-", "Case 2": "--", "Case 3": "-.", "No sealants": ":"}

DISC = np.array([(1.0 + DISCOUNT_RATE) ** (-t / 2.0) for t in range(N_CYCLES + 1)])


# =====================================================================
# 1. PARAMETER DRAWS  (distributions of Table 1)
# =====================================================================
def tri_pct(rng: np.random.Generator, base: float, pct: float, n: int) -> np.ndarray:
    """Triangular(base·(1−pct), base, base·(1+pct)) — the manuscript's ±20 % distributions."""
    return rng.triangular(base * (1 - pct), base, base * (1 + pct), size=n)


def draw_parameters(rng: np.random.Generator, n: int) -> Dict[str, np.ndarray]:
    return {
        "R":        rng.triangular(*R_TRI, size=n),                      # 6-month sealant retention
        "AR_1M":    tri_pct(rng, AR_BASE["1M"], PARAM_RANGE_PCT, n),      # 6-month attack rate, 1M
        "AR_2M":    tri_pct(rng, AR_BASE["2M"], PARAM_RANGE_PCT, n),      # 6-month attack rate, 2M
        "DV_mult":  rng.triangular(1 - DV_RANGE_PCT, 1.0, 1 + DV_RANGE_PCT, size=n),   # × age-specific DV
        "c_seal":   rng.triangular(*SEALANT_COST_TRI, size=n),           # $ per sealant
        "c_fill":   rng.triangular(*FILLING_COST_TRI, size=n),           # $ per restoration
        "p_part":   rng.beta(IMPL_PARTICIPATION_MEAN * IMPL_PARTICIPATION_N,
                             (1 - IMPL_PARTICIPATION_MEAN) * IMPL_PARTICIPATION_N, size=n),  # implementation only
    }


PARAMETER_TABLE = [  # (name, distribution, base value, source) → written to the xlsx and docs
    ("6-month sealant retention (R)", f"Triangular({R_TRI[0]}, {R_TRI[1]}, {R_TRI[2]})", R_TRI[1], "Quiñonez et al. 2005; ±20% on the 6-month loss rate 1−R"),
    ("6-month attack rate, 2M", f"Triangular(±{PARAM_RANGE_PCT:.0%} of base)", AR_BASE["2M"], "NHANES 2011–16"),
    ("6-month attack rate, 1M", f"Triangular(±{PARAM_RANGE_PCT:.0%} of base)", AR_BASE["1M"], "NHANES 2011–16"),
    ("Annual dental-visit probability (age-specific, ages 12–18)", f"Common multiplier ~ Triangular({1-DV_RANGE_PCT:.1f}, 1, {1+DV_RANGE_PCT:.1f})", "see model", "MEPS 2015–18; ±20% mirrors the deterministic SA"),
    ("Sealant fee per tooth ($)", f"Triangular({SEALANT_COST_TRI[0]}, {SEALANT_COST_TRI[1]}, {SEALANT_COST_TRI[2]})", SEALANT_COST_TRI[1], "ADA 2022 fee survey (D1351) × 49.4%, 25th–75th pct"),
    ("Cost of restoration per tooth (filling, $)", f"Triangular({FILLING_COST_TRI[0]}, {FILLING_COST_TRI[1]}, {FILLING_COST_TRI[2]})", FILLING_COST_TRI[1], "ADA 2022 fee survey (D2392) × 49.4%, 25th–75th pct"),
    ("Annual discount rate", "fixed", DISCOUNT_RATE, "varied in the deterministic SA only"),
    ("Eligible / sealed / unerupted teeth at age 12; eruption schedule", "fixed", "Table 1 / Appendix", "NHANES 2011–16 (structural inputs)"),
    ("Consent rate — implementation scenario", f"Beta(mean {IMPL_PARTICIPATION_MEAN:.2f}, n = {IMPL_PARTICIPATION_N:g})", IMPL_PARTICIPATION_MEAN, "authors' survey: ~half the 30% overall median for 2M programs"),
    ("Additional program cost per consenting adolescent ($) — implementation scenario", "scenario values " + ", ".join(f"{c:g}" for c in IMPL_SCREENING_COSTS), "—", "ILLUSTRATIVE — e.g. Medicaid screening/assessment fee (D0190/D0191 × 49.4%) or per-adolescent set-up/logistics cost"),
]


# =====================================================================
# 2. MODEL EVALUATION PER DRAW
# =====================================================================
def evaluate_draw(R: float, ar_1m: float, ar_2m: float, dv_mult: float,
                  c_seal: float, c_fill: float) -> Dict[str, Dict[str, float]]:
    """Expected discounted costs and outcomes of every modelled scenario for one parameter set."""
    dv = np.clip(DV_PER_PERIOD * dv_mult, 0.0, 1.0)
    out = {}
    for sn, cfg in SCENARIOS.items():
        ar = ar_1m if cfg.tooth_group == "1M" else ar_2m
        seals, fills, decay = run_one_scenario_expected(cfg, R, ar, dv)
        out[sn] = {
            "sealants":     float(seals.sum()),
            "decay":        float(decay),
            "fillings":     float(fills.sum()),
            "sealant_cost": c_seal * float(seals @ DISC),
            "filling_cost": c_fill * float(fills @ DISC),
        }
    for sn, cfg in SCENARIOS.items():
        base = f"{cfg.tooth_group} Base Case"
        out[sn]["net_cost"]        = 0.0 if sn == base else out[sn]["sealant_cost"] - (out[base]["filling_cost"] - out[sn]["filling_cost"])
        out[sn]["averted_decay"]   = out[base]["decay"]    - out[sn]["decay"]
        out[sn]["averted_fillings"]= out[base]["fillings"] - out[sn]["fillings"]
    return out


# Strategy = which modelled scenarios apply to the cohort's 1M and 2M
STRATEGY_MAP = {
    "No sealants": {"1M": "1M Base Case", "2M": "2M Base Case"},
    "Case 1":      {"1M": "1M Base Case", "2M": "2M Case 2"},
    "Case 2":      {"1M": "1M Case 2",    "2M": "2M Case 2"},
    "Case 3":      {"1M": "1M Case 3",    "2M": "2M Case 3"},
}


def run_psa(n: int = N_PSA, seed: int = PSA_SEED):
    rng = np.random.default_rng(seed)
    P = draw_parameters(rng, n)
    keys = ["net_cost", "averted_decay", "averted_fillings", "sealants", "decay", "fillings"]
    # per strategy (cohort totals) and per tooth-group scenario (for Table-2-style intervals)
    S  = {s: {k: np.zeros(n) for k in keys} for s in STRATEGIES}
    G  = {sn: {k: np.zeros(n) for k in keys} for sn in SCENARIOS}
    for i in range(n):
        o = evaluate_draw(P["R"][i], P["AR_1M"][i], P["AR_2M"][i], P["DV_mult"][i], P["c_seal"][i], P["c_fill"][i])
        for sn in SCENARIOS:
            for k in keys:
                G[sn][k][i] = o[sn][k]
        for s, m in STRATEGY_MAP.items():
            for k in keys:
                S[s][k][i] = o[m["1M"]][k] + o[m["2M"]][k]
    return P, S, G


# =====================================================================
# 3. SUMMARIES
# =====================================================================
def pct(x: np.ndarray, q: float) -> float:
    return float(np.quantile(x, q))


def summarize(x: np.ndarray) -> Tuple[float, float, float, float]:
    return float(np.mean(x)), float(np.median(x)), pct(x, 0.025), pct(x, 0.975)


def ceac_pairwise(S, wtp: np.ndarray) -> Dict[str, np.ndarray]:
    """P(NMB > 0) vs No sealants, per strategy and WTP."""
    out = {}
    for s in STRATEGIES[1:]:
        nmb = wtp[None, :] * S[s]["averted_decay"][:, None] - S[s]["net_cost"][:, None]
        out[s] = (nmb > 0).mean(axis=0)
    return out


def ceac_multiway(S, wtp: np.ndarray):
    """P(strategy has the highest NMB) for each of the four mutually exclusive strategies; frontier = argmax E[NMB]."""
    nmb = np.stack([wtp[None, :] * S[s]["averted_decay"][:, None] - S[s]["net_cost"][:, None] for s in STRATEGIES], axis=2)  # n × wtp × strat
    best = nmb.argmax(axis=2)
    probs = {s: (best == k).mean(axis=0) for k, s in enumerate(STRATEGIES)}
    frontier = nmb.mean(axis=0).argmax(axis=1)                       # index of strategy with highest expected NMB
    evpi = nmb.max(axis=2).mean(axis=0) - nmb.mean(axis=0).max(axis=1)   # per cohort of 1,000 adolescents
    return probs, frontier, evpi


def incremental_analysis(S) -> List[Tuple[str, float, float, float, str]]:
    """Fully incremental ICERs along the ordered strategies (expected values)."""
    rows = []
    prev_c = prev_e = 0.0; prev_name = "No sealants"
    for s in STRATEGIES[1:]:
        c, e = S[s]["net_cost"].mean(), S[s]["averted_decay"].mean()
        rows.append((s, c, e, (c - prev_c) / (e - prev_e), f"vs {prev_name}"))
        prev_c, prev_e, prev_name = c, e, s
    return rows


def implementation_scenario(P, S, wtp: np.ndarray):
    """
    Case IMPL_CASE delivered by an SBSP to a cohort of COHORT_CHILDREN 12-year-old adolescents.
    Per draw:  consenting adolescents = p × cohort;  ΔE = p × ΔE_case;
               ΔC = p × (ΔC_case + cohort × c_s × Σ_visits disc[t]).
    """
    p = P["p_part"]
    dE = S[IMPL_CASE]["averted_decay"]; dC = S[IMPL_CASE]["net_cost"]
    screen_disc = float(sum(DISC[t] for t in IMPL_SCREEN_PERIODS[IMPL_CASE]))
    rows = []; curves = {}
    for cs in IMPL_SCREENING_COSTS:
        dC_cs = dC + COHORT_CHILDREN * cs * screen_disc          # per 1,000 enrolled if everyone participated
        nmb = wtp[None, :] * dE[:, None] - dC_cs[:, None]
        curves[cs] = (nmb > 0).mean(axis=0)
        icer = dC_cs / dE
        rows.append({
            "screening_cost": cs,
            "ICER_mean_ratio": dC_cs.mean() / dE.mean(),
            "ICER_2.5": pct(icer, 0.025), "ICER_97.5": pct(icer, 0.975),
            "participants_mean": (p * COHORT_CHILDREN).mean(),
            "averted_decay_per_1000_enrolled": summarize(p * dE),
            "net_cost_per_1000_enrolled": summarize(p * dC_cs),
            "national_averted_decay": summarize(p * dE * NATIONAL_6TH_GRADERS / COHORT_CHILDREN),
            "national_net_cost": summarize(p * dC_cs * NATIONAL_6TH_GRADERS / COHORT_CHILDREN),
            "P_CE_at_fill_fee": float((wtp_nmb_at(dE, dC_cs, FILLING_COST_TRI[1]) > 0).mean()),
        })
    return rows, curves


def wtp_nmb_at(dE, dC, lam):
    return lam * dE - dC


# =====================================================================
# 4. EXCEL / CSV
# =====================================================================
def hdr(ws, row=1):
    for c in ws[row]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def autofit(ws):
    for col in ws.columns:
        w = max((len(str(c.value)) for c in col if c.value is not None), default=8)
        ws.column_dimensions[col[0].column_letter].width = min(max(w + 2, 10), 60)


def write_outputs(P, S, G, pair, multi, frontier, evpi, incr, impl_rows, impl_curves) -> Path:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    wb = Workbook(); wb.remove(wb.active)

    ws = wb.create_sheet("PSA parameters")
    ws.append(["Parameter", "Distribution", "Base value", "Source / note"]); hdr(ws)
    for r in PARAMETER_TABLE: ws.append(list(r))
    ws.append([]); ws.append(["N draws", N_PSA]); ws.append(["Seed", PSA_SEED]); ws.append(["Discount rate", DISCOUNT_RATE])
    ws.append(["Model form", "expected-value (cohort) Markov; second-order uncertainty only"]); autofit(ws)

    ws = wb.create_sheet("Strategy summary")
    ws.append(["Strategy (cohort of 1,000 12-year-old adolescents)", "Outcome", "Mean", "Median", "2.5th pct", "97.5th pct"]); hdr(ws)
    for s in STRATEGIES:
        for k, lab in [("sealants", "Sealants applied"), ("decay", "Cumulative new decay (1M+2M)"), ("fillings", "Cumulative new fillings (1M+2M)"),
                       ("averted_decay", "Averted decay vs No sealants"), ("averted_fillings", "Averted fillings vs No sealants"), ("net_cost", "Net program cost ($, discounted 3%)")]:
            m, md, lo, hi = summarize(S[s][k]); ws.append([s, lab, round(m, 2), round(md, 2), round(lo, 2), round(hi, 2)])
        if s != "No sealants":
            icer = S[s]["net_cost"] / S[s]["averted_decay"]
            ws.append([s, "Cost per averted decay ($) — ratio of means", round(S[s]["net_cost"].mean() / S[s]["averted_decay"].mean(), 2), round(float(np.median(icer)), 2), round(pct(icer, .025), 2), round(pct(icer, .975), 2)])
            icf = S[s]["net_cost"] / S[s]["averted_fillings"]
            ws.append([s, "Cost per averted filling ($) — ratio of means", round(S[s]["net_cost"].mean() / S[s]["averted_fillings"].mean(), 2), round(float(np.median(icf)), 2), round(pct(icf, .025), 2), round(pct(icf, .975), 2)])
    autofit(ws)

    ws = wb.create_sheet("Table 2 PSA intervals")
    ws.append(["Tooth group scenario", "Outcome", "Mean", "2.5th pct", "97.5th pct", "note"]); hdr(ws)
    for sn in SCENARIOS:
        for k in ["sealants", "decay", "fillings", "averted_decay", "averted_fillings", "net_cost"]:
            m, md, lo, hi = summarize(G[sn][k]); ws.append([sn, k, round(m, 2), round(lo, 2), round(hi, 2), "95% credible interval (parameter uncertainty)"])
    autofit(ws)

    ws = wb.create_sheet("CEAC")
    ws.append(["WTP ($ per averted decayed tooth)"] + [f"P(CE) {s} vs No sealants" for s in STRATEGIES[1:]] + [f"P(best) {s}" for s in STRATEGIES] + ["Frontier", "EVPI per 1,000 adolescents ($)"]); hdr(ws)
    for j, lam in enumerate(WTP_GRID):
        if abs(lam - round(lam)) > 1e-9 and lam not in WTP_REPORT: continue
        ws.append([float(lam)] + [round(float(pair[s][j]), 4) for s in STRATEGIES[1:]] + [round(float(multi[s][j]), 4) for s in STRATEGIES] + [STRATEGIES[frontier[j]], round(float(evpi[j]), 2)])
    autofit(ws)

    ws = wb.create_sheet("Incremental (frontier)")
    ws.append(["Strategy", "Expected net cost ($)", "Expected averted decay", "ICER ($ per averted decay)", "Comparator"]); hdr(ws)
    for s, c, e, icer, comp in incr: ws.append([s, round(c, 0), round(e, 1), round(icer, 2), comp])
    autofit(ws)

    ws = wb.create_sheet("Implementation")
    ws.append([f"{IMPL_CASE} delivered by an SBSP to a cohort of {COHORT_CHILDREN:,} 12-year-old adolescents; consent rate ~ Beta(mean {IMPL_PARTICIPATION_MEAN:.0%}); the additional program cost per consenting adolescent is ILLUSTRATIVE"])
    ws.append(["Additional program cost ($ per consenting adolescent)", "ICER (ratio of means)", "ICER 2.5th", "ICER 97.5th", "Consenting adolescents (mean)",
               "Averted decay per 1,000 adolescents: mean", "2.5th", "97.5th", "Net cost per 1,000 adolescents: mean", "2.5th", "97.5th",
               f"National averted decay ({NATIONAL_6TH_GRADERS:,} 6th-grade students in the free/reduced-price lunch program): mean", "2.5th", "97.5th", "National net cost: mean", "2.5th", "97.5th",
               f"P(CE) at λ = cost of restoration per tooth (${FILLING_COST_TRI[1]:.2f})"]); hdr(ws, 2)
    for r in impl_rows:
        a, nc, na, nn = r["averted_decay_per_1000_enrolled"], r["net_cost_per_1000_enrolled"], r["national_averted_decay"], r["national_net_cost"]
        ws.append([r["screening_cost"], round(r["ICER_mean_ratio"], 2), round(r["ICER_2.5"], 2), round(r["ICER_97.5"], 2), round(r["participants_mean"], 1),
                   round(a[0], 1), round(a[2], 1), round(a[3], 1), round(nc[0], 0), round(nc[2], 0), round(nc[3], 0),
                   round(na[0], 0), round(na[2], 0), round(na[3], 0), round(nn[0], 0), round(nn[2], 0), round(nn[3], 0), round(r["P_CE_at_fill_fee"], 4)])
    ws.append([]); ws.append(["CEAC of the implementation scenario"]); ws.append(["WTP ($ per averted decay)"] + [f"c_s = ${cs:g}" for cs in IMPL_SCREENING_COSTS]); hdr(ws, ws.max_row)
    for j, lam in enumerate(WTP_GRID):
        if abs(lam - round(lam)) > 1e-9: continue
        ws.append([float(lam)] + [round(float(impl_curves[cs][j]), 4) for cs in IMPL_SCREENING_COSTS])
    autofit(ws)

    xlsx = OUT_DIR / "psa_results.xlsx"; wb.save(xlsx)

    with open(OUT_DIR / "psa_draws.csv", "w", newline="") as f:
        w = csv.writer(f)
        cols = list(P.keys()) + [f"{s}|{k}" for s in STRATEGIES for k in ["net_cost", "averted_decay", "averted_fillings"]]
        w.writerow(["draw"] + cols)
        for i in range(N_PSA):
            w.writerow([i] + [f"{P[k][i]:.6g}" for k in P] + [f"{S[s][k][i]:.4f}" for s in STRATEGIES for k in ["net_cost", "averted_decay", "averted_fillings"]])
    return xlsx


# =====================================================================
# 5. FIGURES
# =====================================================================
def _mpl():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    plt.rcParams.update({"font.family": "sans-serif", "font.sans-serif": ["Arial", "Liberation Sans", "DejaVu Sans"], "font.size": 10, "axes.spines.top": False, "axes.spines.right": False})
    return plt


def figure_ce_plane_ceac(S, pair, fname="figure_psa_ce_plane_ceac.png"):
    plt = _mpl()
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 4.6), dpi=150)
    rng = np.random.default_rng(1)
    idx = rng.choice(N_PSA, size=min(2500, N_PSA), replace=False)
    for s in STRATEGIES[1:]:
        ax1.scatter(S[s]["averted_decay"][idx], S[s]["net_cost"][idx] / 1000, s=6, alpha=0.18, color=COLORS[s], linewidths=0, rasterized=True)
    for s in STRATEGIES[1:]:
        e, c = S[s]["averted_decay"].mean(), S[s]["net_cost"].mean() / 1000
        ax1.plot(e, c, marker="o", ms=9, mfc="white", mec=COLORS[s], mew=2, zorder=5)
        ax1.annotate(f"{s}\n\\${1000*c/e:.0f}/averted decay", (e, c), xytext=(8, -4), textcoords="offset points", fontsize=8.5, color="#222222")
    ax1.axhline(0, color="#999999", lw=0.8)
    ax1.set_xlabel("Averted decay by age 19 (teeth, per 1,000 12-year-olds)")
    ax1.set_ylabel("Incremental net cost vs no sealants (\\$ thousands)")
    ax1.set_title("(a) Cost-effectiveness plane, 10,000 PSA draws", loc="left", fontweight="bold")
    ax1.grid(True, ls=":", color="#dddddd")

    for s in STRATEGIES[1:]:
        ax2.plot(WTP_GRID, pair[s], color=COLORS[s], ls=LINESTYLES[s], lw=2, label=f"{s} vs no sealants")
    ax2.axvline(FILLING_COST_TRI[1], color="#888888", ls=":", lw=1.2)
    ax2.text(FILLING_COST_TRI[1] + 2, 0.60, "Cost of restoration\nper tooth (Table 1)", fontsize=8, color="#555555")
    ax2.set_xlabel("Willingness to pay ($ per averted decay)")
    ax2.set_ylabel("Probability cost-effective")
    ax2.set_ylim(0, 1.02); ax2.set_xlim(0, WTP_GRID[-1])
    ax2.set_yticks(np.arange(0, 1.01, 0.2))
    ax2.set_title("(b) Cost-effectiveness acceptability curves", loc="left", fontweight="bold")
    ax2.grid(True, ls=":", color="#dddddd"); ax2.legend(loc="lower right", frameon=True, framealpha=1, edgecolor="#dddddd", fontsize=9)
    fig.tight_layout(); fig.savefig(OUT_DIR / fname); plt.close(fig)
    print(f"✓  Saved:  {OUT_DIR / fname}")


def figure_multiway(multi, frontier, evpi, fname="figure_psa_ceac_multiway.png"):
    plt = _mpl()
    fig, ax = plt.subplots(figsize=(7.5, 4.6), dpi=150)
    for s in STRATEGIES:
        ax.plot(WTP_GRID, multi[s], color=COLORS[s], ls=LINESTYLES[s], lw=1.8, label=s)
    # frontier: probability of the strategy with the highest expected NMB
    front = np.array([multi[STRATEGIES[k]][j] for j, k in enumerate(frontier)])
    ax.plot(WTP_GRID[::6], front[::6], ls="", marker="o", ms=3.2, mfc="white", mec="black", mew=0.9, label="Acceptability frontier")
    ax.set_xlabel("Willingness to pay ($ per averted decay)"); ax.set_ylabel("Probability of highest net monetary benefit")
    ax.set_ylim(0, 1.02); ax.set_xlim(0, WTP_GRID[-1]); ax.grid(True, ls=":", color="#dddddd")
    ax.legend(loc="center right", frameon=False, fontsize=9)
    ax.set_title("Multi-way CEAC (four mutually exclusive strategies)", loc="left", fontweight="bold")
    fig.tight_layout(); fig.savefig(OUT_DIR / fname); plt.close(fig)
    print(f"✓  Saved:  {OUT_DIR / fname}")


def figure_implementation(impl_curves, impl_rows, fname="figure_psa_implementation.png"):
    plt = _mpl()
    import matplotlib.cm as cm
    fig, ax = plt.subplots(figsize=(7.5, 4.6), dpi=150)
    shades = cm.Blues(np.linspace(0.45, 0.95, len(IMPL_SCREENING_COSTS)))
    for (cs, curve), col, r in zip(impl_curves.items(), shades, impl_rows):
        ax.plot(WTP_GRID, curve, color=col, lw=2, label=f"\\${cs:g} per consenting adolescent  (ICER \\${r['ICER_mean_ratio']:.0f})")
    ax.axvline(FILLING_COST_TRI[1], color="#888888", ls=":", lw=1.2)
    ax.text(FILLING_COST_TRI[1] + 2, 0.60, "Cost of restoration\nper tooth (Table 1)", fontsize=8, color="#555555")
    ax.set_xlabel("Willingness to pay (\\$ per averted decay)"); ax.set_ylabel("Probability cost-effective vs no program")
    ax.set_ylim(0, 1.02); ax.set_xlim(0, WTP_GRID[-1]); ax.grid(True, ls=":", color="#dddddd")
    ax.legend(loc="lower right", frameon=True, framealpha=1, edgecolor="#dddddd", fontsize=9, title="Additional program cost (illustrative values)", title_fontsize=8.5)
    ax.set_title(f"Implementation scenario: {IMPL_CASE} delivered by an SBSP\n(consent rate ~ Beta, mean {IMPL_PARTICIPATION_MEAN:.0%}; additional program cost per consenting adolescent)",
                 loc="left", fontweight="bold", fontsize=10)
    fig.tight_layout(); fig.savefig(OUT_DIR / fname); plt.close(fig)
    print(f"✓  Saved:  {OUT_DIR / fname}")


# =====================================================================
# MAIN
# =====================================================================
def main():
    import time
    t0 = time.time()
    print(f"PSA: {N_PSA:,} parameter draws × {len(SCENARIOS)} scenarios (expected-value model), seed {PSA_SEED}")
    P, S, G = run_psa()
    print(f"  model evaluations done in {time.time()-t0:.1f}s")
    pair = ceac_pairwise(S, WTP_GRID)
    multi, frontier, evpi = ceac_multiway(S, WTP_GRID)
    incr = incremental_analysis(S)
    impl_rows, impl_curves = implementation_scenario(P, S, WTP_GRID)
    xlsx = write_outputs(P, S, G, pair, multi, frontier, evpi, incr, impl_rows, impl_curves)
    print(f"✓  Saved:  {xlsx}\n✓  Saved:  {OUT_DIR / 'psa_draws.csv'}")

    # console summary
    print("\nStrategy summary (per 1,000 12-year-olds; mean [95% credible interval]):")
    for s in STRATEGIES[1:]:
        e = summarize(S[s]["averted_decay"]); c = summarize(S[s]["net_cost"]); icer = S[s]["net_cost"] / S[s]["averted_decay"]
        print(f"  {s}: averted decay {e[0]:.1f} [{e[2]:.1f}, {e[3]:.1f}] | net cost ${c[0]:,.0f} [${c[2]:,.0f}, ${c[3]:,.0f}] | "
              f"$/averted decay {S[s]['net_cost'].mean()/S[s]['averted_decay'].mean():.2f} (2.5–97.5 pct of draws: {pct(icer,.025):.2f}–{pct(icer,.975):.2f})")
    print("\nP(cost-effective vs no sealants) at WTP per averted decayed molar:")
    for lam in WTP_REPORT:
        j = int(np.argmin(np.abs(WTP_GRID - lam)))
        print(f"  λ = ${lam:>6.2f}: " + "  ".join(f"{s} {pair[s][j]:.3f}" for s in STRATEGIES[1:]))
    print("\nIncremental analysis (expected values):")
    for s, c, e, icer, comp in incr: print(f"  {s:8} net ${c:,.0f}, averted {e:.1f} → ICER ${icer:.2f} {comp}")
    j = int(np.argmin(np.abs(WTP_GRID - FILLING_COST_TRI[1])))
    print(f"\nEVPI per 1,000 adolescents at λ = ${FILLING_COST_TRI[1]:.2f}: ${evpi[j]:,.0f}; frontier there: {STRATEGIES[frontier[j]]}")
    print(f"\nImplementation scenario ({IMPL_CASE}, consent rate mean {IMPL_PARTICIPATION_MEAN:.0%}):")
    for r in impl_rows:
        a = r["averted_decay_per_1000_enrolled"]; nc = r["net_cost_per_1000_enrolled"]
        print(f"  extra cost ${r['screening_cost']:>4.0f}/adolescent: ICER ${r['ICER_mean_ratio']:.2f} [{r['ICER_2.5']:.2f}, {r['ICER_97.5']:.2f}] | "
              f"per 1,000 adolescents: averted {a[0]:.0f} [{a[2]:.0f}, {a[3]:.0f}], net ${nc[0]:,.0f} | P(CE @ ${FILLING_COST_TRI[1]:.0f}) = {r['P_CE_at_fill_fee']:.3f}")

    try:
        figure_ce_plane_ceac(S, pair); figure_multiway(multi, frontier, evpi)
        if any(c > 0 for c in IMPL_SCREENING_COSTS):
            figure_implementation(impl_curves, impl_rows)
        else:
            print("(implementation figure skipped: with no extra cost input its CEAC equals Case 1)")
    except ImportError as e:
        print(f"(figures skipped: {e})")


if __name__ == "__main__":
    main()
