"""
=============================================================================
Stochastic Markov simulation — adolescent school sealant programs
=============================================================================

Model behind Table 2, Figure 2 and Figure 3 of
"Impact and Cost-Effectiveness of Adolescent School Sealant Programs"
(Journal of Dental Research, under review).

Settings used for the manuscript
  1. N_SIM = 1,000 Monte Carlo iterations, SEED = 12345
  2. Costs discounted at 3 %/yr; for half-year period t the discount factor
     is (1 + DISCOUNT_RATE) ** (-t / 2), t = 0..14
  3. Table 2 intervals: 95 % CI of the mean across iterations, mean ± 1.96·SE
  4. Cost per averted outcome: ratio of means, 95 % CI by the delta method

Running this file regenerates Table 2 (sheet "Table 2 (paper)" of the Excel
output), Figure 2 (figure2_cumulative_decay.png) and Figure 3
(figure3_tornado_2M.png, figure3_tornado_1M.png) in ../results/, and checks
every Table 2 cell and Figure 3 value against the manuscript (values
embedded at the bottom of this file).

Scenarios
---------
  1M Base Case : 1332 pre-sealed (33.33%) + 1696 unsealed — no new sealing
  1M Case 2    : seal all 1696 sound-unsealed at baseline
  1M Case 3    : seal at baseline + reseal sealant losses for periods 1-4
  2M Base Case : 528 pre-sealed (13.2%) + 2136 unsealed + 1148 unerupted
  2M Case 2    : seal all 2136 sound-unsealed at baseline (→ 2664 total sealed)
  2M Case 3    : seal at baseline + reseal losses + seal erupted, periods 1-4

=============================================================================

"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple

import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# =====================================================================

N_SIM         = 1_000       # iterations used for the manuscript
SEED          = 12345
N_TOTAL_TEETH = 4000
N_CYCLES      = 14          # half-year cycles → periods 0–14, ages 12–19
OUTPUT_XLSX   = "Sealant_Markov_Stochastic_Results_submitted.xlsx"
OUT_DIR       = Path(__file__).resolve().parent.parent / "results"   # all outputs go here

DISCOUNT_RATE = 0.03        # annual; Methods: "a standard discount rate of 3% was used"
# Discount factor for costs incurred in half-year period t (t = 0 is model start)
DISCOUNT_FACTORS = np.array([(1.0 + DISCOUNT_RATE) ** (-t / 2.0) for t in range(N_CYCLES + 1)])

MAKE_FIGURES  = True        # Figure 2 + Figure 3 (needs matplotlib)


# =====================================================================
# MODEL PARAMETERS
# =====================================================================
AR_BASE = {"1M": 0.02895,   "2M": 0.030051}

# Retention rate: Triangular(low=93.82%, mode=94.85%, high=95.88%) — same for 1M and 2M
R_TRI = (0.9382, 0.9485, 0.9588)   # (low, mode, high)

PARAM_RANGE_PCT    = 0.20           # ±20% triangular on AR only
PARAM_DISTRIBUTION = "triangular"  

# ── Age-specific annual dental-visit probabilities (ages 12–18) ──────────────
# SDs ≈ sqrt(p·(1–p)) — natural Bernoulli SDs; no extra DV uncertainty added.
yearly_visit_probs = np.array([
    0.680999,   # age 12
    0.609101,   # age 13
    0.681724,   # age 14
    0.592226,   # age 15
    0.585417,   # age 16
    0.463168,   # age 17
    0.553096,   # age 18
])
sd_visit = np.array([
    0.466090, 0.487952, 0.465807, 0.491421,
    0.492650, 0.498642, 0.497173,
])

# Per-period DV (÷2 for 6-month transitions).
# Period t (1–14) maps to age index = (t–1)//2:
#   t=1,2 → age 12 | t=3,4 → age 13 | … | t=13,14 → age 18
DV_PER_PERIOD = np.array(
    [yearly_visit_probs[(t - 1) // 2] / 2.0 for t in range(1, N_CYCLES + 1)]
)

# ── Costs ─────────────────────────────────────────────────────────────────────
SEALANT_COST_TRI = (28.85,  33.75,  38.17)     # $/tooth Triangular(low, mode, high)
FILLING_COST_TRI = (112.14, 130.89, 146.22)    # $/tooth

# ── Eruption schedules (14 values, periods 1–14) ─────────────────────────────
ERUPTION_2M = [359, 360, 140, 140, 36, 36, 20, 57, 0, 0, 0, 0, 0, 0]
ERUPTION_1M = [0] * N_CYCLES


# =====================================================================
# SCENARIO CONFIGURATION
# =====================================================================
@dataclass(frozen=True)
class ScenarioConfig:
    sheet_name: str
    tooth_group: str              # "1M" or "2M"
    init_sound_sealed: int        # sealed teeth at period 0 (pre-existing)
    init_sound_unsealed: int      # sound unsealed at period 0
    init_unerupted: int           # unerupted at period 0 (2M only)
    eruption_schedule: List[int]  # teeth erupting per period (14 values)
    seal_at_baseline: bool        # seal ALL sound-unsealed teeth
    reseal_loss_cycles: int = 0   # reseal S→NS losses for periods 1..n
    seal_eruption_cycles: int = 0 # seal newly erupted teeth for periods 1..n
    notes: str = ""


SCENARIOS: Dict[str, ScenarioConfig] = {
    # ── 1M ──────────────────────────────────────────────────────────────────
    "1M Base Case": ScenarioConfig(
        sheet_name="1M Base Case",
        tooth_group="1M",
        init_sound_sealed=1332,    # 33.33% × 4000 pre-existing sealants
        init_sound_unsealed=1696,
        init_unerupted=0,
        eruption_schedule=ERUPTION_1M,
        seal_at_baseline=False,
        notes=(
            "No new sealants. 1332 pre-existing sealed (33.33%) + 1696 sound-unsealed "
            "follow natural Markov history. Sealed teeth lose sealants at (1-R)/period → "
            "become unsealed → decay at AR."
        ),
    ),
    "1M Case 2": ScenarioConfig(
        sheet_name="1M Case 2",
        tooth_group="1M",
        init_sound_sealed=1332,
        init_sound_unsealed=1696,
        init_unerupted=0,
        eruption_schedule=ERUPTION_1M,
        seal_at_baseline=True,         # seals 1696 → total sealed at t=0 = 3028
        reseal_loss_cycles=0,
        seal_eruption_cycles=0,
        notes="Seal all 1696 sound-unsealed 1M teeth at baseline. No resealing.",
    ),
    "1M Case 3": ScenarioConfig(
        sheet_name="1M Case 3",
        tooth_group="1M",
        init_sound_sealed=1332,
        init_sound_unsealed=1696,
        init_unerupted=0,
        eruption_schedule=ERUPTION_1M,
        seal_at_baseline=True,
        reseal_loss_cycles=4,          # reseal S→NS losses in periods 1–4
        seal_eruption_cycles=0,
        notes=(
            "Seal all 1696 sound-unsealed at baseline + reseal sealant losses "
            "in periods 1–4. After period 4, losses go to NS."
        ),
    ),
    # ── 2M ──────────────────────────────────────────────────────────────────
    "2M Base Case": ScenarioConfig(
        sheet_name="2M Base Case",
        tooth_group="2M",
        init_sound_sealed=528,     # 13.2% × 4000 
        init_sound_unsealed=2136,
        init_unerupted=1148,
        eruption_schedule=ERUPTION_2M,
        seal_at_baseline=False,
        notes=(
            "No new sealants. 528 pre-existing sealed (13.2%) + 2136 sound-unsealed "
            "+ 1148 unerupted follow natural Markov history."
        ),
    ),
    "2M Case 2": ScenarioConfig(
        sheet_name="2M Case 2",
        tooth_group="2M",
        init_sound_sealed=528,     # pre-existing (not sealed by program)
        init_sound_unsealed=2136,
        init_unerupted=1148,
        eruption_schedule=ERUPTION_2M,
        seal_at_baseline=True,    # seals 2136 → total sealed = 528+2136 = 2664
        reseal_loss_cycles=0,
        seal_eruption_cycles=0,
        notes=(
            "Seal all 2136 sound-unsealed 2M teeth at baseline (total sealed = 2664). "
            "Erupted teeth enter NS naturally."
        ),
    ),
    "2M Case 3": ScenarioConfig(
        sheet_name="2M Case 3",
        tooth_group="2M",
        init_sound_sealed=528,
        init_sound_unsealed=2136,
        init_unerupted=1148,
        eruption_schedule=ERUPTION_2M,
        seal_at_baseline=True,
        reseal_loss_cycles=4,
        seal_eruption_cycles=4,    # seal erupted teeth during periods 1–4
        notes=(
            "Seal all 2136 sound-unsealed at baseline + reseal sealant losses + "
            "seal newly erupted teeth, all during periods 1–4."
        ),
    ),
}

# Overall case column mapping for Results sheet
OVERALL_CASES: Dict[str, Dict[str, str]] = {
    "No sealants": {"1M": "1M Base Case", "2M": "2M Base Case"},
    "Case 1":      {"1M": "1M Base Case", "2M": "2M Case 2"},
    "Case 2":      {"1M": "1M Case 2",    "2M": "2M Case 2"},
    "Case 3":      {"1M": "1M Case 3",    "2M": "2M Case 3"},
}


# =====================================================================
# METRICS TRACKED PER PERIOD
# =====================================================================
METRICS = [
    "sound_sealed",
    "sound_unsealed",
    "untreated_decay_end",
    "unerupted",
    "erupted_this_period",
    "cumulative_erupted",
    "new_decay_from_sealed",
    "new_decay_from_unsealed",
    "new_decay_total",
    "new_fillings",
    "resealed_losses",
    "sealed_newly_erupted",
    "sealants_this_period",
    "cumulative_sealants",
    "cumulative_new_decay",
    "cumulative_new_fillings",
]

DETAIL_COL_ORDER = [
    ("sound_sealed",            "Sound Sealed"),
    ("sound_unsealed",          "Sound Unsealed"),
    ("untreated_decay_end",     "Untreated Decay End"),
    ("unerupted",               "Unerupted"),
    ("erupted_this_period",     "Erupted This Period"),
    ("cumulative_erupted",      "Cumulative Erupted"),
    ("new_decay_from_sealed",   "New Decay from Sealed"),
    ("new_decay_from_unsealed", "New Decay from Unsealed"),
    ("new_decay_total",         "New Decay Total"),
    ("new_fillings",            "New Fillings"),
    ("resealed_losses",         "Resealed Losses"),
    ("sealed_newly_erupted",    "Newly Erupted Sealed"),
    ("sealants_this_period",    "Sealants This Period"),
    ("cumulative_sealants",     "Cumulative Sealants"),
    ("cumulative_new_decay",    "Cumulative New Decay"),
    ("cumulative_new_fillings", "Cumulative New Fillings"),
]


# =====================================================================

def sample_range_parameter(
    rng: np.random.Generator, baseline: float, pct: float
) -> float:
    """Sample from Triangular(baseline*(1-pct), baseline, baseline*(1+pct))."""
    low  = max(0.0,      baseline * (1.0 - pct))
    high = min(0.999999, baseline * (1.0 + pct))
    mode = min(max(baseline, low), high)
    if np.isclose(low, high):
        return float(mode)
    if PARAM_DISTRIBUTION.lower() == "uniform":
        return float(rng.uniform(low, high))
    return float(rng.triangular(left=low, mode=mode, right=high))


def summary_stats(values: np.ndarray) -> Tuple[float, float, float]:
    v = np.asarray(values, dtype=float)
    return float(np.mean(v)), float(np.quantile(v, 0.025)), float(np.quantile(v, 0.975))


def safe_ratio(num: np.ndarray, den: np.ndarray) -> np.ndarray:
    """Element-wise division; iterations where denominator <= 0 -> NaN (excluded from stats)."""
    with np.errstate(invalid="ignore", divide="ignore"):
        return np.where(den > 0, num / den, np.nan)


def summary_stats_nan(values: np.ndarray) -> Tuple[float, float, float]:
    """Mean and PSA-style 95% uncertainty interval, ignoring NaN entries."""
    v = np.asarray(values, dtype=float)
    return (float(np.nanmean(v)),
            float(np.nanquantile(v, 0.025)),
            float(np.nanquantile(v, 0.975)))


def summary_stats_ci(values: np.ndarray) -> Tuple[float, float, float]:
    """Mean and conventional 95% CI for the mean: mean ± 1.96 × SE."""
    v = np.asarray(values, dtype=float)
    n = v.size
    mean = float(np.mean(v))
    if n <= 1:
        return mean, mean, mean
    se = float(np.std(v, ddof=1) / np.sqrt(n))
    half_width = 1.96 * se
    return mean, mean - half_width, mean + half_width


def summary_stats_nan_ci(values: np.ndarray) -> Tuple[float, float, float]:
    """Mean and conventional 95% CI for the mean, ignoring NaN entries."""
    v = np.asarray(values, dtype=float)
    v = v[~np.isnan(v)]
    n = v.size
    mean = float(np.mean(v))
    if n <= 1:
        return mean, mean, mean
    se = float(np.std(v, ddof=1) / np.sqrt(n))
    half_width = 1.96 * se
    return mean, mean - half_width, mean + half_width


def ratio_of_means_ci(num: np.ndarray, den: np.ndarray) -> Tuple[float, float, float]:
    """
    Cost-effectiveness ratio as used in the manuscript:
    point estimate = mean(num) / mean(den); 95% CI by the delta method,
        SE(R) = sqrt( (Var(num) - 2·R·Cov(num,den) + R²·Var(den)) / n ) / mean(den).
    """
    x = np.asarray(num, dtype=float)
    y = np.asarray(den, dtype=float)
    n = x.size
    r = float(np.mean(x) / np.mean(y))
    cov = np.cov(x, y, ddof=1)
    var_r = (cov[0, 0] - 2.0 * r * cov[0, 1] + r * r * cov[1, 1]) / n
    half_width = 1.96 * float(np.sqrt(max(var_r, 0.0))) / float(np.mean(y))
    return r, r - half_width, r + half_width


def fmt_ci(m: float, lo: float, hi: float, d: int = 1) -> str:
    return f"({lo:,.{d}f}, {hi:,.{d}f})"


def fmt_val(m: float, denom: int | None = None, d: int = 1) -> str:
    if denom is None:
        return f"{m:,.{d}f}"
    return f"{m:,.{d}f}  ({100.0*m/denom:.2f}%)"


# ── Manuscript-style cell formats ("mean [lo, hi]") ─────────────────────────
def paper_count(s: Tuple[float, float, float], denom: int | None = None, d: int = 1) -> str:
    m, lo, hi = s
    pct = f" ({100.0*m/denom:.2f}%)" if denom else ""
    return f"{m:,.{d}f}{pct} [{lo:,.{d}f}, {hi:,.{d}f}]"


def paper_cost(s: Tuple[float, float, float], d: int = 0) -> str:
    m, lo, hi = s
    return f"${m:,.{d}f} [${lo:,.{d}f}, ${hi:,.{d}f}]"


# =====================================================================
# CORE SIMULATION
# =====================================================================
def run_one_scenario(
    cfg: ScenarioConfig,
    r_value: float,
    ar_value: float,
    dv_per_period: np.ndarray,
    rng: np.random.Generator,
) -> Dict[str, np.ndarray]:
    """
    One stochastic Markov cohort path for a single scenario.

    """
    path = {m: np.zeros(N_CYCLES + 1, dtype=np.int32) for m in METRICS}

    sound_sealed    = int(cfg.init_sound_sealed)
    sound_unsealed  = int(cfg.init_sound_unsealed)
    untreated_decay = 0
    unerupted       = int(cfg.init_unerupted)

    cum_erupted  = 0
    cum_sealants = 0
    cum_decay    = 0
    cum_fillings = 0

    # Period 0 — optional baseline sealing then record
    if cfg.seal_at_baseline:
        n_seal        = sound_unsealed
        sound_sealed += n_seal
        sound_unsealed = 0
        cum_sealants  += n_seal
        path["sealants_this_period"][0] = n_seal

    path["sound_sealed"][0]           = sound_sealed
    path["sound_unsealed"][0]         = sound_unsealed
    path["untreated_decay_end"][0]    = untreated_decay
    path["unerupted"][0]              = unerupted
    path["cumulative_erupted"][0]     = cum_erupted
    path["cumulative_sealants"][0]    = cum_sealants
    path["cumulative_new_decay"][0]   = cum_decay
    path["cumulative_new_fillings"][0] = cum_fillings

    p_loss     = max(0.0, min(1.0, 1.0 - r_value))
    p_s_to_c   = max(0.0, min(1.0, ar_value / 2.0))   # prob decay | lost sealant
    p_ns_to_c  = max(0.0, min(1.0, ar_value))

    for t in range(1, N_CYCLES + 1):
        fill_prob = float(np.clip(dv_per_period[t - 1], 0.0, 1.0))

        # Step 1 — Eruptions: count & remove from unerupted;
        erupted      = int(cfg.eruption_schedule[t - 1])
        unerupted   -= erupted
        cum_erupted += erupted

        # Step 2 — S → {stay S, NS-buffer, C}
        losses     = int(rng.binomial(sound_sealed, p_loss))   if sound_sealed > 0 else 0
        s_to_c     = int(rng.binomial(losses, p_s_to_c))       if losses > 0        else 0
        s_to_ns    = losses - s_to_c                           # would become NS
        s_retained = sound_sealed - losses

        # Step 3 — NS → C  using NS_START (before any new arrivals)
        ns_to_c  = int(rng.binomial(sound_unsealed, p_ns_to_c)) if sound_unsealed > 0 else 0
        ns_kept  = sound_unsealed - ns_to_c

        # Intervention
        reseal_s2ns  = s_to_ns if t <= cfg.reseal_loss_cycles   else 0
        unseal_s2ns  = 0       if t <= cfg.reseal_loss_cycles   else s_to_ns
        seal_erupt   = erupted if t <= cfg.seal_eruption_cycles else 0
        unseal_erupt = 0       if t <= cfg.seal_eruption_cycles else erupted
        seals_t      = reseal_s2ns + seal_erupt

        # Step 4 — C → F  on (C_old + new_decay_total)
        new_decay         = s_to_c + ns_to_c
        caries_pool       = untreated_decay + new_decay
        filled            = int(rng.binomial(caries_pool, fill_prob)) if caries_pool > 0 else 0
        untreated_decay   = caries_pool - filled

        # Step 5 — Update state counts
        sound_sealed   = s_retained + reseal_s2ns + seal_erupt
        sound_unsealed = ns_kept + unseal_s2ns + unseal_erupt

        cum_sealants += seals_t
        cum_decay    += new_decay
        cum_fillings += filled

        path["sound_sealed"][t]            = sound_sealed
        path["sound_unsealed"][t]          = sound_unsealed
        path["untreated_decay_end"][t]     = untreated_decay
        path["unerupted"][t]               = unerupted
        path["erupted_this_period"][t]     = erupted
        path["cumulative_erupted"][t]      = cum_erupted
        path["new_decay_from_sealed"][t]   = s_to_c
        path["new_decay_from_unsealed"][t] = ns_to_c
        path["new_decay_total"][t]         = new_decay
        path["new_fillings"][t]            = filled
        path["resealed_losses"][t]         = reseal_s2ns
        path["sealed_newly_erupted"][t]    = seal_erupt
        path["sealants_this_period"][t]    = seals_t
        path["cumulative_sealants"][t]     = cum_sealants
        path["cumulative_new_decay"][t]    = cum_decay
        path["cumulative_new_fillings"][t] = cum_fillings

    return path


# =====================================================================
# EXCEL
# =====================================================================
def auto_fit(ws) -> None:
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        mx = max((len(str(c.value)) for c in col if c.value is not None), default=10)
        ws.column_dimensions[col_letter].width = min(max(mx + 2, 11), 38)


def hdr_row(ws, row: int, color: str = "1F4E78") -> None:
    fill = PatternFill("solid", fgColor=color)
    for cell in ws[row]:
        cell.fill      = fill
        cell.font      = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = Border(bottom=Side(style="thin", color="FFFFFF"))


def sub_row(ws, row: int) -> None:
    for cell in ws[row]:
        cell.fill      = PatternFill("solid", fgColor="D9EAF7")
        cell.font      = Font(bold=True)
        cell.alignment = Alignment(vertical="center", wrap_text=True)


# =====================================================================
# SHEET WRITERS
# =====================================================================
def write_inputs_sheet(wb: Workbook) -> None:
    ws = wb.create_sheet("Inputs")
    ws.append(["Parameter", "Value", "Notes"])
    hdr_row(ws, 1)

    for row in [
        ["N_SIM",            N_SIM,              "Monte Carlo iterations"],
        ["SEED",             SEED,               "Random seed"],
        ["N_TOTAL_TEETH",    N_TOTAL_TEETH,      "Denominator for % outputs"],
        ["N_CYCLES",         N_CYCLES,           "14 half-year cycles (ages 12–19)"],
        ["Interval method",   "Mean ± 1.96×SE",   "All outcomes: conventional 95% CI for the mean across iterations; cost-effectiveness = ratio of means with delta-method CI"],
        ["DISCOUNT_RATE",     DISCOUNT_RATE,      "Annual rate; costs in period t discounted by (1+r)^(-t/2)"],
        ["R_TRI_low",        R_TRI[0],           "Retention rate triangular lower bound"],
        ["R_TRI_mode",       R_TRI[1],           "Retention rate triangular mode (base)"],
        ["R_TRI_high",       R_TRI[2],           "Retention rate triangular upper bound"],
        ["AR_BASE_1M",       AR_BASE["1M"],      "±20% triangular uncertainty applied"],
        ["AR_BASE_2M",       AR_BASE["2M"],      "±20% triangular uncertainty applied"],
        ["PARAM_RANGE_PCT",  PARAM_RANGE_PCT,    "Fractional uncertainty on AR only"],
        ["Sealant Cost Low", SEALANT_COST_TRI[0],"$/tooth"],
        ["Sealant Cost Mode",SEALANT_COST_TRI[1],"$/tooth"],
        ["Sealant Cost High",SEALANT_COST_TRI[2],"$/tooth"],
        ["Filling Cost Low", FILLING_COST_TRI[0],"$/tooth"],
        ["Filling Cost Mode",FILLING_COST_TRI[1],"$/tooth"],
        ["Filling Cost High",FILLING_COST_TRI[2],"$/tooth"],
        ["2M eruption schedule", ", ".join(map(str, ERUPTION_2M)), "Periods 1–14"],
    ]:
        ws.append(row)

    ws.append([])
    ws.append(["Age-Specific Dental Visit Probabilities (annual → ÷2 per 6-month period)"])
    ws.append(["Age", "Annual Prob", "Per-Period DV (annual÷2)", "SD (Bernoulli)"])
    hdr_row(ws, ws.max_row)
    for i, (yp, sd) in enumerate(zip(yearly_visit_probs, sd_visit)):
        ws.append([12 + i, round(float(yp), 6), round(float(yp/2), 6), round(float(sd), 6)])

    ws.append([])
    ws.append(["DV per period (t=1..14):"] + [round(float(v), 6) for v in DV_PER_PERIOD])
    ws.append(["Period-to-age mapping: age = 12 + (t-1)//2  (t=1,2→age 12 ; t=3,4→age 13 …)"])

    ws.append([])
    ws.append(["Scenario", "Group", "Init Sealed", "Init Unsealed", "Init Unerupted",
               "Seal Baseline", "Reseal Cycles", "Erupt Seal Cycles", "Notes"])
    hdr_row(ws, ws.max_row)
    for k, cfg in SCENARIOS.items():
        ws.append([k, cfg.tooth_group, cfg.init_sound_sealed, cfg.init_sound_unsealed,
                   cfg.init_unerupted, cfg.seal_at_baseline,
                   cfg.reseal_loss_cycles, cfg.seal_eruption_cycles, cfg.notes])

    auto_fit(ws)
    ws.freeze_panes = "A2"


def build_detail_sheet(
    wb: Workbook,
    sheet_name: str,
    sc_paths: Dict[str, np.ndarray],
    cfg: ScenarioConfig,
) -> None:
    ws = wb.create_sheet(sheet_name)
    ws.append([cfg.sheet_name])
    ws.append([cfg.notes])

    header = ["Period"]
    for _, label in DETAIL_COL_ORDER:
        header.append(f"{label} Mean")
    ws.append(header)
    hdr_row(ws, 3)

    for t in range(N_CYCLES + 1):
        age_lbl = f"P{t} Age {12+t//2}{'½' if t%2 else ''}"
        row = [age_lbl]
        for metric, _ in DETAIL_COL_ORDER:
            m, _, _ = summary_stats(sc_paths[metric][:, t])
            row.append(int(round(m)))
        ws.append(row)

    ws["A1"].font = Font(bold=True, size=12)
    ws["A2"].alignment = Alignment(wrap_text=True)
    ws.row_dimensions[2].height = 36
    auto_fit(ws)
    ws.freeze_panes = "A4"


def write_results_sheet(
    wb: Workbook,
    sc_finals: Dict[str, Dict[str, np.ndarray]],
) -> None:
    ws = wb.create_sheet("Results")
    ws.append([None,
               "No sealants", "(95% interval)",
               "Case 1",      "(95% interval)",
               "Case 2",      "(95% interval)",
               "Case 3",      "(95% interval)"])
    hdr_row(ws, 1)
    cases = ["No sealants", "Case 1", "Case 2", "Case 3"]

    def sec(title):
        ws.append([title])
        sub_row(ws, ws.max_row)

    def metric_row(label, stats_map, denom=None, d=1):
        row = [label]
        for cn in cases:
            s = stats_map.get(cn)
            row += (["N/A", ""] if s is None
                    else [fmt_val(s[0], denom, d), fmt_ci(*s, d)])
        ws.append(row)

    def cost_row(label, stats_map, d=0):
        row = [label]
        for cn in cases:
            s = stats_map.get(cn)
            row += (["N/A", ""] if s is None
                    else [f"${s[0]:,.{d}f}", fmt_ci(*s, d)])
        ws.append(row)

    def ce_row(label, stats_map, d=2):
        """Cost-effectiveness row: $/averted outcome = ratio of means; interval = delta-method 95% CI."""
        row = [label]
        for cn in cases:
            s = stats_map.get(cn)
            row += (["N/A", ""] if s is None
                    else [f"${s[0]:,.{d}f}", fmt_ci(*s, d)])
        ws.append(row)

    # 1M
    sec("1M")
    metric_row("# Sealants applied", {
        "No sealants": summary_stats_ci(np.zeros(N_SIM)),
        "Case 1":      summary_stats_ci(np.zeros(N_SIM)),
        "Case 2":      summary_stats_ci(sc_finals["1M Case 2"]["sealants"]),
        "Case 3":      summary_stats_ci(sc_finals["1M Case 3"]["sealants"]),
    })
    metric_row("# Cumulative New Decay", {
        "No sealants": summary_stats_ci(sc_finals["1M Base Case"]["decay"]),
        "Case 1":      summary_stats_ci(sc_finals["1M Base Case"]["decay"]),
        "Case 2":      summary_stats_ci(sc_finals["1M Case 2"]["decay"]),
        "Case 3":      summary_stats_ci(sc_finals["1M Case 3"]["decay"]),
    }, denom=N_TOTAL_TEETH)
    metric_row("# Cumulative New Fillings", {
        "No sealants": summary_stats_ci(sc_finals["1M Base Case"]["fillings"]),
        "Case 1":      summary_stats_ci(sc_finals["1M Base Case"]["fillings"]),
        "Case 2":      summary_stats_ci(sc_finals["1M Case 2"]["fillings"]),
        "Case 3":      summary_stats_ci(sc_finals["1M Case 3"]["fillings"]),
    }, denom=N_TOTAL_TEETH)
    metric_row("Averted decay", {
        "No sealants": None, "Case 1": None,
        "Case 2": summary_stats_ci(sc_finals["1M Base Case"]["decay"] - sc_finals["1M Case 2"]["decay"]),
        "Case 3": summary_stats_ci(sc_finals["1M Base Case"]["decay"] - sc_finals["1M Case 3"]["decay"]),
    })
    metric_row("Averted fillings", {
        "No sealants": None, "Case 1": None,
        "Case 2": summary_stats_ci(sc_finals["1M Base Case"]["fillings"] - sc_finals["1M Case 2"]["fillings"]),
        "Case 3": summary_stats_ci(sc_finals["1M Base Case"]["fillings"] - sc_finals["1M Case 3"]["fillings"]),
    })
    cost_row("Sealant cost", {
        "No sealants": summary_stats(np.zeros(N_SIM)),
        "Case 1":      summary_stats(np.zeros(N_SIM)),
        "Case 2":      summary_stats_ci(sc_finals["1M Case 2"]["sealant_cost"]),
        "Case 3":      summary_stats_ci(sc_finals["1M Case 3"]["sealant_cost"]),
    })
    cost_row("Filling cost", {
        "No sealants": summary_stats_ci(sc_finals["1M Base Case"]["filling_cost"]),
        "Case 1":      summary_stats_ci(sc_finals["1M Base Case"]["filling_cost"]),
        "Case 2":      summary_stats_ci(sc_finals["1M Case 2"]["filling_cost"]),
        "Case 3":      summary_stats_ci(sc_finals["1M Case 3"]["filling_cost"]),
    })
    cost_row("Cost of program (net)", {
        "No sealants": summary_stats(np.zeros(N_SIM)),
        "Case 1":      summary_stats(np.zeros(N_SIM)),
        "Case 2":      summary_stats_ci(sc_finals["1M Case 2"]["net_program_cost"]),
        "Case 3":      summary_stats_ci(sc_finals["1M Case 3"]["net_program_cost"]),
    })
    # ── 1M Cost-Effectiveness (computed per-simulation iteration) ──────────
    _1m_avd_c2 = sc_finals["1M Base Case"]["decay"]    - sc_finals["1M Case 2"]["decay"]
    _1m_avd_c3 = sc_finals["1M Base Case"]["decay"]    - sc_finals["1M Case 3"]["decay"]
    _1m_avf_c2 = sc_finals["1M Base Case"]["fillings"] - sc_finals["1M Case 2"]["fillings"]
    _1m_avf_c3 = sc_finals["1M Base Case"]["fillings"] - sc_finals["1M Case 3"]["fillings"]
    ce_row("Cost per averted decay ($/decay)", {
        "No sealants": None, "Case 1": None,
        "Case 2": ratio_of_means_ci(sc_finals["1M Case 2"]["net_program_cost"], _1m_avd_c2),
        "Case 3": ratio_of_means_ci(sc_finals["1M Case 3"]["net_program_cost"], _1m_avd_c3),
    })
    ce_row("Cost per averted filling ($/filling)", {
        "No sealants": None, "Case 1": None,
        "Case 2": ratio_of_means_ci(sc_finals["1M Case 2"]["net_program_cost"], _1m_avf_c2),
        "Case 3": ratio_of_means_ci(sc_finals["1M Case 3"]["net_program_cost"], _1m_avf_c3),
    })
    ws.append([])

    # 2M
    sec("2M")
    metric_row("# Sealants applied", {
        "No sealants": summary_stats_ci(np.zeros(N_SIM)),
        "Case 1":      summary_stats_ci(sc_finals["2M Case 2"]["sealants"]),
        "Case 2":      summary_stats_ci(sc_finals["2M Case 2"]["sealants"]),
        "Case 3":      summary_stats_ci(sc_finals["2M Case 3"]["sealants"]),
    })
    metric_row("# Cumulative New Decay", {
        "No sealants": summary_stats_ci(sc_finals["2M Base Case"]["decay"]),
        "Case 1":      summary_stats_ci(sc_finals["2M Case 2"]["decay"]),
        "Case 2":      summary_stats_ci(sc_finals["2M Case 2"]["decay"]),
        "Case 3":      summary_stats_ci(sc_finals["2M Case 3"]["decay"]),
    }, denom=N_TOTAL_TEETH)
    metric_row("# Cumulative New Fillings", {
        "No sealants": summary_stats_ci(sc_finals["2M Base Case"]["fillings"]),
        "Case 1":      summary_stats_ci(sc_finals["2M Case 2"]["fillings"]),
        "Case 2":      summary_stats_ci(sc_finals["2M Case 2"]["fillings"]),
        "Case 3":      summary_stats_ci(sc_finals["2M Case 3"]["fillings"]),
    }, denom=N_TOTAL_TEETH)
    metric_row("Averted decay", {
        "No sealants": None,
        "Case 1": summary_stats_ci(sc_finals["2M Base Case"]["decay"] - sc_finals["2M Case 2"]["decay"]),
        "Case 2": summary_stats_ci(sc_finals["2M Base Case"]["decay"] - sc_finals["2M Case 2"]["decay"]),
        "Case 3": summary_stats_ci(sc_finals["2M Base Case"]["decay"] - sc_finals["2M Case 3"]["decay"]),
    })
    metric_row("Averted fillings", {
        "No sealants": None,
        "Case 1": summary_stats_ci(sc_finals["2M Base Case"]["fillings"] - sc_finals["2M Case 2"]["fillings"]),
        "Case 2": summary_stats_ci(sc_finals["2M Base Case"]["fillings"] - sc_finals["2M Case 2"]["fillings"]),
        "Case 3": summary_stats_ci(sc_finals["2M Base Case"]["fillings"] - sc_finals["2M Case 3"]["fillings"]),
    })
    cost_row("Sealant cost", {
        "No sealants": summary_stats(np.zeros(N_SIM)),
        "Case 1":      summary_stats_ci(sc_finals["2M Case 2"]["sealant_cost"]),
        "Case 2":      summary_stats_ci(sc_finals["2M Case 2"]["sealant_cost"]),
        "Case 3":      summary_stats_ci(sc_finals["2M Case 3"]["sealant_cost"]),
    })
    cost_row("Filling cost", {
        "No sealants": summary_stats_ci(sc_finals["2M Base Case"]["filling_cost"]),
        "Case 1":      summary_stats_ci(sc_finals["2M Case 2"]["filling_cost"]),
        "Case 2":      summary_stats_ci(sc_finals["2M Case 2"]["filling_cost"]),
        "Case 3":      summary_stats_ci(sc_finals["2M Case 3"]["filling_cost"]),
    })
    cost_row("Cost of program (net)", {
        "No sealants": summary_stats(np.zeros(N_SIM)),
        "Case 1":      summary_stats_ci(sc_finals["2M Case 2"]["net_program_cost"]),
        "Case 2":      summary_stats_ci(sc_finals["2M Case 2"]["net_program_cost"]),
        "Case 3":      summary_stats_ci(sc_finals["2M Case 3"]["net_program_cost"]),
    })
    # ── 2M Cost-Effectiveness (computed per-simulation iteration) ──────────
    _2m_avd_c1 = sc_finals["2M Base Case"]["decay"]    - sc_finals["2M Case 2"]["decay"]
    _2m_avd_c3 = sc_finals["2M Base Case"]["decay"]    - sc_finals["2M Case 3"]["decay"]
    _2m_avf_c1 = sc_finals["2M Base Case"]["fillings"] - sc_finals["2M Case 2"]["fillings"]
    _2m_avf_c3 = sc_finals["2M Base Case"]["fillings"] - sc_finals["2M Case 3"]["fillings"]
    ce_row("Cost per averted decay ($/decay)", {
        "No sealants": None,
        "Case 1": ratio_of_means_ci(sc_finals["2M Case 2"]["net_program_cost"], _2m_avd_c1),
        "Case 2": ratio_of_means_ci(sc_finals["2M Case 2"]["net_program_cost"], _2m_avd_c1),
        "Case 3": ratio_of_means_ci(sc_finals["2M Case 3"]["net_program_cost"], _2m_avd_c3),
    })
    ce_row("Cost per averted filling ($/filling)", {
        "No sealants": None,
        "Case 1": ratio_of_means_ci(sc_finals["2M Case 2"]["net_program_cost"], _2m_avf_c1),
        "Case 2": ratio_of_means_ci(sc_finals["2M Case 2"]["net_program_cost"], _2m_avf_c1),
        "Case 3": ratio_of_means_ci(sc_finals["2M Case 3"]["net_program_cost"], _2m_avf_c3),
    })
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, 1).value
        ws.cell(r, 1).font = Font(bold=(v in {"1M", "2M"}))
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=2, max_col=9):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)
    auto_fit(ws)
    ws.freeze_panes = "A2"


def table2_cells(sc_finals: Dict[str, Dict[str, np.ndarray]]) -> Dict[Tuple[str, str], List[str]]:
    """
    Table 2 of the manuscript as strings, keyed by (tooth group, row label), columns in the order
    No sealants | Case 1 | Case 2 | Case 3.  Column mapping follows OVERALL_CASES:
    Case 1 seals 2M only (1M shown as N/A, as in the manuscript), Case 2 adds 1M, Case 3 adds resealing.
    """
    def block(tg: str, base: str, c2: str, c3: str, case1_is_base: bool) -> Dict[str, List[str]]:
        F = sc_finals
        zeros = summary_stats_ci(np.zeros(N_SIM))
        avd = lambda sn: F[base]["decay"] - F[sn]["decay"]
        avf = lambda sn: F[base]["fillings"] - F[sn]["fillings"]
        # Case 1 column: 2M → identical to Case 2 (2M sealed in both); 1M → N/A (not part of Case 1)
        c1 = (lambda f: "N/A") if case1_is_base else (lambda f: f(c2))
        return {
            "# Sealants applied":       [paper_count(zeros), c1(lambda sn: paper_count(summary_stats_ci(F[sn]["sealants"]))),
                                         paper_count(summary_stats_ci(F[c2]["sealants"])), paper_count(summary_stats_ci(F[c3]["sealants"]))],
            "# Cumulative New Decay":   [paper_count(summary_stats_ci(F[base]["decay"]), N_TOTAL_TEETH),
                                         c1(lambda sn: paper_count(summary_stats_ci(F[sn]["decay"]), N_TOTAL_TEETH)),
                                         paper_count(summary_stats_ci(F[c2]["decay"]), N_TOTAL_TEETH), paper_count(summary_stats_ci(F[c3]["decay"]), N_TOTAL_TEETH)],
            "# Cumulative New Fillings":[paper_count(summary_stats_ci(F[base]["fillings"]), N_TOTAL_TEETH),
                                         c1(lambda sn: paper_count(summary_stats_ci(F[sn]["fillings"]), N_TOTAL_TEETH)),
                                         paper_count(summary_stats_ci(F[c2]["fillings"]), N_TOTAL_TEETH), paper_count(summary_stats_ci(F[c3]["fillings"]), N_TOTAL_TEETH)],
            "Averted decay":            ["N/A", c1(lambda sn: paper_count(summary_stats_ci(avd(sn)))),
                                         paper_count(summary_stats_ci(avd(c2))), paper_count(summary_stats_ci(avd(c3)))],
            "Averted fillings":         ["N/A", c1(lambda sn: paper_count(summary_stats_ci(avf(sn)))),
                                         paper_count(summary_stats_ci(avf(c2))), paper_count(summary_stats_ci(avf(c3)))],
            "Sealant cost":             [paper_cost(zeros), c1(lambda sn: paper_cost(summary_stats_ci(F[sn]["sealant_cost"]))),
                                         paper_cost(summary_stats_ci(F[c2]["sealant_cost"])), paper_cost(summary_stats_ci(F[c3]["sealant_cost"]))],
            "Filling cost":             [paper_cost(summary_stats_ci(F[base]["filling_cost"])), c1(lambda sn: paper_cost(summary_stats_ci(F[sn]["filling_cost"]))),
                                         paper_cost(summary_stats_ci(F[c2]["filling_cost"])), paper_cost(summary_stats_ci(F[c3]["filling_cost"]))],
            "Cost of program (net)":    [paper_cost(zeros), c1(lambda sn: paper_cost(summary_stats_ci(F[sn]["net_program_cost"]))),
                                         paper_cost(summary_stats_ci(F[c2]["net_program_cost"])), paper_cost(summary_stats_ci(F[c3]["net_program_cost"]))],
            "Cost per averted decay ($/decay)":
                                        ["N/A", c1(lambda sn: paper_cost(ratio_of_means_ci(F[sn]["net_program_cost"], avd(sn)), 2)),
                                         paper_cost(ratio_of_means_ci(F[c2]["net_program_cost"], avd(c2)), 2), paper_cost(ratio_of_means_ci(F[c3]["net_program_cost"], avd(c3)), 2)],
            "Cost per averted filling ($/filling)":
                                        ["N/A", c1(lambda sn: paper_cost(ratio_of_means_ci(F[sn]["net_program_cost"], avf(sn)), 2)),
                                         paper_cost(ratio_of_means_ci(F[c2]["net_program_cost"], avf(c2)), 2), paper_cost(ratio_of_means_ci(F[c3]["net_program_cost"], avf(c3)), 2)],
        }
    out: Dict[Tuple[str, str], List[str]] = {}
    for tg, base, c2, c3, c1_na in [("2M", "2M Base Case", "2M Case 2", "2M Case 3", False),
                                    ("1M", "1M Base Case", "1M Case 2", "1M Case 3", True)]:
        for row, cells in block(tg, base, c2, c3, c1_na).items():
            out[(tg, row)] = cells
    return out


def write_table2_sheet(wb: Workbook, cells: Dict[Tuple[str, str], List[str]]) -> None:
    """Table 2 exactly as laid out in the submitted manuscript (2M block first, then 1M)."""
    ws = wb.create_sheet("Table 2 (paper)")
    ws.append(["Table 2: Tooth level results for 1,000 12 year olds under three intervention scenarios"])
    ws.append([None, "No sealants", "Case 1", "Case 2", "Case 3"])
    hdr_row(ws, 2)
    for tg in ["2M", "1M"]:
        ws.append([f"{tg}  [95% CI]"]); sub_row(ws, ws.max_row)
        for row in TABLE2_ROWS:
            ws.append([row] + cells[(tg, row)])
    for r in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=2, max_col=5):
        for c in r:
            c.alignment = Alignment(wrap_text=True)
    auto_fit(ws)
    ws.freeze_panes = "A3"


TABLE2_ROWS = ["# Sealants applied", "# Cumulative New Decay", "# Cumulative New Fillings",
               "Averted decay", "Averted fillings", "Sealant cost", "Filling cost", "Cost of program (net)",
               "Cost per averted decay ($/decay)", "Cost per averted filling ($/filling)"]


def write_raw_sheet(
    wb: Workbook,
    sc_finals: Dict[str, Dict[str, np.ndarray]],
) -> None:
    ws = wb.create_sheet("Final Metrics Raw")
    ws.append(["Level", "Name", "Metric", "Mean", "LCL (mean − 1.96·SE)", "UCL (mean + 1.96·SE)"])
    hdr_row(ws, 1)
    for sc_name, finals in sc_finals.items():
        for metric in ["sealants","decay","fillings","sealant_cost","filling_cost","net_program_cost"]:
            m, lo, hi = summary_stats_ci(finals[metric])
            ws.append(["Scenario", sc_name, metric, round(m,2), round(lo,2), round(hi,2)])
    auto_fit(ws)
    ws.freeze_panes = "A2"


# =====================================================================
# MAIN
# =====================================================================
def main() -> Path:
    import time
    print(f"Stochastic Markov Simulation  —  {N_SIM:,} iterations × {len(SCENARIOS)} scenarios")
    print(f"DV per period: {[round(float(v),4) for v in DV_PER_PERIOD]}")

    rng = np.random.default_rng(SEED)

    # Storage
    sc_paths = {
        sn: {m: np.zeros((N_SIM, N_CYCLES + 1), dtype=np.int32) for m in METRICS}
        for sn in SCENARIOS
    }
    sc_finals = {
        sn: {k: np.zeros(N_SIM, dtype=float)
             for k in ["sealants","decay","fillings","sealant_cost","filling_cost","net_program_cost"]}
        for sn in SCENARIOS
    }

    t0 = time.time()
    for i in range(N_SIM):
        if i % 2000 == 0 and i > 0:
            e = time.time() - t0
            print(f"  {i:>6,}/{N_SIM:,}  ({e:.1f}s elapsed, ~{e/i*(N_SIM-i):.0f}s left)")

        # R ~ Triangular(0.9382, 0.9485, 0.9588); AR ~ Triangular ±20% around base
        r_val = float(rng.triangular(R_TRI[0], R_TRI[1], R_TRI[2]))
        r_1m  = r_val
        r_2m  = r_val
        ar_1m = sample_range_parameter(rng, AR_BASE["1M"], PARAM_RANGE_PCT)
        ar_2m = sample_range_parameter(rng, AR_BASE["2M"], PARAM_RANGE_PCT)

        seal_unit = float(rng.triangular(*SEALANT_COST_TRI))
        fill_unit = float(rng.triangular(*FILLING_COST_TRI))

        for sn, cfg in SCENARIOS.items():
            r  = r_1m  if cfg.tooth_group == "1M" else r_2m
            ar = ar_1m if cfg.tooth_group == "1M" else ar_2m
            p  = run_one_scenario(cfg, r, ar, DV_PER_PERIOD, rng)

            for m in METRICS:
                sc_paths[sn][m][i, :] = p[m]

            fs = float(p["cumulative_sealants"][-1])
            fd = float(p["cumulative_new_decay"][-1])
            ff = float(p["cumulative_new_fillings"][-1])

            sc_finals[sn]["sealants"][i]     = fs
            sc_finals[sn]["decay"][i]        = fd
            sc_finals[sn]["fillings"][i]     = ff
            # Costs: each period's sealants / fillings × unit cost, discounted to period 0
            sc_finals[sn]["sealant_cost"][i] = seal_unit * float(p["sealants_this_period"] @ DISCOUNT_FACTORS)
            sc_finals[sn]["filling_cost"][i] = fill_unit * float(p["new_fillings"] @ DISCOUNT_FACTORS)

        # Net program cost per tooth group vs its own base case
        for tg, base in {"1M": "1M Base Case", "2M": "2M Base Case"}.items():
            base_fc = sc_finals[base]["filling_cost"][i]
            for sn, cfg in SCENARIOS.items():
                if cfg.tooth_group != tg:
                    continue
                net = (0.0 if sn == base else
                       sc_finals[sn]["sealant_cost"][i]
                       - (base_fc - sc_finals[sn]["filling_cost"][i]))
                sc_finals[sn]["net_program_cost"][i] = net

    print(f"Simulation done in {time.time()-t0:.1f}s  →  building Excel …")

    wb = Workbook()
    wb.remove(wb.active)

    write_inputs_sheet(wb)
    cells = table2_cells(sc_finals)
    write_table2_sheet(wb, cells)
    for sn, cfg in SCENARIOS.items():
        build_detail_sheet(wb, cfg.sheet_name, sc_paths[sn], cfg)
    write_results_sheet(wb, sc_finals)
    write_raw_sheet(wb, sc_finals)
    sens = one_way_sensitivity()
    write_sensitivity_sheet(wb, sens)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUT_DIR / OUTPUT_XLSX
    wb.save(out)
    print(f"\n✓  Saved:  {out}")
    print(f"   Sheets: {[ws.title for ws in wb.worksheets]}")

    verify_table2(cells)

    if MAKE_FIGURES:
        try:
            plot_figure2(sc_paths)
            plot_figure3(sens)
        except ImportError as e:
            print(f"(figures skipped: {e})")
    return out


# =====================================================================
# FIGURE 3 — ONE-WAY (DETERMINISTIC) SENSITIVITY ANALYSIS
# =====================================================================
def run_one_scenario_expected(
    cfg: ScenarioConfig, r_value: float, ar_value: float, dv_per_period: np.ndarray,
) -> Tuple[np.ndarray, np.ndarray, float]:
    """
    Expected-value (deterministic) twin of run_one_scenario: identical bookkeeping with every
    binomial draw replaced by its expectation.  Returns (sealants per period, fillings per period,
    cumulative new decay) for the tornado diagram.
    """
    S  = float(cfg.init_sound_sealed)
    NS = float(cfg.init_sound_unsealed)
    C  = 0.0
    seals = np.zeros(N_CYCLES + 1); fills = np.zeros(N_CYCLES + 1); cum_decay = 0.0
    if cfg.seal_at_baseline:
        seals[0] = NS; S += NS; NS = 0.0
    p_loss = 1.0 - r_value
    for t in range(1, N_CYCLES + 1):
        erupted   = float(cfg.eruption_schedule[t - 1])
        losses    = S * p_loss
        s_to_c    = losses * (ar_value / 2.0)
        s_to_ns   = losses - s_to_c
        ns_to_c   = NS * ar_value
        reseal    = s_to_ns if t <= cfg.reseal_loss_cycles   else 0.0
        seal_er   = erupted if t <= cfg.seal_eruption_cycles else 0.0
        new_decay = s_to_c + ns_to_c
        pool      = C + new_decay
        filled    = pool * float(np.clip(dv_per_period[t - 1], 0.0, 1.0))
        C         = pool - filled
        S         = (S - losses) + reseal + seal_er
        NS        = (NS - ns_to_c) + (s_to_ns - reseal) + (erupted - seal_er)
        seals[t] = reseal + seal_er; fills[t] = filled; cum_decay += new_decay
    return seals, fills, cum_decay


def cost_per_averted_decay_det(
    tooth_group: str, case: str, *, R: float = R_TRI[1], ar_mult: float = 1.0, dv_mult: float = 1.0,
    seal_cost: float = SEALANT_COST_TRI[1], fill_cost: float = FILLING_COST_TRI[1], rate: float = DISCOUNT_RATE,
) -> float:
    """Deterministic $/averted decay for `case` vs its base case (modal inputs unless overridden)."""
    ar   = AR_BASE[tooth_group] * ar_mult
    dv   = DV_PER_PERIOD * dv_mult
    disc = np.array([(1.0 + rate) ** (-t / 2.0) for t in range(N_CYCLES + 1)])
    sb, fb, db = run_one_scenario_expected(SCENARIOS[f"{tooth_group} Base Case"], R, ar, dv)
    sc, fc, dc = run_one_scenario_expected(SCENARIOS[case], R, ar, dv)
    net = seal_cost * float(sc @ disc) - fill_cost * float((fb - fc) @ disc)
    return net / (db - dc)


SENS_PCT = 0.20   # ±20% on every input (retention: ±20% on the loss rate 1−R → R_TRI bounds)

def one_way_sensitivity() -> Dict[str, List[Tuple[str, float, float]]]:
    """
    For Figure 3: baseline and the (low-input, high-input) endpoints of $/averted decay when each
    input is varied ±20% with all others at their base value.  Keys "2M" (Case 1: 2M sealed) and
    "1M" (Case 2: 1M sealed).  Each list = [(parameter label, value at low input, value at high input), ...].
    """
    out = {}
    for tg, case in [("2M", "2M Case 2"), ("1M", "1M Case 2")]:
        f = lambda **kw: cost_per_averted_decay_det(tg, case, **kw)
        loss = 1.0 - R_TRI[1]
        rows = [
            ("Sealant cost per tooth",     f(seal_cost=SEALANT_COST_TRI[1] * (1 - SENS_PCT)), f(seal_cost=SEALANT_COST_TRI[1] * (1 + SENS_PCT))),
            (f"{tg} cavity attack rate",   f(ar_mult=1 - SENS_PCT),                          f(ar_mult=1 + SENS_PCT)),
            ("Filling cost per tooth",     f(fill_cost=FILLING_COST_TRI[1] * (1 - SENS_PCT)), f(fill_cost=FILLING_COST_TRI[1] * (1 + SENS_PCT))),
            ("Retention rate",             f(R=1.0 - loss * (1 + SENS_PCT)),                 f(R=1.0 - loss * (1 - SENS_PCT))),
            ("Dental visit probability",   f(dv_mult=1 - SENS_PCT),                          f(dv_mult=1 + SENS_PCT)),
            ("Annual discount rate",       f(rate=DISCOUNT_RATE * (1 - SENS_PCT)),           f(rate=DISCOUNT_RATE * (1 + SENS_PCT))),
        ]
        out[tg] = {"baseline": f(), "rows": rows}
    return out


def write_sensitivity_sheet(wb: Workbook, sens) -> None:
    ws = wb.create_sheet("Sensitivity (Fig 3)")
    ws.append(["Group / case", "Parameter", "Low input → $/averted decay", "High input → $/averted decay", "Baseline"])
    hdr_row(ws, 1)
    for tg, case in [("2M", "Case 1"), ("1M", "Case 2")]:
        for name, lo, hi in sens[tg]["rows"]:
            ws.append([f"{tg} {case}", name, round(lo, 2), round(hi, 2), round(sens[tg]["baseline"], 2)])
    auto_fit(ws)


# =====================================================================
# FIGURES
# =====================================================================
def plot_figure2(sc_paths: Dict[str, Dict[str, np.ndarray]], fname: str = "figure2_cumulative_decay.png") -> None:
    """Figure 2: cumulative decay progression (% of the 4,000 molars), 2M panel above 1M panel."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    plt.rcParams["font.family"] = "serif"
    periods = np.arange(N_CYCLES + 1)
    xlabels = [f"P{t}\n({12 + t // 2}{'½' if t % 2 else ''}y)" for t in periods]
    mean_cum = lambda sn: sc_paths[sn]["cumulative_new_decay"].mean(axis=0)

    panels = [
        ("2(a) Cumulative decay progression - 2nd molar (2M)", "2M Base Case", "2M Case 2", "2M Case 3",
         "Case 1/Case 2: seal at baseline (no resealing)", (0, 40)),
        ("2(b) Cumulative decay progression - 1st molar (1M)", "1M Base Case", "1M Case 2", "1M Case 3",
         "Case 2: seal at baseline (no resealing)", (20, 50)),
    ]
    fig, axes = plt.subplots(2, 1, figsize=(6.2, 7.75), dpi=100)
    for ax, (title, base, mid, top, mid_label, ylim) in zip(axes, panels):
        cfg = SCENARIOS[base]
        pre = N_TOTAL_TEETH - cfg.init_sound_sealed - cfg.init_sound_unsealed - cfg.init_unerupted  # already decayed/filled/missing at age 12
        pre_pct = 100.0 * pre / N_TOTAL_TEETH
        y_base = pre_pct + 100.0 * mean_cum(base) / N_TOTAL_TEETH
        y_mid  = pre_pct + 100.0 * mean_cum(mid)  / N_TOTAL_TEETH
        y_top  = pre_pct + 100.0 * mean_cum(top)  / N_TOTAL_TEETH
        ax.fill_between(periods, y_mid, y_base, color="#f4c2c2", alpha=0.45, linewidth=0)
        ax.fill_between(periods, y_top, y_mid,  color="#bcd4ef", alpha=0.45, linewidth=0)
        ax.plot(periods, y_base, ":", color="#555555", marker="s", ms=5, lw=1.6, label="No sealants")
        ax.plot(periods, y_mid,  "-", color="#e41a1c", marker="o", ms=5, lw=1.8, label=mid_label)
        ax.plot(periods, y_top,  "-", color="#1f77d4", marker="s", ms=5, lw=1.8, label="Case 3: seal + reseal (periods 1-4)")
        ax.axhline(pre_pct, color="#999999", ls="-.", lw=0.9)
        ax.set_title(f"{title}\nStarting from {pre_pct:.1f}% pre-existing decay at age 12", fontweight="bold", fontsize=11)
        ax.set_ylabel("Total cumulative\ndecayed teeth (%)", fontsize=9.5)
        ax.set_ylim(*ylim); ax.set_xlim(-0.3, N_CYCLES + 0.3)
        ax.set_xticks(periods); ax.set_xticklabels(xlabels, fontsize=7.5)
        ax.yaxis.set_major_formatter(matplotlib.ticker.PercentFormatter(decimals=0))
        ax.tick_params(axis="y", labelsize=8.5)
        ax.grid(True, ls=":", color="#cccccc", lw=0.7)
        ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.20), ncol=3, fontsize=7.2,
                  columnspacing=1.0, handlelength=2.2, frameon=True)
    axes[1].set_xlabel("Half-yearly period (1 period = 6 months)", fontsize=9.5)
    fig.tight_layout(h_pad=2.2)
    fig.savefig(OUT_DIR / fname, dpi=100)
    plt.close(fig)
    print(f"✓  Saved:  {OUT_DIR / fname}")


def plot_figure3(sens, fname_prefix: str = "figure3_tornado") -> None:
    """Figure 3: one-way sensitivity tornado diagrams for 2M Case 1 and 1M Case 2."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    plt.rcParams["font.family"] = "sans-serif"
    for tg, case_label in [("2M", "Case 1"), ("1M", "Case 2")]:
        base = sens[tg]["baseline"]
        rows = sorted(sens[tg]["rows"], key=lambda r: abs(r[2] - r[1]), reverse=True)
        fig, ax = plt.subplots(figsize=(10.6, 5.6), dpi=100)
        for k, (name, lo_in, hi_in) in enumerate(rows):
            y = len(rows) - 1 - k
            left, right = min(lo_in, hi_in), max(lo_in, hi_in)
            ax.barh(y, right - left, left=left, height=0.5, color="#d3e4f5", edgecolor="#333333", lw=0.8)
            # marker shows which input value produced each endpoint
            ax.plot(lo_in, y, "o", ms=8, mfc="white", mec="#1f77b4", mew=1.6, zorder=3)
            ax.plot(hi_in, y, "^", ms=9, color="#d95f02", zorder=3)
            ax.text(left - 0.6, y, f"{left:.2f}", ha="right", va="center", fontsize=9)
            ax.text(right + 0.6, y, f"{right:.2f}", ha="left", va="center", fontsize=9)
        ax.axvline(base, color="black", ls="--", lw=1.5)
        ax.set_yticks(range(len(rows))); ax.set_yticklabels([r[0] for r in rows][::-1], fontsize=10)
        ax.set_xlim(0, max(r[2] for r in rows + [(None, 0, base)]) * 1.08 + 5)
        ax.set_xlabel("Cost per averted decay ($/decay)", fontsize=11)
        ax.set_title(f"One-way sensitivity analysis for {tg} {case_label}: Cost per averted decay\n", fontweight="bold", fontsize=13)
        ax.text(0.5, 1.02, f"Baseline = {base:.2f}", transform=ax.transAxes, ha="center", fontweight="bold", fontsize=10.5)
        ax.grid(axis="x", ls=":", color="#dddddd")
        ax.spines[["top", "right"]].set_visible(False)
        from matplotlib.lines import Line2D
        ax.legend(handles=[Line2D([], [], marker="o", ls="", mfc="white", mec="#1f77b4", mew=1.6, ms=8, label="Low input value"),
                           Line2D([], [], marker="^", ls="", color="#d95f02", ms=9, label="High input value")],
                  title="Endpoint direction", loc="lower right", fontsize=9, title_fontsize=9.5, frameon=False)
        fig.text(0.5, 0.005, "Markers show which parameter value produced each endpoint; lower x-axis values indicate lower cost per averted decay.",
                 ha="center", fontsize=8.5, color="#333333")
        fig.tight_layout(rect=(0, 0.03, 1, 1))
        fname = f"{fname_prefix}_{tg}.png"
        fig.savefig(OUT_DIR / fname, dpi=100)
        plt.close(fig)
        print(f"✓  Saved:  {OUT_DIR / fname}")


# =====================================================================
# SELF-CHECK AGAINST THE MANUSCRIPT (Table 2 and Figure 3 values as reported)
# =====================================================================
PAPER_TABLE2: Dict[Tuple[str, str], List[str]] = {
    ("2M", "# Sealants applied"):        ["0.0 [0.0, 0.0]", "2,136.0 [2,136.0, 2,136.0]", "2,136.0 [2,136.0, 2,136.0]", "3,773.3 [3,769.8, 3,776.8]"],
    ("2M", "# Cumulative New Decay"):    ["1,131.4 (28.28%) [1,126.2, 1,136.5]", "621.5 (15.54%) [618.3, 624.8]", "621.5 (15.54%) [618.3, 624.8]", "263.2 (6.58%) [261.3, 265.0]"],
    ("2M", "# Cumulative New Fillings"): ["921.7 (23.04%) [917.4, 926.0]", "473.1 (11.83%) [470.5, 475.7]", "473.1 (11.83%) [470.5, 475.7]", "173.7 (4.34%) [172.3, 175.0]"],
    ("2M", "Averted decay"):             ["N/A", "509.8 [506.7, 513.0]", "509.8 [506.7, 513.0]", "868.2 [864.1, 872.3]"],
    ("2M", "Averted fillings"):          ["N/A", "448.6 [445.8, 451.4]", "448.6 [445.8, 451.4]", "748.1 [744.4, 751.7]"],
    ("2M", "Sealant cost"):              ["$0 [$0, $0]", "$71,694 [$71,437, $71,951]", "$71,694 [$71,437, $71,951]", "$124,832 [$124,376, $125,288]"],
    ("2M", "Filling cost"):              ["$106,173 [$105,549, $106,797]", "$53,574 [$53,222, $53,926]", "$53,574 [$53,222, $53,926]", "$19,219 [$19,053, $19,386]"],
    ("2M", "Cost of program (net)"):     ["$0 [$0, $0]", "$19,095 [$18,645, $19,545]", "$19,095 [$18,645, $19,545]", "$37,878 [$37,186, $38,571]"],
    ("2M", "Cost per averted decay ($/decay)"):     ["N/A", "$37.46 [$36.41, $38.50]", "$37.46 [$36.41, $38.50]", "$43.63 [$42.69, $44.57]"],
    ("2M", "Cost per averted filling ($/filling)"): ["N/A", "$42.56 [$41.35, $43.77]", "$42.56 [$41.35, $43.77]", "$50.63 [$49.53, $51.73]"],
    ("1M", "# Sealants applied"):        ["0", "N/A", "1,696.0 [1,696.0, 1,696.0]", "2,310.5 [2,307.1, 2,313.8]"],
    ("1M", "# Cumulative New Decay"):    ["708.4 (17.71%) [705.1, 711.8]", "N/A", "315.8 (7.90%) [313.7, 317.9]", "186.6 (4.67%) [185.2, 188.0]"],
    ("1M", "# Cumulative New Fillings"): ["575.3 (14.38%) [572.5, 578.1]", "N/A", "230.4 (5.76%) [228.8, 232.1]", "122.8 (3.07%) [121.8, 123.8]"],
    ("1M", "Averted decay"):             ["N/A", "N/A", "392.6 [390.1, 395.1]", "521.8 [519.1, 524.5]"],
    ("1M", "Averted fillings"):          ["N/A", "N/A", "344.9 [342.7, 347.1]", "452.5 [450.0, 454.9]"],
    ("1M", "Sealant cost"):              ["$0 [$0, $0]", "N/A", "$56,926 [$56,722, $57,130]", "$76,802 [$76,509, $77,096]"],
    ("1M", "Filling cost"):              ["$66,359 [$65,958, $66,759]", "N/A", "$25,923 [$25,718, $26,129]", "$13,610 [$13,489, $13,731]"],
    ("1M", "Cost of program (net)"):     ["$0 [$0, $0]", "N/A", "$16,490 [$16,142, $16,839]", "$24,054 [$23,614, $24,493]"],
    ("1M", "Cost per averted decay ($/decay)"):     ["N/A", "N/A", "$42.00 [$40.93, $43.08]", "$46.10 [$45.10, $47.10]"],
    ("1M", "Cost per averted filling ($/filling)"): ["N/A", "N/A", "$47.82 [$46.57, $49.06]", "$53.16 [$51.99, $54.34]"],
}

PAPER_FIGURE3 = {  # baseline, then (low-input, high-input) endpoints in the order of one_way_sensitivity()
    "2M": (37.02, [(8.82, 65.23), (65.80, 17.71), (57.82, 16.23), (46.13, 28.50), (43.12, 32.97), (34.92, 39.07)]),
    "1M": (41.42, [(12.37, 70.46), (71.26, 21.39), (62.18, 20.65), (50.81, 32.61), (47.53, 37.35), (39.30, 43.46)]),
}


def verify_table2(cells: Dict[Tuple[str, str], List[str]]) -> None:
    """Compare every generated Table 2 cell with the manuscript (numerically, at the printed precision)."""
    import re
    nums = lambda s: re.findall(r"-?\d[\d,]*\.?\d*", s)
    n_ok = n_tot = 0; bad = []
    for key, paper_cells in PAPER_TABLE2.items():
        for col, (pc, gc) in enumerate(zip(paper_cells, cells[key])):
            if pc == "N/A":
                continue
            n_tot += 1
            if nums(pc) == nums(gc) or (pc == "0" and nums(gc) == ["0.0", "0.0", "0.0"]):
                n_ok += 1
            else:
                bad.append((key, col, pc, gc))
    print(f"\nTable 2 self-check vs submitted manuscript: {n_ok}/{n_tot} cells identical")
    for key, col, pc, gc in bad:
        print(f"   MISMATCH {key} col {col}: paper '{pc}'  vs  generated '{gc}'")
    sens = one_way_sensitivity(); ok = 0; tot = 0
    for tg, (base, ends) in PAPER_FIGURE3.items():
        tot += 1; ok += round(sens[tg]["baseline"], 2) == base
        for (lo, hi), (_, glo, ghi) in zip(ends, sens[tg]["rows"]):
            tot += 2; ok += round(glo, 2) == lo; ok += round(ghi, 2) == hi
    print(f"Figure 3 self-check vs submitted manuscript: {ok}/{tot} values identical")


if __name__ == "__main__":
    main()