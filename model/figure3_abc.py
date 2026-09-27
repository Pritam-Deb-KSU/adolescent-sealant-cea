"""
Figure 3 with the probabilistic sensitivity analysis added as panel (c).

Panels (a) and (b) are the submitted Figure 3 PDF, reused unchanged (vector content, not re-drawn).
Panel (c) — cost-effectiveness acceptability curves — is drawn in the same style and appended below:
  • same page width (501.2 pt) and the same axes left/right edges as (a)/(b) (x = 109.48–490.00 pt)
  • same fonts: Times-Roman / Times-Bold as standard (core) PDF fonts, like the submitted file
  • same text sizes: titles 13 pt bold (two lines, 13.2 pt apart), axis labels 11 pt,
    tick labels 8 pt, annotation 9.5 pt, legend 8.2 pt
  • same frame (0.9 pt box), grid (0.45 pt dashed light grey) and reference line (1.4 pt dashed black)

Inputs : results/figure3_ab_submitted.pdf  (the submitted Figure 3, panels a–b)
Outputs: results/figure3_abc.pdf            (Figure 3 with panel c; vector)
         results/figure3_abc.png            (600 dpi preview for the Word draft)
"""
from __future__ import annotations

from pathlib import Path

import numpy as np

import psa_ceac as P
from sealant_markov_model import FILLING_COST_TRI

RESULTS  = Path(__file__).resolve().parent.parent / "results"
AB_PDF   = RESULTS / "figure3_ab_submitted.pdf"
OUT_PDF  = RESULTS / "figure3_abc.pdf"
OUT_PNG  = RESULTS / "figure3_abc.png"

# ── geometry of the submitted page (points) ───────────────────────────────────
PAGE_W, AB_H   = 501.2, 621.3536
AX_L, AX_R     = 109.48, 490.00          # axes left/right edges of panels (a) and (b)
AX_H           = 160.0                   # panel (c) axes height
AX_B           = 36.0                    # panel (c) axes bottom
GAP_TO_AB      = 16.0                    # clearance below the (a)/(b) legend
AB_INK_BOTTOM  = 9.0                     # lowest ink of the submitted page (legend text)
TITLE_DY1, TITLE_DY2 = 21.21, 8.00       # title baselines above axes top, as in (a)/(b)
TITLE_ASCENT   = 30.0                    # top of the (c) title ink above the axes top (21.21 + cap height)
C_H = AX_B + AX_H + TITLE_ASCENT + GAP_TO_AB - AB_INK_BOTTOM   # height of the panel (c) strip (= 233 pt)

COL = {"Case 1": "#0072B2", "Case 2": "#D55E00", "Case 3": "#009E73"}
LS  = {"Case 1": "-", "Case 2": "--", "Case 3": "-."}
LABEL = {"Case 1": "Case 1: eligible 2M sealed at age 12",
         "Case 2": "Case 2: Case 1 + eligible 1M sealed at age 12",
         "Case 3": "Case 3: Case 2 + lost sealants re-applied at ages 13 and 14"}
GRID = dict(color="#d3d3d3", linestyle="--", linewidth=0.45)


def draw_panel_c(path: Path) -> float:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    plt.rcParams.update({
        "pdf.use14corefonts": True, "font.family": "serif", "font.serif": ["Times", "Times New Roman"],
        "axes.linewidth": 0.9, "font.size": 10,
    })
    Pd, S, _ = P.run_psa()
    wtp = P.WTP_GRID
    pair = P.ceac_pairwise(S, wtp)

    fig = plt.figure(figsize=(PAGE_W / 72.0, C_H / 72.0))
    ax = fig.add_axes([AX_L / PAGE_W, AX_B / C_H, (AX_R - AX_L) / PAGE_W, AX_H / C_H])

    ax.grid(True, axis="both", **GRID); ax.set_axisbelow(True)
    for s in ["Case 1", "Case 2", "Case 3"]:
        ax.plot(wtp, pair[s], color=COL[s], ls=LS[s], lw=1.8, label=LABEL[s], zorder=3)
    fee = FILLING_COST_TRI[1]
    ax.axvline(fee, color="black", ls="--", lw=1.4, zorder=2)
    ax.text(fee + 3, 0.62, "Cost of restoration\nper tooth (Table 1)", fontsize=9.5, va="center", ha="left")

    ax.set_xlim(0, 200); ax.set_ylim(0, 1.03)   # headroom so the plateau at 1.0 is not hidden by the frame
    ax.set_xticks(np.arange(0, 201, 25)); ax.set_yticks(np.arange(0, 1.01, 0.2))
    ax.set_yticklabels([f"{v:.1f}" for v in np.arange(0, 1.01, 0.2)])
    ax.tick_params(labelsize=8)
    ax.set_xlabel("Willingness to pay ($ per averted decay)", fontsize=11)
    ax.set_ylabel("Probability cost-effective\nvs no sealants", fontsize=11)
    ax.legend(loc="lower right", fontsize=8.2, frameon=True, framealpha=1, edgecolor="#d3d3d3",
              fancybox=False, handlelength=2.6, borderpad=0.5)

    cx = (AX_L + AX_R) / 2 / PAGE_W
    top = AX_B + AX_H
    fig.text(cx, (top + TITLE_DY1) / C_H, "3(c) Cost-effectiveness acceptability curves",
             ha="center", va="baseline", fontsize=13, fontweight="bold")
    fig.text(cx, (top + TITLE_DY2) / C_H, "Probabilistic sensitivity analysis, 10,000 iterations",
             ha="center", va="baseline", fontsize=13, fontweight="bold")
    fig.savefig(path)
    plt.close(fig)
    return C_H


def stack(ab_pdf: Path, c_pdf: Path, out_pdf: Path, c_h: float) -> None:
    from pypdf import PageObject, PdfReader, PdfWriter, Transformation
    ab = PdfReader(str(ab_pdf)).pages[0]
    c = PdfReader(str(c_pdf)).pages[0]
    assert abs(float(ab.mediabox.width) - PAGE_W) < 0.01 and abs(float(ab.mediabox.height) - AB_H) < 0.01, "unexpected Figure 3 page size"
    page = PageObject.create_blank_page(width=PAGE_W, height=AB_H + c_h)
    page.merge_transformed_page(c, Transformation())                      # panel (c) at the bottom
    page.merge_transformed_page(ab, Transformation().translate(0, c_h))   # submitted (a)+(b) on top, unchanged
    w = PdfWriter(); w.add_page(page)
    w.add_metadata({"/Title": "Figure 3", "/Creator": "Matplotlib (panels a-c); pypdf (stacking)"})
    with open(out_pdf, "wb") as f:
        w.write(f)


def main():
    tmp_c = RESULTS / "_figure3c_panel.pdf"
    c_h = draw_panel_c(tmp_c)
    stack(AB_PDF, tmp_c, OUT_PDF, c_h)
    tmp_c.unlink()
    import subprocess
    subprocess.run(["pdftoppm", "-png", "-r", "600", "-singlefile", str(OUT_PDF), str(OUT_PNG.with_suffix(""))], check=False)
    print(f"✓  Saved:  {OUT_PDF}  ({PAGE_W / 72:.2f} × {(AB_H + c_h) / 72:.2f} in)")
    if OUT_PNG.exists():
        print(f"✓  Saved:  {OUT_PNG}")


if __name__ == "__main__":
    main()
