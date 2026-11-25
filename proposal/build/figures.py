"""Draw the three figures used in the proposal (black-and-white, print-ready)."""
import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

OUT = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
OUT.mkdir(parents=True, exist_ok=True)
plt.rcParams.update({"font.family": "Liberation Serif", "font.size": 11})

INK = "#1a1a1a"
FILL_IV = "#e8eef7"
FILL_DV = "#eaf4ec"
FILL_NOTE = "#f4f4f4"


def box(ax, x, y, w, h, text, fill, bold_title=None, fs=10.5, ls="-"):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.12",
                                fc=fill, ec=INK, lw=1.2, ls=ls))
    if bold_title:
        ax.text(x + w / 2, y + h - 0.28, bold_title, ha="center", va="top", fontsize=fs + 0.5,
                fontweight="bold", color=INK)
        n = bold_title.count("\n") + 1
        ax.text(x + w / 2, y + h - 0.42 - 0.36 * n, text, ha="center", va="top", fontsize=fs,
                color=INK, linespacing=1.45)
    else:
        ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=fs, color=INK,
                linespacing=1.45)


def arrow(ax, x1, y1, x2, y2, label=None, lx=0, ly=0.18, ls="-"):
    ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="-|>", mutation_scale=16,
                                 lw=1.4, color=INK, ls=ls, shrinkA=0, shrinkB=0))
    if label:
        ax.text((x1 + x2) / 2 + lx, (y1 + y2) / 2 + ly, label, ha="center", va="bottom",
                fontsize=10, color=INK, style="italic",
                bbox=dict(fc="white", ec="none", pad=1.2))


# ---------------------------------------------------------------- Figure 2.1
fig, ax = plt.subplots(figsize=(9.2, 5.6))
ax.set_xlim(0, 12)
ax.set_ylim(0, 7.4)
ax.axis("off")

# boundary
ax.add_patch(FancyBboxPatch((0.15, 0.15), 11.7, 7.05, boxstyle="round,pad=0,rounding_size=0.2",
                            fc="none", ec="#777777", lw=1, ls=(0, (4, 3))))
ax.text(0.45, 6.92, "Study boundary: non-financial LQ45 constituents, IDX, Q1 2016 – Q4 2025",
        fontsize=9.5, color="#444444", style="italic", va="top")

box(ax, 0.6, 2.0, 3.7, 3.9,
    "SUE = (EPS$_{t}$ − EPS$_{t-4}$) / P$_{t}$\nranked into quintiles (SUER)\n\n"
    "D1: Direction of surprise\n(positive / negative)\nD2: Magnitude of surprise",
    FILL_IV, bold_title="Earnings Surprise (X)\nIndependent variable", fs=10)

box(ax, 7.4, 4.75, 4.0, 1.55, "CAR(−1,+1)\nimmediate reaction", FILL_DV,
    bold_title="Announcement reaction", fs=10)
box(ax, 7.4, 2.35, 4.0, 2.05, "CAR(+2,+61), main measure\nCAR(+2,+31), supplementary",
    FILL_DV, bold_title="Post-Earnings Announcement\nDrift (Y), dependent variable", fs=10)
box(ax, 7.4, 0.45, 4.0, 1.5,
    "Drift follows the sign of the surprise:\nD5 drifts up, D1 drifts down", FILL_DV,
    bold_title="Direction of drift", fs=9.6)

arrow(ax, 4.3, 4.6, 7.4, 5.5, "H2 (+)", ly=0.12)
arrow(ax, 4.3, 3.9, 7.4, 3.3, "H1 (+), main hypothesis", ly=0.12)
arrow(ax, 4.3, 2.6, 7.4, 1.25, "H3a (+) / H3b (−)", ly=0.02)

ax.text(0.6, 0.45,
        "No mediating or moderating variable in the main model (by design).\n"
        "Quarter fixed effects absorb market-wide shocks in the regression test.",
        fontsize=9, color="#444444", ha="left", va="bottom", linespacing=1.4)
fig.tight_layout()
fig.savefig(OUT / "fig_conceptual_model.png", dpi=220)
plt.close(fig)

# ---------------------------------------------------------------- Figure 3.1
steps = [
    ("1. Define the universe", "Historical LQ45 lists (Feb/Aug reviews), 2016–2025; drop financial firms"),
    ("2. Build quarterly EPS", "Convert year-to-date reports to discrete quarters; adjust EPS$_{t-4}$ for corporate actions"),
    ("3. Score the surprise", "SUE = (EPS$_{t}$ − EPS$_{t-4}$) / P$_{t}$; winsorize 1%/99%"),
    ("4. Rank every quarter", "Quintiles D1 (most negative) … D5 (most positive) → SUER"),
    ("5. Measure returns", "AR = R$_{i}$ − R$_{IHSG}$; CAR(−1,+1), CAR(+2,+31), CAR(+2,+61) around day 0"),
    ("6. Test A: portfolio spread", "Quarterly D5 − D1 spread → one-sample t-test over ~40 quarters"),
    ("7. Test B: regression", "CAR on scaled SUER, quarter fixed effects, firm-clustered SE"),
    ("8. Robustness", "Alternative benchmark, terciles, Wilcoxon test, excluding 2020"),
]
fig, ax = plt.subplots(figsize=(8.6, 9.0))
ax.set_xlim(0, 10)
ax.set_ylim(0, len(steps) * 1.12 + 0.2)
ax.axis("off")
h = 0.82
for i, (t, d) in enumerate(steps):
    y = (len(steps) - 1 - i) * 1.12 + 0.3
    fill = FILL_DV if i >= 5 else FILL_IV
    ax.add_patch(FancyBboxPatch((0.4, y), 9.2, h, boxstyle="round,pad=0.02,rounding_size=0.1",
                                fc=fill, ec=INK, lw=1.1))
    ax.text(5, y + h - 0.2, t, ha="center", va="top", fontsize=11.5, fontweight="bold")
    ax.text(5, y + 0.2, d, ha="center", va="bottom", fontsize=10)
    if i < len(steps) - 1:
        ax.add_patch(FancyArrowPatch((5, y), (5, y - 0.3 + 0.02), arrowstyle="-|>",
                                     mutation_scale=13, lw=1.2, color=INK))
fig.tight_layout()
fig.savefig(OUT / "fig_research_procedure.png", dpi=220)
plt.close(fig)

# ---------------------------------------------------------------- Figure 3.2 (Gantt)
months = ["Dec\n2026", "Jan\n2027", "Feb\n2027", "Mar\n2027", "Apr\n2027", "May\n2027",
          "Jun\n2027", "Jul\n2027"]
tasks = [
    ("Proposal revision and supervisor approval", 0, 1.5),
    ("Collect LQ45 lists, announcement dates, EPS", 1, 2.0),
    ("Collect daily prices and IHSG; data cross-checks", 1.5, 1.5),
    ("Construct SUE, SUER and CAR; data cleaning", 2.5, 1.5),
    ("Hypothesis tests and robustness checks", 3.5, 1.5),
    ("Write Chapter IV (results and discussion)", 4.5, 1.5),
    ("Write Chapter V; full draft review", 5.5, 1.0),
    ("Journal manuscript draft and 3MT video", 6.0, 1.0),
    ("Final revision, plagiarism check, defense", 6.5, 1.5),
]
fig, ax = plt.subplots(figsize=(9.4, 4.6))
for i, (name, start, dur) in enumerate(tasks):
    ax.barh(i, dur, left=start, height=0.55, color="#9db4d6" if i < 5 else "#a8cfa9",
            edgecolor=INK, lw=0.8)
ax.set_yticks(range(len(tasks)))
ax.set_yticklabels([t[0] for t in tasks], fontsize=10)
ax.invert_yaxis()
ax.set_xticks([m + 0.5 for m in range(len(months))])
ax.set_xticklabels(months, fontsize=9.5)
ax.set_xticks(range(len(months) + 1), minor=True)
ax.set_xlim(0, len(months))
ax.grid(axis="x", which="minor", color="#cccccc", lw=0.8)
ax.tick_params(axis="x", which="both", length=0)
ax.tick_params(axis="y", length=0)
for s in ["top", "right", "left"]:
    ax.spines[s].set_visible(False)
fig.tight_layout()
fig.savefig(OUT / "fig_gantt.png", dpi=220)
plt.close(fig)
print("figures written to", OUT)
