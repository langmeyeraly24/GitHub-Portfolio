"""
media_market_visualizations.py
-------------------------------
Charts that tell the "where to buy media" story for the Mike Collins campaign
plan. These pair with the two Caliper maps Alyssa uploaded (TV DMA + radio
metros) - the maps show WHERE markets are geographically, these charts show
WHAT'S inside each market politically.

Outputs (all PNGs, plan-embedding sized):
    1. tv_market_priority.png        - voters per market vs partisan lean
    2. tv_market_net_margin.png      - net R vote margin available per market
    3. tv_market_strategy_quadrant.png  - quadrant view (share x lean, bubble = voters)
    4. radio_coverage_gap.png        - share of GA voters with NO radio reach

Author: Alyssa Langmeyer  (GOVT 520 - Mike Collins for U.S. Senate 2026)
"""

from pathlib import Path
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

from ga_media_markets import attach_markets

# ---------------------------------------------------------------------------
# Setup: load data and reproduce the market summaries
# ---------------------------------------------------------------------------

HERE = Path(__file__).parent
master = pd.read_excel(HERE / "ga_senate_2026_county_master.xlsx",
                       sheet_name="county_data")
merged = attach_markets(master)

# Read pre-computed summaries from the merged workbook so I don't have to
# duplicate the wavg logic here
summary_xlsx = HERE / "ga_county_master_WITH_MEDIA_MARKETS.xlsx"
tv = pd.read_excel(summary_xlsx, sheet_name="tv_market_summary")
radio = pd.read_excel(summary_xlsx, sheet_name="radio_market_summary")

# Campaign color palette - keep this consistent across all charts
RED   = "#B92532"   # Republican / Collins
BLUE  = "#1F4E89"   # Democratic / Ossoff
NEUTRAL = "#888888"
LIGHT = "#E8E8E8"


# ---------------------------------------------------------------------------
# Chart 1: TV market priority - voters by market, colored by partisan lean
# ---------------------------------------------------------------------------

fig, ax = plt.subplots(figsize=(11, 6.5))
tv_sorted = tv.sort_values("reg_voters_2026", ascending=True)

# Color each bar based on which side the partisan blend leans toward
def lean_color(pct_r):
    if pct_r >= 55:   return "#A02029"   # strong R
    if pct_r >= 50:   return "#D87680"   # lean R
    if pct_r >= 45:   return "#B0B0B0"   # competitive
    if pct_r >= 40:   return "#7B97C0"   # lean D
    return "#1F4E89"                      # strong D

colors = [lean_color(p) for p in tv_sorted["partisan_blend_r"]]
bars = ax.barh(tv_sorted["market"], tv_sorted["reg_voters_2026"] / 1000,
               color=colors, edgecolor="white", linewidth=0.5)

# Label each bar with R% and voter count
for bar, pct_r, voters in zip(bars, tv_sorted["partisan_blend_r"],
                              tv_sorted["reg_voters_2026"]):
    width = bar.get_width()
    label = f"{pct_r:.0f}% R   ({voters/1000:.0f}K voters)"
    ax.text(width + 50, bar.get_y() + bar.get_height()/2, label,
            va="center", ha="left", fontsize=9, color="#222")

ax.set_xlabel("Registered voters (thousands)", fontsize=11)
ax.set_title("Georgia TV Markets - Voter Universe & Partisan Lean\n"
             "Atlanta dominates the state (68% of all voters) but leans Ossoff",
             fontsize=12, loc="left", pad=12)
ax.set_xlim(0, tv_sorted["reg_voters_2026"].max() / 1000 * 1.30)
ax.grid(axis="x", linestyle="--", alpha=0.3)
ax.spines[["top", "right"]].set_visible(False)

# Legend
legend_elements = [
    mpatches.Patch(color="#A02029", label="Strong R market (55%+ R)"),
    mpatches.Patch(color="#D87680", label="Lean R market (50-55%)"),
    mpatches.Patch(color="#B0B0B0", label="Competitive (45-50%)"),
    mpatches.Patch(color="#7B97C0", label="Lean D (40-45%)"),
    mpatches.Patch(color="#1F4E89", label="Strong D (<40%)"),
]
ax.legend(handles=legend_elements, loc="lower right", fontsize=8,
          framealpha=0.95, title="Partisan blend (R%)", title_fontsize=8)

plt.tight_layout()
plt.savefig(HERE / "tv_market_priority.png", dpi=200, bbox_inches="tight")
plt.close()
print("Wrote: tv_market_priority.png")

"""
The story this chart tells:

Atlanta is so huge (5.5M voters, 68% of the state) that it visually dwarfs
everything else - which is exactly the point. The bar is light blue
(lean D) which is the Collins problem in one image: we cannot win the
biggest market on our side, so the question becomes how much we lose it by.

Below Atlanta you can see the secondary markets stack up - Savannah, Macon,
Augusta-Aiken are each 4-7% of the state and run roughly 49-54% R. These
are the "swing the margin" markets - not as big as Atlanta but reachable,
and right at the persuasion sweet spot.

Below those, Chattanooga / Jacksonville / Greenville are tiny (2-3% each)
but DEEP red (75%+ R). That's the GOTV layer of the strategy.
"""


# ---------------------------------------------------------------------------
# Chart 2: Net R margin available per TV market
# ---------------------------------------------------------------------------

fig, ax = plt.subplots(figsize=(11, 6))
tv_margin = tv.sort_values("est_r_margin_2026", ascending=True)
colors = [RED if v >= 0 else BLUE for v in tv_margin["est_r_margin_2026"]]
ax.barh(tv_margin["market"], tv_margin["est_r_margin_2026"] / 1000,
        color=colors, edgecolor="white", linewidth=0.5)

# Annotate each bar with the value
for i, (m, v) in enumerate(zip(tv_margin["market"], tv_margin["est_r_margin_2026"])):
    offset = 8 if v >= 0 else -8
    align = "left" if v >= 0 else "right"
    ax.text(v / 1000 + offset, i, f"{v/1000:+,.0f}K",
            va="center", ha=align, fontsize=9, color="#222")

ax.axvline(0, color="#444", linewidth=1)
ax.set_xlabel("Estimated net R vote margin (thousands)", fontsize=11)
ax.set_title("Net R-vs-D Vote Margin Available Per TV Market (2026 baseline)\n"
             "Atlanta's deficit (-549K) is bigger than every R-positive market combined",
             fontsize=12, loc="left", pad=12)
ax.grid(axis="x", linestyle="--", alpha=0.3)
ax.spines[["top", "right"]].set_visible(False)

# Add a "total" annotation
total_r_margin = tv["est_r_margin_2026"].sum()
ax.text(0.02, 0.02,
        f"Statewide baseline net margin: {total_r_margin/1000:+,.0f}K votes "
        f"(based on partisan blend applied to 2026 reg file)",
        transform=ax.transAxes, fontsize=8, style="italic",
        color="#444")

plt.tight_layout()
plt.savefig(HERE / "tv_market_net_margin.png", dpi=200, bbox_inches="tight")
plt.close()
print("Wrote: tv_market_net_margin.png")

"""
This is the chart that shows the path-to-victory math at the market level:

The Atlanta bar extends LEFT to -549K. That's how many R votes we are
"underwater" in Atlanta DMA based on baseline partisan turnout assumptions.

For Collins to win statewide, the SUM of all the right-pointing bars
(Chattanooga +111K, Jacksonville +70K, Albany +59K, Savannah +46K,
Macon +45K, Tallahassee +40K, Greenville +42K, Dothan +3K) has to
exceed the Atlanta+Columbus deficit (-549K + -23K + -8K = -580K).

Current total right-side: ~416K of R margin
Current total left-side:  ~580K of D margin
Statewide baseline gap: ~-164K (Ossoff favored by ~2-3 points, matches polling)

So the strategic question becomes: where can we move the most votes per
dollar spent? That's a media efficiency question for the next phase.
"""


# ---------------------------------------------------------------------------
# Chart 3: Strategy quadrant (share of voters x partisan lean, sized by voters)
# ---------------------------------------------------------------------------

fig, ax = plt.subplots(figsize=(11, 7))

x = tv["share_of_statewide_reg"]
y = tv["partisan_blend_r"]
sizes = tv["reg_voters_2026"] / 1500   # scale bubbles

# Color by media strategy bucket
def quadrant_color(pct_r, share):
    if share < 1:
        return "#CCCCCC"
    if pct_r >= 55:
        return "#A02029"     # mobilize / run up
    if pct_r >= 45:
        return "#D6A02E"     # persuasion battleground
    return "#1F4E89"          # limit losses (D-leaning)

colors = [quadrant_color(pct_r, sh) for pct_r, sh in zip(y, x)]

ax.scatter(x, y, s=sizes, c=colors, alpha=0.78, edgecolors="white", linewidth=1.5)

# Label each bubble
for _, row in tv.iterrows():
    offset_x = 0.4 if row["share_of_statewide_reg"] < 50 else -0.4
    ha = "left" if row["share_of_statewide_reg"] < 50 else "right"
    ax.annotate(row["market"],
                (row["share_of_statewide_reg"], row["partisan_blend_r"]),
                xytext=(offset_x, 0.5), textcoords="offset fontsize",
                fontsize=9, ha=ha, va="bottom", color="#222")

# Reference lines: 50% partisan lean, and a "small market" cutoff
ax.axhline(50, color="#444", linewidth=1, linestyle="--", alpha=0.6)
ax.axhline(55, color="#A02029", linewidth=0.8, linestyle=":", alpha=0.5)
ax.axhline(45, color="#1F4E89", linewidth=0.8, linestyle=":", alpha=0.5)

# Zone labels
ax.text(35, 75, "RUN UP MARGINS\n(mobilize the base)", fontsize=9,
        color="#A02029", ha="center", va="center", fontweight="bold")
ax.text(35, 50, "PERSUASION BATTLEGROUND\n(swing the margin)", fontsize=9,
        color="#8B6914", ha="center", va="center", fontweight="bold")
ax.text(35, 41, "LIMIT LOSSES\n(reach soft Ds + indies)", fontsize=9,
        color="#1F4E89", ha="center", va="center", fontweight="bold")

ax.set_xlabel("Share of GA registered voters (%)", fontsize=11)
ax.set_ylabel("Partisan blend - Republican share (%)", fontsize=11)
ax.set_title("Georgia TV Market Strategy Quadrant\n"
             "Bubble size = total registered voters in the market",
             fontsize=12, loc="left", pad=12)
ax.set_xlim(-2, 75)
ax.set_ylim(35, 85)
ax.grid(linestyle="--", alpha=0.3)
ax.spines[["top", "right"]].set_visible(False)

plt.tight_layout()
plt.savefig(HERE / "tv_market_strategy_quadrant.png", dpi=200, bbox_inches="tight")
plt.close()
print("Wrote: tv_market_strategy_quadrant.png")

"""
This is the "where do we put the money" picture at one glance:

- Top of the chart (deep red zone): small, cheap, deep-red markets where
  Collins voters live. Money here is for GOTV - turnout operations,
  mail, digital. Not persuasion TV (no one to persuade).

- Middle band (yellow zone, 45-55% R): the persuasion sweet spot. Macon,
  Savannah, Augusta-Aiken, and Columbus all live here. They're 2-7% of
  the state each, big enough that TV is efficient, with enough swing
  voters that ad spend can move the dial. This is the heaviest paid TV
  band.

- Bottom right (Atlanta, huge blue bubble): the giant we cannot ignore.
  Strategy here is suburban-focused persuasion + targeted GOTV in red
  pockets like Cherokee/Forsyth/Hall.
"""


# ---------------------------------------------------------------------------
# Chart 4: Radio coverage gap - rural Collins voters with no radio reach
# ---------------------------------------------------------------------------

fig, axes = plt.subplots(1, 2, figsize=(12, 5.5))

# left panel: share of statewide voters in each radio market vs "no market"
radio_sorted = radio.sort_values("reg_voters_2026", ascending=True)
ax = axes[0]
colors_radio = []
for m, p in zip(radio_sorted["market"], radio_sorted["partisan_blend_r"]):
    if m == "No rated radio market":
        colors_radio.append("#8B6914")
    elif p >= 55:
        colors_radio.append(RED)
    elif p >= 45:
        colors_radio.append("#B0B0B0")
    else:
        colors_radio.append(BLUE)

ax.barh(radio_sorted["market"], radio_sorted["reg_voters_2026"] / 1000,
        color=colors_radio, edgecolor="white", linewidth=0.5)
for i, (m, v, p) in enumerate(zip(radio_sorted["market"],
                                  radio_sorted["reg_voters_2026"],
                                  radio_sorted["partisan_blend_r"])):
    ax.text(v / 1000 + 30, i, f"{p:.0f}% R", va="center", fontsize=9, color="#222")
ax.set_xlabel("Registered voters (thousands)", fontsize=10)
ax.set_title("Radio metro coverage by registered voters", fontsize=11, loc="left")
ax.grid(axis="x", linestyle="--", alpha=0.3)
ax.spines[["top", "right"]].set_visible(False)

# right panel: focused pie - share of GA voters in NO radio market
ax = axes[1]
no_radio = radio[radio["market"] == "No rated radio market"]["reg_voters_2026"].sum()
yes_radio = radio[radio["market"] != "No rated radio market"]["reg_voters_2026"].sum()

# But the interesting question: of voters NOT in a radio market,
# what's their partisan lean?
no_radio_subset = merged[merged["radio_market"] == "No rated radio market"]
no_radio_r_pct = (
    (no_radio_subset["reg_total_2026"] * no_radio_subset["partisan_blend_r"]).sum()
    / no_radio_subset["reg_total_2026"].sum()
)

wedges, texts, autotexts = ax.pie(
    [yes_radio, no_radio],
    labels=["In a rated radio metro", "No rated radio market"],
    colors=["#1F4E89", "#8B6914"],
    autopct="%1.1f%%",
    startangle=90,
    textprops={"fontsize": 10},
    wedgeprops={"edgecolor": "white", "linewidth": 2},
)
ax.set_title(f"GA voter universe by radio reach\n"
             f"The 'no radio' bucket is {no_radio_r_pct:.0f}% R-leaning",
             fontsize=11, loc="left")

plt.suptitle("Radio buys reach the urban core; rural Collins voters need other channels",
             fontsize=12, y=1.02)
plt.tight_layout()
plt.savefig(HERE / "radio_coverage_gap.png", dpi=200, bbox_inches="tight")
plt.close()
print("Wrote: radio_coverage_gap.png")

"""
The radio chart is doing a specific job for the campaign plan:

The pie shows that ~25% of GA voters live in counties with no rated radio
market. The chart subtitle adds the kicker - those 25% are ~68% R-leaning
on partisan blend. In other words, the radio-invisible portion of the state
is disproportionately Collins's base.

So the strategic takeaway is: do not buy radio expecting to reach rural
voters - the rated radio markets actually skew toward where Ossoff's
voters live (urban Atlanta, urban Augusta, urban Columbus). For rural R
mobilization use mail / digital / local field / earned media instead.
"""

print("\nAll charts saved.")
