"""
media_market_analysis.py
-------------------------
Take the master county table, attach TV/radio market columns using the
crosswalk in ga_media_markets.py, and build market-level summary tables +
choropleth visualizations for the campaign plan.

Author: Alyssa Langmeyer  (GOVT 520 - Mike Collins for U.S. Senate 2026)
"""

from pathlib import Path
import pandas as pd
import numpy as np

from ga_media_markets import attach_markets, TV_MARKETS, RADIO_MARKETS


# ---------------------------------------------------------------------------
# 1. Load master county table and attach market columns
# ---------------------------------------------------------------------------

HERE = Path(__file__).parent
MASTER_XLSX = HERE / "ga_senate_2026_county_master.xlsx"

master = pd.read_excel(MASTER_XLSX, sheet_name="county_data")
print(f"Loaded master table: {master.shape[0]} counties, "
      f"{master.shape[1]} columns")

merged = attach_markets(master, county_col="county")

# Sanity: every county should now have a TV market assigned
assert merged["tv_market"].notna().all(), \
    "Some counties failed to merge to a TV DMA - check spelling in master table"
print(f"Merged: every county has tv_market + radio_market set")

"""
What this step does:

I'm loading my master county table (the one with 57 columns of vote share,
turnout, demographics, etc.) and using the attach_markets() function from
ga_media_markets.py to bolt on two new columns: tv_market and radio_market.

The assertion is a guardrail. If any county failed to merge - usually because
of a spelling mismatch like "Dekalb" vs "DeKalb" or "Mc Duffie" vs "McDuffie"
- the script will halt loudly instead of silently dropping that row from
downstream summaries. Better to fail here than to find a hole in the analysis
two days from now.
"""


# ---------------------------------------------------------------------------
# 2. Market-level summary: TV
# ---------------------------------------------------------------------------

# Weighted averages need a helper because pandas .mean() ignores weights
def wavg(values, weights):
    """Weighted average that ignores NaN pairs."""
    v = pd.Series(values).astype(float)
    w = pd.Series(weights).astype(float)
    mask = v.notna() & w.notna() & (w > 0)
    if not mask.any():
        return np.nan
    return float((v[mask] * w[mask]).sum() / w[mask].sum())


def market_summary(df: pd.DataFrame, market_col: str) -> pd.DataFrame:
    """Roll county rows up to market rows with the metrics that matter."""
    rows = []
    for market, sub in df.groupby(market_col, sort=False):
        # vote totals
        pres20_tot = sub["pres20_total_votes"].sum()
        pres24_tot = sub["pres24_total_votes"].sum()
        gov22_tot  = sub["gov22_total_votes"].sum()
        reg_2026   = sub["reg_total_2026"].sum()

        # weighted vote shares (weighted by that race's total votes in each county)
        trump20_pct = wavg(sub["pres20_pct_r"], sub["pres20_total_votes"])
        trump24_pct = wavg(sub["pres24_pct_r"], sub["pres24_total_votes"])
        kemp22_pct  = wavg(sub["gov22_pct_r"],  sub["gov22_total_votes"])
        walker22_pct = wavg(sub["sen22_pct_r"], sub["sen22_total_votes"])
        blend_r     = wavg(sub["partisan_blend_r"], sub["reg_total_2026"])

        # estimated R votes available in 2026 (registered voters * blended R%)
        est_r_votes_2026 = (
            sub["reg_total_2026"] * sub["partisan_blend_r"] / 100.0
        ).sum()
        est_d_votes_2026 = (
            sub["reg_total_2026"] * (100 - sub["partisan_blend_r"]) / 100.0
        ).sum()

        # how the lean buckets distribute inside this market
        bucket_counts = sub["lean_bucket"].value_counts().to_dict()

        rows.append({
            "market": market,
            "counties": len(sub),
            "reg_voters_2026": int(reg_2026),
            "pres24_total_votes": int(pres24_tot),
            "trump_2020_pct": round(trump20_pct, 2),
            "trump_2024_pct": round(trump24_pct, 2),
            "kemp_2022_pct": round(kemp22_pct, 2),
            "walker_2022_pct": round(walker22_pct, 2),
            "partisan_blend_r": round(blend_r, 2),
            "est_r_votes_2026": int(est_r_votes_2026),
            "est_d_votes_2026": int(est_d_votes_2026),
            "est_r_margin_2026": int(est_r_votes_2026 - est_d_votes_2026),
            "strong_r_counties": bucket_counts.get("Strong R", 0),
            "lean_r_counties":   bucket_counts.get("Lean R", 0),
            "competitive_counties": bucket_counts.get("Competitive", 0),
            "lean_d_counties":   bucket_counts.get("Lean D", 0),
            "strong_d_counties": bucket_counts.get("Strong D", 0),
        })
    out = pd.DataFrame(rows).sort_values("reg_voters_2026", ascending=False)
    out["share_of_statewide_reg"] = (
        out["reg_voters_2026"] / out["reg_voters_2026"].sum() * 100
    ).round(2)
    return out.reset_index(drop=True)


tv_summary = market_summary(merged, "tv_market")
radio_summary = market_summary(merged, "radio_market")

print("\n=== TV market summary ===")
print(tv_summary.to_string(index=False))
print("\n=== Radio market summary ===")
print(radio_summary.to_string(index=False))

"""
What I'm doing here, in plain English:

I'm rolling 159 county rows up into ~11 TV market rows (and ~10 radio market
rows including the 'no rated radio' bucket). For each market I'm computing:

- How many counties are in this market
- Total registered voters in 2026 (the universe size)
- Trump 2020 %, Trump 2024 %, Kemp 2022 %, Walker 2022 % - weighted by each
  race's total votes (NOT a simple average across counties, because Fulton
  with 800K votes should count way more than Webster with 1.5K votes)
- The blended-R partisan score from my master table, weighted by registered
  voters
- An estimate of how many R votes vs D votes are likely available in 2026
  based on the partisan blend applied to the 2026 registered voter file
- The R margin available in that market in raw votes
- A breakdown of how many Strong R / Lean R / Competitive / Lean D / Strong D
  counties are in each market

The "share_of_statewide_reg" column tells me what % of GA registered voters
sit in each market. Atlanta will dominate - probably 55-60% of the state's
voters live in that DMA. That number alone is the single most important fact
for the budget: it's why Collins can't ignore Atlanta even though he loses it.

Why weighted averages: A market's "Trump %" should reflect what would happen
if you actually pooled all the votes in that market, not the unweighted
average of county shares. If I just did .mean() on pres20_pct_r, tiny rural
counties would have the same weight as Fulton, and the number would look way
more Republican than the market actually votes. So wavg() weights each
county's vote share by that county's actual vote total in that race.
"""


# ---------------------------------------------------------------------------
# 3. Targeting recommendations per TV market
# ---------------------------------------------------------------------------

def media_strategy(row):
    """
    Turn each market's profile into a campaign-plan-ready strategy label.
    These categories drive how we allocate paid media spend.
    """
    blend = row["partisan_blend_r"]
    margin = row["est_r_margin_2026"]
    share = row["share_of_statewide_reg"]

    if share < 1:
        return "Minimal spend (tiny market, low ROI)"
    if blend >= 65:
        return "RUN UP MARGINS - mobilize base, heavy GOTV"
    if blend >= 55:
        return "PROTECT LEAD - mobilize + light persuasion"
    if 45 <= blend < 55:
        return "PERSUASION BATTLEGROUND - heaviest spend"
    if blend < 45 and share >= 5:
        return "LIMIT LOSSES - reach soft D / independents only"
    return "Low-priority defensive spend"


tv_summary["media_strategy"] = tv_summary.apply(media_strategy, axis=1)

# print the strategy view
print("\n=== TV market - media strategy view ===")
strategy_view = tv_summary[[
    "market", "counties", "reg_voters_2026", "share_of_statewide_reg",
    "partisan_blend_r", "est_r_margin_2026", "media_strategy"
]]
print(strategy_view.to_string(index=False))

"""
The media_strategy() function is my first cut at translating data into a
campaign decision. The bucketing isn't perfect - this is a sketch the team
can refine in our strategy meetings - but it gets us a defensible starting
position for the budget conversation:

  - RUN UP MARGINS (blend >= 65%): Chattanooga, Greenville spillover, etc.
    These are deep-red rural DMAs. Collins doesn't need to persuade anyone
    here, he just needs to TURN OUT the base. Spend on GOTV, not on TV ads
    with persuasion messaging.

  - PROTECT LEAD (55-65%): Mid-red DMAs. Mostly mobilization, some light
    persuasion at the margins for soft Rs and conservative-leaning indies.

  - PERSUASION BATTLEGROUND (45-55%): True swing territory. This is where
    persuasion ad spend has the highest expected return per dollar.

  - LIMIT LOSSES (<45% AND large): Atlanta is the obvious one. Collins
    will lose this DMA, but he can't afford to lose it 70-30. Heavy spend
    aimed at suburban Atlanta women, indies, and soft Ds is the play.

  - Minimal spend (tiny markets): Dothan only covers 2 GA counties.
    Buying Dothan TV to reach 2 GA counties is wildly inefficient. Better
    to hit Early & Miller via mail or digital.

I want to be transparent about a limitation: the partisan_blend_r in my
master table is the BASELINE expected R share absent persuasion. Real
campaigns also factor in things like recent polling movement, candidate-
specific favorability, and ad market cost per point (CPP). I don't have
CPP data yet - that's a follow-up if we need to refine the spend mix.
"""


# ---------------------------------------------------------------------------
# 4. Save the merged county table + summaries to Excel
# ---------------------------------------------------------------------------

OUT_XLSX = HERE / "ga_county_master_WITH_MEDIA_MARKETS.xlsx"

with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
    merged.to_excel(writer, sheet_name="county_data_with_markets", index=False)
    tv_summary.to_excel(writer, sheet_name="tv_market_summary", index=False)
    radio_summary.to_excel(writer, sheet_name="radio_market_summary", index=False)

    # county -> markets crosswalk as its own sheet for easy reference
    crosswalk = merged[["county", "tv_market", "radio_market"]].copy()
    crosswalk.to_excel(writer, sheet_name="county_market_crosswalk", index=False)

# also save the crosswalk as a standalone CSV
CROSSWALK_CSV = HERE / "ga_county_media_markets.csv"
crosswalk.to_csv(CROSSWALK_CSV, index=False)

print(f"\nWrote: {OUT_XLSX.name}")
print(f"Wrote: {CROSSWALK_CSV.name}")
