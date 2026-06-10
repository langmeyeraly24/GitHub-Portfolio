# GA Senate 2026 — Voter Targeting & Media Market Strategy

Data science work completed as part of the **Campaign Management Institute (GOVT520)** at American University, May 2026. This project builds a county-level targeting model for a hypothetical Georgia U.S. Senate race, integrating historical election results, voter file data, and media market cost data.

---

## What This Does

Georgia has 159 counties and a competitive Senate environment. The goal was to answer:
- **Where are the persuadable voters?** (split-ticket, low-partisan-lean counties)
- **Where is GOTV the highest-leverage investment?** (strong partisans with low historical turnout)
- **How should limited media budget be allocated across TV/radio markets?**

The output is a **159-county master dataset** with 47 columns covering vote history, partisan scoring, turnout tiers, and budget recommendations — all reproducible from the notebook.

---

## Key Findings

- Georgia is "purple but mostly red": 94 of 159 counties are Strong R (59%) but hold only 34% of voters. Lean+Strong D = 22 counties (14%) but **51% of voters**.
- **Suburban Atlanta is the persuasion universe.** Top Kemp-Walker gap counties (Oconee, Forsyth, Cobb, Cherokee, Hall, Fayette) hold ~1.18M active voters — if 6–7% are persuadable, that's 70–80K winnable votes. Walker lost the 2022 Senate race by ~37K.
- **Highest-leverage GOTV cell: Strong R + Low turnout** — 26 counties of existing R voters who aren't showing up. Cheapest per-vote mobilization.
- Method 2 projections show metro D-base counties fading slightly in turnout while exurbs (Jackson +25%, Cherokee +12%, Hall +13%) are surging.

---

## Data Sources

All election data is publicly available:

| Source | Data |
|---|---|
| [OpenElections Georgia](https://github.com/openelections/openelections-data-ga) | 2020, 2021, 2022, 2024 precinct and county results |
| [Georgia Secretary of State](https://sos.ga.gov/page/voter-registration-statistics) | Active/inactive voter registration by county |
| [U.S. Census TIGER/Line](https://www.census.gov/geographies/mapping-files/time-series/geo/cartographic-boundary.html) | County shapefiles |
| NYT / 538 | 2026 Senate polling |

---

## Repository Structure

```
├── targeting_analysis.ipynb           # Main EDA and targeting analysis
├── ga_senate_2026_county_master.xlsx  # Output: 159 counties × 47 columns
├── ga_county_media_markets.csv        # County → TV/radio market crosswalk
├── media_market_analysis.py           # TV market cost and reach analysis
├── media_market_visualizations.py     # Visualization scripts
├── targeting_visuals.py               # Targeting chart generation
├── ga_field_office_map.py             # Field office placement map
├── ga_field_office_map.png            # Field office map output
├── images/targeting_viz_package/      # 11 charts used in final campaign plan
└── data/
    ├── election-results/              # Raw OpenElections + SOS files
    ├── polling/                       # 2026 Senate polling data
    └── Online Fundraising - *.xlsx    # Fundraising projections and budget models
```

---

## Visualizations

The `images/targeting_viz_package/` folder contains 11 charts from the final campaign plan:

1. Path to victory map
2. Lean bucket overview (county distribution)
3. Kemp-Walker persuasion gap
4. Targeting matrix (lean × turnout)
5. Top 10 priority counties
6. Projection method comparison
7. Demographic composition
8. TV market net partisan margin
9. TV market priority tiers
10. TV strategy quadrant
11. Radio coverage gap analysis

---

## Tools

- **Python:** pandas, geopandas, matplotlib, openpyxl
- **R:** county map visualization (ggplot2, sf)
- **Quarto:** voter targeting writeup

---

## Context

This work was completed for the **Campaign Management Institute**, a DCCC-affiliated program at American University that trains students in professional campaign strategy. The analytical framework — partisan blend scoring, lean bucketing, turnout tier modeling — reflects standard targeting methodology used in competitive Senate races.
