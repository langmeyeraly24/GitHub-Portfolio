# GitHub Portfolio Strategy — Alyssa Langmeyer

_Built June 2026. Goal: look like a working data scientist, not a student._

---

## The Problem with Your Current GitHub

- 1 public repo (the `.github.io` site), sparse contributions
- No profile README — blank first impression
- Work exists but is invisible

## The Fix: 3 repos + a profile README

---

## Step 1: Profile README (do first)

**Repo name:** `langmeyeraly24/langmeyeraly24` (same as your username — GitHub auto-renders this as your profile page)

**File:** `README.md` → see `profile-README.md` in this folder

**How to set it up:**
1. Go to github.com → New repository
2. Name it exactly `langmeyeraly24`
3. Check "Add a README file"
4. Paste the content from `profile-README.md`
5. Update the repo links once you've created the other repos below

---

## Step 2: Repos to Create

### Repo 1: `ga-senate-2026-targeting` ⭐ (your best work — pin this)

**What to include:**
- `CMI Campaign Plan/georgia_targeting_eda.ipynb` → rename to `targeting_analysis.ipynb`
- `CMI Campaign Plan/ga_senate_2026_county_master.xlsx`
- `CMI Campaign Plan/ga_county_media_markets.csv`
- `CMI Campaign Plan/media_market_analysis.py`
- `CMI Campaign Plan/media_market_visualizations.py`
- `CMI Campaign Plan/targeting_visuals.py`
- `CMI Campaign Plan/ga_field_office_map.py`
- `CMI Campaign Plan/images/targeting_viz_package/` (all 11 charts)
- `CMI Campaign Plan/ga_field_office_map.png`
- `Data/georgia voter + election data/` (all election CSVs)
- `Data/Online Fundraising - *.xlsx` (the 3 fundraising files)
- `README.md` → use the content from `campaign-management-README.md`

**What NOT to include:**
- `Handoff + reference docs/HANDOFF_TO_CLAUDE_CODE.md` (internal notes)
- `.claude/` folder
- `__pycache__/`
- `shapefile_cache/` (big binary files, not useful on GitHub)
- Cover letters, networking CSVs, thank you notes
- Syllabi

**Add a `.gitignore`:**
```
__pycache__/
*.pyc
shapefile_cache/
.claude/
*.docx
*.pdf
~$*
```

---

### Repo 2: `media-bias-detection`

**What to include:**
- `CSC680 - Data Mining/final project/media_bias_analysis.ipynb`
- Any data files used (check the notebook for what CSV/data it loads)
- A README explaining the task, approach, and results

**Note:** This is a class final project — frame it clearly as such in the README. That's fine and normal.

---

### Repo 3: `ms-data-science-coursework`

**Purpose:** Show breadth. One notebook per course, only your finished work, no starter files.

**What to include (suggested selection):**

| Folder | What to include | Why |
|---|---|---|
| CSC680 - Data Mining | Finished notebooks: decision trees, ensemble, evaluation, SVM, KNN, linear regression | Core ML workflow in Python |
| GOVT650 - Political Analysis | `Classwork/GOVT_650_001_Introduction.ipynb` | Political data work |
| STAT615 - Regression | HW1–HW6 Rmd files | Statistical modeling in R |
| STAT627 - Statistical Machine Learning | HW3–HW6 Rmd files | KNN, logistic, CV, bootstrap in R |

**Folder structure in the repo:**
```
ms-data-science-coursework/
├── README.md
├── data-mining/                (CSC680)
├── political-analysis/         (GOVT650)
├── regression/                 (STAT615)
└── statistical-machine-learning/ (STAT627)
```

**README should include:** program description, list of courses, what each folder contains. Frame as: "Selected coursework from my M.S. in Data Science for Applied Public Policy at American University (2025–2026)."

---

## Step 3: What to Skip Entirely

| What | Why |
|---|---|
| DATA793 - AI Practicum | Group work, not yours |
| Any notebook with "Brandt" in filename | Not your work |
| DATA612 - R Programming (most of it) | R homework, lower signal |
| Concert Archive Analysis | You contributed but data work isn't yours — skip or add later with clear attribution |
| Class starter/template files | They're the professor's, not yours |
| `mess/` folders, `downloaded files from canvas/` | Noise |
| Cover letters, networking docs | Wrong venue |

---

## Step 4: Pinned Repos Order

On your profile, pin in this order:
1. `ga-senate-2026-targeting` (best project)
2. `media-bias-detection`
3. `ms-data-science-coursework`
4. `langmeyeraly24.github.io` (your existing site)

---

## Step 5: Contribution Activity

Once you push these repos, your contribution graph will fill in. A few extra things that help:
- Add a **description and topic tags** to each repo (GitHub lets you add tags like `machine-learning`, `political-data`, `python`, `nlp`)
- Link your GitHub on your LinkedIn
- As you finish coursework this fall, push final projects directly to your coursework repo — every commit counts toward the green squares

---

## Notes on Academic Work

It's completely normal and expected for early-career data scientists to have coursework on GitHub. The key framing: label it clearly as coursework, only show your *finished* work (not starter files), and let the quality speak for itself. The campaign management project shows real applied work — that's the anchor. Everything else builds context around it.
