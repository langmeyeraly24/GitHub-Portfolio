# Media Bias Detection — NLP Classification

Final project for **CSC680: Data Mining** at American University, Spring 2026.

This project builds a text classification model to detect political media bias across news sources using natural language processing techniques.

---

## Problem

News articles vary significantly in how they frame political events. This project explores whether machine learning can reliably classify an article's political lean based on its text content alone — without relying on the publication name.

---

## Approach

1. **Data:** Cleaned dataset of labeled news articles (`data/dataset_cleaned.parquet`)
2. **Text preprocessing:** Tokenization, stopword removal, TF-IDF vectorization
3. **Models:** Compared multiple classifiers (see notebook for full results)
4. **Evaluation:** Accuracy, precision/recall, confusion matrix

---

## Results

Key visualizations in `figures/`:

- `results_fig1.png` — Class distribution
- `results_fig2.png` — Model performance comparison
- `results_fig3.png` — Confusion matrix
- `results_fig4.png` — Feature importance / top terms by class
- `results_fig5.png` — Precision-recall breakdown

---

## Files

```
├── media_bias_analysis.ipynb   # Full analysis notebook
├── data/
│   └── dataset_cleaned.parquet
└── figures/
    └── results_fig1–5.png
```

---

## Tools

Python · scikit-learn · pandas · matplotlib · NLP (TF-IDF)
