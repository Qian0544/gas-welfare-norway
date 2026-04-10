# Who Gains When Gas Prices Surge?

**A Welfare Analysis of Norway's Gas Economy Under the Ukraine War Supply Shock**

EBA3650 Quantitative Economics | BI Norwegian Business School | Spring 2026

---

## Project Structure

| File | Description |
|------|-------------|
| `gas_welfare_analysis.ipynb` | **Main deliverable** -- full analysis notebook with code, plots, and narrative |
| `exam_proposal.pdf` | Project proposal (submitted earlier) |
| `generate_proposal.py` | Script that generates the proposal PDF |
| `build_notebook.py` | Script that generates the notebook from scratch |

## How to Run

```bash
# Open the notebook
jupyter notebook gas_welfare_analysis.ipynb

# Or regenerate from scratch
python3 build_notebook.py
jupyter notebook gas_welfare_analysis.ipynb
```

### Requirements

- Python 3.8+
- numpy
- scipy
- matplotlib

```bash
pip install numpy scipy matplotlib
```

## Division of Work

| Member | Section | What to Own |
|--------|---------|-------------|
| 1 | Section 3 (Partial Equilibrium Model) | Gas market supply/demand, equilibrium solving, producer surplus |
| 2 | Section 4 (Consumer Welfare) | Quasi-linear utility, compensating variation, household welfare loss |
| 3 | Section 5 (Petroleum Tax) | Pigou/Ramsey framing, tax decomposition, per-unit tax comparison |
| 4 | Sections 1-2, 6-7 (Integration) | Intro, background, sensitivity analysis, conclusion, presentation |

## Key Results (Baseline)

- **Pre-war gas price:** 20 EUR/MWh --> **Post-shock:** 43.89 EUR/MWh (+119%)
- **Norwegian PS gain:** +28.0 bn EUR
- **Government revenue gain (78% tax):** +21.8 bn EUR
- **Consumer welfare loss:** -5.0 bn EUR
- **Net welfare:** +23.0 bn EUR -- **Norway is a net winner**
