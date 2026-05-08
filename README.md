# Money, Science, and the Clean Energy Transition

> How Technology Maturity Shapes the Feedback Between Public R&D
> Investment and Research Output

**Author:** Ilayda Kucukafacan
**Affiliation:** Koç University, M.Sc. Computational Social Sciences
**Course:** Data Science for Sustainable Development Goals (Spring 2026)
**SDG Link:** SDG 7 (Affordable and Clean Energy) & SDG 13 (Climate Action)

## Abstract

This project links public R&D spending (IEA Energy Technology RD&D Budgets, 1974–2023) to research output (OpenAlex publication counts, queried by tech-specific keyword) across 34 OECD-style countries and 9 energy technologies. Using first-differenced panel Granger causality with country fixed effects, the spending↔publications relationship is statistically robust only for **Hydrogen & fuel cells**, surviving leave-one-out drops, GDP normalization, structural-break splits at 2015, HAC standard errors, varying lag lengths (K=2,3,5), and a 500-iteration year-shuffle placebo. Mature technologies (Nuclear, Hydropower) show no robust link, while growing technologies (Wind, Ocean, CO2 capture) show one-way spending→publications causality. Research efficiency varies 60× across technologies; portfolios have de-concentrated dramatically since the 1970s; and a carbon-management spillover cluster connects Hydrogen, CO2 capture, and Biofuels.

## Key Findings

1. **Maturity gradient in causal direction:** Hydrogen shows robust
   bidirectional Granger causality (8/8 stress tests passed). Growing
   technologies show one-way spending→publications. Mature technologies
   show no robust link.
2. **60× efficiency gap:** Hydro produces 24 pubs/M$ vs nuclear at 0.4.
3. **Portfolio rotation:** Every major OECD country diversified away from
   nuclear since 1980. Specialization within a fixed budget slightly
   outperforms diversification.
4. **Carbon-management spillover cluster:** Hydrogen spending predicts
   CO2 capture publications (F=20.9, p<1e-8). Spending on one technology
   in the cluster partially funds research in the others.

## Repository Structure

```
DS4SDGs-Project/
├── README.md
├── LICENSE
├── requirements.txt
├── .gitignore
├── data/
│   ├── raw/                    # IEA source files (not tracked — see Data Access)
│   └── processed/              # Merged panels ready for analysis
│       ├── rdd_public_panel.csv
│       ├── openalex_full_panel.csv
│       ├── openalex_pilot.csv
│       └── merged_panel.csv
├── notebooks/
│   ├── 01_data_cleaning.ipynb
│   ├── 02_openalex_pull.ipynb
│   ├── 03_exploratory_analysis.ipynb
│   ├── 04_granger_causality.ipynb
│   ├── 05_robustness.ipynb
│   ├── 06_efficiency_and_scale.ipynb
│   ├── 07_portfolio_analysis.ipynb
│   └── 08_crosslinks.ipynb
├── figures/                    # All output figures (19 PNGs)
├── results/                    # All output CSVs (9 files)

```

## Data Access

The IEA Energy Technology RD&D Budgets dataset requires free registration
at [IEA Data Explorer](https://www.iea.org/data-and-statistics/data-tools/energy-technology-rdd-budgets-data-explorer).
Download the RDDPUBLIC series and place it in `data/raw/`.

OpenAlex data is pulled via the free public API — no key needed.
Notebook 02 handles the full pull.

## Methodology

The pipeline merges the IEA RDDPUBLIC panel (filtered to USD constant PPP, Total public, RD&D, non-confidential, with IEA aggregates removed) with OpenAlex publication counts queried by tech-specific keyword sets at the country-year level. Causal inference uses panel Granger regressions on first-differenced series with country fixed effects across 9 technologies and 34 OECD countries (1974–2023). A robustness battery covers lag variation (K=2,3,5), leave-one-out country sensitivity, GDP normalization, pre/post-Paris splits, HAC standard errors, and 500-iteration year-shuffle placebos. Efficiency is measured as publications per million USD spent; portfolio concentration uses the Herfindahl–Hirschman Index (HHI). Cross-technology spillovers are tested with the same Granger framework on candidate pairs (Hydrogen→CO2, CO2→Biofuels, Solar→Wind, Nuclear→Hydrogen).

## How to Reproduce

1. Clone the repo
2. Download IEA RDDPUBLIC data and place in `data/raw/`
3. Install dependencies: `pip install -r requirements.txt`
4. Run notebooks in order (01 through 08)
5. All figures and results CSVs will be regenerated in `figures/` and `results/`

## Citation

If you use this work, please cite:

> Kucukafacan, I. (2026). *Money, Science, and the Clean Energy Transition:
> How Technology Maturity Shapes the Feedback Between Public R&D Investment
> and Research Output.* Koç University, Data Science for SDGs.

## License

MIT
