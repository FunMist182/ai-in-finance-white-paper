# AI in Finance: Accounting Risk and Valuation Using Python

**David Hayes | King’s College London | BSc Accounting & Finance**

This independent research project explores how Python can support financial analysis while examining where professional judgement remains essential.

The project combines financial-statement analysis, valuation and model governance across three companies: Tesco, Carillion and London Stock Exchange Group (LSEG).

## Project Overview

### 1. Tesco and Carillion — Financial Reporting Risk

A Python-based financial-statement screen is applied consistently to:

- Tesco, 2010–2015
- Carillion, 2010–2016

The analysis examines working-capital movements, accruals and profit-to-cash conversion.

The comparison is designed not only to identify warning signals, but also to demonstrate the limitations of fixed-rule financial screening. The Carillion analysis produces recurring cash-conversion and working-capital prompts, while the Tesco case demonstrates how a quantitative screen can fail to identify a specific accounting issue.

### 2. LSEG — Valuation and Investment Judgement

A five-year discounted cash flow valuation of London Stock Exchange Group was built in Python.

The model includes:

- Downside, base and upside scenarios
- WACC and terminal-growth sensitivity analysis
- Enterprise-to-equity value reconciliation
- Explicit valuation assumptions and source controls

All three scenario outputs and 50 sensitivity calculations were independently reconciled as part of the model-validation process.

### 3. LSEG — Model Governance

The final case study examines how much confidence should be placed in a financial model rather than simply whether its calculations run correctly.

It considers:

- Calculation checks
- Source and date controls
- Sensitivity to key assumptions
- Terminal-value dependence
- Unresolved assumptions
- The role of professional judgement

## Repository Contents

- **AI_in_Finance_White_Paper.pdf** — Full research paper and recommended starting point
- **Code/** — Python scripts used for the financial-statement analysis and valuation work
- **Tesco Data/** — Source data supporting the Tesco financial-reporting analysis
- **Carillion Data/** — Source data supporting the Carillion financial-reporting analysis
- **Results/** — Outputs generated from the analysis and valuation work
- **requirements.txt** — Python dependencies required to reproduce the analysis

## Suggested Route Through the Project

For a quick review, start with the **white paper** for the conclusions and methodology, then examine **Code/** and **Results/** for the underlying analytical work and reproducibility.

## Use of AI

The Python analysis was developed for this project. AI was used as a checking and review tool for code and analysis rather than as a substitute for the underlying financial analysis.

The numerical evidence comes from the Python scripts, company data and recorded calculation checks.

## Key Takeaway

The project demonstrates that automation can make financial analysis more consistent and repeatable, but reliable investment analysis still depends on understanding the underlying business, challenging assumptions and recognising what a model cannot establish.

---

*Independent research project, 2026.*
