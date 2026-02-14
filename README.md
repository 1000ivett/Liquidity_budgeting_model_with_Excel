# Real Estate Liquidity Budgeting & Analysis Tool

## Project Overview
Developed a sophisticated financial modelling tool for a real estate consulting case to manage the working capital of a multi-unit property portfolio. The tool enables managers to forecast cash flows, determine necessary capital injections, and analyse the financial impact of operational delays or maintenance costs. Detailed instructions for the task are included in the pdf.

## Key Features
* **Dynamic Cash Flow Forecasting:** A 12-month liquidity model that adjusts automatically based on rent, renovation timelines, and variable maintenance schedules.
* **Automated Breakeven Analysis:** Integrated a VBA macro to calculate the precise rental increase required to achieve zero-capital-injection status, featuring an automatic "return-to-default" state for user convenience.
* **Multi-Variable Sensitivity Analysis:** Built 2D Data Tables to visualise the relationship between maintenance timing (e.g., refrigerator replacement) and rental growth.
* **Scenario Management:** * Utilized **Excel Scenario Manager** for high-level strategic shifts.
    * Built a **Manual Scenario Tool** to stress-test renovation durations against rental delays.
* **Professional Reporting:** One-click PDF report generation via VBA, featuring a combined chart of pre-and-post-injection liquidity.

## Technical Stack
* **Excel:** Data Validation (Drop-downs), Advanced Logic (`IF`, `MOD`, `LOOKUP`), Data Tables.
* **VBA (Macros):** Automated Goal Seek, PDF Exporting, and Input Resetting.
* **Financial Concepts:** Working Capital Management, Liquidity Budgeting, Breakeven Analysis.

## How it Works
1.  **Inputs:** Users define monthly rents, loan repayments, and utility costs via a user-friendly dashboard.
2.  **Timing:** Renovation and maintenance costs are assigned to specific months using dynamic drop-down lists.
3.  **Calculations:** The model calculates monthly balances and flags months where liquidity falls below zero.
4.  **Optimisation:** The owners can see exactly how many blocks of €1,000 must be injected to maintain positive cash flow.

## Visuals
**Model inputs, outputs, visuals, analyses**

<img width="2385" height="990" alt="image" src="https://github.com/user-attachments/assets/4e8e5d68-ef63-49c1-bef3-5337530a9f67" />

**Outputs printed with Macro**
<img width="1101" height="1132" alt="image" src="https://github.com/user-attachments/assets/d95b3867-ba81-46cf-88c8-b6fb2e09d6e8" />

---
*Note: This project was developed as part of the Investment and Business Analytics with Excel Course.*
