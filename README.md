# Capex investment case

Unlevered investment case for a fictional high-speed RTD filling line at **Northline Consumer Products** (Plant 2, Line 4): capex timing, volume lift × contribution, cost-out savings, NPV, IRR, and payback.

**File to open:** `Northline_Capex_Investment_Case.xlsx`

## What you will see

- Capex timed over two years (equipment, installation, spares), then sustaining maintenance
- Incremental volume ramp × contribution per case, plus labor / scrap / energy / avoided-maintenance savings
- Unlevered cash flow: NOPAT + D&A − capex − ΔNWC + salvage
- NPV at WACC, IRR, interpolated payback (undiscounted and discounted), profitability index
- A Base / Upside / Downside toggle, a WACC × volume grid, and a tornado
- An investment-committee one-pager with a go / review rule

Yellow cells with blue font are inputs. Black font is formulas.

## How to use

1. Open `01_Assumptions` and change the yellow cells (scenario, WACC, tax, life, capex timing, volume ramp, unit economics, savings).
2. Read the unlevered cash-flow build on `02_Cash_Flows`.
3. Read NPV, IRR, payback, and PI on `03_Returns`.
4. Stress WACC and volume on `04_Sensitivity`.
5. Use `05_One_Pager` as the investment-committee snapshot.

## Tabs

| Tab | Role |
| --- | --- |
| `00_Cover` | Purpose and how to use |
| `01_Assumptions` | Scenario, capex timing, volume / savings ramps, WACC, tax, life |
| `02_Cash_Flows` | Incremental P&L → unlevered FCF, NWC, salvage |
| `03_Returns` | NPV, IRR, payback, PI, recommendation |
| `04_Sensitivity` | WACC one-way, volume two-way, tornado |
| `05_One_Pager` | Investment-committee snapshot |
| `06_Data_Dictionary` | Field definitions |

## Stack

Excel (formulas only — no VBA). Built so another analyst can inherit the file from the data dictionary. Amounts in $000s. Volume in 000 cases.

## Not included on purpose

- Live capex-system or ERP feeds
- Confidential employer data
- Debt sizing, lease vs buy, or a full three-statement LBO (this is the *unlevered project case*)

All sample numbers are fictional.

## Profile

Sai Siri Bandaru — Financial Analyst | FP&A | forecasting, variance analysis, Excel
