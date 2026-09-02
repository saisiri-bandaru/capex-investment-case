# Capex investment case

Capex request pack for a fictional CPG company (**Northline Consumer Products**): project cash flows, NPV, IRR, payback, and a simple capacity / ROI story for a Midwest high-speed filling line.

**File to open:** `Northline_Capex_Investment_Case.xlsx`

## What you will see

- Capex timing, useful life, salvage, WACC, tax, volume lift, conversion savings, and ramps
- An 8-year unlevered cash-flow build (contribution on new volume, conversion savings on the base run, D&A, tax, NWC, salvage)
- NPV, IRR, undiscounted and discounted payback, profitability index, simple ROI vs hurdles
- A WACC × volume-lift NPV grid and a tornado on the main drivers
- A one-pager: ask, returns, capacity (utilization today vs with the line), recommendation

Yellow cells with blue font are inputs. Black font is formulas.

## How to use

1. Open `01_Assumptions` and change the yellow cells (capex, volume lift, savings, WACC, tax, life, ramps).
2. Read the unlevered cash-flow build on `02_Cash_Flows`.
3. Read NPV, IRR, payback, and hurdle tests on `03_Returns`.
4. Stress WACC × volume lift (and the tornado) on `04_Sensitivity`.
5. Use `05_One_Pager` as the capital-committee page.

## Tabs

| Tab | Role |
| --- | --- |
| `00_Cover` | Purpose and how to use |
| `01_Assumptions` | Capex timing, savings, volume lift, WACC, tax, life |
| `02_Cash_Flows` | Unlevered project cash flows, Year 0–8 |
| `03_Returns` | NPV, IRR, payback, PI, hurdles |
| `04_Sensitivity` | WACC × volume-lift grid and tornado |
| `05_One_Pager` | Committee page: ask, returns, capacity, recommendation |
| `06_Data_Dictionary` | Field definitions |

## Stack

Excel (formulas only — no VBA). Built so another analyst can inherit the file from the data dictionary. Amounts in $000s.

## Not included on purpose

- Live capex-system or SAP project pulls
- Confidential employer data
- A full three-statement LBO or debt schedule (this is the *project returns* pack)

All sample numbers are fictional.

## Profile

Sai Siri Bandaru — Financial Analyst | FP&A | forecasting, variance analysis, Excel
