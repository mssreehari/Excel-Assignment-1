# Excel Data Analysis Assignment-1

## Overview

This assignment analyzes a product dataset using basic Excel functions. The dataset contains Product ID, Product Name, Brand Name, Price, Quantity, and Category.

## Analysis Performed

- Calculated the total price of all products using `SUM`.
- Counted the number of products using `COUNT`.
- Calculated the average product price using `AVERAGE`.
- Found the minimum and maximum prices using `MIN` and `MAX`.
- Created a **Price Range** column using `IF`.
- Calculated the total price of Electronics products using `SUMIF`.
- Counted products priced below $100 using `COUNTIF`.
- Extracted Day, Country Code, and Month using `LEFT`, `RIGHT`, and `MID`.

## Excel Functions Used

| Function | Purpose |
|---|---|
| `SUM` | Adds values in a range |
| `COUNT` | Counts numeric values |
| `AVERAGE` | Calculates the average |
| `MIN` | Finds the minimum value |
| `MAX` | Finds the maximum value |
| `IF` | Applies a condition |
| `SUMIF` | Adds values based on a condition |
| `COUNTIF` | Counts values based on a condition |
| `LEFT` | Extracts characters from the left |
| `RIGHT` | Extracts characters from the right |
| `MID` | Extracts characters from the middle |

## Formulas Used

### Total Price

```excel
=SUM(D2:D35)
