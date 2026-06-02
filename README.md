# 🧹 Project 1: Data Cleaning & Preparation
**DecodeLabs Industrial Training | Batch 2026**
**Author:** Shaik | Mohan Babu University (22IT101001)

---

## 📌 Problem Statement

Raw business datasets are rarely clean. Before any analysis or model can be trusted, the underlying data must be audited and corrected. This project takes a raw 1,200-row sales dataset and transforms it into a production-ready, gold-standard source of truth.

> *"Your analysis is only as good as your data."*

---

## 🎯 Goal

Clean a raw dataset by handling missing values, duplicates, and incorrect data formats — proving zero error rate on unique identifiers and date formats before advancing to Project 2.

---

## 📁 Project Structure

```
DecodeLabs_DataCleaning_Project_1/
├── data_cleaning_project1.py        ← Main cleaning script
├── Dataset for Data Analytics.xlsx  ← Raw input dataset
└── Cleaned_Dataset_Project1.xlsx    ← Output (3 sheets)
    ├── Cleaned_Data                 ← Production-ready dataset
    ├── Change_Log                   ← Audit trail (CR001–CR006)
    └── Summary                      ← Before vs After comparison
```

---

## 📊 Dataset Overview

| Property | Value |
|---|---|
| Source | DecodeLabs Training Dataset |
| Raw Rows | 1,200 |
| Columns | 14 |
| Date Range | Jan 2023 – Jun 2025 |

### Columns
`OrderID`, `Date`, `CustomerID`, `Product`, `Quantity`, `UnitPrice`, `ShippingAddress`, `PaymentMethod`, `OrderStatus`, `TrackingNumber`, `ItemsInCart`, `CouponCode`, `ReferralSource`, `TotalPrice`

---

## 🔍 Issues Found & Fixed

### Phase 1 — Strategic Imputation
| Issue | Column | Count | Action |
|---|---|---|---|
| Missing values | `CouponCode` | 309 nulls | Filled with `'NONE'` (no coupon applied) |

> ⚠️ **Why not delete?** Listwise deletion would have removed 25.75% of all records, reducing statistical power. Filling with `'NONE'` is the correct business interpretation — orders without a coupon are still valid orders.

### Phase 2 — Integrity Audit
| Check | Result |
|---|---|
| Full duplicate rows | 0 found ✅ |
| Duplicate `OrderID` values | 0 found ✅ |
| TotalPrice = Qty × UnitPrice | All 1,200 rows pass ✅ |

### Phase 3 — Format Standardisation
| Issue | Column | Action |
|---|---|---|
| Excel serial date format | `Date` | Converted to ISO 8601 `YYYY-MM-DD` |
| Floating point precision | `UnitPrice`, `TotalPrice` | Rounded to 2 decimal places |
| Inconsistent casing / whitespace | 5 text columns | `str.strip()` + `str.title()` applied |

---

## ✅ Verification Gate (Project 2 Threshold)

| Metric | Before | After | Target |
|---|---|---|---|
| Missing Values | 309 | **0** | 0 |
| Duplicate OrderIDs | 0 | **0** | 0 |
| Bad Date Formats | Excel serial | **ISO 8601** | YYYY-MM-DD |
| Numeric Precision | Up to 14 decimals | **2 decimals** | 2 decimals |
| Total Rows | 1,200 | **1,200** | No loss |

**Result: VERIFICATION PASSED ✅ — 0% Error Rate on IDs and Dates**

---

## 🛠️ Tech Stack

| Tool | Purpose |
|---|---|
| Python 3.x | Core scripting |
| pandas | Data loading, cleaning, transformation |
| openpyxl | Styled Excel output with multiple sheets |
| os | Cross-platform file path handling |

---

## ▶️ How to Run

1. Place `data_cleaning_project1.py` and `Dataset for Data Analytics.xlsx` in the **same folder**
2. Open terminal in that folder
3. Install dependencies (first time only):
```bash
pip install pandas openpyxl
```
4. Run the script:
```bash
python data_cleaning_project1.py
```
5. Output file `Cleaned_Dataset_Project1.xlsx` will appear in the same folder

---

## 📋 Change Log Summary

| Change ID | Column | Issue | Status |
|---|---|---|---|
| CR001 | CouponCode | 309 null values | Resolved |
| CR002 | OrderID | Duplicate check | Verified Clean |
| CR003 | Date | Format standardisation | Resolved |
| CR004 | UnitPrice, TotalPrice | Numeric precision | Resolved |
| CR005 | 5 text columns | Whitespace / casing | Resolved |
| CR006 | TotalPrice | Cross-column integrity | Verified Clean |

---

## 💡 Key Learnings

- **Don't just delete nulls** — understand the business reason first
- **Imputation strategy matters**: mean/median for numeric; domain value (`'NONE'`) for categorical
- **Document every change** — if it isn't logged, it didn't happen
- **Data integrity is the foundation** — all downstream models and dashboards inherit your data quality

---

*DecodeLabs | Professional Standard | Batch 2026*
