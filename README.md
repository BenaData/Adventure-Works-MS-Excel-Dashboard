# 📊 Adventure Works MS-Excel Dashboard Project

A data analysis project built on the Adventure Works database, featuring two interactive Excel dashboards that track business performance across revenue, profit, product lines, and customer demographics over a four-year period.

---

## 🗂️ Project Overview

| Detail | Info |
|---|---|
| **Tool** | Microsoft Excel (.xlsm) |
| **Dataset** | Adventure Works Database |
| **Analysis Period** | 2005 – 2008 |
| **Dashboards** | 2 (Time Series + Sales Analysis) |
| **Total Revenue Analysed** | $307.09 Million |
| **Total Profit Analysed** | $126.29 Million (41.1% margin) |

---

## 🎯 Objectives

- Track revenue, profit, and transaction volumes across four years
- Identify seasonal patterns and peak performance periods
- Evaluate product performance by category and colour
- Understand customer demographics and their contribution to profit
- Deliver two interactive, business-ready Excel dashboards

---

## 🗃️ Dataset Structure

The database was provided as an Excel workbook with six relational tables:

| Table | Description |
|---|---|
| `FactInternetSales` | All transactional sales records |
| `DimProduct` | Product catalogue (name, colour, category) |
| `DimSalesTerritory` | Geographical sales regions |
| `DimDate` | Calendar dimension for time-based analysis |
| `DimCustomer` | Customer profiles and demographics |
| `DimGeography` | Location data for customers and territories |

---

## 📈 Dashboard 1 — Time Series Analysis

<img width="940" height="497" alt="image" src="https://github.com/user-attachments/assets/072002b8-64e1-40a6-9cbe-92c616235bec" />


It tracks business performance trends over time across years, quarters, and days of the week.

**Key Findings:**

- 📌 Total revenue of **$307.09M** and profit of **$126.29M** (41.1% margin) over four years
- 📌 Revenue, profit, and transactions all showed an **upward trajectory from 2005 to 2008**
- 📌 **Q2 was the strongest quarter**, contributing 31% of total profit ($39.30M)
- 📌 **Q3 was the weakest quarter**, contributing only 19% of total profit
- 📌 **Wednesday–Friday** drove 43.8% of total profit — mid-to-late week is the prime selling window
- 📌 **2007 delivered the highest profit** of any single year
- 📌 **2007 and 2008 tied for highest revenue** at $102M each
- 📌 **2008 had the most transactions**, indicating growing customer activity

---

## 🛍️ Dashboard 2 — Detailed Sales Analysis

<img width="940" height="426" alt="image" src="https://github.com/user-attachments/assets/710c5a6d-c01e-4329-ada1-29bde8580754" />

It drills into product and customer dimensions to reveal where profit is concentrated.

**Key Findings:**

- 📌 The **top 4 products** contributed 20% of total profit across 606 products
- 📌 **Black products** were the highest-performing colour by profit
- 📌 **Red, black, and silver** products combined contributed **75% of total profit**
- 📌 The top 5 customers contributed just **0.28% of total profit** — revenue is broadly distributed
- 📌 Profit by gender was nearly equal: **50.4% female, 49.6% male**
- 📌 Customers **aged 50+** contributed **39.4% of total profit**

---

## 🛠️ Tools & Techniques

```
✅ Power Pivot        — Data modelling across six relational tables
✅ DAX Formulas       — Custom KPIs and calculated measures
✅ Pivot Tables       — Dynamic aggregation across multiple dimensions
✅ Excel Macros (VBA) — Dashboard automation and interactivity
✅ VLOOKUP / INDEX / MATCH — Cross-sheet data lookups
✅ IF / IFS Functions — Conditional logic and category groupings
✅ Charts & Slicers   — Interactive visualisations with filter controls
✅ Conditional Formatting — Highlighting performance thresholds
```

---

## 💡 Business Recommendations

Based on the analysis:

1. **Capitalise on Q2** — invest in spring marketing campaigns to maximise the strongest seasonal period.
2. **Lift Q3 performance** — develop mid-year promotions to reduce seasonal revenue volatility.
3. **Time campaigns mid-week** — schedule launches and promotions on Wednesday–Friday for peak engagement.
4. **Focus on high-performing colours** — prioritise black, red, and silver product lines that drive 75% of profit.
5. **Target the 50+ segment** — develop loyalty and upsell programmes for this high-value demographic.
6. **Replicate top product traits** — analyse the top 4 products and apply those characteristics to new product development.
---

## 🚀 How to run the Dashboard from your computer

1. **Clone the repository**
   ```bash
   git clone https://github.com/BenaData/Adventure-Works-MS-Excel-Dashboard.git
   ```

2. **Open the workbook**
   - Open `Adventure Dashboard.xlsm` in Microsoft Excel (2016 or later recommended)
   - **Enable macros** when prompted — they power the dashboard interactivity

3. **Explore the dashboards**
   - Navigate between the Time Series and Sales Analysis dashboards using the sheet tabs
   - Use the slicers and filters to interact with the data

> ⚠️ This file is macro-enabled (`.xlsm`). Make sure macros are enabled in your Excel Trust Center settings for full functionality.

---

## 👤 Author

**Benard Mwinzi** — Data Analyst based in Nairobi, Kenya

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-blue?logo=linkedin)](https://www.linkedin.com/in/benard-musyoka1234/)
[![Kaggle](https://img.shields.io/badge/Kaggle-Profile-20BEFF?logo=kaggle)](https://www.kaggle.com/benardmwinzi)
[![Portfolio](https://img.shields.io/badge/Portfolio-Visit-brightgreen)](https://www.datascienceportfol.io/Benard)

---

*Built with Microsoft Excel | Adventure Works Dataset*
