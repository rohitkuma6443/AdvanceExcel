# Excel Charts – Simple Notes

## 1. What is a Chart?

A chart is a graphical representation of data.

It helps us to:

* Compare data
* Find trends
* Understand patterns
* Present data clearly

---

# Practice Data

| Region |  Sales | Profit |
| ------ | -----: | -----: |
| North  | 150000 |  20000 |
| South  | 120000 |  15000 |
| East   | 180000 |  25000 |
| West   | 140000 |  18000 |

---

# 2. Column Chart

A **Column Chart** uses vertical bars to compare values.

![Excel Column Chart]([https://support.microsoft.com/images/en-us/7d3c6c75-4b5c-4e77-a6d4-6f7d2f8b6f1d](https://www.excel-easy.com/examples/images/column-chart/column-chart.png))

### Example

**Question:** Which region has the highest sales?

Select:

```text
Region + Sales
→ Insert
→ Column Chart
→ Clustered Column
```

### Use Case

Use a Column Chart for:

* Sales by Region
* Sales by Product
* Employees by Department
* Students by Course

---

# 3. Bar Chart

A **Bar Chart** uses horizontal bars.

![Excel Bar Chart](<img width="780" height="480" alt="image" src="https://github.com/user-attachments/assets/05baf432-eebf-42d4-827c-e757596efe42" />
)

### Example

| Product |  Sales |
| ------- | -----: |
| Laptop  | 200000 |
| Mobile  | 150000 |
| Printer | 100000 |

```text
Select Product + Sales
→ Insert
→ Bar Chart
```

### Use Case

Use a Bar Chart for:

* Product comparison
* Sales ranking
* Employee performance
* Categories with long names

---

# 4. Line Chart

A **Line Chart** is used to show trends over time.

![Excel Line Chart](https://support.microsoft.com/images/en-us/1a6e4d4c-6b7d-4e8f-9a0b-1c2d3e4f5a6b)

### Example

| Month |  Sales |
| ----- | -----: |
| Jan   | 100000 |
| Feb   | 130000 |
| Mar   | 150000 |
| Apr   | 140000 |

```text
Select Month + Sales
→ Insert
→ Line Chart
```

### Use Case

Use a Line Chart for:

* Monthly sales
* Daily expenses
* Website traffic
* Stock prices
* Production trends

---

# 5. Pie Chart

A **Pie Chart** shows parts of a whole.

![Excel Pie Chart](https://support.microsoft.com/images/en-us/6e5b4c3d-2f1a-4b8c-9d7e-5f6a7b8c9d0e)

### Example

| Category    |  Sales |
| ----------- | -----: |
| Electronics | 300000 |
| Furniture   | 150000 |
| Stationery  |  50000 |

```text
Select Category + Sales
→ Insert
→ Pie Chart
→ 2-D Pie
```

### Use Case

Use a Pie Chart for:

* Sales contribution
* Expense distribution
* Market share
* Category contribution

### Remember

Pie Charts work best when there are **only a few categories**.

---

# 6. Doughnut Chart

A **Doughnut Chart** is similar to a Pie Chart but has a hole in the center.

![Excel Doughnut Chart](https://support.microsoft.com/images/en-us/6e5b4c3d-2f1a-4b8c-9d7e-5f6a7b8c9d0e)

### Example

```text
Electronics → 60%
Furniture   → 30%
Others      → 10%
```

```text
Insert
→ Doughnut Chart
```

### Use Case

* Sales contribution
* Budget distribution
* Expense distribution
* Category contribution

---

# 7. Area Chart

An **Area Chart** is similar to a Line Chart, but the area below the line is filled.

![Excel Area Chart](https://support.microsoft.com/images/en-us/1a6e4d4c-6b7d-4e8f-9a0b-1c2d3e4f5a6b)

### Example

| Month |  Sales |
| ----- | -----: |
| Jan   | 100000 |
| Feb   | 150000 |
| Mar   | 200000 |

### Use Case

Use an Area Chart for:

* Revenue trends
* Production trends
* Website traffic
* Showing magnitude over time

---

# 8. Scatter Chart

A **Scatter Chart** shows the relationship between two numerical values.

![Excel Scatter Chart](https://support.microsoft.com/images/en-us/2b7c4d6e-8f1a-4b3c-9d5e-7f8a9b0c1d2e)

### Example

| Study Hours | Marks |
| ----------: | ----: |
|           2 |    45 |
|           4 |    55 |
|           6 |    68 |
|           8 |    80 |
|          10 |    92 |

```text
X-Axis → Study Hours
Y-Axis → Marks
```

```text
Select Data
→ Insert
→ Scatter
→ Scatter with Markers
```

### Use Case

Use Scatter Charts for:

* Study Hours vs Marks
* Advertising vs Sales
* Price vs Demand
* Experience vs Salary

---

# 9. Histogram

A **Histogram** shows the distribution of numerical data.

![Excel Histogram](https://support.microsoft.com/images/en-us/1a6e4d4c-6b7d-4e8f-9a0b-1c2d3e4f5a6b)

### Example

Student Marks:

```text
25, 35, 42, 45, 51, 55,
60, 65, 72, 78, 85, 90
```

The values can be grouped into:

```text
0–20
21–40
41–60
61–80
81–100
```

```text
Select Marks
→ Insert
→ Statistical Chart
→ Histogram
```

### Use Case

* Student marks
* Salary distribution
* Age distribution
* Sales amount distribution

---

# 10. Pareto Chart

A **Pareto Chart** contains:

```text
Bars + Cumulative Line
```

![Excel Pareto Chart](https://support.microsoft.com/images/en-us/1a6e4d4c-6b7d-4e8f-9a0b-1c2d3e4f5a6b)

### Example

| Problem         | Complaints |
| --------------- | ---------: |
| Late Delivery   |         80 |
| Damaged Product |         50 |
| Wrong Product   |         30 |
| Poor Packaging  |         20 |

```text
Select Problem + Complaints
→ Insert
→ Statistical Chart
→ Pareto
```

### Use Case

Use a Pareto Chart to identify major causes of problems.

Examples:

* Customer complaints
* Product defects
* Quality problems
* Business problems

---

# 11. Waterfall Chart

A **Waterfall Chart** shows how a value increases or decreases from a starting value to an ending value.

![Excel Waterfall Chart](https://support.microsoft.com/images/en-us/1a6e4d4c-6b7d-4e8f-9a0b-1c2d3e4f5a6b)

### Example

| Item            |  Amount |
| --------------- | ------: |
| Opening Balance |  500000 |
| Sales           |  200000 |
| Expenses        | -100000 |
| Tax             |  -50000 |
| Closing Balance |  550000 |

### Use Case

* Profit analysis
* Cash flow
* Budget analysis
* Revenue analysis
* Opening vs Closing Balance

---

# 12. Combo Chart

A **Combo Chart** combines two different chart types.

Most commonly:

```text
Column + Line
```

![Excel Combo Chart](https://support.microsoft.com/images/en-us/1a6e4d4c-6b7d-4e8f-9a0b-1c2d3e4f5a6b)

### Example

| Month |  Sales | Profit % |
| ----- | -----: | -------: |
| Jan   | 100000 |      10% |
| Feb   | 150000 |      12% |
| Mar   | 180000 |      15% |

Use:

```text
Sales    → Column
Profit % → Line
```

### Create

```text
Select Data
→ Insert
→ Combo Chart
→ Column + Line
```

### Use Case

* Sales + Profit %
* Revenue + Growth %
* Actual + Target
* Sales + Quantity

---

# 13. Radar Chart

A **Radar Chart** compares multiple values across different categories.

![Excel Radar Chart](https://support.microsoft.com/images/en-us/1a6e4d4c-6b7d-4e8f-9a0b-1c2d3e4f5a6b)

### Example

| Employee | Excel | SQL | Power BI | Python |
| -------- | ----: | --: | -------: | -----: |
| Amit     |    90 |  70 |       80 |     75 |

### Use Case

* Employee skills
* Student performance
* Product features
* Competency analysis

---

# 14. Map Chart

A **Map Chart** displays data based on geographical locations.

![Excel Map Chart](https://support.microsoft.com/images/en-us/1a6e4d4c-6b7d-4e8f-9a0b-1c2d3e4f5a6b)

### Example

| State       |  Sales |
| ----------- | -----: |
| Maharashtra | 500000 |
| Gujarat     | 350000 |
| Karnataka   | 400000 |

### Use Case

* Sales by State
* Sales by Country
* Customers by Region
* Geographic analysis

---

# 15. Stock Chart

A **Stock Chart** is mainly used for financial data.

![Excel Stock Chart](https://support.microsoft.com/images/en-us/1a6e4d4c-6b7d-4e8f-9a0b-1c2d3e4f5a6b)

### Example

| Date   | Open | High | Low | Close |
| ------ | ---: | ---: | --: | ----: |
| 01-Jan |  100 |  110 |  95 |   108 |
| 02-Jan |  108 |  115 | 102 |   112 |

### Use Case

* Stock market data
* Open/High/Low/Close analysis
* Financial analysis

---

# 16. PivotChart

A **PivotChart** is created from a PivotTable and is useful for interactive data analysis.

### Example

Create a PivotTable:

```text
Rows
→ Region

Values
→ Sum of Sales
```

Then:

```text
PivotTable
→ PivotChart
```

You can also use:

```text
Slicer
Timeline
Filters
```

### Use Case

* Sales dashboard
* Business reports
* Interactive analysis
* Large datasets
* Data Analyst reports
