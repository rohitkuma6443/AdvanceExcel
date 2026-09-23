# Excel Interview Questions with Answers

---

## 1. Advanced Formula & Function Concepts

### 1. What is the difference between XLOOKUP, VLOOKUP, HLOOKUP, and INDEX + MATCH?

**Answer:**

* `VLOOKUP` searches vertically and can return values only to the right of the lookup column.
* `HLOOKUP` searches horizontally.
* `XLOOKUP` can search vertically or horizontally and can return values from either side.
* `INDEX + MATCH` provides flexible lookup capability and works well in older Excel versions.

**Interview point:** `XLOOKUP` is generally easier and more flexible than `VLOOKUP`.

---

### 2. Why is XLOOKUP more flexible than VLOOKUP?

**Answer:**

`XLOOKUP`:

* Can search left or right.
* Uses exact match by default.
* Can return a custom value when nothing is found.
* Can return multiple columns.
* Can search from first to last or last to first.

Example:

```excel
=XLOOKUP(A2,D:D,E:E,"Not Found")
```

---

### 3. What are the limitations of VLOOKUP?

**Answer:**

Major limitations:

1. Lookup column must be the first column of the selected range.
2. It normally returns values only from columns to the right.
3. Inserting columns can affect the column-index argument.
4. Approximate matching can produce incorrect results if the data is not appropriately sorted.

---

### 4. How does INDEX + MATCH work?

**Answer:**

`MATCH` finds the position of a value, while `INDEX` returns the value from that position.

```excel
=INDEX(C2:C10,MATCH(E2,A2:A10,0))
```

Here:

* `MATCH` finds where `E2` exists in `A2:A10`.
* `INDEX` returns the corresponding value from `C2:C10`.

---

### 5. What is exact match vs approximate match?

**Answer:**

**Exact match** searches for an identical value.

```excel
=MATCH(A2,B:B,0)
```

`0` means exact match.

**Approximate match** searches for the closest appropriate value based on the lookup rules and is commonly used for ranges such as tax slabs, grades, or commission rates.

---

### 6. What is XMATCH?

**Answer:**

`XMATCH` is the modern replacement for `MATCH`.

It supports:

* Exact matching.
* Approximate matching.
* Reverse searching.
* Wildcards.
* Flexible search modes.

Example:

```excel
=XMATCH("Mumbai",A2:A100)
```

---

### 7. Difference between SUMIF and SUMIFS?

**Answer:**

`SUMIF` applies **one criterion**.

```excel
=SUMIF(A:A,"West",B:B)
```

`SUMIFS` can apply **multiple criteria**.

```excel
=SUMIFS(C:C,A:A,"West",B:B,"Laptop")
```

---

### 8. Difference between COUNTIF and COUNTIFS?

**Answer:**

`COUNTIF` counts records based on one condition.

`COUNTIFS` counts records based on multiple conditions.

```excel
=COUNTIF(A:A,"Mumbai")
```

```excel
=COUNTIFS(A:A,"Mumbai",B:B,"Premium")
```

---

### 9. Difference between AVERAGEIF and AVERAGEIFS?

**Answer:**

`AVERAGEIF` calculates an average based on one condition.

`AVERAGEIFS` calculates an average based on multiple conditions.

---

### 10. Can SUMIFS use multiple criteria?

**Answer:**

Yes.

For example:

```excel
=SUMIFS(E:E,A:A,"West",B:B,"Electronics",C:C,">10000")
```

This sums values where all three conditions are satisfied.

---

### 11. How can SUMIFS be used with dates?

**Answer:**

You can specify a start and end date.

```excel
=SUMIFS(C:C,A:A,">="&DATE(2026,1,1),A:A,"<="&DATE(2026,1,31))
```

This calculates sales between January 1 and January 31, 2026.

---

### 12. SUMPRODUCT vs SUMIFS?

**Answer:**

`SUMIFS` is designed specifically for conditional summation.

`SUMPRODUCT` can perform more flexible mathematical calculations across arrays.

For example:

```excel
=SUMPRODUCT(B2:B10,C2:C10)
```

calculates:

```text
B2*C2 + B3*C3 + ... + B10*C10
```

---

### 13. Why is SUMPRODUCT useful for analysts?

**Answer:**

Because it can combine arrays and conditions without necessarily requiring helper columns.

It can be used for:

* Weighted averages.
* Conditional calculations.
* Multiplication of multiple columns.
* Complex analytical formulas.

---

### 14. Difference between COUNT, COUNTA, COUNTBLANK and COUNTIF?

**Answer:**

| Function   | Purpose                             |
| ---------- | ----------------------------------- |
| COUNT      | Counts numbers                      |
| COUNTA     | Counts non-empty cells              |
| COUNTBLANK | Counts blank cells                  |
| COUNTIF    | Counts cells satisfying a condition |

---

### 15. Difference between MAX and MAXIFS?

**Answer:**

`MAX` returns the largest value from a range.

```excel
=MAX(B2:B100)
```

`MAXIFS` returns the largest value subject to conditions.

```excel
=MAXIFS(C:C,A:A,"West")
```

---

### 16. Difference between LARGE and MAX?

**Answer:**

`MAX` returns only the largest value.

`LARGE` can return the 1st, 2nd, 3rd, etc. largest value.

```excel
=LARGE(A2:A100,3)
```

returns the third-largest value.

---

### 17. Difference between IFERROR and IFNA?

**Answer:**

`IFERROR` handles almost all Excel errors.

```excel
=IFERROR(A2/B2,0)
```

`IFNA` specifically handles `#N/A`.

```excel
=IFNA(XLOOKUP(A2,B:B,C:C),"Not Found")
```

---

### 18. When should IFERROR be used?

**Answer:**

Use it when an expected calculation may legitimately generate an error and you want a controlled output.

However, it should not be used simply to hide errors caused by bad formulas or bad data.

---

### 19. What is a nested IF?

**Answer:**

A nested `IF` contains another `IF`.

Example:

```excel
=IF(A2>=90,"A",IF(A2>=75,"B",IF(A2>=60,"C","D")))
```

It is useful for multiple conditions but can become difficult to maintain when there are many conditions.

---

### 20. What problems can excessive nested IF formulas create?

**Answer:**

They can:

* Become difficult to read.
* Become difficult to debug.
* Increase maintenance effort.
* Create logical errors.

For complex rules, `IFS`, `SWITCH`, lookup tables, or Power Query may be better.

---

# 2. Dynamic Array & Modern Excel

### 21. What is a Dynamic Array?

**Answer:**

A dynamic array formula can return multiple results automatically into neighboring cells.

Example:

```excel
=FILTER(A2:C100,C2:C100="West")
```

Excel automatically spills the results into multiple cells.

---

### 22. What is a spilled array?

**Answer:**

When one formula produces multiple results across several cells, those results are called a **spilled array**.

Only the first cell contains the formula.

---

### 23. What causes #SPILL!?

**Answer:**

`#SPILL!` occurs when Excel cannot place the results of a dynamic array into the required cells.

Common reasons:

* Cells contain existing data.
* Merged cells are blocking the range.
* The spill range is otherwise unavailable.

---

### 24. What does FILTER do?

**Answer:**

`FILTER` returns records satisfying specified conditions.

Example:

```excel
=FILTER(A2:D100,D2:D100="West")
```

It dynamically returns rows belonging to the West region.

---

### 25. What does UNIQUE do?

**Answer:**

`UNIQUE` returns distinct values from a range.

```excel
=UNIQUE(A2:A100)
```

It is useful for creating unique customer, city, product, or category lists.

---

### 26. Difference between SORT and SORTBY?

**Answer:**

`SORT` sorts data based on a column position.

`SORTBY` sorts one range based on another range.

`SORTBY` is useful when the sorting logic is based on a separate column or array.

---

### 27. Can FILTER, SORT and UNIQUE be combined?

**Answer:**

Yes.

Example:

```excel
=SORT(UNIQUE(FILTER(A2:A100,B2:B100="West")))
```

This:

1. Filters West records.
2. Removes duplicates.
3. Sorts the result.

---

### 28. What is SEQUENCE?

**Answer:**

`SEQUENCE` generates a sequence of numbers.

```excel
=SEQUENCE(10)
```

returns numbers from 1 to 10.

It is useful for generating:

* Serial numbers.
* Dates.
* Periods.
* Test data.

---

### 29. What is RANDARRAY?

**Answer:**

`RANDARRAY` generates an array of random numbers.

Example:

```excel
=RANDARRAY(10,2,1,100,TRUE)
```

This can generate a 10 × 2 array of random integers between 1 and 100.

---

### 30. What is TRANSPOSE?

**Answer:**

`TRANSPOSE` changes rows into columns and columns into rows.

```excel
=TRANSPOSE(A1:D1)
```

converts a horizontal range into a vertical array.

---

### 31. What is LET?

**Answer:**

`LET` allows you to assign names to intermediate calculations within a formula.

Example:

```excel
=LET(
sales,B2*C2,
discount,sales*D2,
sales-discount
)
```

It can make complex formulas easier to read and can avoid repeating calculations.

---

### 32. How can LET improve performance?

**Answer:**

If the same calculation is repeated several times, `LET` can calculate it once and reuse the result.

This can improve readability and potentially reduce unnecessary calculation.

---

### 33. What is LAMBDA?

**Answer:**

`LAMBDA` allows users to create reusable custom functions without VBA.

For example, a custom calculation can be defined once and reused throughout a workbook.

---

### 34. What is an advantage of LAMBDA?

**Answer:**

It allows organizations to standardize repeated business calculations and reduce duplicated formulas.

---

### 35. Why are dynamic arrays useful for Data Analysts?

**Answer:**

They make many analytical tasks easier because results automatically expand or contract as the source data changes.

They are particularly useful for:

* Dynamic reports.
* Unique lists.
* Filtering.
* Ranking.
* Lookup results.
* Dashboard components.

---

# 3. Advanced Lookup

### 36. What is a two-way lookup?

**Answer:**

A two-way lookup searches using **both a row criterion and a column criterion**.

For example:

```text
Product + Month → Sales
```

Both the product and month determine the result.

---

### 37. How can INDEX and MATCH perform a two-way lookup?

**Answer:**

Typically:

```excel
=INDEX(data,MATCH(row_value,row_labels,0),MATCH(column_value,column_labels,0))
```

The first `MATCH` finds the row and the second finds the column.

---

### 38. How can XLOOKUP perform a two-way lookup?

**Answer:**

Nested `XLOOKUP` functions can be used.

Conceptually:

```excel
=XLOOKUP(Row_Value,Row_Range,
         XLOOKUP(Column_Value,Column_Headers,Data))
```

---

### 39. Can XLOOKUP return multiple columns?

**Answer:**

Yes.

If the return range contains multiple columns, `XLOOKUP` can spill those columns into adjacent cells.

---

### 40. How can you perform a multi-criteria lookup?

**Answer:**

You can combine criteria.

For example, conceptually:

```excel
=XLOOKUP(1,(A2:A100="Mumbai")*(B2:B100="Premium"),C2:C100)
```

Both conditions must be true.

---

### 41. What is a multi-criteria lookup?

**Answer:**

It is a lookup where more than one condition determines the matching record.

Example:

```text
Customer = Rohit
AND
Product = Laptop
```

---

### 42. How can you return the second matching record?

**Answer:**

One approach is to use `FILTER` and then select the required item.

For example:

```excel
=INDEX(FILTER(C2:C100,A2:A100="Rohit"),2)
```

This returns the second matching value.

---

### 43. How can XLOOKUP display a custom message when no result exists?

**Answer:**

Use its `if_not_found` argument:

```excel
=XLOOKUP(A2,B:B,C:C,"Not Found")
```

---

### 44. How can you retrieve the last matching value?

**Answer:**

`XLOOKUP` supports reverse search.

Conceptually:

```excel
=XLOOKUP(value,lookup_array,return_array,,0,-1)
```

The `-1` search mode searches from the last item toward the first.

---

### 45. What happens when duplicate lookup values exist?

**Answer:**

A normal lookup usually returns the **first matching result**, depending on the function and search mode.

If all matching records are required, functions such as `FILTER` are more appropriate.

---

# 4. Date & Time

### 46. Difference between TODAY and NOW?

**Answer:**

`TODAY()` returns the current date.

`NOW()` returns the current date and current time.

---

### 47. How does Excel store dates?

**Answer:**

Excel generally stores dates as serial numbers.

For example, a date is represented internally by a numeric value, while formatting makes it appear as a calendar date.

---

### 48. Why can dates behave like numbers?

**Answer:**

Because Excel stores dates as serial numbers.

Therefore, date arithmetic is possible.

For example:

```excel
=B2-A2
```

can calculate the number of days between two dates.

---

### 49. DATEDIF vs DAYS?

**Answer:**

`DAYS` calculates the number of days between two dates.

`DATEDIF` can calculate differences in:

* Years
* Months
* Days

depending on its unit argument.

---

### 50. Difference between EDATE and EOMONTH?

**Answer:**

`EDATE` returns a date shifted by a specified number of months.

`EOMONTH` returns the last day of the resulting month.

---

### 51. What does NETWORKDAYS do?

**Answer:**

It calculates the number of working days between two dates, normally excluding Saturday and Sunday.

---

### 52. What does NETWORKDAYS.INTL do?

**Answer:**

It provides more control over weekends.

It allows you to define which days should be treated as weekends.

---

### 53. How can employee tenure be calculated?

**Answer:**

A common approach is:

```excel
=DATEDIF(Start_Date,TODAY(),"Y")
```

This returns completed years.

---

### 54. What is Month-over-Month growth?

**Answer:**

It compares the current month's value with the previous month's value.

Formula:

```text
(Current Month - Previous Month) / Previous Month
```

---

### 55. What is Year-over-Year growth?

**Answer:**

It compares a metric with the same period from the previous year.

```text
(Current Year - Previous Year) / Previous Year
```

---

### 56. How can you identify a quarter?

**Answer:**

One approach is:

```excel
="Q"&ROUNDUP(MONTH(A2)/3,0)
```

For example:

* January → Q1
* April → Q2
* July → Q3
* October → Q4

---

### 57. How can you find the last day of a month?

**Answer:**

Use:

```excel
=EOMONTH(A2,0)
```

---

### 58. How can holidays be excluded from working days?

**Answer:**

Pass the holiday range to `NETWORKDAYS`:

```excel
=NETWORKDAYS(A2,B2,Holidays)
```

---

### 59. Why can date formulas return unexpected results?

**Answer:**

Common reasons include:

* Dates stored as text.
* Different regional date formats.
* Invalid dates.
* Hidden spaces.
* Imported data containing mixed formats.

---

### 60. How can you identify dates stored as text?

**Answer:**

One approach is to test whether Excel recognizes the value as a number/date.

You can also inspect the cell formatting and use functions such as `ISNUMBER`.

---

# 5. Text & Data Cleaning

### 61. Difference between LEFT, RIGHT and MID?

**Answer:**

* `LEFT` extracts characters from the beginning.
* `RIGHT` extracts characters from the end.
* `MID` extracts characters from a specified position.

---

### 62. FIND vs SEARCH?

**Answer:**

`FIND` is case-sensitive.

`SEARCH` is not case-sensitive and supports wildcard characters.

---

### 63. SUBSTITUTE vs REPLACE?

**Answer:**

`SUBSTITUTE` replaces specific text.

`REPLACE` replaces characters based on their position.

---

### 64. TRIM vs CLEAN?

**Answer:**

`TRIM` removes unnecessary spaces from text.

`CLEAN` removes certain non-printable characters.

For imported data, both can be useful.

---

### 65. What does TEXTSPLIT do?

**Answer:**

It splits text into rows or columns using delimiters.

Example:

```excel
=TEXTSPLIT(A2,",")
```

can split comma-separated data.

---

### 66. TEXTBEFORE vs TEXTAFTER?

**Answer:**

`TEXTBEFORE` returns text before a specified delimiter.

`TEXTAFTER` returns text after a specified delimiter.

---

### 67. How can you extract the domain from an email?

**Answer:**

Using `TEXTAFTER`:

```excel
=TEXTAFTER(A2,"@")
```

---

### 68. How can you separate first and last names?

**Answer:**

Modern Excel can use:

```excel
=TEXTBEFORE(A2," ")
```

for the first part and:

```excel
=TEXTAFTER(A2," ")
```

for the remaining text.

For more complex names, `TEXTSPLIT` can be useful.

---

### 69. Why might identical-looking names not match?

**Answer:**

They may contain:

* Leading spaces.
* Trailing spaces.
* Non-printing characters.
* Different Unicode characters.
* Different spellings.

For example:

```text
"Rohit"
"Rohit "
```

look similar but are technically different strings.

---

### 70. How do hidden spaces affect lookup?

**Answer:**

A lookup searches for the actual underlying value. Extra spaces can cause the lookup value and source value to be different.

Cleaning with functions such as `TRIM` can help.

---

### 71. How can text be standardized?

**Answer:**

Common functions include:

```excel
=UPPER(A2)
=LOWER(A2)
=PROPER(A2)
=TRIM(A2)
=CLEAN(A2)
```

The appropriate function depends on the required standard.

---

### 72. How can numbers stored as text be converted?

**Answer:**

Possible methods include:

```excel
=VALUE(A2)
```

or multiplying by 1:

```excel
=A2*1
```

provided the text represents a valid number.

---

### 73. How can inconsistent capitalization be identified?

**Answer:**

Functions such as `UPPER`, `LOWER`, and `PROPER` can standardize text.

For analysis, standardization is often preferable to merely identifying differences.

---

### 74. What is data cleaning?

**Answer:**

Data cleaning is the process of identifying and correcting problems such as:

* Duplicates.
* Missing values.
* Incorrect data types.
* Invalid dates.
* Inconsistent text.
* Incorrect categories.
* Formatting problems.

---

### 75. Which Excel tools can perform data cleaning?

**Answer:**

Important tools include:

* Power Query.
* Remove Duplicates.
* Text to Columns.
* Find & Replace.
* Flash Fill.
* Data Validation.
* Excel formulas.
* Conditional Formatting.

---

# 6. Logical & Conditional Analysis

### 76. IF vs IFS vs SWITCH?

**Answer:**

`IF` handles conditional logic.

`IFS` handles multiple conditions sequentially.

`SWITCH` compares an expression against multiple specified values.

---

### 77. What is AND?

**Answer:**

`AND` returns TRUE only when **all specified conditions are TRUE**.

```excel
=AND(A2>100,B2="West")
```

---

### 78. What is OR?

**Answer:**

`OR` returns TRUE when **at least one condition is TRUE**.

```excel
=OR(A2="West",A2="North")
```

---

### 79. How can AND and OR be combined?

**Answer:**

They can be nested to create complex business rules.

Example:

```excel
=IF(AND(A2>=50000,OR(B2="Premium",B2="Corporate")),"High Value","Other")
```

---

### 80. How can Excel create customer categories?

**Answer:**

Using business rules with `IF`, `IFS`, `SWITCH`, or lookup tables.

For example:

```text
Sales >= 100000 → Platinum
Sales >= 50000  → Gold
Sales >= 25000  → Silver
Otherwise       → Bronze
```

---

### 81. Nested IF vs IFS?

**Answer:**

`IFS` is usually easier to read when there are many sequential conditions.

Nested `IF` is useful when the logic requires more complex branching.

---

### 82. When is SWITCH useful?

**Answer:**

`SWITCH` is useful when one expression needs to be compared against several known values.

For example:

```text
"IN" → India
"US" → United States
"UK" → United Kingdom
```

---

### 83. How should blanks be handled in formulas?

**Answer:**

First determine whether the blank means:

* Missing data.
* Not applicable.
* Zero.
* Unknown.

These have different analytical meanings and should not automatically be treated as zero.

---

### 84. Blank cell vs formula returning ""?

**Answer:**

A genuinely empty cell contains nothing.

A formula such as:

```excel
=""
```

returns an empty text string.

These can behave differently in some formulas and functions.

---

### 85. How can contradictory business rules be avoided?

**Answer:**

Rules should be:

1. Clearly defined.
2. Ordered correctly.
3. Mutually consistent.
4. Tested against boundary values.

For example, if one rule says `Sales >= 50,000` and another earlier rule catches all sales above `40,000`, the second rule may never be reached.

---

# 7. PivotTables & Data Analysis

### 86. What is a PivotTable?

**Answer:**

A PivotTable is an Excel analytical tool used to summarize and analyze large datasets without manually writing many formulas.

It can summarize:

* Sales.
* Quantity.
* Profit.
* Customer counts.
* Categories.
* Regions.
* Dates.

---

### 87. What does Sum, Count and Average mean in a PivotTable?

**Answer:**

* **Sum:** Total numerical value.
* **Count:** Number of records/non-empty values depending on field type.
* **Average:** Arithmetic mean.

The aggregation must match the business meaning of the metric.

---

### 88. What is a calculated field?

**Answer:**

A calculated field in a traditional PivotTable creates a calculation using fields from the PivotTable's source data.

For example:

```text
Profit = Sales - Cost
```

---

### 89. What is a PivotChart?

**Answer:**

A PivotChart is a chart connected to a PivotTable.

When the PivotTable is filtered or reorganized, the PivotChart responds accordingly.

---

### 90. What is a slicer?

**Answer:**

A slicer is a visual filtering control.

For example, a sales dashboard could have slicers for:

```text
Region
Year
Category
Salesperson
```

---

### 91. What is a Timeline?

**Answer:**

A Timeline is a visual date-filtering control for PivotTables.

It allows users to filter dates by:

* Years.
* Quarters.
* Months.
* Days.

---

### 92. Can one slicer control multiple PivotTables?

**Answer:**

Yes, if the PivotTables share the appropriate data source or Data Model and are connected through slicer/report connections.

---

### 93. How can dates be grouped in a PivotTable?

**Answer:**

Dates can commonly be grouped into:

* Years.
* Quarters.
* Months.
* Days.

This allows time-series analysis without manually creating those columns.

---

### 94. How can percentage of total be shown in a PivotTable?

**Answer:**

Use:

**Value Field Settings → Show Values As → % of Grand Total**

Other options include:

* % of Row Total.
* % of Column Total.
* % Difference From.
* Running Total.

---

### 95. How can month-over-month change be shown?

**Answer:**

A PivotTable can use **Show Values As → % Difference From** and compare each month with the previous month.

---

### 96. % of Grand Total vs % of Row Total?

**Answer:**

**% of Grand Total** compares a value with the entire PivotTable total.

**% of Row Total** compares a value with the total for its row.

The denominator is therefore different.

---

### 97. Why might a PivotTable show unexpected totals?

**Answer:**

Possible reasons include:

* Incorrect source data.
* Blank records.
* Incorrect data types.
* Duplicate records.
* Incorrect aggregation.
* Old data because the PivotTable wasn't refreshed.
* Incorrect relationships in the Data Model.

---

### 98. What does Refresh do?

**Answer:**

Refresh updates the PivotTable using the latest available source data.

It is important when the underlying dataset has changed.

---

### 99. What is the Excel Data Model?

**Answer:**

The Data Model allows multiple related tables to be connected and analyzed together.

For example:

```text
Customers
     ↓
Orders
     ↓
Products
```

Instead of putting everything into one huge table, relationships can connect the tables.

---

### 100. What is the difference between a normal PivotTable and a Power Pivot/Data Model PivotTable?

**Answer:**

A normal PivotTable generally works from a single source/range or can use certain supported sources.

A Data Model/Power Pivot PivotTable can work with:

* Multiple related tables.
* Relationships.
* DAX measures.
* Larger analytical models.

This makes it much more suitable for advanced Data Analytics.

---

# 8. 20 Expert-Level Follow-Up Questions with Answers

### 101. Why can a lookup return the wrong result even when the value appears correct?

**Answer:**

Possible reasons:

* Extra spaces.
* Text vs number mismatch.
* Duplicate lookup values.
* Incorrect match mode.
* Incorrect lookup range.
* Hidden characters.
* Incorrect sorting when approximate matching is used.

---

### 102. What happens when XLOOKUP finds duplicate values?

**Answer:**

By default, it returns the **first matching result**.

If you need all matches, use `FILTER`.

---

### 103. When might INDEX + MATCH still be used?

**Answer:**

Reasons include:

* Compatibility with older Excel versions.
* Existing legacy workbooks.
* Familiarity with the pattern.
* Complex lookup models already built using it.

---

### 104. What are volatile functions?

**Answer:**

Volatile functions recalculate when Excel recalculates the workbook, even when their direct inputs have not changed.

Examples include:

```text
NOW()
TODAY()
RAND()
RANDBETWEEN()
OFFSET()
INDIRECT()
```

---

### 105. Why can volatile functions affect performance?

**Answer:**

If a workbook contains thousands of formulas depending on volatile functions, repeated recalculation can increase calculation time.

---

### 106. Why can thousands of lookup formulas slow Excel?

**Answer:**

Each formula may require Excel to perform searches and calculations.

Large datasets combined with many complex formulas can increase calculation workload.

Alternatives can include:

* Power Query.
* PivotTables.
* Data Model.
* More efficient formulas.
* Reducing unnecessary calculations.

---

### 107. What is an Excel Table?

**Answer:**

An Excel Table is a structured data range with features such as:

* Automatic expansion.
* Structured references.
* Built-in filtering.
* Automatic formula propagation.
* Easier PivotTable/data-source management.

---

### 108. What are structured references?

**Answer:**

Instead of:

```excel
=C2*D2
```

a Table can use:

```excel
=[@Quantity]*[@Unit_Price]
```

This makes formulas more readable and automatically adapts to table expansion.

---

### 109. What is Power Query?

**Answer:**

Power Query is Excel's data extraction and transformation tool.

It can:

* Import data.
* Clean data.
* Merge datasets.
* Append datasets.
* Remove duplicates.
* Transform columns.
* Change data types.
* Automate repeatable data preparation.

---

### 110. Power Query vs Excel formulas?

**Answer:**

**Excel formulas** are excellent for calculations and interactive worksheet analysis.

**Power Query** is generally better for repeatable ETL/data-cleaning workflows.

A Data Analyst may use both.

---

### 111. What is Power Pivot?

**Answer:**

Power Pivot is Excel's advanced data modeling capability.

It allows analysts to work with:

* Multiple tables.
* Relationships.
* Large datasets.
* DAX calculations.
* Measures.

---

### 112. What is DAX?

**Answer:**

DAX stands for **Data Analysis Expressions**.

It is used in Power Pivot and Power BI to create calculations such as:

```text
Total Sales
Profit Margin
Year-over-Year Growth
Running Total
```

---

### 113. Calculated column vs Measure?

**Answer:**

A **calculated column** calculates a value for each row.

A **measure** calculates dynamically based on the current filter/context.

Measures are generally preferred for aggregated analytical calculations in a Data Model.

---

### 114. What is a star schema?

**Answer:**

A star schema generally consists of:

```text
        Product
           |
Customer — Fact Sales — Date
           |
        Geography
```

A central **fact table** contains transactions, while **dimension tables** describe entities such as customers, products and dates.

---

### 115. Why separate raw data, calculations and reporting?

**Answer:**

It improves:

* Maintainability.
* Data integrity.
* Debugging.
* Reusability.
* Auditability.

A common structure is:

```text
Raw Data
   ↓
Data Cleaning
   ↓
Calculations / Model
   ↓
Dashboard / Report
```

---

### 116. Data Validation vs Conditional Formatting?

**Answer:**

**Data Validation** controls or restricts what users can enter.

Example:

```text
Allowed values:
Pending
Approved
Rejected
```

**Conditional Formatting** changes the visual appearance based on conditions.

Example:

```text
Profit < 0 → highlight
```

---

### 117. What is a circular reference?

**Answer:**

A circular reference occurs when a formula directly or indirectly refers back to itself.

Example:

```text
A1 → B1 → C1 → A1
```

This creates a calculation loop.

---

### 118. What is iterative calculation?

**Answer:**

Iterative calculation allows Excel to repeatedly calculate formulas containing circular references until specified calculation limits are reached.

It can be useful for certain specialized financial models but should not be used to hide accidental circular references.

---

### 119. What does #VALUE! mean?

**Answer:**

`#VALUE!` generally indicates that a formula received an inappropriate data type or argument.

Example:

Trying to perform an arithmetic operation on incompatible text.

---

### 120. What are common Excel errors?

**Answer:**

| Error     | Typical Meaning             |
| --------- | --------------------------- |
| `#N/A`    | Value not available/found   |
| `#VALUE!` | Incorrect value/type        |
| `#REF!`   | Invalid cell reference      |
| `#DIV/0!` | Division by zero            |
| `#NAME?`  | Unrecognized function/name  |
| `#NUM!`   | Invalid numeric calculation |
| `#SPILL!` | Dynamic array cannot spill  |

---
