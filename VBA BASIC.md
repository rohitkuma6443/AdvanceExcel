## Step 1 — Create a Macro

Open Excel → press **Alt + F11**

Then:

```text
Insert → Module
```

Write:

```vba
Sub MyFirstMacro()

    MsgBox "Hello World"

End Sub
```

Run with:

```text
F5
```

### Output

```text
Hello World
```

---

# Step 2 — Understand the Basic Structure

Every basic Macro looks like:

```vba
Sub MacroName()

    'Code goes here

End Sub
```

Example:

```vba
Sub Test()

    MsgBox "Welcome to VBA"

End Sub
```

### Important

```vba
Sub
```

starts the Macro.

```vba
End Sub
```

ends the Macro.

---

# Step 3 — Write Data into a Cell

```vba
Sub WriteData()

    Range("A1").Value = "Rohit"

End Sub
```

Run it.

Excel:

| A     |
| ----- |
| Rohit |

---

# Step 4 — Write Multiple Cells

```vba
Sub WriteData()

    Range("A1").Value = "Name"
    Range("B1").Value = "Age"
    Range("C1").Value = "City"

End Sub
```

Output:

| Name | Age | City |
| ---- | --: | ---- |
|      |     |      |

---

# Step 5 — Write Different Types of Data

```vba
Sub WriteData()

    Range("A1").Value = "Rohit"
    Range("B1").Value = 30
    Range("C1").Value = 50000
    Range("D1").Value = True

End Sub
```

Here:

```text
"Rohit"     → Text
30          → Number
50000       → Number
True        → Boolean
```

---

# Step 6 — Read Data from a Cell

Suppose:

```text
A1 = Rohit
```

Code:

```vba
Sub ReadData()

    MsgBox Range("A1").Value

End Sub
```

Output:

```text
Rohit
```

---

# Step 7 — Copy Data from One Cell to Another

```vba
Sub CopyData()

    Range("B1").Value = Range("A1").Value

End Sub
```

If:

```text
A1 = Rohit
```

then:

```text
B1 = Rohit
```

---

# Step 8 — Use `Cells`

Instead of:

```vba
Range("A1").Value = "Rohit"
```

you can write:

```vba
Cells(1, 1).Value = "Rohit"
```

The syntax is:

```vba
Cells(row, column)
```

Examples:

```vba
Cells(1, 1)   'A1
Cells(1, 2)   'B1
Cells(2, 1)   'A2
Cells(5, 3)   'C5
```

Example:

```vba
Sub Test()

    Cells(1, 1).Value = "Name"
    Cells(1, 2).Value = "Sales"

End Sub
```

---

# Step 9 — Variables

Use `Dim` to create a variable.

```vba
Sub VariableExample()

    Dim name As String

    name = "Rohit"

    MsgBox name

End Sub
```

Another example:

```vba
Sub VariableExample()

    Dim sales As Double

    sales = 50000

    MsgBox sales

End Sub
```

---

# Step 10 — Use Variables with Excel

Suppose:

```text
A2 = Rohit
B2 = 50000
```

Code:

```vba
Sub ReadSales()

    Dim customer As String
    Dim sales As Double

    customer = Range("A2").Value
    sales = Range("B2").Value

    MsgBox customer & " Sales = " & sales

End Sub
```

Output:

```text
Rohit Sales = 50000
```

---

# Step 11 — If Statement

Suppose B2 contains sales.

```vba
Sub CheckSales()

    If Range("B2").Value >= 50000 Then

        MsgBox "Good Sales"

    End If

End Sub
```

---

# Step 12 — If Else

```vba
Sub CheckSales()

    If Range("B2").Value >= 50000 Then

        MsgBox "Good Sales"

    Else

        MsgBox "Low Sales"

    End If

End Sub
```

---

# Step 13 — Write Result into Excel

Instead of displaying a message:

```vba
Sub CheckSales()

    If Range("B2").Value >= 50000 Then

        Range("C2").Value = "Good"

    Else

        Range("C2").Value = "Low"

    End If

End Sub
```

Dataset:

| Customer | Sales | Status |
| -------- | ----: | ------ |
| Rohit    | 75000 | Good   |

---

# Step 14 — Multiple Conditions

Use `ElseIf`.

```vba
Sub CheckSales()

    If Range("B2").Value >= 100000 Then

        Range("C2").Value = "Excellent"

    ElseIf Range("B2").Value >= 50000 Then

        Range("C2").Value = "Good"

    Else

        Range("C2").Value = "Low"

    End If

End Sub
```

Logic:

```text
100000 or more → Excellent
50000–99999    → Good
Below 50000    → Low
```

---

# Step 15 — For Loop

Now we start real automation.

```vba
Sub NumberSeries()

    Dim i As Integer

    For i = 1 To 10

        Cells(i, 1).Value = i

    Next i

End Sub
```

Result:

|  A |
| -: |
|  1 |
|  2 |
|  3 |
|  4 |
|  5 |
|  6 |
|  7 |
|  8 |
|  9 |
| 10 |

---

# Step 16 — Loop Through Sales Data

Create:

| Customer | Sales |
| -------- | ----: |
| Amit     | 45000 |
| Rahul    | 75000 |
| Priya    | 30000 |
| Neha     | 90000 |
| Raj      | 55000 |

Now:

```vba
Sub CheckAllSales()

    Dim i As Integer

    For i = 2 To 6

        If Cells(i, 2).Value >= 50000 Then

            Cells(i, 3).Value = "Good"

        Else

            Cells(i, 3).Value = "Low"

        End If

    Next i

End Sub
```

Result:

| Customer | Sales | Status |
| -------- | ----: | ------ |
| Amit     | 45000 | Low    |
| Rahul    | 75000 | Good   |
| Priya    | 30000 | Low    |
| Neha     | 90000 | Good   |
| Raj      | 55000 | Good   |

This is the **core concept of Excel automation**:

```text
Read Excel
   ↓
Process data
   ↓
Apply logic
   ↓
Write result
```

---

# Step 17 — Find Last Row

This is very important for professional VBA.

Instead of:

```vba
For i = 2 To 6
```

we can automatically find the last row.

```vba
Sub FindLastRow()

    Dim lastRow As Long

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    MsgBox lastRow

End Sub
```

If your data ends at row 100:

```text
lastRow = 100
```

---

# Step 18 — Dynamic Sales Automation

Now combine everything:

```vba
Sub CheckAllSales()

    Dim i As Long
    Dim lastRow As Long

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow

        If Cells(i, 2).Value >= 50000 Then

            Cells(i, 3).Value = "Good"

        Else

            Cells(i, 3).Value = "Low"

        End If

    Next i

End Sub
```

Now you can add 10, 100, or 10,000 rows.

The Macro automatically finds the last row.

---

# Step 19 — Formatting

### Bold

```vba
Range("A1:C1").Font.Bold = True
```

### Font Size

```vba
Range("A1:C1").Font.Size = 14
```

### AutoFit

```vba
Columns("A:C").AutoFit
```

### Borders

```vba
Range("A1:C10").Borders.LineStyle = xlContinuous
```

### Center

```vba
Range("A1:C1").HorizontalAlignment = xlCenter
```

---

# Step 20 — Complete Small Project

Dataset:

| Customer | Sales | Quantity |
| -------- | ----: | -------: |
| Amit     | 45000 |        5 |
| Rahul    | 75000 |        8 |
| Priya    | 30000 |        3 |
| Neha     | 90000 |       10 |
| Raj      | 55000 |        6 |

Macro:

```vba
Option Explicit

Sub SalesReport()

    Dim i As Long
    Dim lastRow As Long

    'Add heading
    Range("D1").Value = "Status"

    'Find last row
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    'Check sales
    For i = 2 To lastRow

        If Cells(i, 2).Value >= 50000 Then

            Cells(i, 4).Value = "Good"

        Else

            Cells(i, 4).Value = "Low"

        End If

    Next i

    'Format heading
    Range("A1:D1").Font.Bold = True

    'Add borders
    Range("A1:D" & lastRow).Borders.LineStyle = xlContinuous

    'AutoFit
    Columns("A:D").AutoFit

    MsgBox "Sales Report Completed"

End Sub
```





```text
1. MsgBox
2. InputBox
3. Range
4. Cells
5. Variables
6. Data Types
7. If / Else
8. Select Case
9. For Loop
10. For Each Loop
11. Do While / Do Until
12. Last Row / Last Column
13. Worksheets
14. Workbooks
15. Range Operations
16. Find / Replace
17. Sort / Filter
18. Excel Tables
19. Functions
20. Arrays
21. Dictionary
22. Error Handling
23. Events
24. UserForms
25. Buttons
26. File Handling
27. Email Automation
28. PDF Automation
29. MIS Automation Project
30. Complete VBA Project
```

**Start with Step 1: `MsgBox`** and master each topic before moving to the next.
