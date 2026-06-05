# Excel Skills Portfolio

## Overview
This repository demonstrates my Microsoft Excel skills for data analysis, reporting, and data management. I have hands-on experience using formulas, functions, and logical calculations to clean, analyze, and summarize data.

---

# Skills and Examples

## 1. SUM Function
Used to add numbers in a range.

### Formula
```excel
=SUM(B2:B10)
```

### Example
| Sales |
|---------|
| 100 |
| 200 |
| 150 |

Result:
```excel
=SUM(A2:A4)
```
Output: **450**

---

## 2. COUNT Function
Counts cells containing numbers.

### Formula
```excel
=COUNT(A2:A10)
```

### Example
| Values |
|---------|
| 10 |
| 20 |
| Text |
| 30 |

Output: **3**

---

## 3. AVERAGE Function
Calculates the average value.

### Formula
```excel
=AVERAGE(B2:B10)
```

### Example
| Marks |
|---------|
| 80 |
| 90 |
| 70 |

Output: **80**

---

## 4. MIN and MAX Functions

### MIN Formula
```excel
=MIN(A2:A10)
```

### MAX Formula
```excel
=MAX(A2:A10)
```

### Example
Values: 25, 10, 35, 15

- MIN = 10
- MAX = 35

---

## 5. IF Function
Performs logical tests.

### Formula
```excel
=IF(B2>=50,"Pass","Fail")
```

### Example

| Marks | Result |
|---------|---------|
| 75 | Pass |
| 40 | Fail |

---

## 6. SUMIF Function
Adds values based on a condition.

### Formula
```excel
=SUMIF(A2:A10,"Electronics",B2:B10)
```

### Example

| Category | Sales |
|-----------|--------|
| Electronics | 1000 |
| Furniture | 500 |
| Electronics | 1500 |

Output: **2500**

---

## 7. COUNTIF Function
Counts cells that meet a condition.

### Formula
```excel
=COUNTIF(A2:A10,"Completed")
```

### Example

| Status |
|----------|
| Completed |
| Pending |
| Completed |

Output: **2**

---

## 8. AVERAGEIF Function
Calculates average based on a condition.

### Formula
```excel
=AVERAGEIF(A2:A10,"North",B2:B10)
```

### Example

| Region | Sales |
|---------|---------|
| North | 100 |
| South | 150 |
| North | 200 |

Output: **150**

---

## 9. AVERAGEIFS Function
Calculates average using multiple conditions.

### Formula
```excel
=AVERAGEIFS(C2:C10,A2:A10,"North",B2:B10,"Electronics")
```

---

## 10. LEFT Function
Extracts characters from the left side of text.

### Formula
```excel
=LEFT(A2,3)
```

### Example

Text: `EXCEL123`

Output: `EXC`

---

## 11. RIGHT Function
Extracts characters from the right side.

### Formula
```excel
=RIGHT(A2,3)
```

### Example

Text: `EXCEL123`

Output: `123`

---

## 12. MID Function
Extracts text from the middle of a string.

### Formula
```excel
=MID(A2,2,4)
```

### Example

Text: `EXCEL123`

Output: `XCEL`

---

# Data Analysis Skills

- Data Cleaning
- Removing Duplicates
- Handling Missing Values
- Data Validation
- Conditional Formatting
- Report Generation
- Dashboard Preparation
- Data Summarization

---

# Sample Business Use Cases

### Sales Analysis
- Calculate total sales using SUM
- Find average sales by region using AVERAGEIF
- Identify highest and lowest sales using MAX and MIN

### Employee Performance Tracking
- Calculate average employee scores
- Determine Pass/Fail status using IF
- Count completed tasks using COUNTIF

### Customer Data Management
- Extract customer codes using LEFT, RIGHT, and MID
- Categorize customers based on conditions

---

# Tools

- Microsoft Excel
- GitHub
- CSV Data Processing

---

# Learning Journey

I continuously improve my Excel skills through practice, data analysis projects, and business reporting tasks. I am interested in learning advanced Excel topics such as:

- Pivot Tables
- XLOOKUP
- INDEX & MATCH
- Power Query
- Power Pivot
- Excel Dashboards
- VBA Automation

---

# Author

Your Name

GitHub Profile: https://github.com/Surya-prakash-a11y
