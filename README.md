# 📊 Excel Lookup Functions – VLOOKUP, HLOOKUP, INDEX & XLOOKUP

## 📌 Project Overview

This project demonstrates the most commonly used **Excel Lookup Functions** for retrieving information from datasets. It includes practical examples using employee records to show how different lookup functions work.

The project is designed for beginners and aspiring Data Analysts to understand lookup techniques used in Excel for data retrieval and analysis.

---

## 🎯 Objectives

- Learn Vertical Lookup (VLOOKUP)
- Learn Horizontal Lookup (HLOOKUP)
- Retrieve values using INDEX
- Perform flexible searches using XLOOKUP
- Compare different lookup methods

---

## 🛠 Tools Used

- Microsoft Excel
- Excel Formulas

---

## 📂 Dataset

The sample dataset contains employee information:

| Column |
|---------|
| Employee ID |
| Location |
| Email |
| Phone Number |

Example:

| Employee ID | Location | Email | Phone |
|-------------|----------|--------------------|------------|
| E1023 | Los Angeles | user23@example.com | 9876512904 |
| E1045 | Los Angeles | user45@example.com | 9876596399 |
| E1006 | New York | user6@example.com | 9876530256 |

---

# 📖 Lookup Functions

## 1️⃣ VLOOKUP (Vertical Lookup)

### Purpose
Searches for a value in the **first column** of a table and returns a value from another column.

### Formula

```excel
=VLOOKUP(A2,$A$2:$D$9,2,FALSE)
```

### Returns

- Location
- Email
- Phone Number

---

## 2️⃣ HLOOKUP (Horizontal Lookup)

### Purpose

Searches across the **first row** of a table and returns a value from a specified row.

### Formula

```excel
=HLOOKUP("Location",A1:D9,2,FALSE)
```

---

## 3️⃣ INDEX Function

### Purpose

Returns the value at a specified row and column.

### Formula

```excel
=INDEX(B2:B6,3)
```

Or with MATCH

```excel
=INDEX(B2:B6,MATCH(E1002,A2:A6,0))
```

---

## 4️⃣ XLOOKUP

### Purpose

Modern replacement for VLOOKUP and HLOOKUP.

### Formula

```excel
=XLOOKUP(F2,C2:C9,A2:A9)
```

Example:

Search using Email and return:

- Employee ID
- Location
- Phone Number

---

# 📈 Key Features

- Employee Lookup by ID
- Email Lookup
- Location Retrieval
- Phone Number Retrieval
- INDEX + MATCH Example
- XLOOKUP Example
- Beginner Friendly

---

# 📊 Skills Demonstrated

- Excel Formulas
- Data Lookup
- Data Retrieval
- Spreadsheet Analysis
- Data Cleaning Basics

---

# 📁 Repository structure

```
Excel-Lookup-Functions/
│
├── README.md
├── Lookup_Functions.xlsx
├── Screenshots/
│   ├── VLOOKUP.png
│   ├── HLOOKUP.png
│   ├── INDEX.png
│   └── XLOOKUP.png
└── Dataset/
    └── Employee_Data.xlsx
```


# 🚀 Learning Outcome

After completing this project, you will understand:

- Difference between VLOOKUP and HLOOKUP
- When to use INDEX
- Why XLOOKUP is more powerful
- Lookup techniques used in Data Analytics

---

# ⭐ Future Improvements

- Add INDEX + MATCH examples
- Dynamic lookup using Drop-down Lists
- Error Handling using IFERROR
- Multiple Criteria Lookup
- Interactive Dashboard

---

## 👨‍💻 Author

**Bhakya M**

Aspiring Data Analyst

### Skills

- Excel
- Power BI
-

---

## ⭐ If you found this project helpful

Please ⭐ Star this repository and share it with others learning Excel and Data Analytics.
