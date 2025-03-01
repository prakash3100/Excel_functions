# **Excel Functions for Business Data Analysis** 📊

## **Introduction**  
This document provides an overview of key Excel functions used for business data analysis, specifically in salary and employee-related queries. It includes functions for calculations, lookups, filtering, and data summarization.

---

## **Table of Contents**  
1. [Total Salary and Headcount by Department](#1-total-salary-and-headcount-by-department)  
2. [Average Salary by Department](#2-average-salary-by-department)  
3. [Employees Earning More than $100k](#3-employees-earning-more-than-100k)  
4. [Filtering Female Employees Earning More than $100k](#4-filtering-female-employees-earning-more-than-100k)  
5. [Filtering Female Employees (Specific Join Year)](#5-filtering-female-employees-specific-join-year)  
6. [Lowest, Highest, and Top 5 Salaries](#6-lowest-highest-and-top-5-salaries)  
7. [Lowest, Highest, and Top 5 Salaries by Gender](#7-lowest-highest-and-top-5-salaries-by-gender)  
8. [List of All Departments](#8-list-of-all-departments)  
9. [List of All Departments in One Cell](#9-list-of-all-departments-in-one-cell)  
10. [Employee Details Lookup](#10-employee-details-lookup)  
11. [Finding the Highest Paid Employee](#11-finding-the-highest-paid-employee)  
12. [Employees Who Joined in March](#12-employees-who-joined-in-march)  
13. [Female Employees Who Started on a Monday](#13-female-employees-who-started-on-a-monday)  
14. [Department-wise Salary and Headcount Analysis](#14-department-wise-salary-and-headcount-analysis)  
15. [Median Salary and Female Ratio Calculation](#15-median-salary-and-female-ratio-calculation)  

---

## **1. Total Salary and Headcount by Department**  
📌 **Functions Used:** `SUMIF`, `COUNTIF`  
🔹 **Purpose:** Calculate the total salary and number of employees in each department.  

```excel
=SUMIF(DepartmentRange, "Sales", SalaryRange)  // Total salary for Sales
=COUNTIF(DepartmentRange, "Sales")  // Number of employees in Sales
```

---

## **2. Average Salary by Department**  
📌 **Functions Used:** `AVERAGEIF`, `AVERAGEIFS`  
🔹 **Purpose:** Find the average salary for each department.  

```excel
=AVERAGEIF(DepartmentRange, "HR", SalaryRange)
```

---

## **3. Employees Earning More than $100k**  
📌 **Functions Used:** `FILTER`, `CHOOSECOLS`  
🔹 **Purpose:** Extract employees whose salary is greater than $100,000.  

```excel
=FILTER(DataRange, SalaryRange > 100000)
```

---

## **4. Filtering Female Employees Earning More than $100k**  
📌 **Functions Used:** `FILTER`, `*`  
🔹 **Purpose:** Filter employees who are female and earn more than $100k.  

```excel
=FILTER(DataRange, (SalaryRange > 100000) * (GenderRange = "Female"))
```

---

## **5. Filtering Female Employees (Specific Join Year)**  
📌 **Functions Used:** `FILTER`  
🔹 **Purpose:** Find female employees who earn more than $100k and joined in 2020 or later.  

```excel
=FILTER(DataRange, (SalaryRange > 100000) * (GenderRange = "Female") * (JoinYearRange >= 2020))
```

---

## **6. Lowest, Highest, and Top 5 Salaries**  
📌 **Functions Used:** `MIN`, `MAX`, `LARGE`, `SORT`, `TAKE`  
🔹 **Purpose:** Identify lowest, highest, and top 5 salaries.  

```excel
=MIN(SalaryRange) // Lowest salary
=MAX(SalaryRange) // Highest salary
=LARGE(SalaryRange, 5) // 5th highest salary
```

---

## **7. Lowest, Highest, and Top 5 Salaries by Gender**  
📌 **Functions Used:** `MINIFS`, `MAXIFS`  
🔹 **Purpose:** Find salaries based on gender criteria.  

```excel
=MINIFS(SalaryRange, GenderRange, "Female")  // Lowest salary for females
=MAXIFS(SalaryRange, GenderRange, "Male")    // Highest salary for males
```

---

## **8. List of All Departments**  
📌 **Functions Used:** `UNIQUE`, `COUNTA`, `SORT`  
🔹 **Purpose:** Get a unique list of departments.  

```excel
=SORT(UNIQUE(DepartmentRange))
```

---

## **9. List of All Departments in One Cell**  
📌 **Functions Used:** `TEXTJOIN`  
🔹 **Purpose:** Combine department names into a single cell.  

```excel
=TEXTJOIN(", ", TRUE, UNIQUE(DepartmentRange))
```

---

## **10. Employee Details Lookup**  
📌 **Functions Used:** `VLOOKUP`, `INDEX + MATCH`, `XLOOKUP`, `IFERROR`  
🔹 **Purpose:** Retrieve employee details using lookup functions.  

```excel
=XLOOKUP(EmployeeID, IDRange, NameRange, "Not Found")
```

---

## **11. Finding the Highest Paid Employee**  
📌 **Functions Used:** `XLOOKUP`, `MAX`  
🔹 **Purpose:** Find the employee with the highest salary.  

```excel
=XLOOKUP(MAX(SalaryRange), SalaryRange, EmployeeNameRange)
```

---

## **12. Employees Who Joined in March**  
📌 **Functions Used:** `FILTER`, `MONTH`  
🔹 **Purpose:** Filter employees who joined in March.  

```excel
=FILTER(DataRange, MONTH(JoinDateRange) = 3)
```

---

## **13. Female Employees Who Started on a Monday**  
📌 **Functions Used:** `FILTER`, `WEEKDAY`  
🔹 **Purpose:** Find female employees who started on a Monday.  

```excel
=FILTER(DataRange, (GenderRange = "Female") * (WEEKDAY(JoinDateRange) = 2))
```

---

## **14. Department-wise Salary and Headcount Analysis**  
📌 **Functions Used:** `UNIQUE`, `SUMIFS`, `COUNTIFS`, `#`, `CONDITIONAL FORMATTING`  
🔹 **Purpose:** Generate a summary report.  

```excel
=SUMIFS(SalaryRange, DepartmentRange, "Sales")  // Total salary for Sales
=COUNTIFS(DepartmentRange, "Sales")  // Employee count in Sales
```

---

## **15. Median Salary and Female Ratio Calculation**  
📌 **Functions Used:** `MEDIAN`, `COUNTIFS`  
🔹 **Purpose:** Calculate median salary and female employee ratio.  

```excel
=MEDIAN(SalaryRange)  // Median salary
=COUNTIFS(GenderRange, "Female") / COUNTA(GenderRange)  // Female ratio
```

---

## **Conclusion**  
This document covers essential Excel functions for analyzing employee salary and department data. These formulas help in efficient data analysis, reporting, and decision-making. 🚀
