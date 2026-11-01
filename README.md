# 📊 Employee Data Analysis

Data cleaning and exploratory analysis of an HR employee dataset with **Python and Pandas**. The raw data is cleaned, analysed by department and demographic group, and exported to a multi-sheet Excel report.

## 🔍 What it does

**1. Data cleaning**
- Strips whitespace from column names and text values
- Converts `Annual Salary` (`$75,000` → `75000.0`) and `Bonus %` (`10%` → `10.0`) to numbers
- Parses `Hire Date` and `Exit Date` as dates, and `Age` as numeric
- Drops rows missing essential fields (name, job title, salary)

**2. Analysis**
- Finds the employee with the highest salary
- Average age and salary per department
- Maximum and minimum age, and median salary, per department and ethnicity

**3. Output**
- `Employee_Analysis_Output.xlsx`, with three sheets: *Cleaned Data*, *Avg by Department*, *Dept+Ethnicity Stats*

## 🛠️ Tech stack

Python · Pandas · openpyxl (Excel export)

## ▶️ Run

```bash
pip install pandas openpyxl
python employee_analysis.py
```

Place the input file (`Employee Sample Data.txt`, tab-separated) in the same folder as the script, or update `file_path` at the top of the script.

## 📁 Files

```
employee_analysis.py           cleaning + analysis script
Employee_Analysis_Output.xlsx  generated report
```

---

Developed by **Abdullah Asasfeh** as part of a data analysis task.
