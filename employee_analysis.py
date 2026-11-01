import pandas as pd

# Step 1: Load and clean the data
file_path = "Employee Sample Data.txt"  # tab-separated input file in the same folder

try:
    df = pd.read_csv(file_path, delimiter="\t", encoding='latin1')
except FileNotFoundError:
    print("File not found. Please check the file path.")
    exit()
except UnicodeDecodeError:
    print("File encoding error. Try another encoding like 'latin1'.")
    exit()

# Clean column names
df.columns = df.columns.str.strip()

# Remove extra spaces
df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

# Clean 'Annual Salary'
df['Annual Salary'] = df['Annual Salary'].replace(r'[\$,]', '', regex=True).astype(float)

# Clean 'Bonus %'
df['Bonus %'] = df['Bonus %'].str.replace('%', '', regex=True).astype(float)

# Convert dates
df['Hire Date'] = pd.to_datetime(df['Hire Date'], errors='coerce')
df['Exit Date'] = pd.to_datetime(df['Exit Date'], errors='coerce')

# Convert Age to numeric
df['Age'] = pd.to_numeric(df['Age'], errors='coerce')

# Drop rows missing essential values
df.dropna(subset=['Full Name', 'Job Title', 'Annual Salary'], inplace=True)

# Reset index
df.reset_index(drop=True, inplace=True)

# Step 2: Replace first 5 rows with example data
example_rows = [
    {'EEID': 'E00001', 'Full Name': 'Ali Ahmad', 'Job Title': 'Engineer', 'Department': 'Engineering',
     'Business Unit': 'Manufacturing', 'Gender': 'Male', 'Ethnicity': 'Asian', 'Age': 30,
     'Hire Date': pd.to_datetime('2020-01-01'), 'Annual Salary': 75000.0, 'Bonus %': 10.0,
     'Country': 'Jordan', 'City': 'Amman', 'Exit Date': pd.NaT},

    {'EEID': 'E00002', 'Full Name': 'Sara Naser', 'Job Title': 'Analyst', 'Department': 'Finance',
     'Business Unit': 'Corporate', 'Gender': 'Female', 'Ethnicity': 'Black', 'Age': 27,
     'Hire Date': pd.to_datetime('2019-05-15'), 'Annual Salary': 64000.0, 'Bonus %': 8.0,
     'Country': 'Jordan', 'City': 'Irbid', 'Exit Date': pd.NaT},

    {'EEID': 'E00003', 'Full Name': 'Omar Saleh', 'Job Title': 'Developer', 'Department': 'IT',
     'Business Unit': 'Research & Development', 'Gender': 'Male', 'Ethnicity': 'White', 'Age': 35,
     'Hire Date': pd.to_datetime('2018-03-10'), 'Annual Salary': 90000.0, 'Bonus %': 12.0,
     'Country': 'Jordan', 'City': 'Zarqa', 'Exit Date': pd.NaT},

    {'EEID': 'E00004', 'Full Name': 'Lina Ayoub', 'Job Title': 'Manager', 'Department': 'Marketing',
     'Business Unit': 'Corporate', 'Gender': 'Female', 'Ethnicity': 'Latino', 'Age': 40,
     'Hire Date': pd.to_datetime('2016-11-23'), 'Annual Salary': 120000.0, 'Bonus %': 15.0,
     'Country': 'Jordan', 'City': 'Amman', 'Exit Date': pd.NaT},

    {'EEID': 'E00005', 'Full Name': 'Hassan Zayed', 'Job Title': 'Data Scientist', 'Department': 'IT',
     'Business Unit': 'Research & Development', 'Gender': 'Male', 'Ethnicity': 'Asian', 'Age': 29,
     'Hire Date': pd.to_datetime('2022-06-01'), 'Annual Salary': 98000.0, 'Bonus %': 9.0,
     'Country': 'Jordan', 'City': 'Aqaba', 'Exit Date': pd.NaT},
]

for i in range(5):
    df.loc[i] = example_rows[i]

# Step 3: Print the row with the largest salary
max_salary_row = df.loc[df['Annual Salary'].idxmax()]
print("\nRow with the largest salary:")
print(max_salary_row)

# Step 4: Group by Department and calculate average Age and Salary
grouped = df.groupby('Department')[['Age', 'Annual Salary']].mean().round(2)
print("\nAverage Age and Salary by Department:")
print(grouped)

# Step 5: Group by Department and Ethnicity
grouped_dept_ethnicity = df.groupby(['Department', 'Ethnicity']).agg({
    'Age': ['max', 'min'],
    'Annual Salary': 'median'
}).round(2)
print("\nMax Age, Min Age, and Median Salary by Department and Ethnicity:")
print(grouped_dept_ethnicity)

# Step 6: Save all results to an Excel file
output_file = "Employee_Analysis_Output.xlsx"
with pd.ExcelWriter(output_file) as writer:
    df.to_excel(writer, sheet_name="Cleaned Data", index=False)
    grouped.to_excel(writer, sheet_name="Avg by Department")
    grouped_dept_ethnicity.to_excel(writer, sheet_name="Dept+Ethnicity Stats")

print(f"\n✅ Results saved successfully to {output_file}")
