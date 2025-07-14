import pandas as pd
from io import BytesIO
import xlsxwriter

data = {
    "Supply Chain": pd.DataFrame({
        "Month": ["Jan", "Feb", "Mar"],
        "Purchases Made": [25, 30, 28],
        "Procurement Plan Monitoring": [80, 85, 90],
        "Special Group Contracts": [5, 6, 4],
        "Suppliers Registered": [10, 15, 12]
    }),
    "Human Resources": pd.DataFrame({
        "Month": ["Jan", "Feb", "Mar"],
        "Staff Training": [2, 3, 4],
        "Staff Welfare Activities": [1, 2, 2],
        "Complaints Received": [3, 4, 2],
        "Complaints Resolved": [2, 4, 2],
        "Students on Industrial Training": [5, 6, 7]
    }),
    "Road Asset and Corridor Management": pd.DataFrame({
        "Month": ["Jan", "Feb", "Mar"],
        "Road Works Progress (%)": [60, 75, 85],
        "Inspections Done": [8, 10, 12],
        "Achievements Reported": [5, 6, 7]
    }),
    "Transport": pd.DataFrame({
        "Month": ["Jan", "Feb", "Mar"],
        "Service and Maintenance": [10, 12, 9],
        "Fuel Consumption (Litres)": [500, 600, 550]
    }),
    "Survey": pd.DataFrame({
        "Month": ["Jan", "Feb", "Mar"],
        "Surveys Completed": [10, 12, 14],
        "Pending Reports": [2, 1, 0]
    }),
    "Finance and Accounts": pd.DataFrame({
        "Contracts Paid": ["Contract A", "Contract B", "Contract C"],
        "Amount Paid": [10000, 15000, 12000],
        "Per Diem Paid": [3000, 2500, 2700],
        "Budget Consumption": [18000, 20000, 19000]
    })
}

sheet_names_fixed = {
    "Supply Chain": "Supply Chain",
    "Human Resources": "HR",
    "Road Asset and Corridor Management": "Road Assets",
    "Transport": "Transport",
    "Survey": "Survey",
    "Finance and Accounts": "Finance"
}

output = BytesIO()
with pd.ExcelWriter("Department_Report_with_Charts.xlsx", engine='xlsxwriter') as writer:
    workbook = writer.book
    for dept, df in data.items():
        sheet_name = sheet_names_fixed[dept]
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        worksheet = writer.sheets[sheet_name]
        if "Month" in df.columns:
            chart = workbook.add_chart({'type': 'line'})
            for i, col in enumerate(df.columns[1:], start=1):
                chart.add_series({
                    'name':       [sheet_name, 0, i],
                    'categories': [sheet_name, 1, 0, len(df), 0],
                    'values':     [sheet_name, 1, i, len(df), i]
                })
            chart.set_title({'name': f'{sheet_name} - Multi-Month Trend'})
            chart.set_x_axis({'name': 'Month'})
            chart.set_y_axis({'name': 'Value'})
            worksheet.insert_chart('G2', chart)

