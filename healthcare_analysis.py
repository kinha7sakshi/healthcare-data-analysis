
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList

# ==========================================
# 1. GENERATE SYNTHETIC HEALTHCARE DATA
# ==========================================

np.random.seed(42)

n = 150

data = {
    "Patient_ID": [f"P{i:03d}" for i in range(1, n + 1)],
    "Age": np.random.randint(18, 80, n),
    "Gender": np.random.choice(["Male", "Female"], n),
    "Modality": np.random.choice(
        ["X-Ray", "CT", "MRI", "Ultrasound"], n
    ),
    "Body_Part": np.random.choice(
        ["Chest", "Head", "Abdomen", "Spine", "Knee"], n
    ),
    "Examination": np.random.choice(
        ["Routine", "Emergency", "Follow-up"], n
    ),
    "Result": np.random.choice(
        ["Normal", "Abnormal"], n, p=[0.65, 0.35]
    )
}

df = pd.DataFrame(data)

df.to_csv("healthcare_data.csv", index=False)

print("Healthcare dataset created!")
print("Total patients:", len(df))

# ==========================================
# 2. DATASET INSPECTION
# ==========================================

print("\n--- DATASET INFORMATION ---")
print("Rows:", df.shape[0])
print("Columns:", df.shape[1])

print("\n--- MISSING VALUES ---")
print(df.isnull().sum())

print("\n--- DUPLICATE RECORDS ---")
print(df.duplicated().sum())

print("\n--- DATA TYPES ---")
print(df.dtypes)

# ==========================================
# 3. HEALTHCARE STATISTICS
# ==========================================

print("\n--- HEALTHCARE STATISTICS ---")

total_patients = len(df)
average_age = round(df["Age"].mean(), 2)
youngest_age = df["Age"].min()
oldest_age = df["Age"].max()

print("Total Patients:", total_patients)
print("Average Age:", average_age)
print("Youngest Age:", youngest_age)
print("Oldest Age:", oldest_age)

# ==========================================
# 4. CATEGORY ANALYSIS
# ==========================================

def create_analysis(column, display_name):
    result = df[column].value_counts().reset_index()
    result.columns = [display_name, "Patients"]

    result["Percentage"] = (
        result["Patients"] / len(df) * 100
    ).round(2)

    return result


modality_df = create_analysis("Modality", "Modality")
body_part_df = create_analysis("Body_Part", "Body Part")
result_df = create_analysis("Result", "Result")
gender_df = create_analysis("Gender", "Gender")
examination_df = create_analysis("Examination", "Examination")

print("\n--- MODALITY ANALYSIS ---")
print(modality_df)

print("\n--- BODY PART ANALYSIS ---")
print(body_part_df)

print("\n--- RESULT ANALYSIS ---")
print(result_df)

print("\n--- GENDER ANALYSIS ---")
print(gender_df)

print("\n--- EXAMINATION ANALYSIS ---")
print(examination_df)

# ==========================================
# 5. SUMMARY INFORMATION
# ==========================================

summary_df = pd.DataFrame({
    "Metric": [
        "Total Patients",
        "Average Age",
        "Youngest Patient",
        "Oldest Patient",
        "Normal Results",
        "Abnormal Results"
    ],
    "Value": [
        total_patients,
        average_age,
        youngest_age,
        oldest_age,
        int((df["Result"] == "Normal").sum()),
        int((df["Result"] == "Abnormal").sum())
    ]
})

# ==========================================
# 6. CREATE EXCEL WORKBOOK
# ==========================================

file_name = "healthcare_analysis_report.xlsx"

with pd.ExcelWriter(file_name, engine="openpyxl") as writer:

    # Dashboard
    summary_df.to_excel(
        writer, sheet_name="Dashboard", index=False,
        startrow=2
    )

    # Patient records
    df.to_excel(
        writer, sheet_name="Patient_Data", index=False
    )

    # Summary
    summary_df.to_excel(
        writer, sheet_name="Summary", index=False
    )

    # Category analysis
    modality_df.to_excel(
        writer, sheet_name="Modality_Analysis", index=False
    )

    body_part_df.to_excel(
        writer, sheet_name="Body_Part_Analysis", index=False
    )

    result_df.to_excel(
        writer, sheet_name="Result_Analysis", index=False
    )

    gender_df.to_excel(
        writer, sheet_name="Gender_Analysis", index=False
    )

    examination_df.to_excel(
        writer, sheet_name="Examination_Analysis", index=False
    )

print("\nExcel workbook created!")

# ==========================================
# 7. FORMAT EXCEL WORKBOOK
# ==========================================

workbook = load_workbook(file_name)

for sheet in workbook.worksheets:

    # Format column headings
    header_row = 3 if sheet.title == "Dashboard" else 1

    for cell in sheet[header_row]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78"
        )
        cell.alignment = Alignment(horizontal="center")

    # Adjust column widths
    for column in sheet.columns:
        max_length = 0
        column_letter = column[0].column_letter

        for cell in column:
            if cell.value is not None:
                max_length = max(
                    max_length, len(str(cell.value))
                )

        sheet.column_dimensions[column_letter].width = min(
            max_length + 3, 30
        )

    sheet.freeze_panes = f"A{header_row + 1}"

    # Enable filters on data tables
    sheet.auto_filter.ref = sheet.dimensions

# Dashboard title
dashboard = workbook["Dashboard"]

dashboard.merge_cells("A1:D1")
dashboard["A1"] = "HEALTHCARE DATA ANALYSIS DASHBOARD"

dashboard["A1"].font = Font(
    bold=True,
    color="FFFFFF",
    size=16
)

dashboard["A1"].fill = PatternFill(
    fill_type="solid",
    fgColor="17365D"
)

dashboard["A1"].alignment = Alignment(
    horizontal="center",
    vertical="center"
)

dashboard.row_dimensions[1].height = 35

# Set dashboard column widths
dashboard.column_dimensions["A"].width = 24
dashboard.column_dimensions["B"].width = 20
dashboard.column_dimensions["C"].width = 4
dashboard.column_dimensions["D"].width = 20

# ==========================================
# 8. ADD EXCEL CHARTS TO DASHBOARD
# ==========================================

def add_bar_chart(sheet_name, title, anchor):

    source_sheet = workbook[sheet_name]

    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = title
    chart.y_axis.title = "Number of Patients"
    chart.x_axis.title = "Category"
    chart.height = 8
    chart.width = 14

    data_ref = Reference(
        source_sheet,
        min_col=2,
        min_row=1,
        max_row=source_sheet.max_row
    )

    category_ref = Reference(
        source_sheet,
        min_col=1,
        min_row=2,
        max_row=source_sheet.max_row
    )

    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(category_ref)

    dashboard.add_chart(chart, anchor)


# Modality chart
add_bar_chart(
    "Modality_Analysis",
    "Patients by Imaging Modality",
    "D3"
)

# Body part chart
add_bar_chart(
    "Body_Part_Analysis",
    "Patients by Body Part",
    "L3"
)

# Result distribution pie chart
pie = PieChart()
pie.title = "Normal vs Abnormal Results"
pie.height = 8
pie.width = 14

result_sheet = workbook["Result_Analysis"]

pie_data = Reference(
    result_sheet,
    min_col=2,
    min_row=1,
    max_row=result_sheet.max_row
)

pie_labels = Reference(
    result_sheet,
    min_col=1,
    min_row=2,
    max_row=result_sheet.max_row
)

pie.add_data(pie_data, titles_from_data=True)
pie.set_categories(pie_labels)

pie.dataLabels = DataLabelList()
pie.dataLabels.showPercent = True
pie.dataLabels.showLegendKey = False

dashboard.add_chart(pie, "D20")

# Save final workbook
workbook.save(file_name)

print("Excel formatting completed!")
print("Excel charts added to Dashboard!")

# ==========================================
# 9. CREATE MATPLOTLIB VISUALIZATIONS
# ==========================================

fig, axes = plt.subplots(3, 1, figsize=(9, 15))

# Imaging modality
modality_counts = df["Modality"].value_counts()

bars1 = axes[0].bar(
    modality_counts.index,
    modality_counts.values
)

axes[0].bar_label(bars1)
axes[0].set_title("Patients by Imaging Modality")
axes[0].set_ylabel("Number of Patients")

# Body part
body_part_counts = df["Body_Part"].value_counts()

bars2 = axes[1].bar(
    body_part_counts.index,
    body_part_counts.values
)

axes[1].bar_label(bars2)
axes[1].set_title("Patients by Body Part")
axes[1].set_ylabel("Number of Patients")

# Result distribution
result_counts = df["Result"].value_counts()

axes[2].pie(
    result_counts.values,
    labels=result_counts.index,
    autopct="%1.1f%%",
    startangle=90
)

axes[2].set_title("Normal vs Abnormal Results")

plt.tight_layout()

# Save graph image
plt.savefig(
    "healthcare_visualizations.png",
    dpi=200,
    bbox_inches="tight"
)

print("\nVisualization image saved!")

# Display graphs
plt.show()
plt.close(fig)

# ==========================================
# 10. FINAL PROJECT STATUS
# ==========================================

print("\n========== PROJECT COMPLETED ==========")
print("1. healthcare_data.csv")
print("2. healthcare_analysis_report.xlsx")
print("3. healthcare_visualizations.png")
print("=======================================")
