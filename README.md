Healthcare Data Analysis Using Python

Project Overview

This project focuses on analysing synthetic healthcare data using Python. It generates patient records, performs data analysis, creates an automated Excel report, and produces data visualizations.

Technologies Used

- Python
- Pandas
- NumPy
- Matplotlib
- OpenPyXL

Dataset Description

The project generates 150 synthetic patient records with the following information:

- Patient ID
- Age
- Gender
- Imaging Modality
- Body Part
- Examination Type
- Examination Result

Key Features

1. Synthetic Data Generation

- Generates 150 synthetic patient records.
- Uses NumPy for data generation.
- Exports the dataset to CSV format.

2. Data Inspection

- Checks rows and columns.
- Identifies missing values.
- Checks duplicate records.
- Displays data types.

3. Healthcare Statistics

Calculates:

- Total number of patients.
- Average patient age.
- Youngest patient age.
- Oldest patient age.

4. Category Analysis

Analyses patient distribution by:

- Imaging modality.
- Body part.
- Examination result.
- Gender.
- Examination type.

Calculates patient counts and percentages for each category.

5. Automated Excel Report

Creates an Excel workbook containing eight worksheets:

1. Dashboard
2. Patient_Data
3. Summary
4. Modality_Analysis
5. Body_Part_Analysis
6. Result_Analysis
7. Gender_Analysis
8. Examination_Analysis

6. Excel Dashboard

The dashboard includes:

- Healthcare summary statistics.
- Patients by imaging modality chart.
- Patients by body part chart.
- Normal vs Abnormal results pie chart.

7. Data Visualization

Creates three visualizations using Matplotlib:

- Patients by Imaging Modality.
- Patients by Body Part.
- Normal vs Abnormal Results.

Project Files

File Name| Description
healthcare_analysis.py| Python source code
healthcare_data.csv| Synthetic healthcare dataset
healthcare_analysis_report.xlsx| Automated Excel report and dashboard
healthcare_visualizations.png| Data visualization image

Libraries Required

Install the required Python libraries:

pip install pandas numpy matplotlib openpyxl

How to Run

1. Clone or download this repository.
2. Install the required libraries.
3. Open the project folder.
4. Run the Python script:

python healthcare_analysis.py

The script automatically generates the CSV dataset, Excel report, and visualization image.

Learning Outcomes

Through this project, I practised:

- Python programming.
- Data generation using NumPy.
- Data manipulation using Pandas.
- Exploratory data analysis.
- Statistical calculations.
- Excel report automation using OpenPyXL.
- Data visualization using Matplotlib.

Future Improvements

- Interactive healthcare dashboards.
- Advanced statistical analysis.
- More detailed imaging modality analysis.
- Interactive data visualizations.
- Machine learning experiments using synthetic healthcare datasets.

Disclaimer

This project is developed for educational and portfolio purposes only.

All patient records and examination results are synthetically generated. They do not represent real patients or actual clinical findings.

This project is not intended for medical diagnosis or clinical decision-making.

Author

Sakshi


Exploring the application of Python, data analysis, automation, and visualization in healthcare.
