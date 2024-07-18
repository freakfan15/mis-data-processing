import pandas as pd

# Load the input and output Excel files
input_file_path = 'input.xlsx'
output_file_path = 'output.xlsx'

input_df = pd.read_excel(input_file_path)
output_template_df = pd.read_excel(output_file_path, header=None)

# Display column names to check for correct references
print("Column names in the input file:", input_df.columns)

# Function to generate a report for a single row
def generate_report(row):
    report_df = output_template_df.copy()
    report_df.iloc[2, 1] = row["Customer Name"]
    report_df.iloc[2, 3] = row["Customer's Father/Husband name"]
    report_df.iloc[4, 1] = row["F.I. Send Date"]
    report_df.iloc[4, 3] = row["F.I. Send Time"]
    report_df.iloc[4, 5] = row["Agency Name"]
    report_df.iloc[2, 7] = row["Area"]
    report_df.iloc[2, 9] = row["LEAD ID "]
    report_df.iloc[2, 14] = row["Customer Remark(Negative/Positive) / GTR Remark(Negative/Positive)"]
    return report_df

if not input_df.empty:
    first_row = input_df.iloc[0]
    report = generate_report(first_row)

    # Save the report to a new Excel file
    output_report_path = './generated_report_first_row.xlsx'
    with pd.ExcelWriter(output_report_path) as writer:
        report.to_excel(writer, sheet_name='Report', index=False, header=False)

# Generate reports for all rows
# reports = []
# for _, row in input_df.iterrows():
#     if pd.notna(row["LEAD ID "]):  # Skip rows with NaN values
#         report = generate_report(row)
#         reports.append(report)

# # Save the reports to separate sheets in a new Excel file
# output_reports_path = './generated_reports.xlsx'
# with pd.ExcelWriter(output_reports_path) as writer:
#     for i, report in enumerate(reports):
#         report.to_excel(writer, sheet_name=f'Report_{i+1}', index=False, header=False)