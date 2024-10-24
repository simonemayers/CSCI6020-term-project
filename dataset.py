from fpdf import FPDF
import pandas as pd
from tabulate import tabulate
import numpy as np

# Load the uploaded Excel file to inspect the sheet names first
file_path = './Preliminary-2020-National-Summary-Clinic-Table.xlsx'
excel_data = pd.ExcelFile(file_path)

# Let's focus on the 'Clinic Table Data Records' sheet first and load it to analyze the data quality
clinic_table_data = excel_data.parse('Clinic Table Data Records')

# Replace all occurrences of NaN with empty strings
clinic_table_data.replace(to_replace=np.nan, value="", inplace=True)

# Replace "*" with empty strings across all columns, treating all as strings
clinic_table_data = clinic_table_data.astype(str).replace(r'\*', '', regex=True)

# Check each column to see if it can be converted to a numeric type
for col in clinic_table_data.columns:
    cleaned_col = clinic_table_data[col].replace({',': '', '%': ''}, regex=True).str.strip()

    # Skip completely empty columns
    if cleaned_col.eq("").all():
        continue

    try:
        converted_col = pd.to_numeric(cleaned_col, errors='coerce')
        if not converted_col.isnull().all():
            clinic_table_data[col] = converted_col
    except Exception as e:
        pass

# Load the dictionary sheet to rename columns
dictionary_data = excel_data.parse('Clinic Table Dictionary')

# Assuming the dictionary sheet has columns named 'Variable' and 'Variable Description'
column_mapping = dict(zip(dictionary_data['Variable'], dictionary_data['Variable Description']))

# Get the first column name from clinic_table_data
first_column_name = clinic_table_data.columns[0]

# Exclude the first column from the renaming process
columns_to_rename = {col: column_mapping[col] for col in clinic_table_data.columns[1:] if col in column_mapping}

# Rename the columns in the main data using the mapping (skipping the first column)
clinic_table_data.rename(columns=columns_to_rename, inplace=True)

# Data Quality Report Preparation
data_quality_report = {
    'Column Name': clinic_table_data.columns,
    'Data Type': clinic_table_data.dtypes,
    'Missing Values': clinic_table_data.isnull().sum(),
    'Unique Values': clinic_table_data.nunique()
}

data_quality_report_df = pd.DataFrame(data_quality_report)

# Improved PDF Generation
class PDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 14)
        self.cell(0, 10, "Data Quality Report - Clinic Table Data Records", border=0, ln=1, align="C")
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, f"Page {self.page_no()}", 0, 0, "C")

    def create_table(self, dataframe):
        self.set_font("Arial", size=10)
        col_widths = [50, 40, 30, 30]  # Set column widths to ensure proper layout

        # Add headers
        self.set_font("Arial", "B", 12)
        for i, col_name in enumerate(dataframe.columns):
            self.cell(col_widths[i], 10, col_name, border=1, align="C")
        self.ln()

        # Add rows
        self.set_font("Arial", size=10)
        for row in dataframe.itertuples(index=False):
            self.cell(col_widths[0], 10, str(row[0]), border=1)
            self.cell(col_widths[1], 10, str(row[1]), border=1)
            self.cell(col_widths[2], 10, str(row[2]), border=1, align="C")
            self.cell(col_widths[3], 10, str(row[3]), border=1, align="C")
            self.ln()

# Create PDF and add table
pdf = PDF()
pdf.add_page()
pdf.create_table(data_quality_report_df)

# Save the PDF
pdf_output_path = "./data_quality_report_clinic_table.pdf"
pdf.output(pdf_output_path)

# Indicate the path of the saved PDF file
pdf_output_path

# Separate the columns into numerical and categorical
numerical_cols = clinic_table_data.select_dtypes(include=[np.number])
categorical_cols = clinic_table_data.select_dtypes(include=[object])

# Data Quality Report for Numerical Columns
numerical_report = pd.DataFrame({
    'Column Name': numerical_cols.columns,
    'Data Type': numerical_cols.dtypes,
    'Missing Values (%)': (numerical_cols.isnull().sum() / len(clinic_table_data)).round(2) * 100,
    'Mean': numerical_cols.mean().round(2),
    'Median': numerical_cols.median().round(2),
    'Std Dev': numerical_cols.std().round(2),
    'Min': numerical_cols.min().round(2),
    'Max': numerical_cols.max().round(2),
    'Unique Values': numerical_cols.nunique()
}).fillna("")

# Data Quality Report for Categorical Columns
categorical_report = pd.DataFrame({
    'Column Name': categorical_cols.columns,
    'Data Type': categorical_cols.dtypes,
    'Missing Values (%)': (categorical_cols.isnull().sum() / len(clinic_table_data)) * 100,
    'Unique Values': categorical_cols.nunique(),
    'Most Frequent Value': categorical_cols.mode().iloc[0],
    'Frequency of Most Frequent': categorical_cols.apply(lambda col: col.value_counts().iloc[0] if not col.value_counts().empty else "")
}).fillna("")

# PDF Report Generation
class PDF(FPDF):
    #put in landscape mode 
    def __init__(self, orientation='L'):
        super().__init__(orientation, 'mm', 'A4')
        self.set_auto_page_break(auto=True, margin=15)

    def header(self):
        self.set_font("Arial", "B", 12)
        self.cell(0, 10, "Data Quality Report - Clinic Table Data Records", border=0, ln=True, align="C")
        #self.ln(10)
        # self.set_font("Arial", "B", 12)
        # for i, col_name in enumerate(dataframe.columns):
            # self.cell(self.col_widths[i], 10, col_name, border=1, align="C")
        # self.ln(10)  # Move to the next line after headers
        #for i, col_name in enumerate(dataframe.columns):
        #    self.cell(col_widths[i], 10, col_name, border=1, align="C")
        #self.ln()

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, f"Page {self.page_no()}", 0, 0, "C")

    def add_section_title(self, title):
        self.set_font("Arial", "B", 12)
        self.cell(0, 10, title, ln=True, align="L")
        # self.ln(5)

    def create_table(self, dataframe, section_title):
        # self.add_section_title(section_title)
        self.set_font("Arial", size=10)
        col_widths = [self.get_string_width(str(col)) + 10 for col in dataframe.columns]
        col_widths = [max(width, 30) for width in col_widths]
        col_names = dataframe.columns


        # Add headers
        #self.set_font("Arial", "B", 12)
        for i, col_name in enumerate(dataframe.columns):
            self.cell(col_widths[i], 10, col_name, border=1, align="C")
        
        #self.ln()
        # self.header()

        # Add rows
        self.set_font("Arial", size=10)
        for row in dataframe.itertuples(index=False):
            for i, cell_value in enumerate(row):
                self.cell(col_widths[i], 10, str(cell_value), border=1, align="C")
        #    self.ln()

        # Add rows
        for row in dataframe.itertuples(index=False):
            max_height = 10  # Default cell height
            for i, cell_value in enumerate(row):
                cell_text = str(cell_value)
                if i == 0:  # Apply word wrap to the first column
                    x_before = self.get_x()
                    y_before = self.get_y()
                    self.multi_cell(col_widths[i], 10, cell_text, border=1, align="L")
                    x_after = self.get_x()
                    y_after = self.get_y()
                    max_height = max(max_height, y_after - y_before)
                    self.set_xy(x_before + col_widths[i], y_before)  # Move to the right of the wrapped cell
                else:
                    self.cell(col_widths[i], max_height, cell_text, border=1, align="C")
            
            self.ln(max_height)  # Move to the next line after each row

            # Check if a page break is needed
            if self.get_y() + max_height +20 > self.page_break_trigger:
                self.add_page()
                self.set_xy(10, 20)  # Reset x position after page break
                # self.ln()
                #for i, col_name in enumerate(dataframe.columns):
                #    self.cell(col_widths[i], 10, col_name, border=1, align="C")
        

            # self.ln(max_height)  # Move to the next line after each row

# Create PDF and add tables
pdf = PDF()
pdf.add_page()

# Add numerical and categorical reports to the PDF
pdf.create_table(numerical_report, "Numerical Columns Report")
pdf.add_page()
pdf.create_table(categorical_report, "Categorical Columns Report")

# Save the PDF
pdf_detailed_output_path = "./detailed_data_quality_report.pdf"
pdf.output(pdf_detailed_output_path)

# Indicate the path of the saved PDF file
pdf_detailed_output_path