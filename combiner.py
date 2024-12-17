import pandas as pd
from openpyxl import Workbook

def process_program_registration(file_path, output_file):
    # Load the input Excel file
    data = pd.read_excel(file_path)
    
    # Extract relevant columns: 'STUDENT: Person ID', 'Student', 'Amount Billed'
    selected_columns = data[['Student', 'STUDENT: Person ID', 'Amount Billed']].copy()
    
    # Rename columns for clarity
    selected_columns = selected_columns.rename(columns={
        'Student': 'Student Name',
        'STUDENT: Person ID': 'PersonID',
        'Amount Billed': 'Gross Amount'
    })
    
    # Clean data: drop rows with missing PersonID or Student Name, ensure Gross Amount is numeric
    selected_columns = selected_columns.dropna(subset=['PersonID', 'Student Name'])
    selected_columns['Gross Amount'] = pd.to_numeric(selected_columns['Gross Amount'], errors='coerce').fillna(0)
    
    # Group by PersonID and Student Name, summing up all Gross Amounts
    grouped = selected_columns.groupby(['PersonID', 'Student Name'])['Gross Amount'].apply(list).reset_index()
    
    # Function to create SUM formula with individual row values
    def create_sum_formula(values):
        return f"=SUM({','.join([str(v) for v in values])})"
    
    # Add a new column with the SUM formula
    grouped['Sum Formula'] = grouped['Gross Amount'].apply(create_sum_formula)
    
    # Drop the Gross Amount column (intermediate data)
    grouped = grouped.drop(columns=['Gross Amount'])
    
    # Export to Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    
    # Add column headers
    headers = ['Student Name', 'PersonID', 'Sum of Gross Amounts']
    ws.append(headers)
    
    # Add data rows
    for _, row in grouped.iterrows():
        ws.append([row['Student Name'], row['PersonID'], row['Sum Formula']])
    
    # Save the workbook with the requested name
    wb.save(output_file)
    print(f"Summary Excel file created: {output_file}")

if __name__ == "__main__":
    # Input and output file paths
    input_file = "Current Year Program Registrations with Amounts.xlsx"
    output_file = "current_year_program_registration_totals.xlsx"
    
    # Process the file
    process_program_registration(input_file, output_file)
