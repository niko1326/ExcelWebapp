import os
import pandas as pd
from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Folder where processed Excel files will be stored
PROCESSED_FOLDER = 'processed_files'

# Ensure the processed folder exists
if not os.path.exists(PROCESSED_FOLDER):
    os.makedirs(PROCESSED_FOLDER)


# Function to process the Excel file using a selected script
def process_excel(file_path, script_choice, new_file_name):
    data = pd.read_excel(file_path)

    if script_choice == 'script1':
        # Example script: Calculate Row Sums
        data['RowSum'] = data.sum(axis=1)

    elif script_choice == 'script2':
        # Generate a new file with sheets per user summarizing total hours and percentage
        processed_file_path = os.path.join(PROCESSED_FOLDER, f"{new_file_name}.xlsx")
        with pd.ExcelWriter(processed_file_path, engine='openpyxl') as writer:
            users = data['User'].unique()  # Get unique users
            for user in users:
                user_data = data[data['User'] == user]  # Filter data by user

                # Replace empty Project fields with "BRAK WYBRANEGO PROJEKTU"
                user_data['Project'] = user_data['Project'].fillna('BRAK WYBRANEGO PROJEKTU')

                # Add Month-Year column
                user_data['Month-Year'] = pd.to_datetime(user_data['Date']).dt.strftime('%b-%y')

                # Group by Month-Year and Project, then sum the Hours
                summary = (
                    user_data.groupby(['Month-Year', 'Project'])['Hours']
                    .sum()
                    .reset_index()
                )
                summary['Total Hours'] = summary['Hours']

                # Calculate percentages per Month-Year
                monthly_totals = summary.groupby('Month-Year')['Total Hours'].transform('sum')
                summary['Percentage'] = summary['Total Hours'] / monthly_totals

                summary = summary[['Month-Year', 'Project', 'Total Hours', 'Percentage']]

                # Extract the value from the "Spółka (user field)" column
                spolki_values = user_data['Spółka (user field)'].unique()
                spolki_values = [val for val in spolki_values if isinstance(val, str)]  # Filter out NaN
                sheet_suffix = "_".join(sorted(spolki_values))  # Join all unique values

                # Name the sheet with User and Spółka values
                sheet_name = f"{user}_{sheet_suffix}" if sheet_suffix else user
                sheet_name = sheet_name[:31]  # Limit sheet name to 31 czharacters

                # Write the summary to the Excel sheet
                summary.to_excel(writer, index=False, sheet_name=sheet_name)

                # Format the Percentage column as a percentage
                workbook = writer.book
                worksheet = writer.sheets[sheet_name]
                percentage_col_index = summary.columns.get_loc('Percentage') + 1  # Adjust for Excel's 1-based indexing

                for row_num in range(2, len(summary) + 2):  # Start from row 2 (Excel rows are 1-based)
                    cell = worksheet.cell(row=row_num, column=percentage_col_index)
                    cell.number_format = '0.00%'  # Format as percentage

        return processed_file_path

    else:
        raise ValueError("Unsupported script choice.")

    # Save the modified DataFrame back to Excel
    processed_file_path = os.path.join(PROCESSED_FOLDER, f"{new_file_name}.xlsx")
    with pd.ExcelWriter(processed_file_path, engine='openpyxl') as writer:
        data.to_excel(writer, index=False, sheet_name='ProcessedData')

    return processed_file_path


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        script_choice = request.form['script_choice']
        new_file_name = request.form['new_file_name'] or 'processed'  # Default name if none provided

        if file and script_choice:
            # Save the uploaded file
            file_path = os.path.join(PROCESSED_FOLDER, secure_filename(file.filename))
            file.save(file_path)

            # Process the file with the selected script
            processed_file_path = process_excel(file_path, script_choice, new_file_name)

            # Send the processed file back to the user
            return send_file(processed_file_path, as_attachment=True)

    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True)
