from flask import Flask, request, render_template, send_file
import pandas as pd
import random
import os
from openpyxl import load_workbook

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        if 'interns_file' not in request.files or 'facilities_file' not in request.files:
            return 'No file part'
        
        interns_file = request.files['interns_file']
        facilities_file = request.files['facilities_file']

        interns_df = pd.read_excel(interns_file)
        facilities_df = pd.read_excel(facilities_file, sheet_name='CARRYING CAPCITY OF INTERNSHIP ')

        # Print columns to debug
        print("Interns DataFrame columns:", interns_df.columns)
        print("Facilities DataFrame columns:", facilities_df.columns)

        # Check if the expected columns are in the dataframe
        if 'Qualification' not in interns_df.columns:
            return "The 'Qualification' column is missing from the interns file."

        # Transform the facilities dataframe to match expected format
        facilities = transform_facilities_df(facilities_df)

        # Distribution logic
        assigned_df = distribute_interns(interns_df, facilities)

        # Sort the assigned dataframe by 'Internship Center' column
        sorted_df = assigned_df.sort_values(by=['Internship Center'])

        # Save the output
        output_file = 'assigned_interns_sorted.xlsx'
        sorted_df.to_excel(output_file, index=False)

        # Adjust column widths
        adjust_column_widths(output_file)

        return send_file(output_file, as_attachment=True)
    
    except Exception as e:
        # Log the exception
        print(f"Error: {e}")
        return str(e)

def transform_facilities_df(facilities_df):
    # Unpivot the facilities dataframe
    facilities = facilities_df.melt(id_vars=['Internship Centre'], 
                                    value_vars=['MBChB', 'BDS', 'B.PHARM', 'BSN', 'BSM'],
                                    var_name='Qualification', 
                                    value_name='Available Positions')
    facilities = facilities[facilities['Available Positions'] > 0]
    return facilities

def distribute_interns(interns_df, facilities_df):
    facilities = facilities_df.to_dict(orient='records')
    interns_df['Internship Center'] = None

    for index, intern in interns_df.iterrows():
        qualified_facilities = [f for f in facilities if f['Qualification'] == intern['Qualification']]
        if qualified_facilities:
            chosen_facility = random.choice(qualified_facilities)
            interns_df.at[index, 'Internship Center'] = chosen_facility['Internship Centre']
            chosen_facility['Available Positions'] -= 1
            if chosen_facility['Available Positions'] == 0:
                facilities.remove(chosen_facility)

    return interns_df

def adjust_column_widths(file_path):
    # Load the workbook and select the active sheet
    wb = load_workbook(file_path)
    ws = wb.active

    # Set the width of each column to be adequate for its content
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width

    # Save the workbook
    wb.save(file_path)


if __name__ == '__main__':
    app.run(debug=True)
