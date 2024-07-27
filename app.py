from flask import Flask, request, render_template, send_file
import pandas as pd
import random
import os

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

        # Save the output
        output_file = 'assigned_interns.xlsx'
        assigned_df.to_excel(output_file, index=False)

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

if __name__ == '__main__':
    app.run(debug=True)
