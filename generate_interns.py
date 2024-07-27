import pandas as pd
import random

# Generate a dummy list of 2000 interns with relevant columns
names = [f'Intern_{i+1}' for i in range(2000)]
sexes = [random.choice(['M', 'F']) for _ in range(2000)]
qualifications = [random.choice(['MBChB', 'BDS', 'B.PHARM', 'BSN', 'BSM']) for _ in range(2000)]
universities = [f'Uni_{random.choice(["A", "B", "C", "D", "E"])}' for _ in range(2000)]
years_of_completion = [random.choice([2019, 2020, 2021, 2022, 2023]) for _ in range(2000)]
national_ids = [f'ID_{i+1:05d}' for i in range(2000)]
nationalities = [f'Country_{random.choice(["A", "B", "C", "D", "E"])}' for _ in range(2000)]

# Create the DataFrame
interns_data = {
    'Name': names,
    'Sex': sexes,
    'Qualification': qualifications,
    'University': universities,
    'Year of Completion': years_of_completion,
    'National Identification Number': national_ids,
    'Nationality': nationalities
}
interns_df = pd.DataFrame(interns_data)

# Save the DataFrame to an Excel file
large_dummy_interns_file = 'large_dummy_interns.xlsx'
interns_df.to_excel(large_dummy_interns_file, index=False)

print(f'Dummy interns file saved to {large_dummy_interns_file}')
