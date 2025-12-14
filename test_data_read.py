import pandas as pd

# Replace this with the full path to your Excel/CSV file
data_file = "data.xlsx"


try:
    if data_file.endswith('.csv'):
        data = pd.read_csv(data_file)
    else:
        data = pd.read_excel(data_file)
    print("File read successfully!")
    print(data.head())  # Prints first 5 rows
except Exception as e:
    print("Error:", e)
