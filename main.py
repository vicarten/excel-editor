import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
import os

def split_and_trim_excel(input_file, output_folder):
    # import all libraries
    import pandas as pd
    import openpyxl as xls


    # Read the Excel file into a Pandas DataFrame
    fps = pd.read_excel(input_file)

    # Replace "deal" (case-insensitive) with an empty string in the first column starting from row 2
    first_column = fps.iloc[:, 0]  # Access the first column by its index (0-based)
    first_column = first_column.str.replace('deal', '', case=False)
    fps.iloc[:, 0] = first_column  # Assign the modified column back to the DataFrame

    # Find the column number that contains "Address" (case-insensitive) in the header
    target_column_name = "Address"
    address_number = None

    for idx, col_name in enumerate(fps.columns):
        if target_column_name.lower() in col_name.lower():
            address_number = idx
            break
    if address_number is not None:
        # Replace ", USA" (case-insensitive) with an empty string in the first column starting from row 2
        first_column = fps.iloc[:, address_number]  # Access the "Address column"
        first_column = first_column.str.replace(', USA', '', case=False)
        fps.iloc[:, address_number] = first_column  # Assign the modified column back to the DataFrame

    if address_number is not None:
        # Trim the "address" column
        fps.iloc[:, address_number] = fps.iloc[:, address_number].str.strip()

        # Extract the value after the last comma or space and store it in the "zip" column
        fps["Zip"] = fps.iloc[:, address_number].str.extract(r'([^, ]+)(?:[, ]*$)')
        # Remove letters
        fps["Zip"] = fps["Zip"].str.replace(r'[a-zA-Z]', '', regex=True)
        # Trim the "Zip" column
        fps["Zip"] = fps["Zip"].str.strip()
        #Check that it is a zip
        fps.loc[fps['Zip'].str.len() < 4, 'Zip'] = ""
        # Trim the "address" column
        fps.iloc[:, address_number] = fps.iloc[:, address_number].str.strip()
        # Remove "zip" from "address" column
        fps.iloc[:, address_number] = fps.apply(lambda row: str(row[address_number])[:-len(str(row["Zip"]))] if row["Zip"] != "" else str(row[address_number]), axis=1)
        # Trim the "address" column
        fps.iloc[:, address_number] = fps.iloc[:, address_number].str.strip()
        # Remove ","
        fps.iloc[:, address_number] = fps.iloc[:, address_number].str.rstrip(',')
        # Trim the "address" column
        fps.iloc[:, address_number] = fps.iloc[:, address_number].str.strip()

        # Reorder the columns to have "address" and "zip" next to each other
        columns = fps.columns.tolist()
        columns.remove("Zip")
        columns.insert(address_number + 1, "Zip")
        fps = fps[columns]

    if address_number is not None:
        # Trim the "address" column
        fps.iloc[:, address_number] = fps.iloc[:, address_number].str.strip()

        # Extract the value after the last comma
        fps["State"] = fps.iloc[:, address_number].str.rsplit(',', n=1).str.get(-1)
        # Trim the "address" column
        fps.iloc[:, address_number] = fps.iloc[:, address_number].str.strip()
        # Check that it is a state
        fps.loc[fps['State'].str.len() > 3, 'State'] = ""
        # Check if contains numbers
        fps["State"] = fps["State"].str.replace(r'[0-9]', '', regex=True)
        # Trim the "address" column
        fps.iloc[:, address_number] = fps.iloc[:, address_number].str.strip()
        # Remove "State" from "address" column
        fps.iloc[:, address_number] = fps.apply(lambda row: str(row[address_number])[:-len(str(row["State"]))] if row["State"] != "" else str(row[address_number]), axis=1)
        # Trim the "address" column
        fps.iloc[:, address_number] = fps.iloc[:, address_number].str.strip()
        # Remove ","
        fps.iloc[:, address_number] = fps.iloc[:, address_number].str.rstrip(',')
        # Trim the "address" column
        fps.iloc[:, address_number] = fps.iloc[:, address_number].str.strip()

        # Reorder the columns to have "address" and "State" next to each other
        columns = fps.columns.tolist()
        columns.remove("State")
        columns.insert(address_number + 1, "State")
        fps = fps[columns]

    if address_number is not None:
        # Trim the "address" column
        fps.iloc[:, address_number] = fps.iloc[:, address_number].str.strip()

        # Extract the value after the last comma
        fps["City"] = fps.iloc[:, address_number].str.rsplit(',', n=1).str.get(-1)
        # Trim the "address" column
        fps.iloc[:, address_number] = fps.iloc[:, address_number].str.strip()
        # Check that it is a City
        fps.loc[fps['City'].str.len() < 3, 'City'] = ""
        # Check if contains numbers
        fps["City"] = fps["City"].str.replace(r'[0-9]', '', regex=True)
        # Trim the "City" column
        fps["City"] = fps["City"].str.strip()
        # Trim the "address" column
        fps.iloc[:, address_number] = fps.iloc[:, address_number].str.strip()
        # Remove "City" from "address" column
        fps.iloc[:, address_number] = fps.apply(lambda row: str(row[address_number])[:-len(str(row["City"]))] if row["City"] != "" else str(row[address_number]), axis=1)
        # Trim the "address" column
        fps.iloc[:, address_number] = fps.iloc[:, address_number].str.strip()
        # Remove ","
        fps.iloc[:, address_number] = fps.iloc[:, address_number].str.rstrip(',')
        # Trim the "address" column
        fps.iloc[:, address_number] = fps.iloc[:, address_number].str.strip()

        # Reorder the columns to have "address" and "City" next to each other
        columns = fps.columns.tolist()
        columns.remove("City")
        columns.insert(address_number + 1, "City")
        fps = fps[columns]



    # Save the modified DataFrame to a new Excel file
    output_file_path = r"C:\Users\Victoria Ten\Downloads\Modified_CA_Rework.xlsx"
    fps.to_excel(output_file_path, index=False)  # Set index to False if you don't want to save the index column

    print(f"DataFrame saved to {output_file_path}")




def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
    if file_path:
        entry_file_path.delete(0, tk.END)
        entry_file_path.insert(0, file_path)

def browse_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        entry_output_folder.delete(0, tk.END)
        entry_output_folder.insert(0, folder_path)

def start_edit():
    input_file = entry_file_path.get()
    output_folder = entry_output_folder.get()

    if input_file and output_folder:
        split_and_trim_excel(input_file, output_folder)

# Create the main window
root = tk.Tk()
root.title("Excel File Editor")

# Create and place widgets
label_file_path = tk.Label(root, text="Select Excel File:")
label_output_folder = tk.Label(root, text="Select Output Folder:")
entry_file_path = tk.Entry(root, width=40)
entry_output_folder = tk.Entry(root, width=40)
button_browse_file = tk.Button(root, text="Browse", command=browse_file)
button_browse_folder = tk.Button(root, text="Browse", command=browse_folder)
button_start = tk.Button(root, text="Start Editing", command=start_edit)

label_file_path.grid(row=0, column=0, padx=10, pady=5, sticky=tk.E)
entry_file_path.grid(row=0, column=1, padx=10, pady=5)
button_browse_file.grid(row=0, column=2, padx=10, pady=5)
label_output_folder.grid(row=1, column=0, padx=10, pady=5, sticky=tk.E)
entry_output_folder.grid(row=1, column=1, padx=10, pady=5)
button_browse_folder.grid(row=1, column=2, padx=10, pady=5)
button_start.grid(row=2, column=1, padx=10, pady=10)

root.mainloop()
