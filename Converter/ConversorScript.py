import tkinter as tk
from tkinter import filedialog
import pandas as pd
import sys
import os

if getattr(sys, 'frozen',False):
    import pyi_splash

# Initialize the main window
root = tk.Tk()
root.title("Dext to Capium Converter")
root.geometry("400x233")
# Define a variable for input path
inputpath = tk.StringVar()

# Function to select input file
def select_input_file():
    input_file = filedialog.askopenfilename(title="Select Input File", filetypes=[("Excel files", "*.xlsx")])
    inputpath.set(input_file)
    input_label.config(text=f"Input File: {input_file}")

# Function to perform conversion
def converter():
    try:
        # Read the input file
        dfDext = pd.read_excel(inputpath.get(), engine='openpyxl')
        
        # Define the columns for the output DataFrame
        columns = ['ContactName', 'TransactionTypeName', 'VDate', 'Description', 'VNo', 'DueDate', 'TaxAmount', 'VatIncluded', 'CurrencyRate', 'AccountName', 'AccountCode', 'Price', 'TaxName', 'VatRate']
        dfCapium = pd.DataFrame(columns=columns)
        
        # Populate the output DataFrame
        dfCapium.ContactName = dfDext['Supplier']
        dfCapium.TransactionTypeName = dfDext['Type']
        dfCapium.VDate = pd.to_datetime(dfDext['Date']).dt.date
        dfCapium.Description = dfDext['Image']
        dfCapium.TaxAmount = dfDext['VAT (GBP)']
        dfCapium.Price = dfDext['Total (GBP)']
        dfCapium.AccountName = dfDext['Category']
        dfCapium.VatIncluded = 'VAT INC'
        dfCapium.TaxName = 'Custom VAT'
        
        # Get the directory of the input file
        input_dir = os.path.dirname(inputpath.get())
        
        # Define the initial output file path
        output_file = os.path.join(input_dir, 'converted_output.xlsx')
        
        # Check if the file exists and create a new filename if it does
        base, ext = os.path.splitext(output_file)
        counter = 1
        while os.path.exists(output_file):
            output_file = f"{base}({counter}){ext}"
            counter += 1
        
        # Save the output DataFrame to the Excel file
        dfCapium.to_excel(output_file, index=False)
        
        result_label.config(text="Conversion successful!", fg="green")
    except Exception as e:
        result_label.config(text=f"Error: {e}", fg="red")

# Add a Label widget
label = tk.Label(root, text="Dext to Capium Converter")
label.pack(pady=5)

# Add a Label widget
label_instruction = tk.Label(root, text="OutputFile will be generated as 'converted_output.xlsx'")
label_instruction.pack(pady=5)

# Add Labels to display selected paths
spacer=tk.Label(root, text="")
spacer.pack(pady=5)

# Add Labels to display selected paths
input_label = tk.Label(root, text="Input File: Not selected")
input_label.pack(pady=5)

# Add Button widgets for file selection
input_button = tk.Button(root, text="Select Input File ('Dext')", command=select_input_file)
input_button.pack(pady=5)

# Add Button widget for conversion
convert_button = tk.Button(root, text="Convert!", command=converter)
convert_button.pack(pady=20)

# Add a Label to display the result of the conversion
result_label = tk.Label(root, text="")
result_label.pack(pady=20)


if getattr(sys,'frozen',False):
    pyi_splash.close()
# Run the main event loop
root.mainloop()