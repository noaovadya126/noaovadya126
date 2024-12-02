import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
from datetime import datetime
import os
import subprocess

# Initialize global variables
source_file_path = None
compare_file_path = None
error_df = pd.DataFrame()

# Specify columns to ignore if hidden
hidden_columns = ['HiddenColumn1', 'HiddenColumn2']  # Replace with actual hidden column names if needed

# Create the main application window
root = tk.Tk()
root.title("Data Comparison Tool")

# Function to select source file
def select_source_file():
    global source_file_path
    source_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
    if source_file_path:
        source_button.config(text="Source: Selected")
        log_text.insert(tk.END, "Source file selected.\n")

# Function to select comparison file
def select_compare_file():
    global compare_file_path
    compare_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
    if compare_file_path:
        compare_button.config(text="Compare: Selected")
        log_text.insert(tk.END, "Compare file selected.\n")

# Function to load file into a DataFrame, handling merged cells and ignoring hidden columns
def load_file(file_path):
    if file_path.endswith('.xlsx'):
        df = pd.read_excel(file_path, header=0)
    else:
        df = pd.read_csv(file_path, header=0)

    # Drop hidden columns
    df = df[[col for col in df.columns if col not in hidden_columns]]

    # Forward-fill merged cells (assuming merged cells are represented by NaN values in the Excel output)
    df.ffill(inplace=True)

    return df

# Function to calculate differences
def calculate_difference():
    global error_df
    if not source_file_path or not compare_file_path:
        messagebox.showwarning("Missing Files", "Please select both Source and Compare files.")
        return

    # Load files into data frames
    source_df = load_file(source_file_path)
    compare_df = load_file(compare_file_path)

    # Create an empty error DataFrame
    error_data = []

    # Compare each cell in the data frames
    for row_idx in range(max(len(source_df), len(compare_df))):
        for col in source_df.columns:
            if col in compare_df.columns:
                source_val = source_df.at[row_idx, col] if row_idx < len(source_df) else None
                compare_val = compare_df.at[row_idx, col] if row_idx < len(compare_df) else None

                if pd.notna(source_val) and pd.notna(compare_val):
                    # Check for type mismatch
                    if type(source_val) != type(compare_val):
                        error_data.append({
                            "Row": row_idx + 1,
                            "Column": col,
                            "Error Type": "Type Mismatch",
                            "Source Value": source_val,
                            "Compare Value": compare_val
                        })
                    # Check for value mismatch
                    elif source_val != compare_val:
                        error_data.append({
                            "Row": row_idx + 1,
                            "Column": col,
                            "Error Type": "Value Mismatch",
                            "Source Value": source_val,
                            "Compare Value": compare_val
                        })
                elif source_val != compare_val:  # One of the values is NaN (None)
                    error_data.append({
                        "Row": row_idx + 1,
                        "Column": col,
                        "Error Type": "Value Mismatch",
                        "Source Value": source_val,
                        "Compare Value": compare_val
                    })

    # Convert the error data to a DataFrame
    error_df = pd.DataFrame(error_data)
    log_text.insert(tk.END, "Comparison completed. Errors found: " + str(len(error_data)) + "\n")

# Function to download the error log and open the result file automatically
def download_error_log():
    if error_df.empty:
        messagebox.showinfo("No Errors", "No errors to save.")
        return

    # Get the current date for the filename
    current_date = datetime.now().strftime("%d-%m")
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        initialfile=f"Result_{current_date}.xlsx"
    )

    if file_path:
        error_df.to_excel(file_path, index=False)
        log_text.insert(tk.END, f"Error log saved to {file_path}\n")

        # Automatically open the file after saving
        try:
            if os.name == 'nt':  # Windows
                os.startfile(file_path)
            elif os.name == 'posix':  # macOS/Linux
                subprocess.call(['open' if os.uname().sysname == 'Darwin' else 'xdg-open', file_path])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file: {e}")

# Function to clear all fields and reset values
def clear_all():
    global source_file_path, compare_file_path, error_df
    source_file_path = None
    compare_file_path = None
    error_df = pd.DataFrame()
    source_button.config(text="Source")
    compare_button.config(text="Compare")
    log_text.delete(1.0, tk.END)
    log_text.insert(tk.END, "Cleared all fields.\n")

# Create buttons and text box for the UI
source_button = tk.Button(root, text="Source file", command=select_source_file, width=15)
source_button.grid(row=0, column=0, padx=10, pady=5)

compare_button = tk.Button(root, text="Compare file", command=select_compare_file, width=15)
compare_button.grid(row=1, column=0, padx=10, pady=5)

calculate_button = tk.Button(root, text="Calculate", command=calculate_difference, width=15)
calculate_button.grid(row=2, column=0, padx=10, pady=5)

download_button = tk.Button(root, text="Download result", command=download_error_log, width=15)
download_button.grid(row=3, column=0, padx=10, pady=5)

# Add Clear button
clear_button = tk.Button(root, text="Clear", command=clear_all, width=15)
clear_button.grid(row=4, column=0, padx=10, pady=5)

# Log text area for displaying messages
log_text = scrolledtext.ScrolledText(root, width=40, height=10)
log_text.grid(row=0, column=1, rowspan=5, padx=10, pady=5)

# Run the main loop
root.mainloop()
