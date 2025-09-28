import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd


def analyze_files(input_dir, output_dir):
    if getattr(sys, 'frozen', False):
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    driver_file = os.path.join(script_dir, 'truck-drivers.xlsx')
    driver_mapping = {}
    if os.path.exists(driver_file):
        driver_df = pd.read_excel(driver_file)
        for _, row in driver_df.iterrows():
            driver_mapping[str(row['vehicle_number'])] = row['driver_name']

    # List of files (assuming they are in input_dir)
    itur_files = [f for f in os.listdir(input_dir) if f.startswith('ייצוא-Excel-דוח') and f.endswith('.xlsx')]

    if not itur_files:
        messagebox.showerror("Error", "No Ituran files found in the input directory.")
        return

    # Read Ituran files
    all_data = []
    for file in itur_files:
        file_path = os.path.join(input_dir, file)
        xl = pd.ExcelFile(file_path)
        for sheet in xl.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet, header=7)
            df['date'] = pd.to_datetime(df['זמן הודעה'], format='%d/%m/%Y %H:%M:%S').dt.date
            # Clean מרחק בק"מ: extract numbers
            df['מרחק בק"מ'] = pd.to_numeric(df['מרחק בק"מ'], errors='coerce').fillna(0)
            all_data.append(df)

    combined_df = pd.concat(all_data, ignore_index=True)

    # Determine date range for filename
    min_date = combined_df['date'].min()
    max_date = combined_df['date'].max()
    if min_date == max_date:
        date_str = min_date.strftime('%Y-%m-%d')
    else:
        date_str = f"{min_date.strftime('%Y-%m-%d')}_to_{max_date.strftime('%Y-%m-%d')}"

    # Group by vehicle and date first
    daily_grouped = combined_df.groupby(['תג זיהוי', 'date']).agg({
        'מרחק בק"מ': lambda x: x.max() - x.min() if len(x) > 1 else (x.iloc[0] if len(x) > 0 else 0),
        'כתובת': lambda x: list(x.dropna().astype(str)),
        'שם נהג': lambda x: x.dropna().iloc[0] if not x.dropna().empty else 'Unknown'
    }).reset_index()

    # Aggregate route for each day
    def aggregate_route(addresses):
        unique_addresses = sorted(set(addresses))
        if len(unique_addresses) > 3:
            return ', '.join(unique_addresses[:3]) + ' и др.'
        return ', '.join(unique_addresses)

    daily_grouped['מקומות'] = daily_grouped['כתובת'].apply(aggregate_route)

    final_grouped = daily_grouped

    # Rename properly
    final_grouped = final_grouped.rename(columns={
        'תג זיהוי': "מס' רכב",
        'מרחק בק"מ': 'Суммарные км',
        'שם נהג': 'שם הנהג',
        'date': 'Дни'
    })

    final_grouped['Дни'] = final_grouped['Дни'].astype(str)

    # Filter vehicles with data
    final_grouped = final_grouped[final_grouped['Суммарные км'] >= 0]

    # Apply driver mapping
    for index, row in final_grouped.iterrows():
        vehicle = str(row["מס' רכב"])
        if vehicle in driver_mapping:
            final_grouped.at[index, 'שם הנהג'] = driver_mapping[vehicle]

    # Sort
    final_grouped = final_grouped.sort_values(by=["מס' רכב", 'Дни'])

    print("Combined Report:")
    print(final_grouped[["מס' רכב", 'שם הנהג', 'Дни', 'Суммарные км', 'מקומות']])

    # Save to Excel
    output_file = os.path.join(output_dir, f'truck_drivers_reports_{date_str}.xlsx')
    print(f"Saving to {output_file}")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        final_grouped[["מס' רכב", 'שם הנהג', 'Дни', 'Суммарные км', 'מקומות']].to_excel(writer, sheet_name='Report', index=False)
    print("Saved successfully")
    messagebox.showinfo("Success", f"Report saved to {output_file}")

def select_input_dir():
    dir_path = filedialog.askdirectory(title="Select Input Directory")
    if dir_path:
        input_dir_var.set(dir_path)
        save_paths()

def select_output_dir():
    dir_path = filedialog.askdirectory(title="Select Output Directory")
    if dir_path:
        output_dir_var.set(dir_path)
        save_paths()

def save_paths():
    with open('paths.txt', 'w') as f:
        f.write(f"{input_dir_var.get()}\n{output_dir_var.get()}")

def load_paths():
    if os.path.exists('paths.txt'):
        with open('paths.txt', 'r') as f:
            lines = f.readlines()
            if len(lines) >= 2:
                input_dir_var.set(lines[0].strip())
                output_dir_var.set(lines[1].strip())

def run_analysis():
    input_dir = input_dir_var.get()
    output_dir = output_dir_var.get()
    analyze_files(input_dir, output_dir)

# GUI
root = tk.Tk()
root.title("Truck Drivers Analysis Tool")
root.geometry("650x220")
root.configure(bg='#e8f4f8')

# Set icon
try:
    root.iconbitmap('car.ico')
except Exception:
    pass  # If icon not found, skip

# Style
style = ttk.Style()
style.configure('TButton', font=('Arial', 10, 'bold'), padding=5)
style.configure('TLabel', font=('Arial', 12, 'bold'), background='#e8f4f8')

# Fonts
label_font = ('Arial', 12, 'bold')
entry_font = ('Arial', 10)
button_font = ('Arial', 10, 'bold')

input_dir_var = tk.StringVar(value='tmp')
output_dir_var = tk.StringVar(value='tmp')

load_paths()

# Title
title_label = ttk.Label(root, text="🚛 Truck Drivers Analysis Tool", font=('Arial', 18, 'bold'), background='#e8f4f8', foreground='#2c3e50')
title_label.grid(row=0, column=0, columnspan=3, pady=15)

ttk.Label(root, text="Input Directory:", style='TLabel').grid(row=1, column=0, sticky='w', padx=5)
tk.Entry(root, textvariable=input_dir_var, width=55, font=entry_font).grid(row=1, column=1, padx=5)
ttk.Button(root, text="Browse", command=select_input_dir, style='TButton').grid(row=1, column=2, padx=5)

ttk.Label(root, text="Output Directory:", style='TLabel').grid(row=2, column=0, sticky='w', padx=5)
tk.Entry(root, textvariable=output_dir_var, width=55, font=entry_font).grid(row=2, column=1, padx=5)
ttk.Button(root, text="Browse", command=select_output_dir, style='TButton').grid(row=2, column=2, padx=5)

ttk.Button(root, text="🚀 Run Analysis", command=run_analysis, style='TButton').grid(row=3, column=0, pady=25, padx=5)
ttk.Button(root, text="Exit", command=root.quit, style='TButton').grid(row=3, column=2, pady=25, padx=5)

root.mainloop()
