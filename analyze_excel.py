import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd


def analyze_files(input_dir, output_dir):
    # Load driver mapping from file
    driver_file = 'vehicle/truck-drivers.xlsx'
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
        'מרחק בק"מ': 'sum',
        'כתובת': lambda x: list(x.dropna().astype(str)),
        'שם נהג': lambda x: x.dropna().iloc[0] if not x.dropna().empty else 'Unknown'
    }).reset_index()

    # Now group by vehicle
    def aggregate_vehicle(group):
        km_dict = {}
        for _, row in group.iterrows():
            date_str = row['date'].strftime('%Y-%m-%d')
            km = row['מרחק בק"מ']
            if km > 0:
                km_dict[date_str] = km
        total_km = sum(km_dict.values())
        days_str = ', '.join(sorted(km_dict.keys()))

        # Collect addresses in chronological order
        addresses_in_order = []
        for _, row in group.sort_values('date').iterrows():
            addresses_in_order.extend(row['כתובת'])

        # Extract cities (last part after comma)
        cities = []
        for addr in addresses_in_order:
            if ',' in addr:
                city = addr.split(',')[-1].strip()
                if city and city not in cities:
                    cities.append(city)

        if cities:
            # Split into lines of 7 cities each
            route_parts = [' - '.join(cities[i:i+7]) for i in range(0, len(cities), 7)]
            route_str = 'Старт\n' + '\n'.join(route_parts) + '\nФиниш'
        else:
            route_str = 'Нет данных'

        driver = group['שם נהג'].iloc[0] if not group['שם נהג'].empty else 'Unknown'

        return pd.Series({
            'Дни': days_str,
            'Суммарные км': total_km,
            'שם הנהג': driver,
            'מקומות': route_str
        })

    final_grouped = daily_grouped.groupby('תג זיהוי').apply(aggregate_vehicle).reset_index()

    # Rename properly
    final_grouped = final_grouped.rename(columns={
        'תג זיהוי': "מס' רכב"
    })

    # Filter vehicles with movement
    final_grouped = final_grouped[final_grouped['Суммарные км'] > 0]

    # Apply driver mapping
    for index, row in final_grouped.iterrows():
        vehicle = str(row["מס' רכב"])
        if vehicle in driver_mapping:
            final_grouped.at[index, 'שם הנהג'] = driver_mapping[vehicle]

    # Sort
    final_grouped = final_grouped.sort_values(by=["מס' רכב"])

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
