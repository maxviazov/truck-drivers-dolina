# Truck Drivers Analysis Project

This project analyzes Excel files from the Ituran system to generate a report on vehicle movements.

## Features
- GUI for selecting input and output directories.
- Automatic loading of driver mappings from `vehicle/truck-drivers.xlsx`.
- Includes all vehicles in the report.
- Generates readable Excel report with sorting, auto-filters, and route paths.
- Separates reports by days (one row per vehicle per day).
- Saves paths between sessions.
- Modern GUI with icon.

## Files
- `analyze_excel.py`: Main script with GUI.
- `vehicle/truck-drivers.xlsx`: Driver mapping file.
- `car.ico`: Application icon (download a car icon and rename to car.ico).
- `tmp/report.xlsx`: Generated report file.

## How to Run
1. Ensure Python environment is configured.
2. Install dependencies: pandas, openpyxl, xlrd.
3. Download a car icon (.ico) and save as `car.ico` in the project root.
4. Run: `python analyze_excel.py`
5. Use the GUI to select directories.
6. Click "Run Analysis".

## Building Standalone Windows App
To create a standalone .exe:
1. Install pyinstaller: `pip install pyinstaller`
2. Run: `pyinstaller --onefile --windowed --icon=car.ico analyze_excel.py`
3. The .exe will be in `dist/` folder.

## Report Columns
- מס' רכב: Vehicle number
- שם הנהג: Driver name
- תאריך: Date
- מרחק בק"מ: Distance in km
- מקומות: Route (addresses in order)
- סה"כ ק"מ: Total km
