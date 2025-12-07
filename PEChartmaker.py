# Make sure you install: "pip install pandas xlsxwriter openpyxl"
# also IMPORTANT - make sure you have the file that your data is pulling from and the graph that your data will add to are CLOSED before you run the program

import pandas as pd
import sys
import os
import math

# ==========================================
# CONFIGURATION
# ==========================================
input_filename = 'CRM_PE_CHART.xlsx' #this is the name of the file you data is pulled from      
sheet_name = 'Worksheet' #this is the name of the sheet your data is pulled from                  
output_filename = 'CEM_ONE_YEAR_NTM_PE.xlsx' #set this whatever name, this will create a new excel with your graph in your files 
# ==========================================

try:
    # 1. READ DATA
    print(f"Reading data from '{input_filename}'...")
    if input_filename.endswith('.csv'):
        df = pd.read_csv(input_filename)
    else:
        df = pd.read_excel(input_filename, sheet_name=sheet_name, engine='openpyxl')

    # Clean data
    df = df.iloc[:, :2]
    df.columns = ['Date', 'PE']
    df['Date'] = pd.to_datetime(df['Date'])

    # 2. CALCULATE MEDIAN
    median_val = df['PE'].median()
    df['Median'] = median_val

    # 3. SETUP EXCEL WRITER
    writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    num_rows = len(df)

    # 4. CALCULATE Y-AXIS BOUNDS
    pe_max = df['PE'].max()
    pe_min = df['PE'].min()
    data_range = pe_max - pe_min
    
    # Calculate 10% Buffer
    if data_range == 0:
        buffer = pe_max * 0.10
    else:
        buffer = data_range * 0.10

    raw_min = pe_min - buffer
    raw_max = pe_max + buffer

    # =========================================================
    # CONDITIONAL LOGIC: Only round if Top Number > 3.5 - This makes it so the graph top/bottom spacing does not look weird if the multiple is under 3.5x
    # =========================================================
    if raw_max > 3.5:
        # --- Apply Integer Rounding & Safety Check ---
        y_axis_min = round(raw_min)
        y_axis_max = round(raw_max)

        # Safety Check: Expand if rounding causes clipping
        if y_axis_min > pe_min:
            y_axis_min -= 1
        if y_axis_max < pe_max:
            y_axis_max += 1
            
        print("Logic: Top number > 3.5. Applied Integer Rounding.")
    else:
        # --- Small Numbers: Keep Decimals (No Rounding) ---
        y_axis_min = raw_min
        y_axis_max = raw_max
        print("Logic: Top number <= 3.5. Kept decimal precision.")

    # Calculate interval to get exactly 5 spaces (6 tick marks)
    interval_unit = (y_axis_max - y_axis_min) / 5

    # 5. CREATE CHART
    chart = workbook.add_chart({'type': 'line'})

    # Series 1: P/E (Blue #00355F)(0-53-95)
    chart.add_series({
        'name':       ['Sheet1', 0, 1],
        'categories': ['Sheet1', 1, 0, num_rows, 0],
        'values':     ['Sheet1', 1, 1, num_rows, 1],
        'line':       {'color': '#00355F', 'width': 2.25},
        'smooth':     True,
    })

    # Series 2: Median (Red #760000, Dashed) (118-0-0)
    chart.add_series({
        'name':       ['Sheet1', 0, 2],
        'categories': ['Sheet1', 1, 0, num_rows, 0],
        'values':     ['Sheet1', 1, 2, num_rows, 2],
        'line':       {'color': '#760000', 'dash_type': 'dash', 'width': 2.25},
        'smooth':     True,
    })

    # 6. FORMATTING

    # TITLE
    file_title = os.path.splitext(input_filename)[0]
    chart.set_title({
        'name': file_title,
        'name_font': {'name': 'Abadi', 'size': 10, 'bold': False, 'color': 'black'}
    })

    chart.set_legend({'none': True})

    # X-AXIS
    if not df.empty:
        days_span = (df['Date'].max() - df['Date'].min()).days
        x_interval = int(days_span / 5) if days_span > 0 else 1
    else:
        x_interval = 1

    chart.set_x_axis({
        'name_font': {'none': True},
        'num_font':  {'name': 'Abadi Extra Light', 'size': 8, 'color': 'black'},
        'date_axis': True,
        'num_format': 'mmm-yy',
        'major_unit': x_interval,
        'major_unit_type': 'days',
        'line': {'color': 'black'},
        'major_tick_mark': 'outside',
    })

    # Y-AXIS
    chart.set_y_axis({
        'name_font': {'none': True},
        'num_font':  {'name': 'Abadi Extra Light', 'size': 8, 'color': 'black'},
        'num_format': '0.0"x"',
        'major_gridlines': {'visible': False},
        'line': {'color': 'black'},
        'major_tick_mark': 'outside',
        'min': y_axis_min,
        'max': y_axis_max,
        'major_unit': interval_unit
    })

    # Insert Chart
    worksheet.insert_chart('E2', chart)
    writer.close()
    
    print(f"Success! Chart created in '{output_filename}'.")
    print(f"Y-Axis Range: {y_axis_min} to {y_axis_max}")

except PermissionError:
    print(f" ERROR: Please close the file '{output_filename}' and run the script again.")
except FileNotFoundError:
    print(f" ERROR: The file '{input_filename}' was not found.")
except Exception as e:
    print(f" An unexpected error occurred: {e}")
