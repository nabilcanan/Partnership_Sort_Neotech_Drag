from tkinter import filedialog, messagebox
import numpy as np
import pandas as pd


def compare_neotech_drag(current_week_file, last_week_file):
    print("compare_neotech_drag function called!") # Debugging line
    # Load current week's data
    current_week_data = pd.read_excel(current_week_file, engine='xlrd')
    current_week_data.columns = current_week_data.columns.str.upper().str.strip()

    if 'PARTNUM' not in current_week_data.columns:
        messagebox.showerror("Error", "'PARTNUM' column not found in the new week's file.")
        return

    # Remove duplicates from the current week's data
    current_week_dupes_removed = current_week_data.drop_duplicates(subset='PARTNUM', keep='first')

    # Try to load the "Dupes Removed" sheet from the previous week's file.
    # If it doesn't exist, load the entire file.
    try:
        prev_week_dupes_removed = pd.read_excel(last_week_file, sheet_name='Dupes Removed', engine='xlrd')
        prev_week_dupes_removed.columns = prev_week_dupes_removed.columns.str.upper().str.strip()
    except Exception:
        prev_week_data = pd.read_excel(last_week_file, engine='xlrd')
        prev_week_data.columns = prev_week_data.columns.str.upper().str.strip()
        prev_week_dupes_removed = prev_week_data.drop_duplicates(subset='PARTNUM', keep='first')

    # Identify 'PartNum' values from the previous week that are not in the current week
    lost_items = prev_week_dupes_removed[
        ~prev_week_dupes_removed['PARTNUM'].isin(current_week_dupes_removed['PARTNUM'])]

    # Merge with previous week data to get the 'MINORDERQTY' and 'BaseUnitPrice' columns
    current_week_dupes_removed = pd.merge(
        current_week_dupes_removed,
        prev_week_dupes_removed[['PARTNUM', 'MINORDERQTY', 'BASEUNITPRICE']],
        on='PARTNUM', how='left', suffixes=('', '_Last_Week')
    )

    # Populate 'MOQ Changed From' column
    condition_moq_change = current_week_dupes_removed['MINORDERQTY'] != current_week_dupes_removed[
        'MINORDERQTY_Last_Week']
    current_week_dupes_removed['MOQ Changed From'] = np.where(condition_moq_change,
                                                              current_week_dupes_removed['MINORDERQTY_Last_Week'],
                                                              np.nan)

    # Create 'Contract Change' column based on 'BaseUnitPrice' comparison
    conditions = [
        (current_week_dupes_removed['BASEUNITPRICE'] > current_week_dupes_removed['BASEUNITPRICE_Last_Week']),
        (current_week_dupes_removed['BASEUNITPRICE'] < current_week_dupes_removed['BASEUNITPRICE_Last_Week']),
        (current_week_dupes_removed['BASEUNITPRICE_Last_Week'].isnull())
    ]
    choices = ['Price Increased', 'Price Decreased', 'New Item']
    current_week_dupes_removed['Contract Change'] = np.select(conditions, choices, default='No Change')

    # Save the data to Excel
    with pd.ExcelWriter(current_week_file, engine='xlsxwriter') as writer:
        current_week_data.to_excel(writer, sheet_name="Original Data", index=False)
        current_week_dupes_removed.to_excel(writer, sheet_name="Dupes Removed", index=False)
        lost_items.to_excel(writer, sheet_name='Lost Items', index=False)

        workbook = writer.book
        wrap_format = workbook.add_format({'text_wrap': True})

        for sheet_name in ["Original Data", "Dupes Removed", "Lost Items"]:
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes(1, 10)

            for col_num, value in enumerate(current_week_data.columns.values):
                worksheet.write(0, col_num, value, wrap_format)

            worksheet.autofilter(0, 0, len(current_week_data), len(current_week_data.columns) - 1)

    messagebox.showinfo("Success", "Operation completed successfully!")
