# import pandas as pd
# from tkinter import filedialog, messagebox
# import os
#
#
# def perform_vlookup(folder_path):
#     # Prompt user to select the DATA file (source)
#     source_file = filedialog.askopenfilename(initialdir=folder_path, title="Select the source file (DATA file)",
#                                              filetypes=[("Excel files", "*.xls")])
#     if not source_file:
#         return
#
#     # Locate the target file (output from compare_neotech_drag)
#     # Assume it's the .xlsx file with "_output" in its name
#     target_files = [f for f in os.listdir(folder_path) if f.endswith('_output.xlsx')]
#     if not target_files:
#         messagebox.showerror("Error",
#                              "No output file from the compare_neotech_drag function found in the specified folder.")
#         return
#
#     target_file = os.path.join(folder_path, target_files[0])
#
#     # Read the files using appropriate engines
#     df_source = pd.read_excel(source_file, engine='xlrd')
#
#     # Load all sheets from the target file
#     all_sheets = pd.read_excel(target_file, engine='openpyxl', sheet_name=None)
#
#     # Assuming "Dupes Removed" is the sheet you want to perform VLOOKUP on
#     df_target = all_sheets['Dupes Removed']
#
#     # Perform the vlookup using merge
#     result = pd.merge(df_target, df_source[['PARTNUM', 'DESCRIPTION', 'BASEUNITPRICE']], on='PARTNUM', how='left',
#                       suffixes=('', '_Source'))
#
#     # Save the result back into "Dupes Removed" sheet and retain other sheets
#     with pd.ExcelWriter(target_file, engine='openpyxl') as writer:
#         for sheet_name, df in all_sheets.items():
#             if sheet_name == "Dupes Removed":
#                 result.to_excel(writer, sheet_name=sheet_name, index=False)
#             else:
#                 df.to_excel(writer, sheet_name=sheet_name, index=False)
#
#     messagebox.showinfo("Success", "VLOOKUP operation completed successfully!")
