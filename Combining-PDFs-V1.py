# --------------------------------------------
# TITLE: Combining PDFs Script
# AUTHOR: Mark Ayiah
# Updated: July 19, 2024
# --------------------------------------------

import os
import importlib

def install_and_import(package):
    try:
        # Try to import the package
        importlib.import_module(package)
    except ImportError:
        # If the package is not found, install it
        os.system(f"{os.sys.executable} -m pip install {package}")
    finally:
        # Import the package after installation
        globals()[package] = importlib.import_module(package)

for library in ["pypdf", "pandas", 'IPython']:
    install_and_import(library)
    
# Importing Libraries
from pypdf import PdfWriter
import pandas as pd
import IPython
import re
import os

def input_q(prompt):
    user_input = input(prompt + '\n[Type "quit" anytime to quit]: ')
    if user_input.lower().strip().replace("'", "").replace('"', '') == 'quit':
        IPython.display.clear_output()
        print("Quitting program... exit successful.")
        return 'quit-quit'
    return user_input

# Defining Combination Function
def combine_pdfs():
    # Creating template, downloads onto device, prompts user to make sure their data is in an acceptable format
    data = {
        "Line Item": ["1", "2", "3"],
        "Support 1": ["example_file_A.pdf", "example_file_B.pdf", "example_file_C.pdf" ],
        "...": ["...", "...", "..."],
        "Support 5": ["example_file_X.pdf", "example_file_Y.pdf", "example_file_Z.pdf"],
        "Final Name": ["files_merged_1.pdf", "files_merged_2.pdf", "files_merged_3.pdf"]
    }
    
    example_df = pd.DataFrame(data)
    print("\n\nBefore we begin, please make sure that your data meets the following 3 conditions:\n\n1. It is in the general format show below (more information below the table):\n")
    print(example_df)

    while True:
        try:
            example_df.drop(columns='...').to_excel("Mapping Template.xlsx", sheet_name="Template", index=False)
            break
        except PermissionError:
            print("\n\nERROR: You likely have the Mapping Template excel sheet still open. Please rename it or close it and try again")
            input("\n\nPlease press ENTER once you have renamed/closed the excel sheet to try again.")
        except Exception as e:
                # Handle any other exception
                print(f"An error occurred: {e}")
                break
    
    print("\n\nThere should be separate column for the names of the files you would like to be merged in order and a column that specifies the names of the final, merged files.\n\nThere is now a file in the same folder where you have saved this Python script called 'Mapping Template.xlsx' that you can use. Please make a copy and save it under another name if you wish to use it.")
    print("\n\n2. This Python script is in the same folder as the mapping excel sheet.")
    print("\n\n3. The PDFs that you wish to merge are in their own seperate folder")
    while True:
        format_flag = input_q("\n\nIs your data in the correct format? Type 'YES' to continue").strip().upper().replace("'", "").replace('"', '')
        if format_flag == 'quit-quit':
            return
        while format_flag == 'YES':
                # Declaring Parameters
                directory_name = input_q("\n\nEnter the name of the subfolder that has the PDFs you would like to merge [CASE SENSITIVE]").strip()
                if directory_name == 'quit-quit':
                    return
                while True:
                    try:
                        files = [os.path.join(directory_name, file_name).replace(directory_name + '\\', '') for file_name in os.listdir(directory_name)]
                        break
                    except FileNotFoundError:
                        directory_name = input_q("\n\nERROR: This folder does not exist. Please try again").strip()
                        if directory_name == 'quit-quit':
                            return
                    except Exception as e:
                        print(f"An error occurred: {e}")
                        break
                        
                output_directory = input_q("\n\nEnter the name of the folder where you would like to save the new PDFs - if you enter a name of a folder that does not exist, it will be created.\n(press ENTER if you would like to leave it as the default 'Merged PDFs')").strip() 
                if output_directory == 'quit-quit':
                    return
                if output_directory == '':
                    output_directory = "Merged PDFs"
                    
                # Importing Mapping File, Checking if File Names Are Valid
                mapping_file = input_q("\n\nEnter the name of the excel file that has the merge mappings (template provided in 'Mapping Template.xlsx') [CASE SENSITIVE, please include the .xlsx]").strip()
                if mapping_file == 'quit-quit':
                    return
                if mapping_file.endswith(tuple(['.pdf', '.csv', '.dta', '.txt', '.xlsm', '.xls', '.xlsb'])):
                    mapping_file = input_q("\n\nERROR: Wrong mapping file format - should be .xlsx. Please try again")
                    if mapping_file == 'quit-quit':
                        return
                if not '.xlsx' in mapping_file:
                    mapping_file = mapping_file + '.xlsx'
                while True:
                    try:
                        mapping = pd.read_excel(mapping_file)
                        break
                    except FileNotFoundError:
                        mapping_file = input_q("\n\nERROR: This file does not exist, please try again")
                        if mapping_file == 'quit-quit':
                            return
                        if any(ext in mapping_file for ext in ['.pdf', '.csv', '.dta']):
                            print("\n\nERROR: Wrong mapping file format - should be .xlsx")
                        if not '.xlsx' in mapping_file:
                            mapping_file = mapping_file + '.xlsx'
                    except Exception as e:
                        print(f"An error occurred: {e}")
                        break
                
                # Checks for column names
                file_columns = []
                col_count = 1
                while True:
                    if col_count == 1:
                        col_name = input_q("Enter the name of the first column for PDF files to be merged (press ENTER to stop)").strip()
                        if col_name == 'quit-quit':
                            return
                    elif col_count > 1:
                        col_name = input_q(f"Enter the name of the next column for PDF files to be merged (currently at file #{col_count} -- press ENTER to stop)").strip()
                        if col_name == 'quit-quit':
                            return
                    if col_name == "":
                        break
                    if col_name not in mapping.columns:
                        print(f"\n\nERROR: Column {col_name} not found in the Excel file. Please check and try again.")
                    else:
                        col_count += 1
                        file_columns.append(col_name)
                        
                if not file_columns:
                    print("\n\nERROR: You must provide at least one column for the PDF files to be merged.")
                    break
                
                final_name = input_q("Column for merged PDF name (default 'Final Name')").strip() or 'Final Name'
                if final_name == 'quit-quit':
                    return

                while True:
                    try:
                        mapping = mapping[file_columns + [final_name]]
                        break
                    except KeyError:
                        print("\nERROR: There is an issue with one or more of your column names. Please double-check and try again")
                        continue
                    except Exception as e:
                        print(f"An error occurred: {e}")
                        break
                
                for col in file_columns:
                    mapping[f'{col} in File List'] = mapping[col].isin(files) | mapping[col].isna()
            
                # Checks if files match, quits if they don't
                if not (mapping[file_columns].isin(files) | mapping[file_columns].isna()).all(axis=1).all():
                    print(f"\n\nError: Some file names in the mapping sheet do not match PDF names in the folder. Please check your data and try again.")
                    print(mapping[~mapping[file_columns].isin(files).all(axis=1)][file_columns + [f'{col} in File List' for col in file_columns]])
                    return
                    
                mapping = mapping[mapping[file_columns].apply(lambda row: all(item in files or pd.isna(item) for item in row), axis=1)]
            
                # Merges pdfs
                os.makedirs(output_directory, exist_ok=True)
                for i, row in mapping.iterrows():
                    merger = PdfWriter()
                    pdfs = [os.path.join(directory_name, row[col]) for col in file_columns if pd.notna(row[col])]
                    for pdf in pdfs:
                        if not pd.isna(pdf):
                            merger.append(pdf)
                    merger.write(os.path.join(output_directory, row[final_name]))
                    merger.close()
                print("\n\nSuccess!!")
                break
        break
    else:
        print("\n\nERROR, please try again")

combine_pdfs()