import os
import re
from datetime import datetime
import logging
import warnings
import numpy as np
import pandas as pd
from openpyxl.styles import PatternFill

'''
Combine the sending file (WGS_XXXX) with 'star' files :
    according to the sample code in sending file, fetch the samples in the 'star' files
    (these samples need to be run first)
'''

warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

def create_dir(d):
    if not os.path.exists(d):
        try:
            os.makedirs(d) 
            logging.info(f"[create_dir] Successfully created directory: {results_dir}\n")
        except Exception as e:
            logging.error(f"[create_dir] Error occurred while creating directory {d}: {e}\n")

def return_extention(f):
    _, file_extension = os.path.splitext(f)
    return file_extension

def read_table_file(file_name, ext):
    if ext == '.csv':
        df = pd.read_csv(file_name)
    elif ext == '.xlsx':
        df = pd.read_excel(file_name)
    else:
        exit("Unsupported file format")
    return df

def fetch_filename(data_folder, file_name):
    file_path = os.path.join(data_folder, file_name)
    file_name = os.path.basename(file_path)
    return file_name

def fetch_columns(df, required_columns):
    """
    Fetches columns from the DataFrame by matching them against the list of required columns.
    Handles cases where there are extra spaces, capitalization differences, etc.
    
    Args:
    df: The pandas DataFrame.
    required_columns: A list of required column names to fetch.
    
    Returns:
    A DataFrame with only the required columns, if they are found.
    """
    # Standardize column names in the DataFrame
    df.columns = df.columns.str.strip().str.lower()

    # Standardize the required columns list (strip spaces and convert to lowercase)
    required_columns_standardized = [col.strip().lower() for col in required_columns]

    # Map the required columns to the actual columns in the DataFrame
    available_columns = []
    for col in required_columns_standardized:
        for df_col in df.columns:
            if col in df_col:  # Check if the required column is a substring of the actual column
                available_columns.append(df_col)
                break
    
    if not available_columns:
        raise ValueError("None of the required columns were found.")

    return available_columns

def replace_hyphen_by_dot(df, col_name):
    pd.Series(df[col_name]).str.replace('-', '.', regex=False)
    return df

def sort_files(d):
    files = os.listdir(d)
    pattern = r'Star_(0\d{2})_(.+)'  
    file_data = []

    for file in files:
        match = re.match(pattern, file)
        if match:
            numeric_part = int(match.group(1)) 
            file_data.append((numeric_part, file))  

    file_data.sort(key=lambda x: x[0])  

    return [file for _, file in file_data]

def sort_df(df, col_name):
    df['numeric_part'] = df[col_name].str.extract('(\d+)').astype(int)
    return df.sort_values(by='numeric_part', ascending=True).drop(columns='numeric_part')

def create_plate_df():
    df = pd.DataFrame(np.empty((8, 12), dtype=object))
    df.columns = [str(i + 1) for i in range(12)]
    df.index = [chr(65 + i) for i in range(8)]

    return df

# Try to fix the issue of different lower/upper case col header
def manage_columns_typo(df):
    return df.columns.str.strip().str.lower()

def create_table_df():
    index = [f"{chr(col)}{row}" for row in range(1, 13) for col in range(ord('A'), ord('H') + 1)]

    df = pd.DataFrame(index=index)
    df[SOURCE_NAME_COL] = '' 
    return df

def color_header(ws, color=None):
    for cell in ws[1]:  
        cell.fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type='solid')

    for cell in ws['A']:
        cell.fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type='solid')


def cells_to_highlight(df, star_run):
    cells = []
    cells_positions = {}
   
    current_df = df[df[STAR_RUN_COL] == star_run]

    for _, row in current_df.iterrows():
        cells.append(row[SAMPLES_CODES_COL])
        cells_positions[row[TARGET_POSITION_COL]] = row[SAMPLES_CODES_COL]

    return cells, cells_positions

def populate_table(d, table, col):
    for key, value in d.items():
        table.loc[key, col] = value

def create_new_tables_sheet(writer, tables, sheet_name):
    rows = []
    for key, df in tables.items():
        rows.append([key,'',''])
        rows.append(['','',SOURCE_NAME_COL])
        for index, row in df.iterrows():
            rows.append(['', index, row.values[0]])
             
         
        rows.append(['','',''])
            
    result_df = pd.DataFrame(rows, columns=[None, None, None]) 
    result_df.to_excel(writer, sheet_name=sheet_name, index=False)
            
def get_position(df, star_run):
    cells_positions = []
   
    current_df = df[df[STAR_RUN_COL] == star_run]

    for _, row in current_df.iterrows():
        cells_positions.append(row[TARGET_POSITION_COL])

    return cells_positions

def highlighting_cells(ws, cells):
    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')         
    for row in ws.iter_rows():
        for cell in row:
            if cell.value in cells: 
                cell.fill = yellow_fill        

def fetch_and_concatenate(df, data_folder):
    df_result = df.copy()

    for file_name in os.listdir(data_folder):
        if file_name.startswith('Star_0'):
            file_path = os.path.join(data_folder, file_name)
            file_name = os.path.basename(file_path)
            run_name = file_name.split('_')[1] 
            # run_name = '_'.join(run_name_parts)

            ext = return_extention(file_path)

            df_star = read_table_file(file_path, ext)

     
            # Ensure that columns are properly formatted (strip whitespace, lowercase)
            df_star.columns = manage_columns_typo(df_star)

            # We need to replace '.' by '-' in the file to have matchs with Star files codes
            # The star files have codes with '-'
            df_star = replace_hyphen_by_dot(df_star, SOURCE_BARCODE_COL)
            
            # Add the Star run
            df_star[STAR_RUN_COL] = run_name

            if SOURCE_BARCODE_COL in df_star.columns and SAMPLES_CODES_COL in df.columns:
                for i, sample_code in enumerate(df[SAMPLES_CODES_COL]):
                    matched_rows = df_star[df_star[SOURCE_BARCODE_COL] == sample_code]
                    
                    # If there's a match, append the row to the corresponding row in df_result
                    if not matched_rows.empty:
                        for column in matched_rows.columns:
                            # Add the matched row to df_result (or merge the columns if already exists)
                            df_result.loc[i, column] = matched_rows.iloc[0][column]
    
        
    return df_result[[SAMPLES_CODES_COL, SOURCE_BARCODE_COL, TARGET_POSITION_COL, STAR_RUN_COL]]

def create_tables(data_folder, df_sorted, output_file):
    tables_by_run = {}  # Create tables to input Myra
    
    with pd.ExcelWriter(output_file,  engine='openpyxl', mode='a') as writer:
        files = sort_files(data_folder)

        for file_name in files:
            cells = []

            if file_name.startswith('Star_0') or file_name.startswith('star_0'):
                file_name = fetch_filename(data_folder, file_name)
                run_name = file_name.split('_')[1] 
                
                # Create df like 96w plate A -> H (rows), 1 -> 12 (cols)
                df_plate = create_plate_df() 
                df_table = create_table_df() 
                
               
                file_name = os.path.join(data_folder, file_name)
                
                ext = return_extention(file_name)

                df = read_table_file(file_name, ext)
                df.columns = manage_columns_typo(df)

                df = replace_hyphen_by_dot(df, SOURCE_BARCODE_COL)
                df = df.dropna()        # remove nan

                for _, row in df.iterrows():
                    target_position = str(row[TARGET_POSITION_COL])
                    sample_code = row[SOURCE_BARCODE_COL]
                    
                    if target_position:
                        plate_row = target_position[0]
                        plate_col = target_position[1:]
                        
                        df_plate.loc[plate_row, plate_col] = sample_code


                cells, positions = cells_to_highlight(df_sorted, run_name)
                
                df_plate.to_excel(writer, sheet_name=run_name)
               
                populate_table(positions, df_table, SOURCE_NAME_COL)

                # Color cells
                ws = writer.sheets[run_name]
                color_header(ws,'ADD8E6')
                highlighting_cells(ws, cells)

                # Create another table for Myra input
                tables_by_run[run_name] = df_table

    # We want a tables sheet containing all tables for Myra input
    # The Myra need position (A1, B1...), and the sample next to each position
    # If we dont have sample for a position, the cell will be empty
    with pd.ExcelWriter(output_file,  engine='openpyxl', mode='a') as writer:
        create_new_tables_sheet(writer, tables_by_run, 'tables')


SERIAL_NUMBER_COL     = 'serial number'
SAMPLES_CODES_COL     = "samples codes"
PRODUCTION_DATE_COL   = "production date"
SOURCE_BARCODE_COL    = "source barcode" 
TARGET_POSITION_COL   = "target position"
SAMPLES_CODES_COL     = 'samples codes'
STAR_RUN_COL          = 'star run'
SOURCE_NAME_COL       = 'source name'

wd = os.path.dirname(os.path.realpath(__file__))
results_dir = os.path.join(wd, 'results')
logs_dir = os.path.join(wd, 'logs')

now = datetime.now()
formatted_date = now.strftime("%Y%m%d_%H%M%S")

create_dir(results_dir)
create_dir(logs_dir)

log_file = os.path.join(logs_dir, f'logs_{formatted_date}.txt')
logging.basicConfig(filename=log_file, level=logging.DEBUG, 
                    format='%(asctime)s - %(levelname)s - %(message)s')


data_dir        = os.path.join(wd, 'data')
results_file    = os.path.join(results_dir, f'results_{formatted_date}.xlsx')
filtered_file   = 'filtered_result.xlsx'
concatened_file = 'concatenated_result.xlsx'


################################################################

# [1 ]- Fetch 'Samples codes' and 'Production date' columns
try:
    for i in os.listdir(data_dir):
        if i.startswith('WGS') or i.startswith('wgs'):
            sending_file = os.path.join(data_dir, i)
            break
    else:
        msg = 'Cannot find your sending file'
        logging.error(f"[1.1] {msg} \n")
        exit(msg) 

    logging.info(f"[1.1] Successfully fetched 'sending' file\n")
except Exception as e:
    logging.error(f"[1.1] Error occurred while fetching 'sending' file: {e}\n")


try:
    df = pd.read_excel(sending_file, header=None)
    # Reaching the table because it doesn't start from the first file row
    header_row = None
    for i, row in df.iterrows():
        row_str = row.astype(str).str.strip().str.lower()   # fix upper/lower case issue
        if SERIAL_NUMBER_COL in row_str.values:
            header_row = i
            break

    if header_row is not None:
        df = pd.read_excel(sending_file, header=header_row)
        
        columns = fetch_columns(df, [SAMPLES_CODES_COL, PRODUCTION_DATE_COL])

        df_filtered = df[columns]
        df_filtered = replace_hyphen_by_dot(df_filtered, SAMPLES_CODES_COL)
        logging.info(f"[1.2] Successfully fetching header form'sending' file\n")

    else:
        logging.error(f"[1.2] Error occurred while fetching header form'sending' file: {e}\n")

        exit("Could not find the table header.")

except Exception as e:
    logging.error(f"[1.2] Error occurred while fetching header form'sending' file: {e}\n")

# [2] - Concate Star_0XX files
directory = 'data'  
try:
    df_results = fetch_and_concatenate(df_filtered, directory)
    logging.info(f"[2.1] Successfully fetching & concate 'star' files.\n")
except Exception as e:
    msg = f"[2.1] Error occurred while fetching & concate 'star' files: {e}\n"
    logging.error(msg)
    exit(msg)

has_nan = df_results.isna().any().any()

if has_nan:
    rows_with_nan = df_results[df_results.isna().any(axis=1)]
    
    msg = f'[2.2] Your data has Null values: {rows_with_nan}\n'
    logging.error(msg)
    exit(msg)


try:
    df_sorted = sort_df(df_results, STAR_RUN_COL)
    logging.info(f"[2.3] Successfully sorted values\n")
except Exception as e:
    msg = f"[2.3] Error occurred while sorting values: {e}\n"
    logging.error(msg)
    exit(msg)

try:
    with pd.ExcelWriter(results_file, engine='openpyxl') as writer:
        df_filtered.to_excel(writer, sheet_name=filtered_file, index=False)
        df_sorted.to_excel(writer, sheet_name=concatened_file, index=False)
        
        logging.info(f"[2.4] Successfully wrote excel table\n")
except Exception as e:
    msg = f"[2.4] Error occurred while writing excel table: {e}\n"
    logging.error(msg)
    exit(msg)

# [3] - Create Tables
# create like 96w table A1, B1, C1 ....
# color the eva's samples
try:
    create_tables(directory, df_sorted, results_file)
    logging.info('[2.5] Successfully created tables sheets (plates, long table)\n')
    msg = 'The script ran successfully\n\n--- THE END ---'
    logging.info(msg)
    print(msg)
except Exception as e:
    logging.error(f"[2.5] Error occurred while creating tables sheets: {e}")

