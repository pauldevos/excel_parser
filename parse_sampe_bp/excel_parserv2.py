import os
import sys
import glob
import pandas as pd

def find_tables(sheet_df):
    """ 
    Function gets the length of the first row with non-null values and uses that as the length of the table.
    It then iterates through the rows and appends them to the table if they have the same length as the first row.
    """
    tables = []
    table = None
    current_row_length = 0
    
    # Iterate over the rows in the sheet to get the lenght of the first row with non-null values and compare to subsequent rows to ascertain if they are part of the same table
    for index, row in sheet_df.iterrows():
        non_null_count = row.count()
        
        if non_null_count == 0:  # Skip rows with all nulls
            continue
        
        if table is None:
            current_row_length = non_null_count
            table = {'header': row.index[row.notna()].tolist(), 'data': [row.dropna().tolist()]}

        elif non_null_count == current_row_length:
            table['data'].append(row.dropna().tolist())

        else:
            if table is not None:
                tables.append(table)
                table = None
                
            if non_null_count > 0:
                current_row_length = non_null_count
                table = {'header': row.index[row.notna()].tolist(), 'data': [row.dropna().tolist()]}

    if table is not None:
        tables.append(table)

    # REVISE: Come up with a Header Matching Function or Algorithm or Mapper

    # Remove tables with only one row as they are not considered tables
    # tables = [table for table in tables if len(table['data']) > 1]  

    # # Combine tables with the same header 
    # combined_tables = []
    # previous_table = None
    
    for table in tables:
        if previous_table is None:
            combined_tables.append(table)
            previous_table = table
        elif table['header'] == previous_table['header']:
            previous_table['data'].extend(table['data'])
        else:
            combined_tables.append(table)
            previous_table = table
    
    tables = combined_tables

    return tables

def save_tables_to_csv(tables, sheet_name, output_folder):
    os.makedirs(output_folder, exist_ok=True)
    for i, table in enumerate(tables):
        table_df = pd.DataFrame(table['data'], columns=table['header'])
        print(f"Saving table {i+1} to {output_folder}/{sheet_name}_table{i+1}.csv")
        table_df.to_csv(f"{output_folder}/{sheet_name}_table{i+1}.csv", index=False)

def parse_excel_file(file_path, output_directory):
    excel_data = pd.ExcelFile(file_path, engine='openpyxl')
    for sheet_name in excel_data.sheet_names:
        sheet_df = pd.read_excel(file_path, sheet_name=sheet_name)
        # find if the text "labor category" is in the entire sheet, if it is then process the sheet, if it isn't, then skip the sheet
        sheet_as_a_string = sheet_df.to_string().lower()
        # Check if 'labor category' is in the string
        if 'labor category' in sheet_as_a_string.lower():
            print(f"Processing sheet: {sheet_name}")
            tables = find_tables(sheet_df)
            save_tables_to_csv(tables, sheet_name, output_folder=output_directory)
        else:
            print(f"Skipping sheet {sheet_name} as it does not contain 'labor category' in the sheet")

        

# if __name__ == "__main__":
#     if len(sys.argv) != 3:
#         print("Usage: python3 excel_parser.py <input_directory> <output_directory>")
#         sys.exit(1)
    
#     input_directory = sys.argv[1]
#     output_directory = sys.argv[2]

########### Specific Use Case, Delete For Production ###########
input_directory = "raw_files"
output_directory = "parsed_files"
excel_files = glob.glob(f"{input_directory}/*.xls*")
for file_path in excel_files:
    print(f"\nParsing {file_path}\n")

    # # parse out the filename without the extension
    # file_name_as_directory = os.path.basename(file_path).split('.')[0]
    # # Check if the output directory exists, if not, create it
    # output_directory = f"{output_directory}/{file_name_as_directory}/"
    # if not os.path.exists(output_directory):
    #     os.makedirs(output_directory)
    parse_excel_file(file_path, output_directory)
################################################################