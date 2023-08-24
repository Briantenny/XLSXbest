import pandas as pd
from fuzzywuzzy import fuzz
import os
import sys
from tqdm import tqdm


def select_file():
    """
    Asks the user to select an Excel file from the files in the current directory.
    Then, asks the user input parameters for these files:
        - How many rows to skip,
        - Which columns contain dates
    
    Returns a dictionary containing:
        - filepath : filepath of the selected file,
        - skiprows : number of rows to skip at the beginning of the file,
        - fuzzycolumn : column index (1-indexed) of the column to use for fuzzy-lookup
    """
    
    print("Here are the .xlsx files in the current directory:")
    xlsx_files = [x for x in os.listdir() if x.endswith('.xlsx')]
    if len(xlsx_files) == 0:
        print("The current directory must contain at least 1 .xlsx file!")
        print("Please copy the files you want to perform a Fuzzy Lookup on within the same directory as this program.")
        print("Exiting...")
        sys.exit(0)
              
    for i, file in enumerate(xlsx_files):
        print("{}: {}".format(i, file))
    file_selected = False
    while not file_selected:
        # ask for user selection of the file to process
        file_idx = input("Please select a file (use the index number / exit with 'q'): ")
        print("\n")
        if file_idx == "q":
            sys.exit(0)
        # ask for user input for heading row skip
        header_row_input = input("Please provide the row number where the header is located (1-indexed): (exit with 'q'):")
        print("\n")
        if header_row_input == "q":
            sys.exit(0)
     
        try:
            file_idx = int(file_idx)
            filepath_str = xlsx_files[file_idx]
            file_selected = True
            return dict(filepath=filepath_str, header_row=int(header_row_input))
        
        except ValueError:
            print("You must enter a valid index number!")
        except IndexError:
            print("This number does not point to an existing file!")


def read_file_for_fuzzy(filepath, header_row):
    """
    """
    print("Loading {}, please wait...".format(filepath))
    
    df = pd.read_excel(filepath, header=header_row-1)
    print("{} loaded successfully!".format(filepath))
    return df


def select_fuzzy_column(filename, df):  
    """
    Allows for user selection of the column to use for Fuzzy Lookup within a pandas dataframe
    
    Returns the 0-index value of the selected column
    """
    
    print("Here are the columns found in {}:".format(filename))
    for idx, col in enumerate(df.columns):
        print("{}: {}".format(idx, col))
    
    fuzzy_column_selected = False
    while fuzzy_column_selected == False:
        fuzzy_column_idx = input("Please select a column to use for the fuzzy match (use the index number / exit with 'q'): ")
        if fuzzy_column_idx == "q":
                sys.exit(0)
        try:
            fuzzy_column_idx = int(fuzzy_column_idx)
            fuzzy_column_selected = True
        except ValueError:
            print("You must enter a valid number!")
        except IndexError:
            print("This number does not point to an existing column!")
    
    print("Column '{}' selected.".format(df.columns[fuzzy_column_idx]))
    print("\n")
    return fuzzy_column_idx

# Main thread
def main():
    # Startup
    print("------------------------------------")
    print("__    ____         __    __    __ ",
           "{}   /{}{}|       {}{}   {}   /{}",
           " {} /{} {}|     ({}   {}  {} /{} ",
           "  {}{}  {}|       {}       {}{}  ",
           " /{}{}  {}|         {}    /{}{}  ",
           "/{}  {} {}|___  {}    {} /{}  {} ",
           "{}    {}{}{}{}{}  {}{}  /{}    {}", sep="\n")
    print("----------------------------------")    
    print("\n")
    print("Author: Brian",
          "author_email:'tennybriank@gmail.com'", sep="\n")
    
    # Setting global input files parameters
    print("\n")
    print("STEP 1: SELECT A FILE WITH THE VALUES TO MATCH")
    first_file = select_file()
    print("\n")
    print("STEP 2: SELECT A FILE WITH THE VALUES TO BE MATCHED AGAINST")
    second_file = select_file()
    
    # Reading files as pandas dataframes
    first_fuzzydf = read_file_for_fuzzy(first_file['filepath'], first_file['header_row'])
    print("\n")
    second_fuzzydf = read_file_for_fuzzy(second_file['filepath'], second_file['header_row'])
    
    print("\n")
    
    print("STEP 3: SELECT THE COLUMNS TO MATCH")
    # Select columns to match:
    first_fuzzycolumn = select_fuzzy_column(first_file['filepath'], first_fuzzydf)
    print("\n")
    second_fuzzycolumn = select_fuzzy_column(second_file['filepath'], second_fuzzydf)
    
    print("\n")
    print("SETUP COMPLETE")

    # Read in the first xlsx file
    #first_fuzzydf = pd.read_excel('File2.xlsx')

    # Read in the second xlsx file
    #second_fuzzydf = pd.read_excel('file1.xlsx')

    #Create a new column in the second dataframe with the comparison results
    second_fuzzydf['On Scan'] = second_fuzzydf['number'].apply(lambda x: 'Yes' if x in first_fuzzydf['matched_job'].values else 'No')

    # Write the modified second dataframe to a new xlsx file
    second_fuzzydf.to_excel('File3 - Sample of Weekly report.xlsx', index=False)
    
    open_file = input("XLSX Match Complete! File 'File3 - Sample of Weekly report.xlsx' written to disk. Do you wish to open it (y/n)?")
    if open_file == 'y':
        os.system('start File3 - Sample of Weekly report.xlsx')
    else:
        pass


    
if __name__ == "__main__":
    main()