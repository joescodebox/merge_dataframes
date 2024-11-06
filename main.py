import os
import pandas as pd
import numpy as np

#Harvested from Control Group Data, this is all the names used for color tests (blank was also used, but not useful for this)
COLOR_TEST_NAMES = ['DE', 'DELTA E', 'DL', 'B', 'L', 'A', 'DA', 'COLOR', 'DB', 'D.L', 'D.A', 'D.B']

#Working Directory location
FOLDER_PATH  = (" ")

#Silence warning
pd.set_option('future.no_silent_downcasting', True)

#
#Joseph Demey
#main.py for Color Data Extraction
#November 4, 2024
#Version 00.00A
#       Built and planned functions, used skeletal code to express
#       Worked in IDLE to trial and error functions
#
#November 5, 2024
#Version 01.00A
#       Prototype of working version. Has all required functionality except write to excel, creates a big dataframe with all data.
#       Utilized program to gather data and save
#

def load_documents(FOLDER_PATH):
    """
    Iterate through directory of quality control cards and open each file, then adds it to a list of dataframes
    which can be concatenated and then written to a spreadsheet
    """
    df_list=[]
    for filename in os.listdir(FOLDER_PATH):
        if filename.endswith('.xls') or filename.endswith('.xlsx'):
            file_path = os.path.join(FOLDER_PATH, filename)
            try:
                df = pd.read_excel(file_path, usecols="A:P", header=None)
            except:
                print("error in : ", str(file_path))

            #Get the part number for this file, then get rid of all rows up to the test block
            part_num = df.iloc[0][1]
            df = condition_df(df)

            #Run the filter columns function to extract all columns where index
            df = filter_columns(df, part_num)

            df_list.append(df)
        #If the file is not an excel file, go to the next file
        else:
            pass
    return(df_list)

def condition_df(df):
    '''
    Condition the data frame to be more workable in the future.

    Steps:
        Trim the header off the file now that the part num has been captured in the head
        Reset the index to avoid messing up calls to specific cells
        Remove all white space and turn all cells to upper case to ensure better matching
    
    Input: Dataframe fresh from read_excel
    Output: Conditioned dataframe    
    '''
    df = df[:][4:] #sets the dataframe equal to itself, without the first four rows
    df.reset_index(inplace=True, drop=True) #resets the index to make it a little easier for me to visualize while programming

    #Make all data upper case and remove all white space to ensure matching works properly
    df = df.map(lambda x: x.strip().upper() if isinstance(x, str) else x) 
    
    return df

def filter_columns(df, part_num):
    '''
    Iterate through columns of excelsheet dataframe to find the desired values in rows 5, 7, or 9. We keep these if all three are blank
    in the use case of an excel document which has data without labels in the following section.

    Input: df - the dataframe from excel
            part_num - the part number from the df

    Output: new_df - the dataframe stripped of the columns which are not color related. In the case of no color related columns, returns
            an empty dataframe    
    '''
    #Bradley final results occur every other row, which is very usefully, the rows we need to check
    #for if they are labeled for color values, so we will filter out every other row
    sub_df = df[::2][:]

    #Then reset the index before transposing so it's less complicated to call cells
    sub_df.reset_index(inplace=True, drop=True)

    #I find it easier to work with if the rows are the tests, so flip it
    sub_df = sub_df.transpose()

    #Test for if the data is part of the color test names
    sub_df = sub_df[(sub_df[0].isin(COLOR_TEST_NAMES)) | (sub_df[1].isin(COLOR_TEST_NAMES)) | (sub_df[2].isin(COLOR_TEST_NAMES))]

    #Replace all white space with NaN and then remove all NaN columns from dataset
    sub_df = sub_df.replace('',np.nan)
    sub_df = sub_df.dropna(how='all', axis=1)

    #Return it back to the column header style for use in the main function
    sub_df = sub_df.transpose()

    #Reset the index once again
    sub_df.reset_index(inplace=True,drop=True)

    if sub_df.empty:
        return sub_df
    else:
        #Set the headers to be equal to the color tests which were found, keep the data
        sub_df, sub_df.columns = sub_df[1:] , sub_df.iloc[0]
        #insert part number into dataframe cell
        sub_df.insert(loc=0,column='PART NUMBER',value=str(part_num))
        #Pass it back, filtering is done
        return sub_df

def send_to_spreadsheet(df):
    '''
    This function will take the qc_test_specs dataframe and send it into an excel formated document
    '''
    df.to_excel("/output/output.xlsx")

#iterate through all files and build a list of QC tests to make the column of spreadsheet
total = load_documents(FOLDER_PATH)
final_df = pd.concat(total)

print(final_df)
send_to_spreadsheet(final_df)