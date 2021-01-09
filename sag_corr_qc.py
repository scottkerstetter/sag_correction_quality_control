"""
SAG CORRECTION QUALITY CONTROL

Loop through files in a directory, extract data, load into dataframe,
and export to a csv file.

Designed for use with corrected survey sheets collected from MagVar (Saphira).

Directories are defined by user.  Make sure output files are saved in
a different location than survey sheets.
"""

# script and version info
script = 'sag_corr_qc.py'
version = '1.0.0'
author = 'S Kerstetter'
created = '2021-01-08'
last_modified = '2021-01-08'


# import modules
import pandas as pd
import datetime as dt
import os


# user-defined paths
directory = "C:\\Users\\kerst997\\Documents\\survey_sheets\\"
save_directory = "C:\\Users\\kerst997\\Documents\\sag_correction_reports\\"


# output file name
today = dt.date.today()
output_file = 'sag_correction_data_' + str(today) + ".csv"


# loop through files in a directory
def loop_files(directory):
    
    output_df = pd.DataFrame(columns=['wellname', 'sag_corrected', 'start_location', 'start_md', 'tool_code', 'previous_svy', 'previous_code'])
    
    for filename in os.listdir(directory):
        # select file and extract data
        target_file = os.path.join(filename)
        extracted_data = extract_data(dirpath, target_file)
        print(f"In file {target_file} we found: {extracted_data}")

        # append extracted data to end of dataframe
        end = len(output_df)
        output_df.loc[end] = extracted_data

    return output_df


# extract data from file, wellname and start of sag corrections
def extract_data(dirpath, file):
    # extract well name
    temp_df = pd.read_excel(dirpath+file, sheet_name='Minimum Curvature', header=None)
    wellname = temp_df.loc[2,2]
    
    # reload data with column headers, drop extraneous columns and rows
    temp_df = pd.read_excel(dirpath+file, sheet_name='Minimum Curvature', header=5)
    temp_df = temp_df.drop(['Azimuth', 'Vertical Depth', 'Northings Local', 'Eastings Local', 'Vertical Section', 'Dogleg Rate'], axis=1)
    temp_df = temp_df.drop(0)
    
    #extract data from file
    for i in range (1, len(temp_df)):
        if "SAG" in temp_df['Tool Code'][i]:
            sag_corrected = True
            start_md = temp_df['Measured Depth'][i]
            tool_code = temp_df['Tool Code'][i]
            previous_svy = temp_df['Measured Depth'][i-1]
            previous_code = temp_df['Tool Code'][i-1]
            break
        else:
            sag_corrected = False
    
    # if well has any amount of sag correction, find location of sag correction and compile data.
    # else compile compile well name and false sag correction
    if sag_corrected is True:
        start_location = find_hole_section(temp_df, start_md)
        file_data = [wellname, sag_corrected, start_location, start_md, tool_code, previous_svy, previous_code]
    else:
        file_data = [wellname, sag_corrected, '', '', '', '', '']
    
    return file_data 


# determines position in curve via inc and neighboring surveys
def find_hole_section(df, md):
    
    # collect KOP, LP, and TD values
    kop, lp, td = find_kop_lp_td(df)
    
    # using KOP, LP, and TD, determine where in well sag corrections began.
    if md < kop and md >= 0:
        hole_section = 'VERTICAL'
    elif md == kop:
        hole_section = 'KOP'
    elif md == lp:
        hole_section = 'LP'
    elif md > kop and md < lp:
        hole_section = 'CURVE'
    elif md > lp and md < td:
        hole_section = 'LATERAL'
    else:
        hole_section = "INVALID"
    
    return hole_section


# determines location of KOP, LP, and TD
def find_kop_lp_td(df):
    # find TD
    end = len(df)
    td = df['Measured Depth'][end]
    
    # find KOP
    while True:
        if df['Inclination'][end] < 4:
            kop = df['Measured Depth'][end]
            break
        end -= 1
    
    # find LP
    for i in range(1, len(df)):
        if df['Inclination'][i] >= 88:
            lp = df['Measured Depth'][i]
            break

    return kop, lp, td


# *** Start Script ****
print(f"""
SAG CORRECTION QUALITY CONTROL

Credits
Script: {script}
Date Created: {created}
Author: {author}
Version: {version}
Last Modified: {last_modified}
""")

print("Processing Feed:")
output_df = loop_files(dirpath)
output_df.head()
output_df.to_csv(save_directory+output_file)