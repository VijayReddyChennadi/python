*** Settings ***
Library    RPA.Excel.Files
Library    Collections

*** Variables ***
${FILE1_PATH}     path/to/your/File1.xlsx
${FILE2_PATH}     path/to/your/File2.xlsx
${SHEET_NAME}     Sheet1
${COLUMN1}        A     # Column to compare in File1
${COLUMN2}        A     # Column to compare in File2

*** Test Cases ***
Compare Columns in Two Excel Files
    [Documentation]    This test case compares columns in two different Excel files.
    
    # Open the first file and get the column data
    Open Workbook    ${FILE1_PATH}
    ${file1_data}=    Read Column Values    ${SHEET_NAME}    ${COLUMN1}
    Close Workbook

    # Open the second file and get the column data
    Open Workbook    ${FILE2_PATH}
    ${file2_data}=    Read Column Values    ${SHEET_NAME}    ${COLUMN2}
    Close Workbook

    # Convert both columns to sets for easy comparison
    ${file1_set}=    Convert To Set    ${file1_data}
    ${file2_set}=    Convert To Set    ${file2_data}

    # Find values in File1 that are not in File2 and vice versa
    ${missing_in_file2}=    Set Difference    ${file1_set}    ${file2_set}
    ${missing_in_file1}=    Set Difference    ${file2_set}    ${file1_set}

    # Log the differences for review
    Log    Values in File1 not in File2: ${missing_in_file2}
    Log    Values in File2 not in File1: ${missing_in_file1}

    # Optionally, add validation steps to check if they are the same
    Should Be Empty    ${missing_in_file2}    msg=Some values from File1 are missing in File2
    Should Be Empty    ${missing_in_file1}    msg=Some values from File2 are missing in File1
