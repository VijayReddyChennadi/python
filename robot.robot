*** Settings ***
Library    RPA.Excel.Files
Library    Collections
Library    OperatingSystem

*** Variables ***
${FILE1_PATH}     path/to/your/File1.xlsx
${FILE2_PATH}     path/to/your/File2.xlsx
${SHEET_NAME}     Sheet1
${COLUMN_NAME}    ColumnA    # The header name of the column you want to compare

*** Test Cases ***
Compare Columns in Two Excel Files
    [Documentation]    This test case compares columns in two different Excel files.
    
    # Open the first file and get the column data
    Open Workbook    ${FILE1_PATH}
    ${file1_table}=    Read Worksheet As Table    ${SHEET_NAME}
    Close Workbook

    # Extract the column data by the column header
    @{file1_data}=    Get Column Values    ${file1_table}    ${COLUMN_NAME}

    # Open the second file and get the column data
    Open Workbook    ${FILE2_PATH}
    ${file2_table}=    Read Worksheet As Table    ${SHEET_NAME}
    Close Workbook

    # Extract the column data by the column header
    @{file2_data}=    Get Column Values    ${file2_table}    ${COLUMN_NAME}

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

*** Keywords ***
Get Column Values
    [Arguments]    ${table}    ${column_name}
    ${column_values}=    Create List
    FOR    ${row}    IN    @{table}
        ${value}=    Get From Dictionary    ${row}    ${column_name}
        Append To List    ${column_values}    ${value}
    END
    [Return]    ${column_values}

Convert To Set
    [Arguments]    ${list}
    ${set}=    Evaluate    set(${list})    # Converts list to set using Python's set function
    [Return]    ${set}

Set Difference
    [Arguments]    ${set1}    ${set2}
    ${difference}=    Evaluate    ${set1} - ${set2}
    [Return]    ${difference}
