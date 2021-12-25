# step 1
# Import the Librarys

import openpyxl as xl
import xlsxwriter
import pandas as pd
import timestring
from datetime import datetime

# step 2
# Set the Excel file path
try:
    path = input('Enter the Excel file path:  \n')
except:
    print('Make sure the path is correct.')


def load_DataFrame(path, skip_rows=3,
                    sheet_index=1):  # takes path of excel file and return a dataframe from the excel sheet  #1
    #  skip_rows: number of rows to skip.

    # sheet_index: the index of the sheet. Note: indexing start from 0.

    # Workbook named "WB-todaydate.xlsx"
    today_date = datetime.today().strftime('%Y-%m-%d')

    new_workbook = 'WB3_' + today_date + path[:10] + '.xlsx'

    # step 3
    try:
        # Load the sheet in an object and skip the unneeded rows
        main_DataFrame = pd.read_excel(path, sheet_index, skiprows=skip_rows)
        print(main_DataFrame.head(10))
        return main_DataFrame, new_workbook
    except:
        print('Make sure the path is correct.')


# Qualification Completion Records Report Aug,23.xlsx
load_DataFrame_return = load_DataFrame(path)

ws2 = load_DataFrame_return[0]  # DataFrame
new_workbook = load_DataFrame_return[1]  # sheet index

# step 4
# make list of the required columns names.
Columns1 = ['Name', 'External ID', 'Completed']
Columns2 = ['Activity', 'External ID', 'Completed']



# step 5
# make new table from the required columns
required_dataframe = ws2[Columns1]

# step 6
# Convert the pivot column(Activity or Name) to list (this part used to help filtring the courses na)
activity_list = required_dataframe['Name'].to_list()


# step 7
# exctractingCoursesNames function exctract the courses names to nameList to help in the filtering part

def exctracting_courses_names(list):    # 2
    n = len(list)
    newList = []
    for i in range(0, n):
        if i == 0:
            newList.append(list[i])

        elif i > 0:
            if list[i] == list[i - 1]:

                continue
            else:
                newList.append(list[i])

    return newList


coursesList = exctracting_courses_names(activity_list)




# step 8
# timestring library is a great library that converts any string to date type object.
# set the dates
# Date in : yyyy-mm-dd
startDate = timestring.Date(input('Enter start Date: '))
endDate = timestring.Date(input('Enter End Date: '))

# step 9
required_dataframe = required_dataframe[required_dataframe['Completed'] > startDate]
required_dataframe.reset_index(drop=True)  # reset the index to start from Zero insted of the defult indexes values.

# step 10
required_dataframe['External ID'] = required_dataframe['External ID'].astype(int)  # convert "External ID" to int.

# step 11
# Create new dataframe for each course and save into list of dataframes called {finalTable}.
finalTable = {}  # this list contains dataframe for each course in {coursesList}.

# this loop takes courses names from {coursesList} to varible then use the name of the course to create table for that course. 
for i in range(len(coursesList)):
    CourseName = [coursesList[i]]  # this list containes name of the course. 

    print('Course Name: ', CourseName)  # Example  ['Confined Space Entry - Permit Required']

    finalTable[i] = required_dataframe[required_dataframe.Name.isin(CourseName)]  # isin() function only accept list.
    finalTable[i] = finalTable[i].reset_index(drop=True)



# step 12
# TEST IT !!!!
# I want to make dataframe for each course

df_list = []  # df_list is list contains a 'pandas.core.frame.DataFrame' in each index

for i in range(len(finalTable)):
    c = ['External ID']
    ID = finalTable[i][c]  # [i] is the index of the DF , and [c] is the column name to call.
    df_list.append(ID)
    df_list[i] = df_list[i].reset_index(drop=True)

print(type(df_list[0]))

# step 13 
# create new excel file to write the results on it
writer = pd.ExcelWriter(new_workbook, engine='openpyxl')

# Write the DFs in the excel work book
for i in range(len(finalTable)):
    print(i)
    print(df_list[i], ' - ', coursesList[i])
    df_list[i].to_excel(writer, sheet_name=coursesList[i][:10])  # write DF on sheet named with the first 10 charachters of Activity name

# Save the workbook on the excel file.
writer.save()
