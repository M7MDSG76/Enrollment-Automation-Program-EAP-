# step 1
# Import the Librarys
import openpyxl as xl
import xlsxwriter
import pandas as pd 
import timestring

# step 2
# Set the Excel file path
workBook1 = 'C:\\Users\\M7MD\\Desktop\\Work\\aaa23august.xlsx'
workBook2 = 'Qualification Completion Records Report Aug,23.xlsx'
new_workbook = 'workbook.xlsx'

# step 3
# Load the sheet in an object and skip the unneeded rows
ws1 = pd.read_excel(workBook1, 1,skiprows = 3)
ws2 = pd.read_excel(workBook2, 1,skiprows = 3)


# step 4

# make list of the required columns names.
ws2Columns = ['Name', 'External ID', 'Completed']
ws1Columns = ['Activity', 'External ID', 'Completed']

# step 5
#make new table from the required columns
table = ws2[ws2Columns]



# step 6
# Convert the pivot column(Activity or Name) to list (this part used to help filtring the courses na)
nameList = table['Name'].to_list()


# step 7
# exctractingCoursesNames function exctract the courses names to nameList to help in the filtering part

def exctractingCoursesNames(list):
    n = len(list)
    newList = []
    for i in range(0,n):
        if i == 0 :
            newList.append(list[i])
            
        elif i > 0 :
            if list[i] == list[i-1]:
               
                continue
            else:
                     newList.append(list[i]) 
                     
    return  newList         
        
coursesList = exctractingCoursesNames(nameList)   


# step 8
# set the dates
startDate = timestring.Date(input('Enter start Date:'))
endDate = timestring.Date(input('Enter End Date: '))



# timestring library is a great library that converts any string to date type object. 


# filter the table by date ((NOT USED))
def filterDFDate(dataFrame, startDate, EndDate):
    
    for i in range(len(dataFrame)):
        timestring.Date(dataFrame['Completed'][i])    # convert the date from string type to date type
    dataFrame[dataFrame['Completed']>startDate]
    
    return dataFrame

# step 9

table = table[table['Completed']>startDate]
table.reset_index(drop=True) # reset the index to start from Zero insted of the defult indexes values.


# step 10

table['External ID'] = table['External ID'].astype(int) # convert "External ID" to int.



# step 11 
# Create new dataframe for each course and save into list of dataframes called {finalTable}.


finalTable = {} # this list contains dataframe for each course in {coursesList}.

# this loop takes courses names from {coursesList} to varible then use the name of the course to create table for that course. 
for i in range(len(coursesList)):
    CourseName = [coursesList[i]]  # this list containes name of the course. 
    
    print('Course Name: ',CourseName) # Example  ['Confined Space Entry - Permit Required']
    
    finalTable[i] = table[table.Name.isin(CourseName)] # isin() function only accept list.
    finalTable[i] = finalTable[i].reset_index(drop= True)
    

    
# step 12

# TEST IT !!!!
# I want to make dataframe for each course

df_list = [] # df_list is list contains a 'pandas.core.frame.DataFrame' in each index

for i in range(len(finalTable)): 
    c = ['External ID']         
    ID = finalTable[i][c] # [i] is the index of the DF , and [c] is the column name to call.
    df_list.append(ID)
    df_list[i] = df_list[i].reset_index(drop=True)
    
print(type(df_list[0]))

# step 13 
# create new excel file to write the results on it
writer = pd.ExcelWriter(new_workbook, engine='openpyxl')

# Write the DFs in the excel work book
for i in range(len(finalTable)):
    
    print(i)
    print(df_list[i],' - ', coursesList[i])
    df_list[i].to_excel(writer, sheet_name=coursesList[i][:10]) # write DF on sheet named with the first 10 charachters of Activity name
    
# Save the workbook on the excel file.
writer.save() 

