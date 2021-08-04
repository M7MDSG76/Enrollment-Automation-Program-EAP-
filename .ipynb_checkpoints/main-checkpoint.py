import pandas as pd
import openpyxl


# setting up the workbook 
file = ('WB.xlsx')
xlsx = pd.ExcelFile('WB.xlsx')



# check the file if working
print(xlsx.sheet_names)

# Create Data Frame
df3 = pd.DataFrame(data = [[1,55,'=SUM(A2:B2)'],[1,4,'=SUM(A3:B3)'],[22,2,'=SUM(A4:B4)']], columns = ['A','B','SUM'])


# write the Data Frame to excel file
df3.to_excel(xlsx, sheet_name = 'sum', index = False)


workbook = openpyxl.load_workbook(xlsx, data_only = True)
worksheet = workbook['sum'] 
dataFrame = pd.DataFrame(worksheet.values)
workbook.save(xlsx)

r_df = pd.read_excel(xlsx, sheet_name = 'sum')
print(r_df)
