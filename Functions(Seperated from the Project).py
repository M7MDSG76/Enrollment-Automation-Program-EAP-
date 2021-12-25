import openpyxl

def read_excel(filePath):  
    workbook = openpyxl.load_workbook(filePath, data_only = True)
    worksheet = wb['sum'] 
    dataFrame = pd.DataFrame(worksheet.values)
    workbook.save(filepath)
    return print(dataFrame)