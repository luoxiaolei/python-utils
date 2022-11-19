import os
import xlsxwriter as xw
import xlrd

EXCEL_FILE_SUFFIX = ".xls"

def helloWorld():
    print("Hello World!")

# 新建Excel
def createExcel(path,fileName):
    fileFullName = path+os.sep+fileName+EXCEL_FILE_SUFFIX
    workbook = xw.Workbook(fileFullName)  
    worksheet = workbook.add_worksheet("sheet1")
    workbook.close()

# 新建Excel并设置表头
def createExcelAndTitle(path,fileName,titleArr):
    fileFullName = path+os.sep+fileName+EXCEL_FILE_SUFFIX
    workbook = xw.Workbook(fileFullName)  
    worksheet = workbook.add_worksheet("sheet1")
    worksheet.write_row('A1', titleArr)
    workbook.close()

# 获取指定单元格内容
def getCellContent(fileFullPath,sheetIndex,rowIndex,collIndex):
    xlsx = xlrd.open_workbook(fileFullPath)
    sheet = xlsx.sheets()[sheetIndex]
    return str(sheet.row(rowIndex)[collIndex].value).strip()