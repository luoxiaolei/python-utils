import ExcelUtils
import os

TEST_DIR = "temp"

ExcelUtils.helloWorld() 

ExcelUtils.createExcel(os.getcwd()+os.sep+TEST_DIR,"testCreateExcel")

ExcelUtils.createExcelAndTitle(os.getcwd()+os.sep+TEST_DIR,"testCreateExcelAndTitle",['序号','姓名'])

print(ExcelUtils.getCellContent(os.getcwd()+os.sep+TEST_DIR+os.sep+"testCreateExcelAndTitle.xls",0,0,0))