# 依赖
# pip install xlrd==1.2.0
# pip install xlwt==1.3.0
# pip install xlutils==2.0.0

# 实现逻辑
# 读取总表，循环逐行读取。
# 读一行，通过模板复制一个文件，写入信息。
# 读不到了，退出循环。

import xlrd  # 导入库
import xlwt
import os
from shutil import copyfile
from xlutils.copy import copy

xlsx = xlrd.open_workbook(os.path.join(os.path.abspath(os.path.dirname(__file__)), '总表.xls'))
sheet = xlsx.sheets()[0]

initRowNum = 2
collMinNum = 0
collMaxNum = 23
xuhao =  str(sheet.row(2)[0].value).strip()

# 修改单元格的值丢失边框，通过添加边框样式来增加边框。
borders = xlwt.Borders()
borders.left = xlwt.Borders.THIN    # 添加边框-虚线边框
borders.right = xlwt.Borders.THIN  
borders.top = xlwt.Borders.THIN   
borders.bottom = xlwt.Borders.THIN  

borders.left_colour = 0x00            # 边框上色
borders.right_colour = 0x00
borders.top_colour = 0x00
borders.bottom_colour = 0x00

style = xlwt.XFStyle()
style.borders = borders

# 添加左边框
bordersLeft = xlwt.Borders()
bordersLeft.left = xlwt.Borders.THIN    # 添加边框-虚线边框

bordersLeft.left_colour = 0x00            # 边框上色

styleLeft = xlwt.XFStyle()
styleLeft.borders = bordersLeft

while len(xuhao) != 0:
    # 从模板复制一个文件，并读取这个文件
    fileName = str(sheet.row(initRowNum)[6].value).strip()
    newFilePath = os.path.join(os.path.abspath(os.path.dirname(__file__)),xuhao+fileName+'.xls')
    copyfile(os.path.join(os.path.abspath(os.path.dirname(__file__)), '分表.xls'), newFilePath)
    # 读取文件必须为xls，否则会报错
    currentXlsx = xlrd.open_workbook(newFilePath,formatting_info=True)
    wb = copy(currentXlsx)
    currentSheet = wb.get_sheet(0)
    # 遍历每一行的每一个单元格
    for collIndex in range(collMinNum,collMaxNum+1):
        # print(str(sheet.row(initRowNum)[collIndex].value).strip())
        # 获取当前单元格的值
        cellValue = str(sheet.row(initRowNum)[collIndex].value).strip()
        # 根据列的索引确定单元格内容
        match collIndex:
            case 0:
                # 序号
                currentSheet.write(2,2,cellValue) # 将单元格内容写入新建的Excel对应单元格
            case 1:
                # 地区
                currentSheet.write(3,2,cellValue,style)
            case 2:
                # 市县区
                currentSheet.write(3,4,cellValue,style)
            case 3:
                # 井盖类型
                currentSheet.write(2,4,cellValue)
            case 4:
                # 编号
                currentSheet.write(2,6,cellValue)
            case 5:
                # 公共区域及场所
                currentSheet.write(4,2,cellValue,style)
            case 6:
                # 位置描述
                currentSheet.write(5,2,cellValue,style)
                currentSheet.write(5,7,'',styleLeft) # 合并后的单元格缺少右边框，添加下一个单元格的左边框进行实现。
            case 7:
                # 坐标
                currentSheet.write(4,4,cellValue,style)
            case 8:
                # 产权单位
                currentSheet.write(6,2,cellValue,style)
                currentSheet.write(6,7,'',styleLeft)
            case 9:
                # 生产厂家
                currentSheet.write(7,2,cellValue,style)
                currentSheet.write(7,7,'',styleLeft)
            case 10:
                # 井盖生产时间
                currentSheet.write(8,2,cellValue,style)
                currentSheet.write(8,7,'',styleLeft)
            case 11:
                # 井盖材质
                currentSheet.write(9,2,cellValue,style)
                currentSheet.write(9,7,'',styleLeft)
            case 12:
                # 是否为智能井盖
                currentSheet.write(10,2,cellValue,style)
                currentSheet.write(10,7,'',styleLeft)
            case 13:
                # 井盖规格
                currentSheet.write(11,2,cellValue,style)
                currentSheet.write(11,7,'',styleLeft)
            case 14:
                # 井盖承载等级
                currentSheet.write(12,2,cellValue,style)
                currentSheet.write(12,7,'',styleLeft)
            case 15:
                # 病害问题
                currentSheet.write(13,2,cellValue,style)
                currentSheet.write(13,7,'',styleLeft)
            case 16:
                # 风险等级
                currentSheet.write(14,2,cellValue,style)
                currentSheet.write(14,7,'',styleLeft)
            case 17:
                # 调查时间
                currentSheet.write(15,2,cellValue,style)
                currentSheet.write(15,7,'',styleLeft)
            case 18:
                # 计划整改时间
                currentSheet.write(16,2,cellValue,style)
                currentSheet.write(16,7,'',styleLeft)
            case 19:
                # 备注
                currentSheet.write(17,2,cellValue,style)
                currentSheet.write(17,7,'',styleLeft)
            case 20:
                # 填报单位
                currentSheet.write(19,2,cellValue)
            case 21:
                # 填报人
                currentSheet.write(19,4,cellValue)
            case 22:
                # 联系电话
                currentSheet.write(20,2,cellValue)
            case 23:
                # 填报时间
                currentSheet.write(20,4,cellValue)

        wb.save(newFilePath)
        # 写入对应的单元格
    initRowNum = initRowNum + 1
    xuhao =  str(sheet.row(initRowNum)[collMinNum].value).strip()