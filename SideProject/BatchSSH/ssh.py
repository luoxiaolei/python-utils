# 依赖
# pip install paramiko
# pip install XlsxWriter==3.0.3

# 遍历Excel
# 执行登录操作
# 保存IP、登录成功的密码、返回的结果
# 写入result

import sys
import paramiko
import xlrd  # 导入库
import xlwt
import os
from shutil import copyfile
from xlutils.copy import copy
import xlsxwriter as xw

sshExcel = xlrd.open_workbook(os.path.join(os.path.abspath(os.path.dirname(__file__)), 'ssh.xlsx'))
sheet = sshExcel.sheets()[0]

workbook = xw.Workbook(os.path.join(os.path.abspath(os.path.dirname(__file__)), 'result.xlsx'))  # 创建工作簿
worksheet1 = workbook.add_worksheet("sheet1")  # 创建子表
worksheet1.activate()  # 激活表
title = ['IP', '真实密码', '查询结果']  # 设置表头
worksheet1.write_row('A1', title)


initRowNum = 1
collMinNum = 0
collMaxNum = 4
rowCount = sheet.nrows # 获取总行数
rowNum = 1
ip = ""
port = ""
username = ""
password = ""
passwordB = ""
truePassword = ""
info = ""
rowResult = []

def getSSHInfo(ip,port,username,password,passwordB):
    print(ip)
    print(port)
    print(username)
    print(password)
    print(passwordB)
    #创建一个ssh的客户端，用来连接服务器
    ssh = paramiko.SSHClient()
    #创建一个ssh的白名单
    know_host = paramiko.AutoAddPolicy()
    #加载创建的白名单
    ssh.set_missing_host_key_policy(know_host)
    
    #连接服务器
    try:
        ssh.connect(
            hostname = ip,
            port = 22,
            username = username,
            password = password
        )
        truePassword = password
    except paramiko.ssh_exception.AuthenticationException:
        print("第一次密码错误")
        try:
            ssh.connect(
                hostname = ip,
                port = 22,
                username = username,
                password = passwordB
            )
            truePassword = passwordB
        except paramiko.ssh_exception.AuthenticationException:
            print("两次密码都错误")
            truePassword = "密码错误"
            return ip,truePassword,"密码错误无法获取命令结果"
        else:
            print("第二次登录成功")
    else:
        print("第一次登录成功")

    #执行命令
    stdin,stdout,stderr = ssh.exec_command("cd /etc && cat system-release")
    #stdin  标准格式的输入，是一个写权限的文件对象
    #stdout 标准格式的输出，是一个读权限的文件对象
    #stderr 标准格式的错误，是一个写权限的文件对象
    info = stdout.read().decode()
    # print(stdout.read().decode())
    stdin.close()
    ssh.close()
    return ip,truePassword,info

while rowNum < rowCount:
    # 遍历每一行的每一个单元格
    for collIndex in range(collMinNum,collMaxNum+1):
        # print(str(sheet.row(initRowNum)[collIndex].value).strip())
        # 获取当前单元格的值
        cellValue = str(sheet.row(initRowNum)[collIndex].value).strip()
        #print(cellValue)
        match collIndex:
            case 0:
                # IP
                ip = cellValue
            case 1:
                # 端口
                port = cellValue
            case 2:
                # 用户名
                username = cellValue
            case 3:
                # 密码1
                password = cellValue
            case 4:
                # 密码2
                passwordB = cellValue
    rowResult = getSSHInfo(ip,port,username,password,passwordB)
    print(rowResult)   
    worksheet1.write_row('A' + str(rowNum+1), rowResult)
    rowNum = rowNum + 1
    initRowNum = initRowNum + 1
 
workbook.close()