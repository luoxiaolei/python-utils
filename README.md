# python-utils
开箱即用的Python小工具。

# 初衷
工作中常会用Python写一些小工具，写的多了发现有好多代码片段是可以复用的，一并放到这里。

# 实现功能
## Excel
1. [新建Excel]()
2. [新建Excel并设置表头]()
3. [获取指定单元格内容]()

## 数据库 
1. [获取MySQL连接](https://github.com/luoxiaolei/python-utils/blob/9739de9406e5ec33ef439307b50dca500ddec56d/utils/DBUtils.py#L10)
2. [获取MongoDB连接](https://github.com/luoxiaolei/python-utils/blob/9739de9406e5ec33ef439307b50dca500ddec56d/utils/DBUtils.py#L18)

## 网络

## 依赖模块
DB
```
pip install pymongo==3.12.3
pip install mysql-connector==2.2.9
```

Excel
```
pip install XlsxWriter==3.0.3
pip install xlrd==1.2.0
```

# Side Project
1. [拆分Excel](SideProject/SplitExcel/main.py)

    将一个Excel按行拆分为多个Excel。

2. [批量SSH后获取指定命令结果](SideProject/BatchSSH/ssh.py)

    从Excel批量读取登录信息，执行命令后的结果写入新的Excel。

3. [登录获取token后批量下载文件](SideProject/BatchDownload/batchDownload.py)

    模拟登录操作，获取token，携带token请求文件下载地址实现。