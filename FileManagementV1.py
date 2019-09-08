'''
@Description: File System Management
@Author: AppleCai
@Date: 2019-09-08 9:25:19
@LastEditTime: 2019-09-08 14:24:13
'''
import os
import time
import xlrd
import xlwt
from xlutils.copy import copy

myTagSearchKeyWord="Eng"  #搜索标签列中的关键字
myFileSystem=[]    #excel表格中的list
myfilelist=[]      #实际扫描当前的文件list

excelfile = 'AppleCai的文件管理系统.xls' #文件信息系统名称
file_dir = r"F:\t1" #需要管理的文件夹
col_data=1  #文件修改时间在第1列
col_name=0  #文件名在第0列

'''
 * @description: 对已经存在于表格中的数据进行更新
 * @param：path是Excel文件名称，value是要写入文件的所有数据
 * @return: 无
'''
def UpdateExcel(path, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlwt.Workbook()  # 新建一个工作簿
    sheet = workbook.add_sheet("目录")  # 在工作簿中新建一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            sheet.write(i, j, value[i][j])  # 像表格中写入数据（对应的行和列）
    workbook.save(path)  # 保存工作簿
    print("表格更新数据成功！")

'''
 * @description: 读取已存在文件的信息
 * @param：path是Excel文件名称
 * @return: 读取的行数
'''
def readExcel(path):
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    for i in range(worksheet.nrows):
        col = worksheet.row_values(i)  ##获取每一列数据
        #print(col)
        myFileSystem.append(col)
    return(len(myFileSystem))

'''
 * @description: 即时扫描文件夹获取目录下的文件信息
 * @param：file_dir是当前扫描的目录
 * @return: 
'''
def fileReader(file_dir):
    for root, dirs, files in os.walk(file_dir):
        for file in files:
            full_path = os.path.join(root, file)
            #需要进行管理的文件名称后缀
            if((os.path.splitext(full_path)[1]=='.txt') or
                (os.path.splitext(full_path)[1] == '.xls') or
                (os.path.splitext(full_path)[1] == '.xlsx') or
                (os.path.splitext(full_path)[1] == '.pdf')  or
                (os.path.splitext(full_path)[1] == '.doc')  or
                (os.path.splitext(full_path)[1] == '.docx') or
                (os.path.splitext(full_path)[1] == '.md')):
                absPath=os.path.dirname(full_path)       # 打印出来为双斜杠，所以需要修改
                absFileName=os.path.basename(full_path)
                mtime = os.path.getmtime(full_path)
                file_modify_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(mtime))
                #格式为文件名，修改时间，Tag，文件路径
                myfilelist.append((absFileName,file_modify_time,"",absPath.replace('\\','/')))
                #myfilelist.append(os.path.split(file_path)) #返回二元，一个为路径，一个为名称，但是没法直接替换斜杠

'''
 * @description: 将扫描后的文件信息和Excel中已经存在的信息对比，进行添加，删除，修改。并保持到myFileSystem中
 * @param：
 * @return: 
'''
def fileHandler():

    listdata = [x[col_data] for x in myFileSystem[1:]]  #获取所有文件修改时间
    listName = [x[col_name] for x in myFileSystem[1:]]  #获取所有文件名称
    listNameNew= [x[col_name] for x in myfilelist]  #获取所有目前扫描出的文件名称
    #print(max(listdata))
    latestRecordData=max(listdata)
    Newlist = (list(set(listNameNew).difference(set(listName))))  # 实际，但是系统信息中没有，需要添加

    #通过判断最新时间来判断需要新增到Excel的追加列表
    for item in myfilelist:
        if((item[col_data]>latestRecordData) and item[col_name] not in listName):
            myFileSystem.append(item)
        elif((item[col_data]>latestRecordData) and item[col_name] in listName):
            index=listName.index(item[col_name])+1         #找到需要修改的行数，因为listName用了[1:]等于减去目录的一行，所以需要添加一行
            myFileSystem[index][col_data]=item[col_data]   #更新此行的文件修改时间
        elif((item[col_data]==latestRecordData) and item[col_name] not in listName): #重命名文件名是不修改文件时间的
            myFileSystem.append(item)
        elif((item[col_data]<latestRecordData) and item[col_name] in Newlist): #修复bug1
            myFileSystem.append(item)

    dellist=(list(set(listName).difference(set(listNameNew))))  #表格系统中有，但是实际文件中已经没有，需要删除
    for item in dellist:
        index = listName.index(item) + 1
        del myFileSystem[index]
'''
 * @description: 将list中的信息追加如Excel中
 * @param：path是Excel文件名称，value是要写入文件的所有数据
 * @return: 无
'''
def write_excel_append(path, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            new_worksheet.write(i+rows_old, j, value[i][j])  # 追加写入数据，注意是从i+rows_old行开始写入
    new_workbook.save(path)  # 保存工作簿
    print("写入文件信息成功！")
'''
 * @description: 按关键字myTagSearchKeyWord搜索标签列，并且打印出对应文件的信息
 * @param：
 * @return: 无
'''
def mysearch():
    for item in myFileSystem:
        if(myTagSearchKeyWord in item[3]):
            print("搜索到第",myFileSystem.index(item),"行，文件信息为",item)

if __name__ == "__main__":
    lineNum=readExcel(excelfile) #读取表格中已有文件信息
    fileReader(file_dir) #通过扫描文件夹获取当前文件信息
    if(lineNum>1): #原来文件数据不为空，则需要处理
        fileHandler()
        UpdateExcel(excelfile,myFileSystem)
    elif(lineNum==0): #原来为空文件，直接追加读取的文件信息
        myhead=[("文件名","修改时间","标签","文件路径")]
        write_excel_append(excelfile,myhead)
        write_excel_append(excelfile, myfilelist)
    else: #原来文件包括标题
        write_excel_append(excelfile, myfilelist)
    #用于表格手工添加完标签后的关键字搜索附加功能，目前尚未开发完全
    mysearch()

'''
bug1（已修复）：原来文件叫图3.txt，把它更新为名称图4.txt后再新建一个新的图3.txt，运行文件管理系统，图4.txt是识别不到的。
'''