import shutil
import os
import openpyxl

directory_path='D:/PythonhSoftware/Orin/'
#整个就是一个for循环
work_book=openpyxl.load_workbook('D:/PythonhSoftware/Orin/文档.xlsx')
sheet_name=work_book.sheetnames
sheet=work_book[sheet_name[0]]#选择哪个工作表
#excel部分

#从excel获取姓名,和学号
name=''
st_code=0
which_type=input("最后后缀:")
for row in sheet.iter_rows(min_row=2,values_only=True):
    name=row[0]
    st_code=row[1]
    # 选取遍历目录的
    # 匹配字符串的部分
    part_name = name
    matching_file = []  # list
    for filename in os.listdir(directory_path):
        if part_name in filename:
            matching_file = filename

    orig_flie_path = f'D:/PythonhSoftware/Orin/{matching_file}'
    new_flie_path = f'D:/PythonhSoftware/copy/{name}-{st_code}-{which_type}.docx'
    shutil.copy(orig_flie_path, new_flie_path)
    # 目前测试成功


