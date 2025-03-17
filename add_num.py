import os
import pandas as pd
import xlrd

# 配置路径（需要根据实际情况修改）
excel_path = "E:/Desktop/序号表.xls"  # excel文件路径（此行代码适用于xls文件）
#df = pd.read_excel(excel_path, engine="openpyxl")   #指定使用openpyxl引擎来读取.xlsx文件
#df = pd.read_excel(excel_path, engine="xlrd")  # 指定使用xlrd引擎读取.xls文件
#excel_path = os.path.expanduser('E:/Desktop/序号表.xlsx')  # Excel文件路径(看表格格式，此行代码适用于xlsx文件)
pdf_folder = os.path.expanduser('E:/Desktop/21104501/外省')  # PDF所在文件夹路径

# 读取Excel表格
df = pd.read_excel(excel_path)
name_to_number = dict(zip(df['姓名'], df['序号']))  # 创建姓名->序号的映射字典

# 遍历处理PDF文件~
for filename in os.listdir(pdf_folder):
    if filename.endswith('.pdf'):
        # 分离文件名和扩展名
        base_name, ext = os.path.splitext(filename)
        
        try:
            # 分割学号和姓名
            student_id, student_name = base_name.split('-', 1)
            
            # 查找对应序号
            if student_name not in name_to_number:
                print(f"⚠️ 未找到【{student_name}】对应的序号，跳过处理")
                continue
                
            serial_number = int(name_to_number[student_name])  # 转换为整数
            
            # 构建新文件名
            new_name = f"{student_id}-{serial_number}-{student_name}{ext}"
            
            # 执行重命名
            old_path = os.path.join(pdf_folder, filename)
            new_path = os.path.join(pdf_folder, new_name)
            os.rename(old_path, new_path)
            print(f"✅ 已重命名：{filename} -> {new_name}")
            
        except ValueError:
            print(f"❌ 文件名格式错误：{filename}，不符合'学号-姓名'格式")

print("\n处理完成！")