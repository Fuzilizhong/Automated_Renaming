import os
import pandas as pd

# ================== 配置区域 ==================
excel_path = "E:/Desktop/序号表.xls"  # Excel文件路径
docx_folder = "E:/Desktop/兵役文档"    # DOCX文件夹路径
# =============================================

def main():
    try:
        # 读取Excel（自动识别xls/xlsx）
        df = pd.read_excel(excel_path, header=0, names=['序号', '姓名'])
        
        # 数据清洗
        df = df.dropna(subset=['序号', '姓名'])  # 删除空行（如您示例中的"	张子扬"）
        df['序号'] = df['序号'].astype(int)      # 确保序号为整数
        name_to_number = dict(zip(df['姓名'], df['序号']))
        
    except Exception as e:
        print(f"❌ Excel读取失败: {e}")
        return

    success = 0
    skipped = 0
    errors = 0

    for filename in os.listdir(docx_folder):
        if not filename.endswith('.docx'):
            continue
        
        try:
            # 提取纯姓名（支持多种情况）
            student_name = os.path.splitext(filename)[0]\
                .replace("_副本", "")\
                .replace("（重命名）", "")\
                .strip()
            
            # 姓名匹配（模糊匹配增强）
            if student_name not in name_to_number:
                # 尝试去除空格匹配
                normalized_name = student_name.replace(" ", "")
                if normalized_name in name_to_number:
                    student_name = normalized_name
                else:
                    print(f"⚠️ 未匹配：{student_name}（原文件：{filename}）")
                    skipped +=1
                    continue
            
            # 生成新文件名
            new_filename = f"计算机学院-{name_to_number[student_name]}-{student_name}.docx"
            
            # 处理文件冲突
            counter = 1
            while os.path.exists(os.path.join(docx_folder, new_filename)):
                new_filename = f"计算机学院-{name_to_number[student_name]}-{student_name}_{counter}.docx"
                counter +=1
            
            # 执行重命名
            os.rename(
                os.path.join(docx_folder, filename),
                os.path.join(docx_folder, new_filename)
            )
            print(f"✅ 已处理：{filename} → {new_filename}")
            success +=1
            
        except Exception as e:
            print(f"❌ 处理失败：{filename}（错误：{str(e)}）")
            errors +=1

    # 统计报告
    print("\n=== 处理结果 ===")
    print(f"成功处理：{success} 个")
    print(f"跳过文件：{skipped} 个（未匹配姓名）")
    print(f"失败文件：{errors} 个")
    if df[df.duplicated(['姓名'])].shape[0] >0:
        print("\n警告：Excel中存在重复姓名！")
    if '_副本' in ' '.join(os.listdir(docx_folder)):
        print("提示：检测到包含'_副本'的文件，建议手动确认")

if __name__ == "__main__":
    print("=== DOCX智能重命名程序 ===")
    main()