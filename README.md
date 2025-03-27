"""
Auther: Zhijian Jin
Time:2025.3.14
"""

【使用说明】
此文件用于批量重命名各类文件，例如：.pdf/.doc/.docx文件等
例如“21104501xx-jzj.pdf”，要在学号和姓名中添加可以唯一标识的“序号”，修改为“学号-序号-姓名”，例如“21104501xx-1125-jzj”。
1、要先选择用于唯一标识目标文件的.xls文件（⚠关键！！！）
2、代码会通过表格中的标识栏自动检索目标文件夹中的文件并且建立字典一一对应
3、用户预设置要重命名的样式
4、点击“执行重命名”，即可批量完成重命名

百度网盘下载链接：通过网盘分享的文件：Autorenaming
链接: https://pan.baidu.com/s/1fa4zuJ3ftXM3w4f_77VpPw?pwd=1234 提取码: 1234

**User Manual**  
**Batch File Renaming Tool**  

This tool is designed for batch renaming various file formats (e.g., `.pdf`, `.doc`, `.docx`). For example, a file named `21104501xx-jzj.pdf` can be renamed to include a **unique identifier** (e.g., `21104501xx-1125-jzj`) by inserting a sequence number between the student ID and name.  

### **Key Features**  
1. **Excel Index File Requirement**  
   - A pre-configured `.xls` file containing unique identifiers (e.g., student IDs, sequence numbers, and names) is **mandatory** for operation.  
   - The tool uses this file to map target files to their corresponding identifiers.  

2. **Automated File Matching**  
   - The software automatically scans the target folder and matches files to entries in the Excel index using designated identifier columns.  

3. **Customizable Naming Format**  
   - Users can predefine the renaming format (e.g., `StudentID-SequenceNumber-Name`) through an intuitive GUI.  

4. **Batch Execution**  
   - Click the **Execute Rename** button to apply the renaming rules across all matched files.  

---

### **Download Link**  
- **Baidu Netdisk**:  
  - File: `Autorenaming`  
  - Link: [https://pan.baidu.com/s/1fa4zuJ3ftXM3w4f_77VpPw?pwd=1234](https://pan.baidu.com/s/1fa4zuJ3ftXM3w4f_77VpPw?pwd=1234)  
  - Extraction Code: `1234`  

---

### **Prerequisites**  
- An **index table** (`.xls` file) must be prepared in advance, containing the following:  
  - Unique identifiers (e.g., student IDs).  
  - Sequence numbers or other critical metadata for file mapping.  

### **Notes**  
- Ensure the index table’s format aligns with the target files’ naming conventions to avoid mismatches.  
- Test the tool on a small subset of files before full-scale execution.  

--- 

This tool streamlines file management workflows by ensuring consistency, accuracy, and efficiency in large-scale renaming tasks.
[注意事项]：需要提前拥有一个“序号表”文件用于标识唯一对应的姓名或学号
