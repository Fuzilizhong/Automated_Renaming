import sys
import os
import re
import pandas as pd
from PyQt6.QtWidgets import *
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QTextCursor


class RenameRuleWidget(QWidget):
    def __init__(self, columns, parent=None):
        super().__init__(parent)
        self.layout = QHBoxLayout(self)

        # 规则类型选择
        self.type_combo = QComboBox()
        self.type_combo.addItem("自定义文本")
        self.type_combo.addItems(columns)
        self.type_combo.currentTextChanged.connect(self._on_type_changed)

        # 内容输入
        self.content_input = QLineEdit()
        self.content_input.setPlaceholderText("输入自定义文本")

        # 分隔符选择
        self.separator_combo = QComboBox()
        self.separator_combo.addItems(["-", "_", " ", "自定义"])
        self.custom_separator = QLineEdit()
        self.custom_separator.setVisible(False)
        self.separator_combo.currentTextChanged.connect(self._on_separator_changed)

        # 删除按钮
        self.delete_btn = QPushButton("×")
        self.delete_btn.setFixedSize(20, 20)
        self.delete_btn.clicked.connect(self._delete_self)

        # 布局
        self.layout.addWidget(self.type_combo)
        self.layout.addWidget(self.content_input)
        self.layout.addWidget(QLabel("分隔符:"))
        self.layout.addWidget(self.separator_combo)
        self.layout.addWidget(self.custom_separator)
        self.layout.addWidget(self.delete_btn)

    def _on_type_changed(self, text):
        """当规则类型变化时，禁用或启用自定义输入框"""
        if text != "自定义文本":
            self.content_input.setEnabled(False)
            self.content_input.setText("")
        else:
            self.content_input.setEnabled(True)

    def _on_separator_changed(self, text):
        """当分隔符选择变化时，显示或隐藏自定义分隔符输入框"""
        self.custom_separator.setVisible(text == "自定义")

    def _delete_self(self):
        """删除当前规则项"""
        parent_layout = self.parent().layout()
        for i in range(parent_layout.count()):
            if parent_layout.itemAt(i).widget() == self:
                parent_layout.removeWidget(self)
                self.deleteLater()
                break

    def get_separator(self):
        """获取当前分隔符"""
        if self.separator_combo.currentText() == "自定义":
            return self.custom_separator.text()
        return self.separator_combo.currentText()


class FileSearchThread(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal(dict, list, list)

    def __init__(self, df, folder, columns, file_type, parent=None):
        super().__init__(parent)
        self.df = df
        self.folder = folder
        self.columns = columns
        self.file_type = file_type

    def run(self):
        try:
            results = {}
            unmatched_files = []
            multiple_matches = []
            for filename in os.listdir(self.folder):
                if filename.endswith(self.file_type):
                    self.progress.emit(f"正在处理: {filename}")
                    base_name = os.path.splitext(filename)[0]
                    matches = self.find_matches(base_name)
                    if not matches:
                        unmatched_files.append(filename)
                    elif len(matches) > 1:
                        multiple_matches.append(filename)
                    else:
                        results[filename] = matches[0]
            self.finished.emit(results, unmatched_files, multiple_matches)
        except Exception as e:
            self.progress.emit(f"错误: {str(e)}")
            import traceback
            self.progress.emit(traceback.format_exc())

    def find_matches(self, filename):
        """在 Excel 中查找与文件名匹配的行"""
        matches = []
        for index, row in self.df.iterrows():
            if all(str(row[col]).lower() in filename.lower() for col in self.columns):
                matches.append((index, row))
        return matches


class SmartRenamer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setup_ui()
        self.setup_connections()

    def setup_ui(self):
        # 主界面布局
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)

        # 操作步骤1: 文件选择区域
        file_group = QGroupBox("步骤1: 选择文件")
        file_layout = QGridLayout(file_group)

        # 文件类型选择
        self.type_combo = QComboBox()
        self.type_combo.addItems(["pdf", "docx", "doc", "其他"])

        # Excel文件选择
        self.excel_btn = QPushButton("选择Excel文件")
        self.excel_label = QLabel("未选择文件")

        # 目标文件夹选择
        self.folder_btn = QPushButton("选择目标文件夹")
        self.folder_label = QLabel("未选择文件夹")

        file_layout.addWidget(QLabel("文件类型:"), 0, 0)
        file_layout.addWidget(self.type_combo, 0, 1)
        file_layout.addWidget(self.excel_btn, 1, 0)
        file_layout.addWidget(self.excel_label, 1, 1)
        file_layout.addWidget(self.folder_btn, 2, 0)
        file_layout.addWidget(self.folder_label, 2, 1)

        # 操作步骤2: 匹配设置
        match_group = QGroupBox("步骤2: 匹配设置")
        match_layout = QVBoxLayout(match_group)

        self.column_list = QListWidget()
        self.search_btn = QPushButton("开始搜索匹配")
        self.status_output = QTextEdit()

        match_layout.addWidget(QLabel("选择匹配列（可多选）:"))
        match_layout.addWidget(self.column_list)
        match_layout.addWidget(self.search_btn)
        match_layout.addWidget(QLabel("匹配状态:"))
        match_layout.addWidget(self.status_output)

        # 操作步骤3: 重命名规则
        rule_group = QGroupBox("步骤3: 命名规则")
        rule_layout = QVBoxLayout(rule_group)

        self.rule_list = QListWidget()
        self.rule_list.setDragDropMode(QListWidget.DragDropMode.InternalMove)
        self.add_rule_btn = QPushButton("添加规则项")

        rule_layout.addWidget(QLabel("拖拽调整顺序:"))
        rule_layout.addWidget(self.rule_list)
        rule_layout.addWidget(self.add_rule_btn)

        # 执行按钮
        self.execute_btn = QPushButton("执行重命名")

        # 主布局
        layout.addWidget(file_group)
        layout.addWidget(match_group)
        layout.addWidget(rule_group)
        layout.addWidget(self.execute_btn)

        # 界面美化
        self.setStyleSheet("""
            QGroupBox {
                border: 1px solid gray;
                border-radius: 5px;
                margin-top: 1ex;
                font-weight: bold;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
            }
            QPushButton {
                padding: 5px;
                min-width: 80px;
            }
            QListWidget {
                border: 1px solid #cccccc;
                border-radius: 3px;
            }
        """)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(15)

    def setup_connections(self):
        # 连接信号槽
        self.excel_btn.clicked.connect(self.load_excel)
        self.folder_btn.clicked.connect(self.select_folder)
        self.add_rule_btn.clicked.connect(self.add_rule_item)
        self.search_btn.clicked.connect(self.start_file_search)
        self.execute_btn.clicked.connect(self.execute_rename)

    def load_excel(self):
        try:
            # 打开文件对话框，选择 Excel 文件
            path, _ = QFileDialog.getOpenFileName(
                self,  # 父窗口
                "选择 Excel 文件",  # 对话框标题
                "",  # 初始路径（空表示默认路径）
                "Excel 文件 (*.xls *.xlsx)"  # 文件过滤器
            )
            if path:  # 如果用户选择了文件
                self.excel_path = path
                self.excel_label.setText(os.path.basename(path))  # 显示文件名
                self.load_excel_columns()  # 加载 Excel 表头
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载 Excel 文件失败: {str(e)}")

    def select_folder(self):
        try:
            # 打开文件夹对话框，选择目标文件夹
            folder = QFileDialog.getExistingDirectory(
                self,  # 父窗口
                "选择目标文件夹",  # 对话框标题
                ""  # 初始路径（空表示默认路径）
            )
            if folder:  # 如果用户选择了文件夹
                self.target_folder = folder
                self.folder_label.setText(folder)  # 显示文件夹路径
        except Exception as e:
            QMessageBox.critical(self, "错误", f"选择文件夹失败: {str(e)}")

    def load_excel_columns(self):
        try:
            self.df = pd.read_excel(self.excel_path)
            self.column_list.clear()
            for col in self.df.columns:
                item = QListWidgetItem(col)
                item.setCheckState(Qt.CheckState.Unchecked)
                self.column_list.addItem(item)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取 Excel 失败: {str(e)}")

    def add_rule_item(self):
        if not hasattr(self, 'df') or self.df is None:
            QMessageBox.warning(self, "警告", "请先加载 Excel 文件！")
            return

        # 添加规则项
        rule_widget = RenameRuleWidget(self.df.columns.tolist(), self)
        item = QListWidgetItem()
        self.rule_list.addItem(item)
        self.rule_list.setItemWidget(item, rule_widget)
        item.setSizeHint(rule_widget.sizeHint())

    def start_file_search(self):
        if not self.target_folder or not self.excel_path:
            QMessageBox.warning(self, "警告", "请先选择目标文件夹和 Excel 文件！")
            return

        # 获取用户选择的匹配列
        selected_columns = []
        for index in range(self.column_list.count()):
            item = self.column_list.item(index)
            if item.checkState() == Qt.CheckState.Checked:
                selected_columns.append(item.text())

        if not selected_columns:
            QMessageBox.warning(self, "警告", "请至少选择一个匹配列！")
            return

        # 启动文件搜索线程
        self.search_thread = FileSearchThread(
            self.df, self.target_folder, selected_columns, self.type_combo.currentText(), self
        )
        self.search_thread.progress.connect(self.update_status)
        self.search_thread.finished.connect(self.on_search_finished)
        self.search_thread.start()

    def update_status(self, message):
        # 更新状态显示
        self.status_output.moveCursor(QTextCursor.MoveOperation.End)
        self.status_output.insertPlainText(message + "\n")
        self.status_output.ensureCursorVisible()

    def on_search_finished(self, results, unmatched_files, multiple_matches):
        try:
            if unmatched_files:
                self.update_status("以下文件未匹配到 Excel 数据:")
                for file in unmatched_files:
                    self.update_status(f"  - {file}")
            if multiple_matches:
                self.update_status("以下文件有多个匹配项:")
                for file in multiple_matches:
                    self.update_status(f"  - {file}")
            if results:
                self.update_status("匹配结果:")
                for file, (index, row) in results.items():
                    self.update_status(f"  - {file} → 表格第{index + 1}行")
            self.mapping_rules = results
        except Exception as e:
            QMessageBox.critical(self, "错误", f"处理匹配结果时出错: {str(e)}")

    def execute_rename(self):
        if not hasattr(self, 'mapping_rules') or not self.mapping_rules:
            QMessageBox.warning(self, "警告", "请先完成文件匹配！")
            return

        try:
            # 收集所有规则项
            rules = []
            for i in range(self.rule_list.count()):
                item = self.rule_list.item(i)
                rule_widget = self.rule_list.itemWidget(item)
                if isinstance(rule_widget, RenameRuleWidget):
                    rules.append(rule_widget)

            # 遍历匹配结果进行重命名
            success = 0
            errors = 0
            for filename, (index, row) in self.mapping_rules.items():
                try:
                    parts = []
                    for rule in rules:
                        # 处理规则项
                        if rule.type_combo.currentText() == "自定义文本":
                            text = rule.content_input.text()
                        else:
                            text = str(row[rule.type_combo.currentText()])
                        parts.append(text)
                        
                        # 添加分隔符（最后一项不加）
                        if rule != rules[-1]:
                            sep = rule.get_separator()
                            parts.append(sep)
                    
                    # 构建新文件名
                    new_name = ''.join(parts) + os.path.splitext(filename)[1]
                    old_path = os.path.join(self.target_folder, filename)
                    new_path = os.path.join(self.target_folder, new_name)
                    
                    # 处理重名冲突
                    counter = 1
                    while os.path.exists(new_path):
                        base, ext = os.path.splitext(new_name)
                        new_path = f"{base}_{counter}{ext}"
                        counter += 1
                    
                    os.rename(old_path, new_path)
                    success += 1
                except Exception as e:
                    errors += 1
                    self.update_status(f"❌ 重命名失败: {filename} ({str(e)})")

            # 显示结果
            QMessageBox.information(
                self,
                "完成",
                f"重命名完成！\n成功: {success} 个\n失败: {errors} 个"
            )
        except Exception as e:
            QMessageBox.critical(self, "错误", f"执行重命名时出错: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SmartRenamer()
    window.show()
    sys.exit(app.exec())