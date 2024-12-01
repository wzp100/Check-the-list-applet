import sys
import csv
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QLineEdit, QTableWidget, QTableWidgetItem, QCheckBox, QHeaderView, QAction, QFileDialog, QLabel, QCheckBox as QOptionCheckBox, QHBoxLayout, QMessageBox
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QKeyEvent
from pypinyin import pinyin, Style
from collections import defaultdict

class AttendanceApp(QMainWindow):
    update_signal = pyqtSignal(str)  # 定义信号

    def __init__(self):
        super().__init__()
        self.names = {}  # 存储姓名和考勤状态的字典
        self.initials_map = defaultdict(list)
        self.matched_names = []  # 存储当前可见的匹配姓名列表
        self.show_marked = True  # 是否显示已标记的人员

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('考勤系统')
        self.setGeometry(100, 100, 800, 600)

        # 菜单栏
        menubar = self.menuBar()
        file_menu = menubar.addMenu('文件')

        # 导入名单操作
        import_action = QAction('导入名单', self)
        import_action.triggered.connect(self.import_names)
        file_menu.addAction(import_action)

        # 导出考勤操作
        export_action = QAction('导出考勤情况', self)
        export_action.triggered.connect(self.export_attendance)
        file_menu.addAction(export_action)

        # 帮助菜单
        help_menu = menubar.addMenu('帮助')
        help_action = QAction('关于', self)
        help_action.triggered.connect(self.show_help)
        help_menu.addAction(help_action)

        # 中央小部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()

        # 输入框
        self.input_field = QLineEdit()
        self.input_field.setPlaceholderText('请输入拼音首字母进行搜索...')
        self.input_field.textChanged.connect(self.on_text_change)
        self.input_field.installEventFilter(self)  # 安装事件过滤器
        layout.addWidget(self.input_field)

        # 显示已标记的人员选项
        self.show_marked_checkbox = QOptionCheckBox("显示已标记的人员")
        self.show_marked_checkbox.setChecked(True)
        self.show_marked_checkbox.stateChanged.connect(self.on_show_marked_changed)
        layout.addWidget(self.show_marked_checkbox)

        # 表格小部件
        self.table = QTableWidget(0, 2)  # 初始为空
        self.table.setHorizontalHeaderLabels(['姓名', '考勤标记'])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setStyleSheet("QTableWidget { text-align: center; }")
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)  # 禁止编辑

        layout.addWidget(self.table)

        # 统计信息和搜索结果数量标签
        stats_layout = QHBoxLayout()
        self.stats_label = QLabel()
        self.stats_label.setAlignment(Qt.AlignLeft | Qt.AlignBottom)  # 将标签居于界面底部左端
        stats_layout.addWidget(self.stats_label)

        self.search_count_label = QLabel()
        self.search_count_label.setAlignment(Qt.AlignRight | Qt.AlignBottom)  # 将标签居于界面底部右端
        stats_layout.addWidget(self.search_count_label)

        layout.addLayout(stats_layout)

        self.update_stats()  # 初始化统计信息
        self.update_search_count()  # 初始化搜索结果数量

        central_widget.setLayout(layout)

    def import_names(self):
        file_path, _ = QFileDialog.getOpenFileName(self, '打开CSV文件', '', 'CSV 文件 (*.csv);;所有文件 (*)')
        if file_path:
            self.load_names_from_csv(file_path)

    def load_names_from_csv(self, file_path):
        self.names = {}
        try:
            with open(file_path, newline='', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                for row in reader:
                    if row:  # 跳过空行
                        name = row[0].strip()
                        self.names[name] = False
            self.initials_map = self._generate_initials_map(self.names.keys())
            self.display_all_names()
            self.update_stats()  # 更新统计信息
        except Exception as e:
            print(f"读取文件出错: {e}")

    def export_attendance(self):
        file_path, _ = QFileDialog.getSaveFileName(self, '保存考勤情况', '', 'CSV 文件 (*.csv);;所有文件 (*)')
        if file_path:
            try:
                with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(['姓名', '状态'])
                    for name, status in self.names.items():
                        writer.writerow([name, '出勤' if status else '缺勤'])
                print(f"考勤情况已导出至 {file_path}")
            except Exception as e:
                print(f"写入文件出错: {e}")

    def show_help(self):
        help_text = (
            "考勤系统 v1.0\n\n"
            "使用方法:\n"
            "1. 通过'文件'菜单导入名单(CSV文件格式)。\n"
            "2. 使用输入框输入拼音首字母进行搜索，支持多音字。\n"
            "3. 勾选'显示已标记的人员'可以查看所有人员，包括已标记的。\n"
            "4. 可以通过'导出考勤情况'按钮将考勤记录保存为CSV文件。\n\n"
            "快捷键使用方法:\n"
            "- 在搜索框中按下回车键可以切换唯一匹配人员的出勤状态。\n"
            "- 按数字键1-9可以快速切换对应位置人员的出勤状态。\n"
        )
        QMessageBox.information(self, "帮助", help_text)

    def _generate_initials_map(self, names):
        initials_map = defaultdict(set)  # 使用 set 自动去重
        for name in names:
            # 通过获取每个字符拼音的首字母生成首字母组合，包括多音字处理
            initials_list = pinyin(name, style=Style.FIRST_LETTER, heteronym=True)  # heteronym=True 表示支持多音字
            all_combinations = set()

            # 使用递归生成所有组合
            def generate_combinations(prefix, idx):
                if idx == len(initials_list):
                    all_combinations.add(prefix)
                    return
                for letter in initials_list[idx]:
                    generate_combinations(prefix + letter[0], idx + 1)

            generate_combinations("", 0)

            for initials_key in all_combinations:
                initials_map[initials_key].add(name)  # 使用 set 添加

        # 转换为 list 便于后续使用
        return {k: list(v) for k, v in initials_map.items()}

    def display_all_names(self):
        """
        在表格中显示所有姓名。
        """
        self.table.setRowCount(0)  # 清空表格
        self.matched_names = []  # 重置匹配的姓名列表
        for name in self.names.keys():
            if not self.show_marked and self.names[name]:
                continue  # 跳过已标记的人员
            self.matched_names.append(name)  # 仅存储当前可见的匹配项
            row_position = self.table.rowCount()
            self.table.insertRow(row_position)

            # 添加姓名
            item = QTableWidgetItem(name)
            item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row_position, 0, item)

            # 添加复选框
            checkbox = QCheckBox()
            checkbox.setChecked(self.names[name])
            checkbox.stateChanged.connect(lambda state, n=name: self.mark_attendance(state, n))
            checkbox_widget = QWidget()
            checkbox_layout = QVBoxLayout(checkbox_widget)
            checkbox_layout.addWidget(checkbox)
            checkbox_layout.setAlignment(Qt.AlignCenter)
            checkbox_layout.setContentsMargins(0, 0, 0, 0)
            self.table.setCellWidget(row_position, 1, checkbox_widget)
        self.update_search_count()  # 更新搜索结果数量

    def on_text_change(self):
        initials = self.input_field.text().strip().lower()
        if not initials:
            # 如果输入为空，显示所有姓名
            self.display_all_names()
        else:
            # 根据拼音首字母组合过滤并显示匹配的姓名
            self.matched_names = []
            self.table.setRowCount(0)  # 清空表格
            seen_names = set()  # 记录已经显示的姓名，防止重复
            for initials_key, names in self.initials_map.items():
                if initials_key.startswith(initials):  # 修改为检查首字母组合是否以输入内容开头
                    for name in names:
                        if name in seen_names:
                            continue  # 跳过已显示的姓名，防止重复
                        if not self.show_marked and self.names[name]:
                            continue  # 跳过已标记的人员
                        self.matched_names.append(name)
                        seen_names.add(name)  # 添加到已显示的姓名集合
                        row_position = self.table.rowCount()
                        self.table.insertRow(row_position)

                        # 添加姓名并设置居中对齐
                        item = QTableWidgetItem(name)
                        item.setTextAlignment(Qt.AlignCenter)  # 设置居中对齐
                        self.table.setItem(row_position, 0, item)

                        # 添加复选框
                        checkbox = QCheckBox()
                        checkbox.setChecked(self.names[name])
                        checkbox.stateChanged.connect(lambda state, n=name: self.mark_attendance(state, n))
                        checkbox_widget = QWidget()
                        checkbox_layout = QVBoxLayout(checkbox_widget)
                        checkbox_layout.addWidget(checkbox)
                        checkbox_layout.setAlignment(Qt.AlignCenter)
                        checkbox_layout.setContentsMargins(0, 0, 0, 0)
                        self.table.setCellWidget(row_position, 1, checkbox_widget)
        self.update_search_count()  # 更新搜索结果数量

    def on_show_marked_changed(self, state):
        self.show_marked = state == Qt.Checked
        self.display_all_names()

    def eventFilter(self, source, event):
        if source == self.input_field and event.type() == QKeyEvent.KeyPress:
            key = event.text()
            if key.isdigit() and 1 <= int(key) <= len(self.matched_names):
                # 如果按下的是数字键1-9，并且有足够的匹配项
                index = int(key) - 1
                name = self.matched_names[index]
                self.names[name] = not self.names[name]  # 切换标记状态
                print(f"{name} 已标记为 {'出勤' if self.names[name] else '缺勤'}")
                self.display_all_names()  # 更新表格
                self.input_field.clear()  # 清空输入框
                self.update_stats()  # 更新统计信息
                return True
            elif event.key() == Qt.Key_Return or event.key() == Qt.Key_Enter:
                # 处理回车键
                if len(self.matched_names) == 1:
                    name = self.matched_names[0]
                    self.names[name] = not self.names[name]  # 切换标记状态
                    print(f"{name} 已标记为 {'出勤' if self.names[name] else '缺勤'}")
                    self.display_all_names()  # 更新表格
                    self.input_field.clear()  # 清空输入框
                    self.update_stats()  # 更新统计信息
                    return True
        return super().eventFilter(source, event)

    def mark_attendance(self, state, name):
        self.names[name] = state == Qt.Checked
        print(f"{name} 已标记为 {'出勤' if self.names[name] else '缺勤'}")
        self.update_stats()  # 更新统计信息

    def update_stats(self):
        total = len(self.names)
        present = sum(1 for status in self.names.values() if status)
        absent = total - present
        self.stats_label.setText(f"总人数: {total} | 出勤: {present} | 缺勤: {absent}")

    def update_search_count(self):
        count = len(self.matched_names)
        self.search_count_label.setText(f"当前搜索结果数量: {count}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = AttendanceApp()
    window.show()
    sys.exit(app.exec_())
