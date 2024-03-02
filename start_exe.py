import sys
import os
import subprocess

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from datetime import datetime


class AddDialog(QDialog):
    def __init__(self, parent=None):
        super(AddDialog, self).__init__(parent)
        self.init_ui(parent)

    def init_ui(self, parent):

        '''布局'''
        self.resize(400, 100)
        qbox = QHBoxLayout()

        self.s_btn = QPushButton()
        self.s_btn.setText('保存')
        self.s_btn.clicked.connect(lambda :self.s_btn_click(parent))

        self.c_btn = QPushButton()
        self.c_btn.setText('取消')
        self.c_btn.clicked.connect(self.c_btn_click)

        qbox.addWidget(self.s_btn)
        qbox.addWidget(self.c_btn)

        # 表单
        fbox = QFormLayout()
        self.row_select = parent.table.selectionModel().selectedRows()
        data = parent.data_list[self.row_select[0].row()] if self.row_select else []

        self.loc_lab = QLabel()
        self.loc_lab.setText('软件位置：')
        self.loc_text = QLineEdit()
        self.loc_text.setPlaceholderText('请输入软件文件位置！')
        if data:
            self.loc_text.setText(data[2])

        self.file_lab = QLabel()
        self.file_lab.setText('文件位置：')
        self.file_text = QTextEdit()
        self.file_text.setPlaceholderText('请输入打开文件位置！')
        if data:
            self.file_text.setText(data[3])

        fbox.addRow(self.loc_lab,self.loc_text)
        fbox.addRow(self.file_lab, self.file_text)

        qvbox = QVBoxLayout()
        qvbox.addLayout(fbox)
        qvbox.addLayout(qbox)

        self.setLayout(qvbox)

    def s_btn_click(self, parent):
        loc_text_text =  self.strip_text(self.loc_text.text())
        if loc_text_text != '' :
            file_text_text =  self.strip_text(self.file_text.toPlainText())
            date_time = datetime.today().date()
            name = os.path.basename(loc_text_text)
            fliesion = os.path.splitext(name)[-1]
            data = [name, fliesion, loc_text_text, file_text_text, date_time]
            if self.row_select:
                row = self.row_select[0].row()
                parent.data_list[row] = data
            else:
                parent.data_list.append(data)
            parent.save_data()
            parent.query_data()
            self.close()

    def c_btn_click(self):
        self.close()

    def open_dialog(parent=None):
        dialog = AddDialog(parent)
        return dialog.exec()


    def strip_text(self, text):
        if text is not None:
            text = text.strip().strip("'").strip('"').strip("‘").strip('’').strip("“").strip('”')
        return text


class StartExeManage(QWidget):

    update_tooltip_signal = pyqtSignal(object)
    def __init__(self):
        super(StartExeManage, self).__init__()
        self.settings = QSettings("init.ini", QSettings.IniFormat)
        self.select_app_btn_text = "选择软件"
        self.select_file_btn_text = "选择文件"
        self.add_btn_text = "添加数据"
        self.edit_btn_text = "编辑数据"
        self.del_btn_text = "删除选中"
        self.start_btn_text = "一键启动"
        self.quit_btn_text = "退出工具"
        self.table_header = ['软件名称', '软件类型', '软件位置', '文件位置', '添加日期']
        self.header_height = 30
        self.header_height = 30
        self.COLUMN = len(self.table_header)
        self.ROW = 0
        self.data_list = self.settings.value("DATA") if self.settings.value("DATA") else []
        self.update_tooltip_signal.connect(self.update_tooltip_slot)
        self.init_ui()

    def init_ui(self):
        '''初始化全局设置'''

        self.setWindowTitle('一键启动工具')
        self.resize(1000, 400)

        # 操作按钮
        self.add_edit_btn = self.create_button(self.add_btn_text, self.add_edit_btn_click)
        self.select_btn = self.create_button(self.select_app_btn_text, self.select_btn_click, checkable=True)
        self.del_btn = self.create_button(self.del_btn_text, self.del_btn_click)
        self.start_btn = self.create_button(self.start_btn_text, self.start_btn_click)
        self.quit_btn = self.create_button(self.quit_btn_text, self.quit_btn_click)

        # 列表
        self.table = QTableWidget()
        self.table.setColumnCount(self.COLUMN)
        self.table.setRowCount(self.ROW)
        self.table.setHorizontalHeaderLabels(self.table_header)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)  # 列宽手动调整
        self.table.horizontalHeader().setMinimumHeight(self.header_height)
        self.table.verticalHeader().setVisible(False)
        self.table.setShowGrid(True)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.MultiSelection)
        self.set_column_width([None, 80, None, None, 80], self.table)
        self.table.itemSelectionChanged.connect(self.change_btn_text)
        self.table.setMouseTracking(True)
        self.table.installEventFilter(self)

        # 布局
        self.col_nums = 20
        buttons = [self.select_btn, self.add_edit_btn, self.del_btn, self.start_btn, self.quit_btn]
        widgets =[(b, 1, (self.col_nums-len(buttons)+i), 1, 1) for i, b in enumerate(buttons)]
        q_grid = self.create_grid(widgets)
        q_grid.addWidget(self.table, 0, 0, 1, self.col_nums)

        self.setLayout(q_grid)
        self.query_data()


    def change_btn_text(self):
        row_select = self.table.selectionModel().selectedRows()
        if row_select:
            self.add_edit_btn.setText(self.edit_btn_text)
            self.select_btn.setText(self.select_file_btn_text)
        else:
            self.add_edit_btn.setText(self.add_btn_text)
            self.select_btn.setText(self.select_app_btn_text)

    def set_column_width(self, widths, table):
        """设置列"""
        for i, width in enumerate(widths):
            if width is None:
                table.horizontalHeader().setSectionResizeMode(i, QHeaderView.Stretch)
            else:
                table.setColumnWidth(i, width)
            table.horizontalHeaderItem(i).setTextAlignment(Qt.AlignVCenter)
        return table

    def create_button(self, text, fun, checkable= None):
        """创建按钮"""
        button = QPushButton()
        button.setText(text)
        if checkable is not None:
            button.setCheckable(checkable)
        button.clicked.connect(fun)
        return button

    def create_grid(self, widgets, grid = None):
        """数据列表"""
        grid = grid if grid else QGridLayout()
        for w in widgets:
            grid.addWidget(*w)
        return grid

    def quit_btn_click(self):
        """退出窗口"""
        QApplication.instance().quit()

    def add_edit_btn_click(self):
        """添加或者修改数据"""
        AddDialog.open_dialog(self)

    def select_btn_click(self):
        '''选择添加数据'''
         # 设置文件扩展名过滤，同一个类型的不同格式用空格隔开
        dir_path = "C:"
        btn_text = self.select_btn.text()
        if btn_text == self.select_file_btn_text:
            row_select = self.table.selectionModel().selectedRows()
            row = row_select[0].row() if row_select else 0
            files = self.data_list[row][3].split("\n") if len(row_select) > 0 else []
            dir_path = os.path.dirname(files[0]) if len(files) > 0 else ""
        file_paths, file_type = QFileDialog.getOpenFileNames(self, "请选择{0}".format(btn_text), dir_path, "")
        date_time = datetime.today().date()
        if btn_text == self.select_app_btn_text:
            if self.select_btn.isChecked():
                for path in file_paths:
                    name = os.path.basename(path)
                    fliesion = os.path.splitext(path)[-1]
                    self.data_list.append([name, fliesion, path, '', date_time])
                self.settings.setValue("DATA", self.data_list)
                self.query_data()
        elif btn_text == self.select_file_btn_text:
            if self.select_btn.isChecked() and row >= 0:
                if file_paths:
                    file_paths = "\n".join(file_paths)
                    self.data_list[row][3] = file_paths
                    self.settings.setValue("DATA", self.data_list)
                    self.query_data()
        self.select_btn.toggle()

    def del_btn_click(self):
        '''删除数据'''
        row_select = self.table.selectionModel().selectedRows()
        for model in row_select:
            row = model.row()
            self.table.removeRow(row)
            if len(self.data_list) > row:
                self.data_list[row] = []
        self.data_list = [d for d in self.data_list if d]
        self.settings.setValue("DATA", self.data_list)
        self.query_data()

    def start_btn_click(self):
        '''启动exe '''
        row_select = self.table.selectionModel().selectedRows()
        if len(row_select) == 0:
            for i, _ in enumerate(self.data_list):
                self.table.selectRow(i)
            row_select = self.table.selectionModel().selectedRows()
        for model in row_select:
            row = model.row()
            if len(self.data_list) > row:
                software = self.data_list[row][2]
                files = self.data_list[row][3]
                for file in files.split("\n"):
                    cmd ="{0} {1}".format(software, file) if file else software
                    if cmd:
                        subprocess.Popen(cmd)
            else:
                self.query_data()

    def query_data(self):
        '''查询数据'''
        datas = self.data_list
        self.table.setRowCount(len(datas))
        if len(datas) != 0:
            for i, data in enumerate(datas):
                if len(data) == self.COLUMN:
                    for j, d in enumerate(data):
                        self.table.setItem(i, j, QTableWidgetItem(str(d)))

    def save_data(self):
        self.settings.setValue("DATA", self.data_list)

    def update_tooltip_slot(self,posit):
        """通过计算坐标确定当前位置所属单元格"""
        self.tool_tip = ""
        self.mouse_x, self.mouse_y = posit.x(), posit.y()
        for r in range(self.table.rowCount()):
            row_height = self.table.rowHeight(r)
            #累计列宽
            self.col_width = 0
            if row_height * r + self.header_height <= self.mouse_y <= ( r + 1) * row_height + self.header_height:
                for c in range(self.table.columnCount()):
                    current_col_width = self.table.columnWidth(c)
                    #每一列的列宽可能不同
                    if self.col_width <= self.mouse_x <= self.col_width + current_col_width:
                        item = self.table.item(r, c)
                        self.tool_tip = item.text() if item != None else ""
                        return self.tool_tip
                    else:
                        self.col_width =self.col_width + current_col_width


    def eventFilter(self, object, event):
        """事件过滤器(关键)"""
        if object is self.table:
            self.setCursor(Qt.ArrowCursor)
            if event.type() == QEvent.ToolTip:
                self.update_tooltip_signal.emit(event.pos())
                rect = QRect(self.mouse_x,self.mouse_y, 30, 10)
                QApplication.processEvents()
                QToolTip.showText(QCursor.pos(), self.tool_tip,self.table,rect,1500)
        return QWidget.eventFilter(self, object, event)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    manage = StartExeManage()
    manage.show()
    sys.exit(app.exec_())
