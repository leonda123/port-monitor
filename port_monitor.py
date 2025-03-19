import sys
import psutil
import subprocess
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTableWidget, QTableWidgetItem, 
                             QPushButton, QVBoxLayout, QHBoxLayout, QWidget, QHeaderView, 
                             QMessageBox, QLabel, QDialog, QTextEdit, QCheckBox, QGroupBox,
                             QGridLayout, QComboBox, QLineEdit, QAction, QMenu, QMenuBar)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QIcon, QFont, QPalette, QColor

class AboutDialog(QDialog):
    """关于对话框"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("关于 Windows端口占用监控工具")
        self.resize(500, 300)
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        
        # 标题
        title_label = QLabel("Windows端口占用监控工具")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)
        
        # 版本信息
        version_label = QLabel("版本: 1.0.0")
        version_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(version_label)
        
        # 描述
        desc_text = QTextEdit()
        desc_text.setReadOnly(True)
        desc_text.setHtml("""
        <p>这是一个用于监控Windows系统端口占用情况的工具，主要功能包括：</p>
        <ul>
            <li>实时监控系统端口占用情况</li>
            <li>查看占用端口的进程详细信息</li>
            <li>支持按端口号和进程名过滤</li>
            <li>支持关闭占用端口的进程</li>
            <li>支持深色/浅色主题切换</li>
        </ul>
        <p>项目地址: <a href='https://github.com/leonda123/port-monitor'>https://github.com/leonda123/port-monitor</a></p>
        <p>作者: leonda</p>
        <p>联系方式: dadajiu45@gmail.com</p>
        """)
        layout.addWidget(desc_text)
        
        # 关闭按钮
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)
        
        self.setLayout(layout)

class ProcessDetailDialog(QDialog):
    """进程详细信息对话框"""
    def __init__(self, pid, parent=None):
        super().__init__(parent)
        self.pid = pid
        self.setWindowTitle(f"进程详细信息 (PID: {pid})")
        self.resize(600, 400)
        self.init_ui()
        self.load_process_info()
        
    def init_ui(self):
        layout = QVBoxLayout()
        
        # 创建信息显示区域
        self.info_text = QTextEdit()
        self.info_text.setReadOnly(True)
        layout.addWidget(self.info_text)
        
        # 关闭按钮
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)
        
        self.setLayout(layout)
    
    def load_process_info(self):
        try:
            process = psutil.Process(self.pid)
            info = []
            
            # 基本信息
            info.append(f"进程名称: {process.name()}")
            info.append(f"PID: {process.pid}")
            info.append(f"状态: {process.status()}")
            
            # 内存信息
            mem_info = process.memory_info()
            info.append(f"内存使用: {self.format_bytes(mem_info.rss)}")
            info.append(f"虚拟内存: {self.format_bytes(mem_info.vms)}")
            
            # CPU信息
            info.append(f"CPU使用率: {process.cpu_percent()}%")
            
            # 创建时间
            info.append(f"创建时间: {self.format_time(process.create_time())}")
            
            # 用户信息
            try:
                info.append(f"用户: {process.username()}")
            except:
                info.append("用户: 未知")
            
            # 命令行
            try:
                cmdline = process.cmdline()
                info.append(f"命令行: {' '.join(cmdline)}")
            except:
                info.append("命令行: 无法获取")
            
            # 打开的文件
            try:
                files = process.open_files()
                if files:
                    info.append("\n打开的文件:")
                    for file in files[:20]:  # 限制显示数量
                        info.append(f"  {file.path}")
                    if len(files) > 20:
                        info.append(f"  ... 还有 {len(files) - 20} 个文件未显示")
            except:
                info.append("打开的文件: 无法获取")
            
            # 网络连接
            try:
                connections = process.connections()
                if connections:
                    info.append("\n网络连接:")
                    for conn in connections:
                        local_addr = f"{conn.laddr.ip}:{conn.laddr.port}" if conn.laddr else "N/A"
                        remote_addr = f"{conn.raddr.ip}:{conn.raddr.port}" if conn.raddr else "N/A"
                        info.append(f"  {conn.type} - 本地: {local_addr}, 远程: {remote_addr}, 状态: {conn.status}")
            except:
                info.append("网络连接: 无法获取")
            
            self.info_text.setText("\n".join(info))
            
        except psutil.NoSuchProcess:
            self.info_text.setText("进程不存在或已终止")
        except Exception as e:
            self.info_text.setText(f"获取进程信息时出错: {str(e)}")
    
    def format_bytes(self, bytes):
        """格式化字节大小为人类可读格式"""
        for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
            if bytes < 1024:
                return f"{bytes:.2f} {unit}"
            bytes /= 1024
        return f"{bytes:.2f} PB"
    
    def format_time(self, timestamp):
        """格式化时间戳"""
        import datetime
        return datetime.datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S')

class PortMonitor(QMainWindow):
    """端口监控主窗口"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Windows端口占用监控工具")
        self.resize(900, 600)
        self.dark_mode = False  # 默认使用浅色主题
        self.init_ui()
        self.create_menu()
        self.port_data = []
        self.timer = QTimer()
        self.timer.timeout.connect(self.refresh_data)
        self.timer.start(5000)  # 每5秒刷新一次
        self.refresh_data()  # 初始加载数据
        
    def create_menu(self):
        """创建菜单栏"""
        menubar = self.menuBar()
        
        # 文件菜单
        file_menu = menubar.addMenu("文件")
        
        exit_action = QAction("退出", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # 视图菜单
        view_menu = menubar.addMenu("视图")
        
        theme_action = QAction("切换深色/浅色主题", self)
        theme_action.triggered.connect(self.toggle_theme)
        view_menu.addAction(theme_action)
        
        # 帮助菜单
        help_menu = menubar.addMenu("帮助")
        
        about_action = QAction("关于", self)
        about_action.triggered.connect(self.show_about_dialog)
        help_menu.addAction(about_action)
    
    def toggle_theme(self):
        """切换深色/浅色主题"""
        self.dark_mode = not self.dark_mode
        self.apply_theme()
    
    def apply_theme(self):
        """应用主题"""
        app = QApplication.instance()
        palette = QPalette()
        
        if self.dark_mode:
            # 深色主题
            palette.setColor(QPalette.Window, QColor(53, 53, 53))
            palette.setColor(QPalette.WindowText, Qt.white)
            palette.setColor(QPalette.Base, QColor(25, 25, 25))
            palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
            palette.setColor(QPalette.ToolTipBase, Qt.white)
            palette.setColor(QPalette.ToolTipText, Qt.white)
            palette.setColor(QPalette.Text, Qt.white)
            palette.setColor(QPalette.Button, QColor(53, 53, 53))
            palette.setColor(QPalette.ButtonText, Qt.white)
            palette.setColor(QPalette.BrightText, Qt.red)
            palette.setColor(QPalette.Link, QColor(42, 130, 218))
            palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
            palette.setColor(QPalette.HighlightedText, Qt.black)
            
            # 为表格添加特定的深色主题样式
            self.table.setStyleSheet("""
                QTableWidget {
                    background-color: #2D2D2D;
                    color: white;
                    gridline-color: #3A3A3A;
                    border: 1px solid #3A3A3A;
                }
                QTableWidget::item {
                    background-color: #2D2D2D;
                    color: white;
                }
                QTableWidget::item:selected {
                    background-color: #3A6EA5;
                    color: white;
                }
                QHeaderView::section {
                    background-color: #3A3A3A;
                    color: white;
                    padding: 4px;
                    border: 1px solid #505050;
                }
                QTableCornerButton::section {
                    background-color: #3A3A3A;
                    border: 1px solid #505050;
                }
            """)
            
            # 为按钮添加特定的深色主题样式
            button_style = """
                QPushButton {
                    background-color: #3A3A3A;
                    color: white;
                    border: 1px solid #505050;
                    border-radius: 3px;
                    padding: 5px 10px;
                }
                QPushButton:hover {
                    background-color: #505050;
                }
                QPushButton:pressed {
                    background-color: #3A6EA5;
                }
                QPushButton:disabled {
                    background-color: #2D2D2D;
                    color: #808080;
                    border: 1px solid #404040;
                }
            """
            self.refresh_btn.setStyleSheet(button_style)
            self.view_details_btn.setStyleSheet(button_style)
            self.kill_process_btn.setStyleSheet(button_style)
            self.force_kill_btn.setStyleSheet(button_style)
            self.filter_btn.setStyleSheet(button_style)
            
            # 为输入框添加特定的深色主题样式
            lineedit_style = """
                QLineEdit {
                    background-color: #2D2D2D;
                    color: white;
                    border: 1px solid #505050;
                    border-radius: 3px;
                    padding: 3px;
                }
                QLineEdit:focus {
                    border: 1px solid #3A6EA5;
                }
                QLineEdit::placeholder {
                    color: #A0A0A0;
                }
            """
            self.filter_port.setStyleSheet(lineedit_style)
            self.filter_process.setStyleSheet(lineedit_style)
            
            # 为分组框添加特定的深色主题样式
            groupbox_style = """
                QGroupBox {
                    border: 1px solid #505050;
                    border-radius: 5px;
                    margin-top: 10px;
                    font-weight: bold;
                    color: white;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    subcontrol-position: top center;
                    padding: 0 5px;
                    background-color: #3A3A3A;
                }
                QLabel {
                    color: white;
                }
            """
            filter_group = self.findChild(QGroupBox, "")
            action_group = self.findChild(QGroupBox, "")
            
            # 应用样式到所有QGroupBox
            for group_box in self.findChildren(QGroupBox):
                group_box.setStyleSheet(groupbox_style)
            
            # 为菜单栏添加特定的深色主题样式
            self.menuBar().setStyleSheet("""
                QMenuBar {
                    background-color: #3A3A3A;
                    color: white;
                    border-bottom: 1px solid #505050;
                }
                QMenuBar::item {
                    background-color: transparent;
                    padding: 4px 10px;
                }
                QMenuBar::item:selected {
                    background-color: #505050;
                    color: white;
                }
                QMenuBar::item:pressed {
                    background-color: #3A6EA5;
                    color: white;
                }
                QMenu {
                    background-color: #3A3A3A;
                    color: white;
                    border: 1px solid #505050;
                }
                QMenu::item {
                    padding: 5px 30px 5px 20px;
                }
                QMenu::item:selected {
                    background-color: #3A6EA5;
                    color: white;
                }
            """)
        else:
            # 浅色主题（默认）
            palette = app.style().standardPalette()
            # 清除所有组件的样式表
            self.table.setStyleSheet("")
            self.menuBar().setStyleSheet("")
            self.refresh_btn.setStyleSheet("")
            self.view_details_btn.setStyleSheet("")
            self.kill_process_btn.setStyleSheet("")
            self.force_kill_btn.setStyleSheet("")
            self.filter_btn.setStyleSheet("")
            self.filter_port.setStyleSheet("")
            self.filter_process.setStyleSheet("")
            
            # 清除所有QGroupBox的样式
            for group_box in self.findChildren(QGroupBox):
                group_box.setStyleSheet("")
        
        app.setPalette(palette)
    
    def show_about_dialog(self):
        """显示关于对话框"""
        dialog = AboutDialog(self)
        dialog.exec_()
    
    def init_ui(self):
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 创建顶部控制区域
        control_layout = QHBoxLayout()
        
        # 过滤选项
        filter_group = QGroupBox("过滤选项")
        filter_layout = QGridLayout()
        
        self.filter_port = QLineEdit()
        self.filter_port.setPlaceholderText("按端口号过滤")
        filter_layout.addWidget(QLabel("端口:"), 0, 0)
        filter_layout.addWidget(self.filter_port, 0, 1)
        
        self.filter_process = QLineEdit()
        self.filter_process.setPlaceholderText("按进程名过滤")
        filter_layout.addWidget(QLabel("进程:"), 1, 0)
        filter_layout.addWidget(self.filter_process, 1, 1)
        
        self.filter_btn = QPushButton("应用过滤")
        self.filter_btn.clicked.connect(self.apply_filter)
        filter_layout.addWidget(self.filter_btn, 2, 0, 1, 2)
        
        filter_group.setLayout(filter_layout)
        control_layout.addWidget(filter_group)
        
        # 操作按钮
        action_group = QGroupBox("操作")
        action_layout = QVBoxLayout()
        
        self.refresh_btn = QPushButton("刷新数据")
        self.refresh_btn.clicked.connect(self.refresh_data)
        action_layout.addWidget(self.refresh_btn)
        
        self.view_details_btn = QPushButton("查看进程详情")
        self.view_details_btn.clicked.connect(self.view_process_details)
        action_layout.addWidget(self.view_details_btn)
        
        self.kill_process_btn = QPushButton("关闭进程")
        self.kill_process_btn.clicked.connect(lambda: self.kill_process(False))
        action_layout.addWidget(self.kill_process_btn)
        
        self.force_kill_btn = QPushButton("强制关闭进程")
        self.force_kill_btn.clicked.connect(lambda: self.kill_process(True))
        action_layout.addWidget(self.force_kill_btn)
        
        action_group.setLayout(action_layout)
        control_layout.addWidget(action_group)
        
        main_layout.addLayout(control_layout)
        
        # 创建表格
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["进程ID", "进程名", "本地地址", "远程地址", "状态"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        main_layout.addWidget(self.table)
        
        # 状态栏
        self.statusBar().showMessage("就绪")
        
    def refresh_data(self):
        """刷新端口占用数据"""
        try:
            self.statusBar().showMessage("正在刷新数据...")
            # 获取所有网络连接
            connections = psutil.net_connections(kind='inet')
            self.port_data = []
            
            for conn in connections:
                if conn.laddr:  # 确保有本地地址
                    pid = conn.pid
                    if pid is None:
                        continue
                    
                    try:
                        process = psutil.Process(pid)
                        process_name = process.name()
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        process_name = "未知进程"
                    
                    local_address = f"{conn.laddr.ip}:{conn.laddr.port}"
                    remote_address = "N/A"
                    if conn.raddr:
                        remote_address = f"{conn.raddr.ip}:{conn.raddr.port}"
                    
                    status = conn.status if conn.status else "未知"
                    
                    self.port_data.append({
                        "pid": pid,
                        "name": process_name,
                        "local_address": local_address,
                        "remote_address": remote_address,
                        "status": status
                    })
            
            self.apply_filter()
            self.statusBar().showMessage(f"数据刷新完成，共 {len(self.port_data)} 条记录")
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"刷新数据时出错: {str(e)}")
            self.statusBar().showMessage("刷新数据失败")
    
    def apply_filter(self):
        """应用过滤条件并更新表格"""
        port_filter = self.filter_port.text().strip()
        process_filter = self.filter_process.text().strip().lower()
        
        filtered_data = self.port_data
        
        # 应用端口过滤
        if port_filter:
            filtered_data = [item for item in filtered_data 
                            if port_filter in item["local_address"].split(":")[1]]
        
        # 应用进程名过滤
        if process_filter:
            filtered_data = [item for item in filtered_data 
                            if process_filter in item["name"].lower()]
        
        # 更新表格
        self.table.setRowCount(0)  # 清空表格
        for row, item in enumerate(filtered_data):
            self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(str(item["pid"])))
            self.table.setItem(row, 1, QTableWidgetItem(item["name"]))
            self.table.setItem(row, 2, QTableWidgetItem(item["local_address"]))
            self.table.setItem(row, 3, QTableWidgetItem(item["remote_address"]))
            self.table.setItem(row, 4, QTableWidgetItem(item["status"]))
        
        self.statusBar().showMessage(f"显示 {self.table.rowCount()} 条记录")
    
    def get_selected_pid(self):
        """获取当前选中行的进程ID"""
        selected_rows = self.table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "警告", "请先选择一个进程")
            return None
        
        row = selected_rows[0].row()
        pid_item = self.table.item(row, 0)
        return int(pid_item.text())
    
    def view_process_details(self):
        """查看进程详细信息"""
        pid = self.get_selected_pid()
        if pid is not None:
            dialog = ProcessDetailDialog(pid, self)
            dialog.exec_()
    
    def kill_process(self, force=False):
        """关闭进程"""
        pid = self.get_selected_pid()
        if pid is None:
            return
        
        try:
            process = psutil.Process(pid)
            process_name = process.name()
            
            msg = f"确定要{'强制' if force else ''}关闭进程 {process_name} (PID: {pid})？"
            reply = QMessageBox.question(self, "确认", msg, 
                                       QMessageBox.Yes | QMessageBox.No, 
                                       QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                if force:
                    process.kill()  # 强制终止
                else:
                    process.terminate()  # 正常终止
                
                QMessageBox.information(self, "成功", f"进程 {process_name} 已{'强制' if force else ''}关闭")
                # 短暂延迟后刷新数据
                QTimer.singleShot(1000, self.refresh_data)
        
        except psutil.NoSuchProcess:
            QMessageBox.warning(self, "警告", "进程不存在或已终止")
        except psutil.AccessDenied:
            QMessageBox.warning(self, "警告", "权限不足，无法关闭该进程")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"关闭进程时出错: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PortMonitor()
    window.apply_theme()  # 应用当前主题
    window.show()
    sys.exit(app.exec_())