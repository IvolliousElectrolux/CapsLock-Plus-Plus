#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import os
import configparser
import re
import ctypes
import threading
import time
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                            QLabel, QLineEdit, QPushButton, QListWidget, 
                            QMessageBox, QFileDialog, QDialog, QComboBox,
                            QFormLayout, QDialogButtonBox, QCheckBox, QInputDialog,
                            QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView,
                            QGroupBox)
from PyQt5.QtCore import Qt, QPoint, QTimer, pyqtSignal, QObject
from PyQt5.QtGui import QIcon, QCursor

try:
    from pynput import mouse
    from pynput import keyboard
    PYNPUT_AVAILABLE = True
except ImportError:
    PYNPUT_AVAILABLE = False

# 使用ctypes导入Windows API
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

# 带圈数字 1-10
CIRCLED_NUMBERS = ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩"]
# 禁用符号 - 带圈数字零
DISABLED_SYMBOL = "⓪"  # 带圈数字零 (Unicode: U+24EA)

# 定义Windows API常量和结构体
class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

# 添加捕获信号类
class CaptureSignals(QObject):
    # 定义信号
    captureComplete = pyqtSignal(str, str)
    captureCanceled = pyqtSignal()

class MenuItemDialog(QDialog):
    def __init__(self, parent=None, item_data=None):
        super(MenuItemDialog, self).__init__(parent)
        self.setWindowTitle("编辑菜单项")
        self.resize(400, 300)
        self.item_data = item_data or {}
        
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        # 名称
        self.name_edit = QLineEdit(self.item_data.get("name", ""))
        form_layout.addRow("名称:", self.name_edit)
        
        # 图标
        icon_layout = QHBoxLayout()
        self.icon_edit = QLineEdit(self.item_data.get("icon", ""))
        self.browse_btn = QPushButton("浏览...")
        self.browse_btn.clicked.connect(self.browse_icon)
        icon_layout.addWidget(self.icon_edit)
        icon_layout.addWidget(self.browse_btn)
        form_layout.addRow("图标:", icon_layout)
        
        # 图标类型
        self.icon_type_combo = QComboBox()
        self.icon_type_combo.addItems(["emoji", "file"])
        self.icon_type_combo.setCurrentText(self.item_data.get("icontype", "emoji"))
        form_layout.addRow("图标类型:", self.icon_type_combo)
        
        # 功能/动作
        self.action_edit = QLineEdit(self.item_data.get("action", ""))
        form_layout.addRow("功能:", self.action_edit)
        
        # 常用功能列表
        self.action_examples = QListWidget()
        self.action_examples.addItems([
            "SetPowerPlan(\"节电\")",
            "SetPowerPlan(\"平衡\")",
            "SetPowerPlan(\"性能\")",
            "ManageProcessWithCtrlCheck(\"启用\")",
            "ManageProcessWithCtrlCheck(\"终止\")",
            "GitHubAccelerate()",
            "WebsiteLogin()",
            "ActivateOrRun(\"程序名\", \"路径\")"
        ])
        self.action_examples.currentTextChanged.connect(self.on_example_selected)
        form_layout.addRow("常用功能:", self.action_examples)
        
        layout.addLayout(form_layout)
        
        # 按钮
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
        
    def browse_icon(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "选择图标文件", "", "图标文件 (*.ico *.png *.jpg *.jpeg);;所有文件 (*.*)"
        )
        if file_name:
            self.icon_edit.setText(file_name)
            self.icon_type_combo.setCurrentText("file")
            
    def on_example_selected(self, text):
        if text:
            self.action_edit.setText(text)
            
    def get_data(self):
        return {
            "name": self.name_edit.text(),
            "icon": self.icon_edit.text(),
            "icontype": self.icon_type_combo.currentText(),
            "action": self.action_edit.text()
        }

class MenuConfigTool(QMainWindow):
    def __init__(self):
        super(MenuConfigTool, self).__init__()
        self.setWindowTitle("CapsLock++ 菜单配置工具")
        self.resize(800, 600)
        
        # 正确获取应用程序路径，兼容脚本和打包后的EXE
        if getattr(sys, 'frozen', False):
            # 如果是打包后的EXE（通过PyInstaller）
            application_path = os.path.dirname(sys.executable)
        else:
            # 如果是直接运行的脚本
            application_path = os.path.dirname(os.path.abspath(__file__))
            
        # 设置窗口图标
        icon_path = os.path.join(application_path, "Icon", "自定义.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
            
        # 使用计算出的路径来定位INI文件
        self.ini_file = os.path.join(application_path, "CapsLock++.ini")
        
        # 菜单组状态和数据
        self.group_names = [""] * 11  # 存储组名，索引0不使用
        self.group_enabled = [False] * 11  # 存储组启用状态，索引0不使用
        self.group_items = [[] for _ in range(11)]  # 存储组内项目，索引0不使用
        
        # 进程管理相关数据
        self.processes_to_start = []
        self.processes_to_terminate = []
        self.gui_processes_to_start = []
        self.gui_processes_to_terminate = []
        
        # 网站配置相关数据
        self.default_site = ""
        self.browser_preference = "default"
        self.websites = []  # 每项包含[site, url]
        
        # 黑名单配置
        self.blacklist_processes = []  # 进程黑名单
        self.blacklist_classes = []    # 窗口类名黑名单
        
        # 速记路径配置
        self.note_targets = []  # 速记目标列表，每项包含[keyword, path]
        
        # 暗色模式设置
        self.dark_mode = True
        
        # 为未命名的组设置默认名称
        for i in range(1, 11):
            if not self.group_names[i]:
                self.group_names[i] = f"菜单组 {i}"
        
        self.current_group_idx = 0  # 当前选中的组索引
        
        # 窗口捕获模式标志
        self.is_capturing = False
        self.capture_listener = None
        
        # 捕获信号
        self.capture_signals = CaptureSignals()
        self.capture_signals.captureComplete.connect(self.on_capture_complete)
        self.capture_signals.captureCanceled.connect(self.on_capture_canceled)
        
        self.initUI()
        self.load_config()
        
    def initUI(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout()
        
        # 创建选项卡控件
        self.tabs = QTabWidget()
        
        # 创建黑名单配置页面
        self.blacklist_page = QWidget()
        self.tabs.addTab(self.blacklist_page, "黑名单窗口")
        
        # 创建速记路径配置页面
        self.note_page = QWidget()
        self.tabs.addTab(self.note_page, "速记路径")
        
        # 创建菜单配置页面
        self.menu_page = QWidget()
        self.tabs.addTab(self.menu_page, "菜单配置")
        
        # 创建进程管理页面
        self.process_page = QWidget()
        self.tabs.addTab(self.process_page, "进程管理")
        
        # 创建网站配置页面
        self.website_page = QWidget()
        self.tabs.addTab(self.website_page, "网站配置")
        
        # 初始化黑名单配置页面
        self.init_blacklist_page()
        
        # 初始化速记路径配置页面
        self.init_note_page()
        
        # 初始化菜单配置页面
        self.init_menu_page()
        
        # 初始化进程管理页面
        self.init_process_page()
        
        # 初始化网站配置页面
        self.init_website_page()
        
        main_layout.addWidget(self.tabs)
        
        # 底部按钮
        bottom_layout = QHBoxLayout()
        self.reload_btn = QPushButton("重新加载")
        self.reload_btn.clicked.connect(self.load_config)
        self.save_btn = QPushButton("保存配置")
        self.save_btn.clicked.connect(self.save_config)
        
        bottom_layout.addStretch(1)
        bottom_layout.addWidget(self.reload_btn)
        bottom_layout.addWidget(self.save_btn)
        
        # 整体布局
        main_layout.addLayout(bottom_layout)
        
        central_widget.setLayout(main_layout)
        
    def init_blacklist_page(self):
        """初始化黑名单配置页面"""
        layout = QVBoxLayout()
        
        # 创建水平布局，用于左右分布两个黑名单表格
        lists_layout = QHBoxLayout()
        
        # 左侧 - 进程黑名单
        process_group = QGroupBox("进程黑名单")
        process_layout = QVBoxLayout()
        
        # 进程黑名单表格
        self.process_table = QTableWidget(0, 1)
        self.process_table.setHorizontalHeaderLabels(["进程路径"])
        self.process_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        # 禁用默认的选择行为，我们将自己管理
        self.process_table.setSelectionMode(QTableWidget.NoSelection)
        # 使用鼠标按下事件
        self.process_table.mousePressEvent = lambda event: self.table_mouse_press(self.process_table, event)
        process_layout.addWidget(self.process_table)
        
        process_group.setLayout(process_layout)
        lists_layout.addWidget(process_group)
        
        # 右侧 - 窗口类名黑名单
        class_group = QGroupBox("窗口类名黑名单")
        class_layout = QVBoxLayout()
        
        # 窗口类名黑名单表格
        self.class_table = QTableWidget(0, 1)
        self.class_table.setHorizontalHeaderLabels(["窗口类名"])
        self.class_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        # 禁用默认的选择行为，我们将自己管理
        self.class_table.setSelectionMode(QTableWidget.NoSelection)
        # 使用鼠标按下事件
        self.class_table.mousePressEvent = lambda event: self.table_mouse_press(self.class_table, event)
        class_layout.addWidget(self.class_table)
        
        class_group.setLayout(class_layout)
        lists_layout.addWidget(class_group)
        
        layout.addLayout(lists_layout)
        
        # 底部共享按钮
        button_layout = QHBoxLayout()
        
        # 添加选择窗口按钮 - 使用瞄准镜符号
        self.capture_btn = QPushButton("选择窗口")
        self.capture_btn.setToolTip("按住左键不放，拖动到目标窗口，在目标窗口上释放左键捕获窗口信息")
        # 使用pressed信号而不是clicked信号
        self.capture_btn.pressed.connect(self.capture_window)
        button_layout.addWidget(self.capture_btn)
        
        button_layout.addStretch(1)
        
        self.add_blacklist_btn = QPushButton("添加")
        self.add_blacklist_btn.clicked.connect(self.add_blacklist_item)
        
        self.edit_blacklist_btn = QPushButton("编辑")
        self.edit_blacklist_btn.clicked.connect(self.edit_blacklist_item)
        
        self.delete_blacklist_btn = QPushButton("删除")
        self.delete_blacklist_btn.clicked.connect(self.delete_blacklist_item)
        
        button_layout.addWidget(self.add_blacklist_btn)
        button_layout.addWidget(self.edit_blacklist_btn)
        button_layout.addWidget(self.delete_blacklist_btn)
        
        layout.addLayout(button_layout)
        
        self.blacklist_page.setLayout(layout)
        
        # 初始状态
        self.current_blacklist_table = None
        self.current_blacklist_row = -1  # 当前选中的行
        self.update_blacklist_buttons_state()
    
    def capture_window(self):
        """开始窗口捕获模式"""
        if not PYNPUT_AVAILABLE:
            QMessageBox.warning(self, "功能不可用", "缺少pynput库，请安装后再使用窗口捕获功能。\n可以通过pip install pynput安装。")
            return
            
        # 设置为捕获模式
        self.is_capturing = True
        # 显示状态栏提示
        self.statusBar().showMessage("按住左键拖动到目标窗口，松开左键进行捕获...(按ESC键取消)")
        # 设置光标
        QApplication.setOverrideCursor(Qt.CrossCursor)
        
        # 启动捕获线程
        self.start_capture_listener()
    
    def start_capture_listener(self):
        """启动鼠标监听器进行捕获"""
        def on_click(x, y, button, pressed):
            # 只跟踪左键
            if button == mouse.Button.left:
                if not pressed:  # 左键释放时
                    # 获取当前位置的窗口信息并停止监听
                    return False
            return True  # 继续监听
        
        def on_key_press(key):
            # ESC键取消捕获
            if key == keyboard.Key.esc:
                # 停止监听
                return False
            return True
        
        def capture_thread():
            try:
                # 创建监听器
                with mouse.Listener(on_click=on_click) as mouse_listener, \
                     keyboard.Listener(on_press=on_key_press) as key_listener:
                    self.capture_mouse_listener = mouse_listener
                    self.capture_key_listener = key_listener
                    listener_thread = threading.Thread(target=key_listener.join)
                    listener_thread.daemon = True
                    listener_thread.start()
                    mouse_listener.join()
                
                # 判断是否是按ESC取消的
                if key_listener.is_alive():
                    key_listener.stop()
                    # 监听器结束，获取当前鼠标位置信息
                    if self.is_capturing:  # 确保仍然在捕获模式
                        # 获取当前鼠标位置
                        cursor_pos = QCursor.pos()
                        point = POINT(cursor_pos.x(), cursor_pos.y())
                        hwnd = user32.WindowFromPoint(point)
                        
                        process_name = ""
                        class_name = ""
                        
                        if hwnd:
                            # 获取窗口类名
                            class_buffer = ctypes.create_unicode_buffer(256)
                            user32.GetClassNameW(hwnd, class_buffer, 256)
                            class_name = class_buffer.value
                            
                            # 获取进程ID
                            process_id = ctypes.c_ulong(0)
                            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(process_id))
                            
                            # 打开进程以获取可执行文件路径
                            hProcess = kernel32.OpenProcess(0x1000, False, process_id.value)  # PROCESS_QUERY_LIMITED_INFORMATION
                            
                            if hProcess:
                                # 获取进程路径
                                path_buffer = ctypes.create_unicode_buffer(260)  # MAX_PATH
                                path_len = ctypes.c_ulong(260)
                                if kernel32.QueryFullProcessImageNameW(hProcess, 0, path_buffer, ctypes.byref(path_len)):
                                    process_path = path_buffer.value
                                    # 提取进程名
                                    process_name = os.path.basename(process_path)
                                else:
                                    process_name = f"进程ID: {process_id.value}"
                                
                                # 关闭进程句柄
                                kernel32.CloseHandle(hProcess)
                            else:
                                process_name = f"进程ID: {process_id.value}"
                        
                        # 发送捕获完成信号（以线程安全的方式）
                        self.capture_signals.captureComplete.emit(process_name, class_name)
                else:
                    # ESC键取消捕获
                    self.capture_signals.captureCanceled.emit()
            except Exception as e:
                print(f"捕获错误: {str(e)}")
                # 发送取消信号
                self.capture_signals.captureCanceled.emit()
        
        # 创建并启动线程
        capture_thread = threading.Thread(target=capture_thread)
        capture_thread.daemon = True
        capture_thread.start()
    
    def on_capture_complete(self, process_name, class_name):
        """捕获完成的槽函数"""
        # 退出捕获模式
        self.is_capturing = False
        QApplication.restoreOverrideCursor()
        
        added = []
        
        # 添加进程信息
        if process_name:
            row = self.process_table.rowCount()
            self.process_table.insertRow(row)
            self.process_table.setItem(row, 0, QTableWidgetItem(process_name))
            # 设置为当前选中项
            self.switch_blacklist_table(self.process_table)
            self.select_blacklist_row(row)
            added.append(f"进程: {process_name}")
        
        # 添加类名信息
        if class_name:
            row = self.class_table.rowCount()
            self.class_table.insertRow(row)
            self.class_table.setItem(row, 0, QTableWidgetItem(class_name))
            # 设置为当前选中项
            self.switch_blacklist_table(self.class_table)
            self.select_blacklist_row(row)
            added.append(f"类名: {class_name}")
        
        # 更新状态栏显示捕获的信息，不再弹出对话框
        if added:
            status_msg = "已添加 " + " 和 ".join(added)
            self.statusBar().showMessage(status_msg, 5000)  # 显示5秒
        else:
            self.statusBar().showMessage("未能获取窗口信息", 3000)
    
    def on_capture_canceled(self):
        """捕获取消的槽函数"""
        self.is_capturing = False
        QApplication.restoreOverrideCursor()
        self.statusBar().showMessage("已取消窗口捕获", 3000)
    
    def keyPressEvent(self, event):
        """键盘按下事件"""
        # 如果在捕获模式下按下Esc键，取消捕获
        if self.is_capturing and event.key() == Qt.Key_Escape:
            self.is_capturing = False
            # 停止监听器
            if self.capture_listener:
                self.capture_listener.stop()
                self.capture_listener = None
            
            QApplication.restoreOverrideCursor()
            self.statusBar().clearMessage()
            QMessageBox.information(self, "捕获取消", "已取消窗口捕获")
            event.accept()
            return
        
        # 调用父类方法
        super(MenuConfigTool, self).keyPressEvent(event)
    
    def table_mouse_press(self, table, event):
        """处理表格的鼠标按下事件"""
        # 保存原始事件处理函数
        orig_mousePressEvent = table.__class__.mousePressEvent
        # 调用原始事件处理（获取点击位置等信息）
        orig_mousePressEvent(table, event)
        
        # 设置当前活动表格
        self.switch_blacklist_table(table)
        
        # 如果点击的是一个单元格，则选择它
        item = table.itemAt(event.pos())
        if item:
            row = item.row()
            # 高亮选中行
            self.select_blacklist_row(row)
        
    def switch_blacklist_table(self, table):
        """切换当前活动的黑名单表格"""
        # 清除两个表格的高亮选择
        self.clear_blacklist_selections()
        
        # 设置当前表格
        self.current_blacklist_table = table
        self.current_blacklist_row = -1  # 重置行选择
        
        # 更新按钮状态
        self.update_blacklist_buttons_state()
    
    def select_blacklist_row(self, row):
        """选择黑名单表格的行"""
        if self.current_blacklist_table is None or row < 0 or row >= self.current_blacklist_table.rowCount():
            return
            
        # 清除之前的选择
        self.clear_blacklist_selections()
        
        # 设置行的背景色
        for col in range(self.current_blacklist_table.columnCount()):
            item = self.current_blacklist_table.item(row, col)
            if item:
                item.setSelected(True)
                
        # 更新当前选中行
        self.current_blacklist_row = row
        
        # 更新按钮状态
        self.update_blacklist_buttons_state()
    
    def clear_blacklist_selections(self):
        """清除黑名单表格的所有选择"""
        # 清除进程表格选择
        for row in range(self.process_table.rowCount()):
            for col in range(self.process_table.columnCount()):
                item = self.process_table.item(row, col)
                if item:
                    item.setSelected(False)
        
        # 清除类名表格选择
        for row in range(self.class_table.rowCount()):
            for col in range(self.class_table.columnCount()):
                item = self.class_table.item(row, col)
                if item:
                    item.setSelected(False)
    
    def update_blacklist_buttons_state(self):
        """更新黑名单按钮状态"""
        # 添加按钮总是可用的
        self.add_blacklist_btn.setEnabled(True)
        
        # 编辑和删除按钮只有在有选择并且选择了有效行时才可用
        has_selection = (self.current_blacklist_table is not None and 
                         self.current_blacklist_row >= 0 and 
                         self.current_blacklist_row < self.current_blacklist_table.rowCount())
        
        self.edit_blacklist_btn.setEnabled(has_selection)
        self.delete_blacklist_btn.setEnabled(has_selection)
    
    def add_blacklist_item(self):
        """添加黑名单项"""
        # 如果没有选中的表格，默认为进程表格
        if self.current_blacklist_table is None or self.current_blacklist_table == self.process_table:
            self.add_blacklist_process()
        else:
            self.add_blacklist_class()
    
    def edit_blacklist_item(self):
        """编辑黑名单项"""
        if not self.current_blacklist_table or self.current_blacklist_row < 0:
            return
            
        if self.current_blacklist_table == self.process_table:
            self.edit_blacklist_process()
        else:
            self.edit_blacklist_class()
    
    def delete_blacklist_item(self):
        """删除黑名单项"""
        if not self.current_blacklist_table or self.current_blacklist_row < 0:
            return
            
        self.current_blacklist_table.removeRow(self.current_blacklist_row)
        self.current_blacklist_row = -1  # 重置选中行
        self.update_blacklist_buttons_state()

    def add_blacklist_process(self):
        """添加进程黑名单项"""
        process_path, ok = QInputDialog.getText(self, "添加进程黑名单", "输入进程名或路径:")
        if ok and process_path:
            row = self.process_table.rowCount()
            self.process_table.insertRow(row)
            self.process_table.setItem(row, 0, QTableWidgetItem(process_path))
            
            # 切换到进程表格并选中新添加的行
            self.switch_blacklist_table(self.process_table)
            self.select_blacklist_row(row)
            
    def edit_blacklist_process(self):
        """编辑进程黑名单项"""
        if self.current_blacklist_row < 0:
            return
            
        current_path = self.process_table.item(self.current_blacklist_row, 0).text()
        process_path, ok = QInputDialog.getText(self, "编辑进程黑名单", "输入进程名或路径:", QLineEdit.Normal, current_path)
        if ok and process_path:
            self.process_table.setItem(self.current_blacklist_row, 0, QTableWidgetItem(process_path))
            
    def delete_blacklist_process(self):
        """删除进程黑名单项"""
        if self.current_blacklist_row >= 0:
            self.process_table.removeRow(self.current_blacklist_row)
            self.current_blacklist_row = -1
            self.update_blacklist_buttons_state()
            
    def add_blacklist_class(self):
        """添加窗口类名黑名单项"""
        class_name, ok = QInputDialog.getText(self, "添加窗口类名", "输入窗口类名:")
        if ok and class_name:
            row = self.class_table.rowCount()
            self.class_table.insertRow(row)
            self.class_table.setItem(row, 0, QTableWidgetItem(class_name))
            
            # 切换到类名表格并选中新添加的行
            self.switch_blacklist_table(self.class_table)
            self.select_blacklist_row(row)
            
    def edit_blacklist_class(self):
        """编辑窗口类名黑名单项"""
        if self.current_blacklist_row < 0:
            return
            
        current_class = self.class_table.item(self.current_blacklist_row, 0).text()
        class_name, ok = QInputDialog.getText(self, "编辑窗口类名", "输入窗口类名:", QLineEdit.Normal, current_class)
        if ok and class_name:
            self.class_table.setItem(self.current_blacklist_row, 0, QTableWidgetItem(class_name))
            
    def delete_blacklist_class(self):
        """删除窗口类名黑名单项"""
        if self.current_blacklist_row >= 0:
            self.class_table.removeRow(self.current_blacklist_row)
            self.current_blacklist_row = -1
            self.update_blacklist_buttons_state()
    
    def add_note_target(self):
        """添加速记目标"""
        # 创建添加对话框
        dialog = QDialog(self)
        dialog.setWindowTitle("添加速记目标")
        dialog.resize(400, 150)
        
        layout = QFormLayout()
        
        keyword_edit = QLineEdit()
        layout.addRow("关键词:", keyword_edit)
        
        path_edit = QLineEdit()
        path_layout = QHBoxLayout()
        path_layout.addWidget(path_edit)
        
        browse_btn = QPushButton("浏览...")
        browse_btn.clicked.connect(lambda: self.browse_note_path(path_edit))
        path_layout.addWidget(browse_btn)
        
        layout.addRow("文件路径:", path_layout)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        
        layout.addRow(buttons)
        dialog.setLayout(layout)
        
        if dialog.exec_() == QDialog.Accepted:
            keyword = keyword_edit.text()
            path = path_edit.text()
            
            if keyword and path:
                row = self.note_table.rowCount()
                self.note_table.insertRow(row)
                self.note_table.setItem(row, 0, QTableWidgetItem(keyword))
                self.note_table.setItem(row, 1, QTableWidgetItem(path))
                
    def edit_note_target(self):
        """编辑速记目标"""
        current_row = self.note_table.currentRow()
        if current_row < 0:
            return
            
        current_keyword = self.note_table.item(current_row, 0).text()
        current_path = self.note_table.item(current_row, 1).text()
        
        # 创建编辑对话框
        dialog = QDialog(self)
        dialog.setWindowTitle("编辑速记目标")
        dialog.resize(400, 150)
        
        layout = QFormLayout()
        
        keyword_edit = QLineEdit(current_keyword)
        layout.addRow("关键词:", keyword_edit)
        
        path_edit = QLineEdit(current_path)
        path_layout = QHBoxLayout()
        path_layout.addWidget(path_edit)
        
        browse_btn = QPushButton("浏览...")
        browse_btn.clicked.connect(lambda: self.browse_note_path(path_edit))
        path_layout.addWidget(browse_btn)
        
        layout.addRow("文件路径:", path_layout)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        
        layout.addRow(buttons)
        dialog.setLayout(layout)
        
        if dialog.exec_() == QDialog.Accepted:
            keyword = keyword_edit.text()
            path = path_edit.text()
            
            if keyword and path:
                self.note_table.setItem(current_row, 0, QTableWidgetItem(keyword))
                self.note_table.setItem(current_row, 1, QTableWidgetItem(path))
                
    def delete_note_target(self):
        """删除速记目标"""
        current_row = self.note_table.currentRow()
        if current_row >= 0:
            self.note_table.removeRow(current_row)
            
    def browse_note_path(self, edit):
        """浏览速记文件路径"""
        file_name, _ = QFileDialog.getOpenFileName(
            self, "选择文件", "", "文本文件 (*.txt);;所有文件 (*.*)"
        )
        if file_name:
            edit.setText(file_name)
            
    def move_note_target(self, direction):
        """上移或下移速记目标"""
        current_row = self.note_table.currentRow()
        if current_row < 0:
            return
            
        target_row = current_row + direction
        if target_row < 0 or target_row >= self.note_table.rowCount():
            return
            
        # 交换两行数据
        for col in range(2):
            current_item = self.note_table.takeItem(current_row, col)
            target_item = self.note_table.takeItem(target_row, col)
            
            self.note_table.setItem(target_row, col, current_item)
            self.note_table.setItem(current_row, col, target_item)
        
        # 选中移动后的行
        self.note_table.setCurrentCell(target_row, 0)

    def load_config(self):
        """从CapsLock++.ini文件加载配置"""
        try:
            # 读取INI文件内容
            ini_content = self.read_ini_file()
            if not ini_content:
                return
                
            # 加载黑名单配置
            self.load_blacklist_config(ini_content)
            
            # 加载速记目标配置
            self.load_note_config(ini_content)
                
            # 加载菜单配置
            self.load_menu_config(ini_content)
                
            # 加载进程管理配置
            self.load_process_config(ini_content)
            
            # 加载网站配置
            self.load_website_config(ini_content)
            
            # 加载暗色模式设置
            self.load_color_mode(ini_content)
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载配置失败: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def load_blacklist_config(self, ini_content):
        """加载黑名单配置"""
        # 重置黑名单数据
        self.blacklist_processes = []
        self.blacklist_classes = []
        
        # 解析进程黑名单
        blacklist_section = re.search(r'\[blacklist_virtual_env\](.*?)(?=\[blacklist_classes_virtual_env\]|\[noteTargets\]|\[MenuGroupNum\]|\[MenuGroupsEnable\]|\[MenuGroupName\]|\[MenuGroupCount\]|\[MenuGroupsColourMode\]|\[CommonWebsites\]|$)', ini_content, re.DOTALL)
        if blacklist_section:
            section_content = blacklist_section.group(1)
            i = 1
            while True:
                item_pattern = r'black{}\s*=\s*(.+)'.format(i)
                match = re.search(item_pattern, section_content)
                if not match:
                    break
                # 去除引号
                path = match.group(1).strip()
                if path.startswith('"') and path.endswith('"'):
                    path = path[1:-1]
                self.blacklist_processes.append(path)
                i += 1
        
        # 解析窗口类名黑名单 - 使用更精确的方式匹配节内容
        # 先提取[blacklist_classes_virtual_env]节内容
        classes_content = ""
        in_classes_section = False
        next_section_found = False
        for line in ini_content.splitlines():
            line = line.strip()
            
            # 检查是否找到类名黑名单节
            if line == "[blacklist_classes_virtual_env]":
                in_classes_section = True
                continue
                
            # 检查是否找到下一个节
            if in_classes_section and line and line.startswith("[") and line.endswith("]"):
                next_section_found = True
                break
                
            # 如果在类名黑名单节内，添加行内容
            if in_classes_section and not next_section_found:
                classes_content += line + "\n"
        
        # 从提取的内容中解析类名配置
        if classes_content:
            # 逐行解析，查找blackclassesX = "value"模式
            for line in classes_content.splitlines():
                line = line.strip()
                if not line or line.startswith(";"):  # 跳过空行和注释
                    continue
                    
                if "=" in line:
                    key, value = line.split("=", 1)
                    key = key.strip()
                    value = value.strip()
                    
                    if key.startswith('blackclasses') and value:
                        # 去除引号
                        if value.startswith('"') and value.endswith('"'):
                            value = value[1:-1]
                        self.blacklist_classes.append(value)
        
        # 更新UI
        self.update_blacklist_tables()
        
    def load_note_config(self, ini_content):
        """加载速记目标配置"""
        # 重置速记目标数据
        self.note_targets = []
        
        # 解析速记目标
        note_section = re.search(r'\[noteTargets\](.*?)(?=\[|$)', ini_content, re.DOTALL)
        if note_section:
            section_content = note_section.group(1)
            
            # 使用正则表达式查找所有配对的note<n>1和note<n>2
            keywords = {}
            paths = {}
            
            for line in section_content.splitlines():
                line = line.strip()
                if not line or line.startswith(';'):
                    continue
                    
                if '=' in line:
                    key, value = line.split('=', 1)
                    key = key.strip()
                    value = value.strip()
                    
                    # 去除引号
                    if value.startswith('"') and value.endswith('"'):
                        value = value[1:-1]
                    
                    if key.startswith('note') and len(key) >= 6:
                        # 提取索引和类型
                        match = re.match(r'note(\d+)(\d)', key)
                        if match:
                            index = match.group(1)
                            type_num = match.group(2)
                            
                            if type_num == '1':
                                keywords[index] = value
                            elif type_num == '2':
                                paths[index] = value
            
            # 将配对的keyword和path组合在一起
            for index in sorted(set(keywords.keys()) | set(paths.keys())):
                if index in keywords and index in paths:
                    self.note_targets.append([keywords[index], paths[index]])
        
        # 如果没有设置，添加默认项
        if not self.note_targets:
            self.note_targets = [
                ["论文", "速记\\论文灵感.txt"],
                ["日记", "速记\\日记.txt"],
                ["工作", "速记\\工作.txt"],
                ["想法", "速记\\想法.txt"]
            ]
        
        # 更新UI
        self.update_note_table()
        
    def update_blacklist_tables(self):
        """更新黑名单表格"""
        # 更新进程黑名单表格
        self.process_table.setRowCount(0)
        for i, path in enumerate(self.blacklist_processes):
            self.process_table.insertRow(i)
            self.process_table.setItem(i, 0, QTableWidgetItem(path))
        
        # 更新窗口类名黑名单表格
        self.class_table.setRowCount(0)
        for i, class_name in enumerate(self.blacklist_classes):
            self.class_table.insertRow(i)
            self.class_table.setItem(i, 0, QTableWidgetItem(class_name))
            
    def update_note_table(self):
        """更新速记目标表格"""
        self.note_table.setRowCount(0)
        for i, (keyword, path) in enumerate(self.note_targets):
            self.note_table.insertRow(i)
            self.note_table.setItem(i, 0, QTableWidgetItem(keyword))
            self.note_table.setItem(i, 1, QTableWidgetItem(path))
    
    def save_config(self):
        """保存配置到CapsLock++.ini文件"""
        try:
            # 读取原INI文件内容
            original_content = self.read_ini_file()
            if not original_content:
                return
            
            # 更新暗色模式设置
            if hasattr(self, 'dark_mode_cb'):
                self.dark_mode = self.dark_mode_cb.isChecked()
                
            # 收集黑名单数据
            self.collect_blacklist_data()
            
            # 收集速记目标数据
            self.collect_note_data()
            
            # 更新黑名单配置部分
            new_content = self.save_blacklist_config(original_content)
            
            # 更新速记目标配置部分
            new_content = self.save_note_config(new_content)
            
            # 更新菜单配置部分
            new_content = self.save_menu_config(new_content)
            
            # 更新进程管理配置部分
            new_content = self.save_process_config(new_content)
            
            # 更新网站配置部分
            new_content = self.save_website_config(new_content)
            
            # 更新暗色模式配置部分
            new_content = self.save_color_mode(new_content)
            
            # 保存修改后的内容
            with open(self.ini_file, 'w', encoding='utf-8') as f:
                f.write(new_content)
                
            # 尝试重启主脚本
            try:
                if os.path.exists("CapsLock++.ahk"):
                    import subprocess
                    subprocess.run(["autohotkey.exe", "CapsLock++.ahk", "/restart"], check=False)
                    QMessageBox.information(self, "成功", "配置已保存，并尝试重启主脚本。")
                else:
                    QMessageBox.information(self, "成功", "配置已保存，请手动重启CapsLock++.ahk以应用更改。")
            except Exception:
                QMessageBox.information(self, "成功", "配置已保存，请手动重启CapsLock++.ahk以应用更改。")
                
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存配置失败: {str(e)}")
            import traceback
            traceback.print_exc()  # 打印详细错误信息到控制台
    
    def collect_blacklist_data(self):
        """从表格中收集黑名单数据"""
        # 收集进程黑名单数据
        self.blacklist_processes = []
        for row in range(self.process_table.rowCount()):
            path = self.process_table.item(row, 0).text()
            self.blacklist_processes.append(path)
        
        # 收集窗口类名黑名单数据
        self.blacklist_classes = []
        for row in range(self.class_table.rowCount()):
            class_name = self.class_table.item(row, 0).text()
            self.blacklist_classes.append(class_name)
            
    def collect_note_data(self):
        """从表格中收集速记目标数据"""
        self.note_targets = []
        for row in range(self.note_table.rowCount()):
            keyword = self.note_table.item(row, 0).text()
            path = self.note_table.item(row, 1).text()
            self.note_targets.append([keyword, path])
    
    def save_blacklist_config(self, content):
        """保存黑名单配置部分"""
        # 构建进程黑名单节
        process_section = "[blacklist_virtual_env]\n"
        for i, path in enumerate(self.blacklist_processes, 1):
            # 不对路径进行额外转义，直接使用值
            process_section += f'black{i} = "{path}"\n'
        process_section += "\n"
        
        # 构建窗口类名黑名单节
        class_section = "[blacklist_classes_virtual_env]\n"
        for i, class_name in enumerate(self.blacklist_classes, 1):
            # 不对类名进行额外转义，直接使用值
            class_section += f'blackclasses{i} = "{class_name}"\n'
        class_section += "\n"
        
        # 创建内容的各节分割
        sections = []
        section_names = []
        current_section = ""
        current_section_name = ""
        in_section = False
        
        # 按行处理并识别各个节
        for line in content.splitlines():
            if line.strip() and line.strip().startswith('[') and line.strip().endswith(']'):
                # 保存上一个节
                if in_section:
                    sections.append(current_section)
                    section_names.append(current_section_name)
                
                # 开始新节
                current_section = line + "\n"
                current_section_name = line.strip()
                in_section = True
            elif in_section:
                current_section += line + "\n"
            else:
                # 文件开头的内容
                sections.append(line + "\n")
                section_names.append("")
        
        # 添加最后一个节
        if in_section:
            sections.append(current_section)
            section_names.append(current_section_name)
        
        # 重新构建内容，替换或添加节
        new_content = ""
        blacklist_processed = False
        classes_processed = False
        
        for i, section in enumerate(sections):
            section_name = section_names[i]
            
            if section_name == "[blacklist_virtual_env]":
                new_content += process_section
                blacklist_processed = True
            elif section_name == "[blacklist_classes_virtual_env]":
                new_content += class_section
                classes_processed = True
            else:
                new_content += section
        
        # 如果节不存在，添加到内容中
        if not blacklist_processed:
            new_content = process_section + new_content
        
        if not classes_processed:
            # 找到合适的位置添加类名黑名单节
            if blacklist_processed:
                # 在黑名单节之后添加
                parts = new_content.split(process_section)
                if len(parts) > 1:
                    new_content = parts[0] + process_section + class_section + parts[1]
                else:
                    new_content += class_section
            else:
                new_content = class_section + new_content
        
        return new_content
    
    def save_note_config(self, content):
        """保存速记目标配置部分"""
        # 构建速记目标节
        note_section = "[noteTargets]\n"
        for i, (keyword, path) in enumerate(self.note_targets, 1):
            note_section += f'note{i}1 = "{keyword}"\n'
            note_section += f'note{i}2 = "{path}"\n'
        note_section += "\n"
        
        # 替换速记目标节
        pattern = r'\[noteTargets\](.*?)(?=\[|$)'
        if re.search(pattern, content, re.DOTALL):
            new_content = re.sub(pattern, note_section, content, flags=re.DOTALL)
        else:
            new_content = content + note_section
        
        return new_content

    def load_menu_config(self, ini_content):
        """加载菜单配置部分"""
        # 重置菜单组数据
        self.group_names = [""] * 11
        self.group_enabled = [False] * 11
        self.group_items = [[] for _ in range(11)]
        
        # 读取组名
        name_section = re.search(r'\[MenuGroupName\](.*?)(?=\[|$)', ini_content, re.DOTALL)
        if name_section:
            for line in name_section.group(1).splitlines():
                line = line.strip()
                if line and "=" in line:
                    key, value = line.split("=", 1)
                    key = key.strip()
                    value = value.strip()
                    if key.startswith("name") and len(key) > 4:
                        try:
                            idx = int(key[4:])
                            if 1 <= idx <= 10:
                                self.group_names[idx] = value
                        except ValueError:
                            pass
        
        # 读取组启用状态
        enable_section = re.search(r'\[MenuGroupsEnable\](.*?)(?=\[|$)', ini_content, re.DOTALL)
        if enable_section:
            for line in enable_section.group(1).splitlines():
                line = line.strip()
                if line and "=" in line:
                    key, value = line.split("=", 1)
                    key = key.strip()
                    value = value.strip().lower()
                    if key.startswith("enableGroup") and len(key) > 11:
                        try:
                            idx = int(key[11:])
                            if 1 <= idx <= 10:
                                self.group_enabled[idx] = (value == "true")
                        except ValueError:
                            pass
        
        # 读取组项目数量
        counts = [0] * 11
        count_section = re.search(r'\[MenuGroupCount\](.*?)(?=\[|$)', ini_content, re.DOTALL)
        if count_section:
            for line in count_section.group(1).splitlines():
                line = line.strip()
                if line and "=" in line:
                    key, value = line.split("=", 1)
                    key = key.strip()
                    value = value.strip()
                    if key.startswith("count") and len(key) > 5:
                        try:
                            idx = int(key[5:])
                            counts[idx] = int(value)
                        except ValueError:
                            pass
        
        # 读取各组项目
        for group_idx in range(1, 11):
            if counts[group_idx] > 0:
                section_pattern = r'\[MenuGroups{}Items\](.*?)(?=\[|$)'.format(group_idx)
                section_match = re.search(section_pattern, ini_content, re.DOTALL)
                if section_match:
                    section_content = section_match.group(1)
                    items = []
                    
                    # 收集项目属性
                    item_props = {}
                    for line in section_content.splitlines():
                        line = line.strip()
                        if line and "=" in line:
                            key, value = line.split("=", 1)
                            key = key.strip()
                            value = value.strip()
                            
                            # 解析属性类型和索引
                            prop_match = re.match(r'([a-z]+)(\d+)', key)
                            if prop_match:
                                prop_type = prop_match.group(1)  # name, icon, icontype, action
                                item_idx = int(prop_match.group(2))
                                
                                if item_idx not in item_props:
                                    item_props[item_idx] = {}
                                item_props[item_idx][prop_type] = value
                    
                    # 按索引排序并添加项目
                    for idx in sorted(item_props.keys()):
                        if "name" in item_props[idx]:  # 只添加至少有名称的项目
                            items.append(item_props[idx])
                    
                    self.group_items[group_idx] = items
        
        # 为未命名的组设置默认名称
        for i in range(1, 11):
            if not self.group_names[i]:
                self.group_names[i] = f"菜单组 {i}"
        
        # 更新UI
        self.update_group_list()
    
    def load_process_config(self, ini_content):
        """加载进程管理配置部分"""
        # 重置进程管理数据
        self.processes_to_start = []
        self.processes_to_terminate = []
        self.gui_processes_to_start = []
        self.gui_processes_to_terminate = []
        
        # 解析 ProcessesToStart 节
        start_section = re.search(r'\[ProcessesToStart\](.*?)(?=\[|$)', ini_content, re.DOTALL)
        if start_section:
            section_content = start_section.group(1)
            i = 1
            while True:
                item_pattern = r'item{}\s*=\s*(.+)'.format(i)
                match = re.search(item_pattern, section_content)
                if not match:
                    break
                self.processes_to_start.append(match.group(1).strip())
                i += 1
        
        # 解析 ProcessesToTerminate 节
        term_section = re.search(r'\[ProcessesToTerminate\](.*?)(?=\[|$)', ini_content, re.DOTALL)
        if term_section:
            section_content = term_section.group(1)
            i = 1
            while True:
                item_pattern = r'item{}\s*=\s*(.+)'.format(i)
                match = re.search(item_pattern, section_content)
                if not match:
                    break
                self.processes_to_terminate.append(match.group(1).strip())
                i += 1
        
        # 解析 GUIProcessesToStart 节
        gui_start_section = re.search(r'\[GUIProcessesToStart\](.*?)(?=\[|$)', ini_content, re.DOTALL)
        if gui_start_section:
            section_content = gui_start_section.group(1)
            i = 1
            while True:
                name_pattern = r'Item{}_Name\s*=\s*(.+)'.format(i)
                path_pattern = r'Item{}_Path\s*=\s*(.+)'.format(i)
                checked_pattern = r'Item{}_Checked\s*=\s*(.+)'.format(i)
                
                name_match = re.search(name_pattern, section_content)
                path_match = re.search(path_pattern, section_content)
                
                if not name_match or not path_match:
                    break
                    
                # 检查是否有Checked属性
                checked_match = re.search(checked_pattern, section_content)
                checked = checked_match and checked_match.group(1).strip().lower() == 'true'
                
                self.gui_processes_to_start.append({
                    'name': name_match.group(1).strip(),
                    'path': path_match.group(1).strip(),
                    'checked': checked
                })
                i += 1
        
        # 解析 GUIProcessesToTerminate 节
        gui_term_section = re.search(r'\[GUIProcessesToTerminate\](.*?)(?=\[|$)', ini_content, re.DOTALL)
        if gui_term_section:
            section_content = gui_term_section.group(1)
            i = 1
            while True:
                name_pattern = r'Item{}_Name\s*=\s*(.+)'.format(i)
                path_pattern = r'Item{}_Path\s*=\s*(.+)'.format(i)
                checked_pattern = r'Item{}_Checked\s*=\s*(.+)'.format(i)
                
                name_match = re.search(name_pattern, section_content)
                path_match = re.search(path_pattern, section_content)
                
                if not name_match or not path_match:
                    break
                    
                # 检查是否有Checked属性
                checked_match = re.search(checked_pattern, section_content)
                checked = checked_match and checked_match.group(1).strip().lower() == 'true'
                
                self.gui_processes_to_terminate.append({
                    'name': name_match.group(1).strip(),
                    'path': path_match.group(1).strip(),
                    'checked': checked
                })
                i += 1
        
        # 更新进程管理表格
        self.update_process_tables()
    
    def update_process_tables(self):
        """更新进程管理页面的表格数据"""
        # 更新直接启用表格
        self.start_table.setRowCount(0)
        for i, path in enumerate(self.processes_to_start):
            self.start_table.insertRow(i)
            self.start_table.setItem(i, 0, QTableWidgetItem(path))
        
        # 更新直接终止表格
        self.term_table.setRowCount(0)
        for i, path in enumerate(self.processes_to_terminate):
            self.term_table.insertRow(i)
            self.term_table.setItem(i, 0, QTableWidgetItem(path))
        
        # 更新Ctrl+启用表格
        self.gui_start_table.setRowCount(0)
        for i, item in enumerate(self.gui_processes_to_start):
            self.gui_start_table.insertRow(i)
            self.gui_start_table.setItem(i, 0, QTableWidgetItem(item['name']))
            self.gui_start_table.setItem(i, 1, QTableWidgetItem(item['path']))
            
            checkbox = QTableWidgetItem()
            checkbox.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            checkbox.setCheckState(Qt.Checked if item['checked'] else Qt.Unchecked)
            self.gui_start_table.setItem(i, 2, checkbox)
        
        # 更新Ctrl+终止表格
        self.gui_term_table.setRowCount(0)
        for i, item in enumerate(self.gui_processes_to_terminate):
            self.gui_term_table.insertRow(i)
            self.gui_term_table.setItem(i, 0, QTableWidgetItem(item['name']))
            self.gui_term_table.setItem(i, 1, QTableWidgetItem(item['path']))
            
            checkbox = QTableWidgetItem()
            checkbox.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            checkbox.setCheckState(Qt.Checked if item['checked'] else Qt.Unchecked)
            self.gui_term_table.setItem(i, 2, checkbox)

    def load_website_config(self, ini_content):
        """加载网站配置部分"""
        # 重置网站配置数据
        self.default_site = ""
        self.browser_preference = "default"
        self.websites = []
        
        # 解析 CommonWebsites 节
        website_section = re.search(r'\[CommonWebsites\](.*?)(?=\[|$)', ini_content, re.DOTALL)
        if website_section:
            section_content = website_section.group(1)
            
            # 解析默认网站
            default_site_match = re.search(r'default_site\s*=\s*(.+)', section_content)
            if default_site_match:
                self.default_site = default_site_match.group(1).strip()
            
            # 解析浏览器偏好
            browser_match = re.search(r'browser\s*=\s*(.+)', section_content)
            if browser_match:
                self.browser_preference = browser_match.group(1).strip()
            
            # 解析网站列表
            sites = {}
            urls = {}
            
            for line in section_content.splitlines():
                line = line.strip()
                if line.startswith(';') or not line:  # 跳过注释和空行
                    continue
                    
                if '=' in line:
                    key, value = line.split('=', 1)
                    key = key.strip()
                    value = value.strip()
                    
                    if key.startswith('site') and key[4:].isdigit():
                        index = int(key[4:])
                        sites[index] = value
                    elif key.startswith('url') and key[3:].isdigit():
                        index = int(key[3:])
                        urls[index] = value
            
            # 合并网站和URL
            for index in sorted(set(sites.keys()) | set(urls.keys())):
                if index in sites and index in urls:
                    self.websites.append([sites[index], urls[index]])
        
        # 更新UI
        self.update_website_ui()
        
    def update_website_ui(self):
        """更新网站配置UI"""
        # 更新默认网站编辑框
        self.default_site_edit.setText(self.default_site)
        
        # 更新浏览器下拉框
        index = self.browser_combo.findText(self.browser_preference)
        if index >= 0:
            self.browser_combo.setCurrentIndex(index)
        
        # 更新网站表格
        self.website_table.setRowCount(0)
        for site, url in self.websites:
            row = self.website_table.rowCount()
            self.website_table.insertRow(row)
            self.website_table.setItem(row, 0, QTableWidgetItem(site))
            self.website_table.setItem(row, 1, QTableWidgetItem(url))
            
    def load_color_mode(self, ini_content):
        """加载暗色模式配置"""
        # 解析 MenuGroupsColourMode 节
        color_mode_section = re.search(r'\[MenuGroupsColourMode\](.*?)(?=\[|$)', ini_content, re.DOTALL)
        if color_mode_section:
            section_content = color_mode_section.group(1)
            
            # 解析暗色模式设置
            dark_mode_match = re.search(r'DarkMode\s*=\s*(.+)', section_content)
            if dark_mode_match:
                value = dark_mode_match.group(1).strip().lower()
                self.dark_mode = (value == "true")
        
        # 更新UI
        if hasattr(self, 'dark_mode_cb'):
            self.dark_mode_cb.setChecked(self.dark_mode)
            
    def save_menu_config(self, content):
        """保存菜单配置部分"""
        # 计算启用的组数
        enabled_count = sum(1 for i in range(1, 11) if self.group_enabled[i])
        
        # 更新MenuGroupNum部分
        num_pattern = r'(\[MenuGroupNum\].*?num\s*=\s*)(\d+)(.*?)(?=\[|$)'
        new_content = re.sub(num_pattern, r'\g<1>{}\g<3>'.format(enabled_count), content, flags=re.DOTALL)
        
        # 更新MenuGroupsEnable部分
        enable_section = "[MenuGroupsEnable]\n"
        for i in range(1, 11):
            value = "true" if self.group_enabled[i] else "false"
            enable_section += f"enableGroup{i} = {value}\n"
        enable_section += "\n"  # 添加空行
        
        # 替换MenuGroupsEnable部分
        enable_pattern = r'\[MenuGroupsEnable\].*?(?=\[|$)'
        new_content = re.sub(enable_pattern, enable_section, new_content, flags=re.DOTALL)
        
        # 更新MenuGroupName部分
        name_section = "[MenuGroupName]\n"
        for i in range(1, 11):
            name_section += f"name{i} = {self.group_names[i]}\n"
        name_section += "\n"  # 添加空行
        
        # 替换MenuGroupName部分
        name_pattern = r'\[MenuGroupName\].*?(?=\[|$)'
        new_content = re.sub(name_pattern, name_section, new_content, flags=re.DOTALL)
        
        # 更新MenuGroupCount部分
        count_section = "[MenuGroupCount]\n"
        for i in range(1, 11):
            count_section += f"count{i} = {len(self.group_items[i])}\n"
        count_section += "\n"  # 添加空行
        
        # 替换MenuGroupCount部分
        count_pattern = r'\[MenuGroupCount\].*?(?=\[|$)'
        new_content = re.sub(count_pattern, count_section, new_content, flags=re.DOTALL)
        
        # 更新各组项目
        for group_idx in range(1, 11):
            items = self.group_items[group_idx]
            if not items:
                continue  # 跳过空组
                
            items_section = f"[MenuGroups{group_idx}Items]\n"
            
            for item_idx, item in enumerate(items, 1):
                # 确保所有值都是字符串并对路径中的反斜杠进行转义
                name = str(item.get('name', ''))
                icon = str(item.get('icon', '')).replace('\\', '\\\\')  # 转义反斜杠
                icontype = str(item.get('icontype', 'emoji'))
                action = str(item.get('action', '')).replace('\\', '\\\\')  # 转义反斜杠
                
                items_section += f"name{item_idx} = {name}\n"
                items_section += f"icon{item_idx} = {icon}\n"
                items_section += f"icontype{item_idx} = {icontype}\n"
                items_section += f"action{item_idx} = {action}\n"
            
            items_section += "\n"  # 添加空行
            
            # 替换或添加组项目部分
            items_pattern = r'\[MenuGroups{}Items\].*?(?=\[|$)'.format(group_idx)
            if re.search(items_pattern, new_content, re.DOTALL):
                new_content = re.sub(items_pattern, items_section, new_content, flags=re.DOTALL)
            else:
                # 如果该组不存在，添加到文件末尾
                new_content += items_section
                
        return new_content
    
    def save_process_config(self, content):
        """保存进程管理配置部分"""
        new_content = content
        
        # 更新 ProcessesToStart 部分
        self.collect_process_data()
        
        # 构建新的ProcessesToStart节
        section = "[ProcessesToStart]\n"
        for i, path in enumerate(self.processes_to_start):
            # 处理路径中的反斜杠，转义为双反斜杠
            path = path.replace('\\', '\\\\')
            section += f"item{i+1} = {path}\n"
        section += "\n"
        
        # 替换ProcessesToStart节
        pattern = r'\[ProcessesToStart\](.*?)(?=\[|$)'
        if re.search(pattern, new_content, re.DOTALL):
            new_content = re.sub(pattern, section, new_content, flags=re.DOTALL)
        else:
            new_content += section
        
        # 构建新的ProcessesToTerminate节
        section = "[ProcessesToTerminate]\n"
        for i, path in enumerate(self.processes_to_terminate):
            # 处理路径中的反斜杠，转义为双反斜杠
            path = path.replace('\\', '\\\\')
            section += f"item{i+1} = {path}\n"
        section += "\n"
        
        # 替换ProcessesToTerminate节
        pattern = r'\[ProcessesToTerminate\](.*?)(?=\[|$)'
        if re.search(pattern, new_content, re.DOTALL):
            new_content = re.sub(pattern, section, new_content, flags=re.DOTALL)
        else:
            new_content += section
        
        # 构建新的GUIProcessesToStart节
        section = "[GUIProcessesToStart]\n"
        section += "; 在这里定义Ctrl+点击启用进程时，选择GUI中显示的进程\n"
        section += "; 格式: ItemX_Name=进程友好名称, ItemX_Path=进程路径或启动命令, ItemX_Checked=true/false (默认是否选中)\n"
        
        for i, item in enumerate(self.gui_processes_to_start):
            name = item['name']
            # 处理路径中的反斜杠，转义为双反斜杠
            path = item['path'].replace('\\', '\\\\')
            checked = str(item['checked']).lower()
            
            section += f"Item{i+1}_Name={name}\n"
            section += f"Item{i+1}_Path={path}\n"
            section += f"Item{i+1}_Checked={checked}\n"
        section += "\n"
        
        # 替换GUIProcessesToStart节
        pattern = r'\[GUIProcessesToStart\](.*?)(?=\[|$)'
        if re.search(pattern, new_content, re.DOTALL):
            new_content = re.sub(pattern, section, new_content, flags=re.DOTALL)
        else:
            new_content += section
        
        # 构建新的GUIProcessesToTerminate节
        section = "[GUIProcessesToTerminate]\n"
        section += "; 在这里定义Ctrl+点击终止进程时，选择GUI中显示的进程\n"
        section += "; 格式: ItemX_Name=进程友好名称, ItemX_Path=要终止的进程名, ItemX_Checked=true/false (默认是否选中)\n"
        
        for i, item in enumerate(self.gui_processes_to_terminate):
            name = item['name']
            # 处理路径中的反斜杠，转义为双反斜杠
            path = item['path'].replace('\\', '\\\\')
            checked = str(item['checked']).lower()
            
            section += f"Item{i+1}_Name={name}\n"
            section += f"Item{i+1}_Path={path}\n"
            section += f"Item{i+1}_Checked={checked}\n"
        section += "\n"
        
        # 替换GUIProcessesToTerminate节
        pattern = r'\[GUIProcessesToTerminate\](.*?)(?=\[|$)'
        if re.search(pattern, new_content, re.DOTALL):
            new_content = re.sub(pattern, section, new_content, flags=re.DOTALL)
        else:
            new_content += section
            
        return new_content
        
    def collect_process_data(self):
        """从表格中收集进程数据"""
        # 收集直接启用进程数据
        self.processes_to_start = []
        for row in range(self.start_table.rowCount()):
            path = self.start_table.item(row, 0).text()
            self.processes_to_start.append(path)
        
        # 收集直接终止进程数据
        self.processes_to_terminate = []
        for row in range(self.term_table.rowCount()):
            path = self.term_table.item(row, 0).text()
            self.processes_to_terminate.append(path)
        
        # 收集Ctrl+启用进程数据
        self.gui_processes_to_start = []
        for row in range(self.gui_start_table.rowCount()):
            name = self.gui_start_table.item(row, 0).text()
            path = self.gui_start_table.item(row, 1).text()
            checked = self.gui_start_table.item(row, 2).checkState() == Qt.Checked
            
            self.gui_processes_to_start.append({
                'name': name,
                'path': path,
                'checked': checked
            })
        
        # 收集Ctrl+终止进程数据
        self.gui_processes_to_terminate = []
        for row in range(self.gui_term_table.rowCount()):
            name = self.gui_term_table.item(row, 0).text()
            path = self.gui_term_table.item(row, 1).text()
            checked = self.gui_term_table.item(row, 2).checkState() == Qt.Checked
            
            self.gui_processes_to_terminate.append({
                'name': name,
                'path': path,
                'checked': checked
            })

    def update_group_list(self):
        """更新菜单组列表"""
        self.group_list.clear()
        
        # 存储索引到行的映射
        self.group_idx_map = {}
        row = 0
        
        # 先添加启用的组，使用带圈数字
        enabled_count = 0
        for i in range(1, 11):
            if self.group_names[i] and self.group_enabled[i]:
                # 使用带圈数字表示序号
                if enabled_count < len(CIRCLED_NUMBERS):
                    prefix = CIRCLED_NUMBERS[enabled_count]
                else:
                    prefix = f"{enabled_count+1}."
                    
                display_text = f"{prefix} {self.group_names[i]}"
                
                # 存储索引映射
                self.group_idx_map[row] = i
                
                # 添加到列表
                self.group_list.addItem(display_text)
                row += 1
                enabled_count += 1
        
        # 然后添加禁用的组，使用⓪符号
        for i in range(1, 11):
            if self.group_names[i] and not self.group_enabled[i]:
                # 使用⓪符号表示禁用
                display_text = f"{DISABLED_SYMBOL} {self.group_names[i]}"
                
                # 存储索引映射
                self.group_idx_map[row] = i
                
                # 添加到列表
                self.group_list.addItem(display_text)
                row += 1
                
        self.update_menu_buttons_state()
        
    def update_item_list(self):
        """更新当前选中组的菜单项列表"""
        self.item_list.clear()
        
        if self.current_group_idx < 1:
            return
            
        items = self.group_items[self.current_group_idx]
        for i, item in enumerate(items, 1):
            display_text = f"{i}. {item.get('name', '')}"
            
            # 添加图标类型指示
            icon_type = item.get('icontype', 'emoji')
            if icon_type == 'emoji':
                display_text += " [表情]"
            else:
                display_text += " [图标]"
            
            self.item_list.addItem(display_text)
            
        self.update_menu_buttons_state()
        
    def update_menu_buttons_state(self):
        """更新菜单配置页面的按钮状态"""
        group_selected = self.group_list.currentRow() >= 0
        item_selected = self.item_list.currentRow() >= 0
        
        self.edit_group_btn.setEnabled(group_selected)
        self.toggle_group_btn.setEnabled(group_selected)
        
        # 更新禁用/启用组按钮文本
        if group_selected and self.group_list.currentRow() in self.group_idx_map:
            group_idx = self.group_idx_map[self.group_list.currentRow()]
            if self.group_enabled[group_idx]:
                self.toggle_group_btn.setText("禁用此组")
            else:
                self.toggle_group_btn.setText("启用此组")
        
        self.up_group_btn.setEnabled(group_selected and self.group_list.currentRow() > 0)
        self.down_group_btn.setEnabled(group_selected and self.group_list.currentRow() < self.group_list.count() - 1)
        
        self.add_item_btn.setEnabled(group_selected)
        self.edit_item_btn.setEnabled(item_selected)
        self.delete_item_btn.setEnabled(item_selected)
        
        self.up_item_btn.setEnabled(item_selected and self.item_list.currentRow() > 0)
        self.down_item_btn.setEnabled(item_selected and self.item_list.currentRow() < self.item_list.count() - 1)
        
    def on_group_selected(self, row):
        """组选择变更时的回调"""
        if row >= 0 and row in self.group_idx_map:
            self.current_group_idx = self.group_idx_map[row]
            self.update_item_list()
        else:
            self.current_group_idx = 0
            self.item_list.clear()
            
        self.update_menu_buttons_state()
        
    def edit_group(self):
        """编辑菜单组名称"""
        current_row = self.group_list.currentRow()
        if current_row < 0 or current_row not in self.group_idx_map:
            return
        
        group_idx = self.group_idx_map[current_row]
        current_name = self.group_names[group_idx]
        
        new_name, ok = QInputDialog.getText(
            self, "编辑菜单组", "输入新的菜单组名称:", QLineEdit.Normal, current_name
        )
        
        if ok and new_name:
            self.group_names[group_idx] = new_name
            self.update_group_list()
            
            # 重新选中相同的组
            for row in range(self.group_list.count()):
                if row in self.group_idx_map and self.group_idx_map[row] == group_idx:
                    self.group_list.setCurrentRow(row)
                    break
            
    def toggle_group_enabled(self):
        """启用或禁用菜单组"""
        current_row = self.group_list.currentRow()
        if current_row < 0 or current_row not in self.group_idx_map:
            return
        
        group_idx = self.group_idx_map[current_row]
        group_name = self.group_names[group_idx]
        is_enabled = self.group_enabled[group_idx]
        
        # 根据当前状态准备操作消息
        action = "禁用" if is_enabled else "启用"
        message = f"确定要{action}菜单组 '{group_name}' 吗?"
        
        reply = QMessageBox.question(
            self, f"确认{action}", message,
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            # 切换启用状态
            self.group_enabled[group_idx] = not is_enabled
            
            # 更新UI
            self.update_group_list()
            
            # 寻找并重新选中相同的组
            for row in range(self.group_list.count()):
                if row in self.group_idx_map and self.group_idx_map[row] == group_idx:
                    self.group_list.setCurrentRow(row)
                    break
            
    def add_item(self):
        """添加菜单项"""
        if self.current_group_idx < 1:
            return
            
        dialog = MenuItemDialog(self)
        
        if dialog.exec_() == QDialog.Accepted:
            item_data = dialog.get_data()
            self.group_items[self.current_group_idx].append(item_data)
            self.update_item_list()
            
    def edit_item(self):
        """编辑菜单项"""
        if self.current_group_idx < 1:
            return
            
        item_idx = self.item_list.currentRow()
        if item_idx < 0 or item_idx >= len(self.group_items[self.current_group_idx]):
            return
            
        current_item = self.group_items[self.current_group_idx][item_idx]
        
        dialog = MenuItemDialog(self, current_item)
        
        if dialog.exec_() == QDialog.Accepted:
            self.group_items[self.current_group_idx][item_idx] = dialog.get_data()
            self.update_item_list()
            
    def delete_item(self):
        """删除菜单项"""
        if self.current_group_idx < 1:
            return
            
        item_idx = self.item_list.currentRow()
        if item_idx < 0 or item_idx >= len(self.group_items[self.current_group_idx]):
            return
            
        item_name = self.group_items[self.current_group_idx][item_idx].get("name", "")
        
        reply = QMessageBox.question(
            self, "确认删除", f"确定要删除菜单项 '{item_name}'?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            del self.group_items[self.current_group_idx][item_idx]
            self.update_item_list()
            
    def move_item_up(self):
        """上移菜单项"""
        self._move_item(-1)
        
    def move_item_down(self):
        """下移菜单项"""
        self._move_item(1)
        
    def _move_item(self, direction):
        """移动菜单项"""
        if self.current_group_idx < 1:
            return
            
        item_idx = self.item_list.currentRow()
        if item_idx < 0 or item_idx >= len(self.group_items[self.current_group_idx]):
            return
            
        new_idx = item_idx + direction
        if new_idx < 0 or new_idx >= len(self.group_items[self.current_group_idx]):
            return
            
        # 交换位置
        items = self.group_items[self.current_group_idx]
        items[item_idx], items[new_idx] = items[new_idx], items[item_idx]
        
        # 更新列表并选中移动后的项
        self.update_item_list()
        self.item_list.setCurrentRow(new_idx)

    def move_group_up(self):
        """上移菜单组"""
        self._move_group(-1)
        
    def move_group_down(self):
        """下移菜单组"""
        self._move_group(1)
        
    def _move_group(self, direction):
        """移动菜单组"""
        current_row = self.group_list.currentRow()
        if current_row < 0 or current_row not in self.group_idx_map:
            return
        
        # 获取当前组索引
        current_idx = self.group_idx_map[current_row]
        
        # 找到目标位置的组索引
        target_row = current_row + direction
        if target_row < 0 or target_row >= self.group_list.count() or target_row not in self.group_idx_map:
            return
        
        target_idx = self.group_idx_map[target_row]
        
        # 交换组数据
        self.group_names[current_idx], self.group_names[target_idx] = self.group_names[target_idx], self.group_names[current_idx]
        self.group_enabled[current_idx], self.group_enabled[target_idx] = self.group_enabled[target_idx], self.group_enabled[current_idx]
        self.group_items[current_idx], self.group_items[target_idx] = self.group_items[target_idx], self.group_items[current_idx]
        
        # 更新列表并选中移动后的位置
        self.update_group_list()
        
        # 寻找并选中移动后的项
        for row in range(self.group_list.count()):
            if row in self.group_idx_map and self.group_idx_map[row] == target_idx:
                self.group_list.setCurrentRow(row)
                break

    def read_ini_file(self):
        """读取INI文件的内容，直接返回文件内容字符串"""
        try:
            with open(self.ini_file, 'r', encoding='utf-8') as f:
                return f.read()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取INI文件失败: {str(e)}")
            return ""

    def save_color_mode(self, content):
        """保存暗色模式配置"""
        # 构建新的MenuGroupsColourMode节
        section = "[MenuGroupsColourMode]\n"
        section += f"DarkMode = {str(self.dark_mode).lower()}\n\n"
        
        # 替换MenuGroupsColourMode节
        pattern = r'\[MenuGroupsColourMode\](.*?)(?=\[|$)'
        if re.search(pattern, content, re.DOTALL):
            new_content = re.sub(pattern, section, content, flags=re.DOTALL)
        else:
            new_content = content + section
            
        return new_content

    def init_menu_page(self):
        """初始化菜单配置页面"""
        layout = QHBoxLayout()
        
        # 左侧面板 - 菜单组列表
        left_panel = QWidget()
        left_layout = QVBoxLayout()
        
        self.group_label = QLabel("菜单组:")
        left_layout.addWidget(self.group_label)
        
        self.group_list = QListWidget()
        self.group_list.currentRowChanged.connect(self.on_group_selected)
        left_layout.addWidget(self.group_list)
        
        group_buttons_layout = QHBoxLayout()
        self.edit_group_btn = QPushButton("编辑组名称")
        self.edit_group_btn.clicked.connect(self.edit_group)
        self.toggle_group_btn = QPushButton("启用/禁用")
        self.toggle_group_btn.clicked.connect(self.toggle_group_enabled)
        
        group_buttons_layout.addWidget(self.edit_group_btn)
        group_buttons_layout.addWidget(self.toggle_group_btn)
        
        left_layout.addLayout(group_buttons_layout)
        
        # 添加组上移下移按钮
        group_move_layout = QHBoxLayout()
        self.up_group_btn = QPushButton("上移")
        self.up_group_btn.clicked.connect(self.move_group_up)
        self.down_group_btn = QPushButton("下移")
        self.down_group_btn.clicked.connect(self.move_group_down)
        
        group_move_layout.addWidget(self.up_group_btn)
        group_move_layout.addWidget(self.down_group_btn)
        
        left_layout.addLayout(group_move_layout)
        
        left_panel.setLayout(left_layout)
        
        # 右侧面板 - 菜单项列表
        right_panel = QWidget()
        right_layout = QVBoxLayout()
        
        # 创建菜单项标题和暗色模式复选框所在的水平布局
        item_header_layout = QHBoxLayout()
        self.item_label = QLabel("菜单项:")
        item_header_layout.addWidget(self.item_label)
        
        # 添加弹性空间，将复选框推到右边
        item_header_layout.addStretch(1)
        
        # 添加暗色模式复选框
        self.dark_mode_cb = QCheckBox("暗色模式")
        self.dark_mode_cb.setChecked(self.dark_mode)
        item_header_layout.addWidget(self.dark_mode_cb)
        
        # 将水平布局添加到右侧面板
        right_layout.addLayout(item_header_layout)
        
        self.item_list = QListWidget()
        self.item_list.itemSelectionChanged.connect(self.update_menu_buttons_state)
        right_layout.addWidget(self.item_list)
        
        item_buttons_layout = QHBoxLayout()
        self.add_item_btn = QPushButton("添加项目")
        self.add_item_btn.clicked.connect(self.add_item)
        self.edit_item_btn = QPushButton("编辑项目")
        self.edit_item_btn.clicked.connect(self.edit_item)
        self.delete_item_btn = QPushButton("删除项目")
        self.delete_item_btn.clicked.connect(self.delete_item)
        
        item_buttons_layout.addWidget(self.add_item_btn)
        item_buttons_layout.addWidget(self.edit_item_btn)
        item_buttons_layout.addWidget(self.delete_item_btn)
        
        right_layout.addLayout(item_buttons_layout)
        
        item_move_layout = QHBoxLayout()
        self.up_item_btn = QPushButton("上移")
        self.up_item_btn.clicked.connect(self.move_item_up)
        self.down_item_btn = QPushButton("下移")
        self.down_item_btn.clicked.connect(self.move_item_down)
        
        item_move_layout.addWidget(self.up_item_btn)
        item_move_layout.addWidget(self.down_item_btn)
        
        right_layout.addLayout(item_move_layout)
        right_panel.setLayout(right_layout)
        
        # 添加左右面板到主布局
        layout.addWidget(left_panel, 1)
        layout.addWidget(right_panel, 2)
        
        self.menu_page.setLayout(layout)
        
        # 初始禁用按钮
        self.update_menu_buttons_state()
        
    def init_process_page(self):
        """初始化进程管理页面"""
        layout = QVBoxLayout()
        
        # 创建选项卡控件
        process_tabs = QTabWidget()
        
        # 添加四个选项卡
        process_tabs.addTab(self.create_simple_list_tab(self.processes_to_start, "直接启用进程列表", "start_table"), "直接启用")
        process_tabs.addTab(self.create_simple_list_tab(self.processes_to_terminate, "直接终止进程列表", "term_table"), "直接终止")
        process_tabs.addTab(self.create_gui_list_tab(self.gui_processes_to_start, "Ctrl+点击启用进程列表", "gui_start_table"), "Ctrl+启用")
        process_tabs.addTab(self.create_gui_list_tab(self.gui_processes_to_terminate, "Ctrl+点击终止进程列表", "gui_term_table"), "Ctrl+终止")
        
        layout.addWidget(process_tabs)
        
        self.process_page.setLayout(layout)
        
    def create_simple_list_tab(self, items, title, table_name):
        """创建简单列表选项卡"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        # 添加标题
        layout.addWidget(QLabel(title))
        
        # 创建表格
        table = QTableWidget(0, 1)
        table.setHorizontalHeaderLabels(["进程路径"])
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        
        # 添加数据
        for i, item in enumerate(items):
            table.insertRow(i)
            table.setItem(i, 0, QTableWidgetItem(item))
        
        # 保存表格引用到类实例
        setattr(self, table_name, table)
        
        layout.addWidget(table)
        
        # 添加按钮
        button_layout = QHBoxLayout()
        add_btn = QPushButton("添加")
        add_btn.clicked.connect(lambda: self.add_simple_process(table))
        delete_btn = QPushButton("删除")
        delete_btn.clicked.connect(lambda: self.delete_process(table))
        move_up_btn = QPushButton("上移")
        move_up_btn.clicked.connect(lambda: self.move_process(table, -1))
        move_down_btn = QPushButton("下移")
        move_down_btn.clicked.connect(lambda: self.move_process(table, 1))
        
        button_layout.addWidget(add_btn)
        button_layout.addWidget(delete_btn)
        button_layout.addWidget(move_up_btn)
        button_layout.addWidget(move_down_btn)
        
        layout.addLayout(button_layout)
        widget.setLayout(layout)
        
        return widget
    
    def create_gui_list_tab(self, items, title, table_name):
        """创建GUI列表选项卡"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        # 添加标题
        layout.addWidget(QLabel(title))
        
        # 创建表格
        table = QTableWidget(0, 3)
        table.setHorizontalHeaderLabels(["显示名称", "进程路径", "默认选中"])
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        
        # 添加数据
        for i, item in enumerate(items):
            table.insertRow(i)
            table.setItem(i, 0, QTableWidgetItem(item['name']))
            table.setItem(i, 1, QTableWidgetItem(item['path']))
            
            # 使用复选框表示是否选中
            checkbox = QTableWidgetItem()
            checkbox.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            checkbox.setCheckState(Qt.Checked if item['checked'] else Qt.Unchecked)
            table.setItem(i, 2, checkbox)
        
        # 保存表格引用到类实例
        setattr(self, table_name, table)
        
        layout.addWidget(table)
        
        # 添加按钮
        button_layout = QHBoxLayout()
        add_btn = QPushButton("添加")
        add_btn.clicked.connect(lambda: self.add_gui_process(table))
        delete_btn = QPushButton("删除")
        delete_btn.clicked.connect(lambda: self.delete_process(table))
        move_up_btn = QPushButton("上移")
        move_up_btn.clicked.connect(lambda: self.move_process(table, -1))
        move_down_btn = QPushButton("下移")
        move_down_btn.clicked.connect(lambda: self.move_process(table, 1))
        
        button_layout.addWidget(add_btn)
        button_layout.addWidget(delete_btn)
        button_layout.addWidget(move_up_btn)
        button_layout.addWidget(move_down_btn)
        
        layout.addLayout(button_layout)
        widget.setLayout(layout)
        
        return widget
    
    def add_simple_process(self, table):
        """添加简单进程项"""
        path, ok = QInputDialog.getText(self, "添加进程", "输入进程路径:", QLineEdit.Normal)
        if ok and path:
            row = table.rowCount()
            table.insertRow(row)
            table.setItem(row, 0, QTableWidgetItem(path))
    
    def add_gui_process(self, table):
        """添加GUI进程项"""
        # 创建添加对话框
        dialog = QDialog(self)
        dialog.setWindowTitle("添加进程项")
        dialog.resize(400, 150)
        
        layout = QFormLayout()
        
        name_edit = QLineEdit()
        layout.addRow("显示名称:", name_edit)
        
        path_edit = QLineEdit()
        path_layout = QHBoxLayout()
        path_layout.addWidget(path_edit)
        
        browse_btn = QPushButton("浏览...")
        browse_btn.clicked.connect(lambda: self.browse_process_path(path_edit))
        path_layout.addWidget(browse_btn)
        
        layout.addRow("进程路径:", path_layout)
        
        checked_box = QCheckBox()
        checked_box.setChecked(True)
        layout.addRow("默认选中:", checked_box)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        
        layout.addRow(buttons)
        dialog.setLayout(layout)
        
        if dialog.exec_() == QDialog.Accepted:
            name = name_edit.text()
            path = path_edit.text()
            checked = checked_box.isChecked()
            
            if name and path:
                row = table.rowCount()
                table.insertRow(row)
                table.setItem(row, 0, QTableWidgetItem(name))
                table.setItem(row, 1, QTableWidgetItem(path))
                
                checkbox = QTableWidgetItem()
                checkbox.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
                checkbox.setCheckState(Qt.Checked if checked else Qt.Unchecked)
                table.setItem(row, 2, checkbox)
    
    def browse_process_path(self, edit):
        """浏览文件路径"""
        file_name, _ = QFileDialog.getOpenFileName(
            self, "选择进程文件", "", "可执行文件 (*.exe *.bat *.lnk);;所有文件 (*.*)"
        )
        if file_name:
            edit.setText(file_name)
    
    def delete_process(self, table):
        """删除选中项"""
        rows = set()
        for item in table.selectedItems():
            rows.add(item.row())
        
        # 从后向前删除以避免索引问题
        for row in sorted(rows, reverse=True):
            table.removeRow(row)
    
    def move_process(self, table, direction):
        """上移或下移项目"""
        current_row = -1
        if table.selectedItems():
            current_row = table.selectedItems()[0].row()
        
        if current_row < 0:
            return
            
        target_row = current_row + direction
        
        if target_row < 0 or target_row >= table.rowCount():
            return
            
        # 保存目标行数据
        col_count = table.columnCount()
        target_data = []
        for col in range(col_count):
            item = table.item(target_row, col)
            if item:
                if col == 2 and item.flags() & Qt.ItemIsUserCheckable:  # 处理复选框列
                    target_data.append((item.text(), item.checkState()))
                else:
                    target_data.append(item.text())
            else:
                target_data.append("")
                
        # 保存当前行数据
        current_data = []
        for col in range(col_count):
            item = table.item(current_row, col)
            if item:
                if col == 2 and item.flags() & Qt.ItemIsUserCheckable:  # 处理复选框列
                    current_data.append((item.text(), item.checkState()))
                else:
                    current_data.append(item.text())
            else:
                current_data.append("")
                
        # 交换数据
        for col in range(col_count):
            if col == 2 and isinstance(target_data[col], tuple):  # 处理复选框列
                checkbox = QTableWidgetItem()
                checkbox.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
                checkbox.setCheckState(target_data[col][1])
                table.setItem(current_row, col, checkbox)
                
                checkbox = QTableWidgetItem()
                checkbox.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
                checkbox.setCheckState(current_data[col][1])
                table.setItem(target_row, col, checkbox)
            else:
                table.setItem(current_row, col, QTableWidgetItem(target_data[col]))
                table.setItem(target_row, col, QTableWidgetItem(current_data[col]))
        
        # 选中移动后的行
        table.clearSelection()
        table.selectRow(target_row)

    def init_website_page(self):
        """初始化网站配置页面"""
        layout = QVBoxLayout()
        
        # 默认网站设置
        default_site_layout = QHBoxLayout()
        default_site_layout.addWidget(QLabel("默认网站:"))
        self.default_site_edit = QLineEdit()
        default_site_layout.addWidget(self.default_site_edit)
        layout.addLayout(default_site_layout)
        
        # 浏览器偏好设置
        browser_layout = QHBoxLayout()
        browser_layout.addWidget(QLabel("浏览器偏好:"))
        self.browser_combo = QComboBox()
        self.browser_combo.addItems(["default", "edge", "chrome", "firefox"])
        browser_layout.addWidget(self.browser_combo)
        layout.addLayout(browser_layout)
        
        # 网站列表
        layout.addWidget(QLabel("网站列表:"))
        
        # 创建网站表格
        self.website_table = QTableWidget(0, 2)
        self.website_table.setHorizontalHeaderLabels(["网站名称", "URL"])
        self.website_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.website_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        layout.addWidget(self.website_table)
        
        # 网站操作按钮
        buttons_layout = QHBoxLayout()
        
        add_site_btn = QPushButton("添加网站")
        add_site_btn.clicked.connect(self.add_website)
        buttons_layout.addWidget(add_site_btn)
        
        delete_site_btn = QPushButton("删除网站")
        delete_site_btn.clicked.connect(self.delete_website)
        buttons_layout.addWidget(delete_site_btn)
        
        move_up_btn = QPushButton("上移")
        move_up_btn.clicked.connect(lambda: self.move_website(-1))
        buttons_layout.addWidget(move_up_btn)
        
        move_down_btn = QPushButton("下移")
        move_down_btn.clicked.connect(lambda: self.move_website(1))
        buttons_layout.addWidget(move_down_btn)
        
        layout.addLayout(buttons_layout)
        
        self.website_page.setLayout(layout)
        
    def add_website(self):
        """添加网站"""
        site_name, ok1 = QInputDialog.getText(self, "添加网站", "输入网站名称:")
        if ok1 and site_name:
            url, ok2 = QInputDialog.getText(self, "添加网站", "输入网站URL:")
            if ok2 and url:
                row = self.website_table.rowCount()
                self.website_table.insertRow(row)
                self.website_table.setItem(row, 0, QTableWidgetItem(site_name))
                self.website_table.setItem(row, 1, QTableWidgetItem(url))
                
    def delete_website(self):
        """删除网站"""
        current_row = self.website_table.currentRow()
        if current_row >= 0:
            site_name = self.website_table.item(current_row, 0).text()
            reply = QMessageBox.question(
                self, "确认删除", f"确定要删除网站 '{site_name}'?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.website_table.removeRow(current_row)
                
    def move_website(self, direction):
        """上移或下移网站"""
        current_row = self.website_table.currentRow()
        if current_row < 0:
            return
            
        target_row = current_row + direction
        if target_row < 0 or target_row >= self.website_table.rowCount():
            return
            
        # 交换两行数据
        for col in range(2):
            current_item = self.website_table.takeItem(current_row, col)
            target_item = self.website_table.takeItem(target_row, col)
            
            self.website_table.setItem(target_row, col, current_item)
            self.website_table.setItem(current_row, col, target_item)
        
        # 选中移动后的行
        self.website_table.setCurrentCell(target_row, 0)
        
    def save_website_config(self, content):
        """保存网站配置部分"""
        # 从UI收集数据
        self.collect_website_data()
        
        # 构建新的CommonWebsites节
        section = "[CommonWebsites]\n"
        section += "; 默认网站\n"
        section += f"default_site = {self.default_site}\n"
        section += "; 浏览器偏好: edge, chrome, firefox, default(默认)\n"
        section += f"browser = {self.browser_preference}\n"
        section += "\n"
        section += "; 网站列表\n"
        
        for i, (site, url) in enumerate(self.websites, 1):
            section += f"site{i} = {site}\n"
            section += f"url{i} = {url}\n"
        
        section += "\n"
        
        # 替换CommonWebsites节
        pattern = r'\[CommonWebsites\](.*?)(?=\[|$)'
        if re.search(pattern, content, re.DOTALL):
            new_content = re.sub(pattern, section, content, flags=re.DOTALL)
        else:
            new_content = content + section
            
        return new_content
        
    def collect_website_data(self):
        """从UI收集网站配置数据"""
        # 收集默认网站
        self.default_site = self.default_site_edit.text()
        
        # 收集浏览器偏好
        self.browser_preference = self.browser_combo.currentText()
        
        # 收集网站列表
        self.websites = []
        for row in range(self.website_table.rowCount()):
            site = self.website_table.item(row, 0).text()
            url = self.website_table.item(row, 1).text()
            self.websites.append([site, url])

    def init_note_page(self):
        """初始化速记路径配置页面"""
        layout = QVBoxLayout()
        
        # 速记目标表格
        self.note_table = QTableWidget(0, 2)
        self.note_table.setHorizontalHeaderLabels(["关键词", "文件路径"])
        self.note_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.note_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        layout.addWidget(self.note_table)
        
        # 速记目标按钮
        button_layout = QHBoxLayout()
        add_note_btn = QPushButton("添加")
        add_note_btn.clicked.connect(self.add_note_target)
        edit_note_btn = QPushButton("编辑")
        edit_note_btn.clicked.connect(self.edit_note_target)
        delete_note_btn = QPushButton("删除")
        delete_note_btn.clicked.connect(self.delete_note_target)
        
        button_layout.addWidget(add_note_btn)
        button_layout.addWidget(edit_note_btn)
        button_layout.addWidget(delete_note_btn)
        layout.addLayout(button_layout)
        
        # 移动按钮
        move_layout = QHBoxLayout()
        up_btn = QPushButton("上移")
        up_btn.clicked.connect(lambda: self.move_note_target(-1))
        down_btn = QPushButton("下移")
        down_btn.clicked.connect(lambda: self.move_note_target(1))
        
        move_layout.addStretch(1)
        move_layout.addWidget(up_btn)
        move_layout.addWidget(down_btn)
        move_layout.addStretch(1)
        layout.addLayout(move_layout)
        
        self.note_page.setLayout(layout)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MenuConfigTool()
    window.show()
    sys.exit(app.exec_()) 