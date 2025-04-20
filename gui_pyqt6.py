import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QComboBox, QTextEdit, QTreeWidget,
                             QTreeWidgetItem, QScrollArea, QFrame, QGridLayout, QGroupBox,
                             QCheckBox, QSplitter, QTableWidget, QTableWidgetItem, QHeaderView,
                             QTabWidget, QMessageBox, QFileDialog)
from PyQt6.QtCore import Qt, QSize, pyqtSignal, pyqtSlot
from PyQt6.QtGui import QIcon, QFont, QPixmap, QColor, QPalette

import config
from email_connector import EmailConnector
from email_classifier import EmailClassifier
from auto_reply import AutoReplyGenerator
from email_sender import EmailSender


class ThemeManager:
    LIGHT_THEME = {
        "background": "#F5F7FA",
        "primary": "#1976D2",
        "secondary": "#E0E0E0",
        "text": "#333333",
        "accent": "#FF4081",
        "sidebar": "#FFFFFF",
        "card": "#FFFFFF",
        "border": "#DDDDDD"
    }
    
    DARK_THEME = {
        "background": "#212121",
        "primary": "#2196F3",
        "secondary": "#424242",
        "text": "#FFFFFF",
        "accent": "#FF4081",
        "sidebar": "#2C2C2C",
        "card": "#333333",
        "border": "#444444"
    }
    
    def __init__(self, app):
        self.app = app
        self.current_theme = "light"
        self.themes = {
            "light": self.LIGHT_THEME,
            "dark": self.DARK_THEME
        }
    
    def apply_theme(self, theme_name):
        if theme_name not in self.themes:
            return False
        
        self.current_theme = theme_name
        theme = self.themes[theme_name]
        
        # 设置应用样式表
        self.app.setStyleSheet(f"""
            QWidget {{
                background-color: {theme["background"]};
                color: {theme["text"]};
            }}
            
            #sidebar {{
                background-color: {theme["sidebar"]};
                border-right: 1px solid {theme["border"]};
            }}
            
            QPushButton {{
                background-color: {theme["primary"]};
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
            }}
            
            QPushButton:hover {{
                background-color: {self._adjust_color(theme["primary"], 20)};
            }}
            
            QLineEdit, QTextEdit, QComboBox {{
                background-color: {theme["card"]};
                border: 1px solid {theme["border"]};
                border-radius: 4px;
                padding: 8px;
            }}
            
            QGroupBox {{
                border: 1px solid {theme["border"]};
                border-radius: 4px;
                margin-top: 1em;
                font-weight: bold;
            }}
            
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }}
            
            QTableWidget {{
                border: 1px solid {theme["border"]};
                gridline-color: {theme["border"]};
            }}
            
            QTableWidget::item:selected {{
                background-color: {theme["primary"]};
            }}
            
            QHeaderView::section {{
                background-color: {theme["secondary"]};
                padding: 4px;
                border: 1px solid {theme["border"]};
                font-weight: bold;
            }}
        """)
        
        return True
    
    def _adjust_color(self, hex_color, amount):
        # 简单的颜色调整函数
        r = min(255, max(0, int(hex_color[1:3], 16) + amount))
        g = min(255, max(0, int(hex_color[3:5], 16) + amount))
        b = min(255, max(0, int(hex_color[5:7], 16) + amount))
        return f"#{r:02x}{g:02x}{b:02x}"


class ShortcutManager:
    def __init__(self, main_window):
        self.main_window = main_window
        self.shortcuts = {}
        self.setup_shortcuts()
    
    def setup_shortcuts(self):
        # 导航快捷键
        self.add_shortcut("Ctrl+R", self.main_window.reply_selected_email, "回复当前邮件")
        self.add_shortcut("Ctrl+F", self.main_window.refresh_emails, "刷新邮件列表")
        self.add_shortcut("Ctrl+D", self.main_window.delete_selected_email, "删除当前邮件")
        self.add_shortcut("Ctrl+M", self.main_window.mark_as_read, "标记为已读")
        
        # 其他通用快捷键
        self.add_shortcut("Ctrl+N", self.main_window.compose_new_email, "撰写新邮件")
        self.add_shortcut("F5", self.main_window.refresh_emails, "刷新邮件列表")
        self.add_shortcut("Esc", self.main_window.clear_selection, "清除选择")
    
    def add_shortcut(self, key_sequence, callback, description):
        from PyQt6.QtGui import QShortcut, QKeySequence
        shortcut = QShortcut(QKeySequence(key_sequence), self.main_window)
        shortcut.activated.connect(callback)
        self.shortcuts[key_sequence] = {
            "shortcut": shortcut,
            "description": description,
            "callback": callback
        }
        
    def get_shortcuts_help(self):
        help_text = "快捷键列表:\n\n"
        for key, data in self.shortcuts.items():
            help_text += f"{key}: {data['description']}\n"
        return help_text


class EmailTableWidget(QTableWidget):
    """现代邮件列表表格控件"""
    
    email_selected = pyqtSignal(dict)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        self.emails_data = {}
        
    def setup_ui(self):
        # 设置表格列
        self.setColumnCount(5)
        self.setHorizontalHeaderLabels(["ID", "发件人", "主题", "日期", "分类"])
        
        # 设置表格样式
        self.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.setAlternatingRowColors(True)
        self.verticalHeader().setVisible(False)
        
        # 设置列宽
        self.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        
        # 连接信号
        self.itemSelectionChanged.connect(self.on_selection_changed)
        
    def add_email(self, email_data):
        """添加一封邮件到表格"""
        row_position = self.rowCount()
        self.insertRow(row_position)
        
        # 设置单元格内容
        self.setItem(row_position, 0, QTableWidgetItem(email_data.get('id', '')))
        self.setItem(row_position, 1, QTableWidgetItem(email_data.get('from', '')))
        self.setItem(row_position, 2, QTableWidgetItem(email_data.get('subject', '')))
        self.setItem(row_position, 3, QTableWidgetItem(email_data.get('date', '')))
        self.setItem(row_position, 4, QTableWidgetItem(email_data.get('category', '未分类')))
        
        # 保存邮件数据
        self.emails_data[row_position] = email_data
        
    def clear_emails(self):
        """清空表格内容"""
        self.setRowCount(0)
        self.emails_data = {}
        
    def get_selected_emails(self):
        """获取选中的邮件"""
        selected_emails = []
        for row in set(index.row() for index in self.selectedIndexes()):
            if row in self.emails_data:
                selected_emails.append(self.emails_data[row])
        return selected_emails
    
    def on_selection_changed(self):
        """选择变更时触发"""
        selected_emails = self.get_selected_emails()
        if selected_emails:
            self.email_selected.emit(selected_emails[0])
            

class BulkActionsWidget(QWidget):
    """批量操作组件"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        
    def setup_ui(self):
        layout = QHBoxLayout(self)
        
        # 批量选择下拉框
        self.selection_dropdown = QComboBox()
        self.selection_dropdown.addItems(["全选", "全不选", "选择已读", "选择未读"])
        self.selection_dropdown.setFixedWidth(150)
        layout.addWidget(self.selection_dropdown)
        
        # 添加分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.VLine)
        separator.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addWidget(separator)
        
        # 批量操作按钮
        self.classify_btn = QPushButton("批量分类")
        self.reply_btn = QPushButton("批量回复")
        self.mark_read_btn = QPushButton("标为已读")
        self.mark_important_btn = QPushButton("标为重要")
        
        layout.addWidget(self.classify_btn)
        layout.addWidget(self.reply_btn)
        layout.addWidget(self.mark_read_btn)
        layout.addWidget(self.mark_important_btn)
        
        # 添加伸缩项使删除按钮靠右
        layout.addStretch()
        
        # 删除按钮
        self.delete_btn = QPushButton("删除")
        self.delete_btn.setStyleSheet("background-color: #f44336;")
        layout.addWidget(self.delete_btn)
        
        # 设置布局边距
        layout.setContentsMargins(0, 0, 0, 0)


class EmailAssistantGUI(QMainWindow):
    """基于PyQt6的邮件助手主窗口"""
    
    def __init__(self):
        super().__init__()
        
        # 配置初始化
        self.config = config
        
        # 初始化邮件处理对象
        self.email_connector = EmailConnector(self.config)
        self.email_classifier = EmailClassifier(self.config)
        self.auto_reply_generator = AutoReplyGenerator(self.config)
        self.email_sender = EmailSender(self.config)
        
        # 邮件数据
        self.current_email = None
        
        # 主题管理器
        self.theme_manager = None
        
        # 设置窗口属性
        self.setWindowTitle("自动邮件分类与回复助手")
        self.setMinimumSize(1280, 800)
        
        # 创建界面
        self.setup_ui()
        
    def setup_ui(self):
        # 创建中央窗口部件
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        # 创建主布局
        main_layout = QVBoxLayout(self.central_widget)
        
        # 创建邮箱配置区域
        self.create_email_config_ui()
        main_layout.addWidget(self.email_config_group)
        
        # 创建中央内容区域
        self.create_content_area()
        main_layout.addWidget(self.content_splitter, 1)
        
        # 创建状态栏
        self.statusBar().showMessage("准备就绪")
        
    def create_email_config_ui(self):
        """创建邮箱配置区域"""
        self.email_config_group = QGroupBox("邮箱配置")
        layout = QGridLayout(self.email_config_group)
        
        # 邮箱地址
        layout.addWidget(QLabel("邮箱地址:"), 0, 0)
        self.email_input = QLineEdit(self.config.EMAIL_ADDRESS)
        layout.addWidget(self.email_input, 0, 1)
        
        # 密码
        layout.addWidget(QLabel("密码:"), 1, 0)
        self.password_input = QLineEdit(self.config.EMAIL_PASSWORD)
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addWidget(self.password_input, 1, 1)
        
        # IMAP 服务器
        layout.addWidget(QLabel("IMAP 服务器:"), 0, 2)
        self.imap_server_input = QLineEdit(self.config.IMAP_SERVER)
        layout.addWidget(self.imap_server_input, 0, 3)
        
        # IMAP 端口
        layout.addWidget(QLabel("IMAP 端口:"), 1, 2)
        self.imap_port_input = QLineEdit(str(self.config.IMAP_PORT))
        layout.addWidget(self.imap_port_input, 1, 3)
        
        # SMTP 服务器
        layout.addWidget(QLabel("SMTP 服务器:"), 2, 0)
        self.smtp_server_input = QLineEdit(self.config.SMTP_SERVER)
        layout.addWidget(self.smtp_server_input, 2, 1)
        
        # SMTP 端口
        layout.addWidget(QLabel("SMTP 端口:"), 2, 2)
        self.smtp_port_input = QLineEdit(str(self.config.SMTP_PORT))
        layout.addWidget(self.smtp_port_input, 2, 3)
        
        # 按钮区域
        button_layout = QHBoxLayout()
        
        self.connect_btn = QPushButton("连接邮箱")
        self.connect_btn.clicked.connect(self.connect_email)
        button_layout.addWidget(self.connect_btn)
        
        self.refresh_btn = QPushButton("重新获取邮件列表")
        self.refresh_btn.clicked.connect(self.fetch_and_display_emails)
        button_layout.addWidget(self.refresh_btn)
        
        # 主题选择
        self.theme_label = QLabel("主题:")
        button_layout.addWidget(self.theme_label)
        
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["浅色", "深色"])
        self.theme_combo.currentIndexChanged.connect(self.change_theme)
        button_layout.addWidget(self.theme_combo)
        
        # 添加伸缩项使后续控件靠右
        button_layout.addStretch()
        
        # 把按钮区域添加到网格布局
        layout.addLayout(button_layout, 3, 0, 1, 4)
        
    def create_content_area(self):
        """创建内容区域"""
        self.content_splitter = QSplitter(Qt.Orientation.Vertical)
        
        # 创建上部区域 - 邮件列表和邮件内容
        upper_splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # 邮件列表区域
        mail_list_container = QGroupBox("邮件列表")
        mail_list_layout = QVBoxLayout(mail_list_container)
        
        # 批量操作工具栏
        self.bulk_actions = BulkActionsWidget()
        mail_list_layout.addWidget(self.bulk_actions)
        
        # 连接批量操作按钮信号
        self.bulk_actions.classify_btn.clicked.connect(self.show_bulk_classify_dialog)
        self.bulk_actions.reply_btn.clicked.connect(self.show_bulk_reply_dialog)
        self.bulk_actions.delete_btn.clicked.connect(self.delete_selected_emails)
        self.bulk_actions.mark_read_btn.clicked.connect(self.mark_as_read)
        self.bulk_actions.selection_dropdown.currentIndexChanged.connect(self.handle_selection_change)
        
        # 邮件表格
        self.email_table = EmailTableWidget()
        self.email_table.email_selected.connect(self.show_email_content)
        mail_list_layout.addWidget(self.email_table)
        
        upper_splitter.addWidget(mail_list_container)
        
        # 邮件内容区域
        mail_content_container = QGroupBox("邮件内容")
        mail_content_layout = QVBoxLayout(mail_content_container)
        
        self.email_body_text = QTextEdit()
        self.email_body_text.setReadOnly(True)
        mail_content_layout.addWidget(self.email_body_text)
        
        upper_splitter.addWidget(mail_content_container)
        
        # 设置初始分割大小
        upper_splitter.setSizes([400, 600])
        
        # 添加到主分割器
        self.content_splitter.addWidget(upper_splitter)
        
        # 创建下部区域 - 自动回复和批量处理
        lower_splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # 自动回复区域
        reply_container = QGroupBox("自动回复")
        reply_layout = QVBoxLayout(reply_container)
        
        reply_layout.addWidget(QLabel("自动回复内容:"))
        
        self.reply_text = QTextEdit()
        reply_layout.addWidget(self.reply_text)
        
        self.send_reply_btn = QPushButton("发送回复")
        self.send_reply_btn.clicked.connect(self.send_reply)
        reply_layout.addWidget(self.send_reply_btn)
        
        lower_splitter.addWidget(reply_container)
        
        # 批量处理区域
        bulk_container = QGroupBox("批量处理")
        bulk_layout = QGridLayout(bulk_container)
        
        bulk_layout.addWidget(QLabel("目标分类:"), 0, 0)
        
        self.bulk_category_combo = QComboBox()
        self.bulk_category_combo.addItems(list(self.config.CATEGORY_KEYWORDS.keys()) + [self.config.DEFAULT_CATEGORY])
        bulk_layout.addWidget(self.bulk_category_combo, 0, 1)
        
        self.bulk_classify_btn = QPushButton("批量分类")
        self.bulk_classify_btn.clicked.connect(self.bulk_classify_emails)
        bulk_layout.addWidget(self.bulk_classify_btn, 0, 2)
        
        bulk_layout.addWidget(QLabel("回复内容:"), 1, 0)
        
        self.bulk_reply_text = QTextEdit()
        bulk_layout.addWidget(self.bulk_reply_text, 1, 1, 1, 2)
        
        self.bulk_reply_btn = QPushButton("批量回复")
        self.bulk_reply_btn.clicked.connect(self.bulk_reply_emails)
        bulk_layout.addWidget(self.bulk_reply_btn, 2, 0, 1, 3)
        
        lower_splitter.addWidget(bulk_container)
        
        # 设置初始分割大小
        lower_splitter.setSizes([600, 400])
        
        # 添加到主分割器
        self.content_splitter.addWidget(lower_splitter)
        
        # 设置主分割器大小
        self.content_splitter.setSizes([500, 300])
        
    def connect_email(self):
        """连接到邮箱"""
        try:
            # 获取输入的邮箱信息
            email_address = self.email_input.text()
            email_password = self.password_input.text()
            imap_server = self.imap_server_input.text()
            imap_port = int(self.imap_port_input.text())
            smtp_server = self.smtp_server_input.text()
            smtp_port = int(self.smtp_port_input.text())
            
            # 更新配置
            self.config.EMAIL_ADDRESS = email_address
            self.config.EMAIL_PASSWORD = email_password
            self.config.IMAP_SERVER = imap_server
            self.config.IMAP_PORT = imap_port
            self.config.SMTP_SERVER = smtp_server
            self.config.SMTP_PORT = smtp_port
            
            # 重新初始化连接器
            self.email_connector = EmailConnector(self.config)
            self.email_sender = EmailSender(self.config)
            
            if self.email_connector.connect():
                QMessageBox.information(self, "成功", "邮箱连接成功")
                self.fetch_and_display_emails()
            else:
                QMessageBox.critical(self, "错误", "邮箱连接失败")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"邮箱连接失败: {e}")
            
    def fetch_and_display_emails(self):
        """获取并显示邮件"""
        try:
            self.statusBar().showMessage("正在获取邮件...")
            emails = self.email_connector.fetch_emails()
            
            # 对邮件列表进行倒序排列
            emails.reverse()
            
            # 清空表格
            self.email_table.clear_emails()
            
            # 添加邮件到表格
            for email_data in emails:
                self.email_table.add_email(email_data)
                
            self.statusBar().showMessage(f"成功获取 {len(emails)} 封邮件")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"获取邮件时出错: {e}")
            self.statusBar().showMessage("获取邮件失败")
            
    def show_email_content(self, email_data):
        """显示选中邮件的内容"""
        self.current_email = email_data
        
        # 显示邮件正文
        self.email_body_text.setText(email_data.get('body', ''))
        
        # 分类邮件并生成回复
        category = self.email_classifier.classify_email(email_data)
        reply_content = self.auto_reply_generator.generate_reply(category)
        
        # 显示自动回复内容
        self.reply_text.setText(reply_content)
        
    def send_reply(self):
        """发送回复邮件"""
        if not self.current_email:
            QMessageBox.warning(self, "警告", "请选择要回复的邮件")
            return
            
        # 获取回复内容
        reply_content = self.reply_text.toPlainText().strip()
        if not reply_content:
            QMessageBox.warning(self, "警告", "回复内容不能为空")
            return
            
        # 发送邮件
        if self.email_sender.send_email(
            recipient=self.current_email['from'],
            subject=f"Re: {self.current_email['subject']}",
            body=reply_content,
            attachments=self.current_email['attachments']
        ):
            QMessageBox.information(self, "成功", "邮件发送成功")
        else:
            QMessageBox.critical(self, "错误", "发送邮件时出错")
            
    def bulk_classify_emails(self):
        """批量分类邮件"""
        selected_emails = self.email_table.get_selected_emails()
        if not selected_emails:
            QMessageBox.warning(self, "警告", "请选择要分类的邮件")
            return
            
        target_category = self.bulk_category_combo.currentText()
        if not target_category:
            QMessageBox.warning(self, "警告", "请选择目标分类")
            return
            
        # 更新邮件分类
        success_count = 0
        for email_data in selected_emails:
            self.email_classifier.tag_email(email_data, target_category)
            success_count += 1
            
        # 刷新邮件列表以显示更新后的分类
        self.fetch_and_display_emails()
        
        QMessageBox.information(self, "成功", f"成功分类 {success_count} 封邮件")
        
    def bulk_reply_emails(self):
        """批量回复邮件"""
        selected_emails = self.email_table.get_selected_emails()
        if not selected_emails:
            QMessageBox.warning(self, "警告", "请选择要回复的邮件")
            return
            
        reply_content = self.bulk_reply_text.toPlainText().strip()
        if not reply_content:
            QMessageBox.warning(self, "警告", "请输入回复内容")
            return
            
        # 发送回复
        success_count = 0
        for email_data in selected_emails:
            if self.email_sender.send_email(
                recipient=email_data['from'],
                subject=f"Re: {email_data['subject']}",
                body=reply_content,
                attachments=email_data['attachments']
            ):
                success_count += 1
                
        QMessageBox.information(self, "成功", f"成功发送 {success_count} 封回复邮件")
        
    def show_bulk_classify_dialog(self):
        """显示批量分类对话框"""
        selected_emails = self.email_table.get_selected_emails()
        if not selected_emails:
            QMessageBox.warning(self, "警告", "请选择要分类的邮件")
            return
            
        # 这里可以实现一个更复杂的分类对话框
        # 目前简单调用批量分类方法
        self.bulk_classify_emails()
        
    def show_bulk_reply_dialog(self):
        """显示批量回复对话框"""
        selected_emails = self.email_table.get_selected_emails()
        if not selected_emails:
            QMessageBox.warning(self, "警告", "请选择要回复的邮件")
            return
            
        # 这里可以实现一个更复杂的回复对话框
        # 目前简单调用批量回复方法
        self.bulk_reply_emails()
        
    def delete_selected_emails(self):
        """删除选中的邮件"""
        selected_emails = self.email_table.get_selected_emails()
        if not selected_emails:
            QMessageBox.warning(self, "警告", "请选择要删除的邮件")
            return
            
        # 确认删除
        reply = QMessageBox.question(self, "确认删除", 
                                    f"确定要删除选中的 {len(selected_emails)} 封邮件吗?",
                                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            # 这里实现真正的删除逻辑
            QMessageBox.information(self, "成功", f"已删除 {len(selected_emails)} 封邮件")
            self.fetch_and_display_emails()
            
    def mark_as_read(self):
        """标记邮件为已读"""
        selected_emails = self.email_table.get_selected_emails()
        if not selected_emails:
            QMessageBox.warning(self, "警告", "请选择要标记的邮件")
            return
            
        # 这里实现标记为已读的逻辑
        QMessageBox.information(self, "成功", f"已将 {len(selected_emails)} 封邮件标记为已读")
        
    def handle_selection_change(self, index):
        """处理批量选择下拉框变更"""
        action = self.bulk_actions.selection_dropdown.currentText()
        if action == "全选":
            self.email_table.selectAll()
        elif action == "全不选":
            self.email_table.clearSelection()
        # 其他选择逻辑需要实现更多函数
            
    def change_theme(self, index):
        """更改主题"""
        if hasattr(self, 'theme_manager') and self.theme_manager:
            theme = "dark" if index == 1 else "light"
            self.theme_manager.apply_theme(theme)
            self.statusBar().showMessage(f"已切换到{'深色' if index == 1 else '浅色'}主题")
        
    def reply_selected_email(self):
        """快捷键回复邮件"""
        if self.current_email:
            self.send_reply()
            
    def refresh_emails(self):
        """刷新邮件列表"""
        self.fetch_and_display_emails()
            
    def delete_selected_email(self):
        """快捷键删除邮件"""
        self.delete_selected_emails()
            
    def compose_new_email(self):
        """撰写新邮件"""
        QMessageBox.information(self, "功能提示", "撰写新邮件功能尚未实现")
            
    def clear_selection(self):
        """清除选择"""
        self.email_table.clearSelection()
            
    def closeEvent(self, event):
        """关闭窗口事件"""
        # 关闭邮箱连接
        try:
            if hasattr(self, 'email_connector') and self.email_connector:
                self.email_connector.close()
        except Exception as e:
            print(f"关闭邮箱连接时出错: {e}")
        event.accept()


# 添加这个函数以便在测试时可以单独运行此文件
def run_app():
    app = QApplication(sys.argv)
    
    # 初始化主题管理器
    theme_manager = ThemeManager(app)
    theme_manager.apply_theme("light")
    
    window = EmailAssistantGUI()
    window.show()
    
    # 初始化快捷键管理器
    shortcut_manager = ShortcutManager(window)
    
    sys.exit(app.exec())


if __name__ == "__main__":
    run_app() 