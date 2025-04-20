import os
import datetime
import mimetypes
import shutil
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                            QPushButton, QFileDialog, QListWidget, QListWidgetItem,
                            QMessageBox, QMenu, QInputDialog)
from PyQt6.QtCore import Qt, pyqtSignal, QPoint
from PyQt6.QtGui import QIcon, QAction


class AttachmentHandler:
    """附件处理类，用于管理邮件附件"""
    
    def __init__(self, attachment_dir="attachments"):
        """初始化附件处理器
        
        Args:
            attachment_dir: 附件存储目录
        """
        self.attachment_dir = attachment_dir
        
        # 确保附件目录存在
        if not os.path.exists(attachment_dir):
            os.makedirs(attachment_dir)
            
    def save_attachment(self, attachment_data, email_id=None, custom_filename=None):
        """保存附件到磁盘
        
        Args:
            attachment_data: 附件数据字典
            email_id: 邮件ID，用于创建子目录
            custom_filename: 自定义文件名，如不指定则使用原始文件名
            
        Returns:
            str: 保存路径
        """
        try:
            # 确定保存目录
            save_dir = self.attachment_dir
            if email_id:
                save_dir = os.path.join(self.attachment_dir, email_id)
                if not os.path.exists(save_dir):
                    os.makedirs(save_dir)
            
            # 确定文件名
            filename = custom_filename or attachment_data.get('filename', f"attachment_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}")
            
            # 完整保存路径
            save_path = os.path.join(save_dir, filename)
            
            # 保存文件
            with open(save_path, 'wb') as f:
                f.write(attachment_data.get('data', b''))
                
            return save_path
        except Exception as e:
            print(f"保存附件失败: {e}")
            return None
    
    def save_all_attachments(self, email_data):
        """保存一封邮件的所有附件
        
        Args:
            email_data: 邮件数据
            
        Returns:
            list: 保存路径列表
        """
        saved_paths = []
        
        if 'attachments' in email_data and email_data['attachments']:
            for attachment in email_data['attachments']:
                path = self.save_attachment(attachment, email_data.get('id'))
                if path:
                    saved_paths.append(path)
                    
        return saved_paths
    
    def get_attachment_by_email(self, email_id):
        """获取指定邮件的所有附件
        
        Args:
            email_id: 邮件ID
            
        Returns:
            list: 附件路径列表
        """
        attachment_dir = os.path.join(self.attachment_dir, email_id)
        
        if not os.path.exists(attachment_dir):
            return []
        
        attachments = []
        for filename in os.listdir(attachment_dir):
            attachments.append(os.path.join(attachment_dir, filename))
            
        return attachments
    
    def get_file_icon(self, file_path):
        """获取文件图标类型
        
        Args:
            file_path: 文件路径
            
        Returns:
            str: 图标名称
        """
        mime_type, _ = mimetypes.guess_type(file_path)
        
        if not mime_type:
            return "file"
            
        if mime_type.startswith("image/"):
            return "image"
        elif mime_type.startswith("text/"):
            return "text"
        elif mime_type.startswith("audio/"):
            return "audio"
        elif mime_type.startswith("video/"):
            return "video"
        elif mime_type.startswith("application/pdf"):
            return "pdf"
        elif "spreadsheet" in mime_type or "excel" in mime_type:
            return "excel"
        elif "presentation" in mime_type or "powerpoint" in mime_type:
            return "powerpoint"
        elif "document" in mime_type or "word" in mime_type:
            return "word"
        elif "zip" in mime_type or "rar" in mime_type or "tar" in mime_type or "gzip" in mime_type:
            return "archive"
        else:
            return "file"
    
    def copy_to_directory(self, attachment_path, target_dir):
        """将附件复制到指定目录
        
        Args:
            attachment_path: 附件路径
            target_dir: 目标目录
            
        Returns:
            str: 复制后的路径
        """
        try:
            if not os.path.exists(target_dir):
                os.makedirs(target_dir)
                
            filename = os.path.basename(attachment_path)
            target_path = os.path.join(target_dir, filename)
            
            shutil.copy2(attachment_path, target_path)
            return target_path
        except Exception as e:
            print(f"复制附件失败: {e}")
            return None


class AttachmentListWidget(QListWidget):
    """附件列表控件，显示附件列表并提供操作功能"""
    
    attachment_double_clicked = pyqtSignal(str)
    
    def __init__(self, attachment_handler, parent=None):
        """初始化附件列表控件
        
        Args:
            attachment_handler: 附件处理器实例
            parent: 父部件
        """
        super().__init__(parent)
        
        self.attachment_handler = attachment_handler
        self.attachments = []
        
        # 设置右键菜单
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu)
        
        # 设置双击事件
        self.itemDoubleClicked.connect(self.handle_double_click)
        
    def set_attachments(self, attachments):
        """设置附件列表
        
        Args:
            attachments: 附件路径列表
        """
        self.clear()
        self.attachments = attachments
        
        # 创建通用图标映射（如果文件不存在，使用默认图标）
        icon_map = {
            "image": QIcon("icons/image.png") if os.path.exists("icons/image.png") else QIcon(),
            "text": QIcon("icons/text.png") if os.path.exists("icons/text.png") else QIcon(),
            "audio": QIcon("icons/audio.png") if os.path.exists("icons/audio.png") else QIcon(),
            "video": QIcon("icons/video.png") if os.path.exists("icons/video.png") else QIcon(),
            "pdf": QIcon("icons/pdf.png") if os.path.exists("icons/pdf.png") else QIcon(),
            "excel": QIcon("icons/excel.png") if os.path.exists("icons/excel.png") else QIcon(),
            "word": QIcon("icons/word.png") if os.path.exists("icons/word.png") else QIcon(),
            "powerpoint": QIcon("icons/powerpoint.png") if os.path.exists("icons/powerpoint.png") else QIcon(),
            "archive": QIcon("icons/archive.png") if os.path.exists("icons/archive.png") else QIcon(),
            "file": QIcon("icons/file.png") if os.path.exists("icons/file.png") else QIcon()
        }
        
        for attachment_path in attachments:
            filename = os.path.basename(attachment_path)
            item = QListWidgetItem(filename)
            
            # 设置图标
            icon_type = self.attachment_handler.get_file_icon(attachment_path)
            item.setIcon(icon_map.get(icon_type, icon_map["file"]))
            
            # 存储完整路径
            item.setData(Qt.ItemDataRole.UserRole, attachment_path)
            
            self.addItem(item)
            
    def show_context_menu(self, position):
        """显示右键菜单
        
        Args:
            position: 鼠标位置
        """
        item = self.itemAt(position)
        if not item:
            return
            
        attachment_path = item.data(Qt.ItemDataRole.UserRole)
        
        menu = QMenu(self)
        
        # 打开操作
        open_action = QAction("打开", self)
        open_action.triggered.connect(lambda: self.open_attachment(attachment_path))
        menu.addAction(open_action)
        
        # 保存操作
        save_action = QAction("另存为...", self)
        save_action.triggered.connect(lambda: self.save_attachment(attachment_path))
        menu.addAction(save_action)
        
        # 重命名操作
        rename_action = QAction("重命名", self)
        rename_action.triggered.connect(lambda: self.rename_attachment(attachment_path, item))
        menu.addAction(rename_action)
        
        menu.exec(self.mapToGlobal(position))
        
    def handle_double_click(self, item):
        """处理双击事件"""
        attachment_path = item.data(Qt.ItemDataRole.UserRole)
        self.attachment_double_clicked.emit(attachment_path)
        self.open_attachment(attachment_path)
        
    def open_attachment(self, attachment_path):
        """打开附件
        
        Args:
            attachment_path: 附件路径
        """
        try:
            # 使用系统默认程序打开
            import os
            import platform
            
            if platform.system() == 'Darwin':  # macOS
                os.system(f'open "{attachment_path}"')
            elif platform.system() == 'Windows':  # Windows
                os.startfile(attachment_path)
            else:  # Linux
                os.system(f'xdg-open "{attachment_path}"')
        except Exception as e:
            QMessageBox.warning(self, "打开失败", f"无法打开附件: {e}")
            
    def save_attachment(self, attachment_path):
        """另存为附件
        
        Args:
            attachment_path: 附件路径
        """
        filename = os.path.basename(attachment_path)
        target_path, _ = QFileDialog.getSaveFileName(
            self, "保存附件", filename, "所有文件 (*.*)"
        )
        
        if not target_path:
            return
            
        try:
            shutil.copy2(attachment_path, target_path)
            QMessageBox.information(self, "保存成功", f"附件已保存至: {target_path}")
        except Exception as e:
            QMessageBox.warning(self, "保存失败", f"无法保存附件: {e}")
            
    def rename_attachment(self, attachment_path, item):
        """重命名附件
        
        Args:
            attachment_path: 附件路径
            item: 列表项
        """
        old_name = os.path.basename(attachment_path)
        new_name, ok = QInputDialog.getText(
            self, "重命名附件", "新文件名:", text=old_name
        )
        
        if not ok or not new_name or new_name == old_name:
            return
            
        dir_path = os.path.dirname(attachment_path)
        new_path = os.path.join(dir_path, new_name)
        
        try:
            os.rename(attachment_path, new_path)
            
            # 更新列表项
            item.setText(new_name)
            item.setData(Qt.ItemDataRole.UserRole, new_path)
            
            # 更新附件列表
            index = self.attachments.index(attachment_path)
            self.attachments[index] = new_path
            
        except Exception as e:
            QMessageBox.warning(self, "重命名失败", f"无法重命名附件: {e}")


class AttachmentWidget(QWidget):
    """附件管理控件"""
    
    def __init__(self, attachment_handler, parent=None):
        """初始化附件管理控件
        
        Args:
            attachment_handler: 附件处理器实例
            parent: 父部件
        """
        super().__init__(parent)
        
        self.attachment_handler = attachment_handler
        
        # 创建布局
        self.setup_ui()
        
    def setup_ui(self):
        """设置界面"""
        layout = QVBoxLayout(self)
        
        # 标题
        title_label = QLabel("附件管理")
        title_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        layout.addWidget(title_label)
        
        # 附件列表
        self.attachment_list = AttachmentListWidget(self.attachment_handler)
        layout.addWidget(self.attachment_list)
        
        # 按钮区域
        button_layout = QHBoxLayout()
        
        self.save_all_btn = QPushButton("全部保存")
        self.save_all_btn.clicked.connect(self.save_all_attachments)
        button_layout.addWidget(self.save_all_btn)
        
        self.add_btn = QPushButton("添加附件")
        self.add_btn.clicked.connect(self.add_attachment)
        button_layout.addWidget(self.add_btn)
        
        layout.addLayout(button_layout)
        
    def set_email(self, email_data):
        """设置当前邮件
        
        Args:
            email_data: 邮件数据
        """
        self.current_email = email_data
        
        # 获取附件
        if 'attachments' in email_data and email_data['attachments']:
            # 保存附件到磁盘
            paths = self.attachment_handler.save_all_attachments(email_data)
            self.attachment_list.set_attachments(paths)
            
            # 启用按钮
            self.save_all_btn.setEnabled(True)
        else:
            self.attachment_list.clear()
            self.save_all_btn.setEnabled(False)
            
    def save_all_attachments(self):
        """保存所有附件"""
        if not hasattr(self, 'current_email') or not self.current_email:
            return
            
        target_dir = QFileDialog.getExistingDirectory(
            self, "选择保存目录", ""
        )
        
        if not target_dir:
            return
            
        saved_count = 0
        for attachment_path in self.attachment_list.attachments:
            if self.attachment_handler.copy_to_directory(attachment_path, target_dir):
                saved_count += 1
                
        if saved_count > 0:
            QMessageBox.information(self, "保存成功", f"已成功保存 {saved_count} 个附件到 {target_dir}")
        else:
            QMessageBox.warning(self, "保存失败", "未能保存任何附件")
            
    def add_attachment(self):
        """添加附件"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "选择附件", "", "所有文件 (*.*)"
        )
        
        if not file_paths:
            return
            
        # TODO: 实现添加附件到邮件的功能
        QMessageBox.information(self, "功能提示", "附件添加功能暂未实现，请在回复邮件时手动添加附件。") 