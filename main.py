import sys
import os
from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QFont, QFontDatabase
from PyQt6.QtCore import QDir, QCoreApplication, Qt, QStandardPaths
from gui_pyqt6 import EmailAssistantGUI
from email_connector import EmailConnector
from email_classifier import EmailClassifier
from email_sender import EmailSender
from async_operations import AsyncEmailProcessor
from config import load_config
from attachment_handler import AttachmentHandler
from template_manager import TemplateManager
from email_analytics import EmailAnalytics

def force_chinese_font_support():
    """确保中文字体支持"""
    # 获取系统字体目录，使用正确的PyQt6枚举名
    # 在PyQt6中，FontsPath改为了Fonts
    try:
        font_dirs = QStandardPaths.standardLocations(QStandardPaths.StandardLocation.FontsLocation)
    except:
        try:
            font_dirs = QStandardPaths.standardLocations(QStandardPaths.StandardLocation.Fonts)
        except:
            # 如果都失败，使用一个简单的备用方案
            font_dirs = []
            if sys.platform == 'win32':
                font_dirs.append(os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Fonts'))
            elif sys.platform == 'darwin':  # macOS
                font_dirs.append('/System/Library/Fonts')
                font_dirs.append('/Library/Fonts')
            else:  # Linux和其他系统
                font_dirs.append('/usr/share/fonts')
                font_dirs.append('/usr/local/share/fonts')
    
    print(f"系统字体目录: {font_dirs}")
    
    # 不再尝试设置高DPI属性，因为在PyQt6中这些属性已被重组
    
    # 设置全局字体回退选项
    font = QFont()
    # 在PyQt6中setFamilies可能不存在，改用setFamily
    font.setFamily("Microsoft YaHei UI, Microsoft YaHei, SimSun, WenQuanYi Micro Hei, Arial")
    QApplication.setFont(font)
    
    # 设置全局样式表确保所有控件都使用中文字体
    QApplication.instance().setStyleSheet("""
        * {
            font-family: "Microsoft YaHei UI", "Microsoft YaHei", "SimSun", "WenQuanYi Micro Hei", sans-serif;
        }
    """)
    
    return True

def main():
    # 加载配置
    config = load_config()
    
    # 初始化各组件
    email_connector = EmailConnector()
    email_classifier = EmailClassifier()
    email_sender = EmailSender()
    attachment_handler = AttachmentHandler()
    template_manager = TemplateManager()
    email_analytics = EmailAnalytics()
    
    # 从配置导入默认模板
    template_manager.import_from_config(config)
    
    # 初始化异步处理器
    async_processor = AsyncEmailProcessor(
        email_connector=email_connector,
        email_classifier=email_classifier,
        analytics=email_analytics
    )
    
    # 初始化GUI
    app = QApplication(sys.argv)
    
    # 强制启用中文字体支持
    force_chinese_font_support()
    
    # 设置应用字体，解决中文显示问题
    # 首先尝试添加本地字体文件（如果有的话）
    local_font_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")
    if os.path.exists(local_font_dir):
        for font_file in os.listdir(local_font_dir):
            if font_file.endswith(('.ttf', '.otf')):
                font_path = os.path.join(local_font_dir, font_file)
                QFontDatabase.addApplicationFont(font_path)
                print(f"已添加字体文件: {font_file}")
    
    # 获取系统可用字体
    available_fonts = QFontDatabase.families()
    
    # 中文字体优先级列表（从最优先到最不优先）
    chinese_fonts = [
        "Microsoft YaHei UI", "Microsoft YaHei",  # Windows优选字体
        "SimSun", "SimHei", "FangSong", "KaiTi",  # Windows其他中文字体
        "WenQuanYi Micro Hei", "WenQuanYi Zen Hei",  # Linux常见中文字体
        "Noto Sans CJK SC", "Noto Sans SC", "Noto Sans CJK TC",  # Google Noto字体
        "Source Han Sans CN", "Source Han Sans SC", "Source Han Sans TC",  # Adobe思源字体
        "PingFang SC", "PingFang TC", "Hiragino Sans GB",  # macOS中文字体
        "Heiti SC", "Heiti TC", "STHeiti"  # 其他中文字体
    ]
    
    # 回退字体列表（如果找不到中文字体）
    fallback_fonts = ["Arial", "Helvetica", "sans-serif"]
    
    # 尝试找到可用的中文字体
    found_font = None
    for font_name in chinese_fonts:
        if font_name in available_fonts:
            found_font = font_name
            print(f"已找到中文字体: {font_name}")
            break
    
    if found_font:
        # 设置全局字体
        font = QFont(found_font, 9)
        app.setFont(font)
        
        print(f"已设置应用字体为: {found_font}")
    else:
        print("警告: 未找到合适的中文字体，将使用系统默认字体")
        # 即使没有找到理想字体，也设置font-family以便CSS可以回退
        app.setStyleSheet("* { font-family: sans-serif; }")
    
    # 创建GUI
    gui = EmailAssistantGUI(
        email_connector=email_connector,
        email_classifier=email_classifier,
        email_sender=email_sender,
        attachment_handler=attachment_handler,
        template_manager=template_manager,
        email_analytics=email_analytics,
        async_processor=async_processor
    )
    gui.show()
    
    # 运行应用
    exit_code = app.exec()
    
    # 清理
    async_processor.close()
    
    sys.exit(exit_code)

if __name__ == "__main__":
    main()