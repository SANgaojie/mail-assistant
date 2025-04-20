import sys
from PyQt6.QtWidgets import QApplication
from gui_pyqt6 import EmailAssistantGUI
from email_connector import EmailConnector
from email_classifier import EmailClassifier
from email_sender import EmailSender
from async_operations import AsyncEmailProcessor
from config import load_config
from attachment_handler import AttachmentHandler
from template_manager import TemplateManager
from email_analytics import EmailAnalytics

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