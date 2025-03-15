import tkinter as tk  # 确保导入 tkinter
from email_connector import EmailConnector
from email_classifier import EmailClassifier
from auto_reply import AutoReplyGenerator
from email_sender import EmailSender
import config
from gui import EmailAssistantGUI  # 导入GUI模块

class EmailAssistant:
    def __init__(self):
        self.config = config
        self.email_connector = EmailConnector(self.config)
        self.email_classifier = EmailClassifier(self.config)
        self.auto_reply_generator = AutoReplyGenerator(self.config)
        self.email_sender = EmailSender(self.config)

    def run_cli(self):
        """运行命令行界面"""
        # 连接到邮箱
        if not self.email_connector.connect():
            return

        # 获取邮件
        emails = self.email_connector.fetch_emails()

        # 处理每封邮件
        for email_data in emails:
            print(f"处理邮件: {email_data['subject']}")

            # 分类邮件
            category = self.email_classifier.classify_email(email_data)
            tagged_email = self.email_classifier.tag_email(email_data, category)

            # 生成自动回复建议
            reply_content = self.auto_reply_generator.generate_reply(category)

            # 发送自动回复邮件
            self.email_sender.send_email(
                recipient=email_data['from'],
                subject=f"Re: {email_data['subject']}",
                body=reply_content,
                attachments=email_data['attachments']
            )

            print(f"邮件处理完成: {email_data['subject']}")
    
    # 批量处理方法
    def bulk_process_emails(self, emails, target_category=None, reply_content=None):
        processed_emails = []
        for email_data in emails:
            if target_category:
                email_data = self.email_classifier.tag_email(email_data, target_category)
        
            if reply_content:
                self.email_sender.send_email(
                    recipient=email_data['from'],
                    subject=f"Re: {email_data['subject']}",
                    body=reply_content,
                    attachments=email_data['attachments']
                )
        
            processed_emails.append(email_data)
        return processed_emails

    def close(self):
        """关闭邮箱连接"""
        if self.email_connector.mail:
            self.email_connector.mail.logout()
            print("邮箱连接已关闭")

if __name__ == "__main__":
    # 创建并运行邮件助手
    print("创建邮件助手实例")
    assistant = EmailAssistant()

    # 运行图形界面
    print("启动图形界面")
    root = tk.Tk()  # 确保使用导入的 tkinter
    app = EmailAssistantGUI(root)
    root.mainloop()

    # 关闭邮箱连接
    assistant.close()