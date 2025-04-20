import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

class EmailSender:
    def __init__(self, config=None):
        # 如果没有提供config参数，则自动加载
        if config is None:
            from config import load_config
            config = load_config()
            
        self.config = config

    def send_email(self, recipient, subject, body, attachments=None):
        """发送邮件"""
        try:
            print(f"尝试连接到 SMTP 服务器: {self.config.SMTP_SERVER}:{self.config.SMTP_PORT}")
            # 创建 SMTP 对象并设置超时时间
            server = smtplib.SMTP(self.config.SMTP_SERVER, self.config.SMTP_PORT, timeout=10)
            # 发送 EHLO 命令
            server.ehlo()
            print("已发送 EHLO 命令")
            # 启动 TLS 加密
            server.starttls()
            print("已启动 TLS 加密")
            # 再次发送 EHLO 命令
            server.ehlo()
            print("再次发送 EHLO 命令")
            # 登录邮箱
            server.login(self.config.EMAIL_ADDRESS, self.config.EMAIL_PASSWORD)
            print("已登录到邮箱账号")

            # 创建邮件消息
            msg = MIMEMultipart()
            msg['From'] = self.config.EMAIL_ADDRESS
            msg['To'] = recipient
            msg['Subject'] = subject

            # 添加邮件正文
            msg.attach(MIMEText(body, 'plain'))

            # 添加附件
            if attachments:
                for attachment in attachments:
                    part = MIMEApplication(attachment['data'], Name=attachment['filename'])
                    part['Content-Disposition'] = f'attachment; filename="{attachment["filename"]}"'
                    msg.attach(part)

            # 发送邮件
            server.send_message(msg)
            server.quit()

            return True
        except Exception as e:
            print(f"发送邮件时出错: {e}")
            return False