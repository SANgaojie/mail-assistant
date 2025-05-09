import imaplib
import email
from email.header import decode_header
import re

class EmailConnector:
    def __init__(self, config=None):
        # 如果没有提供config参数，则自动加载
        if config is None:
            from config import load_config
            config = load_config()
            
        self.email_address = config.EMAIL_ADDRESS
        self.password = config.EMAIL_PASSWORD
        self.imap_server = config.IMAP_SERVER
        self.imap_port = config.IMAP_PORT
        self.mail = None

    def connect(self):
        """连接到邮箱服务器"""
        try:
            # 创建IMAP4_SSL对象
            self.mail = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)

            # 登录邮箱
            self.mail.login(self.email_address, self.password)

            # 添加ID字段参数以满足网易邮箱的安全要求
            imaplib.Commands['ID'] = ('AUTH')
            args = ("name", "EmailAssistant", "contact", self.email_address, "version", "1.0.0", "vendor", "myclient")
            self.mail._simple_command('ID', '("' + '" "'.join(args) + '")')

            print("邮箱连接成功")
            return True
        except Exception as e:
            print(f"邮箱连接失败: {e}")
            return False

    def fetch_emails(self, folder='INBOX', search_criteria='ALL'):
        """获取邮件"""
        # 检查邮箱连接是否存在
        if not self.mail:
            print("邮箱未连接，尝试重新连接")
            if not self.connect():
                print("重新连接失败")
                return []
            
        try:
            # 尝试选择邮件文件夹
            try:
                status, messages = self.mail.select(folder)
                if status != "OK":
                    raise Exception(f"选择邮件文件夹失败: {messages}")
            except Exception as e:
                # SELECT命令失败，尝试重新连接并再次选择
                print(f"选择文件夹时出错: {e}，尝试重新连接...")
                self.close()  # 关闭旧连接
                if not self.connect():  # 重新连接
                    return []
                
                # 重新选择文件夹
                status, messages = self.mail.select(folder)
                if status != "OK":
                    print(f"重新选择邮件文件夹失败: {messages}")
                    return []

            # 搜索邮件
            try:
                status, email_ids = self.mail.search(None, search_criteria)
                if status != "OK":
                    raise Exception(f"搜索邮件失败: {email_ids}")
            except Exception as e:
                # SEARCH命令失败，尝试重新连接
                print(f"搜索邮件时出错: {e}，尝试重新连接...")
                self.close()
                if not self.connect():
                    return []
                
                # 重新选择文件夹和搜索
                status, messages = self.mail.select(folder)
                if status != "OK":
                    return []
                    
                status, email_ids = self.mail.search(None, search_criteria)
                if status != "OK":
                    print(f"重新搜索邮件失败: {email_ids}")
                    return []

            emails = []
            for mail_id in email_ids[0].split():
                try:
                    status, msg_data = self.mail.fetch(mail_id, '(RFC822)')
                    if status != "OK":
                        print(f"获取邮件 {mail_id} 失败")
                        continue

                    raw_email = msg_data[0][1]
                    email_message = email.message_from_bytes(raw_email)

                    # 解析邮件头部信息
                    subject, encoding = decode_header(email_message["Subject"])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding or 'utf-8')

                    # 解析发件人信息
                    from_ = email_message["From"]
                    if from_:
                        from_parts = decode_header(from_)
                        decoded_from = []
                        for part, enc in from_parts:
                            if isinstance(part, bytes):
                                part = part.decode(enc or 'utf-8')
                            decoded_from.append(part)
                        from_ = ''.join(decoded_from)
                        # 提取邮箱地址
                        match = re.search(r'<([^>]+)>', from_)
                        if match:
                            from_ = match.group(1)
                    else:
                        from_ = "unknown@example.com"  # 如果无法获取发件人信息，使用默认值

                    # 提取邮件正文
                    body = ""
                    if email_message.is_multipart():
                        for part in email_message.walk():
                            content_type = part.get_content_type()
                            content_disposition = str(part.get("Content-Disposition"))
                            charset = part.get_content_charset()
                            try:
                                payload = part.get_payload(decode=True)
                                if payload:
                                    if charset:
                                        body = payload.decode(charset)
                                    else:
                                        # 尝试多种编码
                                        encodings = ['utf-8', 'gbk', 'gb2312', 'big5']
                                        for encoding in encodings:
                                            try:
                                                body = payload.decode(encoding)
                                                break
                                            except UnicodeDecodeError:
                                                continue
                            except:
                                pass
                            if content_type == "text/plain" and "attachment" not in content_disposition:
                                break
                    else:
                        payload = email_message.get_payload(decode=True)
                        charset = email_message.get_content_charset()
                        if payload:
                            if charset:
                                body = payload.decode(charset)
                            else:
                                # 尝试多种编码
                                encodings = ['utf-8', 'gbk', 'gb2312', 'big5']
                                for encoding in encodings:
                                    try:
                                        body = payload.decode(encoding)
                                        break
                                    except UnicodeDecodeError:
                                        continue

                    # 创建邮件字典
                    email_dict = {
                        'id': mail_id.decode(),
                        'from': from_,
                        'to': email_message["To"],
                        'subject': subject,
                        'date': email_message["Date"],
                        'body': body,
                        'attachments': self._extract_attachments(email_message)
                    }

                    emails.append(email_dict)
                
                except Exception as e:
                    print(f"处理邮件 {mail_id} 时出错: {e}")
                    continue
                
            print(f"成功获取 {len(emails)} 封邮件")
            return emails

        except Exception as e:
            print(f"获取邮件时出错: {e}")
            # 出现异常，重置连接状态
            self.close()
            return []

    def _extract_attachments(self, email_message):
        """提取邮件附件"""
        attachments = []
        for part in email_message.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue

            filename = part.get_filename()
            if filename:
                attachment_data = part.get_payload(decode=True)
                attachments.append({'filename': filename, 'data': attachment_data})

        print(f"提取到 {len(attachments)} 个附件")
        return attachments

    def close(self):
        """关闭邮箱连接"""
        if self.mail:
            try:
                # 尝试注销，不再调用close()
                self.mail.logout()
                print("邮箱连接已关闭")
            except Exception as e:
                print(f"关闭邮箱连接时出错: {e}")
            finally:
                self.mail = None