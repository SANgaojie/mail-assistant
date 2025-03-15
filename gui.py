import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import email
from email.header import decode_header
import re
import config
from email_connector import EmailConnector
from email_classifier import EmailClassifier
from auto_reply import AutoReplyGenerator
from email_sender import EmailSender


class EmailAssistantGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("自动邮件分类与回复助手")
        self.root.geometry("1200x700")

        # 配置
        self.config = config

        # 初始化邮件相关对象
        self.email_connector = EmailConnector(self.config)
        self.email_classifier = EmailClassifier(self.config)
        self.auto_reply_generator = AutoReplyGenerator(self.config)
        self.email_sender = EmailSender(self.config)
        self.bulk_category = tk.StringVar()
        self.create_widgets()
        self.emails_data = {}

    def create_widgets(self):
        # 整体布局框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 邮箱配置框架
        config_frame = ttk.LabelFrame(main_frame, text="邮箱配置")
        config_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

        # 邮箱地址
        ttk.Label(config_frame, text="邮箱地址:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.email_var = tk.StringVar(value=self.config.EMAIL_ADDRESS)
        ttk.Entry(config_frame, textvariable=self.email_var, width=30).grid(row=0, column=1, padx=5, pady=5)

        # 密码
        ttk.Label(config_frame, text="密码:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.password_var = tk.StringVar(value=self.config.EMAIL_PASSWORD)
        ttk.Entry(config_frame, textvariable=self.password_var, show="*", width=30).grid(row=1, column=1, padx=5, pady=5)

        # IMAP 服务器
        ttk.Label(config_frame, text="IMAP 服务器:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.imap_server_var = tk.StringVar(value=self.config.IMAP_SERVER)
        ttk.Entry(config_frame, textvariable=self.imap_server_var, width=30).grid(row=0, column=3, padx=5, pady=5)

        # IMAP 端口
        ttk.Label(config_frame, text="IMAP 端口:").grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.imap_port_var = tk.StringVar(value=str(self.config.IMAP_PORT))
        ttk.Entry(config_frame, textvariable=self.imap_port_var, width=30).grid(row=1, column=3, padx=5, pady=5)

        # SMTP 服务器
        ttk.Label(config_frame, text="SMTP 服务器:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.smtp_server_var = tk.StringVar(value=self.config.SMTP_SERVER)
        ttk.Entry(config_frame, textvariable=self.smtp_server_var, width=30).grid(row=2, column=1, padx=5, pady=5)

        # SMTP 端口
        ttk.Label(config_frame, text="SMTP 端口:").grid(row=2, column=2, padx=5, pady=5, sticky="w")
        self.smtp_port_var = tk.StringVar(value=str(self.config.SMTP_PORT))
        ttk.Entry(config_frame, textvariable=self.smtp_port_var, width=30).grid(row=2, column=3, padx=5, pady=5)

        btn_frame = ttk.Frame(config_frame)
        btn_frame.grid(row=3, column=0, columnspan=4, padx=5, pady=10)
        ttk.Button(btn_frame, text="连接邮箱", command=self.connect_email).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="重新获取邮件列表", command=self.fetch_and_display_emails).pack(side=tk.LEFT, padx=5)

        # 邮件列表与内容区域框架
        middle_frame = ttk.Frame(main_frame)
        middle_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
        middle_frame.grid_columnconfigure(0, weight=1)
        middle_frame.grid_columnconfigure(1, weight=1)

        # 邮件列表框架
        mailbox_frame = ttk.LabelFrame(middle_frame, text="邮件列表")
        mailbox_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        columns = ("ID", "发件人", "主题", "日期", "分类")
        self.mail_tree = ttk.Treeview(mailbox_frame, columns=columns, show="headings", selectmode='extended')
        for col in columns:
            self.mail_tree.heading(col, text=col)
            self.mail_tree.column("ID", width=50)
            self.mail_tree.column("发件人", width=150)
            self.mail_tree.column("主题", width=200)
            self.mail_tree.column("日期", width=120)
            self.mail_tree.column("分类", width=100)
        self.mail_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(mailbox_frame, orient=tk.VERTICAL, command=self.mail_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.mail_tree.configure(yscrollcommand=scrollbar.set)

        # 邮件内容框架
        content_frame = ttk.LabelFrame(middle_frame, text="邮件内容")
        content_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

        self.email_body_text = scrolledtext.ScrolledText(content_frame, width=60, height=30)
        self.email_body_text.pack(fill=tk.BOTH, expand=True)

        # 操作区域框架
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
        action_frame.grid_columnconfigure(0, weight=1)
        action_frame.grid_columnconfigure(1, weight=1)

        # 自动回复框架
        reply_frame = ttk.LabelFrame(action_frame, text="自动回复")
        reply_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        ttk.Label(reply_frame, text="自动回复内容:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.reply_text = scrolledtext.ScrolledText(reply_frame, width=60, height=8)
        self.reply_text.grid(row=1, column=0, padx=5, pady=5)
        ttk.Button(reply_frame, text="发送回复", command=self.send_reply).grid(row=2, column=0, padx=5, pady=10)

        # 批量处理框架
        bulk_frame = ttk.LabelFrame(action_frame, text="批量处理")
        bulk_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

        # 批量分类部分
        ttk.Label(bulk_frame, text="目标分类:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        category_dropdown = ttk.Combobox(bulk_frame, textvariable=self.bulk_category, width=20)
        category_dropdown['values'] = list(self.config.CATEGORY_KEYWORDS.keys()) + [self.config.DEFAULT_CATEGORY]
        category_dropdown.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(bulk_frame, text="批量分类", command=self.bulk_classify_emails).grid(row=0, column=2, padx=5, pady=5)

        # 批量回复部分
        ttk.Label(bulk_frame, text="回复内容:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.bulk_reply_text = scrolledtext.ScrolledText(bulk_frame, width=60, height=4)
        self.bulk_reply_text.grid(row=1, column=1, columnspan=2, padx=5, pady=5)
        ttk.Button(bulk_frame, text="批量回复", command=self.bulk_reply_emails).grid(row=2, column=0, columnspan=3, padx=5, pady=10)

        # 绑定邮件选择事件
        self.mail_tree.bind("<<TreeviewSelect>>", self.show_email_content)

        # 设置整体权重
        main_frame.grid_rowconfigure(1, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

    def connect_email(self):
        """连接到邮箱"""
        try:
            # 获取 GUI 界面上输入的邮箱地址和密码
            email_address = self.email_var.get()
            email_password = self.password_var.get()
            imap_server = self.imap_server_var.get()
            imap_port = int(self.imap_port_var.get())
            smtp_server = self.smtp_server_var.get()
            smtp_port = int(self.smtp_port_var.get())

            # 更新配置
            self.config.EMAIL_ADDRESS = email_address
            self.config.EMAIL_PASSWORD = email_password
            self.config.IMAP_SERVER = imap_server
            self.config.IMAP_PORT = imap_port
            self.config.SMTP_SERVER = smtp_server
            self.config.SMTP_PORT = smtp_port

            # 重新初始化 EmailConnector 和 EmailSender
            self.email_connector = EmailConnector(self.config)
            self.email_sender = EmailSender(self.config)

            if self.email_connector.connect():
                messagebox.showinfo("成功", "邮箱连接成功")
                self.fetch_and_display_emails()
            else:
                messagebox.showerror("错误", "邮箱连接失败")
        except Exception as e:
            messagebox.showerror("错误", f"邮箱连接失败: {e}")

    def fetch_and_display_emails(self):
        """获取并显示邮件"""
        try:
            emails = self.email_connector.fetch_emails()
            # 对邮件列表进行倒序排列
            emails.reverse()

            # 清空树视图
            for item in self.mail_tree.get_children():
                self.mail_tree.delete(item)

            # 插入邮件数据
            self.emails_data = {}
            for email_data in emails:
                item = self.mail_tree.insert("", tk.END, values=(
                    email_data['id'],
                    email_data['from'],
                    email_data['subject'],
                    email_data['date'],
                    email_data.get('category', '未分类')
                ))
                self.emails_data[item] = email_data

        except Exception as e:
            messagebox.showerror("错误", f"获取邮件时出错: {e}")

    def show_email_content(self, event):
        """显示选中邮件的内容"""
        selected_items = self.mail_tree.selection()
        if not selected_items:
            return

        item = selected_items[0]
        # 查找对应的邮件数据
        if item not in self.emails_data:
            messagebox.showerror("错误", "未找到选中邮件的数据")
            return

        email_data = self.emails_data[item]

        # 显示邮件正文
        self.email_body_text.delete(1.0, tk.END)
        self.email_body_text.insert(tk.END, email_data.get('body', ''))

        # 分类邮件并生成回复
        category = self.email_classifier.classify_email(email_data)
        reply_content = self.auto_reply_generator.generate_reply(category)

        # 显示自动回复内容
        self.reply_text.delete(1.0, tk.END)
        self.reply_text.insert(tk.END, reply_content)

    def send_reply(self):
        """发送回复邮件"""
        selected_items = self.mail_tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请选择要回复的邮件")
            return

        item = selected_items[0]
        # 查找对应的邮件数据
        if item not in self.emails_data:
            messagebox.showerror("错误", "未找到选中邮件的数据")
            return

        email_data = self.emails_data[item]

        # 获取回复内容
        reply_content = self.reply_text.get(1.0, tk.END).strip()
        print(f"Recipient: {email_data['from']}")
        print(f"Subject: Re: {email_data['subject']}")
        print(f"Reply Content: {reply_content}")

        # 发送邮件
        if self.email_sender.send_email(
            recipient=email_data['from'],
            subject=f"Re: {email_data['subject']}",
            body=reply_content,
            attachments=email_data['attachments']
        ):
            messagebox.showinfo("成功", "邮件发送成功")
        else:
            messagebox.showerror("错误", "发送邮件时出错")

    def bulk_classify_emails(self):
        selected_items = self.mail_tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请选择要分类的邮件")
            return

        target_category = self.bulk_category.get()
        if not target_category:
            messagebox.showwarning("警告", "请选择目标分类")
            return

        success_count = 0
        for item in selected_items:
            if item not in self.emails_data:
                continue

            email_data = self.emails_data[item]
            # 更新邮件分类
            email_data = self.email_classifier.tag_email(email_data, target_category)
            # 更新显示
            self.mail_tree.item(item, values=(
                email_data['id'],
                email_data['from'],
                email_data['subject'],
                email_data['date'],
                target_category
            ))
            success_count += 1

        messagebox.showinfo("成功", f"成功分类 {success_count} 封邮件")

    def bulk_reply_emails(self):
        selected_items = self.mail_tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请选择要回复的邮件")
            return

        reply_content = self.bulk_reply_text.get(1.0, tk.END).strip()
        if not reply_content:
            messagebox.showwarning("警告", "请输入回复内容")
            return

        success_count = 0
        for item in selected_items:
            if item not in self.emails_data:
                continue

            email_data = self.emails_data[item]
            if self.email_sender.send_email(
                recipient=email_data['from'],
                subject=f"Re: {email_data['subject']}",
                body=reply_content,
                attachments=email_data['attachments']
            ):
                success_count += 1
            else:
                print(f"发送邮件失败: {email_data['subject']}")

        messagebox.showinfo("成功", f"成功发送 {success_count} 封回复邮件")


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailAssistantGUI(root)
    root.mainloop()