# 邮箱配置
EMAIL_ADDRESS = "your_email@example.com"
EMAIL_PASSWORD = "your_email_password"
IMAP_SERVER = "imap.example.com"
IMAP_PORT = 993
SMTP_SERVER = "smtp.example.com"
SMTP_PORT = 25

# 分类配置
DEFAULT_CATEGORY = "其他"
CATEGORY_KEYWORDS = {
    "账单": ["invoice", "payment", "fee"],
    "支付": ["payment", "refund", "charge"],
    "订单": ["order", "purchase", "delivery"],
    "投诉": ["complaint", "issue", "problem"],
    "反馈": ["feedback", "suggestion", "review"],
    "支持": ["support", "help", "assistance"],
    "咨询": ["question", "inquiry", "consultation"],
    "会议": ["meeting", "appointment", "schedule"],
    "提醒": ["reminder", "notice", "alert"],
    "通知": ["notification", "announcement", "update"]
}

# 自动回复模板
AUTO_REPLY_TEMPLATES = {
    "账单": "尊敬的用户，感谢您的邮件。我们已收到您的账单相关问题，我们的财务部门将在1-2个工作日内处理并回复您。感谢您的耐心等待。",
    "支付": "尊敬的用户，感谢您的邮件。我们已收到您的支付相关问题，我们的支付团队将在24小时内处理并回复您。感谢您的耐心等待。",
    "订单": "尊敬的用户，感谢您的邮件。我们已收到您的订单相关问题，我们的订单处理团队将在1个工作日内处理并回复您。感谢您的耐心等待。",
    "投诉": "尊敬的用户，感谢您的反馈。我们非常重视您的投诉，我们的客户服务团队将在24小时内与您联系解决相关问题。感谢您的支持。",
    "反馈": "尊敬的用户，感谢您的反馈。我们会将您的意见传达给相关部门，持续改进我们的产品和服务。感谢您的支持。",
    "支持": "尊敬的用户，感谢您的邮件。我们的技术支持团队已收到您的请求，将在1个工作日内为您提供解决方案。感谢您的耐心等待。",
    "咨询": "尊敬的用户，感谢您的咨询。我们已收到您的问题，相关专业人员将在24小时内回复您。感谢您的关注。",
    "会议": "尊敬的用户，感谢您的邮件。我们已收到您的会议相关请求，相关安排将在1个工作日内确认并回复您。感谢您的耐心等待。",
    "提醒": "尊敬的用户，感谢您的提醒。我们会及时处理相关事项，并在必要时与您联系。感谢您的支持。",
    "通知": "尊敬的用户，感谢您的邮件。我们已收到您的通知，将按照相关要求进行处理。感谢您的告知。",
    "其他": "尊敬的用户，感谢您的邮件。我们会尽快处理您的请求，并在适当的时候回复您。感谢您的耐心等待。"
}

def load_config():
    """加载配置信息
    
    此函数返回当前模块本身，以便其他模块可以访问配置变量。
    
    Returns:
        module: 包含所有配置变量的模块对象
    """
    import sys
    return sys.modules[__name__]