# 自动邮件分类与回复助手

![运行界面](image/1.png)

## 简介
自动邮件分类与回复助手是一个Python项目，旨在帮助用户自动对邮件进行分类并生成相应的回复。该项目通过连接到指定的邮箱，获取邮件，对邮件进行分类，并根据邮件的类别生成自动回复。同时，项目还提供了一个图形用户界面（GUI），方便用户进行操作。

## 功能特点
- **邮件分类**：基于关键词匹配的规则对邮件进行分类，支持多种预设分类。
- **自动回复**：根据邮件的分类生成相应的自动回复内容。
- **批量处理**：支持对多封邮件进行批量分类和回复。
- **图形界面**：提供直观的图形用户界面，方便用户操作。

## 项目结构
```
├── auto_reply.py         # 自动回复生成模块
├── config.py             # 配置文件，包含邮箱和分类相关配置
├── email_classifier.py   # 邮件分类模块
├── email_connector.py    # 邮箱连接和邮件获取模块
├── email_sender.py       # 邮件发送模块
├── gui.py                # 图形用户界面模块
├── main.py               # 主程序入口
└── README.md             # 项目说明文档
```

## 安装与配置
### 安装依赖
确保你已经安装了Python 3.x，并使用以下命令安装所需的依赖库：
```bash
pip install -r requirements.txt
pip install nltk
```
同时，运行以下Python代码下载NLTK数据：
```python
import nltk
nltk.download('punkt')
nltk.download('stopwords')
```

### 配置邮箱信息
打开 `config.py` 文件，修改以下邮箱配置信息：
```python
# 邮箱配置
EMAIL_ADDRESS = "your_email@example.com"
EMAIL_PASSWORD = "your_email_password"
IMAP_SERVER = "imap.example.com"
IMAP_PORT = 993
SMTP_SERVER = "smtp.example.com"
SMTP_PORT = 25
```

### 配置分类和回复模板
在 `config.py` 文件中，可以修改 `CATEGORY_KEYWORDS` 和 `AUTO_REPLY_TEMPLATES` 来调整邮件分类的关键词和自动回复的模板。

## 使用方法
### 命令行界面
运行 `main.py` 文件，程序将自动连接到邮箱，获取邮件，进行分类并发送自动回复：
```bash
python main.py
```

### 图形用户界面
运行 `gui.py` 文件，将打开图形用户界面，你可以在界面上进行以下操作：
1. **连接邮箱**：输入邮箱地址、密码、IMAP和SMTP服务器信息，点击“连接邮箱”按钮。
2. **获取邮件列表**：点击“重新获取邮件列表”按钮，获取邮箱中的邮件列表。
3. **查看邮件内容**：在邮件列表中选择一封邮件，邮件内容将显示在右侧的文本框中，同时自动生成回复内容。
4. **发送回复**：点击“发送回复”按钮，将自动回复发送给发件人。
5. **批量处理**：在批量处理区域，可以选择目标分类对多封邮件进行批量分类，也可以输入回复内容对多封邮件进行批量回复。

## 注意事项
- 请确保你的邮箱账户已经开启了IMAP和SMTP服务，并且使用的密码是授权码或应用密码。
- 如果在运行过程中遇到连接失败或其他错误，请检查邮箱配置信息和网络连接。

## 贡献
如果你有任何建议或改进意见，欢迎提交Pull Request或提出Issue。