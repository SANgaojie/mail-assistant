import re
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
import nltk

nltk.download('punkt')
nltk.download('stopwords')

class EmailClassifier:
    def __init__(self, config=None):
        # 自动加载配置
        if config is None:
            from config import load_config
            config = load_config()
            
        self.stop_words = set(stopwords.words('english'))
        self.category_keywords = config.CATEGORY_KEYWORDS
        self.default_category = config.DEFAULT_CATEGORY

    def preprocess_text(self, text):
        """文本预处理"""
        text = text.lower()
        text = re.sub(r'[^\w\s]', '', text)
        tokens = word_tokenize(text)
        filtered_tokens = [word for word in tokens if word not in self.stop_words]
        return ' '.join(filtered_tokens)

    def classify_email(self, email_data):
        """对邮件进行分类"""
        processed_subject = self.preprocess_text(email_data['subject'])
        processed_body = self.preprocess_text(email_data['body'])

        # 合并主题和正文
        combined_text = processed_subject + ' ' + processed_body

        # 基于关键词分类
        category = self._rule_based_classification(combined_text)
        return category

    def _rule_based_classification(self, text):
        """基于规则的分类"""
        for category, keywords in self.category_keywords.items():
            for keyword in keywords:
                if keyword in text:
                    return category
        return self.default_category

    def tag_email(self, email_data, category):
        """为邮件添加标签"""
        email_data['category'] = category
        return email_data