import re
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
import nltk

nltk.download('punkt')
nltk.download('stopwords')

class EmailClassifier:
    def __init__(self, config):
        self.stop_words = set(stopwords.words('english'))
        self.category_keywords = config.CATEGORY_KEYWORDS
        self.default_category = config.DEFAULT_CATEGORY

    def preprocess_text(self, text):
        """文本预处理"""
        # 转小写
        text = text.lower()
        # 去除标点符号和特殊字符
        text = re.sub(r'[^\w\s]', '', text)
        # 分词
        tokens = word_tokenize(text)
        # 去除停用词
        filtered_tokens = [word for word in tokens if word not in self.stop_words]
        return ' '.join(filtered_tokens)

    def classify_email(self, email_data):
        """对邮件进行分类"""
        processed_subject = self.preprocess_text(email_data['subject'])
        processed_body = self.preprocess_text(email_data['body'])

        # 合并主题和正文进行分类
        combined_text = processed_subject + ' ' + processed_body

        # 基于关键词的分类
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