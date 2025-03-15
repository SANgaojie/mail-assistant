import config

class AutoReplyGenerator:
    def __init__(self, config):
        self.config = config

    def generate_reply(self, category):
        """根据类别生成回复"""
        return self.config.AUTO_REPLY_TEMPLATES.get(category, self.config.AUTO_REPLY_TEMPLATES['其他'])

    def generate_bulk_reply(self, category_list):
        """根据类别列表生成批量回复"""
        reply_list = []
        for category in category_list:
            reply = self.config.AUTO_REPLY_TEMPLATES.get(category, self.config.AUTO_REPLY_TEMPLATES['其他'])
            reply_list.append(reply)
        return reply_list