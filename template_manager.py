import json
import os
import datetime

class TemplateManager:
    """模板管理系统"""
    
    def __init__(self, template_dir="templates"):
        """初始化模板管理器"""
        self.template_dir = template_dir
        
        # 确保模板目录存在
        if not os.path.exists(template_dir):
            os.makedirs(template_dir)
            
        # 模板文件路径
        self.default_template_file = os.path.join(template_dir, "default_templates.json")
        self.user_template_file = os.path.join(template_dir, "user_templates.json")
        
        # 加载模板
        self.templates = self.load_templates()
        
    def load_templates(self):
        """加载所有模板"""
        templates = {}
        
        # 加载默认模板
        if os.path.exists(self.default_template_file):
            try:
                with open(self.default_template_file, 'r', encoding='utf-8') as f:
                    templates.update(json.load(f))
            except Exception as e:
                print(f"加载默认模板失败: {e}")
        
        # 加载用户模板（用户模板优先）
        if os.path.exists(self.user_template_file):
            try:
                with open(self.user_template_file, 'r', encoding='utf-8') as f:
                    templates.update(json.load(f))
            except Exception as e:
                print(f"加载用户模板失败: {e}")
                
        return templates
    
    def save_template(self, name, content, category=None, is_default=False):
        """保存模板"""
        try:
            template_data = {
                "name": name,
                "content": content,
                "category": category,
                "created_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "updated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
            # 确定保存到哪个文件
            target_file = self.default_template_file if is_default else self.user_template_file
            
            # 加载现有模板
            existing_templates = {}
            if os.path.exists(target_file):
                with open(target_file, 'r', encoding='utf-8') as f:
                    existing_templates = json.load(f)
            
            # 更新模板
            existing_templates[name] = template_data
            
            # 保存模板
            with open(target_file, 'w', encoding='utf-8') as f:
                json.dump(existing_templates, f, ensure_ascii=False, indent=4)
            
            # 更新内存中的模板
            self.templates[name] = template_data
            
            return True
        except Exception as e:
            print(f"保存模板失败: {e}")
            return False
    
    def delete_template(self, name):
        """删除模板"""
        try:
            # 检查默认模板
            default_templates = {}
            if os.path.exists(self.default_template_file):
                with open(self.default_template_file, 'r', encoding='utf-8') as f:
                    default_templates = json.load(f)
            
            # 检查用户模板
            user_templates = {}
            if os.path.exists(self.user_template_file):
                with open(self.user_template_file, 'r', encoding='utf-8') as f:
                    user_templates = json.load(f)
            
            # 从相应文件中删除
            if name in default_templates:
                del default_templates[name]
                with open(self.default_template_file, 'w', encoding='utf-8') as f:
                    json.dump(default_templates, f, ensure_ascii=False, indent=4)
            
            if name in user_templates:
                del user_templates[name]
                with open(self.user_template_file, 'w', encoding='utf-8') as f:
                    json.dump(user_templates, f, ensure_ascii=False, indent=4)
            
            # 从内存中删除
            if name in self.templates:
                del self.templates[name]
                
            return True
        except Exception as e:
            print(f"删除模板失败: {e}")
            return False
    
    def get_template(self, name):
        """获取指定名称的模板"""
        return self.templates.get(name)
    
    def get_template_content(self, name):
        """获取指定名称的模板内容"""
        template = self.get_template(name)
        return template["content"] if template else ""
    
    def get_templates_by_category(self, category=None):
        """获取指定分类的所有模板"""
        if category:
            return [t for t in self.templates.values() if t.get("category") == category]
        return list(self.templates.values())
    
    def get_all_categories(self):
        """获取所有分类"""
        categories = set()
        for template in self.templates.values():
            if "category" in template and template["category"]:
                categories.add(template["category"])
        return list(categories)
    
    def import_from_config(self, config):
        """从配置文件导入模板"""
        count = 0
        if hasattr(config, "AUTO_REPLY_TEMPLATES"):
            for category, content in config.AUTO_REPLY_TEMPLATES.items():
                name = f"默认_{category}"
                self.save_template(name, content, category, is_default=True)
                count += 1
        return count
    
    def fill_template(self, template_content, email_data):
        """填充模板变量"""
        # 基本变量替换
        filled = template_content
        filled = filled.replace("{sender}", email_data.get("from", ""))
        filled = filled.replace("{subject}", email_data.get("subject", ""))
        filled = filled.replace("{date}", email_data.get("date", ""))
        
        return filled 