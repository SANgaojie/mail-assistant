import os
import json
import datetime
import sqlite3
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTabWidget, QComboBox, QFrame
from PyQt6.QtCore import Qt
import matplotlib

# 配置matplotlib支持中文显示
def configure_matplotlib_chinese():
    """配置matplotlib支持中文字体"""
    # 尝试加载中文字体
    chinese_fonts = [
        'Microsoft YaHei', 'SimHei', 'SimSun', 'NSimSun', 'FangSong', 'KaiTi',  # Windows
        'WenQuanYi Micro Hei', 'WenQuanYi Zen Hei',  # Linux
        'PingFang SC', 'STHeiti', 'Heiti SC',  # macOS
        'Noto Sans CJK SC', 'Source Han Sans CN',  # 跨平台开源字体
    ]
    
    # 检查可用的字体
    available_fonts = matplotlib.font_manager.findSystemFonts(fontpaths=None)
    chinese_font_found = False
    
    for font_path in available_fonts:
        try:
            font = matplotlib.font_manager.FontProperties(fname=font_path)
            if font.get_name() in chinese_fonts:
                matplotlib.rcParams['font.family'] = font.get_name()
                chinese_font_found = True
                break
        except:
            continue
    
    # 设置matplotlib字体
    if not chinese_font_found:
        plt.rcParams['font.sans-serif'] = chinese_fonts
    plt.rcParams['axes.unicode_minus'] = False  # 正确显示负号
        
    return chinese_font_found

# 为不支持中文的环境提供备用方案
def safe_decode(text):
    return text.encode('utf-8').decode('utf-8', 'ignore')

# 初始化时配置字体
configure_matplotlib_chinese()

class EmailAnalytics:
    """邮件分析统计类"""
    
    def __init__(self, db_path="email_data.db"):
        """初始化分析器
        
        Args:
            db_path: 数据库路径
        """
        self.db_path = db_path
        self.initialize_db()
        
    def initialize_db(self):
        """初始化数据库"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 创建邮件表
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS emails (
            id TEXT PRIMARY KEY,
            sender TEXT,
            subject TEXT,
            body TEXT,
            date TEXT,
            category TEXT,
            is_replied INTEGER DEFAULT 0,
            reply_content TEXT,
            reply_date TEXT,
            response_time REAL,
            created_at TEXT
        )
        ''')
        
        # 创建统计表
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS statistics (
            date TEXT,
            category TEXT,
            count INTEGER,
            reply_count INTEGER,
            avg_response_time REAL,
            PRIMARY KEY (date, category)
        )
        ''')
        
        conn.commit()
        conn.close()
        
    def save_email(self, email_data):
        """保存邮件数据到数据库
        
        Args:
            email_data: 邮件数据
            
        Returns:
            bool: 是否成功
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 提取必要字段
            email_id = email_data.get('id', '')
            sender = email_data.get('from', '')
            subject = email_data.get('subject', '')
            body = email_data.get('body', '')
            date = email_data.get('date', '')
            category = email_data.get('category', '未分类')
            
            # 检查是否已存在
            cursor.execute("SELECT id FROM emails WHERE id = ?", (email_id,))
            if cursor.fetchone():
                # 更新现有记录
                cursor.execute('''
                UPDATE emails SET 
                sender = ?,
                subject = ?,
                body = ?,
                date = ?,
                category = ?,
                is_replied = ?,
                reply_content = ?,
                reply_date = ?,
                response_time = ?
                WHERE id = ?
                ''', (
                    sender, subject, body, date, category, 
                    email_data.get('is_replied', 0),
                    email_data.get('reply_content', ''),
                    email_data.get('reply_date', ''),
                    email_data.get('response_time', 0),
                    email_id
                ))
            else:
                # 插入新记录
                cursor.execute('''
                INSERT INTO emails (id, sender, subject, body, date, category, is_replied, 
                                  reply_content, reply_date, response_time, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    email_id, sender, subject, body, date, category,
                    email_data.get('is_replied', 0),
                    email_data.get('reply_content', ''),
                    email_data.get('reply_date', ''),
                    email_data.get('response_time', 0),
                    datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                ))
            
            conn.commit()
            
            # 更新统计数据
            self.update_statistics()
            
            return True
        except Exception as e:
            print(f"保存邮件数据失败: {e}")
            return False
        finally:
            conn.close()
    
    def update_statistics(self):
        """更新统计数据表"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 获取今天的日期
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        
        # 获取所有分类
        cursor.execute("SELECT DISTINCT category FROM emails")
        categories = [row[0] for row in cursor.fetchall()]
        
        # 对每个分类更新今天的统计
        for category in categories:
            if not category:
                category = "未分类"
                
            # 查询今天的该分类的邮件数
            cursor.execute('''
            SELECT COUNT(*), SUM(is_replied), AVG(response_time)
            FROM emails 
            WHERE category = ? AND date(created_at) = date(?)
            ''', (category, today))
            
            row = cursor.fetchone()
            count = row[0] or 0
            reply_count = row[1] or 0
            avg_response_time = row[2] or 0
            
            # 检查今天的该分类的记录是否存在
            cursor.execute('''
            SELECT count, reply_count, avg_response_time
            FROM statistics
            WHERE date = ? AND category = ?
            ''', (today, category))
            
            if cursor.fetchone():
                # 更新现有记录
                cursor.execute('''
                UPDATE statistics SET
                count = ?, reply_count = ?, avg_response_time = ?
                WHERE date = ? AND category = ?
                ''', (count, reply_count, avg_response_time, today, category))
            else:
                # 插入新记录
                cursor.execute('''
                INSERT INTO statistics (date, category, count, reply_count, avg_response_time)
                VALUES (?, ?, ?, ?, ?)
                ''', (today, category, count, reply_count, avg_response_time))
        
        conn.commit()
        conn.close()
    
    def get_email_categories(self, date_range=None):
        """获取邮件类别分布数据"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        if date_range:
            # 根据日期范围生成查询条件
            if date_range == 'today':
                date_limit = datetime.datetime.now().strftime("%Y-%m-%d")
                cursor.execute('''
                SELECT category, COUNT(*) as count
                FROM emails
                WHERE DATE(created_at) = ?
                GROUP BY category
                ''', (date_limit,))
            elif date_range == 'week':
                week_ago = (datetime.datetime.now() - datetime.timedelta(days=7)).strftime("%Y-%m-%d")
                cursor.execute('''
                SELECT category, COUNT(*) as count
                FROM emails
                WHERE DATE(created_at) >= ?
                GROUP BY category
                ''', (week_ago,))
            elif date_range == 'month':
                month_ago = (datetime.datetime.now() - datetime.timedelta(days=30)).strftime("%Y-%m-%d")
                cursor.execute('''
                SELECT category, COUNT(*) as count
                FROM emails
                WHERE DATE(created_at) >= ?
                GROUP BY category
                ''', (month_ago,))
            elif date_range == 'year':
                year_ago = (datetime.datetime.now() - datetime.timedelta(days=365)).strftime("%Y-%m-%d")
                cursor.execute('''
                SELECT category, COUNT(*) as count
                FROM emails
                WHERE DATE(created_at) >= ?
                GROUP BY category
                ''', (year_ago,))
        else:
            # 不限日期
            cursor.execute('''
            SELECT category, COUNT(*) as count
            FROM emails
            GROUP BY category
            ''')
            
        results = cursor.fetchall()
        
        # 处理结果
        categories = []
        counts = []
        
        # 确保有默认分类
        has_default = False
        
        # 处理空类别或None类别
        for cat, count in results:
            cat_name = cat if cat and cat.strip() else "未分类"
            
            # 合并相同类别
            existing_idx = next((i for i, c in enumerate(categories) if c == cat_name), None)
            if existing_idx is not None:
                counts[existing_idx] += count
            else:
                categories.append(cat_name)
                counts.append(count)
                
        # 确保至少有一个类别
        if not categories:
            # 查询总邮件数
            cursor.execute('SELECT COUNT(*) FROM emails')
            total_count = cursor.fetchone()[0] or 0
            
            categories = ["未分类"]
            counts = [total_count]
        
        conn.close()
        return categories, counts
    
    def get_email_trend(self, days=30, category=None):
        """获取邮件趋势数据"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 获取日期范围
        today = datetime.datetime.now()
        start_date = (today - datetime.timedelta(days=days-1)).strftime("%Y-%m-%d")
        
        # 生成日期列表
        date_list = []
        for i in range(days):
            date = (today - datetime.timedelta(days=days-1-i)).strftime("%Y-%m-%d")
            date_list.append(date)
        
        # 构建查询条件
        if category and category != "全部":
            # 使用 strftime 函数处理日期 - 更可靠的 SQL 方法
            cursor.execute('''
            SELECT date(created_at) as date, COUNT(*) as count
            FROM emails
            WHERE date(created_at) >= ? AND category = ?
            GROUP BY date(created_at)
            ORDER BY date(created_at)
            ''', (start_date, category))
        else:
            cursor.execute('''
            SELECT date(created_at) as date, COUNT(*) as count
            FROM emails
            WHERE date(created_at) >= ?
            GROUP BY date(created_at)
            ORDER BY date(created_at)
            ''', (start_date,))
        
        # 处理结果填充到日期列表
        results = cursor.fetchall()
        date_counts = {}
        
        for date_str, count in results:
            try:
                # 日期格式可能不同，打印日志
                date_counts[date_str] = count
            except Exception as e:
                print(f"日期处理出错: {date_str}, {e}")
        
        # 创建最终的数据集
        counts = []
        for date in date_list:
            counts.append(date_counts.get(date, 0))
        
        conn.close()
        return date_list, counts
    
    def get_response_data(self, days=30):
        """获取回复统计数据"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row  # 使用字典游标
        cursor = conn.cursor()
        
        # 今天的日期
        today = datetime.datetime.now()
        start_date = (today - datetime.timedelta(days=days-1)).strftime("%Y-%m-%d")
        
        # 查询回复数据
        cursor.execute('''
        SELECT 
            is_replied, 
            response_time
        FROM emails
        WHERE date(created_at) >= ?
        ''', (start_date,))
        
        results = cursor.fetchall()
        
        # 统计数据
        total_emails = len(results)
        replied_emails = sum(1 for row in results if row['is_replied'] == 1)
        reply_rate = replied_emails / total_emails if total_emails > 0 else 0
        
        # 计算平均响应时间（小时转分钟）
        response_times = [row['response_time'] * 24 * 60 for row in results if row['is_replied'] == 1 and row['response_time'] is not None]
        avg_response_time = sum(response_times) / len(response_times) if response_times else 0
        
        conn.close()
        
        return {
            'total_emails': total_emails,
            'replied_emails': replied_emails,
            'reply_rate': reply_rate,
            'avg_response_time': avg_response_time  # 分钟
        }
    
    def get_email_stats(self, date_range=None):
        """获取邮件统计数据"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        if date_range:
            # 根据日期范围生成查询条件
            if date_range == 'today':
                date_limit = datetime.datetime.now().strftime("%Y-%m-%d")
                cursor.execute('''
                SELECT COUNT(*) as count, SUM(is_replied) as replied
                FROM emails
                WHERE DATE(created_at) = ?
                ''', (date_limit,))
            elif date_range == 'week':
                week_ago = (datetime.datetime.now() - datetime.timedelta(days=7)).strftime("%Y-%m-%d")
                cursor.execute('''
                SELECT COUNT(*) as count, SUM(is_replied) as replied
                FROM emails
                WHERE DATE(created_at) >= ?
                ''', (week_ago,))
            elif date_range == 'month':
                month_ago = (datetime.datetime.now() - datetime.timedelta(days=30)).strftime("%Y-%m-%d")
                cursor.execute('''
                SELECT COUNT(*) as count, SUM(is_replied) as replied
                FROM emails
                WHERE DATE(created_at) >= ?
                ''', (month_ago,))
            elif date_range == 'year':
                year_ago = (datetime.datetime.now() - datetime.timedelta(days=365)).strftime("%Y-%m-%d")
                cursor.execute('''
                SELECT COUNT(*) as count, SUM(is_replied) as replied
                FROM emails
                WHERE DATE(created_at) >= ?
                ''', (year_ago,))
        else:
            # 不限日期
            cursor.execute('''
            SELECT COUNT(*) as count, SUM(is_replied) as replied
            FROM emails
            ''')
            
        row = cursor.fetchone()
        count = row[0] or 0
        replied = row[1] or 0
        
        conn.close()
        
        return {
            'count': count,
            'replied': replied,
            'reply_rate': replied / count if count > 0 else 0
        }
    
    def generate_category_pie(self, figure, date_range=None):
        """生成分类饼图"""
        # 确保使用中文字体
        configure_matplotlib_chinese()
        
        # 获取数据
        categories, sizes = self.get_email_categories(date_range)
        
        # 如果没有数据，确保至少显示"未分类"
        if not categories or sum(sizes) == 0:
            categories = ["无数据"]
            sizes = [1]
        
        # 创建饼图
        figure.clear()
        ax = figure.add_subplot(111)
        
        # 数据为0时显示一个空白的饼图
        if sum(sizes) == 0:  # 没有数据
            ax.text(0.5, 0.5, '无数据', horizontalalignment='center', verticalalignment='center', transform=ax.transAxes)
            return
        
        # 颜色映射
        colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd',
          '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']
        
        # 确保颜色足够
        while len(colors) < len(categories):
            colors.extend(colors)
        
        # 绘制饼图，添加颜色
        wedges, texts, autotexts = ax.pie(
            sizes, 
            labels=None, 
            autopct='%1.1f%%',
            startangle=90,
            colors=colors[:len(categories)]
        )
        
        # 设置字体大小和颜色
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontsize(9)
        
        ax.axis('equal')  # 使饼图为正圆形
        
        # 如果类别太多，添加图例
        if len(categories) > 0:
            ax.legend(wedges, categories, title="分类", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
            
        return figure
        
    def generate_trend_chart(self, figure, days=30, category=None):
        """生成每日趋势图"""
        # 确保使用中文字体
        configure_matplotlib_chinese()
        
        # 获取数据
        dates, counts = self.get_email_trend(days, category)
        
        # 创建折线图
        figure.clear()
        ax = figure.add_subplot(111)
        
        # 对日期进行排序
        dates = dates
        counts = counts
        
        # 数据为0时显示一个空白的图表
        if sum(counts) == 0:  # 没有数据
            ax.text(0.5, 0.5, '无数据', horizontalalignment='center', verticalalignment='center', transform=ax.transAxes)
            return
        
        # 绘制折线图
        ax.plot(range(len(dates)), counts, marker='o', linestyle='-', color='#1f77b4', linewidth=2)
        
        # 设置x轴标签
        if days <= 31:
            # 如果日期太多，只显示部分
            step = max(1, len(dates) // 10)  # 最多显示10个标签
            ax.set_xticks(range(0, len(dates), step))
            ax.set_xticklabels([dates[i].split('-')[2] for i in range(0, len(dates), step)], rotation=45)
        else:
            # 月份显示
            month_starts = []
            for i, date in enumerate(dates):
                if i == 0 or date.split('-')[1] != dates[i-1].split('-')[1]:
                    month_starts.append((i, date.split('-')[1]))
            
            ax.set_xticks([i for i, _ in month_starts])
            ax.set_xticklabels([f"{m}月" for _, m in month_starts], rotation=45)
            
        # 添加网格线
        ax.grid(True, linestyle='--', alpha=0.7)
        
        # 设置y轴最小值为0
        ax.set_ylim(bottom=0)
        
        ax.set_title(f"{'全部分类' if not category or category == '全部' else category}邮件趋势")
        ax.set_ylabel("邮件数量")
        
        return figure

class EmailStatsWidget(QWidget):
    """统计分析组件"""
    
    def __init__(self, analytics, parent=None):
        super().__init__(parent)
        self.analytics = analytics
        self.setup_ui()
        
    def setup_ui(self):
        """设置UI界面"""
        # 获取统计数据
        total_stats = self.analytics.get_email_stats()
        week_stats = self.analytics.get_email_stats('week')
        month_stats = self.analytics.get_email_stats('month')
        year_stats = self.analytics.get_email_stats('year')
        today_stats = self.analytics.get_email_stats('today')
        response_data = self.analytics.get_response_data()
        
        # 创建布局
        layout = QVBoxLayout(self)
        
        # 今天
        today_label = QLabel(f"今日邮件: {today_stats['count']} 封")
        layout.addWidget(today_label)
        
        # 一周前
        week_label = QLabel(f"本周邮件: {week_stats['count']} 封")
        layout.addWidget(week_label)
        
        # 一个月前（近似30天）
        month_label = QLabel(f"本月邮件: {month_stats['count']} 封")
        layout.addWidget(month_label)
        
        # 一年前
        year_label = QLabel(f"今年邮件: {year_stats['count']} 封")
        layout.addWidget(year_label)
        
        # 默认所有
        total_label = QLabel(f"累计邮件: {total_stats['count']} 封")
        layout.addWidget(total_label)
        
        # 回复率
        reply_rate = response_data['reply_rate'] * 100
        reply_label = QLabel(f"回复率: {reply_rate:.1f}%")
        layout.addWidget(reply_label)
        
        # 平均响应时间
        avg_time = response_data['avg_response_time']
        response_label = QLabel(f"平均响应时间: {avg_time:.1f} 分钟")
        layout.addWidget(response_label)

class AnalyticsWidget(QWidget):
    """图表小部件"""
    
    def __init__(self, analytics, parent=None):
        super().__init__(parent)
        self.analytics = analytics
        self.figure1 = plt.figure(figsize=(4, 3), dpi=100)
        self.figure2 = plt.figure(figsize=(5, 3), dpi=100)
        self.setup_ui()
        
    def setup_ui(self):
        # 布局
        main_layout = QVBoxLayout(self)
        
        # 创建画布
        self.canvas1 = FigureCanvas(self.figure1)
        self.canvas2 = FigureCanvas(self.figure2)
        
        tabs = QTabWidget()
        stats_tab = QWidget()
        self.statistics_widget = EmailStatsWidget(self.analytics)
        
        # 统计数据小部件
        stat_layout = QVBoxLayout(stats_tab)
        
        # 创建布局
        controls_layout = QHBoxLayout()
        
        # 标题
        title_label = QLabel("邮件分类统计")
        title_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        main_layout.addWidget(title_label)
        
        # 图表区域
        chart_layout = QHBoxLayout()
        
        # 分类饼图
        pie_container = QWidget()
        pie_layout = QVBoxLayout(pie_container)
        pie_layout.addWidget(self.canvas1)
        
        # 趋势图
        trend_container = QWidget()
        trend_layout = QVBoxLayout(trend_container)
        
        self.category_combo = QComboBox()
        self.category_combo.addItem("全部")
        categories = self.analytics.get_email_categories()[0]
        for cat in categories:
            self.category_combo.addItem(cat)
        self.category_combo.currentIndexChanged.connect(self.refresh)
        
        trend_layout.addWidget(self.category_combo)
        trend_layout.addWidget(self.canvas2)
        
        chart_layout.addWidget(pie_container)
        chart_layout.addWidget(trend_container)
        main_layout.addLayout(chart_layout)
        
        # 统计数据区域
        stats_layout = QHBoxLayout()
        
        # 邮件总数
        total_frame = QFrame()
        total_frame.setFrameStyle(QFrame.Shape.Box | QFrame.Shadow.Raised)
        total_layout = QVBoxLayout(total_frame)
        
        # 本周邮件数
        week_frame = QFrame()
        week_frame.setFrameStyle(QFrame.Shape.Box | QFrame.Shadow.Raised)
        week_layout = QVBoxLayout(week_frame)
        
        # 回复统计
        reply_frame = QFrame()
        reply_frame.setFrameStyle(QFrame.Shape.Box | QFrame.Shadow.Raised)
        reply_layout = QVBoxLayout(reply_frame)
        
        stat_layout.addWidget(self.statistics_widget)
        tabs.addTab(stats_tab, "统计数据")
        main_layout.addWidget(tabs)
        
        # 初始刷新
        self.refresh()
    
    def refresh(self):
        """刷新统计数据"""
        # 确保使用中文字体
        configure_matplotlib_chinese()
        
        # 重新生成图表
        category = self.category_combo.currentText() if self.category_combo.currentIndex() > 0 else None
        self.analytics.generate_category_pie(self.figure1, 'month')
        self.analytics.generate_trend_chart(self.figure2, 30, category)
        
        # 更新画布
        self.canvas1.draw()
        self.canvas2.draw()
        
        # 更新统计数据
        if hasattr(self, 'statistics_widget'):
            self.statistics_widget = EmailStatsWidget(self.analytics) 