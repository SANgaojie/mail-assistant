import os
import json
import datetime
import sqlite3
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel
from PyQt6.QtCore import Qt

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
            receiver TEXT,
            subject TEXT,
            body TEXT,
            date TEXT,
            category TEXT,
            processed INTEGER DEFAULT 0,
            replied INTEGER DEFAULT 0,
            replied_at TEXT
        )
        ''')
        
        # 创建统计表
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS stats (
            date TEXT,
            category TEXT,
            count INTEGER,
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
            
            # 检查是否已存在
            cursor.execute("SELECT id FROM emails WHERE id = ?", (email_data.get('id', ''),))
            if cursor.fetchone():
                # 更新现有记录
                cursor.execute('''
                UPDATE emails SET
                    sender = ?,
                    receiver = ?,
                    subject = ?,
                    body = ?,
                    date = ?,
                    category = ?
                WHERE id = ?
                ''', (
                    email_data.get('from', ''),
                    email_data.get('to', ''),
                    email_data.get('subject', ''),
                    email_data.get('body', ''),
                    email_data.get('date', ''),
                    email_data.get('category', '未分类'),
                    email_data.get('id', '')
                ))
            else:
                # 插入新记录
                cursor.execute('''
                INSERT INTO emails (
                    id, sender, receiver, subject, body, date, category
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (
                    email_data.get('id', ''),
                    email_data.get('from', ''),
                    email_data.get('to', ''),
                    email_data.get('subject', ''),
                    email_data.get('body', ''),
                    email_data.get('date', ''),
                    email_data.get('category', '未分类')
                ))
            
            conn.commit()
            
            # 更新统计数据
            self.update_stats(email_data.get('category', '未分类'))
            
            return True
        except Exception as e:
            print(f"保存邮件数据失败: {e}")
            return False
        finally:
            conn.close()
    
    def update_stats(self, category):
        """更新统计数据
        
        Args:
            category: 邮件分类
        """
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 检查今天的该分类的记录是否存在
            cursor.execute("SELECT count FROM stats WHERE date = ? AND category = ?", (today, category))
            result = cursor.fetchone()
            
            if result:
                # 更新现有记录
                cursor.execute('''
                UPDATE stats SET count = count + 1
                WHERE date = ? AND category = ?
                ''', (today, category))
            else:
                # 插入新记录
                cursor.execute('''
                INSERT INTO stats (date, category, count)
                VALUES (?, ?, 1)
                ''', (today, category))
            
            conn.commit()
        except Exception as e:
            print(f"更新统计数据失败: {e}")
        finally:
            conn.close()
    
    def mark_email_replied(self, email_id):
        """标记邮件为已回复
        
        Args:
            email_id: 邮件ID
            
        Returns:
            bool: 是否成功
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            replied_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            cursor.execute('''
            UPDATE emails SET
                replied = 1,
                replied_at = ?
            WHERE id = ?
            ''', (replied_at, email_id))
            
            conn.commit()
            return True
        except Exception as e:
            print(f"标记邮件已回复失败: {e}")
            return False
        finally:
            conn.close()
    
    def get_category_distribution(self, time_range=None):
        """获取分类分布数据
        
        Args:
            time_range: 时间范围，如'week', 'month', 'year'，None表示所有
            
        Returns:
            dict: 分类分布数据
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            query = "SELECT category, COUNT(*) as count FROM emails"
            params = []
            
            if time_range:
                today = datetime.datetime.now()
                if time_range == 'week':
                    # 一周前
                    start_date = (today - datetime.timedelta(days=7)).strftime("%Y-%m-%d")
                elif time_range == 'month':
                    # 一个月前（近似30天）
                    start_date = (today - datetime.timedelta(days=30)).strftime("%Y-%m-%d")
                elif time_range == 'year':
                    # 一年前
                    start_date = (today - datetime.timedelta(days=365)).strftime("%Y-%m-%d")
                else:
                    start_date = None
                
                if start_date:
                    query += " WHERE date >= ?"
                    params.append(start_date)
            
            query += " GROUP BY category"
            
            cursor.execute(query, params)
            results = cursor.fetchall()
            
            distribution = {category: count for category, count in results}
            return distribution
        except Exception as e:
            print(f"获取分类分布数据失败: {e}")
            return {}
        finally:
            conn.close()
    
    def get_daily_email_count(self, days=30):
        """获取每日邮件数量
        
        Args:
            days: 天数
            
        Returns:
            dict: 日期到数量的映射
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 获取日期范围
            end_date = datetime.datetime.now()
            start_date = end_date - datetime.timedelta(days=days)
            
            # 生成日期列表
            date_list = []
            current_date = start_date
            while current_date <= end_date:
                date_list.append(current_date.strftime("%Y-%m-%d"))
                current_date += datetime.timedelta(days=1)
            
            # 查询每日邮件数量
            cursor.execute('''
            SELECT SUBSTR(date, 1, 10) as day, COUNT(*) as count
            FROM emails
            WHERE date >= ?
            GROUP BY day
            ''', (start_date.strftime("%Y-%m-%d"),))
            
            results = cursor.fetchall()
            
            # 填充结果
            daily_counts = {date: 0 for date in date_list}
            for day, count in results:
                if day in daily_counts:
                    daily_counts[day] = count
            
            return daily_counts
        except Exception as e:
            print(f"获取每日邮件数量失败: {e}")
            return {}
        finally:
            conn.close()
    
    def get_response_time_stats(self):
        """获取回复时间统计
        
        Returns:
            dict: 回复时间统计
        """
        try:
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row  # 使用字典游标
            cursor = conn.cursor()
            
            cursor.execute('''
            SELECT 
                id, 
                date, 
                replied_at,
                JULIANDAY(replied_at) - JULIANDAY(date) as response_time
            FROM emails
            WHERE replied = 1
            ''')
            
            results = cursor.fetchall()
            
            if not results:
                return {
                    "avg_response_time": 0,
                    "min_response_time": 0,
                    "max_response_time": 0,
                    "total_replies": 0
                }
            
            response_times = [row['response_time'] * 24 * 60 for row in results]  # 转换为分钟
            
            stats = {
                "avg_response_time": sum(response_times) / len(response_times),
                "min_response_time": min(response_times),
                "max_response_time": max(response_times),
                "total_replies": len(response_times)
            }
            
            return stats
        except Exception as e:
            print(f"获取回复时间统计失败: {e}")
            return {
                "avg_response_time": 0,
                "min_response_time": 0,
                "max_response_time": 0,
                "total_replies": 0
            }
        finally:
            conn.close()
    
    def get_top_senders(self, time_range=None, limit=5):
        """获取发件人排行
        
        Args:
            time_range: 时间范围，如'week', 'month', 'year'，None表示所有
            limit: 限制数量
            
        Returns:
            list: 发件人列表
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            query = "SELECT sender, COUNT(*) as count FROM emails"
            params = []
            
            if time_range:
                today = datetime.datetime.now()
                if time_range == 'week':
                    # 一周前
                    start_date = (today - datetime.timedelta(days=7)).strftime("%Y-%m-%d")
                elif time_range == 'month':
                    # 一个月前（近似30天）
                    start_date = (today - datetime.timedelta(days=30)).strftime("%Y-%m-%d")
                elif time_range == 'year':
                    # 一年前
                    start_date = (today - datetime.timedelta(days=365)).strftime("%Y-%m-%d")
                else:
                    start_date = None
                
                if start_date:
                    query += " WHERE date >= ?"
                    params.append(start_date)
            
            query += " GROUP BY sender ORDER BY count DESC LIMIT ?"
            params.append(limit)
            
            cursor.execute(query, params)
            results = cursor.fetchall()
            
            return [{"sender": sender, "count": count} for sender, count in results]
        except Exception as e:
            print(f"获取发件人排行失败: {e}")
            return []
        finally:
            conn.close()
    
    def generate_category_pie_chart(self):
        """生成分类饼图"""
        distribution = self.get_category_distribution()
        
        # 创建饼图
        fig = Figure(figsize=(6, 6))
        ax = fig.add_subplot(111)
        
        labels = list(distribution.keys())
        sizes = list(distribution.values())
        
        if sum(sizes) == 0:  # 没有数据
            ax.text(0.5, 0.5, "暂无数据", horizontalalignment='center', verticalalignment='center')
            ax.axis('off')
        else:
            ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90)
            ax.axis('equal')  # 使饼图为正圆形
            ax.set_title("邮件分类分布")
        
        return fig
    
    def generate_daily_trend_chart(self, days=30):
        """生成每日趋势图"""
        daily_counts = self.get_daily_email_count(days)
        
        # 创建折线图
        fig = Figure(figsize=(8, 4))
        ax = fig.add_subplot(111)
        
        dates = list(daily_counts.keys())
        counts = list(daily_counts.values())
        
        if sum(counts) == 0:  # 没有数据
            ax.text(0.5, 0.5, "暂无数据", horizontalalignment='center', verticalalignment='center')
            ax.axis('off')
        else:
            # 绘制折线图
            ax.plot(dates, counts, marker='o')
            
            # 设置x轴标签
            if len(dates) > 10:
                # 如果日期太多，只显示部分
                step = len(dates) // 10
                ax.set_xticks(dates[::step])
                ax.set_xticklabels(dates[::step], rotation=45)
            else:
                ax.set_xticks(dates)
                ax.set_xticklabels(dates, rotation=45)
            
            ax.set_title("每日邮件数量趋势")
            ax.set_xlabel("日期")
            ax.set_ylabel("邮件数量")
            fig.tight_layout()
        
        return fig
    
    def generate_weekly_report(self):
        """生成每周报告
        
        Returns:
            dict: 报告数据
        """
        # 获取统计数据
        stats = {
            'total_emails': self.count_emails_by_period('week'),
            'category_dist': self.get_category_distribution('week'),
            'avg_response_time': self.get_response_time_stats()['avg_response_time'],
            'top_senders': self.get_top_senders('week', limit=5)
        }
        
        return stats
    
    def count_emails_by_period(self, period):
        """统计特定时间段内的邮件数量
        
        Args:
            period: 时间段，'day', 'week', 'month', 'year'
            
        Returns:
            int: 邮件数量
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            today = datetime.datetime.now()
            
            if period == 'day':
                # 今天
                start_date = today.strftime("%Y-%m-%d")
            elif period == 'week':
                # 一周前
                start_date = (today - datetime.timedelta(days=7)).strftime("%Y-%m-%d")
            elif period == 'month':
                # 一个月前（近似30天）
                start_date = (today - datetime.timedelta(days=30)).strftime("%Y-%m-%d")
            elif period == 'year':
                # 一年前
                start_date = (today - datetime.timedelta(days=365)).strftime("%Y-%m-%d")
            else:
                # 默认所有
                cursor.execute("SELECT COUNT(*) FROM emails")
                return cursor.fetchone()[0]
            
            cursor.execute("SELECT COUNT(*) FROM emails WHERE date >= ?", (start_date,))
            return cursor.fetchone()[0]
        except Exception as e:
            print(f"统计邮件数量失败: {e}")
            return 0
        finally:
            conn.close()


class ChartWidget(QWidget):
    """图表小部件"""
    
    def __init__(self, figure, parent=None):
        """初始化图表部件
        
        Args:
            figure: matplotlib图表
            parent: 父部件
        """
        super().__init__(parent)
        
        # 布局
        layout = QVBoxLayout(self)
        
        # 创建画布
        self.canvas = FigureCanvas(figure)
        layout.addWidget(self.canvas)

        
class StatisticsWidget(QWidget):
    """统计数据小部件"""
    
    def __init__(self, analytics, parent=None):
        """初始化统计部件
        
        Args:
            analytics: 分析器实例
            parent: 父部件
        """
        super().__init__(parent)
        
        self.analytics = analytics
        
        # 创建布局
        main_layout = QVBoxLayout(self)
        
        # 标题
        title_label = QLabel("邮件统计数据")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("font-size: 18px; font-weight: bold;")
        main_layout.addWidget(title_label)
        
        # 图表区域
        charts_layout = QHBoxLayout()
        
        # 分类饼图
        pie_chart = self.analytics.generate_category_pie_chart()
        self.pie_widget = ChartWidget(pie_chart)
        charts_layout.addWidget(self.pie_widget)
        
        # 趋势图
        trend_chart = self.analytics.generate_daily_trend_chart()
        self.trend_widget = ChartWidget(trend_chart)
        charts_layout.addWidget(self.trend_widget)
        
        main_layout.addLayout(charts_layout)
        
        # 统计数据区域
        stats_layout = QHBoxLayout()
        
        # 邮件总数
        total_emails = self.analytics.count_emails_by_period(None)
        total_label = QLabel(f"总邮件数：{total_emails}")
        stats_layout.addWidget(total_label)
        
        # 本周邮件数
        weekly_emails = self.analytics.count_emails_by_period('week')
        weekly_label = QLabel(f"本周邮件数：{weekly_emails}")
        stats_layout.addWidget(weekly_label)
        
        # 回复统计
        response_stats = self.analytics.get_response_time_stats()
        avg_time = response_stats['avg_response_time']
        avg_label = QLabel(f"平均回复时间：{avg_time:.2f}分钟")
        stats_layout.addWidget(avg_label)
        
        main_layout.addLayout(stats_layout)
        
    def refresh(self):
        """刷新统计数据"""
        # 重新生成图表
        pie_chart = self.analytics.generate_category_pie_chart()
        trend_chart = self.analytics.generate_daily_trend_chart()
        
        # 更新画布
        self.pie_widget.canvas.figure = pie_chart
        self.pie_widget.canvas.draw()
        
        self.trend_widget.canvas.figure = trend_chart
        self.trend_widget.canvas.draw() 