import threading
import queue
import time
from PyQt6.QtCore import QObject, pyqtSignal, QThread, QMutex


class Worker(QObject):
    """工作线程对象，用于在后台执行耗时操作"""
    
    # 信号定义
    finished = pyqtSignal(object)  # 任务完成信号，携带结果
    progress = pyqtSignal(int)     # 进度信号，0-100
    error = pyqtSignal(str)        # 错误信号，携带错误信息
    
    def __init__(self, func, *args, **kwargs):
        """初始化工作线程
        
        Args:
            func: 要执行的函数
            *args, **kwargs: 函数参数
        """
        super().__init__()
        self.func = func
        self.args = args
        self.kwargs = kwargs
        self.is_cancelled = False
        
    def run(self):
        """执行任务"""
        try:
            # 如果函数支持取消和进度报告
            if 'callback' in self.kwargs:
                original_callback = self.kwargs['callback']
                
                # 创建新的回调，包装进度报告
                def callback_wrapper(progress):
                    if self.is_cancelled:
                        return False
                    self.progress.emit(progress)
                    if original_callback:
                        return original_callback(progress)
                    return True
                
                self.kwargs['callback'] = callback_wrapper
            
            # 执行函数
            result = self.func(*self.args, **self.kwargs)
            
            # 如果没有被取消，发出完成信号
            if not self.is_cancelled:
                self.finished.emit(result)
                
        except Exception as e:
            if not self.is_cancelled:
                self.error.emit(str(e))
                
    def cancel(self):
        """取消任务"""
        self.is_cancelled = True


class ThreadPool:
    """线程池，管理多个工作线程"""
    
    def __init__(self, max_workers=4):
        """初始化线程池
        
        Args:
            max_workers: 最大工作线程数
        """
        self.max_workers = max_workers
        self.tasks = queue.Queue()
        self.workers = []
        self.is_running = False
        self.lock = threading.Lock()
        
    def start(self):
        """启动线程池"""
        with self.lock:
            if not self.is_running:
                self.is_running = True
                
                # 创建工作线程
                for _ in range(self.max_workers):
                    worker = threading.Thread(target=self._worker_thread)
                    worker.daemon = True
                    worker.start()
                    self.workers.append(worker)
                    
    def stop(self):
        """停止线程池"""
        with self.lock:
            if self.is_running:
                self.is_running = False
                
                # 添加终止任务
                for _ in range(len(self.workers)):
                    self.tasks.put(None)
                    
                # 等待所有线程结束
                for worker in self.workers:
                    if worker.is_alive():
                        worker.join()
                
                self.workers.clear()
                
    def submit(self, func, callback=None, error_callback=None, *args, **kwargs):
        """提交任务
        
        Args:
            func: 要执行的函数
            callback: 完成回调
            error_callback: 错误回调
            *args, **kwargs: 函数参数
            
        Returns:
            任务ID
        """
        task_id = id(func) + int(time.time() * 1000)
        self.tasks.put((task_id, func, callback, error_callback, args, kwargs))
        return task_id
        
    def _worker_thread(self):
        """工作线程函数"""
        while self.is_running:
            try:
                # 获取任务
                task = self.tasks.get(timeout=1)
                
                # 终止信号
                if task is None:
                    break
                    
                task_id, func, callback, error_callback, args, kwargs = task
                
                try:
                    # 执行任务
                    result = func(*args, **kwargs)
                    
                    # 如果有回调，执行回调
                    if callback:
                        callback(result)
                except Exception as e:
                    # 如果有错误回调，执行错误回调
                    if error_callback:
                        error_callback(str(e))
            except queue.Empty:
                continue


class AsyncEmailProcessor:
    """异步邮件处理器，用于在后台线程处理邮件相关操作"""
    
    def __init__(self, email_connector, email_classifier, analytics=None):
        """初始化异步邮件处理器
        
        Args:
            email_connector: 邮件连接器
            email_classifier: 邮件分类器
            analytics: 数据分析器
        """
        self.email_connector = email_connector
        self.email_classifier = email_classifier
        self.analytics = analytics
        self.thread_pool = ThreadPool(max_workers=4)
        self.thread_pool.start()
        self.cache = AsyncOperationCache(max_size=100)
        
    def fetch_emails_async(self, callback=None, error_callback=None, folder='INBOX', search_criteria='ALL'):
        """异步获取邮件
        
        Args:
            callback: 完成回调
            error_callback: 错误回调
            folder: 邮件文件夹
            search_criteria: 搜索条件
            
        Returns:
            任务ID
        """
        # 创建缓存键
        cache_key = f"emails_{folder}_{search_criteria}"
        
        # 检查缓存
        cached_result = self.cache.get(cache_key)
        if cached_result:
            print(f"使用缓存的邮件列表: {len(cached_result)} 封邮件")
            if callback:
                callback(cached_result)
            return None
            
        def cache_result_callback(result):
            # 缓存结果
            self.cache.put(cache_key, result)
            # 调用原回调
            if callback:
                callback(result)
                
        return self.thread_pool.submit(
            self.email_connector.fetch_emails,
            cache_result_callback,
            error_callback,
            folder,
            search_criteria
        )
        
    def classify_emails_async(self, emails, callback=None, error_callback=None):
        """异步分类邮件
        
        Args:
            emails: 邮件列表
            callback: 完成回调
            error_callback: 错误回调
            
        Returns:
            任务ID
        """
        # 创建缓存键 - 使用邮件ID列表
        email_ids = [email.get('id', '') for email in emails]
        cache_key = f"classify_{'_'.join(email_ids)}"
        
        # 检查缓存
        cached_result = self.cache.get(cache_key)
        if cached_result:
            print(f"使用缓存的分类结果")
            if callback:
                callback(cached_result)
            return None
            
        def classify_emails(emails):
            for i, email_data in enumerate(emails):
                # 分类邮件
                category = self.email_classifier.classify_email(email_data)
                self.email_classifier.tag_email(email_data, category)
                
            return emails
            
        def cache_result_callback(result):
            # 缓存结果
            self.cache.put(cache_key, result)
            # 调用原回调
            if callback:
                callback(result)
                
        return self.thread_pool.submit(
            classify_emails,
            cache_result_callback,
            error_callback,
            emails
        )
        
    def batch_process_async(self, emails, target_category=None, reply_content=None, callback=None, error_callback=None):
        """异步批量处理邮件
        
        Args:
            emails: 邮件列表
            target_category: 目标分类
            reply_content: 回复内容
            callback: 完成回调
            error_callback: 错误回调
            
        Returns:
            任务ID
        """
        def batch_process(emails, target_category, reply_content):
            processed = []
            for email_data in emails:
                try:
                    # 分类邮件
                    if target_category:
                        email_data = self.email_classifier.tag_email(email_data, target_category)
                        
                    # 添加到分析
                    if self.analytics:
                        self.analytics.save_email(email_data)
                    
                    processed.append(email_data)
                except Exception as e:
                    print(f"处理邮件失败: {e}")
                    
            return processed
            
        return self.thread_pool.submit(
            batch_process,
            callback,
            error_callback,
            emails,
            target_category,
            reply_content
        )
        
    def invalidate_cache(self, pattern=None):
        """使缓存无效
        
        Args:
            pattern: 缓存键模式，None表示清除所有缓存
        """
        if pattern:
            # 按模式清除特定缓存（未实现）
            # 这里需要实现一个按模式匹配的功能
            pass
        else:
            # 清除所有缓存
            self.cache.clear()
        
    def close(self):
        """关闭处理器"""
        self.thread_pool.stop()


class QtThreadWorker(QThread):
    """Qt线程工作器，用于在Qt应用程序中执行耗时操作"""
    
    # 信号定义
    finished = pyqtSignal(object)  # 任务完成信号，携带结果
    progress = pyqtSignal(int)     # 进度信号，0-100
    error = pyqtSignal(str)        # 错误信号，携带错误信息
    
    def __init__(self, func, *args, **kwargs):
        """初始化Qt线程工作器
        
        Args:
            func: 要执行的函数
            *args, **kwargs: 函数参数
        """
        super().__init__()
        self.func = func
        self.args = args
        self.kwargs = kwargs
        self.mutex = QMutex()
        self.is_cancelled = False
        
    def run(self):
        """执行任务"""
        try:
            # 如果函数支持进度报告
            if 'progress_callback' in self.kwargs:
                original_callback = self.kwargs['progress_callback']
                
                # 创建新的回调
                def callback_wrapper(progress):
                    self.mutex.lock()
                    cancelled = self.is_cancelled
                    self.mutex.unlock()
                    
                    if cancelled:
                        return False
                        
                    self.progress.emit(progress)
                    if original_callback:
                        return original_callback(progress)
                    return True
                
                self.kwargs['progress_callback'] = callback_wrapper
            
            # 执行函数
            result = self.func(*self.args, **self.kwargs)
            
            # 检查是否被取消
            self.mutex.lock()
            cancelled = self.is_cancelled
            self.mutex.unlock()
            
            if not cancelled:
                self.finished.emit(result)
                
        except Exception as e:
            self.mutex.lock()
            cancelled = self.is_cancelled
            self.mutex.unlock()
            
            if not cancelled:
                self.error.emit(str(e))
        
    def cancel(self):
        """取消任务"""
        self.mutex.lock()
        self.is_cancelled = True
        self.mutex.unlock()


class AsyncOperationCache:
    """异步操作缓存，用于缓存耗时操作的结果"""
    
    def __init__(self, max_size=100):
        """初始化缓存
        
        Args:
            max_size: 最大缓存项数
        """
        self.cache = {}
        self.max_size = max_size
        self.access_times = {}
        self.mutex = QMutex()
        
    def get(self, key):
        """获取缓存项
        
        Args:
            key: 缓存键
            
        Returns:
            缓存值，如果不存在则返回None
        """
        self.mutex.lock()
        try:
            if key in self.cache:
                self.access_times[key] = time.time()
                return self.cache[key]
            return None
        finally:
            self.mutex.unlock()
            
    def put(self, key, value):
        """添加缓存项
        
        Args:
            key: 缓存键
            value: 缓存值
        """
        self.mutex.lock()
        try:
            # 如果缓存已满，移除最久未使用的项
            if len(self.cache) >= self.max_size:
                oldest_key = min(self.access_times.items(), key=lambda x: x[1])[0]
                del self.cache[oldest_key]
                del self.access_times[oldest_key]
                
            self.cache[key] = value
            self.access_times[key] = time.time()
        finally:
            self.mutex.unlock()
            
    def clear(self):
        """清空缓存"""
        self.mutex.lock()
        try:
            self.cache.clear()
            self.access_times.clear()
        finally:
            self.mutex.unlock()
            
    def remove(self, key):
        """移除缓存项
        
        Args:
            key: 缓存键
        """
        self.mutex.lock()
        try:
            if key in self.cache:
                del self.cache[key]
                del self.access_times[key]
        finally:
            self.mutex.unlock() 