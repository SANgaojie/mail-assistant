import threading
import queue
import time
from PyQt6.QtCore import QObject, pyqtSignal, QThread, QMutex


class Worker(QObject):
    """工作线程对象"""
    
    finished = pyqtSignal(object)  # 完成信号
    progress = pyqtSignal(int)     # 进度信号
    error = pyqtSignal(str)        # 错误信号
    
    def __init__(self, func, *args, **kwargs):
        """初始化工作线程"""
        super().__init__()
        self.func = func
        self.args = args
        self.kwargs = kwargs
        self.is_cancelled = False
        
    def run(self):
        """执行任务"""
        try:
            # 处理进度回调
            if 'callback' in self.kwargs:
                original_callback = self.kwargs['callback']
                
                def callback_wrapper(progress):
                    if self.is_cancelled:
                        return False
                    self.progress.emit(progress)
                    if original_callback:
                        return original_callback(progress)
                    return True
                
                self.kwargs['callback'] = callback_wrapper
            
            result = self.func(*self.args, **self.kwargs)
            
            if not self.is_cancelled:
                self.finished.emit(result)
                
        except Exception as e:
            if not self.is_cancelled:
                self.error.emit(str(e))
                
    def cancel(self):
        """取消任务"""
        self.is_cancelled = True


class ThreadPool:
    """线程池"""
    
    def __init__(self, max_workers=4):
        """初始化线程池"""
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
                
                for _ in range(len(self.workers)):
                    self.tasks.put(None)
                    
                for worker in self.workers:
                    if worker.is_alive():
                        worker.join()
                
                self.workers.clear()
                
    def submit(self, func, callback=None, error_callback=None, *args, **kwargs):
        """提交任务"""
        print(f"提交任务: {func.__name__}")
        task_id = id(func) + int(time.time() * 1000)
        self.tasks.put((task_id, func, callback, error_callback, args, kwargs))
        return task_id
        
    def _worker_thread(self):
        """工作线程函数"""
        while self.is_running:
            try:
                task = self.tasks.get(timeout=1)
                
                if task is None:
                    break
                    
                task_id, func, callback, error_callback, args, kwargs = task
                
                try:
                    print(f"执行任务: {func.__name__}")
                    result = func(*args, **kwargs)
                    
                    if callback:
                        callback(result)
                except Exception as e:
                    print(f"任务执行出错: {e}")
                    if error_callback:
                        error_callback(str(e))
            except queue.Empty:
                continue
            except Exception as e:
                print(f"工作线程错误: {e}")
                continue


class AsyncEmailProcessor:
    """异步邮件处理器"""
    
    def __init__(self, email_connector, email_classifier, analytics=None):
        """初始化处理器"""
        self.email_connector = email_connector
        self.email_classifier = email_classifier
        self.analytics = analytics
        self.thread_pool = ThreadPool(max_workers=4)
        self.thread_pool.start()
        self.cache = AsyncOperationCache(max_size=100)
        
    def fetch_emails_async(self, callback=None, error_callback=None, folder='INBOX', search_criteria='ALL'):
        """异步获取邮件"""
        cache_key = f"emails_{folder}_{search_criteria}"
        
        cached_result = self.cache.get(cache_key)
        if cached_result:
            print(f"使用缓存的邮件列表: {len(cached_result)} 封邮件")
            if callback:
                callback(cached_result)
            return None
        
        if not self.email_connector.mail:
            print("邮箱未连接，尝试重新连接...")
            if not self.email_connector.connect():
                if error_callback:
                    error_callback("邮箱连接失败，请重新登录")
                return None
        
        def fetch_emails_wrapper():
            """包装获取邮件方法，确保使用已连接的实例，并添加重试机制"""
            # 添加重试机制
            max_retries = 3
            retry_count = 0
            
            while retry_count < max_retries:
                try:
                    # 检查连接是否有效
                    if not self.email_connector.mail:
                        print(f"重试 {retry_count+1}/{max_retries}: 邮箱未连接，尝试重新连接")
                        if not self.email_connector.connect():
                            retry_count += 1
                            continue
                    
                    # 尝试获取邮件
                    result = self.email_connector.fetch_emails(folder=folder, search_criteria=search_criteria)
                    
                    # 如果成功获取邮件，返回结果
                    if result:
                        return result
                    
                    # 如果没有获取到邮件，但没有异常，也算成功（可能就是没有邮件）
                    if result is not None and len(result) == 0:
                        return []
                        
                    # 此时获取邮件失败，但没有引发异常，重试
                    print(f"重试 {retry_count+1}/{max_retries}: 获取邮件失败但未引发异常")
                    retry_count += 1
                    
                except Exception as e:
                    print(f"重试 {retry_count+1}/{max_retries}: 获取邮件出错: {e}")
                    # 出错了，重置连接
                    try:
                        self.email_connector.close()
                    except:
                        pass
                        
                    retry_count += 1
                    if retry_count >= max_retries:
                        raise Exception(f"获取邮件失败，已重试 {max_retries} 次: {e}")
            
            # 所有重试都失败
            raise Exception(f"获取邮件失败，已重试 {max_retries} 次")
            
        def cache_result_callback(result):
            # 缓存结果
            if result:  # 只缓存非空结果
                self.cache.put(cache_key, result)
            # 调用原回调
            if callback:
                callback(result)
                
        return self.thread_pool.submit(
            fetch_emails_wrapper,  # 使用包装函数而不是直接使用email_connector.fetch_emails
            callback=cache_result_callback,
            error_callback=error_callback
        )
        
    def classify_emails_async(self, emails, callback=None, error_callback=None):
        """异步分类邮件"""
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
        """异步批量处理邮件"""
        print(f"开始批量处理 {len(emails)} 封邮件，目标分类: {target_category}")
        
        try:
            def batch_process(emails_to_process, target_cat, reply_text):
                """实际处理邮件的函数"""
                processed = []
                total = len(emails_to_process)
                
                for i, email_data in enumerate(emails_to_process):
                    try:
                        print(f"处理邮件 {i+1}/{total}")
                        
                        email_copy = email_data.copy()
                        
                        if target_cat:
                            email_copy = self.email_classifier.tag_email(email_copy, target_cat)
                            print(f"邮件已标记为类别: {target_cat}")
                            
                        if self.analytics:
                            self.analytics.save_email(email_copy)
                        
                        processed.append(email_copy)
                    except Exception as e:
                        print(f"处理邮件失败: {e}")
                
                print(f"批量处理完成，成功处理 {len(processed)}/{total} 封邮件")
                return processed
                
            return self.thread_pool.submit(
                batch_process,
                callback=callback,
                error_callback=error_callback,
                emails_to_process=emails,
                target_cat=target_category,
                reply_text=reply_content
            )
            
        except Exception as e:
            print(f"启动批量处理任务失败: {e}")
            if error_callback:
                error_callback(f"启动处理失败: {e}")
            return None
        
    def invalidate_cache(self, pattern=None):
        """使缓存无效"""
        if pattern:
            # 按模式清除特定缓存（未实现）
            pass
        else:
            self.cache.clear()
        
    def close(self):
        """关闭处理器"""
        self.thread_pool.stop()


class QtThreadWorker(QThread):
    """Qt线程工作器"""
    
    finished = pyqtSignal(object)  # 完成信号
    progress = pyqtSignal(int)     # 进度信号
    error = pyqtSignal(str)        # 错误信号
    
    def __init__(self, func, *args, **kwargs):
        """初始化Qt线程工作器"""
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
    """异步操作缓存"""
    
    def __init__(self, max_size=100):
        """初始化缓存"""
        self.cache = {}
        self.max_size = max_size
        self.access_times = {}
        self.mutex = QMutex()
        
    def get(self, key):
        """获取缓存项"""
        self.mutex.lock()
        try:
            if key in self.cache:
                self.access_times[key] = time.time()
                return self.cache[key]
            return None
        finally:
            self.mutex.unlock()
            
    def put(self, key, value):
        """添加缓存项"""
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
        """移除缓存项"""
        self.mutex.lock()
        try:
            if key in self.cache:
                del self.cache[key]
                del self.access_times[key]
        finally:
            self.mutex.unlock() 