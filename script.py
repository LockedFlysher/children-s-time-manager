import ntplib
import time
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
import ctypes
from threading import Thread, Event
import signal
import atexit
import sys
import logging
import traceback
from contextlib import suppress


class TimeChecker:
    def __init__(self):
        self.setup_logging()
        self.stop_event = Event()
        self.warning_active = False
        self.countdown_thread = None
        self.ntp_servers = [
            'pool.ntp.org',
            'time.windows.com',
            'time.apple.com',
            'time.google.com',
        ]

        # 定义锁定时间段，使用23:59代替24:00
        self.lock_times = [
            ("11:00", "13:00"),  # 上午11点后锁定到下午1点
            ("17:00", "19:00"),  # 下午5点后锁定到晚上7点
            ("21:00", "23:59"),  # 晚上9点后锁定到午夜
            ("00:00", "07:00"),  # 凌晨到早上7点
        ]

        # 注册清理函数和信号处理
        atexit.register(self.cleanup)
        self.setup_signal_handlers()

        self.logger.info("TimeChecker 初始化完成")

    def setup_logging(self):
        """设置日志系统"""
        self.logger = logging.getLogger('TimeChecker')
        self.logger.setLevel(logging.INFO)

        # 文件处理器
        file_handler = logging.FileHandler('timechecker.log', encoding='utf-8')
        file_handler.setFormatter(logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s'))

        # 控制台处理器
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s'))

        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)

    def setup_signal_handlers(self):
        """设置信号处理器"""
        with suppress(Exception):
            signal.signal(signal.SIGINT, self.signal_handler)
            signal.signal(signal.SIGTERM, self.signal_handler)

    def __del__(self):
        """析构函数"""
        self.logger.info("对象被销毁，执行锁定...")
        self.cleanup()

    def cleanup(self):
        """清理函数"""
        try:
            self.logger.info("执行清理操作...")
            self.stop_event.set()
            if self.countdown_thread and self.countdown_thread.is_alive():
                self.countdown_thread.join(timeout=1)
        except Exception as e:
            self.logger.error(f"清理过程中发生错误: {e}")
            self.logger.error(traceback.format_exc())

    def signal_handler(self, signum, frame):
        """信号处理函数"""
        self.logger.info(f"收到信号 {signum}，准备终止程序...")
        self.cleanup()
        sys.exit(0)

    def is_time_to_lock(self):
        """检查当前时间是否在需要锁定的时间段内"""
        try:
            current_time = datetime.now()
            current_time_str = current_time.strftime("%H:%M")

            for start_time, end_time in self.lock_times:
                # 转换时间字符串为 datetime 对象，只比较时和分
                current = datetime.strptime(current_time_str, "%H:%M")
                start = datetime.strptime(start_time, "%H:%M")
                end = datetime.strptime(end_time, "%H:%M")

                # 处理跨午夜的情况
                if start_time > end_time:
                    if current >= start or current <= end:
                        self.logger.info(f"当前时间 {current_time_str} 在锁定时间段 {start_time}-{end_time} 内")
                        return True
                else:
                    if start <= current <= end:
                        self.logger.info(f"当前时间 {current_time_str} 在锁定时间段 {start_time}-{end_time} 内")
                        return True

            return False
        except Exception as e:
            self.logger.error(f"检查锁定时间时发生错误: {e}")
            self.logger.error(traceback.format_exc())
            return False

    def safe_tk_operation(self, operation):
        """安全的 Tkinter 操作包装器"""
        max_retries = 3
        for attempt in range(max_retries):
            try:
                root = tk.Tk()
                root.withdraw()
                result = operation(root)
                root.destroy()
                return result
            except Exception as e:
                self.logger.error(f"Tkinter操作失败 (尝试 {attempt + 1}/{max_retries}): {e}")
                if attempt == max_retries - 1:
                    self.logger.error("所有Tkinter尝试都失败")
                time.sleep(1)
        return None

    def get_network_time(self):
        """从NTP服务器获取网络时间"""
        for server in self.ntp_servers:
            try:
                ntp_client = ntplib.NTPClient()
                response = ntp_client.request(server, timeout=5)
                self.logger.debug(f"从 {server} 成功获取时间")
                return datetime.fromtimestamp(response.tx_time)
            except Exception as e:
                self.logger.warning(f"从 {server} 获取时间失败: {e}")
                continue
        return None

    def lock_windows(self):
        """锁定Windows系统"""
        try:
            self.logger.info("锁定系统...")
            ctypes.windll.user32.LockWorkStation()
            # 重置警告状态，以便在下次检查时可以再次显示警告
            self.warning_active = False
        except Exception as e:
            self.logger.error(f"锁定系统失败: {e}")
            self.logger.error(traceback.format_exc())

    def show_countdown_warning(self, remaining_minutes, reason="时间同步"):
        """显示倒计时警告窗口"""

        def show_warning(root):
            message = ""
            if reason == "时间同步":
                message = f"系统时间可能不准确！\n系统将在{remaining_minutes}分钟后锁定...\n" + \
                          "如果在此期间成功同步网络时间且时间正常，将取消锁定"
            else:  # reason == "时间段锁定"
                message = f"当前时间已进入预设的锁定时间段！\n系统将在{remaining_minutes}分钟后锁定..."

            return messagebox.showwarning("系统锁定警告", message)

        self.safe_tk_operation(show_warning)

    def show_normal_message(self):
        """显示正常提示窗口"""

        def show_info(root):
            return messagebox.showinfo("系统提示", "时间同步正常，取消锁定")

        self.safe_tk_operation(show_info)

    def time_lock_countdown(self):
        """时间段锁定的倒计时"""
        try:
            self.logger.info("开始时间段锁定倒计时...")
            time.sleep(60)  # 等待1分钟
            if not self.stop_event.is_set():
                self.lock_windows()
        except Exception as e:
            self.logger.error(f"时间段锁定倒计时发生错误: {e}")
            self.logger.error(traceback.format_exc())

    def countdown_check(self):
        """倒计时检查线程"""
        try:
            total_checks = 30
            for i in range(total_checks):
                if self.stop_event.is_set():
                    return

                with suppress(Exception):
                    network_time = self.get_network_time()
                    if network_time is not None:
                        local_time = datetime.now()
                        time_diff = abs((network_time - local_time).total_seconds() / 60)

                        self.logger.info(f"检测成功 - 本地时间：{local_time}")
                        self.logger.info(f"网络时间：{network_time}")
                        self.logger.info(f"时间差：{time_diff:.2f}分钟")

                        if time_diff <= 5:
                            self.logger.info("网络恢复且时间正常，取消锁定")
                            self.warning_active = False
                            self.show_normal_message()
                            return

                remaining = 5 - (i + 1) * 10 / 60
                self.logger.info(f"倒计时检查：还剩 {remaining:.1f} 分钟")
                time.sleep(10)

            if not self.stop_event.is_set():
                self.lock_windows()
        except Exception as e:
            self.logger.error(f"倒计时检查过程中发生错误: {e}")
            self.logger.error(traceback.format_exc())

    def check_time(self):
        """主要的时间检查循环"""
        while not self.stop_event.is_set():
            try:
                # 首先检查是否在锁定时间段内
                if self.is_time_to_lock():
                    if not self.warning_active:
                        self.warning_active = True
                        Thread(target=self.show_countdown_warning, args=(1, "时间段锁定")).start()
                        self.countdown_thread = Thread(target=self.time_lock_countdown)
                        self.countdown_thread.start()
                    time.sleep(60)  # 在锁定时间段内，每分钟检查一次
                    continue

                # 然后检查时间同步
                network_time = self.get_network_time()

                if network_time is None:
                    self.logger.warning("无法获取网络时间")
                    if not self.warning_active:
                        self.warning_active = True
                        Thread(target=self.show_countdown_warning, args=(5, "时间同步")).start()
                        self.countdown_thread = Thread(target=self.countdown_check)
                        self.countdown_thread.start()
                else:
                    local_time = datetime.now()
                    time_diff = abs((network_time - local_time).total_seconds() / 60)

                    self.logger.info(f"本地时间：{local_time}")
                    self.logger.info(f"网络时间：{network_time}")
                    self.logger.info(f"时间差：{time_diff:.2f}分钟")

                    if time_diff > 5 and not self.warning_active:
                        self.warning_active = True
                        Thread(target=self.show_countdown_warning, args=(5, "时间同步")).start()
                        self.countdown_thread = Thread(target=self.countdown_check)
                        self.countdown_thread.start()

            except Exception as e:
                self.logger.error(f"检查时间时发生错误: {e}")
                self.logger.error(traceback.format_exc())

            # 无论是否发生错误，都等待一段时间后继续
            time.sleep(300)  # 每5分钟检查一次

    def run(self):
        """运行程序"""
        self.logger.info("时间同步检测程序已启动...")
        while True:  # 永远运行，除非被明确终止
            try:
                self.check_time()
            except KeyboardInterrupt:
                self.logger.info("检测到键盘中断，尝试重新启动检查...")
                continue
            except Exception as e:
                self.logger.error(f"主循环发生错误: {e}")
                self.logger.error(traceback.format_exc())
                time.sleep(10)  # 发生错误时等待10秒再继续
                continue


if __name__ == "__main__":
    checker = TimeChecker()
    checker.run()
