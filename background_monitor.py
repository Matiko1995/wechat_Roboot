
from wxauto import WeChat
import threading
import time
import index
from datetime import datetime
import win32gui
import win32con
import win32process
import psutil
import os

class BackgroundWeChatMonitor:
    def __init__(self):
        self.wx = WeChat()
        self.running = True
        self.last_active_window = None
        self.wechat_hwnd = None
        self.last_check_time = time.time()
        self.last_summary_dates = {group_name: None for group_name in index.GROUP_NAMES}
        self.last_msg_ids = {}

    def find_wechat_window(self):
        """查找微信窗口句柄"""
        def callback(hwnd, hwnds):
            if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                if "微信" in window_text:
                    hwnds.append(hwnd)
            return True
        
        hwnds = []
        win32gui.EnumWindows(callback, hwnds)
        return hwnds[0] if hwnds else None

    def activate_wechat(self):
        """激活微信窗口"""
        if not self.wechat_hwnd:
            self.wechat_hwnd = self.find_wechat_window()
        
        if self.wechat_hwnd:
            # 保存当前活动窗口
            self.last_active_window = win32gui.GetForegroundWindow()
            # 激活微信窗口
            win32gui.ShowWindow(self.wechat_hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(self.wechat_hwnd)
            time.sleep(0.5)  # 给窗口切换一点时间
            return True
        return False

    def restore_previous_window(self):
        """恢复之前的活动窗口"""
        if self.last_active_window:
            try:
                win32gui.SetForegroundWindow(self.last_active_window)
            except:
                pass  # 如果窗口已关闭，忽略错误

    def check_messages(self):
        """检查新消息并处理"""
        try:
            # 激活微信窗口
            if not self.activate_wechat():
                print("无法找到或激活微信窗口")
                return
            
            # 初始化last_msg_ids（如果为空）
            if not self.last_msg_ids:
                for group_name in index.GROUP_NAMES:
                    if not self.wx.ChatWith(who=group_name):
                        print(f"找不到群聊: {group_name}")
                        continue
                    
                    last_msgs = self.wx.GetAllMessage()
                    if last_msgs and len(last_msgs) > 0 and hasattr(last_msgs[-1], 'id'):
                        self.last_msg_ids[group_name] = last_msgs[-1].id
                        print(f"设置 {group_name} 初始最后消息ID: {self.last_msg_ids[group_name]}")
            
            # 检查每个群的新消息
            for group_name in index.GROUP_NAMES:
                if not self.wx.ChatWith(who=group_name):
                    print(f"找不到群聊: {group_name}")
                    continue
                
                current_msgs = self.wx.GetAllMessage()
                if not current_msgs:
                    continue
                
                # 检查是否有新消息
                has_new_message = False
                if current_msgs and len(current_msgs) > 0:
                    if group_name not in self.last_msg_ids or self.last_msg_ids[group_name] is None:
                        has_new_message = True
                    elif hasattr(current_msgs[-1], 'id') and current_msgs[-1].id != self.last_msg_ids[group_name]:
                        has_new_message = True
                
                if has_new_message:
                    # 有新消息，处理新消息
                    new_msgs = []
                    for msg in reversed(current_msgs):
                        if group_name in self.last_msg_ids and self.last_msg_ids[group_name] is not None and hasattr(msg, 'id') and msg.id == self.last_msg_ids[group_name]:
                            break
                        new_msgs.insert(0, msg)
                    
                    # 更新最后一条消息ID
                    if current_msgs and len(current_msgs) > 0 and hasattr(current_msgs[-1], 'id'):
                        self.last_msg_ids[group_name] = current_msgs[-1].id
                    
                    # 处理新消息
                    for msg in new_msgs:
                        if not hasattr(msg, 'content'):
                            continue
                        
                        msg_content = msg.content
                        msg_sender = getattr(msg, 'sender', '未知用户')
                        
                        # 跳过机器人自己发送的消息
                        if msg_sender == 'self':
                            continue
                        
                        # 检查是否有人@机器人
                        if index.is_bot_mentioned(msg_content):
                            index.handle_mention(msg, group_name)
                        
                        # 检查是否有新订餐
                        order_content, order_count = index.parse_order_message(msg_content)
                        if order_content and order_count:
                            # 实时保存订单
                            orders = index.collect_orders(group_name)
                            index.save_to_excel(orders, group_name)
            
            # 检查是否需要发送汇总
            today = datetime.now().date()
            for group_name in index.GROUP_NAMES:
                if index.check_time_for_summary() and self.last_summary_dates.get(group_name) != today:
                    if self.wx.ChatWith(who=group_name):
                        index.send_summary(group_name)
                        self.last_summary_dates[group_name] = today
        
        finally:
            # 恢复之前的窗口
            self.restore_previous_window()

    def run(self):
        """运行监控线程"""
        print("后台微信监控已启动...")
        
        while self.running:
            try:
                current_time = time.time()
                # 每隔一段时间打印一次心跳信息
                if current_time - self.last_check_time > 60:  # 每分钟打印一次
                    print(f"监控心跳 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                    self.last_check_time = current_time
                
                # 检查新消息
                self.check_messages()
                
                # 短暂休眠，避免过度占用CPU
                time.sleep(10)  # 每10秒检查一次
                
            except Exception as e:
                print(f"监控过程中出错: {e}")
                time.sleep(30)  # 出错后等待较长时间再重试

    def stop(self):
        """停止监控"""
        self.running = False

# 创建GUI界面
def create_gui():
    """创建一个简单的GUI界面来控制后台监控"""
    import tkinter as tk
    from tkinter import messagebox
    
    # 创建监控实例
    monitor = BackgroundWeChatMonitor()
    monitor_thread = None
    
    # 创建窗口
    root = tk.Tk()
    root.title("微信订餐机器人后台监控")
    root.geometry("400x300")
    
    # 状态变量
    status_var = tk.StringVar()
    status_var.set("未启动")
    
    # 启动监控
    def start_monitor():
        nonlocal monitor_thread
        if monitor_thread is None or not monitor_thread.is_alive():
            monitor.running = True
            monitor_thread = threading.Thread(target=monitor.run)
            monitor_thread.daemon = True
            monitor_thread.start()
            status_var.set("监控中...")
            start_btn.config(state=tk.DISABLED)
            stop_btn.config(state=tk.NORMAL)
            messagebox.showinfo("提示", "后台监控已启动")
    
    # 停止监控
    def stop_monitor():
        if monitor_thread and monitor_thread.is_alive():
            monitor.stop()
            status_var.set("已停止")
            start_btn.config(state=tk.NORMAL)
            stop_btn.config(state=tk.DISABLED)
            messagebox.showinfo("提示", "后台监控已停止")
    
    # 手动发送汇总
    def manual_summary():
        try:
            if monitor.activate_wechat():
                for group_name in index.GROUP_NAMES:
                    if monitor.wx.ChatWith(who=group_name):
                        index.send_summary(group_name)
                messagebox.showinfo("提示", "已手动发送汇总")
                monitor.restore_previous_window()
        except Exception as e:
            messagebox.showerror("错误", f"发送汇总失败: {e}")
    
    # 关闭窗口时的处理
    def on_closing():
        if monitor_thread and monitor_thread.is_alive():
            if messagebox.askokcancel("退出", "监控正在运行，确定要退出吗？"):
                monitor.stop()
                root.destroy()
        else:
            root.destroy()
    
    # 创建控件
    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)
    
    title_label = tk.Label(frame, text="微信订餐机器人后台监控", font=("Arial", 14, "bold"))
    title_label.pack(pady=10)
    
    status_frame = tk.Frame(frame)
    status_frame.pack(pady=10, fill=tk.X)
    
    status_label = tk.Label(status_frame, text="状态:")
    status_label.pack(side=tk.LEFT)
    
    status_value = tk.Label(status_frame, textvariable=status_var, fg="blue")
    status_value.pack(side=tk.LEFT, padx=5)
    
    btn_frame = tk.Frame(frame)
    btn_frame.pack(pady=20)
    
    start_btn = tk.Button(btn_frame, text="启动监控", command=start_monitor, width=12)
    start_btn.grid(row=0, column=0, padx=10)
    
    stop_btn = tk.Button(btn_frame, text="停止监控", command=stop_monitor, width=12, state=tk.DISABLED)
    stop_btn.grid(row=0, column=1, padx=10)
    
    summary_btn = tk.Button(btn_frame, text="手动发送汇总", command=manual_summary, width=12)
    summary_btn.grid(row=1, column=0, columnspan=2, pady=10)
    
    # 设置关闭事件
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    # 运行GUI
    root.mainloop()

if __name__ == "__main__":
    create_gui()
