
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import os
import sys
import time
import re
from datetime import datetime
import json

# 导入主程序模块
try:
    from index import WeChat, monitor_group, collect_orders, save_to_excel, BOT_NAME, generate_summary
except ImportError:
    # 如果直接运行GUI，可能需要添加路径
    import sys
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    from index import WeChat, monitor_group, collect_orders, save_to_excel, BOT_NAME, generate_summary

class RedirectText:
    """重定向标准输出到文本控件"""
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.buffer = ""
        
    def write(self, string):
        self.buffer += string
        # 如果包含换行符，则更新文本控件
        if '\n' in self.buffer:
            self.text_widget.config(state=tk.NORMAL)
            lines = self.buffer.split('\n')
            for i, line in enumerate(lines[:-1]):
                if line.strip():  # 只添加非空行
                    self.text_widget.insert(tk.END, line + '\n')
            self.text_widget.see(tk.END)  # 自动滚动到底部
            self.text_widget.config(state=tk.DISABLED)
            self.buffer = lines[-1]
    
    def flush(self):
        if self.buffer:
            self.text_widget.config(state=tk.NORMAL)
            self.text_widget.insert(tk.END, self.buffer)
            self.text_widget.see(tk.END)
            self.text_widget.config(state=tk.DISABLED)
            self.buffer = ""

class WeChatBotApp:
    def __init__(self, root):
        self.root = root
        self.root.title("微信订餐机器人")
        # 增加默认窗口高度，确保状态栏可见
        self.root.geometry("800x620")  # 增加高度
        self.root.minsize(800, 620)    # 设置最小窗口大小
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 创建配置文件路径
        self.config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bot_config.json")
        
        # 加载配置
        self.config = self.load_config()
        
        # 创建主框架
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建标题标签
        title_label = ttk.Label(self.main_frame, text="微信订餐机器人控制面板", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # 创建机器人名称显示
        bot_name_frame = ttk.Frame(self.main_frame)
        bot_name_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(bot_name_frame, text="机器人名称:").pack(side=tk.LEFT, padx=5)
        self.bot_name_var = tk.StringVar(value=BOT_NAME)
        ttk.Label(bot_name_frame, textvariable=self.bot_name_var, font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        
        # 创建群聊配置区域
        group_frame = ttk.LabelFrame(self.main_frame, text="群聊配置")
        group_frame.pack(fill=tk.X, pady=10, padx=5)
        
        # 群聊列表
        self.groups_frame = ttk.Frame(group_frame)
        self.groups_frame.pack(fill=tk.X, pady=5, padx=5)
        
        # 初始化群聊列表
        self.group_entries = []
        self.at_entries = []
        self.order_count_labels = []  # 添加订餐份数标签列表
        
        # 添加已有的群聊配置
        for group_name, at_person in self.config.get("AT_PERSONS", {}).items():
            self.add_group_entry(group_name, at_person)
        
        # 如果没有群聊配置，添加一个空的
        if not self.group_entries:
            self.add_group_entry("", "")
        
        # 添加/删除群聊按钮
        btn_frame = ttk.Frame(group_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(btn_frame, text="添加群聊", command=self.add_group_entry).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="删除最后一个", command=self.remove_last_group).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="保存配置", command=self.save_config).pack(side=tk.RIGHT, padx=5)
        
        # 创建控制按钮
        control_frame = ttk.Frame(self.main_frame)
        control_frame.pack(fill=tk.X, pady=10)
        
        self.start_btn = ttk.Button(control_frame, text="启动机器人", command=self.start_bot)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_btn = ttk.Button(control_frame, text="停止机器人", command=self.stop_bot, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        # 刷新订餐数量按钮
        self.refresh_btn = ttk.Button(control_frame, text="刷新订餐数量", command=self.refresh_order_counts)
        self.refresh_btn.pack(side=tk.LEFT, padx=5)
        
        # 添加打开当天Excel按钮
        self.open_excel_btn = ttk.Button(control_frame, text="打开当天Excel", command=self.open_today_excel)
        self.open_excel_btn.pack(side=tk.LEFT, padx=5)
        
        # 创建日志显示区域，减少其高度以腾出空间给状态栏
        log_frame = ttk.LabelFrame(self.main_frame, text="运行日志")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 5), padx=5)  # 减少底部padding
        
        # 日志文本区域
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 重定向标准输出到日志区域
        self.redirect = RedirectText(self.log_text)
        sys.stdout = self.redirect
        
        # 机器人线程
        self.bot_thread = None
        self.running = False
        
        # 创建底部状态栏，确保时间显示在右下角
        # 使用Frame而不是ttk.Frame，以便更好地控制外观
        status_frame = tk.Frame(self.root, bd=1, relief=tk.SUNKEN)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 左侧状态信息
        self.status_var = tk.StringVar(value="就绪")
        status_label = tk.Label(status_frame, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 右侧时间显示 - 使用固定宽度，确保在默认窗口大小下也能显示
        self.time_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        # 使用Label而不是ttk.Label，以便更好地控制外观
        time_label = tk.Label(
            status_frame, 
            textvariable=self.time_var, 
            bd=1,
            relief=tk.SUNKEN, 
            width=19,  # 减少宽度以确保在小窗口中也能显示
            anchor=tk.E  # 右对齐
        )
        time_label.pack(side=tk.RIGHT)
        
        # 开始时间更新
        self.update_clock()
        
        # 打印初始信息
        print(f"微信订餐机器人界面已启动")
        print(f"当前机器人名称: {BOT_NAME}")
        
        # 初始化订餐数量
        self.refresh_order_counts()
    
    def update_clock(self):
        """更新时钟显示"""
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.time_var.set(current_time)
        # 每秒更新一次
        self.root.after(1000, self.update_clock)
    
    def add_group_entry(self, group_name="", at_person=""):
        """添加一个群聊配置行"""
        frame = ttk.Frame(self.groups_frame)
        frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(frame, text="群聊名称:").pack(side=tk.LEFT, padx=5)
        group_entry = ttk.Entry(frame, width=30)
        group_entry.pack(side=tk.LEFT, padx=5)
        group_entry.insert(0, group_name)
        
        ttk.Label(frame, text="@的人:").pack(side=tk.LEFT, padx=5)
        at_entry = ttk.Entry(frame, width=20)
        at_entry.pack(side=tk.LEFT, padx=5)
        at_entry.insert(0, at_person)
        
        # 添加订餐份数显示
        ttk.Label(frame, text="当前订餐:").pack(side=tk.LEFT, padx=5)
        order_count_var = tk.StringVar(value="0份")
        order_count_label = ttk.Label(frame, textvariable=order_count_var, width=10)
        order_count_label.pack(side=tk.LEFT, padx=5)
        
        # 添加打开Excel按钮到每个群聊行
        excel_btn = ttk.Button(
            frame, 
            text="打开Excel", 
            width=8,
            command=lambda g=group_name: self.open_group_excel(g)
        )
        excel_btn.pack(side=tk.LEFT, padx=5)
        
        self.group_entries.append(group_entry)
        self.at_entries.append(at_entry)
        self.order_count_labels.append(order_count_var)
    
    def remove_last_group(self):
        """删除最后一个群聊配置行"""
        if len(self.group_entries) > 1:  # 至少保留一个
            self.group_entries[-1].master.destroy()
            self.group_entries.pop()
            self.at_entries.pop()
            self.order_count_labels.pop()
    
    def refresh_order_counts(self):
        """刷新所有群聊的订餐数量"""
        print("正在刷新订餐数量...")
        
        for i, group_entry in enumerate(self.group_entries):
            group_name = group_entry.get().strip()
            if not group_name:
                continue
                
            try:
                # 获取订单
                orders = collect_orders(group_name)
                
                if orders:
                    # 计算总人数和总份数
                    people_set = set()
                    total_count = 0
                    
                    for order in orders:
                        people_set.add(order['发送人'])
                        try:
                            count = int(order.get('订餐数量', 1))
                            total_count += count
                        except:
                            total_count += 1
                    
                    people_count = len(people_set)
                    
                    # 更新显示
                    self.order_count_labels[i].set(f"{people_count}人/{total_count}份")
                    print(f"群聊 '{group_name}' 当前订餐: {people_count}人/{total_count}份")
                else:
                    self.order_count_labels[i].set("0人/0份")
                    
            except Exception as e:
                print(f"刷新群聊 '{group_name}' 订餐数量时出错: {e}")
                self.order_count_labels[i].set("错误")
        
        print("订餐数量刷新完成")
    
    def load_config(self):
        """加载配置文件"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"加载配置文件失败: {e}")
        
        # 返回默认配置
        return {
            "AT_PERSONS": {
                "英明中、晚饭订餐群": "布鲁布鲁"
            }
        }
    
    def save_config(self):
        """保存配置到文件"""
        at_persons = {}
        for group_entry, at_entry in zip(self.group_entries, self.at_entries):
            group_name = group_entry.get().strip()
            at_person = at_entry.get().strip()
            if group_name:  # 只保存有群名的配置
                at_persons[group_name] = at_person
        
        config = {
            "AT_PERSONS": at_persons
        }
        
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            
            messagebox.showinfo("保存成功", "配置已保存")
            print("配置已保存")
            
            # 更新全局变量
            import index
            index.AT_PERSONS = at_persons
            index.GROUP_NAMES = list(at_persons.keys())
            
        except Exception as e:
            messagebox.showerror("保存失败", f"保存配置失败: {e}")
            print(f"保存配置失败: {e}")
    
    def start_bot(self):
        """启动机器人"""
        if self.running:
            return
        
        # 先保存配置
        self.save_config()
        
        # 更新状态
        self.running = True
        self.status_var.set("正在运行...")
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        
        # 清空日志
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        # 启动机器人线程
        self.bot_thread = threading.Thread(target=self.run_bot)
        self.bot_thread.daemon = True
        self.bot_thread.start()
        
        print("机器人已启动")
    
    def run_bot(self):
        """运行机器人的线程函数"""
        try:
            print("订餐统计机器人已启动...")
            print(f"机器人名称: {BOT_NAME}")
            
            # 首次运行时收集并保存当前订单
            for group_name in index.GROUP_NAMES:
                orders = index.collect_orders(group_name)
                index.save_to_excel(orders, group_name)
            
            # 调用原始的监控函数
            index.monitor_group()
            
        except Exception as e:
            print(f"程序运行出错: {e}")
            self.stop_bot()
    
    def stop_bot(self):
        """停止机器人"""
        if not self.running:
            return
        
        self.running = False
        self.status_var.set("已停止")
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        
        if self.bot_thread and self.bot_thread.is_alive():
            self.bot_thread.join(timeout=2)
        
        print("机器人已停止")
    
    def open_group_excel(self, group_name=None):
        """打开指定群聊的Excel文件"""
        try:
            # 如果没有指定群聊名称，则使用当前选中的群聊
            if not group_name:
                # 获取第一个有效的群聊名称
                for entry in self.group_entries:
                    group_name = entry.get().strip()
                    if group_name:
                        break
                
                if not group_name:
                    messagebox.showwarning("提示", "请先配置群聊名称")
                    return
            
            # 获取当前月份
            current_month = datetime.now().strftime("%m月")
            
            # 构建Excel文件路径: /订餐统计/月份_群聊名称_订餐统计表.xlsx
            excel_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "订餐统计")
            excel_filename = f"{current_month}_{group_name}_订餐统计表.xlsx"
            excel_path = os.path.join(excel_dir, excel_filename)
            
            print(f"尝试打开Excel文件: {excel_path}")
            
            # 检查文件是否存在
            if os.path.exists(excel_path):
                # 使用系统默认程序打开Excel文件
                if os.name == 'nt':  # Windows
                    os.startfile(excel_path)
                elif os.name == 'posix':  # macOS/Linux
                    import subprocess
                    subprocess.call(('open' if sys.platform == 'darwin' else 'xdg-open', excel_path))
                
                print(f"已打开Excel文件: {excel_path}")
            else:
                # 如果找不到指定月份的文件，尝试查找该群聊的任何Excel文件
                found_files = []
                if os.path.exists(excel_dir):
                    for filename in os.listdir(excel_dir):
                        if group_name in filename and filename.endswith('.xlsx'):
                            found_files.append(filename)
                
                if found_files:
                    # 按文件修改时间排序，打开最新的文件
                    found_files.sort(key=lambda x: os.path.getmtime(os.path.join(excel_dir, x)), reverse=True)
                    latest_file = os.path.join(excel_dir, found_files[0])
                    
                    if os.name == 'nt':  # Windows
                        os.startfile(latest_file)
                    elif os.name == 'posix':  # macOS/Linux
                        import subprocess
                        subprocess.call(('open' if sys.platform == 'darwin' else 'xdg-open', latest_file))
                    
                    print(f"找不到当月文件，已打开最新的Excel文件: {latest_file}")
                else:
                    messagebox.showwarning("文件未找到", f"未找到群聊 '{group_name}' 的Excel文件")
                    print(f"未找到群聊 '{group_name}' 的Excel文件")
        
        except Exception as e:
            error_msg = f"打开Excel文件时出错: {e}"
            messagebox.showerror("错误", error_msg)
            print(error_msg)
    
    def open_today_excel(self):
        """打开当天的Excel文件"""
        try:
            today = datetime.now().strftime("%Y-%m-%d")
            # 假设Excel文件保存在当前目录下的excel文件夹中
            excel_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "excel")
            excel_file = os.path.join(excel_dir, f"{today}.xlsx")
            
            if os.path.exists(excel_file):
                os.startfile(excel_file)  # Windows系统
                print(f"已打开Excel文件: {excel_file}")
            else:
                messagebox.showwarning("文件未找到", f"未找到当天的Excel文件: {excel_file}")
                print(f"Excel文件不存在: {excel_file}")
                
        except Exception as e:
            messagebox.showerror("打开失败", f"打开Excel文件失败: {e}")
            print(f"打开Excel文件失败: {e}")
    
    def on_closing(self):
        """窗口关闭事件"""
        if self.running:
            if messagebox.askokcancel("退出确认", "机器人正在运行中，确定要退出吗？"):
                self.stop_bot()
                self.root.destroy()
        else:
            self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = WeChatBotApp(root)
    root.mainloop()
