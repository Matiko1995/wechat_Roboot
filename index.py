
from wxauto import WeChat
import re
import os
import time
import pandas as pd
from datetime import datetime

# 尝试导入schedule模块，如果不存在则使用自定义的定时功能

try:
    import schedule
    HAS_SCHEDULE = True
except ImportError:
    HAS_SCHEDULE = False
    print("警告: 未安装schedule模块，将使用简单的定时功能")

# 初始化微信实例
wx = WeChat()
# 获取微信窗口标题 - 修复这部分代码
try:
    # 获取当前登录的微信名称
    wx_window_name = wx.GetWeChatTitle()
    print(f"初始化成功，获取到已登录窗口：{wx_window_name}")
    # 更新BOT_NAME为实际的窗口名称
    BOT_NAME = '良行上厨®快餐店订餐机器人'  # 直接使用实际的名称
    print(f"设置机器人名称为：{BOT_NAME}")
except Exception as e:
    print(f"初始化微信窗口失败: {e}")
    # 保持默认值
    BOT_NAME = '良行上厨®快餐店订餐机器人'  # 使用实际名称作为默认值
    print(f"使用默认机器人名称: {BOT_NAME}")

# 修改群聊配置为列表
# GROUP_NAMES = ["订餐测试群聊"]  #, "英明中、晚饭订餐群" 可以添加多个群聊名称
# 需要@的人（可以为每个群设置不同的@人）
AT_PERSONS = {
    "英明中、晚饭订餐群":'布鲁布鲁'
}

# 从AT_PERSONS中获取群聊名称列表，这样只监听设置了@人的群聊
GROUP_NAMES = list(AT_PERSONS.keys())

# 统计文件保存路径
SAVE_DIR = "订餐统计"
# 机器人的微信名称（用于检测是否被@）
# BOT_NAME = '订餐机器人'  #wx_window_name

# 确保保存目录存在
if not os.path.exists(SAVE_DIR):
    os.makedirs(SAVE_DIR)

# 解析订餐信息的正则表达式
# 匹配格式如：人名xx xxx xxx xxx，共xx份
ORDER_PATTERN = r'(.+?)，共(\d+)份'
# 新增匹配格式如：龚建玲 陈可欣 柴奇 高菊玲 齐相霞 徐泽宇 李洋 宋孝营， 共8人
# 支持中英文逗号，支持不同数量的空格
ORDER_PATTERN_PEOPLE = r'(.+?)[,，]\s*共(\d+)人'
# 检测@消息的正则表达式
AT_PATTERN = r'@([^\s]+)'

def get_current_month_year():
    """获取当前月份和年份"""
    now = datetime.now()
    return now.month, now.year

def get_excel_path(group_name=None):
    """获取当前月份的Excel文件路径
    
    Args:
        group_name: 群聊名称，如果提供则生成特定群的文件名
    """
    month, year = get_current_month_year()
    
    # 如果提供了群聊名称，则在文件名中包含群聊名
    if group_name:
        # 替换群聊名中可能导致文件名无效的字符
        safe_group_name = re.sub(r'[\\/:*?"<>|]', '_', group_name)
        filename = f"{month}月_{safe_group_name}_订餐统计表.xlsx"
    else:
        filename = f"{month}月英明精密订餐统计表.xlsx"
    
    return os.path.join(SAVE_DIR, filename)

def get_today_date():
    """获取今天的日期字符串"""
    return datetime.now().strftime("%Y-%m-%d")

def parse_order_message(content):
    """解析订餐消息内容"""
    print(f"尝试解析订餐消息: {content}")
    
    # 检查消息是否包含逗号（中英文/全角半角）
    if not any(char in content for char in [',', '，']):
        print("消息中不包含逗号，不符合订餐格式要求")
        return None, None
    
    # 尝试匹配第一种格式：xxx，共xx份
    match = re.search(ORDER_PATTERN, content)
    if match:
        order_content = match.group(1)  # 订餐内容
        order_count = int(match.group(2))  # 订餐份数
        print(f"成功解析订餐(份数格式): 内容={order_content}, 份数={order_count}")
        return order_content, order_count
    
    # 尝试匹配第二种格式：xxx xxx xxx， 共xx人
    match = re.search(ORDER_PATTERN_PEOPLE, content)
    if match:
        order_content = match.group(1).strip()  # 人员名单，去除首尾空格
        order_count = int(match.group(2))  # 人数
        print(f"成功解析订餐(人数格式): 内容={order_content}, 人数={order_count}")
        return order_content, order_count
    
    # 不再尝试匹配没有逗号的格式，因为我们已经要求必须包含逗号
    
    print("未能匹配订餐格式")
    return None, None

def is_bot_mentioned(content):
    """检查消息中是否@了机器人"""
    print(f"检查是否@机器人: {content}")
    
    # 首先检查是否包含@符号
    if '@' not in content:
        print("消息中不包含@符号")
        return False
    
    # 使用正则表达式匹配@后面的名称
    at_matches = re.findall(AT_PATTERN, content)
    print(f"@匹配结果: {at_matches}")
    
    # 检查是否有匹配结果
    if not at_matches:
        print("没有找到@匹配")
        return False
    
    # 检查是否@了机器人（使用精确的名称）
    for name in at_matches:
        name = name.strip()
        print(f"检查@名称: '{name}'")
        if name == BOT_NAME or name == '良行上厨®快餐店订餐机器人':
            print(f"检测到@机器人: {name}")
            return True
    
    print("未检测到@机器人")
    return False

def save_to_excel(orders, group_name):
    """保存订单到Excel"""
    if not orders:
        print("没有订单数据需要保存")
        return
    
    excel_path = get_excel_path(group_name)
    today = get_today_date()
    
    try:
        # 检查文件是否存在
        if os.path.exists(excel_path):
            # 尝试读取现有文件
            try:
                # 读取现有Excel文件，检查是否有今天的sheet
                import openpyxl
                wb = openpyxl.load_workbook(excel_path)
                
                if today in wb.sheetnames:
                    # 如果已存在今天的sheet，读取数据
                    existing_data = pd.read_excel(excel_path, sheet_name=today, engine='openpyxl')
                    print(f"成功读取现有数据")
                    
                    # 创建一个集合，存储现有数据中的 (发送人, 订餐内容) 组合
                    existing_keys = set()
                    for _, row in existing_data.iterrows():
                        try:
                            # 只使用发送人和订餐内容作为唯一键，不使用时间
                            key = (row['发送人'], row['订餐内容'])
                            existing_keys.add(key)
                        except Exception as e:
                            print(f"处理现有数据行时出错: {e}")
                    
                    # 过滤掉已存在的订单
                    filtered_orders = []
                    for order in orders:
                        # 只使用发送人和订餐内容作为唯一键，不使用时间
                        key = (order['发送人'], order['订餐内容'])
                        if key not in existing_keys:
                            filtered_orders.append(order)
                    
                    print(f"过滤后剩余新订单数量: {len(filtered_orders)}")
                    
                    # 如果没有新订单，直接返回
                    if not filtered_orders:
                        print("没有新订单需要添加")
                        return
                    
                    # 创建新的DataFrame
                    new_df = pd.DataFrame(filtered_orders)
                    
                    # 合并现有数据和新数据
                    combined_df = pd.concat([existing_data, new_df], ignore_index=True)
                    
                    # 保存回Excel，使用mode='a'追加模式，如果sheet已存在则替换
                    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        combined_df.to_excel(writer, sheet_name=today, index=False)
                        print(f"已更新今天({today})的订餐数据到: {excel_path}")
                else:
                    # 如果不存在今天的sheet，直接创建新sheet
                    new_df = pd.DataFrame(orders)
                    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a') as writer:
                        new_df.to_excel(writer, sheet_name=today, index=False)
                        print(f"已创建今天({today})的新sheet并保存订餐数据到: {excel_path}")
                
            except Exception as e:
                print(f"读取或更新Excel文件时出错: {e}")
                # 如果读取失败，创建新文件
                new_df = pd.DataFrame(orders)
                new_df.to_excel(excel_path, sheet_name=today, index=False, engine='openpyxl')
                print(f"由于错误，已创建新的Excel文件: {excel_path}")
        else:
            # 文件不存在，创建新文件
            new_df = pd.DataFrame(orders)
            new_df.to_excel(excel_path, sheet_name=today, index=False, engine='openpyxl')
            print(f"已创建新的Excel文件: {excel_path}")
        
        return True
    except Exception as e:
        print(f"保存Excel时出错: {e}")
        return False

def collect_orders(group_name):
    """收集订单信息"""
    print(f"开始收集 {group_name} 的订餐信息...")
    
    # 切换到目标群聊
    if not wx.ChatWith(who=group_name):
        print(f"找不到群聊: {group_name}")
        return []
    
    # 获取当前聊天窗口消息
    try:
        msgs = wx.GetAllMessage()
        if not msgs:
            print("没有获取到消息")
            return []
    except Exception as e:
        print(f"获取消息时出错: {e}")
        return []
    
    # 今天的日期
    today = datetime.now().strftime("%Y-%m-%d")
    today_datetime = datetime.now()
    
    # 收集今天的订餐信息
    orders = []
    # 用于去重的集合，存储 (发送人, 订餐内容) 元组
    unique_orders = set()
    
    print(f"开始处理 {len(msgs)} 条消息，筛选今天({today})的订餐信息...")
    
    for i, msg in enumerate(msgs):
        try:
            # 检查消息是否有必要的属性
            if not hasattr(msg, 'content'):
                print(f"消息 {i} 没有content属性")
                continue
                
            # 检查消息是否有time属性
            if not hasattr(msg, 'time') or not msg.time:
                # 如果没有time属性或time为空，默认视为当天消息
                msg_time = today_datetime.strftime("%Y-%m-%d %H:%M:%S")
                print(f"消息 {i} 没有有效的time属性，默认视为当天消息: {msg.content[:30]}...")
            else:
                msg_time = msg.time
            
            # 检查消息是否是今天的
            # 更灵活的日期检查，只要包含今天的日期就算
            if today not in msg_time:
                # 尝试其他可能的日期格式
                try:
                    # 尝试解析消息时间
                    msg_date = None
                    # 尝试几种常见的日期格式
                    date_formats = [
                        "%Y-%m-%d %H:%M:%S",
                        "%Y/%m/%d %H:%M:%S",
                        "%Y年%m月%d日 %H:%M:%S",
                        "%m-%d %H:%M:%S"  # 如果只有月日
                    ]
                    
                    for fmt in date_formats:
                        try:
                            if len(msg_time) >= 10:  # 确保字符串长度足够
                                parsed_date = datetime.strptime(msg_time, fmt)
                                if parsed_date.day == today_datetime.day and parsed_date.month == today_datetime.month:
                                    msg_date = parsed_date
                                    break
                        except ValueError:
                            continue
                    
                    # 如果无法解析日期或不是今天，跳过
                    if not msg_date:
                        continue
                except:
                    # 如果解析失败，默认视为当天消息（因为微信通常显示的是最近的消息）
                    print(f"无法解析消息时间，默认视为当天消息: {msg.content[:30]}...")
            
            # 获取发送人
            sender = getattr(msg, 'sender', '未知用户')
            
            # 跳过机器人自己发送的消息
            if sender == 'self':
                print(f"跳过机器人自己发送的消息: {msg.content[:30]}...")
                continue
            
            # 跳过包含"订餐汇总"的消息，这些是机器人发送的汇总信息
            if "订餐汇总" in getattr(msg, 'content', ''):
                print(f"跳过汇总消息: {msg.content[:30]}...")
                continue
            
            # 解析订餐信息
            order_content, order_count = parse_order_message(msg.content)
            if order_content and order_count:
                # 判断是否是人员名单格式（包含"人"字）
                is_people_list = "人" in msg.content
                
                # 创建唯一标识元组 - 只使用发送人和订餐内容，不使用时间
                order_key = (sender, order_content)
                
                # 检查是否已经存在相同的订单
                if order_key in unique_orders:
                    print(f"跳过重复订单: {sender} - {order_content}")
                    continue
                
                # 添加到去重集合
                unique_orders.add(order_key)
                
                # 添加到订单列表
                orders.append({
                    "发送人": sender,
                    "订餐内容": order_content,
                    "订餐份数": order_count,
                    "发送时间": msg_time,
                    "是否人员名单": is_people_list
                })
                print(f"收集到订单: {sender} - {order_content} - {order_count}份")
        except Exception as e:
            print(f"处理消息 {i} 时出错: {e}")
            continue
    
    print(f"收集到 {len(orders)} 条订餐信息")
    return orders

def generate_summary(orders, group_name, from_excel=False):
    """生成订餐统计信息
    
    Args:
        orders: 订单列表
        group_name: 群聊名称
        from_excel: 是否从Excel读取数据
    """
    if not orders:
        return "没有找到订餐信息"
    
    if from_excel:
        # 从Excel读取今天的数据
        excel_path = get_excel_path(group_name)
        today = get_today_date()
        
        try:
            if os.path.exists(excel_path):
                # 读取Excel中当天的所有订单
                excel_orders = pd.read_excel(excel_path, sheet_name=today, engine='openpyxl')
                print(f"从Excel读取到 {len(excel_orders)} 条订单记录")
                
                # 检查是否有人员名单类型的订单
                people_list_orders = excel_orders[excel_orders['是否人员名单'] == True]
                
                if not people_list_orders.empty:
                    # 修改：不再只使用最新的一条人员名单订单，而是累加所有人员名单订单的人数
                    total_people = 0
                    
                    # 遍历所有人员名单类型的订单
                    for _, order in people_list_orders.iterrows():
                        count = order['订餐份数']
                        # 从人员名单中提取人名并计数
                        names_text = order['订餐内容']
                        # 按空格分割人名并过滤空字符串
                        names_list = [name for name in names_text.split() if name.strip()]
                        # 计算实际人数
                        actual_count = len(names_list)
                        # 使用订单中的人数或实际人数中的较大值
                        final_count = count if count > actual_count else actual_count
                        # 累加人数
                        total_people += final_count
                    
                    return f"{today}{group_name}订餐汇总：共{total_people}人"
                else:
                    # 如果是普通订餐类型，统计所有订单
                    # 检查是否有发送人为self的订单，如果有则排除
                    filtered_orders = excel_orders[excel_orders['发送人'] != 'self']
                    total_people = len(filtered_orders['发送人'].unique())
                    total_portions = filtered_orders['订餐份数'].sum()
                    
                    return f"{today}{group_name}订餐汇总：共{total_people}人订餐，{total_portions}份"
            else:
                print(f"Excel文件不存在: {excel_path}")
        except Exception as e:
            print(f"从Excel读取订单时出错: {e}")
            # 如果读取Excel失败，回退到使用当前收集的订单
            print("回退到使用当前收集的订单")
    
    # 如果没有从Excel读取或读取失败，使用传入的orders
    # 检查是否有人员名单类型的订单
    people_list_orders = [order for order in orders if order.get("是否人员名单", False)]
    
    # 如果有人员名单类型的订单，累加所有人员名单的人数
    if people_list_orders:
        # 修改：不再只使用最新的一条人员名单订单，而是累加所有人员名单订单的人数
        total_people = 0
        
        # 遍历所有人员名单类型的订单
        for order in people_list_orders:
            count = order['订餐份数']
            # 从人员名单中提取人名并计数
            names_text = order['订餐内容']
            # 按空格分割人名并过滤空字符串
            names_list = [name for name in names_text.split() if name.strip()]
            # 计算实际人数
            actual_count = len(names_list)
            # 使用订单中的人数或实际人数中的较大值
            final_count = count if count > actual_count else actual_count
            # 累加人数
            total_people += final_count
        
        return f"{today}{group_name}订餐汇总：共{total_people}人"
    else:
        # 如果是普通订餐类型
        # 过滤掉发送人为self的订单
        filtered_orders = [order for order in orders if order.get("发送人") != 'self']
        total_people = len(set([order["发送人"] for order in filtered_orders]))
        total_portions = sum([order["订餐份数"] for order in filtered_orders])
        
        return f"{today}{group_name}订餐汇总：共{total_people}人订餐，{total_portions}份"

def send_summary(group_name):
    """发送每日汇总信息"""
    print(f"开始生成并发送 {group_name} 的每日汇总...")
    
    # 切换到目标群聊
    if not wx.ChatWith(who=group_name):
        print(f"找不到群聊: {group_name}")
        return
    
    # 获取今日订单
    orders = collect_orders(group_name)
    
    # 生成汇总消息
    summary = generate_summary(orders, group_name)
    at_person = AT_PERSONS.get(group_name, "布鲁布鲁")  # 获取该群聊对应的@人
    summary_msg = f"@{at_person} {summary}"
    
    # 发送汇总消息
    wx.SendMsg(msg=summary_msg, who=group_name)
    print(f"已发送汇总消息: {summary_msg}")
    
    # 保存到Excel
    save_to_excel(orders, group_name)

def handle_mention(msg, group_name):
    """处理@机器人的消息"""
    print(f"检测到@消息: {msg.content}")
    
    # 获取今日订单
    orders = collect_orders(group_name)
    
    # 生成汇总消息，从Excel中读取完整数据
    summary = generate_summary(orders, group_name, from_excel=True)
    
    # 回复@消息
    sender = getattr(msg, 'sender', '朋友')
    at_person = AT_PERSONS.get(group_name, "布鲁布鲁")  # 获取该群聊对应的@人
    reply_msg = f"@{sender} {summary}"
    
    print(f"准备回复消息: {reply_msg}")
    
    # 确保在正确的聊天窗口
    if not wx.ChatWith(who=group_name):
        print(f"无法切换到群聊: {group_name}")
        return
    
    # 发送回复消息
    try:
        # 直接发送消息
        wx.SendMsg(reply_msg)
        print(f"已回复@消息: {reply_msg}")
    except Exception as e:
        print(f"发送回复消息失败: {e}")
        # 尝试使用另一种方式发送
        try:
            # 重新切换到群聊并发送
            wx.ChatWith(who=group_name)
            time.sleep(1)  # 等待切换完成
            wx.SendMsg(reply_msg)
            print("使用替代方法发送回复成功")
        except Exception as e2:
            print(f"替代发送方法也失败: {e2}")
            # 最后尝试最简单的方式
            try:
                wx.SendMsg(summary)
                print("使用最简单方式发送成功")
            except Exception as e3:
                print(f"所有发送方法都失败: {e3}")

def check_time_for_summary():
    """检查是否到了发送汇总的时间"""
    current_hour = datetime.now().hour
    current_minute = datetime.now().minute
    
    # 在16:00左右发送汇总
    if current_hour == 16 and 0 <= current_minute <= 5:
        return True
    return False

def monitor_group():
    """监控群聊并定时处理"""
    print(f"开始监控群聊: {GROUP_NAMES}")
    
    # 获取最后一条消息ID，用于后续检查新消息
    last_msg_ids = {}
    # 添加一个集合来跟踪已处理过的@消息ID
    processed_at_msg_ids = set()
    
    try:
        for group_name in GROUP_NAMES:
            # 切换到目标群聊
            if not wx.ChatWith(who=group_name):
                print(f"找不到群聊: {group_name}")
                continue
            
            # 获取初始消息
            last_msgs = wx.GetAllMessage()
            print(f"初始化时获取到 {group_name} 的 {len(last_msgs) if last_msgs else 0} 条消息")
            
            # 打印所有初始消息的基本信息
            for i, msg in enumerate(last_msgs):
                try:
                    msg_id = getattr(msg, 'id', '无ID')
                    msg_content = getattr(msg, 'content', '无内容')
                    msg_sender = getattr(msg, 'sender', '未知发送者')
                    print(f"初始消息 {i}: ID={msg_id}, 发送者={msg_sender}, 内容={msg_content[:20]}...")
                    
                    # 将所有初始消息的ID添加到已处理集合中，避免重复处理
                    if hasattr(msg, 'id') and msg.id:
                        processed_at_msg_ids.add(msg.id)
                except Exception as e:
                    print(f"打印初始消息 {i} 信息时出错: {e}")
            
            if last_msgs and len(last_msgs) > 0 and hasattr(last_msgs[-1], 'id'):
                last_msg_ids[group_name] = last_msgs[-1].id
                print(f"设置 {group_name} 初始最后消息ID: {last_msg_ids[group_name]}")
            else:
                last_msg_ids[group_name] = None
                print(f"初始化时没有获取到 {group_name} 的消息ID")
    except Exception as e:
        print(f"获取初始消息时出错: {e}")
    
    # 用于跟踪上次发送汇总的日期
    last_summary_dates = {group_name: None for group_name in GROUP_NAMES}
    
    # 记录上次检查的时间
    last_check_time = time.time()
    
    print("开始监控循环...")
    while True:
        try:
            current_time = time.time()
            # 每隔一段时间打印一次心跳信息
            if current_time - last_check_time > 300:  # 每5分钟打印一次心跳
                print(f"监控心跳 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                last_check_time = current_time
            
                # 检查是否有新消息
                try:
                    print("获取当前消息...")
                    current_msgs_all = {}
                    for group_name in GROUP_NAMES:
                        try:
                            # 切换到目标群聊
                            if not wx.ChatWith(who=group_name):
                                print(f"找不到群聊: {group_name}")
                                continue
                                
                            current_msgs = wx.GetAllMessage()
                            current_msgs_all[group_name] = current_msgs
                            print(f"获取到 {group_name} 的 {len(current_msgs) if current_msgs else 0} 条消息")
                        except Exception as e:
                            print(f"获取 {group_name} 消息时出错: {e}")
                            current_msgs_all[group_name] = []
                except Exception as e:
                    print(f"获取消息时出错: {e}")
                    time.sleep(30)
                    continue
                    
                # 安全检查消息列表
                for group_name, current_msgs in current_msgs_all.items():
                    if not current_msgs:
                        print(f"{group_name} 没有获取到消息，等待下一轮检查")
                        continue
                    
                    # 打印最新几条消息的基本信息
                    print(f"{group_name} 最新几条消息:")
                    for i in range(min(3, len(current_msgs))):
                        try:
                            idx = len(current_msgs) - 1 - i
                            if idx >= 0:
                                msg = current_msgs[idx]
                                msg_id = getattr(msg, 'id', '无ID')
                                msg_content = getattr(msg, 'content', '无内容')
                                msg_sender = getattr(msg, 'sender', '未知发送者')
                                print(f"最新消息 {i}: ID={msg_id}, 发送者={msg_sender}, 内容={msg_content[:20]}...")
                        except Exception as e:
                            print(f"打印 {group_name} 最新消息信息时出错: {e}")
                    
                    # 检查是否有新消息
                    try:
                        has_new_message = False
                        if current_msgs and len(current_msgs) > 0:
                            if last_msg_ids.get(group_name) is None:
                                has_new_message = True
                                print(f"{group_name} 首次检测，视为有新消息")
                            elif hasattr(current_msgs[-1], 'id') and current_msgs[-1].id != last_msg_ids.get(group_name):
                                has_new_message = True
                                print(f"{group_name} 检测到新消息: 最新ID={current_msgs[-1].id}, 上次ID={last_msg_ids.get(group_name)}")
                            else:
                                print(f"{group_name} 没有检测到新消息")
                        
                        if has_new_message:
                            # 有新消息，更新last_msg_ids
                            new_msgs = []
                            
                            # 安全地处理消息
                            for msg in reversed(current_msgs):
                                if last_msg_ids.get(group_name) is not None and hasattr(msg, 'id') and msg.id == last_msg_ids.get(group_name):
                                    print(f"{group_name} 找到上次的最后消息ID: {last_msg_ids.get(group_name)}")
                                    break
                                new_msgs.insert(0, msg)
                            
                            print(f"{group_name} 共有 {len(new_msgs)} 条新消息")
                            
                            # 更新最后一条消息ID
                            if current_msgs and len(current_msgs) > 0 and hasattr(current_msgs[-1], 'id'):
                                last_msg_ids[group_name] = current_msgs[-1].id
                                print(f"{group_name} 更新最后消息ID为: {last_msg_ids[group_name]}")
                            
                            # 处理新消息
                            for i, msg in enumerate(new_msgs):
                                try:
                                    # 确保消息有content属性和id属性
                                    if not hasattr(msg, 'content'):
                                        print(f"{group_name} 新消息 {i} 没有content属性")
                                        continue
                                    
                                    # 检查消息是否有ID，如果没有则跳过
                                    if not hasattr(msg, 'id') or not msg.id:
                                        print(f"{group_name} 新消息 {i} 没有有效的ID")
                                        continue
                                    
                                    # 检查消息是否已处理过
                                    if msg.id in processed_at_msg_ids:
                                        print(f"{group_name} 新消息 {i} (ID={msg.id})已处理过，跳过")
                                        continue
                                    
                                    # 将消息ID添加到已处理集合
                                    processed_at_msg_ids.add(msg.id)
                                    
                                    msg_content = msg.content
                                    msg_sender = getattr(msg, 'sender', '未知用户')
                                    print(f"{group_name} 处理新消息 {i}: 发送者={msg_sender}, 内容={msg_content}")
                                        
                                    # 检查是否有人@机器人
                                    try:
                                        if is_bot_mentioned(msg_content):
                                            print(f"{group_name} 检测到@机器人消息: {msg_content}")
                                            handle_mention(msg, group_name)
                                    except Exception as e:
                                        print(f"{group_name} 处理@消息时出错: {e}")
                                    
                                    # 检查是否有新订餐
                                    try:
                                        order_content, order_count = parse_order_message(msg_content)
                                        if order_content and order_count:
                                            print(f"{group_name} 检测到订餐消息: {msg_content}")
                                            print(f"{group_name} 检测到新订餐: {msg_sender} - {order_content}，共{order_count}份")
                                            # 实时保存订单
                                            orders = collect_orders(group_name)
                                            save_to_excel(orders, group_name)
                                    except Exception as e:
                                        print(f"{group_name} 处理订餐消息时出错: {e}")
                                except Exception as e:
                                    print(f"{group_name} 处理新消息 {i} 时出错: {e}")
                    except Exception as e:
                        print(f"{group_name} 处理消息列表时出错: {e}")
                
                # 检查是否需要执行定时任务
                try:
                    if HAS_SCHEDULE:
                        schedule.run_pending()
                    else:
                        # 自定义定时逻辑
                        today = datetime.now().date()
                        for group_name in GROUP_NAMES:
                            if check_time_for_summary() and last_summary_dates.get(group_name) != today:
                                print(f"{group_name} 到达汇总时间，开始发送汇总")
                                send_summary(group_name)
                                last_summary_dates[group_name] = today
                except Exception as e:
                    print(f"执行定时任务时出错: {e}")
            
            # 休眠时间，设置为5分钟
            time.sleep(300)  # 5分钟 = 300秒
            
        except Exception as e:
            print(f"监控过程中出错: {e}")
            time.sleep(300)  # 出错后等待5分钟再重试

# 如果有schedule模块，设置每天16:00发送汇总
if HAS_SCHEDULE:
    schedule.every().day.at("16:00").do(lambda: [send_summary(group_name) for group_name in GROUP_NAMES])

if __name__ == "__main__":
    try:
        print("订餐统计机器人已启动...")
        print(f"机器人名称: {BOT_NAME}")
        
        # 首次运行时收集并保存当前订单
        for group_name in GROUP_NAMES:
            orders = collect_orders(group_name)
            save_to_excel(orders, group_name)
        
        # 开始监控
        monitor_group()
    except Exception as e:
        print(f"程序运行出错: {e}")
