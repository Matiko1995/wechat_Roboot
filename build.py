
import os
import sys
import PyInstaller.__main__

# 获取当前目录
current_dir = os.path.dirname(os.path.abspath(__file__))

# 定义图标文件路径（如果有的话）
# icon_path = os.path.join(current_dir, 'icon.ico')

# 定义打包参数
args = [
    'gui_app.py',  # 主程序文件
    '--name=微信订餐机器人',  # 应用名称
    '--onefile',  # 打包成单个可执行文件
    '--noconsole',  # 不显示控制台窗口
    # f'--icon={icon_path}',  # 应用图标（如果有的话）
    '--add-data=index.py;.',  # 添加额外的Python文件
    '--hidden-import=pandas',  # 添加隐式导入的模块
    '--hidden-import=openpyxl',
    '--hidden-import=wxauto',
]


# 运行PyInstaller
PyInstaller.__main__.run(args)

print("打包完成！")
