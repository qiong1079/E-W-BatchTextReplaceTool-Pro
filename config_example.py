# 配置文件示例
# 使用方法：复制此文件并重命名为 config.py，然后根据需要修改

# 替换规则 - 字典格式，键为要查找的文本，值为替换后的文本
REPLACE_RULES = {
    "原始文本1": "替换后文本1",
    "原始文本2": "替换后文本2",
    # 可以根据需要添加更多规则
    # "公司旧名称": "公司新名称",
    # "旧产品编号": "新产品编号",
}

# 文件夹路径设置
SOURCE_FOLDER = r"E:\源文件夹"  # 请修改为您的源文件夹路径
OUTPUT_FOLDER = r"E:\输出文件夹"  # 请修改为您的输出文件夹路径

# 日志设置
LOG_FILE = "processing.log"  # 日志文件名
LOG_LEVEL = "INFO"  # 日志级别: DEBUG, INFO, WARNING, ERROR, CRITICAL

# 界面与操作控制
SHOW_PROGRESS = True  # 是否显示进度信息
BACKUP_ORIGINAL = False  # 是否在处理前备份原始文档
MAX_RETRIES = 3  # 处理文件失败时的最大重试次数
DISABLE_ALERTS = True  # 是否禁用所有Office应用程序弹窗

# Excel特有设置
EXCEL_SETTINGS = {
    'UPDATE_LINKS': False,  # 是否更新链接
    'SHOW_WARNINGS': False,  # 是否显示警告
    'AUTO_MACROS': False,    # 是否允许宏
}

# Word特有设置
WORD_SETTINGS = {
    'UPDATE_LINKS': False,    # 是否更新链接 
    'SHOW_WARNINGS': False,   # 是否显示警告
    'CONFIRM_CONVERSIONS': False,  # 是否确认格式转换
}

# 高级设置（一般情况下不需要修改）
COMPRESSION_LEVEL = 9  # 文档压缩级别 (0-9)，9为最高压缩比
AUTOFIT_TABLES = True  # 表格替换后是否自动调整列宽 