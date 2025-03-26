# -*- coding: utf-8 -*-
"""
文档批量替换工具配置文件
在这里设置替换规则和文件夹路径
"""

# 替换规则：设置要替换的文本和替换后的文本
# 格式：旧文本: 新文本
# 支持正则表达式，例如：r"\d{4}-\d{2}-\d{2}": "<日期>"
REPLACE_RULES = {
    "2019": "2023",
    "2020": "2024",
    # 你可以继续添加更多的替换规则...
}

# 源文件夹路径：存放原始文档的文件夹
# 注意：使用 r 前缀表示原始字符串，避免路径中的反斜杠问题
SOURCE_FOLDER = r"E:\template"

# 输出文件夹路径：处理后的文档将保存在这里
# 如果文件夹不存在，程序会自动创建
OUTPUT_FOLDER = r"E:\output"

# 日志设置
LOG_FILE = "processing.log"  # 日志文件名
LOG_LEVEL = "INFO"  # 日志级别: DEBUG, INFO, WARNING, ERROR, CRITICAL

# 高级设置（一般情况下不需要修改）
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

"""
说明：
1. 本程序会禁用所有Office弹窗和警告，确保批量处理过程中不会中断
2. 如果某些文件处理失败，可能是由于文件已打开、格式异常或权限问题
3. 为防止意外错误，处理前请备份重要文件
4. 程序会自动创建临时文件，处理完成后会自动删除
""" 