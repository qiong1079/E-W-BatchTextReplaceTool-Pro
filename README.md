# 文档批量替换工具

一个强大的批量文本替换工具，支持Word和Excel文档，能保留原始格式并进行智能替换。

## 功能特点

- 支持多种文档格式：`.docx`, `.doc`, `.xlsx`, `.xls`
- 批量处理大量文件，提高工作效率
- 智能替换，保留原始文档格式
- 支持替换文档的各个部分：正文、页眉页脚、表格、文档属性等
- 精确统计替换数量，提供详细日志
- 自动禁用所有Office应用程序弹窗，实现真正的无人值守批处理
- 多种替换方法，确保替换成功
- 安全的文件操作机制，包含自动重试和错误恢复
- 详细的日志记录，方便排查问题

## 安装要求

- Python 3.7 或更高版本
- Microsoft Office（Word和Excel）已安装
- 安装必要的Python库：
  ```
  pip install pywin32
  ```

## 使用方法

1. 克隆或下载本仓库
2. 配置 `config.py` 文件：
   - 设置替换规则
   - 指定源文件夹和输出文件夹
   - 调整其他选项（如日志级别、备份选项等）
3. 运行主程序：
   ```
   python docx_formatted_replace.py
   ```

## 配置选项

在 `config.py` 文件中，您可以设置以下选项：

```python
# 替换规则
REPLACE_RULES = {
    "旧文本1": "新文本1",
    "旧文本2": "新文本2",
    # 可以添加更多规则
}

# 文件夹路径
SOURCE_FOLDER = r"E:\文档\原始文件"  # 源文件夹路径
OUTPUT_FOLDER = r"E:\文档\处理后文件"  # 输出文件夹路径

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
```

## 版本历史

### V1.9
- 修复Word文档打开时的兼容性问题，删除不支持的"WithWindow"参数
- 增强与不同版本Word应用程序的兼容性，确保文档正常打开
- 优化错误处理，在参数不兼容时能够正常处理文档

### V1.8
- 修复程序运行时可能出现的"unsupported format string passed to tuple.__format__"错误
- 优化Word文档处理函数的返回值处理，确保正确计数替换次数
- 增强错误处理机制，提高程序运行稳定性
- 完善日志记录，便于排查问题

### V1.7
- 修复关键问题：解决部分Word文档只计数不进行实际替换的错误
- 引入替换验证机制，确保计数与实际替换一致
- 增加多种替换方法，通过不同层次确保文本被正确替换
  - 查找替换法 - 逐项精确替换
  - 段落级替换 - 确保段落内容完整替换
  - 直接内容替换 - 处理复杂布局内容
- 针对特殊字符和格式添加了更强的兼容性
- 优化替换失败时的强制替换机制
- 更详细的日志记录，精确追踪每次替换操作

### V1.6
- 增强Word文档封面底部日期处理能力，解决公文封面底部日期无法替换的问题
- 针对中文格式日期（如"2020年3月22日"）增加专门处理方法
- 支持多种日期位置的智能定位：文档底部、表格中、页脚、浮动文本框等
- 针对中文公文格式特别优化，确保封面日期能被准确替换
- 优先处理底部最后一个表格，针对常见公文格式做了特殊处理
- 自动识别标准中文日期格式，精确定位并替换

### V1.5
- 增强Word文档字段(Field)处理功能，解决首页特殊字段无法替换的问题
- 支持多种域类型的替换，包括普通域、DocProperty域、表单域等
- 增加对域代码(Field Code)和域结果(Field Result)的双重处理
- 自动锁定已替换的域，避免内容被自动更新覆盖
- 添加域刷新功能，确保替换后域内容立即生效
- 特别处理文档信息域，如标题、作者、公司等与文档属性关联的域

### V1.4
- 增强Word文档封面处理功能，解决封面文本无法替换的问题
- 支持多种特殊封面元素的替换：文本框、WordArt、SmartArt等
- 增加对文档第一页内容的专门处理
- 提供多种封面处理方法，确保不同类型的封面都能被正确处理
- 对页眉页脚中的特殊文本框增加支持

### V1.3
- 全面增强Word文档处理能力，解决文档替换报错问题
- 多种文档内容获取方法，增强兼容性
- 对不同Word文档格式的智能处理与转换
- 增加了多种保存策略，确保文档能够被正确保存
- 更加详细的错误处理和日志记录
- 临时文件处理采用更安全的随机命名机制
- 表格内容专门处理，确保表格中的文本也能被正确替换

### V1.2
- 完全禁用所有Office应用程序弹窗，实现真正的无人值守批处理
- 使用临时文件处理机制，避免文件锁定问题
- 优化错误处理，提高程序稳定性
- 新增Excel和Word特有设置，可针对不同文档类型进行精细化控制
- 增强日志记录，方便排查问题
- 修复替换计数问题，准确统计实际替换的次数而不是规则数量
- 增加更详细的替换过程日志，包括每个工作表、页眉页脚中的替换次数

## 注意事项

- 处理前请备份重要文件
- 使用 `r` 前缀设置路径，避免路径问题，例如：`r"C:\文档\文件"`
- 首次使用时，建议先用少量文件测试
- 程序会自动创建临时文件夹，处理结束后会自动清理

## 许可证

[MIT](LICENSE)
