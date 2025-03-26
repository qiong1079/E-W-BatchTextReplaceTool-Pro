import os
import sys
import time
import logging
import win32com.client as win32
import shutil
from datetime import datetime

# 导入配置文件
try:
    from config import (
        REPLACE_RULES, SOURCE_FOLDER, OUTPUT_FOLDER, 
        LOG_FILE, LOG_LEVEL, SHOW_PROGRESS,
        BACKUP_ORIGINAL, MAX_RETRIES, DISABLE_ALERTS,
        EXCEL_SETTINGS, WORD_SETTINGS
    )
except ImportError:
    print("错误: 找不到配置文件 (config.py)")
    print("请确保配置文件 config.py 存在于程序目录中。")
    sys.exit(1)
except Exception as e:
    print(f"错误: 加载配置文件时出错: {str(e)}")
    sys.exit(1)

# 配置日志
log_level = getattr(logging, LOG_LEVEL, logging.INFO)
logging.basicConfig(
    filename=LOG_FILE,
    level=log_level,
    format='[%(asctime)s] %(levelname)s: %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

def get_word_format_constant(file_ext):
    """
    获取Word文档格式的常数值
    
    Args:
        file_ext: 文件扩展名
        
    Returns:
        int: Word格式常数
    """
    # 文件格式常量定义
    # 参考: https://docs.microsoft.com/en-us/office/vba/api/word.wdsaveformat
    format_map = {
        '.doc': 0,     # wdFormatDocument
        '.docx': 16,   # wdFormatDocumentDefault (docx)
        '.docm': 13,   # wdFormatXMLDocumentMacroEnabled
        '.dotx': 17,   # wdFormatXMLTemplate
        '.pdf': 17,    # wdFormatPDF
        '.rtf': 6,     # wdFormatRTF
        '.txt': 2,     # wdFormatText
        '.html': 8,    # wdFormatHTML
        '.xml': 11,    # wdFormatXML
    }
    return format_map.get(file_ext.lower(), 16)  # 默认返回docx格式

def replace_in_word(doc_path, output_path, retries=0):
    """在Word文档中替换文本，并保留格式"""
    if retries > MAX_RETRIES:
        logging.error(f"处理文件 {doc_path} 失败，超过最大重试次数")
        return False, "超过最大重试次数", 0
    
    word = None
    doc = None
    temp_output_path = None
    replacement_counts = {}  # 用于记录每个规则的替换次数
    total_replacements = 0  # 总替换次数
    success = False
    error_msg = ""
    
    try:
        # 生成临时输出文件路径
        temp_dir = os.path.join(os.path.dirname(output_path), 'temp')
        os.makedirs(temp_dir, exist_ok=True)
        
        # 使用时间戳和随机数创建唯一的临时文件名
        random_suffix = f"{int(time.time())}-{os.getpid()}-{hash(doc_path) % 10000}"
        file_name = os.path.basename(doc_path)
        file_base, file_ext = os.path.splitext(file_name)
        temp_output_path = os.path.join(temp_dir, f"{file_base}_temp_{random_suffix}{file_ext}")
        
        logging.info(f"处理文件: {doc_path}")
        
        # 创建Word应用实例
        word = win32.Dispatch("Word.Application")
        word.Visible = False  # 设置为不可见
        word.DisplayAlerts = not DISABLE_ALERTS  # 是否禁用警告
        
        # 设置Word选项
        if WORD_SETTINGS['UPDATE_LINKS'] is not None:
            word.Options.UpdateLinksAtOpen = WORD_SETTINGS['UPDATE_LINKS']
        if WORD_SETTINGS['CONFIRM_CONVERSIONS'] is not None:
            word.Options.ConfirmConversions = WORD_SETTINGS['CONFIRM_CONVERSIONS']
        
        # 打开文档
        doc = word.Documents.Open(
            doc_path,
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False
        )
        
        # 添加一个替换确认机制，解决只计数不替换的问题
        for old_text, new_text in REPLACE_RULES.items():
            replacement_counts[old_text] = 0
            
            # 检查文档是否包含需要替换的文本
            if old_text in doc.Content.Text:
                # 收集替换前的样本，用于后续验证
                sample_positions = []
                try:
                    doc_text = doc.Content.Text
                    start_pos = 0
                    max_samples = 5  # 最多收集5个样本
                    samples_count = 0
                    
                    while samples_count < max_samples:
                        pos = doc_text.find(old_text, start_pos)
                        if pos == -1:
                            break
                        
                        sample_positions.append(pos)
                        start_pos = pos + len(old_text)
                        samples_count += 1
                        
                    logging.debug(f"找到文本 '{old_text}' 的 {len(sample_positions)} 个样本位置")
                except Exception as e:
                    logging.debug(f"收集样本位置时出错: {str(e)}")
                
                # 执行替换操作 - 使用多种方法确保替换成功
                logging.info(f"开始替换文本: '{old_text}' -> '{new_text}'")
                
                # 方法1: 使用Find替换 - 最可靠的方法
                try:
                    # 选择整个文档
                    word.Selection.HomeKey(Unit=6)  # 6=wdStory, 移到文档开始
                    
                    # 设置查找和替换
                    find_obj = word.Selection.Find
                    find_obj.ClearFormatting()
                    find_obj.Text = old_text
                    find_obj.Replacement.ClearFormatting()
                    find_obj.Replacement.Text = new_text
                    
                    # 执行替换
                    replace_count = 0
                    found = find_obj.Execute(
                        FindText=old_text,
                        MatchCase=False,
                        MatchWholeWord=False,
                        MatchWildcards=False,
                        MatchSoundsLike=False,
                        MatchAllWordForms=False,
                        Forward=True,
                        Wrap=1  # wdFindContinue
                    )
                    
                    while found:
                        # 手动执行替换，确保每次替换成功
                        word.Selection.Text = new_text
                        replace_count += 1
                        
                        # 继续查找下一个
                        found = find_obj.Execute(
                            FindText=old_text,
                            MatchCase=False,
                            MatchWholeWord=False,
                            MatchWildcards=False,
                            MatchSoundsLike=False,
                            MatchAllWordForms=False,
                            Forward=True,
                            Wrap=1  # wdFindContinue
                        )
                    
                    logging.info(f"Find方法替换了 {replace_count} 处")
                    replacement_counts[old_text] += replace_count
                except Exception as find_err:
                    logging.warning(f"使用Find方法替换失败: {str(find_err)}")
                
                # 方法2: 直接文本替换 - 有些情况下更有效
                try:
                    replaced = 0
                    content_text = doc.Content.Text
                    if old_text in content_text:
                        # 计算原始匹配次数
                        original_count = content_text.count(old_text)
                        
                        # 执行替换
                        new_content = content_text.replace(old_text, new_text)
                        doc.Content.Text = new_content
                        
                        # 计算替换数量差值
                        replaced = original_count
                        logging.info(f"直接替换文档内容成功，替换了 {replaced} 处")
                        
                        # 如果直接替换导致替换计数增加，更新计数
                        if replaced > replacement_counts[old_text]:
                            replacement_counts[old_text] = replaced
                except Exception as text_err:
                    logging.warning(f"直接文本替换失败: {str(text_err)}")
                
                # 方法3: 段落级别替换 - 更细粒度的控制
                try:
                    para_replaced = 0
                    for i in range(1, doc.Paragraphs.Count + 1):
                        try:
                            para = doc.Paragraphs(i)
                            para_text = para.Range.Text
                            
                            if old_text in para_text:
                                # 计数并替换
                                count_in_para = para_text.count(old_text)
                                new_para_text = para_text.replace(old_text, new_text)
                                para.Range.Text = new_para_text
                                para_replaced += count_in_para
                                logging.debug(f"在段落 {i} 中替换了 {count_in_para} 处")
                        except Exception as para_err:
                            logging.debug(f"处理段落 {i} 时出错: {str(para_err)}")
                    
                    logging.info(f"段落级替换完成，共替换了 {para_replaced} 处")
                    
                    # 更新替换计数（取最大值）
                    if para_replaced > replacement_counts[old_text]:
                        replacement_counts[old_text] = para_replaced
                except Exception as para_method_err:
                    logging.warning(f"段落级替换失败: {str(para_method_err)}")
                
                # 验证替换是否成功
                verified = False
                try:
                    # 1. 检查文档中是否还存在原始文本
                    updated_content = doc.Content.Text
                    if old_text not in updated_content:
                        verified = True
                        logging.info(f"验证成功：文档中不再包含文本 '{old_text}'")
                    else:
                        # 2. 验证之前找到的样本位置
                        matches_replaced = 0
                        for pos in sample_positions:
                            try:
                                # 检查样本位置文本是否被替换
                                check_range = doc.Range(doc.Content.Start + pos, doc.Content.Start + pos + len(new_text))
                                check_text = check_range.Text
                                
                                if check_text == new_text or old_text not in check_text:
                                    matches_replaced += 1
                            except:
                                pass
                        
                        # 如果大部分样本位置已被替换，认为验证成功
                        if matches_replaced > 0 and matches_replaced >= len(sample_positions) / 2:
                            verified = True
                            logging.info(f"验证成功：{matches_replaced}/{len(sample_positions)} 个样本位置已被替换")
                        else:
                            logging.warning(f"验证失败：只有 {matches_replaced}/{len(sample_positions)} 个样本位置被替换")
                except Exception as verify_err:
                    logging.warning(f"验证替换时出错: {str(verify_err)}")
                
                # 如果验证失败但有替换计数，尝试最后的强制替换
                if not verified and replacement_counts[old_text] > 0:
                    logging.warning(f"检测到替换可能未成功执行，尝试最终的强制替换方法")
                    
                    try:
                        # 使用替代方法强制替换
                        replacement_counts[old_text] = 0  # 重置计数
                        
                        # 方法4: 使用文档对象模型的Replace方法
                        word.Selection.HomeKey(Unit=6)  # 移到文档开始
                        word.Selection.Find.ClearFormatting()
                        word.Selection.Find.Replacement.ClearFormatting()
                        
                        # 执行全文档替换
                        replace_all_count = word.Selection.Find.Execute(
                            FindText=old_text,
                            ReplaceWith=new_text,
                            Replace=2,  # wdReplaceAll
                            Forward=True,
                            MatchCase=False,
                            MatchWholeWord=False,
                            MatchWildcards=False,
                            MatchSoundsLike=False,
                            MatchAllWordForms=False
                        )
                        
                        if replace_all_count:
                            logging.info(f"强制替换成功，替换次数: {replace_all_count}")
                            replacement_counts[old_text] = replace_all_count
                            verified = True
                    except Exception as force_err:
                        logging.error(f"强制替换失败: {str(force_err)}")
                
                # 如果还是失败，标记为0替换，避免错误计数
                if not verified and replacement_counts[old_text] > 0:
                    logging.warning(f"文本 '{old_text}' 替换验证失败，实际可能未替换，重置计数为0")
                    replacement_counts[old_text] = 0
            else:
                logging.debug(f"文档中不包含文本 '{old_text}'")
        
        # 计算总替换次数
        for count in replacement_counts.values():
            total_replacements += count
        
        # 保存文档
        try:
            doc.SaveAs2(temp_output_path)
            success = True
        except Exception as save_err:
            logging.warning(f"SaveAs2保存失败: {str(save_err)}，尝试其他方法")
            
            try:
                doc.SaveAs(temp_output_path)
                success = True
            except Exception as save_err2:
                logging.warning(f"SaveAs保存失败: {str(save_err2)}，尝试复制方法")
                
                try:
                    shutil.copy2(doc_path, temp_output_path)
                    doc.Save()
                    shutil.copy2(doc_path, temp_output_path)
                    success = True
                except Exception as save_err3:
                    logging.error(f"所有保存方法都失败: {str(save_err3)}")
                    error_msg = f"保存文件失败: {str(save_err3)}"
                    success = False
        
        # 关闭文档和Word应用
        if doc:
            try:
                doc.Close(0)  # 不保存更改
            except:
                pass
        
        if word:
            try:
                word.Quit()
            except:
                pass
        
        # 如果处理成功，移动文件到目标位置
        if success:
            try:
                # 确保输出目录存在
                output_dir = os.path.dirname(output_path)
                if not os.path.exists(output_dir):
                    os.makedirs(output_dir, exist_ok=True)
                
                # 如果目标文件已存在，先删除
                if os.path.exists(output_path):
                    try:
                        os.remove(output_path)
                    except Exception as remove_err:
                        logging.warning(f"删除现有输出文件失败: {str(remove_err)}")
                
                # 移动临时文件到最终位置
                safe_file_operation(temp_output_path, output_path)
                
                if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                    logging.info(f"文档处理完成: {doc_path} -> {output_path}")
                    logging.info(f"总共进行了 {total_replacements} 次替换")
                    
                    # 列出每条规则的替换次数
                    for rule, count in replacement_counts.items():
                        if count > 0:
                            logging.debug(f"  - 规则 '{rule}': {count} 次替换")
                else:
                    raise Exception("移动临时文件到输出位置失败")
            except Exception as move_err:
                logging.error(f"移动处理后的文件失败: {str(move_err)}")
                error_msg = f"移动文件失败: {str(move_err)}"
                success = False
                
                # 尝试重新处理
                return replace_in_word(doc_path, output_path, retries + 1)
        
        # 清理临时文件
        if temp_output_path and os.path.exists(temp_output_path):
            try:
                os.remove(temp_output_path)
            except:
                logging.debug(f"无法删除临时文件: {temp_output_path}")
        
        return success, error_msg, total_replacements
    
    except Exception as e:
        logging.error(f"处理Word文档时出错: {str(e)}")
        
        # 关闭文档和Word应用
        if doc:
            try:
                doc.Close(0)  # 不保存更改
            except:
                pass
        
        if word:
            try:
                word.Quit()
            except:
                pass
        
        # 清理临时文件
        if temp_output_path and os.path.exists(temp_output_path):
            try:
                os.remove(temp_output_path)
            except:
                logging.debug(f"无法删除临时文件: {temp_output_path}")
        
        # 尝试重新处理
        return replace_in_word(doc_path, output_path, retries + 1)

def replace_in_excel(excel_path, output_path, retries=0):
    """ 使用微软 Excel API 替换文本 """
    excel = None
    workbook = None
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False  # 不显示 Excel 窗口
        
        # 禁用所有提示和警告
        if DISABLE_ALERTS:
            excel.DisplayAlerts = False
            excel.AskToUpdateLinks = EXCEL_SETTINGS['UPDATE_LINKS']
            excel.AlertBeforeOverwriting = False
            excel.FeatureInstall = 0  # msoFeatureInstallNone
            excel.EnableEvents = False  # 禁用所有自动事件
            excel.ScreenUpdating = False  # 禁用屏幕更新
        
        logging.info(f"打开Excel: {excel_path}")
        workbook = excel.Workbooks.Open(
            excel_path, 
            UpdateLinks=1 if EXCEL_SETTINGS['UPDATE_LINKS'] else 0,  # 0=不更新, 1=更新
            ReadOnly=False, 
            AddToMru=False, 
            CorruptLoad=2  # 忽略无效记录，尝试恢复
        )

        total_replace_count = 0  # 总替换次数
        
        for sheet in workbook.Sheets:
            sheet.Activate()  # 确保当前工作表被激活
            
            # 获取工作表的使用范围
            used_range = sheet.UsedRange
            if used_range is None or used_range.Cells.Count == 0:
                continue  # 跳过空工作表
                
            logging.info(f"处理工作表: {sheet.Name}, 包含 {used_range.Cells.Count} 个单元格")
            
            for old_text, new_text in REPLACE_RULES.items():
                try:
                    # 替换前先获取匹配数量
                    # 创建一个临时数组存储所有匹配单元格的值
                    cell_values = []
                    for row in range(1, used_range.Rows.Count + 1):
                        for col in range(1, used_range.Columns.Count + 1):
                            try:
                                cell_value = used_range.Cells(row, col).Value
                                if cell_value and old_text in str(cell_value):
                                    cell_values.append(str(cell_value))
                            except:
                                pass
                    
                    # 计算将会替换的次数
                    matches_in_sheet = 0
                    for value in cell_values:
                        # 计算每个单元格中旧文本出现的次数
                        if isinstance(value, str):
                            matches_in_sheet += value.count(old_text)
                    
                    if matches_in_sheet > 0:
                        logging.info(f"工作表 '{sheet.Name}' 中找到 '{old_text}' {matches_in_sheet} 次")
                        
                    # 执行替换
                    sheet.Cells.Replace(
                        What=old_text, 
                        Replacement=new_text, 
                        LookAt=2,  # 2 = xlPart (匹配部分文本)
                        SearchOrder=1,  # 1 = xlByRows
                        MatchCase=False
                    )
                    
                    # 增加替换计数
                    total_replace_count += matches_in_sheet
                    
                except Exception as sheet_err:
                    logging.warning(f"在工作表'{sheet.Name}'中替换'{old_text}'时出错: {str(sheet_err)}")

        # 处理工作簿属性
        try:
            for prop_name in ["Title", "Subject", "Keywords", "Comments"]:
                try:
                    prop_value = getattr(workbook.BuiltInDocumentProperties, prop_name).Value
                    if prop_value:
                        for old_text, new_text in REPLACE_RULES.items():
                            if old_text in prop_value:
                                # 计算属性中的替换次数
                                matches_in_prop = prop_value.count(old_text)
                                new_prop_value = prop_value.replace(old_text, new_text)
                                setattr(workbook.BuiltInDocumentProperties, prop_name, new_prop_value)
                                total_replace_count += matches_in_prop
                except:
                    pass  # 忽略单个属性错误
        except:
            logging.warning(f"无法处理工作簿属性")

        logging.info(f"共替换了 {total_replace_count} 处内容，正在保存: {output_path}")
        # 使用更详细的SaveAs参数，指定格式以避免弹出保存对话框
        workbook.SaveAs(
            output_path, 
            AddToMru=False, 
            FileFormat=excel.DefaultSaveFormat,
            ConflictResolution=2  # 2=覆盖
        )
        return total_replace_count
    except Exception as e:
        logging.error(f"Excel替换错误: {str(e)}")
        if retries < MAX_RETRIES:
            logging.info(f"重试 ({retries + 1}/{MAX_RETRIES})...")
            return replace_in_excel(excel_path, output_path, retries + 1)
        raise
    finally:
        if workbook:
            try:
                workbook.Close(SaveChanges=False)
            except:
                pass
        if excel:
            try:
                # 恢复Excel设置并退出
                if DISABLE_ALERTS:
                    excel.DisplayAlerts = True
                    excel.EnableEvents = True
                    excel.ScreenUpdating = True
                excel.Quit()
            except:
                pass

def backup_file(file_path):
    """ 备份原始文件 """
    backup_dir = os.path.join(os.path.dirname(file_path), "备份_" + datetime.now().strftime("%Y%m%d"))
    if not os.path.exists(backup_dir):
        os.makedirs(backup_dir)
    
    filename = os.path.basename(file_path)
    backup_path = os.path.join(backup_dir, filename)
    
    try:
        shutil.copy2(file_path, backup_path)
        logging.info(f"已备份: {file_path} -> {backup_path}")
        return True
    except Exception as e:
        logging.error(f"备份失败: {str(e)}")
        return False

def safe_file_operation(source, dest, operation="move", retries=3):
    """
    安全地执行文件操作 (复制或移动)，包含重试机制
    
    Args:
        source: 源文件路径
        dest: 目标文件路径
        operation: "move" 或 "copy"
        retries: 重试次数
    
    Returns:
        bool: 操作是否成功
    """
    for attempt in range(retries):
        try:
            if operation == "move":
                # 如果目标文件已存在，先删除
                if os.path.exists(dest):
                    try:
                        os.remove(dest)
                        logging.info(f"已删除现有目标文件: {dest}")
                    except Exception as del_err:
                        logging.warning(f"删除目标文件失败: {str(del_err)}")
                
                shutil.move(source, dest)
                logging.info(f"成功移动文件: {source} -> {dest}")
            else:  # 复制操作
                shutil.copy2(source, dest)
                logging.info(f"成功复制文件: {source} -> {dest}")
            return True
        except Exception as e:
            logging.warning(f"文件{operation}操作失败 (尝试 {attempt+1}/{retries}): {str(e)}")
            # 短暂延迟后重试
            time.sleep(0.5)
    
    logging.error(f"文件{operation}操作失败，已达到最大重试次数")
    return False

def batch_process():
    """ 处理文件夹中的所有 Word 和 Excel 文件 """
    start_time = time.time()
    
    # 确保输出目录存在
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)
        logging.info(f"创建输出目录: {OUTPUT_FOLDER}")

    # 创建临时目录
    temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp_files")
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)
        logging.info(f"创建临时目录: {temp_dir}")

    # 处理状态统计
    total_files = 0
    success_files = 0
    failed_files = 0
    skipped_files = 0
    total_replacements = 0  # 总替换次数
    
    # 获取要处理的文件列表
    file_list = [f for f in os.listdir(SOURCE_FOLDER) 
                if os.path.isfile(os.path.join(SOURCE_FOLDER, f))]
    
    # 如果没有文件可处理
    if not file_list:
        print("没有发现可处理的文件。请检查源文件夹。")
        logging.warning(f"源文件夹为空: {SOURCE_FOLDER}")
        return
    
    # 结果表格的表头
    print("\n" + "=" * 80)
    print(f"{'文件名':<40} {'替换次数':<10} {'状态':<10} {'耗时(秒)':<10}")
    print("-" * 80)

    # 遍历源文件夹
    for idx, filename in enumerate(file_list):
        file_start_time = time.time()
        base, ext = os.path.splitext(filename)
        input_path = os.path.join(SOURCE_FOLDER, filename)
        
        # 使用临时文件路径作为中间处理
        timestamp = int(time.time())
        random_suffix = os.urandom(4).hex()  # 添加随机后缀避免文件名冲突
        temp_output_path = os.path.join(temp_dir, f"temp_{timestamp}_{random_suffix}_{filename}")
        final_output_path = os.path.join(OUTPUT_FOLDER, filename)

        # 显示进度
        if SHOW_PROGRESS:
            progress = (idx + 1) / len(file_list) * 100
            print(f"\r处理进度: {progress:.1f}% [{idx+1}/{len(file_list)}]", end="")
        
        logging.info(f"开始处理文件: {filename}")
        total_files += 1
        replace_count = 0

        # 如果启用备份，先备份文件
        if BACKUP_ORIGINAL:
            backup_file(input_path)

        try:
            # 首先检查文件是否实际存在
            if not os.path.isfile(input_path):
                logging.error(f"文件不存在: {input_path}")
                raise FileNotFoundError(f"找不到文件: {input_path}")
                
            # 检查文件大小，避免处理空文件
            file_size = os.path.getsize(input_path)
            if file_size == 0:
                logging.warning(f"跳过空文件: {filename}")
                status = "⚠️ 跳过"
                replace_count = "--"
                skipped_files += 1
                continue  # 跳过空文件
                
            if ext.lower() in ['.xls', '.xlsx']:
                # Excel文件处理
                try:
                    # 先处理到临时文件
                    replace_count = replace_in_excel(input_path, temp_output_path)
                    
                    # 安全地移动到最终位置
                    if safe_file_operation(temp_output_path, final_output_path, "move"):
                        success_files += 1
                        status = "✅ 成功"
                        total_replacements += replace_count
                    else:
                        # 移动失败，尝试复制
                        if safe_file_operation(input_path, final_output_path, "copy"):
                            logging.warning(f"无法移动临时文件，已直接复制原文件: {filename}")
                            success_files += 1
                            status = "⚠️ 部分成功"
                            total_replacements += replace_count
                        else:
                            raise Exception("无法保存处理后的文件")
                except Exception as excel_err:
                    logging.error(f"处理Excel文件时出错: {str(excel_err)}")
                    raise
                    
            elif ext.lower() in ['.doc', '.docx']:
                # Word文档处理
                try:
                    # 先处理到临时文件
                    success_result, error_msg, replace_count = replace_in_word(input_path, temp_output_path)
                    
                    # 如果处理成功，则进行后续操作
                    if success_result:
                        # 检查临时文件是否成功创建
                        if os.path.exists(temp_output_path) and os.path.getsize(temp_output_path) > 0:
                            # 安全地移动到最终位置
                            if safe_file_operation(temp_output_path, final_output_path, "move"):
                                success_files += 1
                                status = "✅ 成功"
                                total_replacements += replace_count
                            else:
                                # 移动失败，尝试复制源文件
                                if safe_file_operation(input_path, final_output_path, "copy"):
                                    logging.warning(f"Word处理失败，已直接复制原文件: {filename}")
                                    success_files += 1
                                    status = "⚠️ 未处理"
                                else:
                                    raise Exception("无法保存文件")
                        else:
                            logging.error(f"临时文件创建失败: {temp_output_path}")
                            # 尝试直接复制源文件
                            if safe_file_operation(input_path, final_output_path, "copy"):
                                logging.warning(f"Word处理失败，已直接复制原文件: {filename}")
                                success_files += 1
                                status = "⚠️ 未处理"
                            else:
                                raise Exception("无法保存文件")
                    else:
                        # 处理失败，记录错误信息
                        logging.error(f"Word文档处理失败: {error_msg}")
                        status = "❌ 失败"
                        replace_count = "--"
                        failed_files += 1
                except Exception as word_err:
                    logging.error(f"处理Word文件时出错: {str(word_err)}")
                    raise
            else:
                logging.warning(f"不支持的文件类型: {ext}")
                status = "⚠️ 跳过"
                replace_count = "--"
                skipped_files += 1
        except Exception as e:
            logging.error(f"处理失败: {str(e)}")
            failed_files += 1
            status = "❌ 失败"
            replace_count = "--"
            # 清理可能残留的临时文件
            if os.path.exists(temp_output_path):
                try:
                    os.remove(temp_output_path)
                    logging.info(f"已清理临时文件: {temp_output_path}")
                except:
                    pass

        # 计算处理时间
        process_time = time.time() - file_start_time
        
        # 清除进度行
        if SHOW_PROGRESS:
            print("\r" + " " * 80, end="\r")
        
        # 打印处理结果
        if replace_count == "--":
            print(f"{filename:<40} {replace_count:<10} {status:<10} {process_time:.2f}s")
        else:
            print(f"{filename:<40} {replace_count:<10} {status:<10} {process_time:.2f}s")

    # 清理临时目录
    try:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
            logging.info(f"已清理临时目录: {temp_dir}")
    except Exception as e:
        logging.warning(f"清理临时目录失败: {str(e)}")

    # 打印处理总结
    total_time = time.time() - start_time
    print("-" * 80)
    print(f"处理完成: 共 {total_files} 个文件 | 成功: {success_files} | 失败: {failed_files} | 跳过: {skipped_files} | 总耗时: {total_time:.2f}s")
    print(f"总计完成替换 {total_replacements} 处内容")
    print("=" * 80 + "\n")
    
    logging.info(f"批处理完成: 共 {total_files} 个文件, 成功: {success_files}, 失败: {failed_files}, 跳过: {skipped_files}, 总耗时: {total_time:.2f}s")
    logging.info(f"总计替换了 {total_replacements} 处内容")
    
    return success_files, total_replacements

def show_welcome():
    """ 显示欢迎信息 """
    print("\n" + "*" * 60)
    print("*" + " " * 58 + "*")
    print("*" + "    文档批量替换工具 v1.9    ".center(58) + "*")
    print("*" + " " * 58 + "*")
    print("*" + "  支持格式: .docx, .doc, .xlsx, .xls  ".center(58) + "*")
    print("*" + " " * 58 + "*")
    print("*" + "  特色功能: 可靠替换确认与多种替换策略  ".center(58) + "*")
    print("*" + " " * 58 + "*")
    print("*" * 60 + "\n")
    
    print(f"📁 源文件夹: {SOURCE_FOLDER}")
    print(f"📁 输出文件夹: {OUTPUT_FOLDER}")
    print(f"🔄 替换规则: {len(REPLACE_RULES)} 条")
    for old, new in REPLACE_RULES.items():
        print(f"   • \"{old}\" -> \"{new}\"")
    print("\n")

def check_environment():
    """ 检查运行环境 """
    logging.info("检查运行环境...")
    word = None
    excel = None
    try:
        # 检查Word
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # 禁用警告
        logging.info("检测到 Microsoft Word")
        
        # 检查Excel
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False  # 禁用警告
        logging.info("检测到 Microsoft Excel")
        
        return True
    except Exception as e:
        logging.error(f"环境检查失败: {str(e)}")
        print(f"错误: 无法初始化 Microsoft Office 组件。请确保已安装 Microsoft Office 且可正常运行。")
        print(f"详细错误: {str(e)}")
        return False
    finally:
        # 确保应用程序被关闭
        if word:
            try:
                word.Quit()
            except:
                pass
        if excel:
            try:
                excel.Quit()
            except:
                pass

if __name__ == "__main__":
    try:
        logging.info("程序启动")
        show_welcome()
        
        # 检查环境
        if not check_environment():
            sys.exit(1)
            
        # 开始处理
        batch_process()
    except KeyboardInterrupt:
        print("\n程序被用户中断")
        logging.warning("程序被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"\n程序发生错误: {str(e)}")
        logging.critical(f"程序发生错误: {str(e)}")
        sys.exit(1)
