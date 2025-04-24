import os
import shutil
import logging
import time
import configparser
from functools import wraps
from head_geter import WordHeadGetter, find_section_offsets

# 设置日志格式
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[logging.StreamHandler()]
)

def timeit_log(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        start_time = time.time()
        logging.info(f"开始执行: {func.__name__}")
        result = func(*args, **kwargs)
        end_time = time.time()
        logging.info(f"结束执行: {func.__name__}，耗时: {end_time - start_time:.2f} 秒")
        return result
    return wrapper

def read_config(config_file='config.ini'):
    """从配置文件读取参数"""
    config = configparser.ConfigParser()
    
    # 检查配置文件是否存在
    if not os.path.exists(config_file):
        logging.error(f"配置文件 {config_file} 不存在")
        return None
    
    # 读取配置文件
    config.read(config_file, encoding='utf-8')
    
    # 读取路径配置
    paths = {}
    if 'Paths' in config:
        for key in config['Paths']:
            paths[key] = config['Paths'][key]
    
    # 读取章节设置
    chapter_settings = {}
    if 'ChapterSettings' in config:
        for key in config['ChapterSettings']:
            # 处理特殊情况：多个关键词用逗号分隔
            if key in ['section1', 'section2']:
                chapter_settings[key] = config['ChapterSettings'][key].split(',')
            else:
                # 尝试将其他参数转换为整数
                try:
                    chapter_settings[key] = int(config['ChapterSettings'][key])
                except ValueError:
                    chapter_settings[key] = config['ChapterSettings'][key]
    
    # 读取处理设置
    processing = {}
    if 'Processing' in config:
        for key in config['Processing']:
            # 尝试将数值参数转换为适当的类型
            if key == 'wait_time':
                processing[key] = float(config['Processing'][key])
            elif key == 'verbose':
                processing[key] = config['Processing'][key].lower() == 'true'
            else:
                processing[key] = config['Processing'][key]
    
    return {
        'paths': paths,
        'chapter_settings': chapter_settings,
        'processing': processing
    }

@timeit_log
def slice_word_by_delete_with_getter(getter, input_path, output_path, section1_keywords, section2_keywords, 
                                     section1_offset=1, section2_offset=1, section1_level=1, section2_level=1):
    """根据关键词列表切片文档"""
    if os.path.abspath(input_path) != os.path.abspath(output_path):
        shutil.copy2(input_path, output_path)
    
    # 确保输出文件不是只读的
    try:
        # 修改文件权限，确保可写
        os.chmod(output_path, 0o666)
    except Exception as e:
        logging.error(f"修改文件权限失败: {e}")
    
    # 获取文档标题结构
    titles, doc = getter.get_document_titles_tree(output_path)
    if titles is None:
        logging.warning(f"无法解析文档结构: {output_path}")
        return False

    # 获取指定等级的标题
    level1_titles = [node for node in titles if node.get("级别", 1) == section1_level]
    level2_titles = [node for node in titles if node.get("级别", 1) == section2_level]
    offsets1 = [(node["标题"], node["偏移量"]) for node in level1_titles]
    offsets2 = [(node["标题"], node["偏移量"]) for node in level2_titles]
    
    # 找到section1和section2的index
    start_idx = end_idx = None
    
    # 遍历所有section1关键词，找到匹配的第一个
    for keyword in section1_keywords:
        if start_idx is not None:
            break
        for idx, (title, _) in enumerate(offsets1):
            if keyword in title:
                start_idx = idx
                logging.info(f"找到起始章节: {keyword} 在 {title}")
                break
    
    # 遍历所有section2关键词，找到匹配的第一个
    for keyword in section2_keywords:
        if end_idx is not None:
            break
        for idx, (title, _) in enumerate(offsets2):
            if keyword in title:
                end_idx = idx
                logging.info(f"找到结束章节: {keyword} 在 {title}")
                break
    
    # 如果未找到指定章节，尝试关闭文档并返回False
    if start_idx is None or end_idx is None:
        logging.warning(f"未找到指定章节: {output_path}")
        try:
            if doc is not None:
                doc.Close(False)
        except Exception as e:
            logging.error(f"关闭文档时发生异常: {e}")
        return False
    
    # 计算实际切片区间
    start_target_idx = start_idx + section1_offset
    end_target_idx = end_idx + section2_offset
    if start_target_idx >= len(offsets1) or end_target_idx > len(offsets2):
        logging.warning(f"切片区间超出标题范围: {output_path}")
        try:
            if doc is not None:
                doc.Close(False)
        except Exception as e:
            logging.error(f"关闭文档时发生异常: {e}")
        return False
    
    start = offsets1[start_target_idx][1]
    if end_target_idx == len(offsets2):
        end = doc.Content.End
    else:
        end = offsets2[end_target_idx][1]

    try:
        # 验证切片点范围
        content_end = doc.Content.End
        if end > content_end:
            end = content_end
        if start < 0:
            start = 0

        # 执行切片操作
        logging.info(f"开始切片: {output_path}，范围: {start} - {end}")
        doc.Range(end, content_end).Delete()
        logging.info(f"删除结束部分: {output_path}")
        doc.Range(0, start).Delete()
        logging.info(f"删除开始部分: {output_path}")
        
        # 保存文件到临时文件，避免只读问题
        temp_save_path = output_path + ".tmp"
        try:
            doc.SaveAs(temp_save_path)
            logging.info(f"临时文件保存成功: {temp_save_path}")
            
            # 关闭文档
            doc.Close(False)  # 不保存原文档
            doc = None  # 确保释放引用
            
            # 如果原文件存在，删除它
            if os.path.exists(output_path):
                os.remove(output_path)
                
            # 重命名临时文件为目标文件
            os.rename(temp_save_path, output_path)
            logging.info(f"保存完成: {output_path}")
            logging.info(f"切片完成: {output_path}")
            return True
        except Exception as e:
            logging.error(f"保存文件失败: {e}")
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception as close_err:
                    logging.error(f"关闭文档时发生异常: {close_err}")
            return False
            
    except Exception as e:
        logging.error(f"处理{output_path}时发生异常: {e}")
        # 尝试关闭文档
        if doc is not None:
            try:
                doc.Close(False)  # 不保存
            except Exception as close_err:
                logging.error(f"关闭文档时发生异常: {close_err}")
        return False

@timeit_log
def process_file(getter, file_path, temp_folder, output_folder, unsupport_folder, old_folder, 
                 section1_keywords, section2_keywords, section1_offset, section2_offset, 
                 section1_level, section2_level, wait_time=1):
    """处理单个文件"""
    ext = os.path.splitext(file_path)[1].lower()
    file_name = os.path.basename(file_path)

    if ext not in ['.doc', '.docx']:
        target_path = os.path.join(unsupport_folder, file_name)
        logging.info(f"{file_path} 不是doc/docx，已剪切到: {target_path}")
        try:
            shutil.copy2(file_path, target_path)  # 复制到不支持文件夹
            os.remove(file_path)  # 删除源文件
        except Exception as e:
            logging.error(f"文件操作失败: {e}")
        return "unsupported"

    base_name = os.path.splitext(file_name)[0]
    temp_path = os.path.join(temp_folder, file_name)
    out_path = os.path.join(output_folder, base_name + '_slice' + ext)

    try:
        # 复制到临时文件夹
        shutil.copy2(file_path, temp_path)
        
        # 处理文件
        logging.info(f"开始处理文件: {file_name}")
        success = slice_word_by_delete_with_getter(
            getter, temp_path, temp_path, section1_keywords, section2_keywords, 
            section1_offset, section2_offset, section1_level, section2_level
        )

        # 确保Word进程不再占用该文件
        time.sleep(wait_time)
        
        # 无论是否成功，都将原文件移动到old_file文件夹
        try:
            # 使用复制+删除替代移动
            old_file_path = os.path.join(old_folder, file_name)
            shutil.copy2(file_path, old_file_path)
            os.remove(file_path)
            logging.info(f"原文件已移动到: {old_file_path}")
        except Exception as e:
            logging.error(f"移动原文件到old_file文件夹失败: {e}")
        
        # 根据处理结果分别处理临时文件
        if success:
            # 等待一下确保文件不被占用
            time.sleep(wait_time / 2)
            try:
                # 使用复制后删除的方式移动到output文件夹
                shutil.copy2(temp_path, out_path)
                os.remove(temp_path)
                logging.info(f"切片成功，结果已保存到: {out_path}")
                return "success"
            except Exception as e:
                logging.error(f"移动文件到output文件夹失败: {e}")
                return "failed"
        else:
            try:
                # 等待一下确保文件不被占用
                time.sleep(wait_time / 2)
                
                # 使用复制后删除的方式移动到unsupport文件夹
                unsupport_path = os.path.join(unsupport_folder, file_name)
                shutil.copy2(temp_path, unsupport_path)
                os.remove(temp_path)
                logging.info(f"切片失败，原文件已移动到: {unsupport_path}")
            except Exception as e:
                logging.error(f"移动文件到unsupport文件夹失败: {e}")
            return "failed"
    except Exception as e:
        logging.error(f"处理文件时发生异常: {e}")
        # 清理临时文件
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception as rm_err:
                logging.error(f"删除临时文件失败: {rm_err}")
                
        # 仍然尝试将原文件移动到old_file文件夹
        try:
            old_file_path = os.path.join(old_folder, file_name)
            shutil.copy2(file_path, old_file_path)
            os.remove(file_path)
            logging.info(f"原文件已移动到: {old_file_path}")
        except Exception as e:
            logging.error(f"移动原文件到old_file文件夹失败: {e}")
            
        return "error"

@timeit_log
def process_folder_by_delete(config):
    """根据配置处理文件夹"""
    # 从配置中获取路径
    base_dir = os.path.dirname(os.path.abspath(__file__))
    input_folder = os.path.join(base_dir, config['paths'].get('input_folder', 'input_file'))
    output_folder = os.path.join(base_dir, config['paths'].get('output_folder', 'output_file'))
    unsupport_folder = os.path.join(base_dir, config['paths'].get('unsupport_folder', 'unsupport_file'))
    old_folder = os.path.join(base_dir, config['paths'].get('old_folder', 'old_file'))
    temp_folder = os.path.join(base_dir, config['paths'].get('temp_folder', 'temp_file'))
    
    # 创建必要的文件夹
    for folder in [output_folder, unsupport_folder, old_folder, temp_folder]:
        if not os.path.exists(folder):
            os.makedirs(folder)
    
    # 从配置中获取章节设置
    section1_keywords = config['chapter_settings'].get('section1', ['总论'])
    section2_keywords = config['chapter_settings'].get('section2', ['建设方案'])
    section1_offset = config['chapter_settings'].get('section1_offset', 1)
    section2_offset = config['chapter_settings'].get('section2_offset', 1)
    section1_level = config['chapter_settings'].get('section1_level', 1)
    section2_level = config['chapter_settings'].get('section2_level', 1)
    
    # 从配置中获取处理设置
    wait_time = config['processing'].get('wait_time', 1)
    verbose = config['processing'].get('verbose', True)
    
    # 设置日志级别
    if verbose:
        logging.getLogger().setLevel(logging.INFO)
    else:
        logging.getLogger().setLevel(logging.WARNING)
    
    # 获取所有文件
    files = os.listdir(input_folder)
    files = [f for f in files if os.path.isfile(os.path.join(input_folder, f))]
    
    if not files:
        logging.info("没有找到需要处理的文件")
        return

    logging.info(f"发现 {len(files)} 个文件待处理")
    logging.info(f"起始章节关键词: {section1_keywords}, 结束章节关键词: {section2_keywords}")
    logging.info(f"章节偏移量: {section1_offset}, {section2_offset}")
    logging.info(f"章节级别: {section1_level}, {section2_level}")
    
    # 创建Word处理器
    getter = WordHeadGetter()
    
    # 处理每个文件
    success_count = failed_count = error_count = unsupported_count = 0
    for file_name in files:
        file_path = os.path.join(input_folder, file_name)
        result = process_file(
            getter, file_path, temp_folder, output_folder, unsupport_folder, old_folder,
            section1_keywords, section2_keywords, section1_offset, section2_offset,
            section1_level, section2_level, wait_time
        )
        
        if result == "success":
            success_count += 1
        elif result == "failed":
            failed_count += 1
        elif result == "error":
            error_count += 1
        elif result == "unsupported":
            unsupported_count += 1
    
    # 所有文件处理完成后，关闭所有标签页
    if getter.word is not None:
        try:
            # 关闭所有打开的标签页
            while getter.word.Windows.Count > 0:
                getter.word.Windows(1).Close()
                logging.info("已关闭一个Word标签页")
            logging.info("所有Word标签页已关闭")
        except Exception as e:
            logging.error(f"关闭标签页失败: {e}")
    
    # 关闭Word处理器
    getter.quit()
    logging.info("Word应用程序已退出")
    
    logging.info(f"处理完成。成功: {success_count}, 失败: {failed_count}, 错误: {error_count}, 不支持: {unsupported_count}")

if __name__ == "__main__":
    # 从配置文件读取参数
    config_file = 'config.ini'
    config = read_config(config_file)
    
    if config is None:
        logging.error(f"无法加载配置文件 {config_file}，程序退出")
        exit(1)
    
    # 处理文件夹
    process_folder_by_delete(config)
