import win32com.client as win32
import logging
import time
import pythoncom
from functools import wraps
from tqdm import tqdm  # 添加进度条库

# 彩色日志装饰器
COLORS = {
    'HEADER': '\033[95m',
    'OKBLUE': '\033[94m',
    'OKCYAN': '\033[96m',
    'OKGREEN': '\033[92m',
    'WARNING': '\033[93m',
    'FAIL': '\033[91m',
    'ENDC': '\033[0m',
    'BOLD': '\033[1m',
    'UNDERLINE': '\033[4m'
}

def color_log(msg, color='OKGREEN'):
    print(f"{COLORS.get(color, COLORS['OKGREEN'])}{msg}{COLORS['ENDC']}")

def timeit_log(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        start_time = time.time()
        color_log(f"开始执行: {func.__name__}", 'OKBLUE')
        result = func(*args, **kwargs)
        end_time = time.time()
        color_log(f"结束执行: {func.__name__}，耗时: {end_time - start_time:.2f} 秒", 'OKCYAN')
        return result
    return wrapper

class WordHeadGetter:

    @timeit_log
    def __init__(self):
        self.word = win32.gencache.EnsureDispatch('Word.Application')
        self.word.Visible = 0

    @timeit_log
    def get_document_titles_tree(self, file_path):
        doc = None
        try:

            # 设置打开文档的选项，提高速度
            doc = self.word.Documents.Open(
                file_path, 
                ReadOnly=True,  # 只读模式
                Visible=False,  # 不可见
                AddToRecentFiles=False  # 不添加到最近文件
            )
            titles = []
            para_count = doc.Paragraphs.Count

            # 一次性提取所有段落的内容和属性，减少COM调用
            paragraphs = [(para.Range.Text.strip(), para.Style.NameLocal, para.Range.Start, para.OutlineLevel)
                          for para in doc.Paragraphs]

            # 遍历提取的段落信息，收集标题
            with tqdm(total=para_count, desc="处理段落", unit="段") as pbar:
                for para_text, style_name, start_offset, outline_level in paragraphs:
                    if style_name.startswith("标题") or style_name.startswith("Heading"):
                        titles.append({
                            "标题": para_text,
                            "偏移量": start_offset,
                            "级别": outline_level,
                            "children": []
                        })
                    pbar.update(1)

            # 根据标题的等级和偏移值构建树状结构
            root = []
            for title in titles:
                if title["级别"] == 1:
                    root.append(title)
                else:
                    # 找到最近的上级标题
                    parent = root[-1]  # 从最后一个一级标题开始
                    while parent["children"] and parent["children"][-1]["级别"] < title["级别"]:
                        parent = parent["children"][-1]
                    parent["children"].append(title)
            return root, doc
        except Exception as e:
            color_log(f"发生错误: {e}", 'FAIL')
            if doc is not None:
                try:
                    doc.Close(False)
                except:
                    pass
            return None, None

    @timeit_log
    def quit(self):
        if self.word is not None:
            try:
                self.word.Quit()
            except:
                pass
            self.word = None

@timeit_log
def find_section_offsets(tree, section1, section2):
    """
    只在1级标题中查找section1和section2的下一个1级标题的偏移量。
    返回(start, end)
    """
    start = end = None
    found_section1 = found_section2 = False
    offsets = []
    # 只收集所有1级标题的偏移量和标题
    for node in tree:
        if node.get("级别", 1) == 1:
            offsets.append((node["标题"], node["偏移量"]))
    # 遍历，找到section1和section2的下一个1级标题
    for idx, (title, offset) in enumerate(offsets):
        if not found_section1 and section1 in title:
            if idx + 1 < len(offsets):
                start = offsets[idx + 1][1]
            found_section1 = True
        if not found_section2 and section2 in title:
            if idx + 1 < len(offsets):
                end = offsets[idx + 1][1]
            found_section2 = True
    return start, end

if __name__ == "__main__":
    # 示例用法
    getter = WordHeadGetter()
    input_path = "D:\WORK\Word-Geter\\2.doc"
    section1 = "研究背景"
    section2 = "实验设计"

    titles, doc = getter.get_document_titles_tree(input_path)
    if titles is None:
        color_log(f"无法解析文档结构: {input_path}", 'FAIL')
    else:
        start, end = find_section_offsets(titles, section1, section2)
        color_log(f"找到的偏移量: {start}, {end}", 'OKGREEN')
        if doc is not None:
            try:
                if doc.Windows.Count > 0:  # 检查是否有打开的标签页
                    doc.Windows(1).Close()  # 只关闭标签页
                else:
                    color_log("没有打开的标签页可关闭。", 'WARNING')
            except Exception as e:
                color_log(f"关闭标签页时发生异常: {e}", 'FAIL')
    getter.quit()