import os
import win32com.client as win32

def fix_word_extension(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    # 只处理doc和docx后缀
    if ext not in ['.doc', '.docx']:
        print(f"{file_path} 不是doc/docx文件，无需修正。")
        return
    # 检查伪docx
    if ext == '.doc':
        try:
            from docx import Document
            Document(file_path)
            # 能用python-docx打开，说明其实是docx
            new_path = file_path + 'x'
            os.rename(file_path, new_path)
            print(f"已将 {file_path} 重命名为 {new_path}（实际为docx）")
            return
        except Exception:
            pass  # 不是docx，继续后面流程
    # 检查伪doc
    if ext == '.docx':
        try:
            from docx import Document
            Document(file_path)
            print(f"{file_path} 是真正的docx文件，无需修正。")
            return
        except Exception:
            print(f"{file_path} 不是标准docx，尝试用win32com打开...")
        try:
            word = win32.gencache.EnsureDispatch('Word.Application')
            word.Visible = 0
            doc = word.Documents.Open(file_path)
            doc.Close()
            word.Quit()
            print(f"{file_path} 虽然不能用python-docx打开，但Word可以打开，建议人工确认。")
        except Exception:
            # 如果win32com也打不开，尝试重命名为.doc
            new_path = file_path[:-1]  # 去掉x，变成.doc
            try:
                os.rename(file_path, new_path)
                print(f"已将 {file_path} 重命名为 {new_path}（实际为doc）")
            except Exception as e:
                print(f"重命名失败: {e}")

if __name__ == "__main__":
    folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'input_file')
    for fname in os.listdir(folder):
        if fname.lower().endswith('.doc') or fname.lower().endswith('.docx'):
            fix_word_extension(os.path.join(folder, fname))