import os
from docx import Document
from langchain.document_loaders import PyPDFLoader
from typing import List

# 处理PDF文件
def process_pdf(file_path: str) -> List[str]:
    loader = PyPDFLoader(file_path)
    pages = loader.load_and_split()
    return [page.page_content for page in pages]


# 处理DOCX文件
def process_docx(file_path: str) -> str:
    doc = Document(file_path)
    return '\n'.join([paragraph.text for paragraph in doc.paragraphs])


# 处理DOC文件（需要安装pywin32库）
import win32com.client as win32

def process_doc(file_path: str) -> str:
    # 启动Word应用程序
    word = win32.gencache.EnsureDispatch('Word.Application')
    # 打开.doc文件
    doc = word.Documents.Open(file_path)
    # 读取文件内容
    content = doc.Range().Text
    # 关闭Word文档
    doc.Close()
    # 退出Word应用程序
    word.Quit()
    return content



# 处理TXT文件
def process_txt(file_path: str) -> str:
    with open(file_path, 'r', encoding='utf-8') as file:
        return file.read()


# 保存解析文本到新的txt文件
def save_to_txt(content: str, save_path: str):
    with open(save_path, 'w', encoding='utf-8') as file:
        file.write(content)


# 根据文件后缀名选择处理函数
def process_file(file_path: str, save_dir: str) -> str:
    _, file_extension = os.path.splitext(file_path)
    save_path = os.path.join(save_dir, os.path.basename(file_path) + '.txt')

    if file_extension.lower() == '.pdf':
        content = '\n'.join(process_pdf(file_path))
    elif file_extension.lower() == '.docx':
        content = process_docx(file_path)
    elif file_extension.lower() == '.doc':
        content = process_doc(file_path)
    elif file_extension.lower() == '.txt':
        content = process_txt(file_path)
    else:
        raise ValueError(f"不支持的文件类型: {file_extension}")

    save_to_txt(content, save_path)
    return save_path


# 主函数
def main(file_paths: List[str], save_dir: str):
    for file_path in file_paths:
        try:
            processed_file_path = process_file(file_path, save_dir)
            print(f"文件 {file_path} 的内容已处理并保存到 {processed_file_path}")
        except Exception as e:
            print(f"处理文件 {file_path} 时发生错误: {e}")


# 测试代码
if __name__ == "__main__":
    files_to_process = [
        input('请输入文件路径：')
        # 'example.docx',
        # # 'example.doc',  # 这将引发未实现异常
        # 'example.txt'
    ]
    save_directory = 'processed_texts'  # 指定保存解析文本的目录
    if not os.path.exists(save_directory):
        os.makedirs(save_directory)  # 如果目录不存在，则创建它
    main(files_to_process, save_directory)


# # 用户输入
# if __name__ == "__main__":
#     input_path = input("请输入文件或文件夹路径: ")
#     save_directory = 'processed_texts'  # 指定保存解析文本的目录
#     if not os.path.exists(save_directory):
#         os.makedirs(save_directory)  # 如果目录不存在，则创建它
#     main(input_path, save_directory)
