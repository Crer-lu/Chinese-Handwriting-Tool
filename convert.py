import os
import random

import win32com.client as win32

# 修改到当前目录
work_dir = os.path.dirname(os.path.abspath(__file__))
# print(work_dir)

# 创建Word应用程序实例
word = win32.gencache.EnsureDispatch('Word.Application')

# 打开源文档
source_doc = word.Documents.Open(work_dir + '/source.docx')

# 选择要复制的文本
source_doc.Content.WholeStory()
source_doc.Content.Copy()

# 打开目标文档
target_doc = word.Documents.Open(work_dir + '/template.docx')

# 将复制的文本粘贴到目标文档（仅文字）
target_doc.Content.PasteSpecial(DataType=win32.constants.wdPasteText)

# 关闭源文档
source_doc.Close()

# 设置要应用的新字体
font_name = '对你不止是喜欢'
font_size = [14, 15, 16, 17, 18]

# 遍历文档中的所有段落，设置字体
for paragraph in target_doc.Paragraphs:
    for run in paragraph.Range.Words:
        run.Font.Name = font_name
        run.Font.Size = font_size[random.randint(0, 4)]

# 保存目标文档
target_doc.Save()


# 关闭目标文档
target_doc.Close()

# 关闭Word应用程序
word.Quit()
