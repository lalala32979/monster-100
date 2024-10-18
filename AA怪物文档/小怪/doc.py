from docx import Document

# 创建一个新的 Word 文档
document = Document()

# 向文档中添加一段话
document.add_paragraph('啦啦啦啦你好')

# 保存文档
document.save('world.docx')

import pandas as pd

# 读取表格数据
df = pd.read_excel('config_moster_data.xlsx')

# 获取第一列数据
first_column = df.iloc[:, 0]

# 打印第一列数据
print(first_column)