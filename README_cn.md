# 工具描述
可对folder_path下的docx和doc文件进行文件名关键字替换，简易实现脱敏的功能。

# 注意
- 1.doc格式处理需要先用liboffice转为docx，所以需要本地有liboffice，cmd中执行soffice是ok的
- 2.会直接删除源文件，务必**额外保存源文件**
- 
# 使用方法
修改代码中
folder_path -> 源及目标目标文件夹
keywords_to_replace -> 敏感词列表 
replacement_text -> 脱敏词
