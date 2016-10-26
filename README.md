# filter_invention
这个工具的作用是过滤部门的专利总览表，然后统计出g各处个人的专利完成情况。分为四列：发明提交、发明受理、实用新型提交、实用新型受理。输入文档必须是csv格式的，且只有一个sheet。本工具基于Python Tkinter制作图形界面。依赖详见import部分。打包成exe格式请使用pyinstall,命令为Python pyinstaller.py -F InvertionFilter.py
