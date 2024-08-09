from PyInstaller.utils.hooks import collect_data_files

# 收集 docxcompose 模块中的所有数据文件
datas = collect_data_files('docxcompose', includes=['templates/*'])
