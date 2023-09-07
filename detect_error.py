# 模块功能：检测底稿是否有错误

import os
import pandas as pd


def detect_column_name(folder_path):
    file_names = os.listdir(folder_path)

    for file_name in file_names:
        file_path = os.path.join(folder_path, file_name)

        if (file_name.endswith('.xlsx') or file_name.endswith('.xls')) and (file_name != 'toplist.xlsx'):
            
            df = pd.read_excel(file_path)
            
            # 如果 df 的列名不是 abcdefg，那么就返回错误
            df_columns = df.columns.tolist()
            
            if df_columns[0:4] != ['日期', '类型', 'group', '机构', '人名']:
                result1, result2 = "## Excel 底稿列名错误 ##\n\n请检查前五列是列名是否为：'日期', '类型', 'group', '机构', '人名'", ''
            else:
                result1, result2 = True, True
    
    return result1, result2
