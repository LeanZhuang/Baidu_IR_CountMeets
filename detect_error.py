'''
辅助模块
模块功能：检测底稿是否有错误
'''

import os
import pandas as pd


def detect_column_name(folder_path):
    file_names = os.listdir(folder_path)
    
    files_to_remove = ['【底稿】当期总表.xlsx', 'toplist.xlsx']
    file_names = [file for file in file_names if file not in files_to_remove]

    for file_name in file_names:
        file_path = os.path.join(folder_path, file_name)

        if (file_name.endswith('.xlsx') or file_name.endswith('.xls')):
            
            df = pd.read_excel(file_path)
            columns = df.columns[:5].tolist()
            
            is_matching = columns == ['日期', '类型', 'group', '机构', '人名']
    
    return is_matching