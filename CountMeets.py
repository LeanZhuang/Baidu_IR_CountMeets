import pandas as pd
import numpy as np
import os


# 定义全局变量
folder_path = './底稿'
new = '0630'


df_total = pd.DataFrame()

# 获取文件夹中的所有文件名
file_names = os.listdir(folder_path)

# 遍历文件夹中的每个文件
for file_name in file_names:
    # 检查文件名是否以'.xlsx'或'.xls'结尾
    if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
        # 构建文件路径
        file_path = os.path.join(folder_path, file_name)
        
        # 读取Excel文件为DataFrame
        df = pd.read_excel(file_path)
        df.dropna(how='all', inplace=True)
        
        if file_name == new + '.xlsx':
            df['if_new'] = 1
        
        if file_name != 'toplist.xlsx':
            df_total = df_total.append(df, ignore_index=True)
        else:
            toplist = df
            toplist.dropna(how='all', inplace=True)


# 本周
df_this_week = df_total[df_total['if_new'] == 1]  # 本周数据

week_in = df_this_week['机构'].nunique()  # 本周机构数
week_people = df_this_week['人名'].nunique()  # 本周人数

# 本周1*1
df_this_week_1on1 = df_this_week[df_this_week['group'].isna()]
one_on_one_in = df_this_week_1on1['机构'].nunique()  # 本周1*1机构数
one_on_one_people = df_this_week_1on1['人名'].nunique()  # 本周1*1人数

# brokers
brokers_list = df_this_week['group'].unique().tolist()
count_brokers = len(brokers_list)


# 自财报第二天开始
total_in = df_total['机构'].nunique()  # 总机构数
total_people = df_total['人名'].nunique()  # 总人数

top10_list = toplist[toplist['rate'] <= 10]['Institution'].unique().tolist()
top30_list = toplist[toplist['rate'] <= 30]['Institution'].unique().tolist()

# Top10
df_top10 = df_total[df_total['机构'].isin(top10_list)]
count_top10 = df_top10['机构'].nunique()
count_top10_list = df_top10['机构'].tolist()

# Top30
df_top30 = df_total[df_total['机构'].isin(top30_list)]
count_top30 = df_top30['机构'].nunique()


# 潜在买家
df_total = df_total.merge(toplist, left_on='机构', right_on='Institution', how='left')
df_potential = df_total[(df_total['ADS'] <= 100000)]
count_potential = df_potential['机构'].nunique()
potential_list = df_potential['机构'].unique().tolist()[0:5]


result1 = f'* 自6月21日至{new}，本周合计与{week_in}个机构{week_people}人沟通，包括：1*1合计{one_on_one_in}个机构（覆盖{one_on_one_people}人），{count_brokers}家brokers（包括：{brokers_list}）举办的NDR或行业会议。'


result2 = f'* 自财报第二天（5月17日），我们已沟通{total_in}家机构{total_people}人，包括Top10中的{count_top10}个({count_top10_list}); 以及另外{count_top30}个Top 30大股东。潜在买家{count_potential}个，包括：{potential_list}。潜在买家定义：持股量少于10万ADR。'


def write_to_file(text, filename):
    with open(filename, 'w') as file:
        file.write(text)

text = result1 + '\n' + result2
filename = "result.txt"
write_to_file(text, filename)

