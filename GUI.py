import tkinter as tk
import tkinter.filedialog as fd
import pandas as pd
import os
import sys


def run_code(folder_path, new_value):
    df_total = pd.DataFrame()

    file_names = os.listdir(folder_path)

    for file_name in file_names:
        file_path = os.path.join(folder_path, file_name)

        if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
            df = pd.read_excel(file_path)
            df.dropna(how='all', inplace=True)

            if file_name == new_value + '.xlsx':
                df['if_new'] = 1

            if file_name != 'toplist.xlsx':
                df_total = pd.concat([df_total, df], ignore_index=True)
            else:
                toplist = df
                toplist.dropna(how='all', inplace=True)

    df_total.to_excel(folder_path + '当期总表.xlsx', index=False)  # 输出当期总表
    
    df_this_week = df_total[df_total['if_new'] == 1]
    week_in = df_this_week['机构'].nunique()
    week_people = df_this_week['人名'].nunique()

    df_this_week_1on1 = df_this_week[df_this_week['group'].isna()]
    one_on_one_in = df_this_week_1on1['机构'].nunique()
    one_on_one_people = df_this_week_1on1['人名'].nunique()

    brokers_list = df_this_week['group'].unique().tolist()
    count_brokers = len(brokers_list)

    total_in = df_total['机构'].nunique()
    total_people = df_total['人名'].nunique()

    top10_list = toplist[toplist['rate'] <= 10]['Institution'].unique().tolist()
    top30_list = toplist[toplist['rate'] <= 30]['Institution'].unique().tolist()

    df_top10 = df_total[df_total['机构'].isin(top10_list)]
    count_top10 = df_top10['机构'].nunique()
    count_top10_list = df_top10['机构'].tolist()

    df_top30 = df_total[df_total['机构'].isin(top30_list)]
    count_top30 = df_top30['机构'].nunique()

    df_total = df_total.merge(toplist, left_on='机构', right_on='Institution', how='left')
    df_potential = df_total[(df_total['ADS'] <= 100000)]
    count_potential = df_potential['机构'].nunique()
    potential_list = df_potential['机构'].unique().tolist()[0:5]

    brokers_list_str = ', '.join('%s'%brokers for brokers in brokers_list)
    count_top10_list_str = ', '.join('%s'%top10 for top10 in count_top10_list)
    
    result1 = f'* 自6月21日至{new_value}，本周合计与{week_in}个机构{week_people}人沟通，包括：1*1合计{one_on_one_in}个机构（覆盖{one_on_one_people}人），{count_brokers}家brokers（包括：{brokers_list_str}）举办的NDR或行业会议。'
    result2 = f'* 自财报第二天（5月17日），我们已沟通{total_in}家机构{total_people}人，包括Top10中的{count_top10}个({count_top10_list_str}); 以及另外{count_top30}个Top 30大股东。潜在买家{count_potential}个，包括：{potential_list}。潜在买家定义：持股量少于10万ADR。'

    text = result1 + '\n' + result2

    # 输出结果到GUI
    result_text.config(state='normal')  # 允许编辑
    result_text.delete(1.0, tk.END)  # 清空文本框内容
    result_text.insert(tk.END, text)  # 插入结果文本
    result_text.config(state='disabled')  # 禁止编辑

    # print(text)



def browse_folder():
    folder_path = fd.askdirectory()
    if folder_path:
        new_value = new_entry.get()
        run_code(folder_path, new_value)


# 创建窗口
window = tk.Tk()
window.title("计算程序")
window.geometry("400x300")

# 创建输入文本框标签
new_label = tk.Label(window, text="输入最新日期:")
new_label.pack(pady=10)

# 创建输入文本框
new_entry = tk.Entry(window)
new_entry.pack(pady=5)

# 创建文件夹选择按钮
button = tk.Button(window, text="选择底稿文件夹", command=browse_folder)
button.pack(pady=10)

# 创建提示文本
warning_label = tk.Label(window, text="注意底稿文件命名规范", fg="red")
warning_label.pack()

# 创建文本框用于显示结果
result_text = tk.Text(window, state='disabled')
result_text.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)

# 运行窗口主循环
window.mainloop()
