'''
数会程序的GUI界面，使用tkinter库。
'''


import tkinter as tk
# import ctypes
import tkinter.filedialog as fd
import tkinter.messagebox as messagebox
from count_meet_func.run_code import try_to_count_meet
import count_meet_func.text_config as text_config


def browse_folder():
    """选择底稿所在的文件夹，同时尝试数会计算。
    """
    folder_path = fd.askdirectory()
    if folder_path:
        new_value = new_entry.get()
        
        result = try_to_count_meet(folder_path, new_value)
        show_results(result)


def show_results(result):
    new_value = new_entry.get()
    if len(new_value) != 4 or not new_value.isdigit():
        text = text_config.date_error  # 日期格式错误
    else:
        text = result

    result_text.config(state='normal')  # 允许编辑
    result_text.delete(1.0, tk.END)  # 清空文本框内容
    result_text.insert(tk.END, text)  # 插入结果文本
    result_text.config(state='disabled')  # 禁止编辑
            

# 创建窗口
window = tk.Tk()
window.title("数会计算程序")
window.geometry("600x400")

# 调用api设置成由应用程序缩放
# ctypes.windll.shcore.SetProcessDpiAwareness(1)
# 调用api获得当前的缩放因子
# ScaleFactor = ctypes.windll.shcore.GetScaleFactorForDevice(0)
# 设置缩放因子
# window.tk.call('tk', 'scaling', ScaleFactor / 75)

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
warning_label = tk.Label(window, text=text_config.notice, fg="red", anchor="w", justify="left")
warning_label.pack()

# 创建文本框用于显示结果
result_text = tk.Text(window, state='disabled')
result_text.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)

# 运行窗口主循环
window.mainloop()
