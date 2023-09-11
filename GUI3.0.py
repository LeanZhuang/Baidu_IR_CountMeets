'''
数会程序的GUI界面，使用tkinter库。
'''

# import ctypes
import time
import tkinter as tk
from tkinter import font
import tkinter.filedialog as fd
import tkinter.messagebox as messagebox

from count_meet_func.run_code import try_to_count_meet
import count_meet_func.text_config as text_config

# import tkinter as tk
from tkinter import ttk
# from ttkthemes import themed_tk as tk

# 创建ttk样式
# style = ttk.Style()
# style.configure('CustomFrame.TFrame', background='light grey', relief='solid', bordercolor='red')

# 创建Frame小部件
# frame = ttk.Frame(window, style='CustomFrame.TFrame', width=200, height=200)
# frame.pack()
# tk.ThemedStyle(theme="arc").theme_use()


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
window.geometry("800x600")
# window.set_theme("clam")


# result_text = tk.Text(window, state='disabled')
custom_font = font.Font(family="Arial", size=14)
custom_font_bold = font.Font(family="Arial", size=14, weight=font.BOLD)

# result_text.configure(font=custom_font)


# 创建容器框架
frame1 = tk.Frame(window, relief="solid", borderwidth=0, highlightthickness=0)
frame1.pack(padx=25, pady=10, expand=True)
frame1.pack_configure(anchor=tk.CENTER)


# 创建输入文本框标签
new_label = tk.Label(frame1, text="输入最新日期:")
new_label.configure(font=custom_font)
new_label.pack(pady=5, side=tk.LEFT)

# 创建输入文本框
new_entry = tk.Entry(frame1)
new_entry.pack(pady=5, side=tk.RIGHT)

# 创建文件夹选择按钮
button = ttk.Button(window, text="选择底稿文件夹并输出结果", command=browse_folder)
# button.configure(font=custom_font)
button.pack(padx=25, pady=10)

# 创建容器框架
frame2 = ttk.Frame(window, style='CustomFrame.TFrame')
frame2.pack(padx=25, pady=10, fill=tk.BOTH, expand=True)

label = ttk.Label(frame2, text="- 注意事项 -")
label.configure(font=custom_font_bold)
label.place(x=2, y=0)

# 创建文本标签
warning_label = tk.Label(frame2, text=text_config.notice, fg="red", anchor="w", justify="left")
warning_label.pack(padx=10, pady=0,)


# 创建容器框架
frame3 = tk.Frame(window, relief="solid", borderwidth=1, name="light grey")
frame3.pack(padx=25, pady=10, fill=tk.BOTH, expand=True)

label = tk.Label(frame3, text="- 结果输出 -")
label.configure(font=custom_font_bold)
label.place(x=2, y=0)


# 创建文本框用于显示结果
result_text = tk.Text(frame3, state='disabled')
result_text.configure(font=custom_font)


result_text.pack(padx=25, pady=25, fill=tk.BOTH, expand=True)
result_text.pack_propagate(False)


# 运行窗口主循环
window.mainloop()
