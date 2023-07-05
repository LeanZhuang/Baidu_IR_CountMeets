import tkinter as tk
# import ctypes
import tkinter.filedialog as fd
import tkinter.messagebox as messagebox
from run_code import run_code


def browse_folder():
    folder_path = fd.askdirectory()
    if folder_path:
        new_value = new_entry.get()
        
        try:
            result1, result2 = run_code(folder_path, new_value)
            show_results(result1, result2)
        except:
            result1, result2 = '-------------------------\n--------- ERROR ---------\n-------------------------\n\n请检查文件夹内是否有正确的文件', ''
            show_results(result1, result2)


def show_results(result1, result2):
    new_value = new_entry.get()
    if len(new_value) != 4 or not new_value.isdigit():
            text = '-------------------------\n--------- ERROR ---------\n-------------------------\n\n请输入正确的日期格式，如：0621'
    else:
        text = result1 + '\n' + result2

    result_text.config(state='normal')  # 允许编辑
    result_text.delete(1.0, tk.END)  # 清空文本框内容
    result_text.insert(tk.END, text)  # 插入结果文本
    result_text.config(state='disabled')  # 禁止编辑
            

# 创建窗口
window = tk.Tk()
window.title("计算程序")
window.geometry("800x600")

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
warning_label = tk.Label(window, text="注意底稿文件命名规范", fg="red")
warning_label.pack()

# 创建文本框用于显示结果
result_text = tk.Text(window, state='disabled')
result_text.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)

# 运行窗口主循环
window.mainloop()
