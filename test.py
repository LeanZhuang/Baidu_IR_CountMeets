import tkinter as tk
from tkinter import ttk

window = tk.Tk()

# 创建一个Style对象
style = ttk.Style()

# 配置按钮的样式
style.configure('Custom.TButton', background='blue', foreground='grey', font=('Helvetica', 12, 'bold'))

# 创建一个按钮
button = ttk.Button(window, text='点击我', style='Custom.TButton')
button.pack(padx=20, pady=10)

window.mainloop()
