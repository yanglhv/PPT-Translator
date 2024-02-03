import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
from translatorByXfChat import TranslatorByXfChat
import json
import webbrowser
from tkinter import messagebox

# 定义帮助文档
def open_website(event):
    # 打开网页
    webbrowser.open('https://github.com/yanglhv/PPT-Translator')

# 定义保存配置的操作,将配置信息保存到json文档
def save_entry_values():
    # 获取输入框的值
    values = {f'entry{i}': entry[i].get() for i in range(lenConfig)}
    # 保存值到文件
    with open('saved_values.json', 'w') as f:
        json.dump(values, f)
    messagebox.showinfo("API配置", "已保存(saved_values.json)")


# 当文件被拖放到框架上时触发此函数。
# 它接收一个事件作为参数，该事件包含有关被拖放文件的数据。
# 该函数从事件数据中检索文件路径，并使用此路径更新输入字段。
def drop(event):
    filepath = event.data
    if filepath.startswith('{') and filepath.endswith('}'):
        filepath = filepath[1:-1]
    try:
        translator = TranslatorByXfChat(app_id=entry[0].get(),
                                    api_key=entry[1].get(),
                                    api_secret=entry[2].get(),
                                    version=float(entry[3].get()),
                                    target_language=var.get()
                                    )
        feedback = translator.translate_presentation_and_save_new(filepath)
        messagebox.showinfo("结果", feedback)
    except Exception as e:
        messagebox.showerror("错误", e)


#   弹出对话框


# 使用 `TkinterDnD.Tk()` 创建应用程序的主窗口。
# 窗口的标题设置为 'PPT Translator'，其大小设置为 400x200 像素。
root = TkinterDnD.Tk()
root.title('PPT Translator')
root.geometry('300x300')

# 创建一个框架
frame = tk.Frame(root,borderwidth=1, relief="solid")
frame.pack(padx=10, pady=10)  # 设置边距

label0 = tk.Label(frame, text="配置API,选择要翻译为的语言(点我查看手册)", fg="grey", cursor="hand2")
label0.bind("<Button-1>", open_website)
label0.grid(row=0, column=0, columnspan=3)

config =['app_id','api_key','api_secret','version']
lenConfig = len(config)
label = [tk.Label(frame, text=config[i]) for i in range(lenConfig)]
entry = [tk.Entry(frame) for i in range(lenConfig)]
defaultEntryString = ['d83fbada','27438b758e01141255604a15b41b6ac3','YWFhZjFlM2JjOTNlZjEwZmFiMTU3OWM5','1.1']
try:
    with open('saved_values.json', 'r') as f:
        saved_values = json.load(f)
    for i in range(lenConfig):
        entry[i].insert(0, saved_values.get(f'entry{i}', ''))
except FileNotFoundError:
    for i in range(lenConfig):
        entry[i].insert(0, defaultEntryString[i])
for i in range(lenConfig):
    label[i].grid(row=i+1, column=0)
    entry[i].grid(row=i+1, column=1)

var = tk.StringVar()
languages = ['中文', '英文', '日语']
lenlanguages = len(languages)
Radiobutton = [tk.Radiobutton(frame, text=languages[i], variable=var, value=languages[i]) for i in range(lenlanguages)]
for i in range(lenlanguages):
    Radiobutton[i].grid(row=i+1, column=2)
var.set(languages[2])


# 创建一个按钮，点击时调用save_entry_values函数
button = tk.Button(frame, text="保存", command=save_entry_values)
button.grid(row=lenlanguages+1, column=2)

# 使用 `tk.Frame` 创建框架并添加到主窗口。
# 通过调用 `drop_target_register(DND_FILES)` 设置框架接受文件拖放。
# `drop` 函数绑定到 '<<Drop>>' 事件，该事件在文件被拖放到框架上时触发。
label2 = tk.Label(root, text="将要翻译的PPT拖拽到下框内即可", fg="grey")
label2.pack()
frame1 = tk.Frame(root, height=100, width=200, bg='grey',borderwidth=4)
frame1.pack_propagate(False)
frame1.pack()
frame1.drop_target_register(DND_FILES)
frame1.dnd_bind('<<Drop>>', drop)
# 窗口中显示文字


# 在主窗口上调用 `mainloop` 方法以启动 Tkinter 事件循环，
# 该循环使应用程序保持运行并对用户交互做出响应。
root.mainloop()
