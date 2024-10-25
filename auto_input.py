import keyboard
import time
import os
import re
from docx import Document
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading

# 预编译正则表达式
CHINESE_REGEX = re.compile(r'[\u4e00-\u9fff]')

# 判断是否包含中文字符
def contains_chinese(text):
    return CHINESE_REGEX.search(text)

# 逐字输入中文字符（优化延时）
def type_chinese(text):
    for char in text:
        keyboard.write(char, delay=0)  # 设置 delay=0 尽可能快地输入
    # 添加一次性的短延时，确保输入法处理完成
    time.sleep(0.05)

# 读取 .py 文件内容
def read_py_file(file_path):
    """
    读取指定路径的 Python (.py) 文件，并返回所有行的文本列表。
    包括空行。
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
        # 去除行末的换行符，但保留空行
        lines = [line.rstrip('\n') for line in lines]
        return lines
    except Exception as e:
        raise Exception(f"读取 .py 文件时发生错误: {e}")

# 读取 .docx 文件内容
def read_docx_file(file_path):
    """
    读取指定路径的 Word (.docx) 文件，并返回所有段落的文本列表。
    包括空段落。
    """
    try:
        document = Document(file_path)
        text = []
        for para in document.paragraphs:
            # 保留空段落，将其表示为空字符串
            text.append(para.text.rstrip())
        return text
    except Exception as e:
        raise Exception(f"读取 .docx 文件时发生错误: {e}")

# 自动输入内容的函数
def auto_input(lines, log_widget, cancel_event, buttons, progress_widget):
    total = len(lines)
    try:
        # 倒计时并提示用户将鼠标焦点移动到答题区域
        message = "即将开始输入，请确保鼠标焦点在目标输入区域...\n"
        log_widget.insert(tk.END, message)
        log_widget.update()
        for i in range(8, 0, -1):
            message = f"{i}秒后开始...\n"
            log_widget.insert(tk.END, message)
            log_widget.update()
            time.sleep(1)
            if cancel_event.is_set():
                message = "输入已取消。\n"
                log_widget.insert(tk.END, message)
                log_widget.update()
                return

        # 自动输入内容
        for idx, line in enumerate(lines, 1):
            if cancel_event.is_set():
                message = "输入已取消。\n"
                log_widget.insert(tk.END, message)
                log_widget.update()
                return
            if line.strip() == "":
                # 如果是空行，仅按下回车键
                keyboard.press_and_release('enter')
                message = f"第{idx}行: 空行\n"
                log_widget.insert(tk.END, message)
            elif contains_chinese(line):
                # 如果包含中文，逐字输入中文
                type_chinese(line)
                keyboard.press_and_release('enter')   # 每行输入后按回车键
                message = f"第{idx}行: 输入中文\n"
                log_widget.insert(tk.END, message)
            else:
                # 否则直接使用 keyboard.write 输入英文，设置较低的 delay
                keyboard.write(line, delay=0.005)  # 根据需要调整
                keyboard.press_and_release('enter')   # 每行输入后按回车键
                message = f"第{idx}行: 输入英文\n"
                log_widget.insert(tk.END, message)
            # 更新进度条
            progress_widget['value'] = (idx / total) * 100
            log_widget.see(tk.END)
            log_widget.update()
        message = "内容输入完成。\n"
        log_widget.insert(tk.END, message)
    except Exception as e:
        message = f"发生错误: {e}\n"
        log_widget.insert(tk.END, message)
        log_widget.update()
    finally:
        # 恢复按钮状态
        buttons['start'].config(state=tk.NORMAL)
        buttons['cancel'].config(state=tk.DISABLED)

# GUI 界面类
class AutoInputApp:
    def __init__(self, root):
        self.root = root
        self.root.title("自动输入工具")
        self.root.geometry("600x600")  # 增加高度以适应水印和广告
        self.file_path = None
        self.cancel_event = threading.Event()

        # 创建界面组件
        self.create_widgets()

    def create_widgets(self):
        # 说明标签
        instruction = tk.Label(
            self.root,
            text="请选择一个 Python (.py) 或 Word (.docx) 文件，点击 '开始输入' 按钮后，点击目标输入区域并等待几秒，脚本将自动输入文件内容。",
            wraplength=580,
            justify=tk.LEFT
        )
        instruction.pack(pady=10)

        # 输入法提示（根据您的反馈，可以选择保留或删除）
        # 如果不需要输入法提示，可以注释或删除以下代码
        input_method_label = tk.Label(
            self.root,
            text="提示：输入中文前请切换到中文输入法，输入英文前请切换到英文输入法。",
            fg='red',
            font=('Arial', 10)
        )
        input_method_label.pack(pady=5)

        # 文件选择部分
        file_frame = tk.Frame(self.root)
        file_frame.pack(pady=5)

        self.file_label = tk.Label(file_frame, text="未选择文件", width=60, anchor='w')
        self.file_label.pack(side=tk.LEFT, padx=5)

        select_button = tk.Button(file_frame, text="选择文件", command=self.select_file)
        select_button.pack(side=tk.LEFT, padx=5)

        # 开始和取消按钮
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)

        self.start_button = tk.Button(
            button_frame,
            text="开始输入",
            command=self.start_input,
            state=tk.DISABLED,
            width=15
        )
        self.start_button.pack(side=tk.LEFT, padx=10)

        self.cancel_button = tk.Button(
            button_frame,
            text="取消",
            command=self.cancel_input,
            state=tk.DISABLED,
            width=15
        )
        self.cancel_button.pack(side=tk.LEFT, padx=10)

        # 进度条
        progress_label = tk.Label(self.root, text="输入进度:")
        progress_label.pack(pady=5)

        self.progress = ttk.Progressbar(
            self.root,
            orient='horizontal',
            length=580,
            mode='determinate'
        )
        self.progress.pack(pady=5)

        # 日志输出框
        log_label = tk.Label(self.root, text="日志输出:")
        log_label.pack(pady=5)

        self.log_widget = scrolledtext.ScrolledText(
            self.root,
            height=15,
            width=70,
            state='normal'
        )
        self.log_widget.pack(padx=10, pady=5)

        # 广告引流文字
        ad_label = tk.Label(
            self.root,
            text="B站UP主：小约翰可汗t",
            fg='blue',
            font=('Arial', 10, 'italic')
        )
        ad_label.pack(pady=10)

        # 水印
        watermark = tk.Label(
            self.root,
            text="© 2024 小约翰可汗t",
            fg='lightgrey',
            font=('Arial', 12),
            bg='white'
        )
        watermark.place(x=20, y=550)  # 向左移动到左下角

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="选择一个 Python (.py) 或 Word (.docx) 文件",
            filetypes=[("Python 文件", "*.py"), ("Word 文件", "*.docx")]
        )
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=file_path)
            self.start_button.config(state=tk.NORMAL)
            self.log_widget.insert(tk.END, f"已选择文件: {file_path}\n")
            self.log_widget.see(tk.END)

    def start_input(self):
        if not self.file_path:
            messagebox.showwarning("警告", "请先选择一个文件。")
            return

        # 根据文件扩展名选择读取方法
        file_ext = os.path.splitext(self.file_path)[1].lower()
        try:
            if file_ext == '.py':
                lines = read_py_file(self.file_path)
            elif file_ext == '.docx':
                lines = read_docx_file(self.file_path)
            else:
                messagebox.showerror("错误", "不支持的文件格式。请选择 .py 或 .docx 文件。")
                return

            if not lines:
                messagebox.showinfo("信息", "文件中没有可输入的内容。")
                return

        except Exception as e:
            messagebox.showerror("错误", str(e))
            return

        # 禁用开始按钮，启用取消按钮
        self.start_button.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.NORMAL)
        self.cancel_event.clear()

        # 清除之前的日志
        self.log_widget.delete(1.0, tk.END)
        self.log_widget.insert(tk.END, "开始输入...\n")
        self.log_widget.see(tk.END)

        # 启动输入过程的线程
        input_thread = threading.Thread(
            target=auto_input,
            args=(
                lines,
                self.log_widget,
                self.cancel_event,
                {'start': self.start_button, 'cancel': self.cancel_button},
                self.progress
            )
        )
        input_thread.start()

    def cancel_input(self):
        self.cancel_event.set()
        self.cancel_button.config(state=tk.DISABLED)
        self.log_widget.insert(tk.END, "取消请求已发送。\n")
        self.log_widget.see(tk.END)

# 主函数
def main():
    root = tk.Tk()
    app = AutoInputApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
