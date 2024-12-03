import tkinter as tk
from tkinter import filedialog, messagebox
import importlib

# 全局变量，存储用户上传的文件路径
selected_files = []


def upload_files():
    global selected_files
    files = filedialog.askopenfilenames(
        title="选择文件",
        filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
    )
    if files:
        selected_files = list(files)
        # 在文件显示区域显示上传的文件
        file_list_text.delete(1.0, tk.END)
        file_list_text.insert(tk.END, "\n".join(selected_files))


def clear_files():
    """清除已上传的文件"""
    global selected_files
    selected_files.clear()
    file_list_text.delete(1.0, tk.END)
    messagebox.showinfo("清除成功", "已清除所有上传的文件。")


def load_modules():
    try:
        # 读取配置文件，格式为：模块名 文件名
        with open("config.txt", "r", encoding="utf-8") as f:
            for line in f:
                module_name, file_name = line.strip().split()
                create_module_button(module_name, file_name)
    except FileNotFoundError:
        messagebox.showerror("错误", "未找到配置文件 config.txt")
    except Exception as e:
        messagebox.showerror("错误", f"加载模块失败：{e}")


def create_module_button(module_name, file_name):
    def load_module():
        try:
            module = importlib.import_module(file_name.replace(".py", ""))
            module.run(selected_files)  # 调用模块的主入口函数
        except Exception as e:
            messagebox.showerror("错误", f"加载模块 {module_name} 时出错：{e}")

    # 创建按钮并动态添加到网格布局
    global module_row, module_col
    button = tk.Button(module_frame, text=module_name, command=load_module, width=12)
    button.grid(row=module_row, column=module_col, padx=10, pady=5)
    module_col += 1
    if module_col >= max_columns:  # 超过最大列数换行
        module_col = 0
        module_row += 1


# 主界面
root = tk.Tk()
root.title("Excel 工具集")
root.geometry("600x400")

# 上传和清除文件按钮容器
top_frame = tk.Frame(root)
top_frame.pack(pady=10)

# 上传文件按钮
upload_button = tk.Button(top_frame, text="上传文件", command=upload_files)
upload_button.pack(side=tk.LEFT, padx=5)

# 清除文件按钮
clear_button = tk.Button(top_frame, text="清除文件", command=clear_files)
clear_button.pack(side=tk.LEFT, padx=5)

# 文件列表显示区域
file_list_text = tk.Text(root, height=10, width=70)
file_list_text.pack(pady=10)

# 动态模块加载区域
module_frame = tk.Frame(root)
module_frame.pack(pady=10)

# 初始化模块按钮布局参数
module_row = 0  # 当前行号
module_col = 0  # 当前列号
max_columns = 4  # 每行最多显示的按钮数量

# 加载模块按钮
load_modules()

# 启动主循环
root.mainloop()
