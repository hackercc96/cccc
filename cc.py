import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import openpyxl


class ExcelMatcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 文件匹配")

        self.base_file = None  # 基本文件
        self.match_file = None  # 匹配文件
        self.matched_data = None
        self.base_columns = []  # 基本文件列名
        self.match_columns = []  # 匹配文件列名
        self.common_columns = []  # 存储两个文件中相同的列
        self.column_vars = {}  # 存储每列的BooleanVar，用于复选框

        # 创建主布局：两行两列
        main_frame = tk.Frame(root)
        main_frame.grid(row=0, column=0, sticky="nswe")

        # 设置两列的比例
        root.grid_columnconfigure(0, weight=3)
        root.grid_columnconfigure(1, weight=1)
        root.grid_rowconfigure(0, weight=1)
        root.grid_rowconfigure(1, weight=2)

        # 创建左侧功能区域（功能区）
        left_frame = tk.Frame(main_frame)
        left_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nswe")

        # 创建按钮
        self.load_base_button = tk.Button(left_frame, text="1读取源文件", command=self.load_base_file)
        self.load_base_button.grid(row=0, column=0, pady=10)

        self.load_match_button = tk.Button(left_frame, text="2读取匹配文件", command=self.load_match_file)
        self.load_match_button.grid(row=1, column=0, pady=10)

        # 选择共同列
        self.match_column_label = tk.Label(left_frame, text="3选择共同列:")
        self.match_column_label.grid(row=2, column=0, pady=5)
        self.match_column = tk.StringVar()  # 用于存储选择的共同列
        self.match_column_menu = tk.OptionMenu(left_frame, self.match_column, "")
        self.match_column_menu.grid(row=3, column=0, pady=5)

        # 选择需要匹配的列
        self.column_selection_label = tk.Label(left_frame, text="4选择需要匹配的列：")
        self.column_selection_label.grid(row=4, column=0, pady=5)

        self.column_selection_frame = tk.Frame(left_frame)
        self.column_selection_frame.grid(row=5, column=0, pady=5)

        self.match_button = tk.Button(left_frame, text="5进行匹配", command=self.match_data, state=tk.DISABLED)
        self.match_button.grid(row=6, column=0, pady=10)

        self.download_button = tk.Button(left_frame, text="6下载匹配文件", command=self.download_file, state=tk.DISABLED)
        self.download_button.grid(row=7, column=0, pady=10)

        self.progress = ttk.Progressbar(left_frame, length=200, mode='indeterminate')
        self.progress.grid(row=8, column=0, pady=10)

        # 创建右侧帮助区域（帮助区）
        self.help_frame = tk.Frame(main_frame, width=400, height=400, bg="#f0f0f0", padx=20, pady=20)
        self.help_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nswe")

        self.help_text = """
                                                                   
        在导入前最好可以预先处理一下被导入的excel文件,在文件中最好不要有空白列,不要有合并的单元格,需要匹配处理的数据最好是从1行开始.
        请按照数字引导1-2-3-4-5-6顺序操作.
        1. 加载源文件：
           源文件:你的源数据文件,其他文件(匹配文件)内的数据会被匹配到这个文件内.
          
        2.加载匹配文件：
           匹配文件:这个文件的内容会被匹配到源文件内. 

        3.选择共同列：
           一般来讲两个文件中只有一个共同列,这时不需要手动选择,只有两个文件中存在两个共同列时才需要手动选择.

        4.选择需要匹配的列：
           勾选需要将哪列数据匹配到源文件中。需要哪列数据就选择哪列数据名称. 

        5.进行数据匹配：
           点击“进行匹配”按钮进行匹配，匹配后的数据会显示在下方的预览区域. 
           
        6.下载匹配文件：
           点击“下载匹配文件”按钮保存合并后的Excel文件.
        """
        self.help_label = tk.Label(self.help_frame, text=self.help_text, justify="left", font=("Arial", 10), anchor="nw")
        self.help_label.grid(row=0, column=0, sticky="nswe")

        # 创建下方预览区域（预览区）
        preview_frame = tk.Frame(root)
        preview_frame.grid(row=1, column=0, columnspan=2, padx=20, pady=20, sticky="nswe")

        # 创建 Treeview 控件，用于显示匹配后的数据预览
        self.tree = ttk.Treeview(preview_frame)
        self.tree.grid(row=0, column=0, pady=10, sticky="nswe")

        # 设置预览数据行数的输入框
        self.preview_label = tk.Label(preview_frame, text="显示预览数据的行数：")
        self.preview_label.grid(row=1, column=0, pady=5)

        self.preview_row_count = tk.IntVar()
        self.preview_row_count.set(10)  # 默认显示10行
        self.preview_row_count_entry = tk.Entry(preview_frame, textvariable=self.preview_row_count)
        self.preview_row_count_entry.grid(row=2, column=0, pady=5)

        # 设置列宽
        preview_frame.grid_columnconfigure(0, weight=1)

    def load_excel_file(self, title):
        """通用函数，用于加载Excel文件"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")], title=title)
        if file_path:
            if not os.path.exists(file_path):  # 增加文件路径验证
                messagebox.showerror("错误", "文件路径不存在！")
                return None
            try:
                df = pd.read_excel(file_path, dtype=str)
                print(df.dtypes)
                return df
            except Exception as e:
                messagebox.showerror("错误", f"加载文件失败: {e}\n详细错误信息：{str(e)}")
        return None

    def load_base_file(self):
        """加载基本文件"""
        self.base_file = self.load_excel_file("读取源文件")
        if self.base_file is not None:
            self.base_columns = self.base_file.columns.tolist()
            print(f"源文件列名：{self.base_columns}")
            self.check_common_columns()
            self.check_buttons_enabled()
            self.handle_long_numeric_columns(self.base_file)
            messagebox.showinfo("成功", f"源文件加载成功！")

    def load_match_file(self):
        """加载匹配文件"""
        self.match_file = self.load_excel_file("读取匹配文件")
        if self.match_file is not None:
            self.match_columns = self.match_file.columns.tolist()
            print(f"匹配文件列名：{self.match_columns}")
            self.check_common_columns()
            self.check_buttons_enabled()
            self.update_column_selection()
            self.handle_long_numeric_columns(self.match_file)
            messagebox.showinfo("成功", f"匹配文件加载成功！")

    def handle_long_numeric_columns(self, df):
        """处理长数字列，确保大于11位的数字列被强制转换为字符串"""
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                df[col] = df[col].apply(lambda x: str(int(x)) if pd.notna(x) and x == int(x) else str(x))
            elif df[col].dtype == float:
                df[col] = df[col].apply(lambda x: str(int(x)) if pd.notna(x) and x == int(x) else str(x))

    def check_common_columns(self):
        """检查两个文件中共有的列"""
        print("正在检查共同列...")
        if self.base_file is not None and self.match_file is not None:
            self.common_columns = list(set(self.base_file.columns) & set(self.match_file.columns))
            print(f"共同列：{self.common_columns}")
            if self.common_columns:
                self.update_match_column_menu()  # 更新共同列下拉框
            else:
                messagebox.showwarning("警告", "两个文件没有共同的列！")
                self.match_column_menu['menu'].delete(0, 'end')
        else:
            self.common_columns = []

    def update_match_column_menu(self):
        """更新共同列下拉框"""
        if self.common_columns:
            self.match_column.set(self.common_columns[0])  # 默认选择第一个共同列
            menu = self.match_column_menu["menu"]
            menu.delete(0, "end")
            for column in self.common_columns:
                menu.add_command(label=column, command=lambda col=column: self.match_column.set(col))
        else:
            messagebox.showwarning("警告", "没有共同列可以选择！")

    def update_column_selection(self):
        """更新选择列的区域"""
        if self.match_file is not None:
            for widget in self.column_selection_frame.winfo_children():
                widget.destroy()  # 清空列选择区域

            # 根据匹配文件的列，动态添加勾选框
            self.column_vars = {}  # 清空之前的变量
            for col in self.match_file.columns:
                var = tk.BooleanVar()
                self.column_vars[col] = var
                check_button = tk.Checkbutton(self.column_selection_frame, text=col, variable=var)
                check_button.grid(sticky="w", pady=2)

    def match_data(self):
        """进行数据匹配"""
        if self.base_file is not None and self.match_file is not None:
            common_col = self.match_column.get()
            if common_col:
                # 获取被选中的列
                selected_columns = [col for col, var in self.column_vars.items() if var.get()]
                if selected_columns:
                    # Perform matching logic
                    self.matched_data = pd.merge(self.base_file, self.match_file[selected_columns + [common_col]], on=common_col, how='left')
                    self.show_preview()  # 显示匹配后的预览
                    self.download_button.config(state=tk.NORMAL)
                else:
                    messagebox.showwarning("警告", "请选择至少一个需要匹配的列！")
            else:
                messagebox.showwarning("警告", "请选择共同列进行匹配！")
        else:
            messagebox.showwarning("警告", "请先加载基本文件和匹配文件！")

        # 增加进度条显示
        self.progress.start()
        self.root.update_idletasks()
        time.sleep(1)  # 模拟耗时操作
        self.progress.stop()

    def show_preview(self):
        """显示数据预览"""
        preview_rows = self.preview_row_count.get()
        preview_data = self.matched_data.head(preview_rows)

        # 清除现有的树状视图
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 删除完全为空的列
        preview_data = preview_data.dropna(axis=1, how='all')  # 删除所有完全为空的列

        # 如果第一列是空列（所有值为空），则删除第一列
        if preview_data.iloc[:, 0].isna().all():
            preview_data = preview_data.drop(preview_data.columns[0], axis=1)

        # 确保数字列在显示时去掉 .0
        for col in preview_data.columns:
            if pd.api.types.is_numeric_dtype(preview_data[col]):
                preview_data[col] = preview_data[col].apply(
                    lambda x: str(int(x)) if pd.notna(x) and x == int(x) else str(x))

        # 设置树状视图的列标题
        self.tree["columns"] = list(preview_data.columns)
        self.tree["show"] = "headings"
        for col in preview_data.columns:
            self.tree.heading(col, text=col)
            # 设置列的对齐方式为居中
            self.tree.column(col, anchor="center")

        # 插入数据行
        for index, row in preview_data.iterrows():
            self.tree.insert("", "end", values=list(row))

        # 动态调整列宽
        for col in preview_data.columns:
            max_width = max(preview_data[col].apply(lambda x: len(str(x))).max(), len(col))  # 获取最大宽度
            self.tree.column(col, width=max_width * 10)  # 乘以系数来设置列宽

    def download_file(self):
        """下载匹配后的文件"""
        if self.matched_data is not None:
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if save_path:
                try:
                    # 打印列的数据类型
                    print("列的数据类型:")
                    print(self.matched_data.dtypes)

                    # 创建新的工作簿
                    wb = openpyxl.Workbook()
                    ws = wb.active

                    # 处理列头
                    columns = self.matched_data.columns.tolist()
                    ws.append(columns)

                    # 遍历并填充数据
                    for row in self.matched_data.itertuples(index=False, name=None):
                        ws.append(row)

                    # 对数字列进行处理
                    for col_idx, col_name in enumerate(columns, start=1):
                        for row_idx in range(2, len(self.matched_data) + 2):  # 从第2行开始
                            cell = ws.cell(row=row_idx, column=col_idx)
                            cell_value = cell.value

                            # 如果单元格值是字符串且是数字
                            if isinstance(cell_value, str) and cell_value.isnumeric():
                                # 获取字符串长度
                                if len(cell_value) > 11:  # 如果数字长度大于11
                                    cell.number_format = '@'  # 设置为文本格式
                                    cell.value = str(cell_value)  # 强制将数字作为文本存储
                                else:  # 如果数字长度小于等于11
                                    cell.number_format = '0'  # 设置为常规数字格式
                                    cell.value = int(cell_value)  # 保持为数字格式

                    # 保存文件
                    wb.save(save_path)
                    messagebox.showinfo("成功", f"匹配文件已成功下载：{save_path}")
                except Exception as e:
                    messagebox.showerror("错误", f"下载文件失败: {e}")
            else:
                messagebox.showwarning("警告", "没有选择保存路径！")
        else:
            messagebox.showwarning("警告", "没有匹配数据可下载！")

    def check_buttons_enabled(self):
        """根据文件加载情况启用按钮"""
        if self.base_file is not None and self.match_file is not None:
            self.match_button.config(state=tk.NORMAL)
        else:
            self.match_button.config(state=tk.DISABLED)


if __name__ == "__main__":
    import os
    import time
    root = tk.Tk()
    app = ExcelMatcherApp(root)
    root.mainloop()
