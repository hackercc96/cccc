import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

from openpyxl import Workbook
from openpyxl.styles import NamedStyle
from openpyxl.utils.dataframe import dataframe_to_rows

from openpyxl import Workbook
from openpyxl.styles import NamedStyle
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

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
        self.load_base_button = tk.Button(left_frame, text="读取基本文件", command=self.load_base_file)
        self.load_base_button.grid(row=0, column=0, pady=10)

        self.load_match_button = tk.Button(left_frame, text="读取匹配文件", command=self.load_match_file)
        self.load_match_button.grid(row=1, column=0, pady=10)

        # 选择共同列
        self.match_column_label = tk.Label(left_frame, text="选择共同列:")
        self.match_column_label.grid(row=2, column=0, pady=5)
        self.match_column = tk.StringVar()  # 用于存储选择的共同列
        self.match_column_menu = tk.OptionMenu(left_frame, self.match_column, "")
        self.match_column_menu.grid(row=3, column=0, pady=5)

        # 选择需要匹配的列
        self.column_selection_label = tk.Label(left_frame, text="选择需要匹配的列：")
        self.column_selection_label.grid(row=4, column=0, pady=5)

        self.column_selection_frame = tk.Frame(left_frame)
        self.column_selection_frame.grid(row=5, column=0, pady=5)

        self.match_button = tk.Button(left_frame, text="进行匹配", command=self.match_data, state=tk.DISABLED)
        self.match_button.grid(row=6, column=0, pady=10)

        self.download_button = tk.Button(left_frame, text="下载匹配文件", command=self.download_file, state=tk.DISABLED)
        self.download_button.grid(row=7, column=0, pady=10)

        self.progress = ttk.Progressbar(left_frame, length=200, mode='indeterminate')
        self.progress.grid(row=8, column=0, pady=10)

        # 创建右侧帮助区域（帮助区）
        self.help_frame = tk.Frame(main_frame, width=400, height=400, bg="#f0f0f0", padx=20, pady=20)
        self.help_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nswe")

        self.help_text = """
        1. 读取基本文件和匹配文件：
        使用“读取基本文件(进行匹配的原始文件数据以此文件为基准)”和“读取匹配文件(匹配文件中的内容将会匹配到基本文件当中)”按钮加载两个Excel文件。

        2. 选择共同列：
           在“选择共同列”下拉框中选择两个文件中共同的列。两个文件当中必须要有相同的列数据作为匹配标准.

        3. 选择需要匹配的列：
           在“选择需要匹配的列”区域勾选需要合并的列。需要将匹配文件中哪列数据匹配到基准文件中就选择那列数据名称.

        4. 进行数据匹配：
           点击“进行匹配”按钮进行匹配，匹配后的数据会显示在下方的预览区域。

        5. 下载匹配文件：
           点击“下载匹配文件”按钮保存合并后的Excel文件。

        6. 注意事项：
           - 请确保文件格式为Excel格式（.xlsx）。
           - 请选择至少一个需要匹配的列。
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

    def load_base_file(self):
        """加载基本文件"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if file_path:
            try:
                # 修改为所有列都读取为字符串
                self.base_file = pd.read_excel(file_path, dtype=str)
                # 打印数据类型，确认是否为字符串类型
                print(self.base_file.dtypes)
                self.base_columns = self.base_file.columns.tolist()
                print(f"基本文件列名：{self.base_columns}")
                self.check_common_columns()  # 检查并更新共同列
                self.check_buttons_enabled()
                self.handle_long_numeric_columns(self.base_file)  # 检查并处理长数字列
                messagebox.showinfo("成功", f"基本文件加载成功！文件路径：{file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"加载基本文件失败: {e}")

    def load_match_file(self):
        """加载匹配文件"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if file_path:
            try:
                self.match_file = pd.read_excel(file_path, dtype=str)
                print(self.match_file.dtypes)
                self.match_columns = self.match_file.columns.tolist()
                print(f"匹配文件列名：{self.match_columns}")
                self.check_common_columns()  # 检查并更新共同列
                self.check_buttons_enabled()
                self.update_column_selection()  # 更新选择匹配列
                self.handle_long_numeric_columns(self.match_file)  # 检查并处理长数字列
                messagebox.showinfo("成功", f"匹配文件加载成功！文件路径：{file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"加载匹配文件失败: {e}")

    def handle_long_numeric_columns(self, df):
        """处理长数字列，确保大于11位的数字列被强制转换为字符串"""
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                df[col] = df[col].apply(lambda x: str(int(x)) if pd.notna(x) and x == int(x) else str(x))
            # 强制将可能会丢失精度的数字列转为字符串
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
                menu.add_command(label=column, command=tk._setit(self.match_column, column))
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
                    wb = Workbook()
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
    root = tk.Tk()
    app = ExcelMatcherApp(root)
    root.mainloop()
