import datetime
import time
import threading
import os
import sys
import json
from tkinter import messagebox

class ExcelReminderApp:
    def __init__(self, excel_path, time_column, content_columns=None):
        self.excel_path = excel_path
        self.time_column = time_column
        self.content_columns = content_columns or []
        self.today_data = []
        self.stop_event = threading.Event()
        self.check_thread = None
        self.cache_file = 'data_cache.json'

    def load_today_data(self):
        today = datetime.date.today()
        try:
            with open(self.cache_file, 'r') as f:
                cache = json.load(f)
                if cache.get('date') == str(today):
                    self.today_data = cache.get('data', [])
                    return True, f"成功从缓存加载 {len(self.today_data)} 条今日记录"
        except (FileNotFoundError, json.JSONDecodeError):
            pass

        try:
            if not os.path.exists(self.excel_path):
                return False, f"错误：Excel文件 '{self.excel_path}' 不存在"

            file_ext = os.path.splitext(self.excel_path)[1].lower()

            if file_ext == '.xlsx':
                from openpyxl import load_workbook
                wb = load_workbook(self.excel_path, data_only=True)
                ws = wb.active

                time_col_idx = None
                content_col_indices = {}

                for col_idx, cell in enumerate(ws[1], 1):
                    if cell.value == self.time_column:
                        time_col_idx = col_idx
                    if cell.value in self.content_columns:
                        content_col_indices[cell.value] = col_idx

                if not time_col_idx:
                    return False, f"找不到时间列 '{self.time_column}'"

                self.today_data = []
                today = datetime.date.today()

                for row in ws.iter_rows(min_row=2, values_only=True):
                    time_value = row[time_col_idx - 1]

                    if time_value:
                        if isinstance(time_value, str):
                            try:
                                time_obj = datetime.datetime.strptime(time_value, '%Y-%m-%d %H:%M:%S')
                            except ValueError:
                                try:
                                    time_obj = datetime.datetime.strptime(time_value, '%Y-%m-%d')
                                except ValueError:
                                    print(f"无法解析时间: {time_value}")
                                    continue
                        elif isinstance(time_value, datetime.datetime):
                            time_obj = time_value
                        elif isinstance(time_value, datetime.date):
                            time_obj = datetime.datetime.combine(time_value, datetime.time())
                        else:
                            print(f"未知时间格式: {time_value}")
                            continue

                        if time_obj.date() == today:
                            record = {'时间': time_obj}
                            for col_name, col_idx in content_col_indices.items():
                                record[col_name] = row[col_idx - 1] if col_idx - 1 < len(row) else None
                            self.today_data.append(record)

                wb.close()

            elif file_ext == '.xls':
                import pandas as pd
                df = pd.read_excel(self.excel_path)

                if self.time_column not in df.columns:
                    return False, f"找不到时间列 '{self.time_column}'"

                try:
                    df['datetime'] = pd.to_datetime(df[self.time_column])
                except:
                    return False, f"无法解析时间列 '{self.time_column}'"

                today = datetime.date.today()
                today_start = datetime.datetime.combine(today, datetime.time.min)
                today_end = datetime.datetime.combine(today, datetime.time.max)

                today_df = df[(df['datetime'] >= today_start) & (df['datetime'] <= today_end)]

                self.today_data = []
                for _, row in today_df.iterrows():
                    record = {'时间': row['datetime'].to_pydatetime()}
                    for col in self.content_columns:
                        record[col] = row[col] if col in row else None
                    self.today_data.append(record)
            else:
                return False, f"错误：不支持的文件类型 '{file_ext}'"

            cache = {
                'date': str(today),
                'data': self.today_data
            }
            with open(self.cache_file, 'w') as f:
                json.dump(cache, f, default=str)

            return True, f"成功加载 {len(self.today_data)} 条今日记录"

        except Exception as e:
            return False, f"加载数据时出错: {str(e)}"

    def start_refreshing(self, interval=60):
        if self.check_thread and self.check_thread.is_alive():
            return

        self.check_thread = threading.Thread(target=self._refresh_loop, args=(interval,))
        self.check_thread.daemon = True
        self.check_thread.start()

    def stop_refreshing(self):
        self.stop_event.set()
        if self.check_thread:
            self.check_thread.join(timeout=1.0)

    def _refresh_loop(self, interval):
        while not self.stop_event.is_set():
            success, message = self.load_today_data()
            print(message)
            time.sleep(interval)

class ExcelReminderGUI:
    def __init__(self, root, excel_path, time_column, content_columns=None, subtitle=None):
        self.root = root
        self.root.title("小美的预约系统")
        self.root.geometry("800x600")
        self.root.configure(bg="#f0f0f0")

        self.app = ExcelReminderApp(excel_path, time_column, content_columns)

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.subtitle = subtitle

        self.create_widgets()
        self.load_data()

        self.auto_refresh_var = tk.BooleanVar()
        self.auto_refresh_var.set(False)

        self.tree = None

    def create_widgets(self):
        title_frame = tk.Frame(self.root, bg="#f0f0f0")
        title_frame.pack(pady=10, fill=tk.X)

        title_label = tk.Label(title_frame, text="小美的预约系统",
                               font=("微软雅黑", 18, "bold"), bg="#f0f0f0")
        title_label.pack(side=tk.LEFT, padx=20)

        subtitle_text = self.subtitle or datetime.datetime.now().strftime("%Y年%m月%d日")
        subtitle_label = tk.Label(title_frame, text=subtitle_text,
                                  font=("微软雅黑", 10), bg="#f0f0f0", fg="#666666")
        subtitle_label.pack(side=tk.RIGHT, padx=20)

        button_frame = tk.Frame(self.root, bg="#f0f0f0")
        button_frame.pack(pady=5, fill=tk.X)

        auto_refresh_button = tk.Button(button_frame, text="自动刷新",
                                        command=self.toggle_auto_refresh,
                                        font=("微软雅黑", 10), bg="#4CAF50", fg="white",
                                        padx=15, pady=5)
        auto_refresh_button.pack(side=tk.LEFT, padx=5)

        refresh_button = tk.Button(button_frame, text="刷新数据",
                                   command=self.load_data,
                                   font=("微软雅黑", 10), bg="#2196F3", fg="white",
                                   padx=15, pady=5)
        refresh_button.pack(side=tk.LEFT, padx=5)

        exit_button = tk.Button(button_frame, text="退出",
                                command=self.on_close,
                                font=("微软雅黑", 10), bg="#f44336", fg="white",
                                padx=15, pady=5)
        exit_button.pack(side=tk.LEFT, padx=5)

        self.status_var = tk.StringVar()
        self.status_var.set("准备加载数据...")
        status_label = tk.Label(button_frame, textvariable=self.status_var,
                                font=("微软雅黑", 10), bg="#f0f0f0", fg="blue")
        status_label.pack(side=tk.RIGHT, padx=20)

    def load_data(self):
        self.status_var.set("正在加载数据...")
        self.root.update()

        success, message = self.app.load_today_data()

        if success:
            if self.tree is None:
                columns = ["时间"] + self.app.content_columns
                self.tree = ttk.Treeview(self.root, columns=columns, show="headings")
                self.tree.column("时间", width=150)
                self.tree.heading("时间", text="时间")
                for col in self.app.content_columns:
                    self.tree.column(col, width=150)
                    self.tree.heading(col, text=col)
                scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
                self.tree.configure(yscroll=scrollbar.set)
                self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=20, pady=10)
                scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=10)

            for item in self.tree.get_children():
                self.tree.delete(item)

            for record in self.app.today_data:
                values = [record["时间"].strftime("%Y-%m-%d %H:%M:%S")]
                for col in self.app.content_columns:
                    values.append(record.get(col, ""))
                self.tree.insert("", tk.END, values=values)

            self.status_var.set(f"加载完成，今日共有 {len(self.app.today_data)} 条记录")
        else:
            self.status_var.set(f"加载失败: {message}")
            messagebox.showerror("错误", message)

    def toggle_auto_refresh(self):
        if self.auto_refresh_var.get():
            self.app.stop_refreshing()
            self.status_var.set("自动刷新已关闭")
            self.auto_refresh_var.set(False)
        else:
            self.app.start_refreshing(interval=3600)
            self.status_var.set("自动刷新已开启 (每5小时)")
            self.auto_refresh_var.set(True)

    def on_close(self):
        if messagebox.askyesno("确认", "确定要退出程序吗？"):
            self.app.stop_refreshing()
            self.root.destroy()

def main():
    excel_path = "/Users/Sun/Desktop/预约/患者管理登记表.xlsx"
    #  excel_path = "患者管理登记表.xlsx"  # 请替换为你的Excel文件路径
    time_column = "复诊时间"  # 请替换为你的时间列名
    content_columns = ["姓名", "处置", "余留问题"]  # 请替换为你要显示的列名
    subtitle = "栋哥特约版V1.0"  # 自定义副标题内容

    silent_mode = '--silent' in sys.argv

    if silent_mode:
        app = ExcelReminderApp(excel_path, time_column, content_columns)
        success, message = app.load_today_data()
        if success:
            print(f"今日共有 {len(app.today_data)} 条记录")
            for record in app.today_data:
                time_str = record["时间"].strftime("%Y-%m-%d %H:%M:%S")
                print(f"{time_str}: {', '.join([f'{k}: {v}' for k, v in record.items() if k != '时间'])}")
        else:
            print(f"错误: {message}")

        app.start_refreshing(interval=300)

        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            app.stop_refreshing()
            print("程序已退出")
    else:
        import tkinter as tk
        from tkinter import ttk
        root = tk.Tk()
        app = ExcelReminderGUI(root, excel_path, time_column, content_columns, subtitle)
        root.mainloop()

if __name__ == "__main__":
    main()
