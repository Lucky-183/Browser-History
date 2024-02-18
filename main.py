import re
import csv
import xlwt
from urllib import parse
from os import remove
from browser_history.utils import default_browser,get_browser
import xlrd
from datetime import datetime
import os
from tkinter import ttk  # 导入ttk模块，用于创建ComboBox
import tkinter as tk
from tkinter import messagebox

class BrowserHistoryProcessor:
    def __init__(self):
        self.pattern1 = re.compile(r'q=.*?&')
        self.pattern2 = re.compile(r'wd=.*?&')
        self.url_pattern = re.compile(r'^(?:http[s]?://)?([^/]+)')
        self.url_pattern2 = re.compile(r'[\u4e00-\u9fa5][\u4e00-\u9fa5\s,，。a-zA-Z0-9?!-]{2,10}')
        self.data = []

    def fetch_history(self,browser=""):
        if browser:
            BrowserClass = get_browser(browser)
        else:
            BrowserClass = default_browser()
        if BrowserClass is None:
            print("Could not get default browser!")
            return False
        else:
            b = BrowserClass()
            outputs = b.fetch_history(desc=True)
            outputs.save("history.csv")
            return True

    def process_history(self):
        with open("history.csv", newline='') as f:
            self.data = [row for row in csv.DictReader(f)]
            for i in range(len(self.data)):
                self.data[i]['URL'] = parse.unquote(self.data[i]['URL'])
                self.extract_keyword(str(self.data[i]['URL']),i)

    def extract_keyword(self, url, i):
        if(url.find("bing")!=-1) or (url.find("google")!=-1):
            m = self.pattern1.search(url)
            index1 = 2
            index2 = 1
        elif (url.find("baidu")!=-1):
            m = self.pattern2.search(url)
            index1 = 3
            index2 = 1
        else:
            index1 = 0
            index2 = 0
            m = self.url_pattern2.search(url)

        try:
            self.data[i]['Key']=url[m.span(0)[0] + index1:m.span(0)[1] - index2].replace('+', ' ')
        except:
            self.data[i]['Key']=""

    def save_processed_history_to_csv(self):
        with open('history.csv', 'w', newline='', encoding='UTF-8') as f:
            writer = csv.DictWriter(f, fieldnames=self.data[0].keys())
            writer.writeheader()
            writer.writerows(self.data)


    def csv_to_xlsx(self):
        with open('history.csv', 'r', encoding='UTF-8') as f:
            read = csv.reader(f)
            workbook = xlwt.Workbook()
            sheet = workbook.add_sheet('data')
            for l, line in enumerate(read):
                for r, i in enumerate(line):
                    sheet.write(l, r, i)
            workbook.save('history_temp.xls')  # 临时保存
        remove("history.csv")


    def merge_histories(self, original_file='history.xls', temp_file='history_temp.xls'):
        if not (os.path.exists('history.xls')):
            os.rename('history_temp.xls','history.xls')
        else:
            # 读取原有Excel文件的第一行时间戳
            original_book = xlrd.open_workbook(original_file)
            original_sheet = original_book.sheet_by_index(0)
            original_first_row = original_sheet.row_values(1)
            original_first_timestamp = datetime.strptime(original_first_row[0], '%Y-%m-%d %H:%M:%S').timestamp()
            header_data= [original_sheet.row_values(0)]
            merged_data = [original_sheet.row_values(row) for row in range(1,original_sheet.nrows)]
            # 读取临时Excel文件的第一行时间戳
            temp_book = xlrd.open_workbook(temp_file)
            temp_sheet = temp_book.sheet_by_index(0)


            for row_index in range(1,temp_sheet.nrows):
                temp_timestamp = datetime.strptime(temp_sheet.row_values(row_index)[0], '%Y-%m-%d %H:%M:%S').timestamp()
            # 确定新旧数据的合并范围
                if temp_timestamp > original_first_timestamp:
                    # 合并数据
                    merged_data += [temp_sheet.row_values(row_index)]
                else:
                    break
            # 对合并后的数据按时间戳降序排序
            merged_data.sort(key=lambda x: datetime.strptime(x[0], '%Y-%m-%d %H:%M:%S').timestamp(), reverse=True)

            # 保存到新的Excel文件
            workbook = xlwt.Workbook()
            sheet = workbook.add_sheet('data')
            for row_index, row_data in enumerate(merged_data):
                for col_index, value in enumerate(row_data):
                    sheet.write(row_index, col_index, value)
            workbook.save('history.xls')
            remove('history_temp.xls')

    def run(self):
        if self.fetch_history():
            self.process_history()
            self.save_processed_history_to_csv()
            self.csv_to_xlsx()
            self.merge_histories()



class BrowserHistoryApp(tk.Tk):
    def __init__(self, processor):
        super().__init__()
        self.processor = processor
        self.title('Browser History Processor')
        self.geometry('400x100')
        self.create_widgets()

    def create_widgets(self):
        # 浏览器选择下拉菜单
        self.browser_var = tk.StringVar()
        self.browser_combobox = ttk.Combobox(self, textvariable=self.browser_var)
        self.browser_combobox['values'] = ('Chrome', 'Firefox', 'Edge', 'Opera', 'OperaGX', 'Brave', 'Vivaldi', 'LibreWolf', 'Safari', 'Epic')
        self.browser_combobox['state'] = 'readonly'
        self.browser_combobox.set(default_browser()().name)  # 默认值
        self.browser_combobox.pack(pady=10)

        # 处理并保存历史按钮
        self.process_button = tk.Button(self, text='Fetch and Save History', command=self.fetch_and_save_history)
        self.process_button.pack(pady=10)


    def fetch_and_save_history(self):
        """处理并保存浏览器历史"""
        selected_browser = self.browser_var.get().lower()  # 获取并转换为小写以匹配类名称
        try:
            self.processor.fetch_history(selected_browser)
            self.processor.process_history()
            self.processor.save_processed_history_to_csv()
            self.processor.csv_to_xlsx()
            self.processor.merge_histories()
        except Exception as e:
            messagebox.showerror("Error"," Make sure the browser exist")
            return
        messagebox.showinfo("Info", "Finished processing and saving browser history.")

if __name__ == '__main__':
    processor = BrowserHistoryProcessor()
    app = BrowserHistoryApp(processor)
    app.mainloop()

