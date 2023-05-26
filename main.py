from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import *
from tkinter.messagebox import *
from typing import Dict
from copy import deepcopy

import pdfplumber
from openpyxl import Workbook

from dd import img
import base64,os


class WinGUI(Tk):
    widget_dic: Dict[str, Widget] = {}

    def __init__(self):
        super().__init__()
        tmp = open("tmp.ico", "wb+")
        tmp.write(base64.b64decode(img))  # 写入到临时文件中
        tmp.close()
        super().iconbitmap("tmp.ico")  # 设置图标
        os.remove("tmp.ico")  # 删除临死图标
        self.table_array = []
        self.tk_table = Treeview()
        self.table_columns = {"序号": 60, "上车时间": 150, "城市": 100, "起点": 413, "终点": 413, "金额(元)": 100}
        self.__win()
        self.file_path = StringVar()
        self.total = DoubleVar()
        self.widget_dic["tk_label_li4p21nm"] = self.__tk_label_li4p21nm(self)
        self.widget_dic["tk_button_li4p2kgv"] = self.__tk_button_li4p2kgv(self)
        self.widget_dic["tk_table_li4p5k6c"] = self.__tk_table_li4p5k6c(self)
        self.widget_dic["tk_button_li4p6pmz"] = self.__tk_button_li4p6pmz(self)
        self.widget_dic["tk_label_li4plyoe"] = self.__tk_label_li4plyoe(self)
        self.widget_dic["tk_label_li4uhiwn"] = self.__tk_label_li4uhiwn(self)
        self.widget_dic["tk_label_li4uhwqc"] = self.__tk_label_li4uhwqc(self)

    def __win(self):
        self.title("行程单生成工具")
        # 设置窗口大小、居中
        width = 1280
        height = 741
        screenwidth = self.winfo_screenwidth()
        screenheight = self.winfo_screenheight()
        geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.geometry(geometry)
        self.resizable(width=False, height=False)

    def scrollbar_autohide(self, bar, widget):
        self.__scrollbar_hide(bar, widget)
        widget.bind("<Enter>", lambda e: self.__scrollbar_show(bar, widget))
        bar.bind("<Enter>", lambda e: self.__scrollbar_show(bar, widget))
        widget.bind("<Leave>", lambda e: self.__scrollbar_hide(bar, widget))
        bar.bind("<Leave>", lambda e: self.__scrollbar_hide(bar, widget))

    def __scrollbar_show(self, bar, widget):
        bar.lift(widget)

    def __scrollbar_hide(self, bar, widget):
        bar.lower(widget)

    def __tk_label_li4p21nm(self, parent):
        label = Label(parent, text="行程单：", anchor="center", )
        label.place(x=20, y=20, width=50, height=30)
        return label

    def __tk_button_li4p2kgv(self, parent):
        btn = Button(parent, text="选择", takefocus=False, command=self.openSelectFile)
        btn.place(x=70, y=20, width=88, height=30)
        return btn

    def __tk_button_li4p6pmz(self, parent):
        btn = Button(parent, text="导出", takefocus=False, command=self.export_excel)
        btn.place(x=1170, y=20, width=88, height=30)
        return btn

    def __tk_label_li4plyoe(self, parent):
        label = Label(parent, textvariable=self.file_path, )
        label.place(x=180, y=20, width=646, height=30)
        return label

    def __tk_label_li4uhiwn(self, parent):
        label = Label(parent, text="合计：", anchor="center", foreground='red', font=("黑体", 16, "bold"))
        label.place(x=940, y=20, width=50, height=30)
        return label

    def __tk_label_li4uhwqc(self, parent):
        label = Label(parent, textvariable=self.total, foreground='red', font=("黑体", 16, "bold"))
        label.place(x=990, y=20, width=97, height=30, )
        return label

    def __tk_table_li4p5k6c(self, parent):
        xbar = Scrollbar(parent, orient='horizontal')
        self.tk_table = Treeview(
            parent,
            show="headings",
            columns=list(self.table_columns),
            xscrollcommand=xbar.set,
        )
        xbar['command'] = self.tk_table.xview
        xbar.pack(side='bottom', fill='x')
        for text, width in self.table_columns.items():
            self.tk_table.heading(text, text=text, anchor='center')
            self.tk_table.column(text, anchor='center', width=width, stretch=False)

        self.tk_table.place(x=20, y=70, width=1238, height=653)

    def openSelectFile(self):
        file_path = askopenfilename(defaultextension=".pdf", filetypes=[("PDF", ".pdf")])
        if file_path is not None and file_path != "":
            self.analytic_data(file_path)

    def export_excel(self):
        if not self.table_array:
            showinfo('提示', '没有要导出的数据！')
            return
        wb = Workbook()
        sheet = wb.active
        for row in self.table_array:
            sheet.append(row)
        filepath = asksaveasfilename(defaultextension=".xlsx")
        wb.save(filename=filepath)

        if os.path.exists(filepath):
            showinfo('提示', '导出成功！')
        else:
            showerror('提示', '导出失败，请联系开发人员！')

    def analytic_data(self, file_path):
        self.table_array = []

        def replace_excess(table_array):
            add_column_index = [0, 2, 3, 4, 5, 7]
            atr = []
            for aci in add_column_index:
                atr.append(table_array[aci].replace("\n", ""))
            return atr

        pdf_verify = True
        with pdfplumber.open(file_path) as p:
            for page in p.pages:
                table_rows = deepcopy(page.extract_table())
                if len(table_rows[0]) != 9:
                    pdf_verify = False
                    break
                del table_rows[0]
                for table_row in table_rows:
                    self.table_array.append(replace_excess(table_row))
        if not pdf_verify:
            showwarning("警告", "请上传正确的 滴滴出行-行程单")
            return
        self.file_path.set(file_path)
        self.insert_data()

    def calc_total(self):
        self.total.set(0.0)
        for table_row in self.table_array:
            self.total.set(round(float(self.total.get()) + float(table_row[5]), 2))

    def insert_data(self):
        self.tk_table.delete(*self.tk_table.get_children())
        for data in self.table_array:
            self.tk_table.insert('', END, values=data, )
        self.calc_total()


class Win(WinGUI):
    def __init__(self):
        super().__init__()
        self.__event_bind()

    def __event_bind(self):
        pass


if __name__ == "__main__":
    win = Win()
    win.mainloop()
