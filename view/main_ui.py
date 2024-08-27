# main_ui
import os.path
from tkinter import *
# from tkinter.ttk import *
import data_center


class WinGUI(Tk):
    
    def __init__(self):
        super().__init__()
        self.__win()
        self.json_path = os.path.join(data_center.get_script_dir(), r"file/config.json")

        self.tax_path_var = StringVar()
        self.tax_path_var.set(data_center.get_value(self.json_path, 'invoice_table_path'))
        self.save_path_var = StringVar()
        self.save_path_var.set(data_center.get_value(self.json_path, 'save_path'))
        # 客户信息
        self.tk_label_custom_info = self.__tk_label_custom_info(self)
        self.tk_button_custom_info_select = self.__tk_button_custom_info_select(self)
        self.tk_button_custom_info_check = self.__tk_button_custom_info_check(self)
        self.tk_button_custom_info_add = self.__tk_button_custom_info_add(self)
        self.tk_button_custom_info_del = self.__tk_button_custom_info_del(self)
        self.tk_button_custom_info_change = self.__tk_button_custom_info_change(self)
        # 发货单样板
        self.tk_label_sample = self.__tk_label_sample(self)
        self.tk_button_sample_select = self.__tk_button_sample_select(self)
        self.tk_button_sample_check = self.__tk_button_sample_check(self)
        # 发票明细表
        self.tk_label_invoice_table = self.__tk_label_invioce_table(self)
        self.tk_button_invoice_table_select = self.__tk_button_invoice_table_select(self)
        self.tk_button_invoice_table_check = self.__tk_button_invoice_table_check(self)
        # 保存路径
        self.tk_label_save_path = self.__tk_label_save_path(self)
        self.tk_button_save_path_select = self.__tk_button_save_path_select(self)
        self.tk_button_save_path_open = self.__tk_button_save_path_open(self)
        # 状态栏
        self.status_var = StringVar()
        self.status_var.set("就绪")  # 初始状态
        self.tk_status_bar = self.__tk_status_bar(self)

        # 清除缓存
        self.tk_button_clean_cache = self.__tk_button_clean_cache(self)
        # 原神，启动！
        self.tk_button_start = self.__tk_button_start(self)

    def __win(self):
        self.title("送货单提取助手")
        # 设置窗口大小、居中
        width = 500
        height = 300
        screenwidth = self.winfo_screenwidth()
        screenheight = self.winfo_screenheight()
        geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.geometry(geometry)
        self.resizable(width=False, height=False)

# 客户信息
    def __tk_label_custom_info(self, parent):
        label = Label(parent, text="客 户 信 息 :", anchor="w", )
        label.place(x=40, y=20, width=80, height=25)
        return label
    def __tk_button_custom_info_select(self, parent):
        btn = Button(parent, text="选择", relief="raised", overrelief="groove")
        btn.place(x=120, y=20, width=50, height=25)
        return btn
    def __tk_button_custom_info_check(self, parent):
        btn = Button(parent, text="查看", relief="raised", overrelief="groove")
        btn.place(x=170, y=20, width=50, height=25)
        return btn
    def __tk_button_custom_info_add(self, parent):
        btn = Button(parent, text="增加", relief="raised", overrelief="groove")
        btn.place(x=220, y=20, width=50, height=25)
        return btn
    def __tk_button_custom_info_del(self, parent):
        btn = Button(parent, text="删除", relief="raised", overrelief="groove")
        btn.place(x=270, y=20, width=50, height=25)
        return btn
    def __tk_button_custom_info_change(self, parent):
        btn = Button(parent, text="更改", relief="raised", overrelief="groove")
        btn.place(x=320, y=20, width=50, height=25)
        return btn

# 发货单样板
    def __tk_label_sample(self, parent):
        label = Label(parent, text="发货单样板 :", anchor="w", )
        label.place(x=40, y=50, width=80, height=25)
        return label
    def __tk_button_sample_select(self, parent):
        btn = Button(parent, text="选择", relief="raised", overrelief="groove")
        btn.place(x=120, y=50, width=50, height=25)
        return btn
    def __tk_button_sample_check(self, parent):
        btn = Button(parent, text="查看", relief="raised", overrelief="groove")
        btn.place(x=170, y=50, width=50, height=25)
        return btn
# 发票明细表
    def __tk_label_invioce_table(self, parent):
        label = Label(parent, text="发票明细表 :", anchor="w", )
        label.place(x=40, y=80, width=80, height=25)
        return label
    def __tk_button_invoice_table_select(self, parent):
        btn = Button(parent, text="选择", relief="raised", overrelief="groove")
        btn.place(x=120, y=80, width=50, height=25)
        return btn
    def __tk_button_invoice_table_check(self, parent):
        btn = Button(parent, text="查看", relief="raised", overrelief="groove")
        btn.place(x=170, y=80, width=50, height=25)
        return btn
# 保存路径
    def __tk_label_save_path(self, parent):
        label = Label(parent, text="保 存 路 径 :", anchor="w", )
        label.place(x=40, y=110, width=80, height=25)
        return label
    def __tk_button_save_path_select(self, parent):
        btn = Button(parent, text="选择", relief="raised", overrelief="groove")
        btn.place(x=120, y=110, width=50, height=25)
        return btn
    def __tk_button_save_path_open(self, parent):
        btn = Button(parent, text="打开", relief="raised", overrelief="groove")
        btn.place(x=170, y=110, width=50, height=25)
        return btn

# 状态条
    def __tk_status_bar(self, parent):
        status_bar = Label(parent, textvariable=self.status_var, relief=SUNKEN, anchor=W)
        status_bar.pack(side=BOTTOM, fill=X)
        return status_bar

# 清除缓存
    def __tk_button_clean_cache(self, parent):
        btn = Button(parent, text="清除缓存", font=("微软雅黑", 10), relief="raised", overrelief="groove")
        btn.place(x=340, y=200, width=70, height=30)
        return btn

# 原神，启动！
    def __tk_button_start(self, parent):
        btn = Button(parent, text="开始处理", font=("等线", 20, "bold"), relief="raised", overrelief="groove")
        btn.place(x=40, y=200, width=120, height=30)
        return btn


class Win(WinGUI):
    def __init__(self, controller):
        self.ctl = controller
        super().__init__()
        self.__event_bind()
        self.ctl.init(self)
    def __event_bind(self):
        # 客户信息
        self.tk_button_custom_info_select.bind('<Button-1>', self.ctl.custom_info_select)
        self.tk_button_custom_info_check.bind('<Button-1>', self.ctl.custom_info_check)
        self.tk_button_custom_info_add.bind('<Button-1>', self.ctl.custom_info_add)
        self.tk_button_custom_info_del.bind('<Button-1>', self.ctl.custom_info_del)
        self.tk_button_custom_info_change.bind('<Button-1>', self.ctl.custom_info_change)
        # 发货单样板
        self.tk_button_sample_select.bind('<Button-1>', self.ctl.sample_select)
        self.tk_button_sample_check.bind('<Button-1>', self.ctl.sample_check)
        # 发票明细表
        self.tk_button_invoice_table_select.bind('<Button-1>', self.ctl.invoice_table_select)
        self.tk_button_invoice_table_check.bind('<Button-1>', self.ctl.invoice_table_check)
        # 保存路径
        self.tk_button_save_path_select.bind('<Button-1>', self.ctl.save_path_select)
        self.tk_button_save_path_open.bind('<Button-1>', self.ctl.save_path_open)
        # 清除缓存
        self.tk_button_clean_cache.bind('<Button-1>', self.ctl.clean_cache)
        # 原神，启动！
        self.tk_button_start.bind('<Button-1>', self.ctl.start)
        pass


if __name__ == "__main__":
    win = WinGUI()
    win.mainloop()
