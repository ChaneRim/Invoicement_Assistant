# main_controller

from win32com import client
from controller.custom_info_add_controller import Controller as addController
from controller.custom_info_del_controller import Controller as delController
from controller.custom_info_change_controller import Controller as changeController
from view.custom_info_add_ui import Win as addWin
from view.custom_info_del_ui import Win as delWin
from view.custom_info_change_ui import Win as changeWin
from view.main_ui import Win
from core.tax_core import *
from tkinter import filedialog, messagebox
import tkinter as tk
import data_center
import os


class Controller:
    # 导入UI类后，替换以下的 object 类型，将获得 IDE 属性提示功能
    ui: Win
    def __init__(self):
        self.sample = {"save_path": "", "invoice_table_path": "", "custom_info_path": "", "sample_path": ""}
        pass

    def init(self, ui):
        """
        得到UI实例，对组件进行初始化配置
        """
        self.ui = ui
        # TODO 组件初始化 赋值操作

    def update_status(self, message):
        self.ui.status_var.set(message)
    def open_wps(self,key):
        path = data_center.get_value(self.ui.json_path, key)
        try:
            app = client.Dispatch("ket.Application")
            app.Visible = True
            wb = app.WorkBooks.Open(path)
        except:
            messagebox.showinfo('错误', '路径错误或不存在该文件！')
    def xlsx_select(self,key):
        file_path = tk.filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            data_center.change_value(self.ui.json_path, key, file_path)
            self.ui.tax_path_var.set(file_path)  # 实时更新 Entry 的内容
            self.update_status(f"客户信息路径更改为：{file_path}")
# 客户信息
    def custom_info_select(self,message):
        self.xlsx_select('custom_info_path')
    def custom_info_check(self,message):
        self.open_wps('custom_info_path')
    def custom_info_add(self,message):
        add_win = addWin(addController())
        add_win.mainloop()
    def custom_info_del(self,message):
        del_win = delWin(delController())
        del_win.mainloop()
    def custom_info_change(self,message):
        change_win = changeWin(changeController())
        change_win.mainloop()
# 发货单样板
    def sample_select(self,message):
        self.xlsx_select('sample_path')
    def sample_check(self,message):
        self.open_wps('sample_path')
# 发票明细表
    def invoice_table_select(self,message):
        self.xlsx_select('invoice_table_path')

    def invoice_table_check(self, message):
        self.open_wps('invoice_table_path')
# 保存路径
    def save_path_select(self, message):
        file_path = tk.filedialog.askdirectory()
        if file_path:
            data_center.change_value(self.ui.json_path, 'save_path', file_path)
            self.ui.tax_path_var.set(file_path)  # 实时更新 Entry 的内容
            self.update_status(f"保存路径更改为：{file_path}")
    def save_path_open(self, message):
        path = data_center.get_value(self.ui.json_path, 'save_path')
        try:
            os.startfile(path)
            self.update_status(f"打开保存路径 : {path}")
        except:
            messagebox.showinfo('路径错误', '路径错误或不存在该路径！')

    def clean_cache(self, message):
        dir_main_ctl = data_center.get_script_dir()
        json_path = os.path.join(dir_main_ctl,r'file/config.json')
        data_center.save_json(json_path, self.sample)
        messagebox.showinfo("提示","清除缓存成功！")

    def start(self, message):
        try:
            dir_main_ctl = data_center.get_script_dir()
            config_path = os.path.join(dir_main_ctl, r"file/config.json")
            invoice_path = data_center.get_value(config_path, "invoice_table_path")
            __custom_info_path = data_center.get_value(config_path, "custom_info_path")
            __save_path = data_center.get_value(config_path, "save_path")
            __sample_path = data_center.get_value(config_path, "sample_path")
            __valid_col_df = get_valid_frame(invoice_path)
            __ws = to_excel(__valid_col_df).active
            __dic = create_dic(__ws)
            __invalid_dic = qty_filter(__dic)
            __valid_data = null_cleaner(__dic, __invalid_dic)
            __final_data = dic_input(__valid_data, __custom_info_path)
            create_excel(__save_path, __sample_path, __final_data)
            self.update_status(f"文件已生成于目录：{__save_path} 中。")

        except FileNotFoundError as e:
            messagebox.showinfo("错误", f"未找到文件,请检查！\n{e}")
        except Exception as e:
            messagebox.showinfo("错误", f"发生未知错误！\n{e}")

