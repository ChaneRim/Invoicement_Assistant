from view.custom_info_add_ui import Win
import data_center
import os
from tkinter import messagebox
class Controller:
    # 导入UI类后，替换以下的 object 类型，将获得 IDE 属性提示功能
    ui: Win
    def __init__(self):
        pass
    def init(self, ui):
        """
        得到UI实例，对组件进行初始化配置
        """
        self.ui = ui
        # TODO 组件初始化 赋值操作
    def confrim_add(self,evt):
        dir_main = data_center.get_script_dir()
        config_path = os.path.join(dir_main, "file/config.json")
        custom_info_path = data_center.get_value(config_path, "custom_info_path")
        name = self.ui.tk_input_custom_info_name.get()
        address = self.ui.tk_input_custom_info_address.get()
        phone = self.ui.tk_input_custom_info_phone.get()
        data_center.add_item(custom_info_path, [name, address, phone])
        messagebox.showinfo('提示', '保存成功')
        self.ui.destroy()

