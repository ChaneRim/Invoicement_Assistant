from tkinter import *
from tkinter.ttk import *


class WinGUI(Tk):
    def __init__(self):
        super().__init__()
        self.__win()
        self.tk_input_custom_info_name = self.__tk_input_custom_info_name(self)
        self.tk_label_custom_info_name = self.__tk_label_custom_info_name(self)
        self.tk_button_confirm = self.__tk_button_confirm(self)

    def __win(self):
        self.title("删除客户信息")
        # 设置窗口大小、居中
        width = 300
        height = 150
        screenwidth = self.winfo_screenwidth()
        screenheight = self.winfo_screenheight()
        geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.geometry(geometry)
        self.resizable(width=False, height=False)


    def __tk_label_custom_info_name(self,parent):
        label = Label(parent,text="客户名称",anchor="center", )
        label.place(x=15, y=15, width=70, height=25)
        return label
    def __tk_input_custom_info_name(self,parent):
        ipt = Entry(parent, )
        ipt.place(x=90, y=15, width=190, height=25)
        return ipt
    def __tk_button_confirm(self,parent):
        btn = Button(parent, text="确定", takefocus=False,)
        btn.place(x=230, y=110, width=50, height=30)
        return btn

class Win(WinGUI):
    def __init__(self, controller):
        self.ctl = controller
        super().__init__()
        self.__event_bind()
        self.ctl.init(self)

    def __event_bind(self):
        self.tk_button_confirm.bind('<Button-1>',self.ctl.confrim_del)
        pass

if __name__ == "__main__":
    win = WinGUI()
    win.mainloop()
