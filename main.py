# main


import os
import data_center
# 导入布局文件
from view.main_ui import Win as MainWin
# 导入窗口控制器
from controller.main_controller import Controller as MainUIController

if not os.path.exists('file/config.json'):
    print('不存在文件 config.json ，正在创建......')
    data = {"save_path": "", "invoice_table_path": "", "custom_info_path": "", "sample_path": ""}
    json_dir = data_center.get_script_dir()
    json_dir = os.path.join(json_dir, 'file/config.json')
    data_center.save_json(json_dir, data)

# 将窗口控制器 传递给UI
app = MainWin(MainUIController())

if __name__ == "__main__":
    # 原神，启动！
    app.mainloop()
