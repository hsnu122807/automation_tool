from setting import UserInformation
from interface import MenuBar

import tkinter as tk

class MainGUI:
    def __init__(self):
    
        self.user_info = UserInformation.UserInformation()
        self.username = self.user_info.get_current_username()

        self.root = tk.Tk()
        self.root.title("自動化工具程式")
        
        # 設定主畫面元件
        self.label_username = tk.Label(self.root, text=f"登入AD帳號：{self.username}")
        self.label_username.pack()

        self.label_welcome = tk.Label(self.root, text="歡迎使用自動化工具程式！")
        self.label_welcome.pack()
        
        # 設定主畫面的選單列
        self.menu = MenuBar.MenuBar(self.root)
        self.root.config(menu=self.menu.menu_bar)

    def run(self):
        self.root.mainloop()

# 使用範例
gui = MainGUI()
gui.run()
