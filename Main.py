from setting import UserInformation
from interface import MenuBar
from interface import LogDisplay

import tkinter as tk

class MainGUI:
    def __init__(self):
    
        self.user_info = UserInformation.UserInformation()
        self.username = self.user_info.get_current_username()

        self.root = tk.Tk()
        self.root.title("自動化工具程式")
        self.root.geometry("1024x768")  # 設定畫面大小為 1024x768
        self.root.iconphoto(False, tk.PhotoImage(file='tree.png'))
        
        # 設定主畫面元件
        self.label_username = tk.Label(self.root, text=f"登入AD帳號：{self.username}")
        self.label_username.pack()

        self.label_welcome = tk.Label(self.root, text="歡迎使用自動化工具程式！")
        self.label_welcome.pack()
        
        # 設定主畫面的選單列
        self.menu = MenuBar.MenuBar(self.root)
        self.root.config(menu=self.menu.menu_bar)
        
        # 設定Log
        #self.log_display = LogDisplay.LogDisplay(self.root)

    def run(self):
        self.root.mainloop()

# 使用範例
gui = MainGUI()
#gui.log_display.info("使用者操作：執行功能1")
#gui.log_display.error("5.menu bar - 主畫面上會有menu bar，menu bar上有3個sub menu，sub menu的文字分別是功能、設定、說明。每一個按下去都會跳出sub menu的選項，先幫我各預留3個sub menu的子選項。5.menu bar - 主畫面上會有menu bar，menu bar上有3個sub menu，sub menu的文字分別是功能、設定、說明。每一個按下去都會跳出sub menu的選項，先幫我各預留3個sub menu的子選項。5.menu bar - 主畫面上會有menu bar，menu bar上有3個sub menu，sub menu的文字分別是功能、設定、說明。每一個按下去都會跳出sub menu的選項，先幫我各預留3個sub menu的子選項。")
#gui.log_display.clear()
gui.run()
