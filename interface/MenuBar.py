import tkinter as tk
from tkinter import messagebox
from interface import AutoTesting

class MenuBar:
    def __init__(self, root):
        self.root = root
        self.root.title("自動化工具程式")

        # 建立選單列
        self.menu_bar = tk.Menu(self.root)

        # 建立功能選單
        self.function_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.function_menu.add_command(label="自動化測試", command=self.auto_testing)
        self.function_menu.add_command(label="改檔", command=self.change_dw_data)
        self.function_menu.add_command(label="功能3", command=self.function3_action)

        # 建立設定選單
        self.setting_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.setting_menu.add_command(label="設定1", command=self.setting1_action)
        self.setting_menu.add_command(label="設定2", command=self.setting2_action)
        self.setting_menu.add_command(label="設定3", command=self.setting3_action)

        # 建立說明選單
        self.help_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.help_menu.add_command(label="說明1", command=self.help1_action)
        self.help_menu.add_command(label="說明2", command=self.help2_action)
        self.help_menu.add_command(label="說明3", command=self.help3_action)

        # 將子選單加入選單列
        self.menu_bar.add_cascade(label="功能", menu=self.function_menu)
        self.menu_bar.add_cascade(label="設定", menu=self.setting_menu)
        self.menu_bar.add_cascade(label="說明", menu=self.help_menu)

        # 設定主畫面的選單列
        self.root.config(menu=self.menu_bar)

    def run(self):
        self.root.mainloop()

    def clear_window(self):
        # 清空主畫面，保留menu
        for widget in self.root.winfo_children():
            if not isinstance(widget, tk.Menu):
                widget.destroy()

    # 功能選單的動作
    def auto_testing(self):
        self.clear_window()
        AutoTesting.AutoTesting(self.root)

    def change_dw_data(self):
        self.clear_window()

    def function3_action(self):
        messagebox.showinfo("功能3", "執行功能3")

    # 設定選單的動作
    def setting1_action(self):
        messagebox.showinfo("設定1", "執行設定1")

    def setting2_action(self):
        messagebox.showinfo("設定2", "執行設定2")

    def setting3_action(self):
        messagebox.showinfo("設定3", "執行設定3")

    # 說明選單的動作
    def help1_action(self):
        messagebox.showinfo("說明1", "顯示說明1")

    def help2_action(self):
        messagebox.showinfo("說明2", "顯示說明2")

    def help3_action(self):
        messagebox.showinfo("說明3", "顯示說明3")


