import tkinter as tk
import time
 


class LogDisplay:
    def __init__(self, root):
        self.root = root

        # 設定 Log 訊息的 Text 物件，位於畫面下半部並佔整個畫面的三分之一
        self.log_text = tk.Text(self.root, pady=3, padx=3)
        

        # 增加垂直滾動卷軸
        self.scroll_y = tk.Scrollbar(self.log_text)
        self.scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=self.scroll_y.set)
        
        self.log_text.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)

    def info(self, message):
        curr_time = time.strftime("%Y-%m-%d %H:%M:%S ", time.localtime())
        self.log_text.insert(tk.END, curr_time + '[訊息]' + message + "\n")
        self.log_text.see(tk.END)

    def error(self, message):
        curr_time = time.strftime("%Y-%m-%d %H:%M:%S ", time.localtime())
        self.log_text.tag_config("error_tag", foreground="red")
        self.log_text.insert(tk.END, curr_time + '[錯誤]' + message + "\n", "error_tag")
        self.log_text.see(tk.END)

    def clear(self):
        self.log_text.delete(1.0, tk.END)
