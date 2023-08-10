import tkinter as tk
from interface import LogDisplay

class AutoTesting:
    def __init__(self, root):
        self.root = root

        # 設定 AutoTesting 畫面的元件
        # 畫面上半部
        self.upper_frame = tk.Frame(self.root)
        self.upper_frame.pack(fill=tk.BOTH, expand=True)

        # 畫面左上
        self.upper_left_frame = tk.Frame(self.upper_frame, pady=10, padx=10)
        self.upper_left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.radio_btn_frame = tk.Frame(self.upper_left_frame, pady=10, padx=10)
        self.radio_btn_frame.pack(fill=tk.BOTH, expand=True)
        
        self.radio_var = tk.StringVar()
        self.radio_var.set('EDW')
        self.radio_btn_edw = tk.Radiobutton(self.radio_btn_frame, text='EDW', variable=self.radio_var, value='EDW')    # 放入第一個單選按鈕
        self.radio_btn_edw.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.radio_btn_edw.select()
        self.radio_btn_etl = tk.Radiobutton(self.radio_btn_frame, text='ETL', variable=self.radio_var, value='ETL')   # 放入第二個單選按鈕
        self.radio_btn_etl.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.radio_var.trace("w", self.radio_selection_changed)
        
        self.show_edw_list_box(self.upper_left_frame)


        # 畫面中上
        self.upper_center_frame = tk.Frame(self.upper_frame, pady=10, padx=10)
        self.upper_center_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        

        self.upper_right_frame = tk.Frame(self.upper_frame, pady=10, padx=10)
        self.upper_right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 畫面下半部
        self.lower_frame = tk.Frame(self.root, pady=10, padx=10) 
        self.log_display = LogDisplay.LogDisplay(self.lower_frame)
        self.lower_frame.pack(fill=tk.BOTH, expand=True)
        
        
    def clear_window(self, frame):
        for widget in frame.winfo_children():
            widget.destroy()


    def radio_selection_changed(self, *args):
        selected_value = self.radio_var.get()
        if selected_value == 'EDW':
            print("選擇了選項EDW")
        elif selected_value == 'ETL':
            print("選擇了選項ETL")


    def on_radio_btn_select(self, evt):
        # Note here that Tkinter passes an event object to onselect()
        w = evt.widget
        index = int(w.curselection()[0])
        value = w.get(index)
        self.log_display.info("使用者操作：執行功能-" + value)
        
    def on_edw_list_box_select(self, evt):
        # Note here that Tkinter passes an event object to onselect()
        w = evt.widget
        index = int(w.curselection()[0])
        value = w.get(index)
        self.log_display.info("使用者操作：執行功能-" + value)
        
        
    def show_edw_list_box(self, frame):
        self.listbox = tk.Listbox(frame)# 放入列表選擇框  
        self.listbox.insert(1, '上傳程式與測試資料')    # 第一個選項
        self.listbox.insert(2, '測試完成進行版控')      # 第二個選項
        self.listbox.insert(3, '檢核測試結果')          # 第三個選項
        self.listbox.bind('<<ListboxSelect>>', self.on_edw_list_box_select)
        self.listbox.pack(fill=tk.BOTH, expand=True, pady=3, padx=3)
        

    def show_etl_list_box(self, frame):
        self.listbox = tk.Listbox(frame)# 放入列表選擇框  
        self.listbox.insert(1, '上傳程式與測試資料')    # 第一個選項
        self.listbox.insert(2, '測試完成進行版控')      # 第二個選項
        self.listbox.insert(3, '檢核測試結果')          # 第三個選項
        self.listbox.bind('<<ListboxSelect>>', self.on_etl_list_box_select)
        self.listbox.pack(fill=tk.BOTH, expand=True, pady=3, padx=3)
