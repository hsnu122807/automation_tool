import tkinter as tk
from tkinter import filedialog
from interface import LogDisplay
import os
from os.path import join,isdir


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
        self.show_edw_etl_radio_btn(self.radio_btn_frame)
        
        self.list_view_frame = tk.Frame(self.upper_left_frame, pady=10, padx=10)
        self.list_view_frame.pack(fill=tk.BOTH, expand=True)
        self.show_edw_list_box(self.list_view_frame)


        # 畫面中上
        self.upper_center_frame = tk.Frame(self.upper_frame, pady=10, padx=10)
        self.upper_center_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.common_info_frame = tk.Frame(self.upper_center_frame, pady=10, padx=10)
        self.common_info_frame.pack(fill=tk.BOTH, expand=True)
        self.show_edw_common_info(self.common_info_frame)
        
        self.step_frame = tk.Frame(self.upper_center_frame, pady=10, padx=10)
        self.step_frame.pack(fill=tk.BOTH, expand=True)
        

        # 畫面右上
        #self.upper_right_frame = tk.Frame(self.upper_frame, pady=10, padx=10)
        #self.upper_right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

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
            self.clear_window(self.list_view_frame)
            self.show_edw_list_box(self.list_view_frame)
        elif selected_value == 'ETL':
            self.clear_window(self.list_view_frame)
            self.show_etl_list_box(self.list_view_frame)


        
    def on_edw_list_box_select(self, evt):
        # Note here that Tkinter passes an event object to onselect()
        w = evt.widget
        index = int(w.curselection()[0])
        value = w.get(index)
        self.log_display.info("使用者操作：執行功能-" + value)
        self.clear_window(self.step_frame)
        if index == 0:
            self.show_edw_create_directory_step(self.step_frame)
        elif index == 1:
            self.show_edw_gen_dir_file_step(self.step_frame)
        elif index == 2:
            self.log_display.info('此功能開發中，尚未啟用')
            #self.show_edw_upload(self.step_frame)
        elif index == 3:
            self.log_display.info('此功能開發中，尚未啟用')
            #self.show_edw_cicd(self.step_frame)
        elif index == 4:
            self.log_display.info('此功能開發中，尚未啟用')
            #self.show_edw_check_doc(self.step_frame)
        
    def on_etl_list_box_select(self, evt):
        # Note here that Tkinter passes an event object to onselect()
        w = evt.widget
        index = int(w.curselection()[0])
        value = w.get(index)
        self.log_display.info("使用者操作：執行功能-" + value)
        if index == 0:
            self.log_display.info('此功能開發中，尚未啟用')
            #self.clear_window(self.step_frame)
            #self.show_etl_cicd(self.step_frame)
        elif index == 1:
            self.log_display.info('此功能開發中，尚未啟用')
            #self.clear_window(self.step_frame)
            #self.show_etl_check_doc(self.step_frame)
        
        
    def show_edw_etl_radio_btn(self, frame):
        self.radio_var = tk.StringVar()
        self.radio_var.set('EDW')
        self.radio_btn_edw = tk.Radiobutton(frame, text='EDW', variable=self.radio_var, value='EDW')    # 放入第一個單選按鈕
        self.radio_btn_edw.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.radio_btn_edw.select()
        self.radio_btn_etl = tk.Radiobutton(frame, text='ETL', variable=self.radio_var, value='ETL')   # 放入第二個單選按鈕
        self.radio_btn_etl.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.radio_var.trace("w", self.radio_selection_changed)
        
    def show_edw_list_box(self, frame):
        self.listbox = tk.Listbox(frame)# 放入列表選擇框  
        self.listbox.insert(1, '建立工作資料夾')          # 第一個選項
        self.listbox.insert(2, '轉換/產出dir檔')          # 第二個選項
        self.listbox.insert(3, '上傳程式與測試資料')      # 第三個選項
        self.listbox.insert(4, '測試完成進行版控')        # 第四個選項
        self.listbox.insert(5, '檢核上線文件')            # 第五個選項
        self.listbox.bind('<<ListboxSelect>>', self.on_edw_list_box_select)
        self.listbox.pack(fill=tk.BOTH, expand=True, pady=3, padx=3)
        

    def show_etl_list_box(self, frame):
        self.listbox = tk.Listbox(frame)# 放入列表選擇框  
        self.listbox.insert(1, '測試完成進行版控')    # 第一個選項
        self.listbox.insert(2, '檢核上線文件')        # 第二個選項
        self.listbox.bind('<<ListboxSelect>>', self.on_etl_list_box_select)
        self.listbox.pack(fill=tk.BOTH, expand=True, pady=3, padx=3)


    def get_work_file_path(self):
        # 選擇檔案後回傳檔案路徑與名稱
        file_path = filedialog.askdirectory()
        self.entry_work_file_path.delete(0,tk.END)
        self.entry_work_file_path.insert(0,file_path)


    def show_edw_common_info(self, frame):
        self.label_work_file_path = tk.Label(frame, text="工作區資料夾:")
        self.label_work_file_path.grid(row=0, column=0, padx=10, pady=5)
        self.entry_work_file_path = tk.Entry(frame)
        self.entry_work_file_path.insert(0,r"D:\自動化測試")
        self.entry_work_file_path.grid(row=0, column=1, padx=10, pady=5)
        self.btn_get_work_file_path = tk.Button(frame, text="選擇資料夾", command=self.get_work_file_path)
        self.btn_get_work_file_path.grid(row=0, column=2, padx=10, pady=5)
        self.label_icontect_no = tk.Label(frame, text="IContect單號:")
        self.label_icontect_no.grid(row=1, column=0, padx=10, pady=5)
        self.entry_icontect_no = tk.Entry(frame)
        self.entry_icontect_no.grid(row=1, column=1, padx=10, pady=5)
        self.label_describe = tk.Label(frame, text="需求簡述:")
        self.label_describe.grid(row=2, column=0, padx=10, pady=5)
        self.entry_describe = tk.Entry(frame)
        self.entry_describe.grid(row=2, column=1, padx=10, pady=5)
    
    
    def create_child_directory(self):
        work_file_path = self.entry_work_file_path.get()
        icontect_no = self.entry_icontect_no.get()
        describe = self.entry_describe.get()
        if not isdir(work_file_path):
            os.mkdir(work_file_path)
            self.log_display.info('建立資料夾成功，路徑: ' + work_file_path)
        self.path = join(work_file_path, icontect_no + '-' + describe)
        if not isdir(self.path):
            os.mkdir(self.path)
            self.log_display.info('建立資料夾成功，路徑: ' + self.path)
        child_dir_list = ['上線文件','上線程式','測試資料']
        for child_dir in child_dir_list:
            cur_dir = join(self.path,child_dir)
            if not isdir(cur_dir):
                os.mkdir(cur_dir)
                self.log_display.info('建立資料夾成功，路徑: ' + cur_dir)
            else:
                self.log_display.info('資料夾 ' + cur_dir + ' 已存在')
    
    def show_edw_create_directory_step(self, frame):
        self.btn_create_directory = tk.Button(frame, text="建立上線文件/上線程式/測試資料資料夾", command=self.create_child_directory)
        self.btn_create_directory.grid(row=0, column=0, padx=10, pady=5)
    def show_edw_gen_dir_file_step(self, frame):
        self.gen_dir_file_type_radio_btn_frame = tk.Frame(frame, pady=10, padx=10)
        self.gen_dir_file_type_radio_btn_frame.pack(fill=tk.BOTH, expand=True)
    
        self.radio_var_gen_dir_type = tk.StringVar()
        self.radio_var_gen_dir_type.set('有flg測試檔(建議)')
        self.radio_btn_have_flg_file = tk.Radiobutton(gen_dir_file_type_radio_btn_frame, text='有flg測試檔(建議)', variable=self.radio_var_gen_dir_type, value='有flg測試檔(建議)')    # 放入第一個單選按鈕
        self.radio_btn_have_flg_file.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.radio_btn_have_flg_file.select()
        self.radio_btn_no_flg_file = tk.Radiobutton(gen_dir_file_type_radio_btn_frame, text='無flg測試檔', variable=self.radio_var_gen_dir_type, value='無flg測試檔')   # 放入第二個單選按鈕
        self.radio_btn_no_flg_file.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.radio_var_gen_dir_type.trace("w", self.radio_selection_changed)
        
        self.gen_dir_file_info_frame = tk.Frame(frame, pady=10, padx=10)
        self.gen_dir_file_info_frame.pack(fill=tk.BOTH, expand=True)
        
        self.label_test_file_path = tk.Label(frame, text="測試檔案資料夾:")
        self.label_test_file_path.grid(row=0, column=0, padx=10, pady=5)
        self.entry_test_file_path = tk.Entry(frame)
        self.entry_test_file_path.insert(0,r"D:\自動化測試")
        self.entry_test_file_path.grid(row=0, column=1, padx=10, pady=5)
        self.btn_get_test_file_path = tk.Button(frame, text="選擇資料夾", command=self.get_test_file_path)
        self.btn_get_test_file_path.grid(row=0, column=2, padx=10, pady=5)
        self.btn_get_test_file_path = tk.Button(frame, text="產出dir檔", command=self.gen_dir_file)
        self.btn_get_test_file_path.grid(row=1, column=1, padx=10, pady=5)


    #def show_edw_upload(self, frame):
    #
    #def show_edw_cicd(self, frame):
    #
    #def show_edw_prepare(self, frame):