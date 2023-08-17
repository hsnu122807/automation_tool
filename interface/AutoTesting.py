import tkinter as tk
from tkinter import filedialog
from interface import LogDisplay
from setting import UserInformation
from setting import LoadResource
import os
from os import listdir
from os.path import isfile, isdir, join
import re
import paramiko

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
        if len(w.curselection()) > 0:
            index = int(w.curselection()[0])
            value = w.get(index)
            self.log_display.info("使用者操作：執行功能-" + value)
            self.clear_window(self.step_frame)
            if index == 0:
                self.show_edw_create_directory_step(self.step_frame)
            elif index == 1:
                self.show_edw_gen_dir_file_step(self.step_frame)
            elif index == 2:
                self.show_edw_upload_step(self.step_frame)
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


    def get_test_file_path(self):
        # 選擇資料夾後回傳檔案路徑與名稱
        file_path = filedialog.askdirectory()
        self.entry_test_file_path.delete(0,tk.END)
        self.entry_test_file_path.insert(0,file_path)

    def get_program_file_path(self):
        # 選擇資料夾後回傳檔案路徑與名稱
        file_path = filedialog.askdirectory()
        self.entry_program_file_path.delete(0,tk.END)
        self.entry_program_file_path.insert(0,file_path)


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
        child_dir_list = ['上線文件','上線程式','測試資料',r'測試資料\APP']
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
        self.str_have_flg_file = '有flg測試檔(建議)'
        self.str_no_flg_file = '無flg測試檔'
        self.radio_var_gen_dir_type = tk.StringVar()
        self.radio_var_gen_dir_type.set(self.str_have_flg_file)
        self.radio_btn_have_flg_file = tk.Radiobutton(frame, text=self.str_have_flg_file, variable=self.radio_var_gen_dir_type, value=self.str_have_flg_file)    # 放入第一個單選按鈕
        self.radio_btn_have_flg_file.grid(row=0, column=0, padx=10, pady=5)
        self.radio_btn_have_flg_file.select()
        self.radio_btn_no_flg_file = tk.Radiobutton(frame, text=self.str_no_flg_file, variable=self.radio_var_gen_dir_type, value=self.str_no_flg_file)   # 放入第二個單選按鈕
        self.radio_btn_no_flg_file.grid(row=0, column=1, padx=10, pady=5)
        self.radio_var_gen_dir_type.trace("w", self.radio_selection_changed)
        
        self.label_test_file_path = tk.Label(frame, text="測試檔案資料夾:")
        self.label_test_file_path.grid(row=1, column=0, padx=10, pady=5)
        work_file_path = self.entry_work_file_path.get()
        icontect_no = self.entry_icontect_no.get()
        describe = self.entry_describe.get()
        self.entry_test_file_path = tk.Entry(frame)
        self.entry_test_file_path.insert(0,join(work_file_path, icontect_no + '-' + describe, '測試資料'))
        self.entry_test_file_path.grid(row=1, column=1, padx=10, pady=5)
        self.btn_get_test_file_path = tk.Button(frame, text="選擇資料夾", command=self.get_test_file_path)
        self.btn_get_test_file_path.grid(row=1, column=2, padx=10, pady=5)
        self.btn_gen_dir_file = tk.Button(frame, text="產出dir檔", command=self.gen_dir_file)
        self.btn_gen_dir_file.grid(row=2, column=1, padx=10, pady=5)


    def gen_dir_file(self):
        path = self.entry_test_file_path.get()
        gen_dir_type = self.radio_var_gen_dir_type.get()
        
        try:
            # 取得所有檔案與子目錄名稱
            files = listdir(path)

            flg_file_name_list = []
            txt_file_name_list = []
            if gen_dir_type == self.str_have_flg_file:
                # 把flg改檔名為dir
                self.log_display.info('有flg檔案，將flg檔改名為dir檔')
                for f in files:
                    fullpath = join(path, f)
                    if isfile(fullpath) and len(re.findall('[0-9]{8}.txt$',f)) > 0:
                        txt_file_name_list.append(f)
                        self.log_display.info('找到txt檔案: ' + f)
                    elif isfile(fullpath) and len(re.findall('.flg$',f))>0:
                        flg_file_name_list.append(f)
                        self.log_display.info('找到flg檔案: ' + f)
                if len(flg_file_name_list) == 0:
                    self.log_display.error('路徑下未找到flg檔案: ' + path)
                else:
                    for flg_file_name in flg_file_name_list:
                        file = flg_file_name[:-4]
                        for txt_file_name in txt_file_name_list:
                            if file in txt_file_name and len(file) + 12 == len(txt_file_name):
                                date = txt_file_name[-12:-4]
                                current_full_path = join(path, flg_file_name)
                                dir_file_name = 'dir.'+file.upper()+date
                                new_full_path = join(path, dir_file_name)
                                os.rename(current_full_path,new_full_path)
                                self.log_display.info('將flg檔案 ' + flg_file_name + ' 改名為 ' + dir_file_name)
                                break
            elif gen_dir_type == self.str_no_flg_file:
                # 直接根據txt產出dir
                self.log_display.info('無flg檔案，根據txt檔生成dir檔')
                self.log_display.hint('若為首次導檔仍建議請前端提供flg，確保flg檔案內容正確性!')
                for f in files:
                    fullpath = join(path, f)
                    if isfile(fullpath) and len(re.findall('[0-9]{8}.txt$',f)) > 0:
                        txt_file_name_list.append(f)
                        self.log_display.info('找到txt檔案: ' + f)
                for txt_file_name in txt_file_name_list:
                    file = txt_file_name[:-4]
                    dir_file_name = 'dir.'+file.upper()
                    f = open(join(path, dir_file_name), 'w')
                    f.write(txt_file_name + '\n')
                    f.close()
                    self.log_display.info('產出dir檔案: ' + dir_file_name)
        except IOError:
            self.log_display.error('此路徑不存在: ' + path)
        except Exception as e:
            self.log_display.error('未預期的錯誤，請通知自動化工具維護者')
            self.log_display.error(str(e))


    def show_edw_upload_step(self, frame):
        work_file_path = self.entry_work_file_path.get()
        icontect_no = self.entry_icontect_no.get()
        describe = self.entry_describe.get()
        
        user_info = UserInformation.UserInformation()
        bank_no = user_info.get_user_bank_no()
        
        self.label_r6_acc = tk.Label(frame, text="UT R6帳號:")
        self.label_r6_acc.grid(row=0, column=0, padx=10, pady=5)
        self.entry_r6_acc = tk.Entry(frame)
        self.entry_r6_acc.insert(0, 'dod'+bank_no[2:])
        self.entry_r6_acc.grid(row=0, column=1, padx=10, pady=5)
        
        self.label_r6_pwd = tk.Label(frame, text="UT R6密碼:")
        self.label_r6_pwd.grid(row=1, column=0, padx=10, pady=5)
        self.entry_r6_pwd = tk.Entry(frame,show='*')
        self.entry_r6_pwd.grid(row=1, column=1, padx=10, pady=5)
        
        
        self.label_program_file_path = tk.Label(frame, text="程式資料夾:")
        self.label_program_file_path.grid(row=2, column=0, padx=10, pady=5)
        self.entry_program_file_path = tk.Entry(frame)
        self.entry_program_file_path.insert(0,join(work_file_path, icontect_no + '-' + describe, r'上線程式\APP'))
        self.entry_program_file_path.grid(row=2, column=1, padx=10, pady=5)
        self.btn_get_program_file_path = tk.Button(frame, text="選擇資料夾", command=self.get_program_file_path)
        self.btn_get_program_file_path.grid(row=2, column=2, padx=10, pady=5)
        self.btn_upload_program = tk.Button(frame, text="上傳 ETL程式 至UT環境", command=self.upload_program)
        self.btn_upload_program.grid(row=2, column=3, padx=10, pady=5)
        
        self.label_test_file_path = tk.Label(frame, text="測試檔案資料夾:")
        self.label_test_file_path.grid(row=3, column=0, padx=10, pady=5)
        self.entry_test_file_path = tk.Entry(frame)
        self.entry_test_file_path.insert(0,join(work_file_path, icontect_no + '-' + describe, '測試資料'))
        self.entry_test_file_path.grid(row=3, column=1, padx=10, pady=5)
        self.btn_get_test_file_path = tk.Button(frame, text="選擇資料夾", command=self.get_test_file_path)
        self.btn_get_test_file_path.grid(row=3, column=2, padx=10, pady=5)
        self.btn_upload_test_data = tk.Button(frame, text="上傳 測試資料 至UT環境", command=self.upload_test_data)
        self.btn_upload_test_data.grid(row=3, column=3, padx=10, pady=5)
        
        self.btn_upload_program_and_test_data = tk.Button(frame, text="上傳 ETL程式&測試資料 至UT環境", command=self.upload_program_and_test_data)
        self.btn_upload_program_and_test_data.grid(row=4, column=1, padx=10, pady=5)

    def upload_program(self):
        
        # 設定SFTP伺服器的連線資訊
        res = LoadResource.LoadResource()
        ut_r6_info = res.get_ut_r6_info()
        sftp_hostname = ut_r6_info['IP']
        sftp_port = ut_r6_info['SSH_PORT']
        sftp_path = ut_r6_info['SFTP_PATH']
        sftp_username = self.entry_r6_acc.get()
        sftp_password = self.entry_r6_pwd.get()
        if len(sftp_username) == 0 or len(sftp_password) == 0:
            self.log_display.error('未輸入帳號或密碼')
        else:
            # 設定本地檔案路徑及檔案名稱
            local_folder_path = self.entry_program_file_path.get()
            
            if local_folder_path[-4:] == r'\APP':
                # 設定遠端AIX主機的檔案路徑及檔案名稱
                remote_folder_path  = sftp_path
                try:
                    # 建立SSH客戶端連線
                    ssh_client = paramiko.SSHClient()
                    ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
                    ssh_client.connect(hostname=sftp_hostname, port=sftp_port, username=sftp_username, password=sftp_password)
                    # 建立SFTP客戶端連線
                    sftp_client = ssh_client.open_sftp()
                    # 執行資料夾上傳
                    self.upload_folder(ssh_client, sftp_client, local_folder_path, remote_folder_path)
                    self.log_display.info('ETL程式上傳完成')
                except Exception as e:
                    if str(e) == 'Authentication failed.':
                        self.log_display.error('帳號或密碼輸入錯誤')
                    else:
                        self.log_display.error('未預期的錯誤，請通知自動化工具維護者')
                        self.log_display.error(str(e))
                        print(e)
                finally:
                    # 關閉SFTP客戶端連線
                    sftp_client.close()
                    # 關閉SSH客戶端連線
                    ssh_client.close()
            else:
                self.log_display.error('程式資料夾命名須為APP')
                self.log_display.hint('此功能會將程式資料夾下的資料夾與ETL檔案SFTP至R6 UT環境的APP目錄下，請確保所選資料夾目錄架構的正確性!')
        

    def upload_test_data(self):
        # 設定SFTP伺服器的連線資訊
        res = LoadResource.LoadResource()
        ut_r6_info = res.get_ut_r6_info()
        sftp_hostname = ut_r6_info['IP']
        sftp_port = ut_r6_info['SSH_PORT']
        sftp_path = ut_r6_info['SFTP_PATH']
        sftp_username = self.entry_r6_acc.get()
        sftp_password = self.entry_r6_pwd.get()
        if len(sftp_username) == 0 or len(sftp_password) == 0:
            self.log_display.error('未輸入帳號或密碼')
        else:
            # 設定本地檔案路徑及檔案名稱
            local_folder_path = self.entry_program_file_path.get()
            
            if local_folder_path[-4:] == r'\APP':
                # 設定遠端AIX主機的檔案路徑及檔案名稱
                remote_folder_path  = sftp_path
                try:
                    # 建立SSH客戶端連線
                    ssh_client = paramiko.SSHClient()
                    ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
                    ssh_client.connect(hostname=sftp_hostname, port=sftp_port, username=sftp_username, password=sftp_password)
                    # 建立SFTP客戶端連線
                    sftp_client = ssh_client.open_sftp()
                    # 執行資料夾上傳
                    self.upload_folder(ssh_client, sftp_client, local_folder_path, remote_folder_path)
                    self.log_display.info('測試資料上傳完成')
                except Exception as e:
                    if str(e) == 'Authentication failed.':
                        self.log_display.error('帳號或密碼輸入錯誤')
                    else:
                        self.log_display.error('未預期的錯誤，請通知自動化工具維護者')
                        self.log_display.error(str(e))
                        print(e)
                finally:
                    # 關閉SFTP客戶端連線
                    sftp_client.close()
                    # 關閉SSH客戶端連線
                    ssh_client.close()
            else:
                self.log_display.error('程式資料夾命名須為APP')
                self.log_display.hint('此功能會將程式資料夾下的資料夾與檔案SFTP至R6 UT環境的APP目錄下，請確保所選資料夾目錄架構的正確性!')

    def upload_program_and_test_data(self):
        self.upload_program()
        self.upload_test_data()
        
    
    def create_remote_directory(self, sftp_client, remote_path):
        try:
            sftp_client.mkdir(remote_path)
            self.log_display.info('於UT環境建立資料夾: ' + remote_item)
        except OSError:
            pass

    def set_remote_permissions(self, ssh_client, remote_path):
        # 執行遠端的chmod 777指令
        command = f"chmod 777 {remote_path}"
        stdin, stdout, stderr = ssh_client.exec_command(command)
        # 可選：若需要檢查指令是否執行成功，可以印出stdout和stderr的內容
    #     print(stdout.read().decode())
    #     print(stderr.read().decode())
        
    def upload_folder(self, ssh_client, sftp_client, local_path, remote_path):
        self.create_remote_directory(sftp_client, remote_path)
        self.set_remote_permissions(ssh_client, remote_path)  # 設定目錄權限為777
        for item in os.listdir(local_path):
            local_item = os.path.join(local_path, item)
            local_item = local_item.replace('\\', '/')
            remote_item = os.path.join(remote_path, item)
            remote_item = remote_item.replace('\\', '/')
            if os.path.isfile(local_item):
                if local_item[-3:] == '.pl' or local_item[-4:] == '.sql':
                    sftp_client.put(local_item, remote_item)
                    self.set_remote_permissions(ssh_client, remote_item)  # 設定檔案權限為777
                    self.log_display.info('上傳 ' + local_item + ' 至 ' + remote_item + ' ，並將權限改為777')
                else:
                    self.log_display.warn(local_item + ' 並非.sql或.pl檔案，沒有上傳至R6 UT環境')
            elif os.path.isdir(local_item):
                self.upload_folder(ssh_client, sftp_client, local_item, remote_item)
    #def show_edw_upload(self, frame):
    #
    #def show_edw_cicd(self, frame):
    #
    #def show_edw_prepare(self, frame):