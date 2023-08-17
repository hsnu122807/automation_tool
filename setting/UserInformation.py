import getpass
import subprocess
from tkinter import messagebox

class UserInformation:
    def __init__(self):
        self.user_cathay_no = getpass.getuser()
        user_dict_file = 'resource/user_list.txt'
        user_dict = {}
        f = open(user_dict_file,'r',encoding='utf-8')
        for line in f.readlines():
            cathay_no,user_bank_no,user_name = line.split()
            user_dict[cathay_no] = [user_bank_no,user_name]
            # print(cathay_no,user_bank_no,user_name)
        f.close()
        
        if self.user_cathay_no in user_dict:
            self.user_bank_no,self.user_name = user_dict[self.user_cathay_no]
        else:
            subprocess.Popen('explorer ' + user_dict_file)
            messagebox.showinfo("錯誤", '使用者員編資訊未登錄，請修改"' + user_dict_file + '"內容')
            self.user_bank_no,self.user_name = '',''


    def get_user_cathay_no(self):
        return self.user_cathay_no

    def get_user_bank_no(self):
        return self.user_bank_no

    def get_user_full_name(self):
        return self.user_name

    def get_user_first_name(self):
        return self.user_name[1:]

    def get_user_last_name(self):
        return self.user_name[0]
