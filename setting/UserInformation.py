import getpass
import platform

class UserInformation:
    def __init__(self):
        self.username = None

    def get_current_username(self):
        if platform.system() == "Windows":
            self.username = getpass.getuser()
        elif platform.system() == "Linux":
            self.username = getpass.getuser()
        elif platform.system() == "Darwin":  # macOS
            self.username = getpass.getuser()
        else:
            self.username = "Unknown"
        return self.username

# 使用範例
# user_info = UserInformation()
# current_username = user_info.get_current_username()
# print(f"目前登入的AD帳號: {current_username}")
