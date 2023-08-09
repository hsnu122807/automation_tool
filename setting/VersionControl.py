import subprocess

class VersionControl:
    def __init__(self, git_repo):
        self.git_repo = git_repo

    def check_latest_version(self):
        try:
            subprocess.check_output(["git", "fetch"], cwd=self.git_repo)
            result = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=self.git_repo)
            current_commit = result.decode().strip()

            subprocess.check_output(["git", "reset", "--hard", "origin/master"], cwd=self.git_repo)
            subprocess.check_output(["git", "pull"], cwd=self.git_repo)
            result = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=self.git_repo)
            latest_commit = result.decode().strip()

            if current_commit != latest_commit:
                return False
            return True

        except subprocess.CalledProcessError:
            return False

# 使用範例
git_repo_path = "D:\Python\Projects\automation_tool"
version_control = VersionControl(git_repo_path)

if version_control.check_latest_version():
    print("程式是最新版本")
else:
    print("程式不是最新版本，需要更新")