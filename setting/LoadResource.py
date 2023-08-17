

class LoadResource:
    def __init__(self):
        pass
        
    def get_ut_r6_info(self):
        result = {}
        f = open('resource/r6_info.txt','r',encoding='utf-8')
        for line in f.readlines():
            if line[:2] == 'UT':
                result_line = line
        f.close()
        result_list = result_line.split()
        result['IP'] = result_list[1]
        result['SSH_PORT'] = result_list[2]
        result['SFTP_PATH'] = result_list[3]
        return result
        