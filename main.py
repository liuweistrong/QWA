import time
from PATH import *
from enroll import enroll_feipin,enroll_pinkun
from load import *
from writeback import *

if __name__ == '__main__':
    print('Welcome to Tangzeling Working Automation Tool')
    print('现在的时间是： ',time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),'\n','--------------------------------')

    jobInfo = get_library(2)
    enrolled_feipin,not_enrolled_feipin = enroll_feipin(jobInfo)
    enrolled_pinkun,not_enrolled_pinkun = enroll_pinkun(jobInfo)

    write_result(enrolled_feipin,enrolled_pinkun,jobInfo)


    print('录取结果在result.xlsx文件中 打开路径：',RESULT_PATH)

    input('这个不看了关了就行了')