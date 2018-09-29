import os
import time
from transform import get_easyList


    # 1 生成待录取精简列表
    # appli_lite_pinkun = [[15，True，10000，10002，12225],[],[]...] [编号，服从调剂，岗位编号1，2，3]
    # appli_lite_feipin
    #0.1版本只读取A+



#对于某一个岗位的录取方法
def exeEnroll(jobNum,pin,easy_list):
    if not pin:

        for appli in easy_list:
            #第一志愿查找
            if appli[2] == jobNum:
                return appli[0]
        for appli in easy_list:
            #第二志愿查找
            if appli[3] == jobNum:
                return appli[0]
        for appli in easy_list:
            #第三志愿查找
            if appli[4] == jobNum:
                return appli[0]
        for appli in easy_list:
            #检索服从调剂的同学
            if appli[1] == '是':
                return appli[0]
        return None
    if pin:

        for appli in easy_list:
            #第一志愿查找
            if appli[2] == jobNum:
                return appli[0]
        for appli in easy_list:
            #第二志愿查找
            if appli[3] == jobNum:
                return appli[0]
        for appli in easy_list:
            #第三志愿查找
            if appli[4] == jobNum:
                return appli[0]
        for appli in easy_list:
            #检索服从调剂的同学
            if appli[1] == '是':
                return appli[0]
        return None


def enroll_feipin(jobInfo):
    print('首先我们来进行非贫困生的录取')
    easy_list = get_easyList(False)
    print('获得待录取资格的非贫困生人数是 {0}'.format(len(easy_list)))
    enrolled = []
    count = 0
    for jobDict in jobInfo:
        n = jobDict.get('feipin')
        if n != 0 and n!= None:
            count = count + n
    if count:
        print('非贫困生录取名额总数是 {0}'.format(count))
        for jobDict in jobInfo: #对招聘信息中招收非贫困生的岗位进行遍历
            if jobDict.get('feipin') != 0 and jobDict.get('feipin') != None:
                n = jobDict['feipin']
                jobNum = jobDict.get('num')
                for _ in range(n): # 找几个人就循环多少次
                    bianhao = exeEnroll(jobNum,pin=False,easy_list=easy_list)
                    if bianhao != None: #录取了某个同学
                        print('  恭喜编号是{0}的同学(非贫困生)被录取岗位{1}'.format(bianhao,jobNum))
                        #操作录取的列表改变  改变jobinfo 增加录取的列表 减少带录取的easylist
                        enrollItem = {}
                        enrollItem[bianhao] = jobNum
                        enrolled.append(enrollItem)

                        jobDict['feipin'] -= 1

                        i = 0
                        for appli in easy_list:
                            if appli[0] == bianhao:
                                del easy_list[i]
                            i+=1
                    else:
                        jobDict['pinkun'] += jobDict['feipin']

                        msg = '  因为非贫困生没能招满名额，对于岗位{0}，将非贫困生剩余的{1}个名额增加到贫困生的录取中'.format(jobNum ,jobDict['feipin'])
                        print(msg)
                        break
        return enrolled,easy_list
        #返回的easylist是没有录取的非贫困生，

    else:
        print('没有非贫困生的名额')
        return None,None

def enroll_pinkun(jobInfo):
    print('------------------------------\n','接下来我们来进行贫困生的录取')
    easy_list = get_easyList(True)
    print('获得待录取资格的贫困生人数是 {0}'.format(len(easy_list)))
    enrolled = []
    count = 0
    for jobDict in jobInfo:
        n = jobDict.get('pinkun')
        if n != 0 and n != None:
            count = count + n
    if count:
        print('贫困生录取 名额 总数是 {0}'.format(count))
        for jobDict in jobInfo:  # 对招聘信息中招收贫困生的 岗位 进行遍历
            if jobDict.get('pinkun') != 0 and jobDict.get('pinkun') != None:
                n = jobDict['pinkun']
                jobNum = jobDict.get('num')
                for _ in range(n):  # 找几个人就循环多少次
                    bianhao = exeEnroll(jobNum, pin=True, easy_list=easy_list)
                    if bianhao != None:  # 录取了某个同学
                        print('  恭喜编号是{0}的同学(贫困生)被录取岗位{1}'.format(bianhao, jobNum))
                        # 操作录取的列表改变  改变jobinfo 增加录取的列表 减少带录取的easylist
                        enrollItem = {}
                        enrollItem[bianhao] = jobNum
                        enrolled.append(enrollItem)

                        jobDict['pinkun'] -= 1

                        i = 0
                        for appli in easy_list:
                            if appli[0] == bianhao:
                                del easy_list[i]
                            i += 1
                    else:
                        msg = '  对于岗位{0}，贫困生剩余{1}个名额'.format(jobNum, jobDict['pinkun'])
                        print(msg)
                        break
        return enrolled, easy_list
        # 返回的easylist是没有录取的非贫困生，

    else:
        print('没有贫困生的名额')
        return None,None








