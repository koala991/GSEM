# -*- coding: utf-8 -*-
"""
Version:
FireFox 57.0.4
geckodriver-v0.19.1-win64
selenium 3.9.0
"""

"""
未完成项
1.根据条件进一步筛选讲座
2.邮件通知
"""



from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
import time
from retrying import retry




@retry(stop_max_attempt_number = 2)
def LoginXmu(driver, user, passwd):
    # 模拟登陆讲座系统的函数
    driver.get('http://event.soe.xmu.edu.cn/')
    driver.find_element_by_name('UserName').send_keys(user)
    driver.find_element_by_name('Password').send_keys(passwd)
    driver.find_element_by_xpath('''//input[@class="click-logon button"]''').click()
    WebDriverWait(driver, 20).until(EC.text_to_be_present_in_element( \
                 (By.XPATH,"//div[@id='default-main']//legend"), "Seminars available for reservation:"))
    return




@retry(stop_max_attempt_number = 5)
def RefreshXmu(driver):
    driver.get("http://event.soe.xmu.edu.cn/LectureOrder.aspx")
    statu_load = WebDriverWait(driver, 20).until(EC.text_to_be_present_in_element( \
                              (By.XPATH,"//div[@id='default-main']//legend"), "Seminars available for reservation:"))
    return



def ReserveSeminar(deriver, s_id):
    # 根据id号选取讲座直至成功或失败
    # 因为还不知道选到讲座后的网页结构，所以使用暴力狂点的方式
    for i in range(20):
        driver.find_element_by_xpath \
        ("//a[@id='ctl00_MainContent_GridView1_ctl%s_btnreceive']"%s_id).click()
    # WebDriverWait(driver, 20).until(EC.text_to_be_present_in_element( \
    #                         (By.XPATH,"//"), "预约成功"))
    
    # 在获得页面结构后通过判定返回 True | False
    return False



def ScreenLec(driver, ids_lec, condition):
    # 输入可选讲座ids和条件，返回第一个符合条件的讲座的详情字典，若无符合条件则返回空字典
    
    return {}



def SendMail(mail_address, content):
    # 输入邮箱地址与发件内容，向邮箱发送邮件
    
    return



def GetSeminars(driver, condition = '', mail_address = '', n_need = 1, time_sep = 1):
    # 根据条件选取足够数量的的讲座
    t_load = 0
    n_get = 0
    while n_get < n_need:
        t_load += 1
        # 变量初始化
        statu_get = False
        lec_id = []
        lec_detail = {}
        # 获取总讲座数
        s_npage = driver.find_element_by_xpath("//div[@id='ctl00_MainContent_AspNetPager']").text
        n_lec = int(re.search('共([0-9]+)条记录', s_npage).groups()[0])
        for lec in range(n_lec):
            lec_statu = driver.find_element_by_xpath \
            ("//span[@id='ctl00_MainContent_GridView1_ctl0%d_Label1']"%(lec+2)).text
            if '预约已满' in lec_statu:
                pass
            elif '预约中' in lec_statu:
                lec_id.append('0%d'%(lec+2))
            else:
                pass
        
        if len(lec_id) > 0:
            # 将可预约的讲座输入判定函数，返回第一个符合条件的讲座的详情
            lec_detail = ScreenLec(driver, lec_id, condition)                
        
        if len(lec_detail) > 0:
            # 将符合条件的讲座id输入函数，对讲座进行选取，并返回选取状态
            statu_get = ReserveSeminar(driver, lec_detail['id'])
        n_get = n_get + statu_get        
        if statu_get:
            # 若抢到讲座，则邮件通知，并附上详情
            # 构造邮件
            mail_content = ''
            SendMail(mail_address, mail_content)
        time.sleep(time_sep)
        print("Load %d times!"%t_load)
        # 刷新页面
        RefreshXmu(driver)
        # 此处没有讲座时需要一个处理
    else:
        print("Got " + str(n_get) + " seminars!")


            
            
            




# stuid = input("请输入学号:")
stuid = "15420161152178"
# passwd = input("请输入密码：")
passwd = "koala19940331"



try:
    # options = webdriver.FirefoxOptions()
    # options.add_argument('--headless')
    # driver = webdriver.Firefox(options=options)
    driver = webdriver.Firefox()
    LoginXmu(driver, stuid, passwd)
    GetSeminars(driver)
    driver.close()
except Exception as err:
    print(err)
