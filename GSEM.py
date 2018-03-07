# -*- coding: utf-8 -*-
# Version 0.0.2 添加延迟运行
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


import re, time, os, sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from retrying import retry
import smtplib
from email.mime.text import MIMEText
from email.header import Header
 



def LoginXmu(driver, user, passwd):
    # 模拟登陆讲座系统的函数
    try:
        driver.get('http://event.soe.xmu.edu.cn/')
        driver.find_element_by_name('UserName').send_keys(user)
        driver.find_element_by_name('Password').send_keys(passwd)
        driver.find_element_by_xpath('''//input[@class="click-logon button"]''').click()
        WebDriverWait(driver, 20).until(EC.text_to_be_present_in_element( \
                     (By.XPATH,"//div[@id='default-main']//legend"), "Seminars available for reservation:"))
    except Exception:
        print("登陆失败，请检查网络状态或账户密码")
    return




@retry(stop_max_attempt_number = 5)
def RefreshXmu(driver):
    driver.get("http://event.soe.xmu.edu.cn/LectureOrder.aspx")
    WebDriverWait(driver, 20).until(EC.text_to_be_present_in_element( \
                 (By.XPATH,"//div[@id='default-main']//legend"), "Seminars available for reservation:"))
    return



def ReserveSeminar(deriver, s_id):
    # 根据id号选取讲座直至成功或失败
    get = False
    t = 0
    while (not get | t <20):
        driver.find_element_by_xpath \
        ("//a[@id='ctl00_MainContent_GridView1_ctl%s_btnreceive']"%s_id).click()
        try:
            WebDriverWait(driver, 5).until(EC.alert_is_present())
        except Exception:
            t += 1
            pass
        else:
            driver.switch_to_alert().accept()
            get = True
            break
    return get



def ScreenLec(driver, ids_lec, condition):
    # 输入可选讲座ids和条件，返回第一个符合条件的讲座的详情字典，若无符合条件则返回空字典
    # 条件目前仅一个最小退订时间
    for l in ids_lec:
        l_time = driver.find_element_by_xpath("//span[@id='ctl00_MainContent_GridView1_ctl%s_orderendtime']"%l).text
        l_time = time.mktime(time.strptime(l_time,"%Y-%m-%d %H:%M:%S"))
        if l_time - time.time() < condition["regret"] * 3600:
            continue
        # 其他条件，开发时直接if接下去就好
        
        # 若运行至此，则满足上列条件，开始获取详情
        detail = {}
        detail["name"] = driver.find_element_by_xpath(
                "//table[@id='ctl00_MainContent_GridView1']/tbody[1]/tr[%d]/td[2]"%int(l)).text
        detail["re_time"] = driver.find_element_by_xpath(
                "//span[@id='ctl00_MainContent_GridView1_ctl%s_orderendtime']"%l).text
        detail["id"] = l
        return detail
    return {}



def SendMail(mail_address, sender, passwd, content):
    # 输入邮箱地址与发件内容，向邮箱发送邮件
    # 目前仅支持从qq邮箱向外发送
    mail_host="smtp.qq.com"  #设置服务器
    mail_user= sender    #用户名
    mail_pass= passwd   #口令 
    receivers = mail_address  # 接收邮件，可设置为你的QQ邮箱或者其他邮箱
     
    message = MIMEText(content, 'plain', 'utf-8')
    message['From'] = Header(sender, 'utf-8')
    message['To'] =  Header(mail_address, 'utf-8')
    subject = "讲座成功选取提醒"
    message['Subject'] = Header(subject, 'utf-8')
    try:
        smtpObj = smtplib.SMTP_SSL(mail_host, 465) 
        smtpObj.login(mail_user,mail_pass)  
        smtpObj.sendmail(sender, receivers, message.as_string())
        print("邮件发送成功")
    except smtplib.SMTPException:
        print("Error: 无法发送邮件")
        return



def GetSeminars(driver, sendmail = False, condition = {}, mail_address = '',
                sender = '', passwd_mail = '', n_need = 1, time_sep = 1):
    # 根据条件选取足够数量的的讲座
    t_load = 0
    n_get = 0
    time_s = time.time()
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
                try:
                    driver.find_element_by_xpath("//a[@id='ctl00_MainContent_GridView1_ctl%d_btnreceive']"%(lec+2))
                except Exception as err:
                    # print(err)
                    pass
                else:
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
        if statu_get & sendmail:
            # 若抢到讲座，则邮件通知，并附上详情
            # 构造邮件
            mail_content = "已成功预约讲座<" + lec_detail["name"] + ">，最后退订时间为" + \
            lec_detail["re_time"] + "请及时查看！"
            SendMail(mail_address, sender, passwd_mail, mail_content)
        time.sleep(time_sep)
        time_n = time.time()
        if (t_load % int(180 / time_sep)) == 1:
            print("已运行 %d 秒， 尝试预约 %d 次， 已预约 %d 场讲座。"%(int(time_n - time_s), t_load, n_get) )
        # 刷新页面
        RefreshXmu(driver)
        # 此处没有讲座时需要一个处理
    else:
        print("完成任务")


            
            
            
###############################################
        
op = {}
op["condition"] = {}
set_load = input("读取配置(Y/N)?\n>>")
if set_load.upper() == "Y":
    f = open("options.ini", "r")
    op = eval(f.read())
    f.close()
    print("Load options", str(op))
else:
    op["stuid"] = input("请输入学号:\n>>")
    op["passwd"] = input("请输入密码:\n>>")
    t_regret = input("最小取消时间(默认12h):\n>>")
    op["condition"]['regret'] = 12 if t_regret == "" else float(t_regret)
    op["set_send"] = input("在选取讲座后发送邮件提醒(Y/N)? \n(需提供qq邮箱作为发信邮箱并提供账号密码，另还需接收邮箱。不推荐)\n>>")
    if op["set_send"].upper() == "Y":
        op["set_send"] = True
        op["sender"] = input("发信邮箱地址:\n>>")
        op["passwd_mail"] = input("发件邮箱密码:\n>>")
        op["mail_address"] = input("收信邮箱地址:\n>>")
    else:
        op["set_send"] = False
        op["sender"] = ""
        op["passwd_mail"] = ""
        op["mail_address"] = ""
    op["need"] = int(input("共需选取讲座数:\n>>"))
    op["sep"] = float(input("刷新间隔(建议为10):\n>>"))
    set_save = input("保存配置(Y/N)?\n>>")
    if set_save.upper() == "Y":
        f = open("options.ini", "wb")
        f.write(str(op).encode())
        f.close()
op["headless"] = True if input("隐藏浏览器(Y/N)?\n>>").upper() == "Y" else False
t_delay = input("延迟运行时间(h):\n(立即运行请回车)\n>>")
t_delay = 0 if t_delay == "" else float(t_delay)
print("延迟%d秒后运行"%(t_delay * 3600))
time.sleep(t_delay * 3600)
try:
    if op["headless"]:
        op_driver = webdriver.FirefoxOptions()
        op_driver.add_argument('--headless')
        driver = webdriver.Firefox(options = op_driver)
    else:
        driver = webdriver.Firefox()
    LoginXmu(driver, op['stuid'], op['passwd'])
    GetSeminars(driver, sendmail = op["set_send"], condition = op["condition"],
                mail_address = op["mail_address"], sender = op["sender"], 
                passwd_mail = op["passwd_mail"], n_need = op["need"],
                time_sep = op["sep"])
    driver.close()
except Exception as err:
    print("error: ", err)
print("预约结束")

