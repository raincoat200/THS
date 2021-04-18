#!coding:utf-8
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
import pandas as pd
import datetime
from time import sleep

def nb_date(date):
    data = browser.find_element_by_class_name('table_data')
    data3 = data.find_element_by_class_name('data_tbody')
    rqs = data3.find_elements_by_xpath("table[@class='top_thead']/tbody/tr/th")
    global nb_num
    nb_num = 0  # 年度序号
    for rq in rqs:
        if rq.text == date:
            break
        elif nb_num > 10:
            nb_num=None
            newdata = browser.find_element_by_class_name('table_data').find_elements_by_class_name('td_w')[0].text
            print("该期年报未发布,最新年报日期为：", newdata)
            break
        else:
            nb_num = nb_num + 1

def su_find(km,num1):
    global nb_num
    #num1为年报序号,0为最新一期,依此类推
    if nb_num is not None:
        data = browser.find_element_by_class_name('table_data')
        data1 = data.find_element_by_class_name('left_thead')
        data3 = data.find_element_by_class_name('data_tbody')
        kms=data1.find_elements_by_xpath("table[@class='tbody']/tbody/tr")
        km_num=1
        for i in kms:
            if i.text==km:
                #km_num=i.get_attribute('data-index') #科目序号
                break
            km_num = km_num + 1
        nr = data3.find_elements_by_xpath('table[@class="tbody"]/tbody/tr[' + str(km_num) + ']/td')[num1]
        nr = nr.get_attribute('textContent').replace(" ", "")
        return bs(nr)
    else:
        return None

def bs(var):
    var=var.replace(" ","").replace("\n","").replace("\t","").replace("--", "")
    if var.find('万亿')>0:
        bsnum=10000
        var=var.replace("万亿","")
    elif var.find('亿')>0:
        bsnum=1
        var = var.replace("亿", "")
    elif var.find('千万')>0:
        bsnum=1000/10000
        var = var.replace("千万", "")
    elif var.find('百万')>0:
        bsnum=100/10000
        var = var.replace("百万", "")
    elif var.find('万')>0:
        bsnum=1/10000
        var = var.replace("万", "")
    elif var.find('%')>0:
        bsnum=1/100
        var = var.replace("%", "")
    elif var.find('-')>0:
        bsnum=1
        var = var.replace("--", "")
    if var!="":
        return round(float(var)*bsnum,2)
    else:
        return ""

def zb_main(str):
    temp=browser.find_element_by_xpath("//p[@class='name' and text()='"+str+"']//parent::div")
    ActionChains(browser).move_to_element(temp)
    var1=temp.find_elements_by_tag_name("span")[nb_num].get_attribute('textContent')
    var2=temp.find_elements_by_tag_name("span")[nb_num+4].get_attribute('textContent')
    #同比增长率
    try:
        var3='{:.2%}'.format((bs(var1)-bs(var2))/bs(var2))
    except:
        var3=""
    return bs(var1),var3

chrome_path = '..\\chromedriver.exe'
chrome_path = 'chromedriver.exe'
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--headless') #增加无界面选项
chrome_options.add_argument('--disable-gpu')
browser = webdriver.Chrome(executable_path=chrome_path, options=chrome_options)
browser = webdriver.Chrome(options=chrome_options)
windows = browser.window_handles

pdata=pd.DataFrame(columns=('证券代码','公司名称','时间','总资产','资产增长率','保证金规模','总负债','总股本',
                            '手续费及佣金收入','投资收益','利息净收入','营业收入','营收增长率','净利润','净利润增长率',
                            '归母净利润','归母净利润增长率','净利润率','ROE','营业支出','管理费用','信用减值损失',
                            '净资本','净资本增长率','净资产','风险覆盖率'))
code_list=['600999','000166','601995','601555','601901','601162','600958','002939','601990','002797','600909','601236',
 '601377','601456','601066','600837','601211','601688','600030','601696','000686','601375','601788','600109',
 '002500','002673','002945','000783','002926','002736','601878','600369','000728','601198','601108',
 '000776','601099','000712','601881','300059','600918','000750']
#code_list=['600030','601788','002500']
nb_text = '2020-12-31'
for code in code_list:
    s = {}
    #http://basic.10jqka.com.cn/601099/finance.html#finance
    browser.get('http://basic.10jqka.com.cn/'+ code +'/finance.html#finance')
    WebDriverWait(browser, 60 * 10, 0.5).until(EC.title_contains("同花顺金融服务网"))
    sleep(2)
    nb_date(nb_text)
    s['证券代码'] = code
    try:
        s['公司名称'] = browser.find_element_by_xpath("//div[@class='code fl']/div").text
        s['时间'] = nb_text
        if nb_num is not None:
            # 6------------------------------------------------------------
            # 杜邦分析结构图
            browser.execute_script("var q=document.documentElement.scrollTop=10000")
            (s['总资产'],s['资产增长率'])=zb_main('资产总额')
            (s['总负债'],var7)=zb_main('负债总额')
            (s['营业收入'],s['营收增长率'])=zb_main('营业总收入')
            (s['净利润'],s['净利润增长率'])=zb_main('净利润')

            # 1------------------------------------------------------------
            # 主要指标TAB
            # 滑动滚动条到TOP
            browser.execute_script("var q=document.documentElement.scrollTop=0")
            tab = browser.find_element_by_xpath("//ul[@class='newTab']/li[1]").click()
            sleep(3)
            browser.find_element_by_link_text('按报告期').click()
            s['ROE'] = '{:.2%}'.format(su_find("净资产收益率-摊薄", nb_num))

            # 2------------------------------------------------------------
            #资产负债表TAB
            browser.execute_script("var q=document.documentElement.scrollTop=0")
            browser.find_element_by_xpath("//ul[@class='newTab']/li[2]").click()
            sleep(3)
            s['保证金规模']=su_find("其中：客户存款(元)",nb_num)
            s['总股本']=su_find("实收资本（或股本）(元)",nb_num)

            # 3------------------------------------------------------------
            #利润表TAB
            browser.execute_script("var q=document.documentElement.scrollTop=0")
            browser.find_element_by_xpath("//ul[@class='newTab']/li[3]").click()
            sleep(3)
            s['手续费及佣金收入']=su_find("手续费及佣金净收入(元)",nb_num)
            s['投资收益']=su_find("投资收益(元)",nb_num)
            s['利息净收入']=su_find("利息净收入(元)",nb_num)
            s['归母净利润']=su_find("归属母公司所有者的净利润(元)",nb_num)
            var7=su_find("归属母公司所有者的净利润(元)",nb_num+4)
            try:
                s['归母净利润增长率']='{:.2%}'.format(((s['归母净利润'])-var7)/var7)
            except:
                s['归母净利润增长率']=""
            try:
                s['净利润率']='{:.2%}'.format(s['归母净利润']/s['营业收入'])
            except:
                s['净利润率'] =""
            s['营业支出']=su_find("二、营业支出(元)",nb_num)
            s['管理费用']=su_find("业务及管理费(元)",nb_num)
            s['信用减值损失']=su_find("信用减值损失(元)",nb_num)

            # 4------------------------------------------------------------
            #现金流量表TAB
            #browser.execute_script("var q=document.documentElement.scrollTop=0")
            #browser.find_element_by_xpath("//ul[@class='newTab']/li[4]").click()
            #sleep(2)

            # 5------------------------------------------------------------
            #券商专项指标TAB
            browser.execute_script("var q=document.documentElement.scrollTop=0")
            browser.find_element_by_xpath("//ul[@class='newTab']/li[5]").click()
            sleep(3)
            s['净资本']=su_find("净资本(元)",nb_num)
            var7=su_find("净资本(元)",nb_num+4)
            try:
                s['净资本增长率']='{:.2%}'.format((s['净资本']-var7)/var7)
            except:
                s['净资本增长率']=""
            s['净资产']=su_find("净资产(元)",nb_num)
            try:
                s['风险覆盖率']='{:.2%}'.format(su_find("净资本/各项风险资本准备之和",nb_num))
            except:
                s['风险覆盖率'] =""
    except:
        print("采集元素异常")
    pdata = pdata.append(s, ignore_index=True)
    print(s)
pd.DataFrame(pdata).to_excel('data.xlsx', sheet_name='Sheet1', index=False, header=True)
browser.quit()