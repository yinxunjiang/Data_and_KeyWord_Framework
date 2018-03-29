#encoding=utf-8
from selenium import webdriver
import time
from Util.Excel import *
from ProjectVar.var import *
from Util.FormatTime import *
from Util.log import *
'''此模块定义了新建联系人的操作，通过调用它可实现 邮箱登录--新建联系人的功能 '''

def new_linkman(driver,name,email,star,phone,remark):
    ad=Address(driver)
    ad.get_add_linkman().click()
    time.sleep(1)
    ad.get_name().send_keys(name)
    time.sleep(1)
    ad.get_email().send_keys(email)
    time.sleep(1)
    if star==u'是':
        ad.get_star().click()
        time.sleep(1)
    ad.get_phone().send_keys(phone)
    time.sleep(1)
    if not remark==None:
        ad.get_remark().send_keys(remark)
        time.sleep(1)
    ad.get_ok().click()
    time.sleep(3)

if __name__=='__main__':
    driver = webdriver.Chrome(executable_path="d:\\chromedriver")
    login(driver,'yinxunjiang123','gloryroad')
    visit_address_page(driver)
    new_linkman(driver,u'尹逊江','764400985@qq.com',u'是','18911734601','new man')

    '''
    driver=webdriver.Chrome(executable_path="d:\\chromedriver")
    path=os.path.join(project_path,'testdata',u"126邮箱联系人.xlsx")
    pe=ParseExcel(path)
    pe.get_sheet_by_name(u"126账号")
    for row in pe.get_all_rows()[1:]:
        username=row[1].value
        password=row[2].value
        dataexcel=row[3].value
        yes_or_no_execute=row[4].value
        try:
            if yes_or_no_execute=='y':
                login(username,password)
                row[5].value='pass'
                pe.get_sheet_by_name(dataexcel)
                try:
                    address_book()
                    for ro in pe.get_all_rows()[1:]:
                        if ro[7].value=='y':
                            new_linkman(ro[1].value,ro[2].value,ro[3].value,ro[4].value,ro[5].value)
                            ro[8].value=date_time_chinese()
                            ro[9].value=u'成功'
                            info('执行成功')
                except Exception,e:
                    ro[8].value=date_time_chinese()
                    ro[9].value=u'失败'
                    print 'failed'
                    error('执行失败')
        except Exception,e:
            row[5].value=u'失败'
            print 'failed'
            error('执行失败')
    pe.workbook_save()
    driver.quit()
    #new_linman(u'尹逊江','test1234@qq.com','18900000001',u'光荣之路测试')
    '''