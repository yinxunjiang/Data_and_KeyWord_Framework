#encoding=utf-8
from selenium import webdriver
import time
from Util.ObjectMap import *
from ProjectVar.Var import *
import traceback
from Util.FormatTime import *

driver=None
#打开浏览器操作
def open_browser(browsername,*args):
    global driver
    try:
        if browsername.lower()=='ie':
            driver=webdriver.Ie(executable_path=ieDriverFilePath)
        elif browsername.lower()=='chrome':
            driver=webdriver.Chrome(executable_path=chromeDriverFilePath)
        elif browsername.lower()=='firefox':
            driver=webdriver.Firefox(executable_path=firefoxDriverFilePath)
    except Exception,e:
        raise e
#访问网址操作
def visit_url(url,*args):
    global driver
    try:
        driver.get(url)
    except Exception,e:
        raise e
#等待操作
def pause(seconds):
    seconds=float(seconds)
    time.sleep(seconds)
#进入frame结构操作
def enter_frame(locatormethod,locatorexpression,*args):
    global driver
    try:
        driver.switch_to.frame(getElement(driver,locatormethod,locatorexpression))
    except Exception,e:
        raise e
#输入操作
def input_string(locatormethod,locatorexpression,content,*args):
    global driver
    try:
        element=getElement(driver,locatormethod,locatorexpression)
        element.clear()
        element.send_keys(content)
    except Exception,e:
        raise e
#点击操作
def click(locatormethod,locatorexpression,*args):
    global driver
    try:
        element=getElement(driver,locatormethod,locatorexpression)
        element.click()
    except Exception,e:
        raise e
#关闭浏览器操作
def close_browser(*args):
    global driver
    try:
        driver.quit()
    except Exception,e:
        raise e
#截图，截图放到ScreenPictures目录下的相应文件夹下，根据日期动态创建保存目录
def get_screen_shot(picturepath,dirname,filename):
    global driver
    try:
        if not os.path.exists(os.path.join(picturepath,dirname)):
            os.mkdir(os.path.join(picturepath,dirname))
        driver.get_screenshot_as_file(os.path.join(picturepath,dirname,filename)+'.png')
    except Exception,e:
        raise e
    #返回图片地址
    return os.path.join(picturepath,dirname,filename)+'.png'
#断言操作，判断某个内容是否在页面源码中
def assert_word(word,*args):
    global driver
    try:
        assert word in driver.page_source
        #print '断言成功！'
        #get_screen_shot(capturepicturepath,dates(),times())
    except Exception,e:
        #断言失败，将页面截图保存，截图存放在当前年月日的目录里，截图名称以当前时分秒命名
        get_screen_shot(errorpicturepath,dates(),times())
        raise e

'''
def login(usernameandpassword,*args):
    username,password=usernameandpassword.split("||")
    open_browser('chrome')
    visit_url('http://www.126.com')
    pause(3)
    enter_frame('id','x-URS-iframe')
    input_string('xpath',"//input[@name='email']",username)
    input_string('xpath',"//input[@name='password']",password)
    click('id','dologin')
    pause(3)
    assert_word(u'退出')
    #get_screen_shot(capturepicturepath+'/11.png')
    #close_browser()

def add_new_contacts(name,email,phone,remark,*args):
    click("xpath","//div[text()='通讯录']")
    click("xpath","//span[text()='新建联系人']")
    pause(0.5)
    input_string("xpath","//a[@title='编辑详细姓名']/preceding-sibling::div/input",name)
    input_string("xpath","//*[@id='iaddress_MAIL_wrap']//input",email)
    input_string("xpath","//*[@id='iaddress_TEL_wrap']//dd//input",phone)
    input_string("xpath","//textarea",remark)
    click("xpath","//span[.='确 定']")
    pause(2)
    assert_word(name)
    #get_screen_shot(capturepicturepath+'/22.png')
'''


if __name__=='__main__':
    login("yinxunjiang123||gloryroad")
    add_new_contacts(u'刘德华','1000001@qq.com','18000000001',u'你最帅')



