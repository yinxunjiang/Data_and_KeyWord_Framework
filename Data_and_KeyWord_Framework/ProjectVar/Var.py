#encoding=utf-8
import os
from Util.FormatTime import *
#打印文件所在目录
file_path=os.path.dirname(__file__)
#打印文件所在的工程目录
project_path=os.path.dirname(os.path.dirname(__file__))
#测试数据路径
test_data_excel_path=os.path.join(project_path,'TestData',u"126邮箱的测试用例.xlsx")
#test_data_excel_path=project_path+u"/testdata/126邮箱联系人.xlsx"
#日志文件
log_file_path=os.path.join(project_path,'conf','Logger.conf')
#获取页面元素配置文件路径
configFilePath=os.path.join(project_path,'conf','PageElementLocator.ini')
#浏览器驱动路径
ieDriverFilePath="d:\\IEDriverServer"
chromeDriverFilePath="d:\\chromedriver"
firefoxDriverFilePath="d:\\geckodriver"
#截图位置：
capturepicturepath=configFilePath=os.path.join(project_path,'ScreenPictures','CapturePicture')
errorpicturepath=configFilePath=os.path.join(project_path,'ScreenPictures','ErrorPicture')
#add sheet页位置
step_no=1
action_no=2
locator_type_no=3
locator_expression_no=4
value_no=5
execute_time_no=6
result_no=7
shot_picture_no=8
exception_no=9
#联系人表格位置
name_no=1
email_no=2
star_no=3
phone_no=4
text_no=5
keyword_no=6
yes_or_no_no=7
test_time_no=8
add_result_no=9



if __name__=='__main__':
    print file_path
    print project_path
    print test_data_excel_path
    print configFilePath
    print capturepicturepath
    print os.path.join(capturepicturepath,time_chinese())