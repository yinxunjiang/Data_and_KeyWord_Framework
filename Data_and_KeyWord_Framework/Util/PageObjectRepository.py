#encoding=utf-8
from ConfigParser import ConfigParser
from ProjectVar.Var import *
import os 
class ParsePageObjectRepositoryConfig(object):
    '''
    从定位配置文件中读取文件内容
    '''
    def __init__(self):
        #configFilePath=os.path.join(ProjectVar.var.project_path,'conf','PageElementLocator.ini')
        self.cf=ConfigParser()
        self.cf.read(configFilePath)

    def getItemsFromSection(self,sectionName):
        items=self.cf.items(sectionName)
        return dict(items)

    def getOptionValue(self,sectionName,optionName):
        return self.cf.get(sectionName,optionName)

if __name__=="__main__":
    pp=ParsePageObjectRepositoryConfig()
    print pp.getItemsFromSection("126mail_login")
    print pp.getOptionValue("126mail_login","login_page.username")
