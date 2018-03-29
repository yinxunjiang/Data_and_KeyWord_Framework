#encoding=utf-8
import logging
import logging.config
import os
from ProjectVar.Var import *
#读取配置日志文件
#log_file_path=ProjectVar.var.project_path+'\\conf\\Logger.conf'
logging.config.fileConfig(log_file_path)
#选择一个日志格式
logger=logging.getLogger("example02") #example01

def error(message):
    #打印debug级别的信息
    logger.error(message)

def info(message):
    #打印 info 级别的信息
    logger.info(message)

def warning(message):
    #打印 warnging级别的信息
    logger.warning(message)

if __name__=="__main__":
    print log_file_path
    error("hello")
    info("world!")
    warning("gloryroad!")
