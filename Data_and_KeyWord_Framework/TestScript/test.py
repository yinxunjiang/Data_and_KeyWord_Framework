#encoding=utf-8
from ProjectVar.Var import *
from Util.Excel import *
from Action.Action import *
from Util.Log import *
import re
'''获取excel表格内容，进行邮箱登录操作,并添加联系人 '''
if __name__=="__main__":
    #获取表格
    test_data_excel_file=ParseExcel(test_data_excel_path)
    #进入登录sheet页，进行登录操作
    test_data_excel_file.get_sheet_by_name('login')
    for row in test_data_excel_file.get_all_rows()[1:]:
            #获取动作、定位方式、定位表达式、操作值
            action_name=row[action_no].value
            locator_type=row[locator_type_no].value
            locator_expression=row[locator_expression_no].value
            action_value=row[value_no].value
            if locator_type is None and locator_expression is None and action_value is not None:
                #针对打开浏览器、访问地址、暂定、断言操作
                command_line='%s(u"%s")'%(action_name,action_value)
                print command_line
            #针对进入frame、点击操作
            elif locator_type is not None and locator_expression is not None and action_value is None:
                command_line='%s("%s","%s")'%(action_name,locator_type,locator_expression)
                print command_line
            #针对输入操作
            elif locator_type is not None and locator_expression is not None and action_value is not None:
                command_line='%s("%s","%s",u"%s")'%(action_name,locator_type,locator_expression,action_value)
                print command_line
            else:#针对关闭浏览器操作
                command_line=action_name+"()"
                print command_line
            try:
                start=time.time()
                #执行命令
                exec(command_line)
                end=time.time()
                execute_time="%.2f"%(end-start)
                #将耗时写入表格
                row[execute_time_no].value=str(execute_time)+'s'
                #将执行结果写入表格
                row[result_no].value=u'成功'
                row[result_no].font= Font(color="008800")#设置成绿色
                info(row[step_no].value+u' 执行成功.')
                #picturepath=get_screen_shot(capturepicturepath,dates(),times())
                #row[shot_picture_no].value=picturepath
            except Exception,e:
                #若用例执行失败，记录耗时、结果，并将错误截图地址写入表格
                print traceback.format_exc()
                error(row[step_no].value+u'执行失败:'+traceback.format_exc())
                row[execute_time_no].value=str(execute_time)+'s'
                row[result_no].value=u'失败'
                row[result_no].font= Font(color="FF0000")#设置成红色
                row[shot_picture_no].value=get_screen_shot(errorpicturepath,dates(),times())
                row[exception_no].value=traceback.format_exc()
    print '*'*50
#-----------------------进入添加联系人sheet页---------------------------------#
    #进入linkman sheet页，获取具体联系人数据
    test_data_excel_file.set_sheet_by_name('linkman')
    for ro in test_data_excel_file.get_all_rows()[1:]:
            name=ro[name_no].value
            email=ro[email_no].value
            phone=ro[phone_no].value
            text=ro[text_no].value
            keyword=ro[keyword_no].value
            #切换到添加操作sheet页，操作页面元素，并根据操作值不同把联系人信息传进去
            test_data_excel_file.get_sheet_by_name('add')
            for row in test_data_excel_file.get_all_rows()[1:-1]:
                    #获取动作、定位方式、定位表达式、操作值
                    action_name=row[action_no].value
                    locator_type=row[locator_type_no].value
                    locator_expression=row[locator_expression_no].value
                    action_value=row[value_no].value
                    if locator_type is None and locator_expression is None and action_value is not None:
                        #针对打开浏览器、访问地址、暂定、断言操作
                        if re.search('\${.*}',str(action_value)):
                            command_line='%s(u"%s")'%(action_name,keyword)
                            print command_line
                        else:
                            command_line='%s(u"%s")'%(action_name,action_value)
                            print command_line

                    #针对进入frame、点击操作
                    elif locator_type is not None and locator_expression is not None and action_value is None:
                        command_line='%s("%s","%s")'%(action_name,locator_type,locator_expression)
                        print command_line
                    #针对输入操作
                    elif locator_type is not None and locator_expression is not None and action_value is not None:
                        if re.search('\${name}',action_value):
                            command_line='%s("%s","%s",u"%s")'%(action_name,locator_type,locator_expression,name)
                            print command_line
                        elif re.search('\${email}',action_value):
                            command_line='%s("%s","%s",u"%s")'%(action_name,locator_type,locator_expression,email)
                            print command_line
                        elif re.search('\${phone}',action_value):
                            command_line='%s("%s","%s",u"%s")'%(action_name,locator_type,locator_expression,phone)
                            print command_line
                        elif re.search('\${text}',action_value):
                            command_line='%s("%s","%s",u"%s")'%(action_name,locator_type,locator_expression,text)
                            print command_line
                    try:
                        start=time.time()
                        #执行命令
                        exec(command_line)
                        end=time.time()
                        execute_time="%.2f"%(end-start)
                        #将耗时写入表格
                        row[execute_time_no].value=str(execute_time)+'s'
                        #将执行结果写入表格
                        row[result_no].value=u'成功'
                        row[result_no].font= Font(color="008800")#设置成绿色
                        info(row[step_no].value+u' 执行成功.')
                    except Exception,e:
                        #若用例执行失败，记录耗时、结果，并将错误截图地址写入表格
                        print traceback.format_exc()
                        error(row[step_no].value+u'执行失败:'+traceback.format_exc())
                        row[execute_time_no].value=str(execute_time)+'s'
                        row[result_no].value=u'失败'
                        row[result_no].font= Font(color="FF0000")#设置成红色
                        row[shot_picture_no].value=get_screen_shot(errorpicturepath,dates(),times())
                        row[exception_no].value=traceback.format_exc()
            test_data_excel_file.set_sheet_by_name('linkman')

    #针对关闭浏览器操作
    test_data_excel_file.set_sheet_by_name('add')
    row=test_data_excel_file.get_single_row(-1)
    action_name=row[action_no].value
    command_line=action_name+"()"
    print command_line
    try:
        start=time.time()
        #执行命令
        exec(command_line)
        end=time.time()
        execute_time="%.2f"%(end-start)
        #将耗时写入表格
        row[execute_time_no].value=str(execute_time)+'s'
        #将执行结果写入表格
        row[result_no].value=u'成功'
        row[result_no].font= Font(color="008800")#设置成绿色
        info(row[step_no].value+u' 执行成功.')
    except Exception,e:
        #若用例执行失败，记录耗时、结果，并将错误截图地址写入表格
        print traceback.format_exc()
        error(row[step_no].value+u'执行失败:'+traceback.format_exc())
        row[execute_time_no].value=str(execute_time)+'s'
        row[result_no].value=u'失败'
        row[result_no].font= Font(color="FF0000")#设置成红色
        row[shot_picture_no].value=get_screen_shot(errorpicturepath,dates(),times())
        row[exception_no].value=traceback.format_exc()
    test_data_excel_file.workbook_save()
