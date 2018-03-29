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
    #获取表格里的全部sheet页名字
    names=test_data_excel_file.get_all_sheet_names()
    #遍历所有sheet页
    for sheet in names[:2]:
        test_data_excel_file.set_sheet_by_name(sheet)
        #遍历从第二行开始的每一行内容，因为第一行是标题
        for row in test_data_excel_file.get_all_rows()[1:]:
            #获取动作、定位方式、定位表达式、操作值
            action_name=row[action_no].value
            locator_type=row[locator_type_no].value
            locator_expression=row[locator_expression_no].value
            action_value=row[value_no].value
            #结合存放测试数据的表格，有些项是没有值的，故需要根据实际情况判断如何拼接执行命令
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
                                #进入linkmansheet页
                test_data_excel_file.set_sheet_by_name('linkman')
                for ro in test_data_excel_file.get_all_rows()[1:]:
                    if re.search('\${name}',action_value):
                        command_line='%s("%s","%s",u"%s")'%(action_name,locator_type,locator_expression,ro[name_no].value)
                        print command_line
                    elif re.search('\${email}',action_value):
                        command_line='%s("%s","%s",u"%s")'%(action_name,locator_type,locator_expression,ro[email_no].value)
                        print command_line
                    elif re.search('\${phone}',action_value):
                        command_line='%s("%s","%s",u"%s")'%(action_name,locator_type,locator_expression,ro[phone_no].value)
                        print command_line
                    elif re.search('\${text}',action_value):
                        command_line='%s("%s","%s",u"%s")'%(action_name,locator_type,locator_expression,ro[text_no].value)
                        print command_line
                #跳回到刚才遍历的sheet页
                test_data_excel_file.set_sheet_by_name(sheet)
            else:#针对关闭浏览器操作
                command_line=action_name+"()"
                print command_line
            '''
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

    #保存表格
    test_data_excel_file.workbook_save()
    '''

