#encoding=utf-8
from ProjectVar.Var import *
from Util.Excel import *
from Action.Action import *

def execute_test_data_file_case(test_data_file_path):
    test_data_excel_file = ParseExcel(test_data_file_path)
    test_data_excel_file.set_sheet_by_name("Sheet1")
    print test_data_excel_file.get_default_name()

    command_line = ""

    action_name = ""
    locator_method = ""
    lcoator_expression = ""
    action_value = ""

    for id, row in enumerate(test_data_excel_file.get_all_rows()[1:]):
        # print row[action_name_col_no].value,row[locator_method_col_no].value,row[locator_expression_col_no].value,row[action_value_col_no].value
        action_name = row[action_name_col_no].value
        locator_method = row[locator_method_col_no].value
        locator_expression = row[locator_expression_col_no].value
        action_value = row[action_value_col_no].value

        if locator_method is None and locator_expression is None and action_value is None:
            command_line = action_name + "()"
            print "command line:", command_line
        elif locator_method is not None and locator_expression is not None and action_value is None:
            command_line = action_name + "('" + locator_method + "','" + locator_expression + "')"
            print "command line:", command_line
        elif locator_method is None and locator_expression is None and action_value is not None:
            command_line = "%s(u'%s')" % (action_name, action_value)
            print "command line:", command_line
        else:
            command_line = '%s("%s","%s",u"%s")' % (
                action_name, locator_method, locator_expression, action_value)
            print "command line:", command_line

        try:
            time1 = time.time()
            result = eval(command_line)
            elapse_time = "%.2f" % (time.time() - time1)
            test_data_excel_file.write_cell_content(id + 2, action_elapse_time_col_no, elapse_time)
            test_data_excel_file.write_cell_content(id + 2, action_result_col_no, u"成功")
            test_data_excel_file.write_cell_content(id + 2, capture_screen_path_col_no, result)
            test_data_excel_file.save_excel_file()
        except Exception, e:
            elapse_time = "%.2f" % (time.time() - time1)
            print traceback.format_exc()
            test_data_excel_file.write_cell_content(id + 2, action_elapse_time_col_no, elapse_time)
            test_data_excel_file.write_cell_content(id + 2, action_result_col_no, u"失败")
            test_data_excel_file.write_cell_content(id + 2, action_excetion_info_col_no, traceback.format_exc())
            result = capture_error_screen()
            test_data_excel_file.write_cell_content(id + 2, capture_screen_path_col_no, result)
            test_data_excel_file.save_excel_file()

if __name__=='__main__':
    execute_test_data_file_case(test_data_excel_path)
