# coding=utf-8
from Share import json
from Share import xlrd
from Share import DataHandle


# 检查json数据有效性 输入参数：表的路径、检查的列（列表）  返回：数据正确返空字典  或 有误的数据
def check_json_data(log_user, table_file, way, work_table_name='Sheet1', *line_list):
    log_user.debug('log_user:%s  table_file:%s way:%s work_table_name:%s line_list:%s'
                   % (log_user, table_file, way, work_table_name, line_list))
    ex = xlrd.open_workbook(table_file)               # 打开文件
    table = ex.sheet_by_name(work_table_name)
    linenum = 1                                         # 行号
    error_info = {}                                     # 错误信息
    while(1):
        for Column_number in line_list:              # 判断对应的列号
            try:
                if way == 'AllUint':
                    rq_date_type = table.cell(linenum, 7).value.encode("utf-8")
                    rp_date_type = table.cell(linenum, 12).value.encode("utf-8")
                    cmp_data = table.cell(linenum, Column_number).value.encode("utf-8")
                    if (rq_date_type == '1' and Column_number == 8) or (rp_date_type == '1' and Column_number == 13):
                        if Column_number == 8:
                            content = eval(table.cell(linenum, 9).value.encode("utf-8"))  # 取替换参数
                            cmp_data = DataHandle.DataHandle('1', cmp_data, content, log_user)
                        else:
                            content = eval(table.cell(linenum, 14).value.encode("utf-8"))  # 取替换参数
                            cmp_data = DataHandle.DataHandle('1', cmp_data, content, log_user)
                        log_user.info('被检查的为数据：%s' % cmp_data)
                    elif (rq_date_type == '2'and Column_number == 8) or (rp_date_type == '2' and Column_number == 13):
                        if Column_number == 8:
                            content = eval(table.cell(linenum, 9).value.encode("utf-8"))  # 取替换参数
                        else:
                            content = eval(table.cell(linenum, 14).value.encode("utf-8"))  # 取替换参数
                        log_user.debug("查询SQL:%s " % content)
                        state, content = DataHandle.PerformDBASql(content, log_user)                      # 执行sql，查询数据
                        if state != 0:
                            cmp_data = DataHandle.DataHandle('2', cmp_data, content, log_user)      # 替换数据
                            log_user.info('被检查的为数据：%s' % cmp_data)
                        else:
                            log_user.info("** 数据库操作失败:%s  **" % content)
                            if linenum in error_info:
                               error_list = error_info[linenum]
                               error_list.append(Column_number)
                               error_info[linenum] = error_list
                               log_user.debug("检查第:%s 第:%s 列的json数据   Fail" % (linenum, Column_number))
                            else:
                               error_info[linenum] = []
                               error_list = error_info[linenum]
                               error_list.append(Column_number)
                               error_info[linenum] = error_list
                               log_user.debug("检查第:%s 第:%s列的json数据   Fail" % (linenum, Column_number))
                            break
                    if cmp_data != '':
                        json.loads(cmp_data)             # 检查数据是否为有效的json
                        log_user.debug("工作表：%s 检查第:%s 行 第:%s 列的json数据  PASS"
                                       % (work_table_name, linenum, Column_number))
                elif way == 'Function':
                    table = ex.sheet_by_name(work_table_name)
                    if work_table_name == 'function':
                        check_json_date = table.cell(linenum, Column_number).value.encode("utf-8")
                        if check_json_date != '':
                            json.loads(check_json_date)             # 检查数据是否为有效的json
                            log_user.debug("工作表：%s 检查第:%s 行 第:%s 列的json数据  PASS"
                                           % (work_table_name, linenum, Column_number))

                    elif work_table_name == 'data':
                        check_json_date = table.cell(linenum, Column_number).value.encode("utf-8")
                        if check_json_date != '':
                            json.loads(check_json_date)             # 检查数据是否为有效的json
                            log_user.debug("工作表：%s 检查第:%s 行 第:%s 列的json数据  PASS"
                                           % (work_table_name, linenum, Column_number))

            except ValueError:
                if linenum in error_info:
                    error_list = error_info[linenum]
                    error_list.append(Column_number)
                    error_info[linenum] = error_list
                    log_user.debug("检查第:%s 第:%s 列的json数据   Fail" % (linenum, Column_number))
                else:
                    error_info[linenum] = []
                    error_list = error_info[linenum]
                    error_list.append(Column_number)
                    error_info[linenum] = error_list
                    log_user.debug("检查第:%s 第:%s列的json数据   Fail" % (linenum, Column_number))
                # log_user.error("errorinfo:%s" %error_info)
            except IndexError:
                return error_info
            # except Exception, e:
            #      log_user.error("异常信息:%s" %e)
        linenum += 1
    table.close()
    return error_info

if __name__ == '__main__':
    testlist=[8, 10, 11, 13, 15, 17, 19, 23]
    fi=r'C:\Users\huangjiajian\Desktop\CashBox_new\TestCase\function.xlsx'
    check_json_data()