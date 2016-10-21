# coding=utf-8
import os
import logging
import logging.config
from Share import xlrd
from Share.string import Template
from Share import Tools
from Share import GlobalConfiguration
from Share.pypinyin import lazy_pinyin, STYLE_TONE2


CurrentPath = os.getcwd()
# CasePath=CurrentPath + '\TestCase\interface.xlsx'
LogPath = CurrentPath + '\Configuration\logging.conf'
ConfPath = CurrentPath + '\Configuration\TestConfiguration.cfg'

RunMode = ''
logging.config.fileConfig(LogPath)
LoggUser = logging.getLogger('Operate')

# 接口用例开始模板
Content = '''# coding=utf-8
import logging

from Share import unittest
from Share import xlrd
from Share import DataHandle
from Share import GlobalConfiguration
# print os.path.dirname(logging.__file__) #查看模块路径


class ${ClassName}(unittest.TestCase):
    u"""${ClassDetails}"""

    def setUp(self):
        try:
            self.loger_user = logging.getLogger('${user}')
            self.loger_user.info('------ 开始初始化 ${ClassDetails} 测试用例 ------ ')
            # ex = xlrd.open_workbook(r'G:\Python\CashBox_new\TestCase\\function.xlsx') # 单文件时调时用
            ex = xlrd.open_workbook(GlobalConfiguration.get_path_info('${CaseFile}', self.loger_user))  # main调时用
            self.table = ex.sheet_by_name('interface')
            self.test_host = GlobalConfiguration.get_service_address()  # 环境选择
            self.loger_user.info('------ 初始化 ${ClassDetails} 测试用例结束 ------ ')
        except Exception, e:
            self.loger_user.error('初始化错误信息:%s' % e)

'''


# 功能用例开始模板
function_start_Content = '''# coding=utf-8
import logging

from Share import unittest
from Share import xlrd
from Share import DataHandle
from Share import GlobalConfiguration
# print os.path.dirname(logging.__file__) #查看模块路径


class ${ClassName}(unittest.TestCase):
    u"""${ClassDetails}"""
    def setUp(self):
        try:
            self.loger_user = logging.getLogger('${user}')
            self.loger_user.info('------ 开始初始化 ${ClassDetails} 功能测试用例 ------ ')
            # self.table = xlrd.open_workbook(r'G:\Python\CashBox_new\TestCase\\function.xlsx') # 单文件时调时用
            self.table = xlrd.open_workbook(GlobalConfiguration.get_path_info('${CaseFile}', self.loger_user))  # main调时用
            # 地址 数据库存 需要改成从公用文件中提取
            self.test_host = GlobalConfiguration.get_service_address()  # 环境选择
            self.loger_user.info('------ 初始化 ${ClassDetails} 功能测试用例结束 ------ ')
        except Exception, e:
            self.loger_user.error('初始化错误信息:%s' % e)

    def data_preparation(self, set_out_sql):                     # 执行初始化或还原脚本
            self.loger_user.info('开始执行数据初始化脚本')
            state, content = DataHandle.PerformDBASql(set_out_sql, self.loger_user)
            if state == 0:
                self.loger_user.info('执行数据初始化脚本失败  状态：%s  返回内容：%s' % (state, content))
                self.assertTrue(state, content)
            self.loger_user.info('执行数据初始化脚本成功 状态：%s  返回内容：%s' % (state, content))

    def check_the_database_data(self, test_case_line):                                    # 检查数据库数据
        self.loger_user.info("开始数据库数据查检")
        ex_table = self.table.sheet_by_name('data')
        data_comparison = eval(ex_table.cell(test_case_line, 9).value.encode("utf-8"))     # 参考数据
        database_data = eval(ex_table.cell(test_case_line, 10).value.encode("utf-8"))      # 数据库数据
        sql_statements = eval(ex_table.cell(test_case_line, 11).value.encode("utf-8"))     # 查询语句
        self.loger_user.info("%s" % sql_statements)
        state, content = DataHandle.PerformDBASql(sql_statements, self.loger_user)          # 执行sql，查询数据
        if state != 0:
            database_data = eval(DataHandle.DataHandle('2', database_data, content, self.loger_user))  # 替换数据
            self.loger_user.debug("替换数据完成")
        else:
            self.loger_user.info("数据操作失败:%s  **" % content)
            self.assertTrue(state, content)
        result, info_a, info_b = DataHandle.cmp_json(database_data, data_comparison, self.loger_user)
        if result != '1':                                           # 调用断言方式 返回1 正常  ，非1不正确
            self.loger_user.info('数据库数据对比失败 数据库的数据:\\n响应返回的值: %s 参考的数据 :%s'
            % (str(info_a), str(info_b)))
            self.assertTrue(result, ('数据库的数据值：' + info_a + '   参考的数据值：'+ info_b))
        self.loger_user.info("数据库数据检查结束")

    def results_the_assertion(self, state, info_a, info_b, test_case_line_num):       # 函数功能结果断言
        ex_table = self.table.sheet_by_name('data')
        test_case_name = ex_table.cell(test_case_line_num, 1).value.encode("utf-8")
        if state == 0:
            self.loger_user.debug('失败-状态：%s  内容_A：%s  内容_b：%s' % (state, info_a, info_b))
            self.loger_user.info('执行用例状态:失败 第 ' + str(test_case_line_num) + '条   ' + test_case_name + ' 测试用例结束')
            self.assertEqual(info_a, info_b)
        elif state in(1, 3):     # 1 执行正常流程成功状态  3 执行不完整流程成功状态
            self.loger_user.debug('状态：%s  内容_A：%s  内容_b：%s' % (state, info_a, info_b))
            self.loger_user.info('执行用例状态:通过 第 ' + str(test_case_line_num) + '条   ' + test_case_name + ' 测试用例结束')
            self.assertEqual(info_a, info_b)
        elif state == 2:
            self.loger_user.debug('状态：%s  内容_A：%s  内容_b：%s' % (state, info_a, info_b))
            self.loger_user.info('执行用例状态:失败 第 ' + str(test_case_line_num) + '条   ' + test_case_name + ' 测试用例结束')
            self.fail(info_b)
        else:
            self.loger_user.info('执行用例状态:未知 执行第 ' + str(test_case_line_num) + '条   ' + test_case_name + ' 测试用例结束')
            self.fail('The state is unknown')

    def ${functionality}(self, table_name, test_case):
        state = ''                 # 用于防止功能组合为空，测试组结果为通过的情况
        self.loger_user.debug('state:%s ' % state)
'''

# 功能组合模板
function_Content = '''        state, interface, info = DataHandle.function_core(table_name, ${function_line}, test_case, self.test_host, self.loger_user)
        self.loger_user.debug('状态：%s  interface：%s  info：%s' % (state, interface, info))
        if state != 1 or state == 3:
            return state, interface, info

'''
# 功能组合结尾
function_Content_end = '''        if state == 1:
            return 1, '', ''
        else:
            return 2, '', 'Function combination content is empty'

'''


# 功能测试用例模板
function_Content_case = '''    def ${test_case_name}(self):
        u"""${FunctionDescribe}"""
        test_case_line_num = ${line}
        self.loger_user.info('开始执行测试用例   ')
        ex_table = self.table.sheet_by_name('data')
        set_out = str(ex_table.cell(test_case_line_num, 3).value)                           # 请求前是否需要初始化标志
        check_data_sign = str(ex_table.cell(test_case_line_num, 8).value)                   # 数据库数据检验标志
        ready_case_test = ex_table.cell(test_case_line_num, 13).value                       # 预执行测试用例
        self.loger_user.info('ready_case_test:%s  type:%s ' % (ready_case_test, type(ready_case_test)))
        if len(ready_case_test) > 0:
            DataHandle.function_preset(self.table, int(ready_case_test), self.test_host, self.loger_user)
        if set_out == '1':
            set_out_sql = eval(ex_table.cell(test_case_line_num, 4).value.encode("utf-8"))  # 获取sql
            self.data_preparation(set_out_sql)
        state, info_a, info_b = self.${function}(self.table, test_case_line_num)
        # 增加数据库数据对比
        if state == 1 and check_data_sign == '1':
            self.check_the_database_data(test_case_line_num)
        ex_table = self.table.sheet_by_name('data')
        set_out = str(ex_table.cell(test_case_line_num, 5).value)                         # 数据还原标志
        if set_out == '1':
            set_out_sql = eval(ex_table.cell(test_case_line_num, 6).value.encode("utf-8"))  # 获取sql
            self.data_preparation(set_out_sql)
        self.results_the_assertion(state, info_a, info_b, test_case_line_num)               # 调用断言

'''

# 测试用例模板
TestCaseContent = '''    def ${function}(self):
        u"""${FunctionDescribe}"""

        state, info_a, info_b = DataHandle.request_core(self.table, ${line}, self.test_host, self.loger_user)
        set_down = str(self.table.cell(${line}, 18).value)              # 请求后还原标志位
        if set_down == '1':
            self.loger_user.info('开始执行执行还原脚本')
            set_down_sql = eval(self.table.cell(${line}, 19).value.encode("utf-8"))
            date_state, content = DataHandle.PerformDBASql(set_down_sql, self.loger_user)
            self.loger_user.info('还原脚本执行结束。状态：%s  返回内容：%s' % (state, content))
            if date_state == 0:
                self.assertEqual('database operation Fail',  '')
        if state in (0, 2):
            self.loger_user.debug('失败-状态：%s  内容_A：%s  内容_b：%s' % (state, info_a, info_b))
            self.loger_user.info('执行第 ${line} 条 ${FunctionDescribe} 测试用例结束')
            self.assertEqual(info_a, info_b)
        elif state in (1, 3):
            self.loger_user.debug('状态：%s  内容_A：%s  内容_b：%s' % (state, info_a, info_b))
            self.loger_user.info('执行第 ${line} 条 ${FunctionDescribe} 测试用例结束')
            self.assertEqual(info_a, info_b)
        else:
            self.loger_user.info('执行第 ${line} 条 ${FunctionDescribe} 测试用例结束')
            self.assertEqual(info_a, info_b)

'''
# 用例结束模板
EndTestCaseConten = '''    def tearDown(self):
            self.loger_user.info('------ 测试还原 ------\\n')

if __name__ == '__main__':
    import logging.config
    logging.config.fileConfig(r'${Log}')
    root_logger = logging.getLogger('root')    # 使用打印日志用户
    GlobalConfiguration.ConfigInit(r'${Config}', root_logger)

    unittest.main()
'''

# main函数固定信息
MainStandard = '''# coding=utf-8
import time
from Share import unittest
from Share import doctest
from Share import Report
from Share import GlobalConfiguration
from TestScript.${FolderName} import ${UnitName}


def perform_test_case(log_user, test_report):
    log_user.info('开始组建测试用例')
    suite = doctest.DocTestSuite()
 '''
# main内容
MainContent = '''
    suite.addTest(unittest.makeSuite(${UnitName}.${ClassName}))'''

MainEnd = '''

    log_user.info('组建测试用例结束')
    log_user.info('测试报告标题：%s' % test_report)
    runner, fp = Report.Report(FilePath=GlobalConfiguration.get_path_info(test_report, log_user), ReportTitle=test_report)
    log_user.info('---- 开始执行测试用例集合 ----')
    start_time = str(time.strftime('%Y-%m-%d  %H:%M:%S', time.localtime(time.time())))
    runner.run(suite)
    end_time = str(time.strftime('%Y-%m-%d  %H:%M:%S', time.localtime(time.time())))
    log_user.info('开始运行时间:%s  运行结束时间:%s' % (start_time, end_time))
    # unittest.TextTestRunner(verbosity=2).run(suite) # 打印详细信息
'''

hint = '''
***请选择创建脚本内容方式，请输入相应方式的序号，再按回车确定***
1.创建--单接口自动化--测试脚本，脚本内容来源于 interface.xlsx 的用例；
2.创建--功能接口自动化--测试脚本，脚本内容来源 function.xlsx 的用例；
3.创建--压力测试--脚本，脚本内容来源 pressure.xlsx 的用例；
4.创建--WebUi自动化测试--脚本，脚本内容来源 ui.xlsx 的用例；
**数据格式校验**
5.只校验 interface.xlsx 的文件的JSON数据有效性；
6.只校验 function.xlsx  的文件的JSON数据有效性；
7.只校验 pressure.xlsx  的文件的JSON数据有效性；
8.只校验 ui.xlsx 的文件的JSON数据有效性；
'''

error_info = {}


# 创建脚本文件 输入参数：测试文件类型  当前路径  表对象 模板内容  行号  功能标识  日志用户   返回：已找开的文件对象  类名  行号
def new_file(case_file, current_path, table, content=Content, line=1, feature_identifier='', logg_user=LoggUser):
    try:
        logg_user.debug('调用 new_file 函数。case_file：%s current_path:%s table:%s content:%s line:%s feature_identifier:%s'
                        % (case_file, current_path, table, content, line, feature_identifier))
        ex_table = table
        case_path = current_path
        way = get_chars(case_path, 2)
        logg_user.debug('way:%s  lien:%s case_file：%s  current_path:%s table:%s  feature_identifier:%s'
                        % (way, line, case_file, current_path, table, feature_identifier))
        if way == 'AllUint':
            while (1):
                table = ex_table.sheet_by_name('interface')
                script_name = table.cell(line, 1).value.encode('utf-8')
                script_name = get_chars(script_name)                     # 取接口名字
                test_sign = table.cell(line, 26).value.encode('utf-8')  # 是否生效标志
                content = Content
                function_name = table.cell(line, 2).value.encode('utf-8')  # 接口说明
                if script_name == '':
                    return '0', '创建结束(名字为空)',''
                elif test_sign == '1':
                    logg_user.info('第 %s 用例标为：%s ，标志为用例不生效，不创建文件' %(line, test_sign))
                    line += 1
                    continue
                unit_file_name = case_path + script_name + '.py'
                unit_test_name = script_name + 'Class'
                logg_file_user = get_logg_user()
                auto_test_name = open(unit_file_name, 'w')
                logg_user.info('创建脚本:%s' % unit_file_name)
                content = Template(content)
                content = content.substitute(CaseFile=case_file, ClassName=unit_test_name,
                                           ClassDetails=function_name, user=logg_file_user)
                logg_user.debug('内容:%s' % content)
                auto_test_name.writelines(content)
                return auto_test_name, unit_test_name, line
        elif way == 'Function':
            befroe_lien = line
            while (1):
                table = ex_table.sheet_by_name('data')
                test_case_service = table.cell(befroe_lien, 11).value.encode('utf-8')
                if test_case_service == '1':
                    befroe_lien += 1
                    logg_user.info('data 第 %s 为：%s ，标志为用例不生效，不创建文件' % (befroe_lien, test_case_service))
                    continue
                feature_identifier = table.cell(befroe_lien, 2).value.encode('utf-8')      # 生效的测试用列标能标志
                line = 1
                while (1):
                    table = ex_table.sheet_by_name('function')
                    class_name = table.cell(line, 3).value                   # 取功能标识
                    if class_name != feature_identifier:
                        line += 1
                        continue
                    script_name = table.cell(line, 4).value                  # 取功能描述
                    function_name = script_name                              # 保存原功能描述
                    test_case_name = lazy_pinyin(script_name, style=STYLE_TONE2)
                    script_name = ''
                    for num in range(len(test_case_name)):                 # 将功能描述拼接成字符串
                        script_name = script_name + test_case_name[num]
                        logg_user.debug('way:%s  class_name:%s   功能函数的名字：%s' % (way, class_name, script_name))
                        logg_user.info('feature_identifier:%s  len:%s' % (feature_identifier, len(feature_identifier)))
                        if script_name == '':
                            return '0', '创建结束(名字为空)',''
                    if check_there(feature_identifier, 'check_function') == 0:
                        logg_user.info('data工作表中的第 %s 用例标为：%s  功能，在function表中没查对应的功能标识'
                                       % (line, feature_identifier))
                        return '00', 'function表中不存在的功能标识:'+ feature_identifier, befroe_lien
                    unit_file_name = case_path + script_name + '.py'
                    unit_test_name = class_name + 'Class'
                    logg_file_user = get_logg_user()
                    table = ex_table.sheet_by_name('data')
                    test_case_line = table.cell(befroe_lien, 0).value.encode('utf-8')
                    test_case_title = table.cell(befroe_lien, 1).value.encode('utf-8')
                    functionality = table.cell(befroe_lien, 2).value.encode('utf-8')
                    auto_test_name = open(unit_file_name, 'w')
                    logg_user.info('创建脚本:%s' % unit_file_name)
                    content = function_start_Content
                    content = Template(content)
                    logg_user.debug('case_file:%s  unit_test_name:%s  test_case_name:%s   functionality:%s'
                                % (case_file, unit_test_name, test_case_title, functionality))
                    content = content.substitute(ClassName=unit_test_name, ClassDetails=function_name, user=logg_file_user,
                                           CaseFile=case_file, test_case_line=test_case_line, test_case_name=test_case_title,
                                           functionality=functionality)
                    logg_user.debug('内容:%s' % content)
                    auto_test_name.writelines(content)
                    combination_function(auto_test_name, unit_test_name, function_Content)
                    return auto_test_name, unit_test_name, befroe_lien
    except IndexError:
        logg_user.info('所有测试用例标志为不生效，不创建文件')
        return '0', '创建结束(名字为空)', ''


# 函数功能：组合接口  输入参数：脚本文件对象 功能标志 行号  日志用户  返回：状态  描述
def combination_function(auto_test_name, function_sign, test_case_content=function_Content, line=1):
    LoggUser.debug('调用 combination_function 函数。传入参数：auto_test_name：%s function_sign:%s test_case_content:%s '
                   ' line:%s' % (auto_test_name, function_sign, test_case_content, line))
    table = ex.sheet_by_name('function')
    function_sign = function_sign[0:-5]
    while(1):
        try:
            next_function_sign = table.cell(line, 3).value.encode('utf-8')
            service_sign = str(table.cell(line, 14).value)
            LoggUser.debug('line:%s function_sign:%s  next_function_sign:%s  service_sign:%s '
                           % (line, function_sign, next_function_sign, service_sign))
            if next_function_sign == function_sign and service_sign != '1':
                function_line_num = table.cell(line, 0).value.encode('utf-8')
                content = Template(test_case_content)
                content = content.substitute(function_line=function_line_num)
                auto_test_name.writelines(content)
                LoggUser.info('功能组合内容:\n%s' % content)
            elif service_sign == '1' and next_function_sign == function_sign:
                LoggUser.info('第 %s 条测试用例标志为不生效，不创建内容' %line)
            line += 1
        except IndexError:
            auto_test_name.writelines(function_Content_end)
            LoggUser.info('function整表已查询完')
            break


# 创建测试用例
def MakeTestCase(AutoTestName, UnitTestName, FilePath, table, CaseFile, TestCaseContent=TestCaseContent, line=1,
                 EndTestCaseConten=EndTestCaseConten, LogPat=LogPath, Conf=ConfPath):
    LoggUser.debug('调用 MakeTestCase 函数。AutoTestName：%s UnitTestName:%s FilePath:%s table:%s CaseFile:%s'
                   ' TestCaseContent:%s line:%s EndTestCaseConten:%s LogPat:%s Conf:%s'
      % (AutoTestName, UnitTestName, FilePath, table, CaseFile,TestCaseContent, line, EndTestCaseConten, LogPat, Conf))
    before_sign = ''
    while( 1 ):
        try:
            way = get_chars(FilePath, 2)
            if way == 'AllUint':
                table = ex.sheet_by_name('interface')
                script_name = table.cell(line, 1).value.encode('utf-8')
                script_name = get_chars(script_name)                      # 取接口名字
                TestCase = 'test_' + table.cell(line, 0).value.encode('utf-8')
                TestDescribe = table.cell(line, 3).value.encode('utf-8')            # 测试用例说明
                test_sign = table.cell(line, 26).value.encode('utf-8')             # 用例生效标志
                NesUnitTestName = script_name + 'Class'
            elif way == 'Function':
                table = ex.sheet_by_name('data')
                TestCase = 'test_' + table.cell(line, 0).value                        # 取data 表的测试用例序号
                TestDescribe = table.cell(line, 1).value.encode('utf-8')              # 取data 表的测试用例标题
                function_data_sing = table.cell(line, 2).value.encode('utf-8')        # 取data 表的功能标志
                test_sign = table.cell(line, 11).value.encode('utf-8')                # 用例生效标志
                if before_sign == '':
                    before_sign = function_data_sing
                elif before_sign != function_data_sing:
                    before_sign = function_data_sing
                    script_name = check_there(function_data_sing, 'get_function_description')   # 取function 表的脚本名
                    test_case_name = lazy_pinyin(script_name, style=STYLE_TONE2)
                    script_name = ''
                    for num in range(len(test_case_name)):                             # 将列表的拼音 拼接成字符串
                        script_name = script_name + test_case_name[num]
                NesUnitTestName = function_data_sing + 'Class'
                LoggUser.debug('UnitTestName:%s  NesUnitTestName:%s  function_data_sing：%s '
                           % (UnitTestName, NesUnitTestName, function_data_sing))
            LoggUser.debug('UnitTestName:%s  NesUnitTestName:%s ' % (UnitTestName, NesUnitTestName))
            if UnitTestName == NesUnitTestName:
                if test_sign != '1':
                    Content = Template(TestCaseContent)
                    if way == 'AllUint':
                        Content = Content.substitute(function=TestCase, FunctionDescribe=TestDescribe, line=line)
                    elif way == 'Function':
                        exist = check_there(function_data_sing)
                        LoggUser.debug('way:%s  function_data_sing:%s exist:%s  exist type:%s' %
                                       (way, function_data_sing, exist, type(exist)))
                        if exist == 0:
                            LoggUser.info('测试用例找不到对应的功能:%s' % function_data_sing)
                            continue
                        Content = Content.substitute(test_case_name=TestCase, FunctionDescribe=TestDescribe, line=line,
                                                   function=function_data_sing)
                    AutoTestName.writelines(Content)
                    LoggUser.info('测试用例内容:\n%s' % Content)
                else:
                    LoggUser.info('第 %s 条测试用例，标志为不生效，不创建测试用例' %line)
                line += 1
            elif NesUnitTestName != '':                  # 开始创建结尾
                Content = Template(EndTestCaseConten)
                Content = Content.substitute(Log=LogPat, Config=Conf)
                AutoTestName.writelines(Content)
                logging.info('开始创造新脚本')  # 结束创建结尾
                logging.debug('CaseFile:%s FilePath:%s table:%s line:%s' % (CaseFile, FilePath, ex, line))  # 结束创建结尾
                if way == 'AllUint':
                    TAutoTestName, TUnitTestName, line = new_file(CaseFile, FilePath, ex, line=line)
                    if TAutoTestName in ('0', '00'):
                        LoggUser.info('创建脚本信息：%s' % TUnitTestName)
                        break
                elif way == 'Function':
                    TAutoTestName, TUnitTestName, line = new_file(CaseFile, FilePath, ex, line=line,
                                                                feature_identifier=function_data_sing)
                    if TAutoTestName in ('0', '00'):
                        LoggUser.info('创建脚本信息：%s' % TUnitTestName)
                        error_info[line] = TUnitTestName
                        break

                if way == 'AllUint':
                    MakeTestCase(AutoTestName=TAutoTestName, UnitTestName=TUnitTestName, FilePath=FilePath, table=table,
                                 CaseFile=CaseFile, line=line)
                elif way == 'Function':
                    MakeTestCase(AutoTestName=TAutoTestName, UnitTestName=TUnitTestName, FilePath=FilePath, table=table,
                                 CaseFile=CaseFile, line=line, TestCaseContent=function_Content_case)
                AutoTestName.close()
                break
            else:
                AutoTestName.writelines(EndTestCaseConten)
                LoggUser.info('测试用例结束的内容：%s' % EndTestCaseConten)
                AutoTestName.close()
                break
        except IndexError:
            Content = Template(EndTestCaseConten)
            Content = Content.substitute(Log=LogPat, Config=Conf)
            AutoTestName.writelines(Content)
            LoggUser.info('测试用例结束的内容：%s' % EndTestCaseConten)
            AutoTestName.close()
            break
        except Exception, e:
            logging.error('运行错误:%s' % e)
            AutoTestName.close()
            error_info[line] = e
            break


# 创建main主函数
def CreateMain(FolderName, CurrentPath = CurrentPath, line=1):
    LoggUser.debug('调用 CreateMain 函数。传入参数：FolderName:%s CurrentPath：%s, line:%s'
                   % (FolderName, CurrentPath, line))
    UnitFileName = CurrentPath + '\Perform_test_case.py'
    AllUintName = ''
    last_class_name = ''
    AllCalssName = {}
    way = get_chars(CurrentPath, 2)
    if way == 'AllUint':
        while (1):
            try:
                table = ex.sheet_by_name('interface')
                ScriptName = table.cell(line, 1).value.encode('utf-8').strip()
                ScriptName = get_chars(ScriptName)
                test_sign = table.cell(line, 26).value.encode('utf-8')   # 测试用例生效标志
                NesUnitTestName = ScriptName + 'Class'
                if AllUintName.count(ScriptName) < 1 and ScriptName != '' and test_sign != '1':
                    if line == 1 or AllUintName == '':
                        AllUintName = ScriptName
                        AllCalssName[line] = {ScriptName: NesUnitTestName}
                    else:
                        AllUintName = ScriptName + ',' + AllUintName
                        AllCalssName[line] = {ScriptName: NesUnitTestName}
                line += 1
            except IndexError:
                LoggUser.info('所有脚本名信息%s' % AllCalssName)
                break
    elif way == 'Function':
        while(1):
            try:
                table = ex.sheet_by_name('data')
                class_name = table.cell(line, 2).value                   # 取功能标识
                ScriptName = check_there(class_name, 'get_function_description')
                if ScriptName == 0:
                    LoggUser.info('%s功能标识，找不到对应的脚本名称' % class_name)
                    break
                test_case_name = lazy_pinyin(ScriptName, style=STYLE_TONE2)
                ScriptName = ''
                NesUnitTestName = class_name + 'Class'
                for num in range(len(test_case_name)):
                    ScriptName = ScriptName + test_case_name[num]
                if last_class_name == '' or last_class_name != class_name:
                    test_sign = str(check_there(class_name, 'check_test_case'))  # 测试用例生效标志
                    last_class_name = class_name
                LoggUser.debug('line:%s 脚本信息的总数:%s  ScriptName:%s test_sign:%s '
                       % (line, AllUintName.count(ScriptName), ScriptName, test_sign))
                if AllUintName.count(ScriptName) < 1 and ScriptName != '' and test_sign != '0':
                    if line == 1 or AllUintName == '':
                        AllUintName = ScriptName
                        AllCalssName[line] = {ScriptName: NesUnitTestName}
                    else:
                        AllUintName = ScriptName + ',' + AllUintName
                        AllCalssName[line] = {ScriptName: NesUnitTestName}
                line += 1
            except IndexError:
                LoggUser.info('所有脚本名信息%s' % AllCalssName)
                break

    AutoTestName = open(UnitFileName, 'w')
    content = Template(MainStandard)
    content = content.substitute(FolderName=FolderName, UnitName=AllUintName)
    LoggUser.info('Content:\n%s' % content)
    AutoTestName.writelines(content)
    for k in sorted(AllCalssName.keys()):  # 按顺序增加
        twematter = AllCalssName[k]
        for key, v in sorted(twematter.items()):    #
            content = Template(MainContent)
            content = content.substitute(UnitName=key, ClassName=v)
            LoggUser.info('主函数内容:%s' % content)
            AutoTestName.writelines(content)
    AutoTestName.writelines(MainEnd)
    LoggUser.info('MainEnd:%s' % MainEnd)
    AutoTestName.close()


# 创建脚本方式处理 传入参数 方式  返回：当前创建模式 用例文件路径 创建脚本路径
def creation_method(run_mode, current_path=CurrentPath):
    if run_mode == '1' or run_mode == '5':
        case_path = current_path + '\TestCase\interface.xlsx'
        file_path = current_path + '\TestScript\AllUint\\'
        return run_mode, case_path, file_path, 'AllUint'
    elif run_mode == '2'or run_mode == '6':
        case_path = current_path + '\TestCase\\function.xlsx'
        file_path = current_path + '\TestScript\Function\\'
        return run_mode, case_path, file_path, 'Function'
    elif run_mode == '3'or run_mode == '7':
        case_path = current_path + '\TestCase\pressure.xlsx'
        file_path = current_path + '\TestScript\Pressure\\'
        return run_mode, case_path, file_path, 'Pressure'
    elif run_mode == '4'or run_mode == '8':
        case_path = current_path + '\TestCase\ui.xlsx'
        file_path = current_path + '\TestScript\AutoUi\\'
        return run_mode, case_path, file_path, 'AutoUi'
    else:
        return 0, 'error', 'Not RunMod', ''  # 返回错误


# 获取日志用户
def get_logg_user():
    if RunMode == '1':
        return 'Example'
    elif RunMode == '2':
        return 'function'
    else:
        return 'Operate'


# 获取接口名称
def get_chars(character, way=1):
    if way == 1:  # 取接口名字
        decollator = character.split('/')
        character = decollator[-1]
        decollator = character.split('.')
        character = decollator[0]
    elif way == 2:  # 取倒最后一个文件名字
        decollator = character.split('\\')
        character = decollator[-2]
    return character


# 检查测试用例 在功能否找到  传要检查的功能标志  返回 1：表示存在  0表示不存在
def check_there(sing, way='check_function'):
    try:
        LoggUser.debug('调用 check_there 函数。sing：%s  way:%s' %(sing, way))
        line = 1
        if way == 'check_function':
            table = ex.sheet_by_name('function')  # 检查测试对应的功能是否存在
            while(1):
                function_sign = table.cell(line, 3).value.encode('utf-8')
                server_sign = table.cell(line, 14).value
                if sing == function_sign and server_sign != '1':
                    return 1
                line += 1
        elif way == 'check_test_case':   # 检查测试用例是否生效
            table = ex.sheet_by_name('data')
            while(1):
                function_sign = table.cell(line, 2).value.encode('utf-8')
                server_sign = table.cell(line, 11).value
                if sing == function_sign and server_sign != '1':
                    return 1
                line += 1
        elif way == 'get_function_description':
            table = ex.sheet_by_name('function')  # 检查测试对应的功能是否存在
            while(1):
                function_sign = table.cell(line, 3).value.encode('utf-8')
                if sing == function_sign:
                    function_descripteion = table.cell(line, 4).value
                    return function_descripteion
                line += 1
    except IndexError:
        LoggUser.info('调用check_there函数。way：%s 没有查到 %s 功能信息' %(way, sing))
        return 0
    except Exception, e:
        logging.error('运行错误:%s' %e)
        return 0


# 删除文件
def file_processing(file_path, operation, target):
    LoggUser.debug('调用 file_processing 函数. file_path:%s operation:%s target:%s' % (file_path, operation, target))
    all_file = os.listdir(file_path)
    if operation == 'remove':
        if target == '*':
            for file_name in all_file:
                os.remove(file_path + '\\' + file_name)
                LoggUser.info('删除文件，文件路径为： ' + file_path + '\\' + file_name)
            return 1, ''
    elif operation == 'only_reserve':
        all_file.remove(target)
        for file_name in all_file:
            os.remove(file_path + '\\' + file_name)
            LoggUser.info('删除文件，文件路径为： ' + file_path + '\\' + file_name)
        return 1, ''

if __name__ == '__main__':
    mode = raw_input(hint)
    RunMode, CasePath, FilePath, FolderName = creation_method(mode)
    GlobalConfiguration.ConfigInit(ConfPath, LoggUser)   # 配置全局变量
    LoggUser.debug('RunMode:%s  CasePath:%s FilePath:%s FolderName:%s' % (RunMode, CasePath, FilePath, FolderName))
    if RunMode != 0:
        LoggUser.info('开始校验测试例的json数据正确性')
        if RunMode in ('1', '5'):                                    # 检查接口的列
            checklist = [5, 6, 8, 9, 13, 14, 17, 19, 21, 22, 23]    # 被检查的列
            result = Tools.check_json_data(LoggUser, CasePath, 'AllUint', 'interface', *checklist)
        elif RunMode in ('2', '6'):                                  # 检查功能的列
            checklist = [6, 7]    # 被检查的列
            result = Tools.check_json_data(LoggUser, CasePath, 'Function', 'function', *checklist)
            if result == {}:
                checklist = [4, 6, 8, 9, 10, 13, 14, 16, 17, 19, 20, 22, 23, 25, 26, 28, 29, 31, 32, 34, 35]
                result = Tools.check_json_data(LoggUser, CasePath, 'Function', 'data', *checklist)
        else:
            LoggUser.info('找不到检查的列数据')
            result = 'no run mode'
        if len(result) == 0 and int(mode) <= 4:
            if RunMode == '1':                                       # 创建接口文件
                ex = xlrd.open_workbook(CasePath)
                state, info = file_processing(FilePath, 'only_reserve', '__init__.py')
                if state == 1:
                    LoggUser.info('开始创建接口测试脚本')
                    parameter_a, parameter_b, table_line = new_file(FolderName, FilePath, ex)
                    if parameter_a != '0':
                        MakeTestCase(parameter_a, parameter_b, FilePath, ex, FolderName, line=table_line)
                        CreateMain(FolderName, FilePath)
                        LoggUser.info('创建接口测试脚本结束')
                    if len(error_info) > 0:
                        LoggUser.info('创建接口脚本错误信息:%s' % error_info)
            elif RunMode == '2':                                     # 创建功能测试文件
                ex = xlrd.open_workbook(CasePath)
                state, info = file_processing(FilePath, 'only_reserve', '__init__.py')
                if state == 1:
                    LoggUser.info('开始创建功能测试脚本')
                    parameter_a, parameter_b, table_line = new_file(FolderName, FilePath, ex)
                    LoggUser.debug('parameter_a:%s parameter_b:%s table_line:%s' % (parameter_a, parameter_b, table_line))
                    if parameter_a not in ('0', '00'):
                        MakeTestCase(parameter_a, parameter_b, FilePath, ex, FolderName, function_Content_case,
                                     line=table_line)
                        CreateMain(FolderName, FilePath)
                    elif parameter_a == '00':
                        error_info[table_line] = parameter_b
                    LoggUser.info('创建功能测试脚本结束')
                    if len(error_info) > 0:
                        for key, value in error_info.items():
                            LoggUser.info('创建功能测试脚本错误信息.第 %s 行测试用例，错误信息:%s' % (str(key), value))
            else:
                 LoggUser.info('此运行方式暂时不支持')
        elif len(result) == 0 and int(mode) >= 4:
            LoggUser.info('测试用例json校验通过')
        else:
            LoggUser.error('错误的json数据:%s' % result)
    else:
        LoggUser.info('不支持的运行模式:%s' % FilePath)