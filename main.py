# coding=utf-8
import os
import logging
import logging.config

from Share import GlobalConfiguration

hint='''
***请选择运行脚本模式，请输入相应方式的序号，再按回车确定***
1.执行--> 单接口自动化 <--测试功能，脚本来源   TestScript->AllUint；
2.执行--> 功能接口自动化 <--测试功能，脚本来源 TestScript->Function；
3.执行--> 压力测试--脚本 <--脚本来源 TestScript->Pressure；
4.执行--> WebUi自动化测试 <--脚本，脚本来源  TestScript->AutoUi；

'''

CurrentPath=os.getcwd()                                        # 取当前的路径
loggconf_path=CurrentPath + '\Configuration\logging.conf'      # 加载配置日志的路径
conf_path=CurrentPath + '\Configuration\TestConfiguration.cfg' # 加载配置的路径

if __name__=='__main__':
    logging.config.fileConfig(loggconf_path)    # 加载配置
    root_logger = logging.getLogger('root')    # 使用打印日志用户
    GlobalConfiguration.ConfigInit(conf_path, root_logger)   # 配置全局变量
    mode=raw_input(hint)
    if mode=='1':
        logger_user = logging.getLogger('Example')  # 使用打印日志用户
        test_report='Interface_Report'              # 测试报告名
        logger_user.info('运行模式：单接口测试')
        from TestScript.AllUint import Perform_test_case
        Perform_test_case.perform_test_case(logger_user, test_report)
        logger_user.info('执行完成--单接口测试--')
    elif mode == '2':
        logger_user = logging.getLogger('function')  # 使用打印日志用户
        logger_user.info('运行模式：功能接口测试')
        test_report='Function_Report'                # 测试报告名
        from TestScript.Function import Perform_test_case
        Perform_test_case.perform_test_case(logger_user, test_report)
        logger_user.info('执行完成--功能接口测试--')
    elif mode == '3':
        root_logger.info('运行模式：压力测试')
        test_report='Pressure_Report'               # 测试报告名
        from TestScript.Pressure import Perform_test_case
        Perform_test_case.perform_test_case(root_logger, test_report)
        root_logger.info('执行完成--压力测试--')
    elif mode == '4':
        test_report='Ui_Report'                    # 测试报告名
        root_logger.info('运行模式：UI自动化测试')
        from TestScript.AutoUi import Perform_test_case
        Perform_test_case.perform_test_case(root_logger, test_report)
        root_logger.info('执行完成--UI自动化测试--')
    else:
        root_logger.info('没有此运行模式，请输入：1-4 范围的模式')