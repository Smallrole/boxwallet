#coding=utf-8
import os, sys
from Share import ConfigParser


def ConfigInit(config_file, log_user):
    try:
        log_user.debug('ConfigInit')
        global host_address
        global DB
        global e_mail
        global error_code
        global conf
        global current_path
        global file_path
        global login_session
        global jsessionid
        global cache_region
        conf = ConfigParser.ConfigParser()
        conf.read(config_file)
        host_address = dict(conf.items('service_address'))  # list->dict
        current_path = os.getcwd()
        # DB = dict(conf.items('DB'))    # list->dict
        e_mail = dict(conf.items('e_mail'))  # list->dict
        error_code = dict(conf.items('error_code'))
        file_path = None
        login_session = None
        jsessionid = None
        cache_region = {}
        log_user.debug('config_init_end')
    except Exception, e:
            log_user.error('error:%s'% e)


# 返回主机地址
def get_service_address():
    for key, value in host_address.items():
        host = host_address[key]
    return host


# 返回一个可用的拼接 如：online_bus/online_bus@172.30.0.155/szdev
def get_db(DBName):
    db_name = 'DB_' + DBName
    db = dict(conf.items(db_name))
    for k, v in db.items():
        k = str(k)
        v = str(v)
        if k == 'user':
            user = v
        elif k == 'password':
            pass_word = v
        elif k == 'host':
            host = v
        elif k == 'dbname':
            name = v
    db_name = user + '/' + pass_word + '@' + host + '/' + name
    return db_name


# 获取邮件配置信息
def get_e_mail():
    for key in e_mail.keys():
        if key == 'user':
            user = e_mail[key]
        elif key == 'password':
            password = e_mail[key]
        elif key == 'recipient':
            recipient = e_mail[key]
    return user, password, recipient


# 获取状态码描述
def get_error_code(status):
    for key in error_code.keys():
        if key == str(status):
            status_info = error_code[key]
            return status_info


# 获取路径信息：测试用例文件路径、测试报告、日志配置、配置(只适用于main调用)
def get_path_info(name, log):
    if name == 'AllUint':      # 接口测试用例
        file_path = current_path + '\TestCase\interface.xlsx'
        log.info('单接口文件路径:%s' % file_path)
        return file_path
    elif name == 'Function':          # 功能测试用例
        file_path = current_path + '\TestCase\\function.xlsx'
        log.info('功能测试文件路径:%s' % file_path)
        return file_path
    elif name == 'Pressure':          # 压力测试用例
        file_path = current_path + '\TestCase\pressure.xlsx'
        log.info('压力测试文件路径:%s' % file_path)
        return file_path
    elif name == 'AutoUi':                 # ui测试用例
        file_path = current_path +'\TestCase\ui.xlsx'
        log.info('Ui测试文件路径:%s' % file_path)
        return file_path
    elif name == 'TestReport':  # 测试报告
        file_path = current_path + '\TestResult\Result'
        log.info('测试报告存放路径:%s' % file_path)
        return file_path
    elif name == 'Logconfig':   # 日志配置
        file_path = current_path + '\Configuration\logging.conf'
        log.info('日志配置文件路径:%s' % file_path)
        return file_path
    elif name == 'infoconfig':  # 信息配置
        file_path = current_path + '\Configuration\TestConfiguration.cfg'
        log.info('测试信息配置文件路径:%s' % file_path)
        return file_path
    elif name in ('Interface_Report', 'Function_Report', 'Pressure_Report', 'Ui_Report'):  # 测试报告文件
        report = '\\' + name + '.html'
        file_path = current_path + '\TestResult\Result' + report
        log.info('测试报告文件路径:%s' % file_path)
        return file_path
    else:
        return 'error'           # 返回错误


# 获取缓存数据
def get_cache_region():
    return cache_region


# 保存缓存数据
def save_cache_region(cache):
    global cache_region
    cache_region = cache

if __name__ == '__main__':
    # ConfigInit(r'G:\Python\CashBox\CashBox\Configuration\TestConfiguration.cfg')
    ConfigInit(r'D:\Procedure\PycharmProjects\CashBox\Configuration\TestConfiguration.cfg')
    # print get_service_address()
    # print get_db('online_bus')
    # print get_e_mail()
    # print get_path_info('TestCase', 1)