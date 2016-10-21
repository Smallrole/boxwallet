# coding=utf-8
import cx_Oracle
import sys
import time
reload(sys)
sys.setdefaultencoding("utf-8")   # 把 str 编码由 ascii 改为 utf8

from Share import requests
from Share import json
from Share import GlobalConfiguration
from Share.string import Template


# 脚本执行 参数：脚本内容，日志用户（要区分增删改查操作,类型是字典需要处理） 返回：执行状态（1成功） 数据，
def PerformDBASql(operation, log):
    try:
        results = {}
        for host in operation.keys():
            hostname = str(host)        # 返回一个可连接的信息
            log.info('使用  %s 用户连接数据库' % hostname)
            # 连接数据
            hostname = GlobalConfiguration.get_db(hostname)
            log.debug('DB_hostname:%s' % hostname)
            conn = cx_Oracle.connect(hostname)
            cursor = conn.cursor()
            second_step = operation[host]                   # 取内容
            for name, sql in sorted(second_step.items()):   # key进行排序升序
                analyze_sql = sql.split(' ')                 # 以空格分割
                if analyze_sql[0] == 'select':             # 获取操作方式
                    log.info('进行 %s 操作 SQL:%s' % (analyze_sql[0], sql))
                    sql = str(sql)
                    cursor.execute(sql)               # 执行SQL
                    one = cursor.fetchone()           # 执行SQL返回结果fetchone或fetchal
                    if one != None:
                        # log.debug('查询到的数据结果类型：%s' %type(one))
                        one = str(one[0])
                        one = ''.join(one)                 # 将结果转化为str,以‘，’分开
                        results[name] = one
                    else:
                        results[name] = 'null'
                    log.debug('查询 %s 数据，结果：%s' % (name, one))    # 返回查找后的数据
                elif analyze_sql[0] == 'update':                              # key进行排序升序
                    log.debug('进行 %s 操作 执行第 %s 条语句，SQL：%s' % (analyze_sql[0], name, sql))
                    # 执行SQL
                    sql = str(sql)
                    cursor.execute(sql)               # 执行SQL
                elif analyze_sql[0] == 'delete':
                    log.debug('进行 %s 操作 执行第 %s 条语句，SQL：%s' % (analyze_sql[0], name,sql))
                    # 执行SQL
                    sql = str(sql)
                    cursor.execute(sql)
                elif analyze_sql[0] == 'insert':
                    log.info('进行 %s 操作' % analyze_sql[0])
                    log.debug('执行第 %s 条语句，SQL：%s' % (name, sql))
                    # 执行SQL
                    sql = str(sql)
                    cursor.execute(sql)
                else:
                    log.info('暂不支持此类型操作,关闭数据连接 %s' % host)
                    cursor.close()
                    conn.close()
                    return 0, '暂不支持此类型操作'   # 暂不支持此类型操作
            log.info('从数据库里查询到的数据:\n %s' % json.dumps(results, ensure_ascii=False, indent=4))
            log.info('关闭  %s 连接数据库,提交数据' % host)
            cursor.close()
            conn.commit()
            conn.close()
        return 1, results
    except Exception, e:
        log.error('数据库操作异常信息：%s  ***如果提示内容为：ORA-00911: invalid character 则是sql语句后面加了“；”号 '% e)
        try:
            cursor.close()
            conn.rollback()
            conn.close()
        except Exception, db:
            return 0, 'Error:%s' % e  # 返回异常信息
        log.error('发生异常，进行数据回滚操作。关闭  %s 连接数据库' % host)  #rollback()
        return 0, 'Error:%s' % e                            # 返回异常信息


# 数据替换（后期增加日志）
def DataHandle(RqDateTpye, Data, Content, log):
    # log.debug('Content数据:\n%s  RqDateTpye:%s' %(json.dumps(Content, ensure_ascii=False, indent=4),RqDateTpye ))
    if str(type(Content)) != "<type 'dict'>":
        Content=eval(Content)       # 一个Bug后期处理
    Data = str(Data)
    if RqDateTpye == '1':          # 数据分离
        Data = Template(Data)
        Data = Data.safe_substitute(Content).encode("utf-8")
        return Data
    elif RqDateTpye == '2':   # 数据分离（查询数据库）
        for k, v in Content.items():
            Content[k] = v       # 数据库结果 增加执行数据库存语句
        Data = Template(Data)
        Data = Data.safe_substitute(Content).encode("utf-8")
        # log.debug('返回Content数据:\n%s' %json.dumps(Data, ensure_ascii=False, indent=4))
        return Data
    elif RqDateTpye == '0' or RqDateTpye == '':
        return 'Default Values'
    else:
        return 'There is no this type'


# 处理长度类型  返回 状态 信息1 信息2
def JsonLen(DataA, DataB, loguser):
    loguser.debug('调用JsonLen函数 开始判断长度，响应返回数据:%s  预期数据:%s' % (DataA, DataB))
    for k, v in DataB.items():
        ReverseLen = len(str(DataA))       # 返回数据的长度
        RefrenceLen = int(DataB[k])        # 参考数据的长度
        loguser.debug('实返回数据长度：%s  预期数据长度：%s' % (ReverseLen, RefrenceLen))
        if ReverseLen != RefrenceLen:
            return 0, ReverseLen, RefrenceLen  # 失败
        return 1, None, None                   # 成功


# 处理列表-列表类型数据 返回参数a 参考参数 b 日志  返回状态1或0 信息1 信息2
def JsonList(ListA, ListB, loguser):
    try:
        loguser.debug('调用JsonList函数 实响应数据: %s 预期数据:%s' % (ListA, ListB))
        for list1 in ListB:                   # 取参考数据列表中的一组字典数据
            loguser.debug('预期数据 %s' % list1)
            for information in ListA:
                loguser.debug('预期数据 %s' % information)
                if information != list1:      # 如果列表的中字典不相同，要增加轮查
                    Signs = 0                 # 默认的标志
                    for repetition in ListA:
                        if repetition == list1:
                            loguser.debug('在轮查中找到并退出.重试的数据：%s 预期数据：%s' % (repetition, list1))
                            Signs = 1                  # 找到的标志
                            break                     # 退出当次循环
                    if Signs == 1:
                        break                         # 退出当次循环
                    for k, v in information.items():
                        loguser.debug('返回数据key:%s　V：%s  预期数据 key:%s V:%s' % (k, v, k, list1 [k] ))
                        if k in list1 and information[k] != list1[k]:
                            if len(list1) >= 3:       # 字典长度大小决定定位的精度
                                CombinationA = {}
                                CombinationB = {}
                                CombinationA[k] = v
                                CombinationB[k] = list1[k]
                                loguser.debug('匹配失败的信息.返回的数据:%s　预期的数据：%s '  % (CombinationA, CombinationB))
                                return 0, CombinationA, CombinationB
                            #elif str(type(list1[k])) == "<type 'dict'>" and str(type(v)) != "<type 'dict'>":    # 动态数据对比数据长度
                            elif isinstance(list1[k], dict) and not isinstance(v, dict):
                                State, InfoA, InfoB = JsonLen(v, list1[k], loguser)
                                if State == 0:
                                    return 0, InfoA, InfoB
                                break
                            else:
                                loguser.debug('匹配失败的信息，返回的数据:%s  预期的数据:%s '% (information, list1))
                                return 0, information, list1
                        elif k not in list1:
                            loguser.debug('返回数据里找不到Key:%s '% k)
                            return 0, 'NoFoundKey', k
                elif information == list1:
                    loguser.debug('匹对成功退出当次循环')
                    break
        return 1, None, None  # 数据正确
    except KeyError, e:
        loguser.error('发生异常！！！错误信息%s' % e)
        return 0, 'KeyError', e


# 处理字典类型数据  字典1，字典2 返回状态1或0 信息1 信息2
def JsonDict(DictA, DictB, loguser):
    try:
        loguser.debug('调用 JsonDict 函数 DictA: %s DictB:%s' % (DictA, DictB))
        for kb, vb in DictB.items():  # 取参考数据中的一组
            loguser.debug('实返回 key:%s 实值：%s 预期key: %s 预期值:%s' % (kb,DictA[kb], kb, vb))
            # if str(type(vb))== "<type 'dict'>" and str(type(DictA[kb]))== "<type 'dict'>":  # 如果参考数据是字典类型
            if isinstance(vb, dict) and isinstance(DictA[kb], dict):
                DictA1 = DictA[kb]                # 取返回对应的字典
                DictB1 = vb                       # 取参考对应的字典
                if DictB1 != DictA1:
                    for k, v in DictA1.items():
                        loguser.debug('返回key: %s 值:%s' % (k, v))
                        if k in DictB1 and DictA1[k] != DictB1[k]:
                            loguser.debug('匹对失败。实返回key: %s 值:%s  预期Key%s  预期值：%s' % (k, v, k ,DictB1[k]))
                            return 0, DictA1[k], DictB1[k]  # 返回失败状态0 返回参数 参考数
                        elif k not in DictB1:
                            loguser.debug('在返回数据中找不到key: %s ' % k)
                            return 0, 'NoFoundKey', k
            # elif str(type(vb))== "<type 'list'>" and str(type(DictA[kb]))== "<type 'list'>":      # 调用列表匹配
            elif isinstance(vb, list) and isinstance(DictA[kb], list):
                JsonList(DictA[kb], vb, loguser)

            # elif str(type(vb)) == "<type 'dict'>" and str(type(DictA[kb])) != "<type 'dict'>":    # 动态数据对比数据长度
            elif isinstance(vb, dict) and not isinstance(DictA[kb], dict):
                State, InfoA, InfoB = JsonLen(DictA[kb], vb, loguser)
                if State == 0:
                    return 0, InfoA, InfoB
                break
            elif kb in DictA and DictA[kb] != vb:
                loguser.debug('匹对失败。返回key: %s 值:%s  参考Key：%s  值：%s' % (kb, DictA[kb], kb ,vb))
                return 0, DictA[kb], DictB[kb]    # 返回失败状态0 返回参数 参考数
            elif kb not in DictA:
                loguser.debug('在返回数据中找不到key: %s ' % kb)
                return 0, 'NoFoundKey', k
        return 1, None, None   # 返回成功
    except KeyError, e:
        loguser.error('错误信息%s' % e)
        return 0, 'KeyError', e


# 函数功能 数据包含方式对比   （后期增加日志） A返回数据 B为参考数据
def cmp_json(jsonA, jsonB, loguser):
    # try:
        for key, value in jsonB.items():
            # loguser.info('Key:%s  value:%s' %(key, value))
            if (key in jsonA):
                # if str(type(jsonB[key])) == "<type 'list'>" and str(type(jsonA[key])) == "<type 'list'>":  # 列表类型数据
                if isinstance(jsonB[key], list) and isinstance(jsonA[key], list):
                    loguser.debug('开始对比--list--类型的数据')
                    State, InfoA, InfoB = JsonList(jsonA[key], jsonB[key], loguser)
                    if State == 0:
                        return State, InfoA, InfoB
                    loguser.debug('结束对比--list--类型的数据')
                    continue
                # elif str(type(jsonB[key])) == "<type 'dict'>" and str(type(jsonA[key])) == "<type 'dict'>":  # 字典类型
                elif isinstance(jsonB[key], dict) and isinstance(jsonA[key], dict):
                    loguser.debug('开始对比--dict--类型的数据')
                    State, InfoA, InfoB= JsonDict(jsonA[key], jsonB[key], loguser)
                    if State== 0:
                        return State, InfoA, InfoB
                    loguser.debug('结束对比--dict--类型的数据')
                    continue
                # elif str(type(jsonB[key])) == "<type 'dict'>" and str(type(jsonA[key])) != "<type 'dict'>":  # 只判断长度类型
                elif isinstance(jsonB[key], dict) and not isinstance(jsonA[key], dict):
                    loguser.debug('开始对比--动态数据长度判断')
                    State, InfoA, InfoB = JsonLen(jsonA[key], jsonB[key], loguser)
                    loguser.debug('结束对比--动态数据长度判断')
                    if State == 0:
                        return 0, InfoA, InfoB
                    continue
                elif key in jsonA and jsonA[key] != value:               # 普通数据判断
                    loguser.info('实返回数据：%s:%s 预期返回数据：%s:%s' % (key, jsonA[key], key, value))
                    return 0, jsonA[key], value
                loguser.debug('普通类型判断 返回数据：KEY:%s  VALUE:%s   参考数据：KEY:%s  VALUE:%s'
                              % (key, str(jsonA[key]), key, value))
            else:                            # key不在返回数据里
                loguser.debug('在响应数据里找不到 %s Key' %key)
                return 0, '%s key not in comparison' % str(key), ''      # 返回失败原因
        return '1', None, None  # pass

"""except KeyError,e:
        loguser.error('错误信息%s' %e)
        return 'KeyError:',e
"""


# 预置功能函数
def preset(test_host, table, row, log_user):
        log_user.info('调用 preset 函数 请求接前执行第 %s 条用例' % row)
        row = int(row)
        state, info_a, info_b = request_core(table, row, test_host, log_user)   # state ：0 失败  1 成功  2其它错误
        set_down = str(table.cell(row, 18).value)              # 请求后还原标志位
        if set_down == '1':
            log_user.debug('开始执行执行还原脚本')
            set_down_sql = eval(table.cell(row, 19).value.encode("utf-8"))
            date_state, content = PerformDBASql(set_down_sql, log_user)
            if date_state == 0:
                return 'database operation Fail', ''
        if state == 0:
            log_user.debug('失败-状态：%s  内容_A：%s  内容_b：%s' % (state, info_a, info_b))
            return 0, info_a, info_b                   # 返回失败
        elif state == 1:
            log_user.debug('状态：%s  内容_A：%s  内容_b：%s' % (state, info_a, info_b))
            return 1, info_a, info_b                   # 返回成功
        elif state == 2:
            log_user.debug('状态：%s  内容_A：%s  内容_b：%s' % (state, info_a, info_b))
            return 0, info_a, info_b                   # 返回失败
        else:
            return 0, info_a, info_b                   # 返回失败


# 功能预执行函数
def function_preset(table, test_case_row, test_host, loger_user):
    loger_user.info('调用 function_preset 函数  table%s： test_case_row：%s test_host：%s' % (table, test_case_row, test_host))
    ex_table = table.sheet_by_name('data')
    test_case_valid = str(ex_table.cell(test_case_row, 11).value)    # 读取测试用例生效标志
    if test_case_valid == '1':
        loger_user.info('无效的测试用例')
        return 0, 'Invalid test case'
    function = ex_table.cell(test_case_row, 2).value.encode("utf-8")  # 读取测试用例的功能标志
    state, data_info = get_line_number(function, table, loger_user)
    loger_user.info('state:%s data_info:%s' % (state, data_info))
    if state == 0:
        return state, data_info
    for line_number in data_info:
        state, interface, info = function_core(table, line_number, test_case_row, test_host, loger_user)
        loger_user.debug('状态：%s  interface：%s  info：%s' % (state, interface, info))
        if state != 1 or state == 3:
            return state, interface, info


# 获取功能对应的列号
def get_line_number(sign, ex, loger_user):
    try:
        loger_user.debug('调用 get_line_number 函数。sing：%s  ex:%s' % (sign, ex))
        line = 1
        table = ex.sheet_by_name('function')  # 检查测试对应的功能是否存在
        function_line_num = []
        while(1):
            function_sign = table.cell(line, 3).value.encode('utf-8')
            server_sign = table.cell(line, 14).value
            if sign == function_sign and server_sign != '0' and line not in function_line_num:
                loger_user.debug('function_line_num:%s' % function_line_num)
                function_line_num.append(line)
            line += 1
    except IndexError:
        if len(function_line_num) > 0:
            return 1, function_line_num
        else:
            return 0, '找不到功能或功能都不生效'
    except Exception, e:
        loger_user.error('运行错误:%s' %e)
        return 0, e


# 输入cookies 信息  返回重构后的cookies
def reset_cookies(cookie):
    cookies = {}
    if GlobalConfiguration.jsessionid == None:
        return 0
    # cookies['JSESSIONID']=GlobalConfiguration.jsessionid 双重jsessionid
    for k, v in cookie.items():
        cookies[k] = v
    return cookies


# 请求处理 读取表 读取的行 主机地址 日志用户  返回 1，参数1 参数2
def request_core(ex_table, table_line, test_host, log_user):
        table_line = int(table_line)
        set_out = str(ex_table.cell(table_line, 16).value)                         # 请求前是否需要初始化标志
        interface_explain = ex_table.cell(table_line, 3).value.encode("utf-8")     # 读取测试用例名
        test_name = ex_table.cell(table_line, 3).value.encode("utf-8")             # 读取测试用例名
        log_user.info('状态：准备 被测试接口名称：%s  执行第 %s 条 %s 测试用例' % (interface_explain, table_line, test_name))
        if set_out == '1':
            set_out_sql = eval(ex_table.cell(table_line, 17).value.encode("utf-8"))  # 获取sql
            log_user.info('开始执行初始化脚本')
            state, content = PerformDBASql(set_out_sql, log_user)
            if state == 0:
                log_user.info('初始化脚本执行失败  状态：%s  返回内容：%s' % (state,content))
                return 0, 'database operation Faile', ''
            log_user.info('初始化脚本执行成功 状态：%s  返回内容：%s' % (state,content))
        login_sign = ex_table.cell(table_line, 24).value.encode("utf-8")    # 获取是否需要登录标志
        execute_row = ex_table.cell(table_line, 25).value.encode("utf-8")   # 获取用例行数
        log_user.info('login_session: %s' % GlobalConfiguration.login_session)
        if login_sign == '1':
            log_user.info('开始执行预置接口')
            state, reinfo_a, reinfo_b = preset(test_host, ex_table, execute_row, log_user)
            if state == 1:
                log_user.info('执行成功')
            else:
                log_user.info('执行失败')
                return state, reinfo_a, reinfo_b
        if GlobalConfiguration.login_session != None and login_sign != '2':
            role=GlobalConfiguration.login_session                            # 使用上一次session
        elif login_sign == '2':                                               # 当为2时不带session
            role = requests
        else:
            role = requests.session()                                         # 设置为带session
        interface_name = ex_table.cell(table_line, 1).value.encode("utf-8")  # 读取请求接口
        test_host = test_host + '/' + interface_name                         # 拼接请求地址
        way = ex_table.cell(table_line, 4).value.encode('utf-8')             # 请求方式
        reset = ex_table.cell(table_line, 6).value.encode('utf-8')           # cookies 信息
        log_user.info('状态：开始执行 被测试接口名称：%s  执行第 %s 条 %s 测试用例' % (interface_explain, table_line, test_name))
        if way == 'POST':
            log_user.info('------ 开始 Post 请求 ------ 请求地址:%s' % test_host)
            rq_header = eval(ex_table.cell(table_line, 5).value.encode("utf-8"))        # 请求头
            data_type = str(ex_table.cell(table_line, 7).value)                         # 数据并接方式
            if data_type == '1' or data_type == '2':
                payload = ex_table.cell(table_line, 8).value.encode("utf-8")            # 获取请求内容
            else:
                payload = json.loads(ex_table.cell(table_line, 8).value.encode("utf-8"))  # 转为字典格式  请求内容
            if data_type == '1':
                content = eval(ex_table.cell(table_line, 9).value.encode("utf-8"))  # 取替换参数
                content = merge_data(content, GlobalConfiguration.get_cache_region())
                payload = json.loads(DataHandle(data_type, payload, content, log_user))
            elif data_type == '2':
                content = eval(ex_table.cell(table_line, 9).value.encode("utf-8"))  # 取替换参数
                log_user.debug("查询SQL:%s " % content)
                state, content = PerformDBASql(content, log_user)                      # 执行sql，查询数据
                if state != 0:
                    content = merge_data(content, GlobalConfiguration.get_cache_region())
                    payload = json.loads(DataHandle(data_type, payload, content, log_user))      # 替换数据
                else:
                    log_user.info("** 数据库操作失败:%s  **" % content)
                    return 0, 'database operation Fail', ''
            elif data_type not in ('0', ''):
                return 0, 'POST Parametric approach does not support', ''
            log_user.info('数据参数化方式：%s 请求数据:\n%s' % (data_type, json.dumps(payload, ensure_ascii=False, indent=4)))
            # 带头部信息
            if rq_header != '':
                header = rq_header
                for k in header.keys():
                    if str(k) == 'Content-Type':
                        if header[k].count('x-www-form-urlencoded') > 0:  # 窗体数据被编码为名称/值对，这是标准的编码格式
                            if reset == '' or reset != '1':
                                rq = role.post(test_host, data=payload, headers=header, timeout=15)
                            else:
                                cookies=reset_cookies(eval(reset))
                                rq = role.post(test_host, data=payload, headers=header, cookies=cookies, timeout=15)
                        elif header[k].count('json') > 0:                 # 数据为键值方式
                            if reset == ''or reset != '1':
                                rq = role.post(test_host, data=json.dumps(payload), headers=header, timeout=15)
                            else:
                                cookies = reset_cookies(eval(reset))
                                log_user.info('cookies:\n%s' % cookies)
                                rq = role.post(test_host, data=json.dumps(payload), headers=header, cookies=cookies, timeout=15)
                        elif header[k].count('form-data') > 0:            # 窗体数据被编码为一条消息
                            if reset == ''or reset != '1':
                                rq = role.post(test_host, files=payload, headers=header, timeout=15)
                            else:
                                cookies = reset_cookies(eval(reset))
                                log_user.info('cookies:\n%s' % cookies)
                                rq = role.post(test_host, files=payload, headers=header, cookies=cookies, timeout=15)
                        break
                    else:
                        log_user.info("不支持的请求头方式：%s" % header)
            else:
                if reset == ''or reset != '1':
                    rq = role.post(test_host, data=json.dumps(payload), timeout=15)
                else:
                    cookies = reset_cookies(eval(reset))
                    log_user.info('cookies:\n%s' % cookies)
                    rq = role.post(test_host, data=json.dumps(payload), cookies=cookies, timeout=15)
        elif way == 'GET':
            log_user.info('------开始 Get 的请求------ 主机地址:%s' % test_host)
            data_type = str(ex_table.cell(table_line, 7).value)                     # 数据并接方式
            if data_type == '1' or data_type == '2':
                payload = ex_table.cell(table_line, 8).value.encode("utf-8")         # 获取请求内容
            else:
                payload = json.loads(ex_table.cell(table_line, 8).value.encode("utf-8"))   # 转为字典格式  请求内容
            if data_type == '1':
                content = json.loads(ex_table.cell(table_line, 9).value.encode("utf-8"))  # 取替换参数
                payload = json.loads(DataHandle(data_type, payload, content, log_user))
            elif data_type == '2':
                content = eval(ex_table.cell(table_line, 9).value.encode("utf-8"))     # 取替换参数
                state, content = PerformDBASql(content, log_user)                      # 执行sql，查询数据
                if state != 0:
                    content = merge_data(content, GlobalConfiguration.get_cache_region())
                    payload=json.loads(DataHandle(data_type, payload, content, log_user))            # 替换数据
                else:
                    log_user.info("**数据库操作失败**:%s" % content)
                    return 0, 'database operation Fail', ''
            else:
                return 0, 'Parametric approach does not support', ''
            rq = role.get(test_host, params=payload)
            log_user.info('GET请求URL:%s' % rq.url)
        else:
            log_user.info('不支持的请求方式')
            return 2, 'Request form mistake', 'POST OR GET'
        rq.raise_for_status()                                         # 作用状态非200是引发异常
        if 'JSESSIONID' in rq.cookies:
            GlobalConfiguration.jsessionid = rq.cookies['JSESSIONID']  # 保存sessionid
        if GlobalConfiguration.login_session == None:
            GlobalConfiguration.login_session = role                    # 返回维持session的会话
        buffer_sign = str(ex_table.cell(table_line, 15).value)
        if buffer_sign != '':
            save_list = eval(ex_table.cell(table_line, 15).value.encode("utf-8"))
            log_user.debug('需要缓存的数据:%s  数据类型:%s' % (save_list, type(save_list)))
            memory = GlobalConfiguration.get_cache_region()
            log_user.debug('请求数据:%s' % payload)
            memory = save_data(save_list, payload, memory)
            GlobalConfiguration.save_cache_region(memory)
            log_user.info('已缓存的数据:%s' % memory)
            memory = save_data(save_list, rq.json(), memory)
            GlobalConfiguration.save_cache_region(memory)
            log_user.info('已缓存的数据:%s' % memory)
        log_user.info('响应返回数据:\n%s' % json.dumps(rq.json(), ensure_ascii=False, indent=4))
        delay = ex_table.cell(table_line, 11).value              # 读取延时
        if delay != '':
            delay = int(delay)
            log_user.info('等待后台处理，延时:%s S' % delay)
            time.sleep(delay)
        log_user.info('------分析响应结果 ------')
        criteria = ex_table.cell(table_line, 10).value.encode("utf-8")                # 断言方式
        data_type = str(ex_table.cell(table_line, 12).value)                          # 返回数据方式
        check_data_sign = str(ex_table.cell(table_line, 20).value)                    # 检查数据标记
        if data_type == '0' or data_type == '':
            comparison = json.loads(ex_table.cell(table_line, 13).value.encode("utf-8"))  # 对比结果
        elif data_type == '1':
            comparison = ex_table.cell(table_line, 13).value.encode("utf-8")          # 读取参数化数据
            content = ex_table.cell(table_line, 14).value.encode('utf-8')             # 读取替换数据
            comparison = eval(DataHandle(data_type, comparison, content, log_user))
        elif data_type == '2':
            comparison = ex_table.cell(table_line, 13).value.encode("utf-8")          # 读取参数化数据
            content = eval(ex_table.cell(table_line, 14).value.encode('utf-8'))
            state, content = PerformDBASql(content, log_user)                           # 执行sql，查询数据
            if state != 0:
                comparison = eval((DataHandle(data_type, comparison, content, log_user)).encode('utf-8'))    # 替换数据
                log_user.debug('替换后的数据:\n%s' % json.dumps(comparison, ensure_ascii=False, indent=4))
            else:
                log_user.info("数据库存操作失败:%s  **" % content)
                return 0, 'database operation Fail', ''
        elif data_type not in ('0', ''):
            return 0, 'Get Parametric approach does not support', ''
        if check_data_sign == '1':
            log_user.info("开始数据数据对比")
            data_comparison = eval(ex_table.cell(table_line, 21).value.encode("utf-8"))    # 参考数据
            database_data = eval(ex_table.cell(table_line, 22).value.encode("utf-8"))      # 数据库数据
            sql_statements = eval(ex_table.cell(table_line, 23).value.encode("utf-8"))     # 查询语句
            state, content = PerformDBASql(sql_statements, log_user)                       # 执行sql，查询数据
            if state != 0:
                database_data = eval(DataHandle('2', database_data, content, log_user))               # 替换数据
                log_user.info("** database Read the complete **")
            else:
                log_user.info("数据操作失败:%s  **" % content)
                return 0, 'database operation Fail', ''
        response_time = rq.elapsed.microseconds / 1000                                    # 响应时间
        status_info = GlobalConfiguration.get_error_code(rq.status_code).encode("utf-8")  # 获取状态码描述
        log_user.info("请求响应时间:%dms 响应状态:%d %s cookies: %s  sessionid: %s"
                      % (response_time, rq.status_code, status_info, rq.cookies, GlobalConfiguration.jsessionid))
        if criteria == 'AssertEqual':                                                  # 断言方式相等
            log_user.info('数据断言方式：AssertEqual  参考响应数据:\n%s' % json.dumps(comparison, ensure_ascii=False, indent=4))
            if comparison == rq.json() and check_data_sign == '1':
                log_user.info('数据库中的数据:\n%s' % json.dumps(database_data, ensure_ascii=False, indent=4))
                log_user.info('参考的数据:\n%s' % json.dumps(data_comparison, ensure_ascii=False, indent=4))
                result, info_a, info_b=cmp_json(database_data, data_comparison, log_user)
                if result != '1':                                           # 调用断言方式 返回1 正常  ，非1不正确
                        log_user.info('数据库数据对比失败 数据库的数据:\nResponse: %s 参考的数据 :%s' % (str(info_a), str(info_b)))
                        log_user.info('测试失败')
                        return 0, info_a, info_b
                else:
                        log_user.info('测试通过')
                        return 1, result, '1'
            elif comparison == rq.json():
                return 1, comparison, rq.json()                  # 返回成功
            else:
                return 0, comparison, rq.json()                  # 增加数据的定位
        elif criteria == 'AssertNotEqual':                                  # 断言方式不相等
            log_user.info('数据断言方式： AssertNotEqual  参考响应数据:\n%s' % json.dumps(comparison, ensure_ascii=False, indent=4))
            if comparison != rq.json() and check_data_sign == '1':
                log_user.info('数据库的数据:\n%s' % json.dumps(database_data, ensure_ascii=False, indent=4))
                log_user.info('参考的数据库数据:\n%s' % json.dumps(data_comparison, ensure_ascii=False, indent=4))
                result, info_a, info_b=cmp_json(database_data, data_comparison, log_user)
                if result != '1':                                           # 调用断言方式 返回1 正常  ，非1不正确
                        log_user.info('数据库数据对比失败  数据库数据: %s 参考数据 :%s' % (str(info_a), str(info_b)))
                        log_user.info('测试失败')
                        return 0, info_a, info_b
                else:
                        log_user.info('测试通过')
                        return 1, result, '1'
            elif comparison != rq.text:
                return 1, comparison, rq.json()                 # 返回成功
            else:
                return 0, comparison, rq.json()                 # 返回失败
        elif criteria == 'Contain':                                       # 断言方式包含
            rq=json.loads(rq.text)                                         # 转为字典数据
            log_user.info('预期响应数据:\n%s' % json.dumps(comparison, ensure_ascii=False, indent=4))
            result, info_a, info_b = cmp_json(rq, comparison, log_user)    # 对比数据
            if result != '1':                                            # 调用断言方式 返回1 正常  ，非1不正确
                log_user.info('响应数据匹配失败:\n响应数据: %s 参考响应数据 :%s' % (str(info_a), str(info_b)))
                log_user.info('响应数据匹配失败，测试失败')
                return 0, info_a, info_b
            else:
                log_user.info('响应数据匹配通过')
                if check_data_sign == '1':
                    log_user.info('数据库数据:\n%s' % json.dumps(database_data, ensure_ascii=False, indent=4))
                    log_user.info('数据库参考数据:\n%s' % json.dumps(data_comparison, ensure_ascii=False, indent=4))
                    result, info_a, info_b = cmp_json(database_data, data_comparison, log_user)
                    if result != '1':                                           # 调用断言方式 返回1 正常  ，非1不正确
                        log_user.info('数据库数据匹配失败，测试失败')
                        return 0, info_a, info_b
                    else:
                        log_user.info('测试通过')
                        return 1, result, '1'
                else:
                    log_user.info('测试通过')
                    return 1, result, '1'
            log_user.info('------ 数据分析结束 ------')
        else:
            log_user.info('不支持的断言方式')
            return 2, 'Judgment way wrong, not Contain or AssertNotEqual or AssertEqual', ''


# 功能测试 读取表 读取的行 主机地址 日志用户  返回 1，参数1 参数2
def function_core(ex_table, table_line, test_case, test_host, log_user):
        log_user.debug('ex_table:%s table_line:%s test_case:%s test_host:%s' % (ex_table, table_line, test_case, test_host))
        ex_object = ex_table
        test_case = int(test_case)
        ex_table = ex_object.sheet_by_name('data')
        test_case_name = ex_table.cell(test_case, 1).value.encode("utf-8")     # 读取测试用例标题
        ex_table = ex_object.sheet_by_name('function')
        test_name = ex_table.cell(table_line, 2).value.encode("utf-8")     # 读取接口名字
        log_user.info('准备执行： %s 用例 接口序号： %s  接口名称： %s' % (test_case_name, table_line, test_name))
        login_sign = ex_table.cell(table_line, 13).value.encode("utf-8")   # 获取是否需要登录标志
        execute_row = ex_table.cell(table_line, 14).value.encode("utf-8")   # 获取用例行数
        log_user.info('login_session: %s' % GlobalConfiguration.login_session)
        if login_sign == '1':
            log_user.info('开始执行预置接口')
            state, reinfo_a, reinfo_b = preset(test_host, ex_table, execute_row, log_user)
            if state == 1:
                log_user.info('执行成功')
            else:
                log_user.info('执行失败')
                return state, reinfo_a, reinfo_b
        if GlobalConfiguration.login_session != None and login_sign != '2':
            role = GlobalConfiguration.login_session                            # 使用上一次session
        elif login_sign == '2':                                                 # 当为2时不带session
            role = requests
        else:
            role = requests.session()                                         # 设置为带session
        interface_name = ex_table.cell(table_line, 1).value.encode("utf-8")  # 读取请求接口
        test_host = test_host + '/' + interface_name                         # 拼接请求地址
        way = ex_table.cell(table_line, 5).value.encode('utf-8')             # 请求方式
        reset = ex_table.cell(table_line, 7).value.encode('utf-8')           # cookies 信息
        log_user.info('开始执行： %s 用例  接口序号：%s  接口名称：%s' % (test_case_name, table_line, test_name))
        if way == 'POST':
            log_user.info('------ 开始 Post 请求 ------ 请求地址:%s' % test_host)
            rq_header = eval(ex_table.cell(table_line, 6).value.encode("utf-8"))        # 请求头
            requests_line_num = int(ex_table.cell(table_line, 8).value)                 # 读取请求数据行号
            requests_line_num = 15 + ((requests_line_num - 1) * 6)
            response_line_num = int(ex_table.cell(table_line, 11).value)                # 读取响应数据行号
            response_line_num = 18 + ((response_line_num - 1) * 6)
            ex_table = ex_object.sheet_by_name('data')                                  # 切换表格
            flow_sign = str(ex_table.cell(test_case, 12).value)
            data_type = str(ex_table.cell(test_case, (requests_line_num - 1)).value)                         # 数据并接方式
            log_user.debug('requests_line_num:%s response_line_num:%s  data_type:%s  flow_sign:%s'
                           % (str(requests_line_num), str(response_line_num), data_type, flow_sign))
            if data_type == '1' or data_type == '2':
                payload = ex_table.cell(test_case, requests_line_num).value.encode("utf-8")        # 获取请求内容
            else:
                payload = ex_table.cell(test_case, requests_line_num).value.encode("utf-8")
                if len(payload) > 0:
                    payload = json.loads(payload)  # 转为字典格式  请求内容
            if data_type == '1':
                content = eval(ex_table.cell(test_case, (requests_line_num + 1)).value.encode("utf-8"))  # 取替换参数
                content = merge_data(content, GlobalConfiguration.get_cache_region())
                payload = json.loads(DataHandle(data_type, payload, content, log_user))
            elif data_type == '2':
                log_user.debug("table_line:%s   requests_line_num :%s" % (test_case, requests_line_num))
                content = eval(ex_table.cell(test_case, (requests_line_num + 1)).value.encode("utf-8"))  # 取替换参数
                log_user.debug("查询SQL:%s " % content)
                state, content = PerformDBASql(content, log_user)                      # 执行sql，查询数据
                if state != 0:
                    content = merge_data(content, GlobalConfiguration.get_cache_region())
                    payload = json.loads(DataHandle(data_type, payload, content, log_user))      # 替换数据
                else:
                    log_user.info("** 数据库操作失败:%s  **" % content)
                    return 2, '', 'database operation Fail'
            elif data_type not in ('0', ''):
                return 2, '', 'POST Parametric approach does not support'
            log_user.info('数据参数化方式：%s 请求数据:\n%s' % (data_type, json.dumps(payload, ensure_ascii=False, indent=4)))
            if flow_sign == '1' and len(payload) == 0:  # 代表该条测试用例流程是不完整的
                log_user.info('第 %s 测试用例是不完整流程，不执行 %s 接口，执行测试例结束' % (table_line, test_name))
                return 3, 1, 1
            # 带头部信息
            if rq_header != '':
                if 'Content-Type' in rq_header:
                    if rq_header['Content-Type'].count('x-www-form-urlencoded') > 0:  # 窗体数据被编码为名称/值对，这是标准的编码格式
                        log_user.info('数据编码方式:x-www-form-urlencoded')
                        if reset == '' or reset != '1':
                            rq = role.post(test_host, data=payload, headers=rq_header, timeout=15)
                        else:
                            cookies=reset_cookies(eval(reset))
                            rq = role.post(test_host, data=payload, headers=rq_header, cookies=cookies, timeout=15)
                    elif rq_header['Content-Type'].count('json') > 0:                 # 数据为键值方式
                        log_user.info('数据编码方式:json')
                        if reset == ''or reset != '1':
                            rq = role.post(test_host, data=json.dumps(payload), headers=rq_header, timeout=15)
                        else:
                            cookies = reset_cookies(eval(reset))
                            log_user.info('cookies:\n%s' % cookies)
                            rq = role.post(test_host, data=json.dumps(payload), headers=rq_header, cookies=cookies, timeout=15)
                    elif rq_header['Content-Type'].count('form-data') > 0:            # 窗体数据被编码为一条消息
                        log_user.info('数据编码方式:form-data')
                        if reset == ''or reset != '1':
                            rq = role.post(test_host, files=payload, headers=rq_header, timeout=15)
                        else:
                            cookies = reset_cookies(eval(reset))
                            log_user.info('cookies:\n%s' % cookies)
                            rq = role.post(test_host, files=payload, headers=rq_header, cookies=cookies, timeout=15)
                else:
                    log_user.info('请求头数据数据方式错误，请选择：form-data 或 json 或 x-www-form-urlencoded')
            else:
                if reset == ''or reset != '1':
                    rq = role.post(test_host, data=json.dumps(payload), timeout=15)
                else:
                    cookies = reset_cookies(eval(reset))
                    log_user.info('cookies:\n%s' % cookies)
                    rq = role.post(test_host, data=json.dumps(payload), cookies=cookies, timeout=15)
        elif way == 'GET':
            ex_table = ex_object.sheet_by_name('function')
            requests_line_num = int(ex_table.cell(table_line, 8).value)
            response_line_num = int(ex_table.cell(table_line, 11).value)
            ex_table = ex_object.sheet_by_name('data')
            log_user.info('------开始 Get 的请求------ 主机地址:%s' % test_host)
            data_type = str(ex_table.cell(test_case, (requests_line_num - 1)).value)               # 数据并接方式
            ex_table = ex_object.sheet_by_name('data')
            if data_type == '1' or data_type == '2':
                payload = ex_table.cell(test_case, requests_line_num).value.encode("utf-8")         # 获取请求内容
            else:
                payload = eval(ex_table.cell(test_case, requests_line_num).value.encode("utf-8"))   # 转为字典格式  请求内容
            if data_type == '1':
                content = eval(ex_table.cell(test_case, (requests_line_num + 1)).value.encode("utf-8"))  # 取替换参数
                payload = eval(DataHandle(data_type, payload, content, log_user))
            elif data_type == '2':
                content = eval(ex_table.cell(test_case, (requests_line_num + 1)).value.encode("utf-8"))  # 取替换参数
                state, content = PerformDBASql(content, log_user)                      # 执行sql，查询数据
                if state != 0:
                    content = merge_data(content, GlobalConfiguration.get_cache_region())
                    payload = eval(DataHandle(data_type, payload, content, log_user))            # 替换数据
                else:
                    log_user.info("**数据库操作失败**:%s" % content)
                    return 2, '', 'database operation Fail'
            else:
                return 2, '', 'Parametric approach does not support'
            rq = role.get(test_host, params=payload)
            log_user.info('GET请求URL:%s' % rq.url)
        else:
            log_user.info('不支持的请求方式')
            return 2, '', ' Request form mistake POST OR GET'
        rq.raise_for_status()                                         # 作用状态非200是引发异常
        if 'JSESSIONID' in rq.cookies:
            GlobalConfiguration.jsessionid = rq.cookies['JSESSIONID']  # 保存sessionid
        if GlobalConfiguration.login_session == None:
            GlobalConfiguration.login_session = role                    # 返回维持session的会话
        ex_table = ex_object.sheet_by_name('function')
        buffer_sign = str(ex_table.cell(table_line, 12).value)
        if buffer_sign != '':
            save_list = eval(ex_table.cell(table_line, 12).value.encode("utf-8"))
            log_user.debug('需要缓存的数据:%s  数据类型:%s' % (save_list, type(save_list)))
            memory = GlobalConfiguration.get_cache_region()
            log_user.debug('请求数据:%s' % payload)
            memory = save_data(save_list, payload, memory)
            GlobalConfiguration.save_cache_region(memory)
            log_user.info('已缓存的数据:%s' % memory)
            memory = save_data(save_list, rq.json(), memory)
            GlobalConfiguration.save_cache_region(memory)
            log_user.info('已缓存的数据:%s' % memory)
        log_user.info('响应返回数据:\n%s' % json.dumps(rq.json(), ensure_ascii=False, indent=4))
        delay = ex_table.cell(table_line, 10).value              # 读取延时
        if delay != '':
            delay = int(delay)
            log_user.info('等待后台处理，延时：%s 秒' % delay)
            time.sleep(delay)
        log_user.info('------分析响应结果 ------')
        criteria = ex_table.cell(table_line, 9).value.encode("utf-8")                # 断言方式
        ex_table = ex_object.sheet_by_name('data')
        data_type = str(ex_table.cell(test_case, (response_line_num - 1)).value)                          # 返回数据方式
        log_user.debug('------test_case:%s test_lien:%s data_type:%s ------' % (test_case, (response_line_num - 1), data_type))
        if data_type == '0' or data_type == '':
            comparison = eval(ex_table.cell(test_case, response_line_num).value.encode("utf-8"))  # 对比结果
        elif data_type == '1':
            comparison = ex_table.cell(test_case, response_line_num).value.encode("utf-8")          # 读取参数化数据
            content = ex_table.cell(test_case, (response_line_num + 1)).value.encode('utf-8')             # 读取替换数据
            comparison = eval(DataHandle(data_type, comparison, content, log_user))
        elif data_type == '2':
            comparison = ex_table.cell(test_case, response_line_num).value.encode("utf-8")          # 读取参数化数据
            content = eval(ex_table.cell(test_case, (response_line_num + 1)).value.encode('utf-8'))
            state, content = PerformDBASql(content, log_user)                           # 执行sql，查询数据
            if state != 0:
                comparison= eval((DataHandle(data_type, comparison, content, log_user)).encode('utf-8'))    # 替换数据
                log_user.debug('替换后的数据:\n%s' % json.dumps(comparison, ensure_ascii=False, indent=4))
            else:
                log_user.info("数据库存操作失败:%s  **" % content)
                return 2, '', 'database operation Fail'
        elif data_type not in ('0', ''):
            return 2, '', 'Get Parametric approach does not suppor'
        response_time = rq.elapsed.microseconds / 1000                                    # 响应时间
        status_info = GlobalConfiguration.get_error_code(rq.status_code).encode("utf-8")  # 获取状态码描述
        log_user.info("请求响应时间:%dms 响应状态:%d %s cookies: %s  sessionid: %s"
                      % (response_time, rq.status_code, status_info, rq.cookies, GlobalConfiguration.jsessionid))
        if criteria == 'AssertEqual':                                                  # 断言方式相等
            log_user.info('数据断言方式：AssertEqual  预期响应数据:\n%s ' % json.dumps(comparison, ensure_ascii=False, indent=4))
            log_user.debug('comparison:%s  ,text:%s' % (type(comparison), rq.json()))
            if comparison == rq.json():
                log_user.info('执行结束：%s 用例 接口序号：%s  接口名称：%s   执行通过' % (test_case_name, table_line, test_name))
                return 1, comparison, rq.json()                  # 返回成功
            else:
                log_user.info('执行结束：%s 用例 接口序号：%s  接口名称：%s   执行失败' % (test_case_name, table_line, test_name))
                return 0, comparison, rq.json()
        elif criteria == 'AssertNotEqual':                                  # 断言方式不相等
            log_user.info('数据断言方式： AssertNotEqual  预期响应数据:\n%s' % json.dumps(comparison, ensure_ascii=False, indent=4))
            if comparison != rq.json():
                log_user.info('执行结束：%s 接口序号：%s  接口名称：%s   执行通过' % (test_case_name, table_line, test_name))
                return 1, comparison, rq.json()                # 返回成功
            else:
                log_user.info('执行结束：%s 接口序号：%s  接口名称：%s   执行失败' % (test_case_name, table_line, test_name))
                return 0, comparison, rq.json()                  # 返回失败
        elif criteria == 'Contain':                                       # 断言方式包含
            rq = json.loads(rq.text)                                         # 转为字典数据
            log_user.info('数据断言方式： Contain   预期响应数据:\n%s' % json.dumps(comparison, ensure_ascii=False, indent=4))
            result, info_a, info_b = cmp_json(rq, comparison, log_user)    # 对比数据
            if result != '1':                                            # 调用断言方式 返回1 正常  ，非1不正确
                log_user.info('响应数据匹配失败:\n响应数据: %s 参考响应数据 :%s' % (str(info_a), str(info_b)))
                log_user.info('执行结束：%s 接口序号：%s  接口名称：%s   执行失败  原因：响应数据匹配失败'
                              % (test_case_name, table_line, test_name))
                return 0, info_a, info_b
            else:
                log_user.info('执行结束： %s 接口序号：%s  接口名称：%s   执行通过' % (test_case_name, table_line, test_name))
                log_user.info('------ 数据分析结束 ------')
                return 1, result, '1'

        else:
            log_user.info('不支持的断言方式')
            return 2, '', 'Judgment way wrong, not Contain or AssertNotEqual or AssertEqual'


# 处理键值类型  返回状态1表示查询到数据
def key_value_type(parameter, target):
    for k, v in parameter.items():
        if k == target:
            return 1, v
        elif isinstance(v, list):
            state, result = key_list_type(v, target)
            if state == 1:
                return state, result
        elif isinstance(v, dict):
            state, result = key_dict_type(v, target)
            if state == 1:
                return state, result
    return 0, None


# 处理字典类型  返回状态1表示查询到数据
def key_dict_type(parameter, target):
    for k, v in parameter.items():
        if k == target:
            return 1, v
        elif isinstance(v, list):
            state, result = key_list_type(v, target)
            if state == 1:
                return state, result
        elif isinstance(v, dict):
            state, result = key_dict_type(v, target)
            if state == 1:
                return state, result
    return 0, None


# 处理列表类型  返回状态1表示查询到数据
def key_list_type(parameter, target):
    list_len = len(parameter)
    for i in range(list_len):
        read_data = parameter[i]
        if isinstance(read_data, dict):
            state, result = key_dict_type(read_data, target)
            if state == 1:
                return state, result
    return 0, None


def read_json_data(parameter, target):
    for k, v in parameter.items():
        if k == target:
            return 1, v
        elif isinstance(v, list):
            state, result = key_list_type(v, target)
            if state == 1:
                return state, result
        elif isinstance(v, dict):
            state, result = key_dict_type(v, target)
            if state == 1:
                return state, result
    return 0, None


def save_data(parameter, response_data, cache_region):
    for p in parameter:
        state, result = read_json_data(response_data, p)
        if state == 1:
                cache_region[p]=result
    return cache_region


# 合并数据
def merge_data(parameter_one, parameter_twe):
    for k, v in parameter_twe.items():
        parameter_one[k] = v
    return parameter_one


if __name__ == '__main__':
    import logging
    import logging.config
    # logging.config.fileConfig(r'G:\Python\CashBox\CashBox\Configuration\logging.conf')
    logging.config.fileConfig(r'G:\Python\boxwallet\Configuration\logging.conf')
    LoggUser = logging.getLogger('Example')
    # GlobalConfiguration.ConfigInit(r'G:\Python\CashBox\CashBox\Configuration\TestConfiguration.cfg')
    GlobalConfiguration.ConfigInit(r'G:\Python\boxwallet\Configuration\TestConfiguration.cfg',LoggUser)

    testsql1='select * for a;selcet * form b;'
    testsql={"t_fspf_omms_n": {
        "select": {
            "phone": "select t.user_login from tbl_mcht_user t where t.mcht_no='001445395001768'and t.user_primary='1'",
            "id": "select t.user_card_no from tbl_mcht_user t where t.mcht_no='001445395001768'and t.user_primary='1'"}}}
    testsqltowe={"online_bus": {
        "updata": {11:"update  tbl_icb_good t set t.good_store_num =11 where t.good_id =2;",
                   12:"update  tbl_icb_good t set t.good_store_num =12 where t.good_id =1"},
        "select": {11:"update  tbl_icb_good t set t.good_store_num =11 where t.good_id =2;",
                   12:"update  tbl_icb_good t set t.good_store_num =12 where t.good_id =1"}
                }}