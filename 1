#!/usr/bin/env python
# -*- coding: utf-8 -*-
# author：YFWang time:2021/12/22

import xlwt
import xlrd
import pymysql
import math
import time
import datetime
import smtplib  # 加载smtplib模块
from email.mime.text import MIMEText
from email.utils import formataddr
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


class SendMail(object):
    def __init__(self, sender, title, content):
        self.sender = sender  # 发送地址
        self.title = title  # 标题
        self.content = content  # 发送内容
        self.sys_sender = 'qazwsx910908@163.com'  # 系统账户
        self.sys_pwd = 'XJZDDLHZYACGJVWM'  # 系统账户密码

    def send(self, file_list):
        """
        发送邮件
        :param file_list: 附件文件列表
        :return: bool
        """
        try:
            # 创建一个带附件的实例
            msg = MIMEMultipart()
            # 发件人格式
            msg['From'] = formataddr(["技术服务部", self.sys_sender])
            # 收件人格式
            msg['To'] = formataddr(["", self.sender])
            # 邮件主题
            msg['Subject'] = self.title

            # 邮件正文内容
            msg.attach(MIMEText(self.content, 'plain', 'utf-8'))

            # 多个附件
            for file_name in file_list:
                print("file_name", file_name)
                # 构造附件
                xlsxpart = MIMEApplication(open(file_name, 'rb').read())
                # filename表示邮件中显示的附件名
                xlsxpart.add_header('Content-Disposition', 'attachment', filename='%s' % file_name)
                msg.attach(xlsxpart)

            # SMTP服务器
            server = smtplib.SMTP_SSL("smtp.163.com", 465, timeout=10)
            # 登录账户
            server.login(self.sys_sender, self.sys_pwd)
            # 发送邮件
            server.sendmail(self.sys_sender, [self.sender, ], msg.as_string())
            # 退出账户
            server.quit()
            return True
        except Exception as e:
            print(e)
            return False


def len_byte(value):
    length = len(value)
    utf8_length = len(value.encode('utf-8'))
    length = (utf8_length - length) / 2 + length
    return int(length)


#   判断当前费率计费单位
def fee_cal(holdt, feet):
    q = 60
    y = 6
    if math.ceil(float(holdt) / int(q)) == float(feet) / int(q):
        return int(q)
    else:
        return int(y)

# def fee_cal(holdt, feet):
#     y = ['60', '6']
#     for x in y:
#         if math.ceil(float(holdt) / int(x)) == float(feet) / int(x):
#             return int(x)


def importExcelToMysql(cur, path):
    # 读取excel文件
    workbook = xlrd.open_workbook(path)
    sheets = workbook.sheet_names()
    worksheet = workbook.sheet_by_name(sheets[0])

    # 将表中数据读到 sqlstr 数组中
    for i in range(1, worksheet.nrows):
        row = worksheet.row(i)

        sqlstr = []

        for j in range(0, worksheet.ncols):
            sqlstr.append(worksheet.cell_value(i, j))
        ###
        valuestr = [str(sqlstr[0]), str(sqlstr[1]), str(sqlstr[2]), str(sqlstr[3]), str(sqlstr[4]), float(sqlstr[5]),
                    int(sqlstr[6]), float(sqlstr[7]), int(sqlstr[8]), int(sqlstr[9]), float(sqlstr[10]),
                    float(sqlstr[11]),
                    int(sqlstr[12]), int(sqlstr[13]), int(sqlstr[14]), str(sqlstr[15]), str(sqlstr[16]),
                    str(sqlstr[17]),
                    str(sqlstr[18]), str(sqlstr[19]), str(sqlstr[20]), str(sqlstr[21]), str(sqlstr[22]), str(sqlstr[23]),
                    str(sqlstr[24])]

        # 将每行数据存到数据库中
        cur.execute("insert into bill_new (date,sell_account,sell_account_name,caller_gw,callerip,sell_price,sell_unit,"
                    "income,sell_duration_60s,sell_duration_6s,cost,cost_price,cost_unit,cost_duration_60s,"
                    "cost_duration_6s,callee_gw,cost_account,cost_account_name,cdr_num,saler,real_cost_name,real_customer_name,cost_class,account_class,vos) "
                    "values (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", valuestr)


def readTable(cursor):
    # 选择全部
    cursor.execute("select * from bill_new")
    # 获得返回值，返回多条记录，若没有结果则返回()
    results = cursor.fetchall()

    for i in range(0, results.__len__()):
        for j in range(0, 20):
            print(results[i][j], end='\t')
        print('\r')


def yes_time():
    now_time = datetime.datetime.now()
    yes_time = now_time + datetime.timedelta(days=-1)
    yes_time_nyr = yes_time.strftime('%Y%m%d')
    return yes_time_nyr


# day_g = input("请输入日期:")
w = 1
port = 3305
while w == 1:
    port = port + 1
    w += 1
    db = eval("pymysql.connect(host='192.168.2.75', user='root', password='123456', db='vos3000', port=%s)" % (port))
    db2 = eval("pymysql.connect(host='192.168.2.75', user='root', password='123456', db='report', port=3306)")
    # db = eval("pymysql.connect(host='127.0.0.1', user='root', password='', db='vos3000', port=%s)" % (port))
    # db2 = eval("pymysql.connect(host='127.0.0.1', user='root', password='', db='report', port=3306)")
    cursor = db.cursor()
    cursor2 = db2.cursor()
    if port == 3306:
        vosname = "111.1.32.152"
    # elif port == 3307:
    #     vosname = "116.62.117.222"
    # elif port == 3308:
    #     vosname = "39.99.185.107"
    # else:
    #     vosname = "39.98.114.7"
    try:
        book = xlwt.Workbook()  # 新建一个Excel
        sheet = book.add_sheet('导出数据')  # 创建sheet
        title = ['日期', '账户号码', '账户名称', '主叫网关', '主叫IP', '客户单价', '客户计费单位', '话费总计', '计费时长(60+)次', '计费时长(6+)次', '结算成本',
                 '结算单价', '结算计费单位', '结算时长(60+)次', '结算时长(6+)次', '落地网关', '结算账户号码', '结算账户名称', '话单数量(>0)', '对应业务员',
                 '真实供应商', '真实客户名', '供应商类型', '客户类型', 'vos']  # 写表头
        # 循环将表头写入到sheet页
        x = 0
        rowxxx = 0
        ret = ["20230410"]
        # ret.append("%s" % str(yes_time()))
        # print(ret)
        for header in title:  # 打印表头
            sheet.write(0, x, header)
            x += 1
        # print(vosname)
        try:
            for date in ret:
                sql1 ="""
                SELECT 
                  e.holdtime,
                  e.feetime,
                  e.fee,
                  e.agentfeetime,
                  e.agentfee,
                  e.customeraccount,
                  e.customername,
                  e.callergatewayid,
                  e.callerip,
                  SUM(e.fee),
                  SUM(CEILING(e.holdtime / 60)) AS holdtime_min,
                  SUM(CEILING(e.holdtime / 6)),
                  SUM(e.agentfee),
                  SUM(CEILING(e.agentfeetime / 60)),
                  SUM(CEILING(e.agentfeetime / 6)),
                  e.calleegatewayid,
                  e.agentaccount,
                  e.agentname,
                  COUNT(*) AS total_count,
                  j.memo,
                  MAX(CASE 
                      WHEN c1.type = 2 THEN d.companyname 
                      ELSE NULL 
                    END) AS real_cost_name,
                  MAX(CASE 
                      WHEN c1.type = 0 THEN d.companyname 
                      ELSE NULL 
                    END) AS real_customer_name,
                  MAX(CASE 
                      WHEN c1.type = 2 THEN d.linkman 
                      ELSE NULL 
                    END) AS cost_class,                    
                  MAX(CASE 
                      WHEN c1.type = 0 THEN d.linkman 
                      ELSE NULL 
                    END) AS account_class
                FROM 
                  e_cdr_{0} e 
                  LEFT JOIN e_customer c1 ON e.customeraccount = c1.account 
                                          OR e.agentaccount = c1.account 
                  LEFT JOIN e_customerdetail d ON c1.id = d.customer_id 
                  LEFT JOIN e_customer j ON e.customeraccount = j.account 
                  LEFT JOIN e_customerdetail k ON j.id = k.customer_id 
                WHERE 
                  holdtime > 0 
                  AND c1.type IN (0, 2)
                GROUP BY 
                  customeraccount,
                  customername,
                  callergatewayid,
                  callerip,
                  calleegatewayid,
                  agentaccount,
                  agentname
                """.format(date)
                #       print(sql1)

                sql2 = "SELECT holdtime, feetime, fee, agentfeetime, agentfee from e_cdr_%s WHERE holdtime > 0 " \
                       "GROUP BY customeraccount,customername,callergatewayid,callerip,calleegatewayid,agentaccount,agentname" % date
                # print(sql1)
                # col_width = []
                cursor.execute(sql1)
                results_sql1 = cursor.fetchall()
                cursor.execute(sql2)
                results_sql2 = cursor.fetchall()
                # print(results_sql1[0])
                for row in range(1, len(results_sql1) + 1):
                    print(results_sql1)
                    for col in range(0, len(results_sql1[row - 1])):
                        if col == 0:
                            sheet.write(row + rowxxx, col, date)
                            # print(row + rowxxx, col, date)
                        elif 0 < col < 5:
                            sheet.write(row + rowxxx, col, results_sql1[row - 1][col + 4])
                            # print(row, col, results_sql1[row - 1][col + 4])
                        # 客户最小单价
                        elif col == 5:
                            if results_sql1[row - 1][1] != 0:
                                fee_unit = fee_cal(results_sql1[row - 1][0], results_sql1[row - 1][1])  # 客户计费单位
                                fee_count = results_sql1[row - 1][1] / fee_unit  # 客户计费次数
                                unit_price = results_sql1[row - 1][2] / fee_count  # 客户最小单价
                                sheet.write(row + rowxxx, col, unit_price)
                                # print(row + rowxxx, col, unit_price)
                            else:
                                sheet.write(row + rowxxx, col, 0)
                                # print(row + rowxxx, col, 0)
                        # 客户计费单位
                        elif col == 6:
                            # print("holdtime:%s", type(results_sql1[row - 1][0]))
                            # print("feetime:%s", type(results_sql1[row - 1][1]))
                            # print("fee:", type(results_sql1[row - 1][2]))
                            if results_sql1[row - 1][1] != 0:
                                fee_unit = fee_cal(results_sql1[row - 1][0], results_sql1[row - 1][1])  # 客户计费单位
                                fee_count = results_sql1[row - 1][1] / fee_unit  # 客户计费次数
                                unit_price = results_sql1[row - 1][2] / fee_count  # 客户最小单价
                                sheet.write(row + rowxxx, col, fee_unit)
                                # print(row + rowxxx, col, fee_unit)
                            else:
                                sheet.write(row + rowxxx, col, 0)
                                # print(row + rowxxx, col, 0)
                        elif 6 < col < 11:
                            sheet.write(row + rowxxx, col, results_sql1[row - 1][col + 2])
                        # 结算最小单价
                        elif col == 11:
                            if results_sql1[row - 1][3] != 0:
                                agent_fee_unit = fee_cal(results_sql1[row - 1][0], results_sql1[row - 1][3])
                                agent_fee_count = results_sql1[row - 1][3] / agent_fee_unit  # 结算计费次数
                                agent_unit_price = results_sql1[row - 1][4] / agent_fee_count  # 结算最小单价
                                sheet.write(row + rowxxx, col, agent_unit_price)
                                # print(row + rowxxx, col, agent_unit_price)
                            else:
                                sheet.write(row + rowxxx, col, 0)
                                # print(row + rowxxx, col, 0)
                        # 结算计费单位
                        elif col == 12:
                            if results_sql1[row - 1][3] != 0:
                                agent_fee_unit = fee_cal(results_sql1[row - 1][0], results_sql1[row - 1][3])  # 结算计费单位
                                agent_fee_count = results_sql1[row - 1][3] / agent_fee_unit  # 结算计费次数
                                agent_unit_price = results_sql1[row - 1][4] / agent_fee_count  # 结算最小单价
                                sheet.write(row + rowxxx, col, agent_fee_unit)
                                # print(row + rowxxx, col, agent_fee_unit)
                            else:
                                sheet.write(row + rowxxx, col, 0)
                                # print(row + rowxxx, col, 0)
                        elif 12 < col < 24:
                            sheet.write(row + rowxxx, col, results_sql1[row - 1][col])
                        else:
                            print(vosname)
                            sheet.write(row + rowxxx, col, vosname)
                        # else:
                        #     sheet.write(row + rowxxx, col, results_sql1[row - 1][col])
                            # print(row + rowxxx, col, results_sql1[row - 1][col])

                rowxxx = rowxxx + len(results_sql1)

            book.save("%s-%s.xls" % (ret[0], vosname))
            importExcelToMysql(cursor2, "%s-%s.xls" % (ret[0], vosname))
            if __name__ == '__main__':
                # 发送地址
                # sender = "huang_api@163.com"
                sender = ["m1@163.com"]
                # 标题
                title = "财务报表"
                # 发送内容
                content = "包含日期:%s\n所在VOS:%s统计表" % (ret, vosname)
                # 附件列表
                file_list = ["%s-%s.xls" % (ret[0], vosname)]
                file_ret = SendMail(sender, title, content).send(file_list)
                for send in sender:
                    file_ret = SendMail(send, title, content).send(file_list)
                    print(file_ret, type(file_ret))
        except Exception as e:
            raise e
    finally:
        db.close()
        db2.commit()
        db2.close()
