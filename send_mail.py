#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from collections import OrderedDict
from jinja2 import Template
import pathlib
import time
import csv

from lib.ibm_note import Mail

# ---SETTING FIELDS---
username = 'Your username'
password = 'Your password'
subject = '{0}，此为2016年X月工资条，请收悉'
sendto_field = '收件人'  # csv文件中收件人列名，ChineseName/OU/OU/O@Domain
delay_time = 10  # 每封邮件间隔时间(秒)

# ---START BATCH SEND MAIL---
mail = Mail()
# 尝试登陆IBM Lotus iNote
try:
    mail.do_login(username, password)
except Exception as e:
    print('出错啦: ', e)

# 读取csv目录
csv_path = pathlib.Path("csv")
csv_file_list = [str(x) for x in csv_path.glob("*.csv")]
if csv_file_list is None:
    raise FileNotFoundError('csv目录下没有.csv文件')

# 读取邮件模板文件
with open('tpl\\mail.tpl', mode='r', encoding='UTF-8') as tpl_file:
    template = Template(tpl_file.read())

# 遍历csv目录
for csv_file_name in csv_file_list:
    # 打开csv文件
    with open(csv_file_name, mode='r') as csv_file:
        rows = csv.DictReader(csv_file)  # 调用csv的字典读取方法，每行会读取为一个字典
        field_list = rows.fieldnames  # 取得表头列表
        for row in rows:
            # 依表头顺序有序生成字典，排除空值
            result = OrderedDict()
            for field in field_list:
                if row[field] not in ('', None):
                    result[field] = row[field]
            # 开始填充发送邮件的关键信息
            mail_to = result[sendto_field]
            result.pop(sendto_field)  # 弹掉收信人地址栏（该信息仅用于存储收件人地址）
            message = template.render(rows=result)  # 渲染邮件正文
            mail.send_mail(subject=subject.format(result['员工/申请者姓名']), mail_to=mail_to, message=message)
            # 显示投递结果
            status_text = '>>> To: {0}\t{1}.'
            if mail.last_response.status_code is 200:
                print(status_text.format(mail_to, 'success'))
            else:
                print(status_text.format(mail_to, 'failed'))
            # 发件延迟
            time.sleep(delay_time)



