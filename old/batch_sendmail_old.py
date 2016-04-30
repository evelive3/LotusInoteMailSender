from iNote import LotusINoteConnector, LotusINoteMailSender, getUserUrl
from jinja2 import Template
import pathlib
import csv
from collections import OrderedDict
import time

# 设定基础信息
baseUrl = 'http://mail.clic'
username = 'XXX'
password = 'XXX'
subject = "{0}，此为2016年二月工资条，请收悉"
userNameIncField = '收件人'  # 包含用于发送的用户名字段

# 登陆邮箱
session = LotusINoteConnector(username, password, baseUrl)

# 获取用户登录后url
userUrl = getUserUrl(session)

# 判断是否已经登陆，才进行下步操作，否则返回错误信息
if session:
    # 读取邮件模板文件
    tplFile = open('tpl\\mail.tpl', mode='r', encoding='UTF-8')
    sendTPL = tplFile.read()
    tplFile.close()
    # 准备发送模板
    template = Template(sendTPL)

    # 生成CSV文件夹内的csv后缀文件名列表
    csvPath = pathlib.Path("csv")
    csvFileList = [str(x) for x in csvPath.glob("*.csv")]

    # 检查csv目录下是否存在csv文件,否则返回错误信息
    if csvFileList:
        # 读取各csv文件
        for csvFileName in csvFileList:
            # 打开csv文件
            with open(csvFileName, mode='r') as csvFile:
                csvReader = csv.DictReader(csvFile)  # 调用csv的直接读取方法，每行会读取为一个字典
                # 逐行读取当前打开csv
                for row in csvReader:
                    csvFieldList = csvReader.fieldnames  # 取得有序的表头列表
                    csvField = csvFieldList.copy()  # 建立有序表头列表的影表
                    # 删除相对于结果字典内value为空的表头列表项
                    for k, v in row.items():
                        if v in ("", None):
                            csvField.remove(k)
                    # 跟据表头列表顺序，有序的组装新的返回字典
                    result = OrderedDict()  # 新建有序字典实列
                    for title in csvField:
                        result[title] = row[title]  # 装填有序字典
                    # 开始填充发送邮件的关键信息
                    sendto = '{0}/{1}/{2}/CLIC@CLIC'.format(result[userNameIncField], '雅安', 'SC')
                    # 弹掉收信人地址栏（该信息并不需要发送给客户）
                    result.pop(userNameIncField)
                    sendbody = template.render(rows=result)
                    # 发送邮件
                    sendState = LotusINoteMailSender(session, userUrl, subject.format(result['员工/申请者姓名']), sendto, sendbody)
                    if sendState:
                        print(">>>{0} 发送成功".format(sendto))
                    else:
                        print(">>>{0} 发送失败".format(sendto))
                    # 暂停几秒后再处理下一行
                    time.sleep(10)
    else:
        print("CSV目录下未找到csv文件")
else:
    print("登陆失败，请核对用户名，密码与邮箱web入口！")
