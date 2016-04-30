# ibm_lotus_inote
### IBM Lotus iNote批量发送器
这是一个每个月批量发送工资条（也可以是别的内容）的脚本。
### 制作原因

 - 从外网通过Internat地址向公司IBM Lotus iNote发送邮件被Ban的机率太高，所以使用requests连接内网web页面
 - 必须清除csv文件里每行记录里的空值，每个人空的地方还不一样，两百条记录左右，人工做会花上几小时（据说