import requests
from lxml import etree
import json


# IBM Lotus iNote网页端连接函数
def LotusINoteConnector(username, password, baseurl):
    connected = False
    with requests.session() as s:
        # 设定请求头为模仿Chrome浏览器
        s.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/47.0.2526.111 Safari/537.36'}
        # 进入邮箱首页，获取必要数据
        p = requests.get(baseurl)
        p = etree.HTML(p.text)
        # 获取登录页url
        login_url = baseurl + p.xpath(u"//form[@name='_DominoForm']")[0].attrib['action']
        # 读取目标form所有input项的name与value并组装为字典
        loginRequiredDict = {x.attrib.get('name'): x.attrib.get('value') for x in p.xpath(u"//form//input")}
        loginRequiredDict.pop(None)  # 弹掉无用的登陆按钮项
        loginRequiredDict['Username'] = username
        loginRequiredDict['password'] = password
        # 模拟用户登录
        login = s.post(url=login_url, data=loginRequiredDict)
        # 查找登录表单，若未找到，说明已登录
        loginPage = etree.HTML(login.text)
        if not loginPage.xpath(u"//form[@name='_DominoForm']"):
            connected = True  # 更改连接状态为已连接
    # 如果连接成功，返回当前连接session，失败返回布尔值假
    if connected:
        return s
    else:
        return False


# 定义用户url获取函数
def getUserUrl(session):
    # 取得已登录session
    s = session
    redirec_nsf = s.get('http://mail.clic/redirec.nsf')
    redirecPage = etree.HTML(redirec_nsf.text)
    # 取得跳转至邮箱主页网址，结果类似'4;URL=//HQMS15.mail.clic/ma4/username.nsf?OpenDatabase'
    redirecUrl = redirecPage.xpath(u"//meta")[0].attrib['content']
    # 按;号分隔结果字符串
    redirecUrl = redirecUrl.split(";")[1]
    # 去掉多余字符并拼接成目标字符串
    redirecUrl = 'http:' + redirecUrl[4:]
    # 跳转至个人邮箱主页
    userIndex = s.get(redirecUrl)

    # 获取用户登陆后url
    curr_url = userIndex.url.split('?')[0]  # 取用户首页url,按?符号分隔，取前一段
    return curr_url


# IBM Lotus iNote网页端邮件发送函数
def LotusINoteMailSender(session, userurl, subject, sendto, sendbody):
    """
    session = session requestSession
    subject = mail_subject str
    sendto = Lotus-iNote-Sytle-Email-Address str  Like this: ChineseName/OU/OU/Domain
    sendfrom = like sendto
    sendbody = mail_body str
    """
    # 拆分收件人邮件地址为系统可识别样式
    sendTo = sendto.split("/")
    # 取得已登录session
    s = session

    # 发送邮件前置准备
    curr_url = userurl  # 取用户首页url,按?符号分隔，取前一段
    sendUrl = curr_url + '/' + "($Inbox)/$new/?EditDocument&Form=h_PageUI&PresetFields=h_EditAction;h_ShimmerEdit,s_ViewName;($Inbox),s_NotesForm;Memo&ui=dwa_form"  # 合成目标发件url
    sendPage = s.get(sendUrl)
    sendPage = etree.HTML(sendPage.text)
    # 初始化post字典
    sendRequiredDict = {
        '%%Nonce': None,
        'h_SceneContext': None,
        'h_EditAction': 'h_Next',
        'h_SetEditCurrentScene': 'l_StdPageEdit',
        'h_PageText': None,
        's_SubjectText': None,
        'JITUsers': None,
        'NotesRecips': None,
        'hFailedUsers': None,
        'SMIMESign': None,
        'SMIMERecips': None,
        'NoDomRecips': None,
        'Sign': 0,
        'Encrypt': 0,
        'TrustInetCerts': 0,
        's_AllRecips': None,
        'h_SetEditNextScene': None,
        'h_SetReturnURL': '[[./&Form=l_CallListener]]',
        'h_NoSceneTrail': 0,
        'h_SetCommand': 'h_ShimmerSendMail',
        'h_SetSaveDoc': 1,
        's_MailSendReturnPage': None,
        's_MailViewBefore': None,
        'h_SetPublishToFolder': None,
        'h_Name': None,
        'h_SetPublishAction': 'h_Publish',
        'h_EditSceneTrail': None,
        'h_WorkflowStage': None,
        'h_DictionaryId': None,
        's_ViewName': None,
        'h_SetDeleteListCS': None,
        's_DisclaimerIsAdded': 0,
        '$Disclaimed': 1,
        'h_SpellCheckStatus': None,
        'SendTo': None,
        'CopyTo': None,
        'BlindCopyTo': None,
        'Body': None,
        'BodyPT': None,
        'NumOfBodyPT': None,
        'BodyImgCids': None,
        'From': None,
        'AltFrom': None,
        '$LangFrom': None,
        's_InetFrom': None,
        'Principal': None,
        '$AltPrincipal': None,
        '$LangPrincipal': None,
        'MailOptions': 1,
        'Form': 'Memo',
        'ReturnReceipt': 0,
        '$KeepPrivate': 0,
        'Importance': 2,
        'DeliveryReport': 'B',
        'DeliveryPriority': 'N',
        'SaveOptions': None,
        'h_SetParentUnid': None,
        'In_Reply_To': None,
        'References': None,
        'h_DestFolder': None,
        'h_Move': None,
        'h_SetDeleteList': None,
        's_SendAndFile': None,
        's_NewFolderList': None,
        's_SaveFollowUp': 0,
        's_SaveFollowUpAlarm': 0,
        's_RemoveFollowUpAlarm': 0,
        'FollowUpStatus': None,
        'FollowUpText': None,
        'FollowUpDate': None,
        'FollowUpTime': None,
        'Alarms': 0,
        '$Alarm': 0,
        '$AlarmOffset': None,
        '$AlarmUnit': None,
        '$AlarmTime': None,
        'h_AlarmOn': None,
        '$AlarmMemoOptions': None,
        '$AlarmSendTo': None,
        '$AlarmSound': None,
        's_SetRFSaveInfo': None,
        's_SetReplyFlag': 0,
        's_SetForwardedFrom': 0,
        's_IgnoreQuota': 0,
        's_LDAPGroup': None,
        's_UsePlainText': 0,
        's_UsePlainTextAndHTML': 0,
        's_PlainEditor': 0,
        'h_NumOfPageText': None,
        's_EmbeddedImageInfo': None,
        's_ImageUseCidRef': None,
        's_CidImageInfo': None,
        's_ConvertImage': 0,
        's_ConvertQuickrIconImageInfo': None,
        's_NotesLinkIconInfo': None,
        'h_ImageCount': 0,
        's_DataUriInfo': None,
        '%%PostCharset': 'UTF-8',
        'RemoveAtClose': None,
        'Classification': None,
        'ConfidentialString': None,
        'DocExExpireDate': None,
        'ExpandPersonalGroups': 1,
        'IsMailStationery': None,
        'MailStationeryName': None,
        's_IsStationery': None,
        'h_AttachmentFileItemNames': None,
        'h_AttachmentRealNames': None,
        'h_AttachmentAppleTypes': None,
        'Subject': None
    }
    # 填写发件人，收件人，主题，正文等
    sendRequiredDict['Subject'] = subject  # 设定标题
    sendRequiredDict['h_Name'] = subject  # 设定标题
    sendRequiredDict['SendTo'] = 'CN={0}/OU={1}/OU={2}/O={3}'.format(sendTo[0], sendTo[1], sendTo[2],
                                                                     sendTo[3])  # 设定收件人
    sendRequiredDict['Body'] = sendbody  # 设定正文

    # 发送邮件
    sendMail = s.post(url=sendUrl, data=sendRequiredDict)
    sendPage = etree.HTML(sendMail.text)
    # TODO 检查邮件发送状态，成功返回逻辑真，失败返回逻辑假
    if sendMail.status_code == 200:
        return True
    else:
        return False


# 定义用户名获取函数
def getUserName(session, userurl, username):
    userUrl = userurl + '/iNotes/Proxy/?EditDocument&Form=s_ValidationJson'
    userdict = {'VAL_NameEntries': username, 'VAL_ExpandGroup': 0, 'VAL_Type': 1, 'VAL_Flags': 0,
                'VAL_SendEncrypted': 0, '%%PostCharset': 'UTF-8'}
    userNameStr = session.post(url=userUrl, data=userdict)
    if userNameStr.status_code == 200:
        # 因为接收到的json字符串带有注释，而标准json是不支持注释的，所以需要将注释去掉
        userNameStr = userNameStr.text.partition('      ')[2]
        # 调用loads方法将字符串加载为json
        userNameJson = json.loads(userNameStr)
        return userNameJson
    else:
        return False
