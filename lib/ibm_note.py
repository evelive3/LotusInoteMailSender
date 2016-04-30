# -*- coding: utf-8 -*-

from lxml import etree
import requests

from .forms import send_data


# IBM Lotus iNote类
class Mail(object):
    def __init__(self):
        self.BASE_URL = 'http://mail.clic'

        self._session = None
        self._user_url = None
        self.last_response = None

    def do_login(self, username, password):
        # 取得登陆session
        s = requests.session()
        # 伪装请求头 <Chrome浏览器>
        s.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/47.0.2526'
                          '.111 Safari/537.36'}
        # 获取登陆页信息
        _response = s.get(self.BASE_URL)
        root = etree.HTML(_response.text)
        login_data = {x.attrib.get('name'): x.attrib.get('value') for x in root.xpath(u"//form//input") if
                      x is not None}  # 提取登陆表单并将提交按钮排除
        login_data['username'], login_data['password'] = username, password
        login_url = self.BASE_URL + root.xpath(u"//form[@name='_DominoForm']")[0].attrib['action']  # 提取登陆验证url
        _response = s.post(login_url, data=login_data)
        # 验证登陆
        root = etree.HTML(_response.text)
        if root.xpath(u"//form[@name='_DominoForm']"):
            raise AttributeError('username or password is wrong')
        else:
            self._session = s
            # 更新用户页url
            _response = s.get(self.BASE_URL + '/redirec.nsf')
            root = etree.HTML(_response.text)
            # 取得跳转至邮箱个人面板网址，结果类似'4;URL=//HQMS15.mail.clic/ma4/username.nsf?OpenDatabase'
            _response = s.get('http:' + root.xpath(u"//meta")[0].attrib['content'].split(";")[1][4:])
            self._user_url = _response.url.split('?')[0]

        # 更新最后访问页
        self.last_response = _response

    def send_mail(self, subject, mail_to, message):
        """
        :param message: mail_subject
        :param mail_to: Lotus-iNote-Sytle-Email-Address  Like this: ChineseName/OU/OU/OU@Domain
        :param subject: mail_body
        """
        send_url = self._user_url + "($Inbox)/$new/?EditDocument&Form=h_PageUI&PresetFields=h_EditAction;h_ShimmerEdi" \
                                    "t,s_ViewName;($Inbox),s_NotesForm;Memo&ui=dwa_form"
        _response = self._session.get(send_url)  # 获取cookie
        mail_to = mail_to.split("/")

        # 填写发件人，收件人，主题，正文等
        send_data['Subject'], send_data['h_Name'] = subject, subject  # 设定标题
        send_data['SendTo'] = 'CN={0}/OU={1}/OU={2}/O={3}'.format(mail_to[0], mail_to[1], mail_to[2],
                                                                  mail_to[3])  # 设定收件人
        send_data['Body'] = message  # 设定正文
        # 发送邮件
        _response = self._session.post(send_url, send_data)

        self.last_response = _response
