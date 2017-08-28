#!/usr/bin/env python
# coding=utf-8

# Todo：消息通知
# Author：tester

import os
import urllib.request
import json
import urllib
import time
import os,sys
import xlwt

try:
    import xlrd
except:
    os.system('pip install -U xlrd')
    import xlrd
try:
    from pyDes import *
except ImportError as e:
    os.system('pip install -U pyDes --allow-external pyDes --allow-unverified pyDes')
    from pyDes import *
import hashlib
import base64
import smtplib
from email.mime.text import MIMEText



# 发送通知邮件
def sendMail(text):
    sender = 'wangshaokun@88gongxiang.com'
    receiver = ['wangshaokun@88gongxiang.com']
    mailToCc = ['wangshaokun@88gongxiang.com']
    subject = '[AutomantionTest]优惠券接口自动化测试报告通知'
    smtpserver = 'smtp.exmail.qq.com'
    username = 'wangshaokun@88gongxiang.com'
    password = 'Aa000000'

    msg = MIMEText(text, 'html', 'utf-8')
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = ';'.join(receiver)
    msg['Cc'] = ';'.join(mailToCc)
    smtp = smtplib.SMTP_SSL()
    smtp.connect(smtpserver)
    smtp.login(username, password)
    smtp.sendmail(sender, receiver + mailToCc, msg.as_string())
    smtp.quit()




# --------------------------------
# 获取企业微信token
# --------------------------------

def get_token():
    corpid = 'wwdf5248c5e061ef5a'
    corpsecret = 'AYBVGwiTiQQhCUhkjD_C1D4V5DaFzzyc6zap_s2eolc'
    url = 'https://qyapi.weixin.qq.com'
    token_url = '%s/cgi-bin/gettoken?corpid=%s&corpsecret=%s' % (url, corpid, corpsecret)
    token = json.loads(urllib.request.urlopen(token_url).read().decode())['access_token']
    return token


# --------------------------------
# 构建告警信息json
# --------------------------------
def messages(title,content,digest):
    send_values = {
        "touser": '@all',  # 企业号中的用户帐号
        # "toparty": "1",  # 企业号中的部门id
        "msgtype": "mpnews",  # 企业号中的应用id，消息类型。
        "agentid": 1000002,
        "mpnews": {
            "articles": [
                {
                    "title": title,
                    "thumb_media_id": "2RFygPDBYzQnTIKPYUj1IsFSJ46WnJkqEP_IKIYpxWoI",
                    "author": "tester",
                    "content_source_url": "URL",
                    "content": content,
                    "digest": digest,
                    "show_cover_pic": "0"
                }
            ]
        }
    }

    #https://work.weixin.qq.com/wework_admin/material/getOpenMsgBuf?type=image&media_id=2tjm0T20g2yFrOTwBKxehfpoAXpOpe1fsKV54xTaByRQ&file_name=executor_failure2.png&download=1
    msges = (bytes(json.dumps(send_values), 'utf-8'))
    return msges


# --------------------------------
# 发送告警信息
# --------------------------------
def send_message(token, data):
    url = 'https://qyapi.weixin.qq.com'

    send_url = '%s/cgi-bin/message/send?access_token=%s' % (url, token)
    respone = urllib.request.urlopen(urllib.request.Request(url=send_url, data=data)).read()
    x = json.loads(respone.decode())['errcode']
    if x == 0:
        print('微信发送状态：Succesfully')
    else:
        print('微信发送状态：Failed')
#邮件格式
def generate_sendmsg_body(errorTest):
    html = '<html><body>接口自动化定期扫描，共有 ' + str(len(
            errorTest)) + ' 个异常接口，列表如下：' + '</p><table border="1" bordercolor="pink" cellspacing="0"><th width="200">接口</th><th width="200">状态</th><th width="400">接口地址</th><th width="600">接口返回值</th></tr>'
    for test in errorTest:
        html = html + '<tr><td>' + test[0] + '</td><td>' + test[1] + '</td><td>' + test[2] + '</td><td>' + test[
            3] + '</td></tr>'
    html = html + '</table></body></html>'

    # send weixin message---------------------------------------------------------------------
    title = '接口自动化定期扫描，共有 ' + str(len(errorTest)) + ' 个异常接口' + '\n' \
                                                               '执行时间：' + time.strftime('%Y-%m-%d %H:%M:%S',
                                                                                       time.localtime(time.time()))

    description = '<h1>接口自动化定期扫描，共有 ' + str(len(errorTest)) + ' 个异常接口，列表如下：</h1>'

    for test in errorTest:
        description = description + '<p>'
        description = description + ' <font size="4" color="red"><strong>接口:' + test[0] + '</strong></font>' \
                                                                                          '<br />' + '<strong>接口地址:</strong><br />' + \
                      test[2] + \
                      '<br />' + '<strong>状态:</strong>' + test[1] + \
                      '<br />' + '<strong>接口返回值:</strong>' + test[3]
        description = description + '</p>'

    digset = '接口自动化定期扫描，共有 ' + str(len(errorTest)) + ' 个异常接口,详细参看【详情】'

    return html,title,description,digset

def send_msg(errorTest):

    (html,title,description,digset)=generate_sendmsg_body(errorTest)
    # send weixin message---------------------------------------------------------------------
    test_token = get_token()
    msg_data = messages(title, description, digset)
    #send_message(test_token, msg_data)
    # send weixin message---------------------------------------------------------------------
    #sendMail(html)



