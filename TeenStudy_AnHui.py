# -*- coding:utf-8 -*-
import requests, xlrd, smtplib
from urllib.parse import urlencode
from email.mime.text import MIMEText

# UrlEncode编码 和 UrlDecode解码


def InitEmailInfo():
    wb = xlrd.open_workbook('/Users/macbook/Documents/email.xls')    # 打开excel
    sh = wb.sheet_by_name('Sheet1')     # 按工作簿定位工作表
    AllList = []
    numList = []
    for i in range(1, sh.nrows):
        listTemp = [sh.row_values(i)[0], sh.row_values(i)[1]]
        AllList.append(listTemp)
        numList.append(sh.row_values(i)[0])
    print(AllList)
    return AllList


def GetNewTableName():  # 获取最新一期的青年大学习
    RankRsp = requests.get("http://dxx.ahyouth.org.cn/api/peopleRankList")
    RankJson = RankRsp.json()
    return RankJson['list'][0]['table_name']


def Send_Email():
    mail_msg ='''
            <!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8">
    <title>一封邮件</title>
</head>

<body style="background-color: #fff;">
    <table cellpadding="0" cellspacing="0" border="0" width="100%" style="max-width: 600px; margin: auto; font-family: Arial, sans-serif; font-size: 16px; line-height: 1.4; color: #333;">
        <!-- Header -->
        <tr>
            <td style="padding: 20px 0; text-align: center; background-color: #00bfff;">
                <h1 style="font-size: 36px; margin: 0; color: #fff;">一封邮件</h1>
            </td>
        </tr>
        <!-- Content -->
        <tr>
            <td style="padding: 20px;">
                <p style="font-weight: bold">嗨，伙计！当你看到这封邮件你一定是这周青年大学习还没学习!!</p>
                <p>你相信光吗？</p>
                <p>那天我见到了真正的迪迦奥特曼</p>
                <p>“迪迦你还记得我吗 我们曾经并肩战斗过一次。”</p>
                <p>“这样啊，我想起来了一些，可我那时候红球被抢走，你没有技能是如何帮助我拿回来的？”</p>
                <p>“谁说我没有的！我倔强仰起头，眼神坚毅，我学了青年大学习的，ta给了我强大无比的能量，你怎么能忘！”</p>
                <p>“所以好兄弟！等你学完青年大学习，我们一起守护世界！”</p>
            </td>
        </tr>
        <!-- Footer -->
        <tr>
            <td style="padding: 20px; background-color: #eee; text-align: center;">
                <p style="margin: 0;">&copy; 20级移动应用开发班</p>
            </td>
        </tr>
    </table>
</body>

</html>
    '''
    msg = MIMEText(mail_msg,'html','utf-8')  # 构造邮件，内容为青年大学习
    msg["Subject"] = "青年大学习"  # 设置邮件主题
    msg["From"] = '辛苦勤奋艰苦奋斗的团支书'  # 寄件者
    msg["To"] = '未做人员'              # 收件者
    from_addr = '24684623@qq.com'      # TODO 这里记得写你的邮箱地址
    password = 'yoockluklzakbgee'       # TODO 这里记得更改上面为from_addr的邮箱的申请码
    smtp_server = 'smtp.qq.com'  # smtp服务器地址
    to_addr = '24684623@qq.com'  # 收件人地址
    print(to_addr)
    try:
        # smtp协议的默认端口是25，QQ邮箱smtp服务器端口是465,第一个参数是smtp服务器地址，第二个参数是端口，第三个参数是超时设置,这里必须使用ssl证书，要不链接不上服务器
        server = smtplib.SMTP_SSL(smtp_server, 465, timeout=2)
        server.login(from_addr, password)  # 登录邮箱
        # 发送邮件，第一个参数是发送方地址，第二个参数是接收方列表，列表中可以有多个接收方地址，表示发送给多个邮箱，msg.as_string()将MIMEText对象转化成文本
        server.sendmail(from_addr, to_addr, msg.as_string())
        server.quit()
        print('发送邮件成功')
    except Exception as e:
        print('发送邮件失败: ', e)


def GetNotFinishList():  # 获取名单
    ClassAllListEmail = [['刘冰倩', '1064207149@qq.com'], ['莫玉洁', '2839252974@qq.com'], ['汤红红', '1522567746@qq.com'], ['赵康康', '24684623@qq.com'], ['曹席席', '2784157406@qq.com'], ['常雨欣', '3140453265@qq.com'], ['冯勤勤', '2286903066@qq.com'], ['高磊', '1909558259@qq.com'], ['高政', '1526387535@qq.com'], ['蒋逸歌', '3054413035@qq.com'], ['孔梦聃', '2793314958@qq.com'], ['廖菊', '2487276748@qq.com'], ['刘欣', '2695847218@qq.com'], ['卢雨晴', '3237707114@qq.com'], ['马德金', '2231057665@qq.com'], ['孟明丽', '3552267405@qq.com'], ['钱郡如', '2239884071@qq.com'], ['乔龙海', '3480960849@qq.com'], ['齐薇', '2472538151@qq.com'], ['宋晶晶', '2113592454@qq.com'], ['时诗瑤', '1392933913@qq.com'], ['沈万志', '2532508949@qq.com'], ['王光雪', '825904433@qq.com'], ['汪海婷', '2779287824@qq.com'], ['王俊', '3168519552@qq.com'], ['王坤', '2237631857@qq.com'], ['王开涛', '214071843@qq.com'], ['王临风', '3485460278@qq.com'], ['王紫薇', '2597976265@qq.com'], ['薛蕾', '3217985110@qq.com'], ['徐媛', '1586217801@qq.com'], ['钟华盛', '2609079534@qq.com'], ['张怀宇', '2445820246@qq.com'], ['张俊杰', '1251186942@qq.com'], ['支靓', '2628025183@qq.com'], ['张刘中祎', '2623735036@qq.com'], ['张永军', '1291459322@qq.com'], ['赵紫悦', '2833771561@qq.com']]
    ClassHaveDone = []
    ClassNotDoneName = []
    ClassNotDoneEmail = []
    headers = {
        "Host": "dxx.ahyouth.org.cn",
        "Cookie": "PHPSESSID=662c565370b0dc460fdaae9ef0aa7004",
        "User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 13_7 like Mac OS X) AppleWebKit/605.1.15 (KHTML, "
                      "like Gecko) Mobile/15E148 MicroMessenger/8.0.7(0x18000731) NetType/WIFI Language/zh_CN",
        "Accept-Language": "zh-cn",
        "Referer": "http://dxx.ahyouth.org.cn/",
        "Accept-Encoding": "gzip, deflate",
        "Connection": "keep-alive",
    }
    params = {
        "level1": "地市",
        "level2": "马鞍山市",
        "level3": "马鞍山师范高等专科学校",
        "level4": "软件工程系",
        "level5": "20级移动应用开发班"
    }
    info = urlencode(params)
    tableName = GetNewTableName()
    URL = "http://dxx.ahyouth.org.cn/api/peopleRankStage?table_name=" + tableName + "&" + info
    info_rsp = requests.get(url=URL, headers=headers)
    info_json = info_rsp.json()
    AllInfoList = info_json['list']['list']
    for people in AllInfoList:
        ClassHaveDone.append(people['username'])
    for k in ClassAllListEmail:
        if k[0] not in ClassHaveDone:       # 如果这个同学没做的话
            ClassNotDoneEmail.append(k[1])   # 就添加Email在这其中
            ClassNotDoneName.append(k[0])   # 就添加Name在这其中
    for p in ClassNotDoneName:
        print("未完成同学：", p)
    return ClassNotDoneEmail


if __name__ == '__main__':
    #InitEmailInfo()
    GetNotFinishList()  # 拿到未完成名单但是不发邮件
    #Send_Email()      # 拿到未完成名单发送邮件
