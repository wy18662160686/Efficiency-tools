import requests,re,xlwt


def RegFunction(html, reg):
    '''通用正则匹配函数'''
    reg = reg
    res = re.compile(reg)
    list = re.findall(res, html)
    return list

def oaLogin():
    requests.packages.urllib3.disable_warnings()
    username = "wuyang@kedacom.com"
    passwd = "Wu1991y0520"
    login_url = 'https://sso.kedacom.com:8443/CasServer/login'

    login_head = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64; rv:59.0) Gecko/20100101 Firefox/59.0',
        'Referer': 'https://sso.kedacom.com:8443/CasServer/login'
    }
    se = requests.session()
    #print(se.get(login_url, verify=False).text)
    #print(RegFunction(se.get(login_url, verify=False).text, r'name="loginTicket" value="(.*)"/>'))
    lt = RegFunction(se.get(login_url, verify=False).text, r'name="loginTicket" value="(.*)"/>')[0]

    login_data = {
        'username': username,
        'password': passwd,
        'loginTicket': lt,
        'execution': 'e1s1',
        '_eventId': 'submit',
        'vcode': '',
        'submit': ''
    }
    time_data = {
        "page": 1,
        "limit": 500,
        "start": 0
    }
    time_url = "https://oa.kedacom.com/report/report/approve/initData.do"

    se.post(login_url, login_data, login_head, verify=False)
    worktime = se.post(time_url, time_data, verify=False)
    return worktime.json()


workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('工时汇总')
worksheet.write(0, 0, "项目名称")
worksheet.write(0, 1, "姓名")
worksheet.write(0, 2, "填报日期")
worksheet.write(0, 3, "填写工时")

oa_content=oaLogin()
contet_date=oa_content['content']

x=1
for i in contet_date:
    worksheet.write(x, 0, i['projectName'])
    worksheet.write(x, 1, i['fullname'])

    worksheet.write(x, 2, i['taskDate'])
    worksheet.write(x, 3, i['hours'])
    x+=1

workbook.save('D:/工时提交情况/49周工时提交情况.xlsx')
