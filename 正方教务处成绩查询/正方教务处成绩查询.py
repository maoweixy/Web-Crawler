# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
import xlwt
import urllib2
import urllib
import cookielib
import sys
reload(sys)
sys.setdefaultencoding("utf-8")
class HBJJ:
    def __init__(self, baseUrl):
        # 登陆界面URL-post
        self.baseURL = baseUrl
        # 验证码的URL-get
        self.CaptchaUrl = "http://218.197.80.27/CheckCode.aspx"
        # 成绩查询URL-get
        self.graduURL1 = "http://218.197.80.27/xscjcx.aspx?xh=13150122&xm=%E6%AF%9B%E4%BC%9F&gnmkdm=N121605"
        # 历年成绩查询URL-post
        self.graduURL2 = "http://218.197.80.27/xscjcx.aspx?xh=13150122&xm=%C3%AB%CE%B0&gnmkdm=N121605"
        self.user_agent = "Mozilla/4.0 (compatible; MISE 5.5; windows NT)"
        # 声明一个CookieJar对象实例来保存cookie
        self.cookie = cookielib.CookieJar()
        self.headers = {"user-Agent": self.user_agent}
        # 登陆界面URL的header,    Referer详解见代码分析
        self.headers_gra1 = {'Referer':'http://218.197.80.27/xs_main.aspx?xh=13150122',
                                      'user-Agent': self.user_agent}
        # 历年成绩界面URL的header
        self.headers_gra2 = {'Referer': 'http://218.197.80.27/xscjcx.aspx?xh=13150122&xm=%C3%AB%CE%B0&gnmkdm=N121605',
                             'user-Agent': self.user_agent}
        # 先利用urllib2库的HTTPCookieProcessor对象来创建cookie处理器
        # 再通过handler来构建opener
        self.opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(self.cookie))
        # 用openr获取验证码pageCode
        picture = self.opener.open(self.CaptchaUrl).read()
        # 保存验证码到本地
        local = open('image.jpg', 'wb')
        local.write(picture)
        local.close()
        # 打开保存的验证码图片 输入
        SecretCode = raw_input('输入验证码： ')
        # 建立登陆界面URL的header的data
        self.postData = urllib.urlencode({
            '__VIEWSTATE': 'dDwyODE2NTM0OTg7Oz76KAIzEOLDS5fWN9NIrqLdruD9ag==',
            'txtUserName': '用户名',
            'TextBox2': '密码',
            'txtSecretCode': SecretCode,
            'RadioButtonList1': '学生',
            'Button1': '',
            'lbLanguage': '',
            'hidPdrs': '',
            'hidsc': ''
        })
        # 建立历年成绩界面URL的data
        self.postData_Gra = urllib.urlencode({
            '__EVENTTARGET':'',
            '__EVENTARGUMENT':'',
            'btn_zcj':'历年成绩',
            '__VIEWSTATE':'dDw2NDI3MTcwOTk7dDxwPGw8U29ydEV4cHJlcztzZmRjYms7ZGczO2R5YnlzY2o7U29ydERpcmU7eGg7c3RyX3RhYl9iamc7Y2pjeF9sc2I7enhjamN4eHM7PjtsPGtjbWM7XGU7YmpnOzE7YXNjOzEzMTUwMTIyO3pmX2N4Y2p0al8xMzE1MDEyMjs7MDs+PjtsPGk8MT47PjtsPHQ8O2w8aTw0PjtpPDEwPjtpPDE5PjtpPDI0PjtpPDMyPjtpPDM0PjtpPDM2PjtpPDM4PjtpPDQwPjtpPDQyPjtpPDQ0PjtpPDQ2PjtpPDQ4PjtpPDUyPjtpPDU0PjtpPDU2Pjs+O2w8dDx0PHA8cDxsPERhdGFUZXh0RmllbGQ7RGF0YVZhbHVlRmllbGQ7PjtsPFhOO1hOOz4+Oz47dDxpPDQ+O0A8XGU7MjAxNS0yMDE2OzIwMTQtMjAxNTsyMDEzLTIwMTQ7PjtAPFxlOzIwMTUtMjAxNjsyMDE0LTIwMTU7MjAxMy0yMDE0Oz4+Oz47Oz47dDx0PHA8cDxsPERhdGFUZXh0RmllbGQ7RGF0YVZhbHVlRmllbGQ7PjtsPGtjeHptYztrY3h6ZG07Pj47Pjt0PGk8MTI+O0A85b+F5L+u6K++O+WFrOWFsemAieS/ruivvjvkuJPkuJrpgInkv67or7475LiT5Lia5b+F5L+u6K++O+S4k+S4mumZkOmAieivvjvpgJror4blv4Xkv67or7475a2m56eR5Z+656GA6K++O+S6uuaWh+iJuuacr+mZkOmAiTvpgJror4bmlZnogrLku7vpgIk75a6e6aqM6K++O+Wunui3teaVmeWtpueOr+iKgjtcZTs+O0A8MDE7MDI7MDM7MDQ7MDU7MDY7MDc7MDg7MDk7MTA7MTE7XGU7Pj47Pjs7Pjt0PHA8cDxsPFZpc2libGU7PjtsPG88Zj47Pj47Pjs7Pjt0PHA8cDxsPFZpc2libGU7PjtsPG88Zj47Pj47Pjs7Pjt0PHA8cDxsPFZpc2libGU7PjtsPG88Zj47Pj47Pjs7Pjt0PHA8cDxsPFRleHQ7PjtsPFxlOz4+Oz47Oz47dDxwPHA8bDxUZXh0O1Zpc2libGU7PjtsPOWtpuWPt++8mjEzMTUwMTIyO288dD47Pj47Pjs7Pjt0PHA8cDxsPFRleHQ7VmlzaWJsZTs+O2w85aeT5ZCN77ya5q+b5LyfO288dD47Pj47Pjs7Pjt0PHA8cDxsPFRleHQ7VmlzaWJsZTs+O2w85a2m6Zmi77ya5L+h5oGv5bel56iL5a2m6ZmiO288dD47Pj47Pjs7Pjt0PHA8cDxsPFRleHQ7VmlzaWJsZTs+O2w85LiT5Lia77yaO288dD47Pj47Pjs7Pjt0PHA8cDxsPFRleHQ7VmlzaWJsZTs+O2w86L2v5Lu25bel56iLO288dD47Pj47Pjs7Pjt0PHA8cDxsPFRleHQ7PjtsPOS4k+S4muaWueWQkTo7Pj47Pjs7Pjt0PHA8cDxsPFRleHQ7VmlzaWJsZTs+O2w86KGM5pS/54+t77ya6L2v5Lu2UTEzNDE7bzx0Pjs+Pjs+Ozs+O3Q8QDA8cDxwPGw8VmlzaWJsZTs+O2w8bzxmPjs+Pjs+Ozs7Ozs7Ozs7Oz47Oz47dDw7bDxpPDE+O2k8Mz47aTw1PjtpPDc+O2k8OT47aTwxMz47aTwxNT47aTwxOT47aTwyMT47aTwyMj47aTwyMz47aTwyNT47aTwyNz47aTwyOT47aTwzMT47aTwzMz47aTw0MT47aTw0Nz47aTw0OT47aTw1MD47PjtsPHQ8cDxwPGw8VmlzaWJsZTs+O2w8bzxmPjs+Pjs+Ozs+O3Q8QDA8cDxwPGw8VmlzaWJsZTs+O2w8bzxmPjs+PjtwPGw8c3R5bGU7PjtsPERJU1BMQVk6bm9uZTs+Pj47Ozs7Ozs7Ozs7Pjs7Pjt0PDtsPGk8MTM+Oz47bDx0PEAwPDs7Ozs7Ozs7Ozs+Ozs+Oz4+O3Q8cDxwPGw8VGV4dDtWaXNpYmxlOz47bDzoh7Pku4rmnKrpgJrov4for77nqIvmiJDnu6nvvJo7bzx0Pjs+Pjs+Ozs+O3Q8QDA8cDxwPGw8UGFnZUNvdW50O18hSXRlbUNvdW50O18hRGF0YVNvdXJjZUl0ZW1Db3VudDtEYXRhS2V5czs+O2w8aTwxPjtpPDA+O2k8MD47bDw+Oz4+O3A8bDxzdHlsZTs+O2w8RElTUExBWTpibG9jazs+Pj47Ozs7Ozs7Ozs7Pjs7Pjt0PEAwPHA8cDxsPFZpc2libGU7PjtsPG88Zj47Pj47cDxsPHN0eWxlOz47bDxESVNQTEFZOm5vbmU7Pj4+Ozs7Ozs7Ozs7Oz47Oz47dDxAMDxwPHA8bDxWaXNpYmxlOz47bDxvPGY+Oz4+O3A8bDxzdHlsZTs+O2w8RElTUExBWTpub25lOz4+Pjs7Ozs7Ozs7Ozs+Ozs+O3Q8QDA8Ozs7Ozs7Ozs7Oz47Oz47dDxAMDxwPHA8bDxWaXNpYmxlOz47bDxvPGY+Oz4+O3A8bDxzdHlsZTs+O2w8RElTUExBWTpub25lOz4+Pjs7Ozs7Ozs7Ozs+Ozs+O3Q8QDA8cDxwPGw8VmlzaWJsZTs+O2w8bzxmPjs+PjtwPGw8c3R5bGU7PjtsPERJU1BMQVk6bm9uZTs+Pj47Ozs7Ozs7Ozs7Pjs7Pjt0PEAwPHA8cDxsPFZpc2libGU7PjtsPG88Zj47Pj47Pjs7Ozs7Ozs7Ozs+Ozs+O3Q8QDA8cDxwPGw8VmlzaWJsZTs+O2w8bzxmPjs+PjtwPGw8c3R5bGU7PjtsPERJU1BMQVk6bm9uZTs+Pj47Ozs7Ozs7Ozs7Pjs7Pjt0PEAwPHA8cDxsPFZpc2libGU7PjtsPG88Zj47Pj47cDxsPHN0eWxlOz47bDxESVNQTEFZOm5vbmU7Pj4+Ozs7Ozs7Ozs7Oz47Oz47dDxAMDw7QDA8OztAMDxwPGw8SGVhZGVyVGV4dDs+O2w85Yib5paw5YaF5a65Oz4+Ozs7Oz47QDA8cDxsPEhlYWRlclRleHQ7PjtsPOWIm+aWsOWtpuWIhjs+Pjs7Ozs+O0AwPHA8bDxIZWFkZXJUZXh0Oz47bDzliJvmlrDmrKHmlbA7Pj47Ozs7Pjs7Oz47Ozs7Ozs7Ozs+Ozs+O3Q8cDxwPGw8VGV4dDtWaXNpYmxlOz47bDzmnKzkuJPkuJrlhbE1MOS6ujtvPGY+Oz4+Oz47Oz47dDxwPHA8bDxWaXNpYmxlOz47bDxvPGY+Oz4+Oz47Oz47dDxwPHA8bDxWaXNpYmxlOz47bDxvPGY+Oz4+Oz47Oz47dDxwPHA8bDxWaXNpYmxlOz47bDxvPGY+Oz4+Oz47Oz47dDxwPHA8bDxUZXh0Oz47bDxIQlVFOz4+Oz47Oz47dDxwPHA8bDxJbWFnZVVybDs+O2w8Li9leGNlbC8xMzE1MDEyMi5qcGc7Pj47Pjs7Pjs+Pjt0PDtsPGk8Mz47PjtsPHQ8QDA8Ozs7Ozs7Ozs7Oz47Oz47Pj47Pj47Pj47PusyeoKZgpYdAAQhMwCtDaTiai6H',
            'hidLanguage': '',
            'ddLXN':'',
            'ddLXQ':'',
            'ddl_kcxz':'',
        })
    # 获得历年成绩界面的pageCode
    def getPage(self):
        try:
            request = urllib2.Request(self.baseURL, self.postData, self.headers)
            request_gra1 =urllib2.Request(self.graduURL1, headers=self.headers_gra1)
            request_gra2 = urllib2.Request(self.graduURL2, self.postData_Gra, self.headers_gra2)
            result = self.opener.open(request)
            result = self.opener.open(request_gra1)
            result = self.opener.open(request_gra2)
            return result.read().decode("gb2312")
        except urllib2.URLError, e:
            if hasattr(e,"reason"):
                print "连接错误， 错误原因"+e.reason
                return None

    def writeIntoExcel(self):
        pageCode = self.getPage()
        # 通过解析pageCode来生产解析后的代码，解析后的代码有很多属性，和dom类似，构造函数的第二个参数见代码详解
        soup = BeautifulSoup(pageCode, 'html.parser')
        # 找出第一个table标签
        table = soup.find("table", class_="datelist")
        # 创建一个Workbook对象，这就相当于创建了一个Excel文件，将编码设置成utf-8，就可以在excel中输出中文了。默认是ascii
        # style_compression表示压缩格式
        book = xlwt.Workbook(encoding="utf-8", style_compression=0)
        # 创建一个sheet对象，一个sheet对象对应Excel文件中的一张表格，这里我命名为mao，并且写入可覆盖
        sheet = book.add_sheet("mao", cell_overwrite_ok=True)

        trs = table.find("tr")
        tds = trs.find_all("td")
        col = 0
        # 存入表格第一行，即每一列的说明
        for i in range(len(tds)):
            if i == 0 or i == 1 or i == 3:
                sheet.write(0, col, tds[i].find('a').string.decode("utf-8"))
                col += 1
            if i == 12:
                sheet.write(0, col, tds[i].string.decode("utf-8"))
                col += 1
        # 存入详细成绩
        row = 0
        trs = table.find_all("tr")
        for i in range(len(trs)):
            if i > 0:
                tds = trs[i].find_all("td")
                row += 1
                col = 0
                for j in range(len(tds)):
                    if j == 0 or j == 1 or j == 3 or j == 12:
                        sheet.write(row, col, tds[j].string.decode("utf-8"))
                        col += 1
        # 最后，将以上操作保存到指定的Excel文件中
        book.save("123.xls")
        print "写入EXCEL完毕!"
baseUrl = "http://218.197.80.27/default2.aspx"
test = HBJJ(baseUrl)
test.writeIntoExcel()
