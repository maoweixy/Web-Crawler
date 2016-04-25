# -*- coding:utf-8 -*-

import urllib2
import re
import sys
reload(sys)
sys.setdefaultencoding("utf-8")
class Tool:
    # delete img tag
    removeImg = re.compile('<img.*?>')
    # delete hyperlink tag
    removeLink = re.compile('<a.*?>|</a>')
    # replace line code with "\n"
    replaceLine = re.compile('<tr>')
    replaceTD = re.compile('<td>')
    replacePara = re.compile('<p.*?>')
    replaceBR = re.compile('<br><br>|<br>')
    removeExtraTag = re.compile('<.*?>')
    def replace(self, x):
        x = re.sub(self.removeImg, "", x)
        x = re.sub(self.removeLink, "", x)
        x = re.sub(self.replaceLine, "\n", x)
        x = re.sub(self.replaceTD, "\t", x)
        x = re.sub(self.replacePara, "\n    ", x)
        x = re.sub(self.replaceBR, "\n", x)
        x = re.sub(self.removeExtraTag, "", x)
        return x.strip()

class BDTB:
    def __init__(self, baseUrl, seeLz):
        self.tool = Tool()
        self.baseURL = baseUrl
        self.seeLZ = "?see_lz="+str(seeLz)
        self.user_agent = "Mozilla/4.0 (compatible; MISE 5.5; windows NT)"
        self.headers = {"User-Agent": self.user_agent}
        self.file = None
    def getPage(self, pageNum):
        try:
            url = self.baseURL+self.seeLZ+"&pn="+str(pageNum)
            request = urllib2.Request(url, headers=self.headers)
            response = urllib2.urlopen(request)
            pageCode = response.read().decode("utf-8")
            return pageCode
        except urllib2.URLError, e:
            if hasattr(e,"reason"):
                print "连接错误，错误原因：" + e.reason
                return None

    def setTitle(self, page):
        pattern = re.compile(r'<h3 class="core_title_txt.*?>(.*?)</h3>', re.S)
        result = re.search(pattern,page)
        title = result.group(1).strip()
        if result:
            self.file = open(title+".txt", "w+")
            content = result.group(1).strip()+"\n"
            self.file.write(content.encode("utf-8"))

    def getPageNum(self, pageCode):
        pattern = re.compile(r'<li class="l_reply_num.*?<span class="red.*?>(.*?)</span>.*?'
                             r'<span class="red.*?>(.*?)</span>', re.S)
        result = re.search(pattern,pageCode)
        content = result.group(1).strip()+"回复贴"+",共"+result.group(2).strip()+"页\n"
        self.file.write(content.encode("utf-8"))
        return result.group(2)
        # if result:
        #     print
        #     # print result.group(1).strip()
        # else:
        #     print "123"

    def getContent(self, pageCode):
        i = 1;
        pattern = re.compile(r'<div id="post_content.*?>(.*?)</div>', re.S)
        items = re.findall(pattern, pageCode)
        contents = []
        for item in items:
            content = "\n"+self.tool.replace(item)+"\n"
            contents.append(content.encode("utf-8"))
        return contents

    # def setFileTitle(self, title):
    #     if title is not None:
    #         self.file = open(title+".txt", "w+")
    #     # else:
    #     #     self.file = open("百度贴吧"+".txt", "w+")

    def writeDtat(self, content,page):
        i = 1
        for item in content:
            floorLine = "\n"+"第"+str(page)+"页————"+str(i)+"楼:"
            i+=1
            self.file.write(floorLine)
            self.file.write(item)

    def start(self):
        firstPage = self.getPage(1)
        self.setTitle(firstPage)
        pageNum = self.getPageNum(firstPage)
        if pageNum == None:
            print "URL已经失效"
            return
        try:
            print "该帖子共有"+pageNum+"页"
            for i in range(1, int(pageNum)+1):
                print "正在写入第"+str(i)+"页"
                pageCode = self.getPage(i)
                content = self.getContent(pageCode)
                self.writeDtat(content,i)
        except IOError, e:
            print "写入异常，原因是:" + e.message
        finally:
            print "任务完成！"
            self.file.close()
baseURL = 'http://tieba.baidu.com/p/3138733512'
test = BDTB(baseURL,1)
test.start()