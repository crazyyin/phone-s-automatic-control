# -*- coding: utf-8 -*-

import uiautomator2 as u2
from time import sleep
import wx
import wx.xrc
import openpyxl as excel



# 文档说明
# 此程序使用有几点需要注意的地方
# 1.由于使用的是模拟器，必须打开模拟器所在目录，输入cmd 连接 127.0.0.1：7555
# 2 enter 命令和真机有区别，必须 使用两次
# 3 使用问题搜索前 需要输入全名一次 此处可以换个方法 直接自己打开相应的话题，然后再开启程序即可

# 以300次滑动为一个月计算

# 首先建立问题库

question=['在吗','还有吗','多少钱','租金','房租','多少','哪一站','多少平米','多大','月租','合租','坐标','房源','详细','私聊','位置','单人','双人','有房子吗','看房','微信','有没有','什么价','空房','多钱','在哪','有房吗']
# 建立已发客户列表

customer_sended=['游牧人Hiro', '走猫扛炮', '今天也不是好东西', 'SeasonWang_', '前门大鲤鱼', '李菲菲invicibility', '日兴小酱', '仙妮蒙', '猪居然', '龙猫不是不是不是猫', '小酒哥在东京_东京租房', '东京租赁买卖-柯南', 'HarryPotter_DD', '黄艺伟', '蘇雲依Phoniex', '-桃子组阳酱-', '令和公主', '多元希', '葵记奶茶', '网路恶霸小猫驴', '两岁丢掉的牛', 'HilaryLFrankie', '七月_lan', '一起Superlit', '平成処女', '安德烈只钟意dollar', '小玉爱吃鱼cc', '嘉訢Michelle']
zhongjie_name=['东京','租']


d=u2.connect('127.0.0.1:62001')
print(d.info)



def page_down(content):
    d.swipe(163,910,330,500)
    sleep(1)
    if d(resourceId="com.weico.international:id/item_timeline_toolbar_comment"):

        if d(resourceId="com.weico.international:id/item_timeline_toolbar_comment").get_text():
            number=int(d(resourceId="com.weico.international:id/item_timeline_toolbar_comment").get_text())
            if 1<=number<=20:
                d(resourceId="com.weico.international:id/item_timeline_toolbar_comment").click()
                sleep(2)

                for i in range(len(d(resourceId="com.weico.international:id/detail_item_content"))):
                    # 用户名
                    customer=d(resourceId="com.weico.international:id/detail_item_screenname")[i]
                    customer_name=d(resourceId="com.weico.international:id/detail_item_screenname")[i].get_text()
                    # 评论
                    comment=d(resourceId="com.weico.international:id/detail_item_content")[i].get_text()
                    send_message(customer,customer_name,comment,customer_sended,question,content)
                    print(customer_name)
                    print('第'+str(i)+'条评论')
                    sleep(1)
                d(description=u"转到上一层级").click()
                sleep(1)

    else:
        d.swipe(163,910,330,500)
def send_message(customer,name,comment,customer_sended,question,subject):
    if name in customer_sended:
        print("用户已发送")

    else:
        def judge_zhongjie(name):

            for i in range(len(zhongjie_name)):
                if zhongjie_name[i] in name:
                    return ('这是中介')
                else:
                    pass
        if judge_zhongjie(name)=='这是中介':
            pass

        else:
            for i in range(len(question)):
                if question[i] in comment:
                    if name in customer_sended:
                        pass
                    else:
                        customer_sended.append(name)
                        print(customer_sended)
                        # 点击用户名
                        customer.click()
                        # 点击私信
                        sleep(1)
                        d(resourceId="com.weico.international:id/profile_header_dm").click()
                        # 添加私信内容
                        d(resourceId="com.weico.international:id/msg_text").set_text(subject)
                        # 发送
                        d(resourceId="com.weico.international:id/send_layout").click()

                        # 返回


                        d(description=u"转到上一层级").click()
                        sleep(1)
                        d(description=u"转到上一层级").click()


# 此部分用于输出excel表
def excel_output():
    wb=excel.Workbook()
    ws=wb.active

    for i in range(len(customer_sended)):

        ws.cell(row = i+1, column = 1,value=customer_sended[i])

    wb.save('C:\\Users\\user\\Desktop\\output\\weibo.xlsx')
    print('对比完成，文档输出')






class MyFrame1 ( wx.Frame ):
    def __init__( self, parent ):
        wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = wx.EmptyString, pos = wx.DefaultPosition, size = wx.Size( 153,151 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )
        self.SetSizeHintsSz( wx.DefaultSize, wx.DefaultSize )

        bSizer1 = wx.BoxSizer( wx.VERTICAL )

        self.m_textCtrl1 = wx.TextCtrl( self, wx.ID_ANY,  u"300", wx.DefaultPosition, wx.DefaultSize, 0 )
        bSizer1.Add( self.m_textCtrl1, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )

        self.m_textCtrl2 = wx.TextCtrl( self, wx.ID_ANY, u"东京租", wx.DefaultPosition, wx.DefaultSize, 0 )
        bSizer1.Add( self.m_textCtrl2, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )

        self.m_button2 = wx.Button( self, wx.ID_ANY, u"开始", wx.DefaultPosition, wx.DefaultSize, 0 )
        bSizer1.Add( self.m_button2, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )


        self.SetSizer( bSizer1 )
        self.Layout()

        self.Centre( wx.BOTH )

        # Connect Events
        self.m_button2.Bind( wx.EVT_BUTTON, self.main )
    def __del__( self ):
        pass
    # Virtual event handlers, overide them in your derived class
    def main( self, event ):
        number = int(self.m_textCtrl1.GetLineText(1))
        subject = self.m_textCtrl2.GetLineText(1)
        print(number)
        print(subject)

        # 启动微博app


        d.app_start("com.weico.international")
# 点击搜索框
        d(resourceId="com.weico.international:id/menu_index_search").click()
# 等待搜索框出现
        d(resourceId="com.weico.international:id/act_search_input").wait(timeout=3.0)
# 点击搜索框 输入内容
        d(resourceId="com.weico.international:id/act_search_input").set_text(subject)
# 点击回车 搜索内容
        sleep(2)
# 此处需使用两次 进入 才可以
        d.press("enter")
        d.press("enter")

        sleep(2)
        for i in range(number):
            page_down('亲，您好，我专业为留学生提供房产租赁服务，有大量房源可供您选择。亲，可以加我V信：ay125890 详聊。')
            print('滑动'+str(i)+'次')
        excel_output()



app = wx.App(False)
frame = MyFrame1(None)
frame.Show(True)
#start the applications
app.MainLoop()
