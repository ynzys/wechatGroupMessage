# -*- coding: utf-8-*-
import subprocess
import uiautomation as auto
import time
import tkinter . filedialog
import tkinter . messagebox
import os 
import csv

class WxContacts(object):
    '''
    从微信获取联系人的列表类
    wechat_exe_dir:微信程序exe所在绝对目录
    class_name:微信主窗口的类名，默认'WeChatMainWndForPC'
    window_name:窗口名称
    后两个名字在windows系统下，默认'微信'
    在windows系统可以使用 Accessibility Insights For Windows 程序获得。下载地址百度

    '''
    # 构造函数
    def __init__(self):
        wxconfig_csv = '__wxconfig__.csv'
        csv_lines = []
        # 从文件读取微信配置
        with open(wxconfig_csv, 'r', encoding='utf-8') as f:
            csv_reader = csv.reader(f)
            for L in csv_reader:
                csv_lines.append(L[0])

        self.wechat_exe_dir = csv_lines[0]
        self.class_name = csv_lines[1]
        self.window_name =csv_lines[2]
    
    # 通过手动点击的方式 获取联系人列表
    def GetContactsListByClick(self,isneedsave = False):
        '''
            通过手动点击的方式 获取联系人列表
            左键点击 某联系人，会获取当前联系人列表中第二个联系人。
            按下空格键后停止选择，程序退出
        '''
        # 启动微信，为了防止登录耽误时间。需要手动事先启动最好
        subprocess.Popen(self.wechat_exe_dir)
        #time.sleep(10) # 预留10s登录时间，如果手动已经实现启动，则不需要
        time.sleep(0.5) # 等待0.5s
        wechat_Window = auto.WindowControl(searchDepth=1, className=self.class_name, Name=self.window_name)
        
        # 点击通讯录
        button = wechat_Window.ButtonControl(Name='通讯录')
        button.Click()
        # 
        # 请事先手动定位联系人列表
        list_control_contacts = wechat_Window.ListControl(Name="联系人")
        #list_control_contacts.MoveCursorToMyCenter()

        # list_contacts 表示联系人名称列表，count 表示计数器,flag_continue 控制代码是否退出
        list_contacts = []
        count = 1
        flag_continue = True # True表示持续循环，Flase，表示停止

        start_time = time.time()
        #print("""开始时间：{}""".format(start_time))

        while flag_continue:
            # 空格，判断鼠标点击动作发生
            
            if auto.IsKeyPressed(auto.Keys.VK_A): # 按下A键，代表添加
            # 尝试根据鼠标位置来判断当前联系人是第几个？
                mouse_y = auto.GetCursorPos()[1]# 获得鼠标的y位置
                for control_contact in list_control_contacts.GetChildren():
                    if mouse_y<=control_contact.BoundingRectangle.bottom and mouse_y>= control_contact.BoundingRectangle.top:
                        # 落在某个联系人上面
                        if control_contact.Name not in list_contacts:
                            if control_contact.Name !='':
                                #print(count, '', nick_control.Name)
                                list_contacts.append(control_contact.Name)
                                count += 1
            
            if auto.IsKeyPressed(auto.Keys.VK_E): #按下E键，停止获取
                flag_continue = False
                #tkinter.messagebox.showinfo('提示','好友列表获取完成,请关闭微信窗口!')
        if isneedsave:
            self.contacts2csv(list_contacts=list_contacts)
            tkinter.messagebox.showinfo('提示','好友列表获取完成，请查看__contacts__.csv!')

        return list_contacts #返回列表

    # 通过自动滚动的方式获取所有联系人。
    def GetContactsAllAuto(self,wheel_pause_time=0.5,isneedsave = False):
        '''
            自动获取列表中所有联系人
        '''
        # 启动微信，为了防止登录耽误时间。需要手动事先启动最好
        subprocess.Popen(self.wechat_exe_dir)
        #time.sleep(10) # 预留10s登录时间，如果手动已经实现启动，则不需要
        time.sleep(0.5) # 等待0.5s
        wechat_Window = auto.WindowControl(searchDepth=1, className=self.class_name, Name=self.window_name)
        
        # 点击通讯录
        button = wechat_Window.ButtonControl(Name='通讯录')
        button.Click()
        
        # 请事先手动定位联系人列表
        list_control_contacts = wechat_Window.ListControl(Name="联系人") #  我理解是得到了联系人列表的指针
        list_control_contacts.MoveCursorToMyCenter()

        # list_contacts 表示联系人名称列表，count 表示计数器,flag_continue 控制代码是否退出
        list_contacts = []
        count = 1
        flag_continue = True # True表示持续循环，Flase，表示停止

        start_time = time.time()
        #print("""开始时间：{}""".format(start_time))

        while flag_continue:

            # 所以采用下面的方案，直接取GetChildren()[0]的Name属性，不管是否是ListItemControl都能取出Name，正确运行
            nick_control = list_control_contacts.GetChildren()[0] # 注意当该Item到最顶端的时候才生效,GetChildren这个只能获得名内的列表
            
            '''判断是不是在列表list_contacts中，如果列表中没有，则添加'''
            if nick_control.Name not in list_contacts:
                if nick_control.Name !='':
                    #print(count, '', nick_control.Name)
                    list_contacts.append(nick_control.Name)
                    count += 1
            # 滚轮下滚
            auto.WheelDown(waitTime=wheel_pause_time)
            if auto.IsKeyPressed(auto.Keys.VK_E):# 按下E键终止
                #tkinter.messagebox.showinfo('提示','好友列表获取完成,请关闭微信窗口!')
                flag_continue = False
                
            
            # 手动实现滚轮滚动到底操作
            # 空格，要根据情况手动敲入 空格实现结束动作。
            if auto.IsKeyPressed(auto.Keys.VK_SPACE):
                #print("到底了")

                for nick_control in list_control_contacts.GetChildren()[1:]:# 屏幕列表中剩余的昵称遍历取出
                    nickname = nick_control.Name
        
                    '''判断是不是在列表a中，如果列表中没有，则添加'''
                    if nick_control.Name not in list_contacts:
                        if nick_control.Name !='':
                            #print(count, '', nick_control.Name)
                            list_contacts.append(nick_control.Name)
                            count += 1
                #print("""运行时间：{}s""".format(sum_time)) 
                #tkinter.messagebox.showinfo('提示','好友列表获取完成,请关闭微信窗口!')       
                flag_continue = False
       
        if isneedsave:
            self.contacts2csv(list_contacts=list_contacts)
            tkinter.messagebox.showinfo('提示','好友列表获取完成，请查看__contacts__.csv!')
        return list_contacts #返回列表
    # 把列表写入文件
    def contacts2csv(self,contacts_csv_file='__contacts__.csv',list_contacts=[],blessword='祝您节日快乐！'): 
        with open(contacts_csv_file, 'w', encoding='utf-8') as f:
            f.write('')
        with open(contacts_csv_file, 'a', encoding='utf-8') as f:  
            for contact in list_contacts:
                f.write(contact + '%' + contact + '%' + blessword  + '\n')
        