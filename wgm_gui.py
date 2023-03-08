# -*- coding: utf-8-*-
from tkinter import *
from tkinter.ttk import *
import tkinter . filedialog
import tkinter . messagebox
import csv
import time
import os 
import sys
from get_contacts_from_wx import *
import subprocess
import uiautomation as auto
import pyperclip
from PIL import Image, ImageTk
import requests
from io import BytesIO

# GUI 界面对 好友列表进行管理，并提供一系列辅助功能。

class WxGroupMsgGUI(object):
    '''
    界面GUI类
    '''
    def __init__(self):
        # 
        self.initWindow()
        self.readWxConfig()
        self.drawMenu()
        self.drawToolbar()
        self.drawTableFrame()
        self.drawBlessworsFrame()
        self.drawTable()
        self.drawBlessworsTable()
        #self.drawMoneyImage() # 不打赏版本
        
    
    # 初始化窗口大小位置，设定布局组件和表格的宽度大小。
    def initWindow(self):
        self.win = Tk() # 创建一个主窗口对象
        # 自动设置好窗口的大小
        # 获取屏幕大小
        screen_width = self.win.winfo_screenwidth()
        screen_height = self.win.winfo_screenheight()
        # 计算窗口的大小比屏幕大小 各缩进80
        self.window_width = screen_width - 120
        self.window_height = screen_height - 120
        # 设置窗口内部的布局组件Frame的大小
        self.frame_width = self.window_width-10 # 不用设置高度height，高度会根据内容自动调整
        # 设置表格组件table的大小
        self.table_width = self.frame_width - 10  # 表格的高度height，暂不设置。？？todo

        # 初始化窗口
        geostr = str(self.window_width) + 'x' + str(self.window_height) + '+10+10'
        self.win.geometry(geostr) # 初始化大小
        self.win.resizable(False,False) # 宽度和高度均不能修改 
        self.win.iconbitmap('icon.ico')
        self.inti_title = '微信祝福群发助手3.0'
        self.win.title(self.inti_title)

        self.opened_file_name = ''# 当前被打开的文件名称
        self.wxconfig_csv = '__wxconfig__.csv'
        self.wxapp =''

    # 绘制主菜单#######################################################################
    def drawMenu(self):
         
        self.main_menu = Menu (self.win) # 主菜单
        # 文件菜单
        self.file_menu = Menu(self.win,tearoff=False) # 文件菜单
        self.main_menu.add_cascade(label='文件',menu=self.file_menu) # 把文件菜单配置在主菜单
        self.file_menu.add_command (label="新建文件",command=self.menuNewFile) # 菜单项绑定方法
        self.file_menu.add_command (label="打开文件",command=self.menuOpenFile) # 菜单项绑定方法
        self.file_menu.add_command (label="关闭文件",command=self.menuCloseFile) # 菜单项绑定方法
        self.file_menu.add_command (label="文件保存",command=self.menuSaveFile) # 菜单项绑定方法
        self.file_menu.add_command (label="文件另存为",command=self.menuSaveAsFile) # 菜单项绑定方法
        self.file_menu.add_separator()
        self.file_menu.add_command (label="从微信选择获取",command=self.getWX) # 菜单项绑定方法
        self.file_menu.add_command (label="从微信自动获取",command=self.getWXAll) # 菜单项绑定方法
        self.file_menu.add_separator()

        self.file_menu.add_command (label="程序退出",command=self.menuCommandExit) # 菜单项绑定方法

        # 帮助菜单
        self.help_menu = Menu(self.win,tearoff=False) # 帮助菜单
        self.main_menu.add_cascade(label='帮助配置',menu=self.help_menu) # 把文件菜单配置在主菜单
        self.help_menu.add_command (label="参数配置",command=self.wxConfig) # 菜单项绑定方法
        self.help_menu.add_separator()
        self.help_menu.add_command (label="使用说明",command=self.helpfile) # 菜单项绑定方法
        self.help_menu.add_command (label="代码说明",command=self.codefile) # 菜单项绑定方法
        self.help_menu.add_separator()
        self.help_menu.add_command (label="版权信息",command=self.author) # 菜单项绑定方法
        #显示菜单对象
        self.win.config (menu = self.main_menu)   
        
        # 定义关闭按钮的事件
        self.win.protocol('WM_DELETE_WINDOW',self.menuCommandExit)
       
    # 绘制工具栏#######################################################################
    def drawToolbar(self):
        
        self.toolbar_frame = Frame(self.win,relief=RAISED,borderwidth=1)  # 这是第一个菜单栏下面第一个Frame
        # 工具栏按钮
        
        button_width = 10
       
        button1 = Button(self.toolbar_frame,text='New',command=self.menuNewFile,width=button_width)
        button1.pack(side=LEFT,padx=3) 

        button1 = Button(self.toolbar_frame,text='Open',command=self.menuOpenFile,width=button_width)
        button1.pack(side=LEFT,padx=3)  

        button1 = Button(self.toolbar_frame, text='Close',command=self.menuCloseFile,width=button_width)
        button1.pack(side=LEFT,padx=3) 

        button1 = Button(self.toolbar_frame,text='Save',command=self.menuSaveFile,width=button_width)
        button1.pack(side=LEFT,padx=3) 

        button1 = Button(self.toolbar_frame, text='SaveAs',command=self.menuSaveAsFile,width=button_width)
        button1.pack(side=LEFT,padx=3) 

        button1 = Button(self.toolbar_frame,text='GetWX',command=self.getWX,width=button_width)
        button1.pack(side=LEFT,padx=3) 

        button1 = Button(self.toolbar_frame,text='GetWXAll',command=self.getWXAll,width=button_width)
        button1.pack(side=LEFT,padx=3) 

        button1 = Button(self.toolbar_frame,text='Exit',command=self.menuCommandExit,width=button_width)
        button1.pack(side=LEFT,padx=3)  
        
        # 工具栏按钮绘制完毕后 pack 工具栏
        self.toolbar_frame.pack(side=TOP,fill=X) # 
        
    # 从微信列表获取好友 todo
    def getWX(self):
        # 不用先存盘了，直接在后面追加吧，这样比较好。有重复的就不添加。
        # 首先获取列表中的 nicknames列表，用来防止重复
        nicknames = []
        for item in self.tree_contacts.get_children():
            nickname = self.tree_contacts.item(item, 'values')[1]
            nicknames.append(nickname)
        
        wxcts = WxContacts()
        list_contacts = wxcts.GetContactsListByClick()
        for contact in list_contacts:
            if contact not in nicknames:# 如果不重复
                self.tree_contacts.insert("",index=END,values=(1,contact,contact,"祝您节日快乐！"))
        self.refreshIndex()
    
    # 从微信列表获取所有好友,可用
    def getWXAll(self):
        # 不用先存盘了，直接在后面追加吧，这样比较好。有重复的就不添加。
        # 首先获取列表中的 nicknames列表，用来防止重复
        nicknames = []
        for item in self.tree_contacts.get_children():
            nickname = self.tree_contacts.item(item, 'values')[1]
            nicknames.append(nickname)
        
        wxcts = WxContacts()
        list_contacts = wxcts.GetContactsAllAuto()
        for contact in list_contacts:
            if contact not in nicknames:# 如果不重复
                self.tree_contacts.insert("",index=END,values=(1,contact,contact,"祝您节日快乐！"))
        self.refreshIndex()

    # 菜单对应的命令
    # 新建文件
    def menuNewFile(self):
        # 询问下是否需要被保存？
        self.isSaveAs()
         
        #然后重新绘制table
        self.drawTable()
        
    # 打开文件
    def menuOpenFile(self):
        self.isSaveAs()
        currentdir = os.getcwd()
        self.opened_file_name  = tkinter.filedialog.askopenfilename(initialdir=currentdir,defaultextension='friends-001',filetypes=[("csv文件", ".csv")])
        self.drawTable(open_csv_file=self.opened_file_name)
 
    # 保存当前文件
    def menuSaveFile(self):
        if self.opened_file_name =='':
            currentdir = os.getcwd()
            self.opened_file_name = tkinter.filedialog.asksaveasfilename(initialdir=currentdir,
                                                                        defaultextension='contatcs-001',
                                                                        filetypes=[("csv文件", ".csv")])
        
        self.saveFileAs()
        self.win.title(self.inti_title + '——'+ self.opened_file_name)

    # 另存当前文件
    def menuSaveAsFile(self):
        currentdir = os.getcwd()
        self.opened_file_name = tkinter.filedialog.asksaveasfilename(initialdir=currentdir,
                                                                        defaultextension='contatcs-001',
                                                                        filetypes=[("csv文件", ".csv")])
        self.saveFileAs()
        self.win.title(self.inti_title + '——'+ self.opened_file_name)

    # 关闭文件
    def menuCloseFile(self):
        self.isSaveAs()
        # 首先 清楚掉table_frame中的所有子控件
        for widget in self.table_frame.winfo_children():
            widget.destroy()
        self.opened_file_name = ''
        self.win.title(self.inti_title)
    
    # 配置参数
    def wxConfig(self):
        subprocess.Popen(['notepad.exe', '__wxconfig__.csv'])
    
    # 打开帮助
    def helpfile(self):
        
        subprocess.Popen(['notepad.exe', 'help.txt'])

    # 打开帮助
    def codefile(self):
        subprocess.Popen(['notepad.exe', 'codefile.txt'])

    # 版权信息
    def author(self):
        tkinter.messagebox.showinfo('提示','该程序不得用于商业以及非法用途，如因使用被封号，本人概不负责！作者微信公众号：雷斯林日记，联系可关注公号后留言，谢谢！')

    # 退出程序
    def menuCommandExit(self):
        self.isSaveAs()
        self.win.destroy()
        sys.exit()
    
    # 询问是否需要保存，如有需要则保存
    def isSaveAs(self):
        # 询问下是否需要被保存？
        is_need_save = tkinter.messagebox.askyesno(title='是否需要保存', message='请问当前文件是否需要保存？')   
        if is_need_save: # 返回为True
            # 先保存
            # 判断当前已经打开的文件名称存在，则存在现有的文件中，否则弹出文件输入框后再保存。
            if self.opened_file_name =='':
                # 如果已经打开的文件为空，表示新建文件后没有保存过，要先弹出保存路径文件输入
                currentdir = os.getcwd()
                self.opened_file_name = tkinter.filedialog.asksaveasfilename(initialdir=currentdir,defaultextension='contatcs-001',filetypes=[("csv文件", ".csv")])
            
            self.saveFileAs()
            tkinter.messagebox.showinfo('提示','文件已经保存到'+self.opened_file_name) #
            # 进行保存
            
    
    # 绘制表格所在的Frame空间#########################################
    def drawTableFrame(self):
        self.table_frame = LabelFrame(self.win,width=self.frame_width,height=50,text='联系人列表')
        self.table_frame.pack()
    
    # 表格绘制
    def drawTable(self,open_csv_file=''):
        # 
        # 首先 清楚掉table_frame中的所有子控件
        for widget in self.table_frame.winfo_children():
            widget.destroy()
        
        xbar = Scrollbar(self.table_frame, orient='horizontal')  # 水平滚动条
        ybar = Scrollbar(self.table_frame,orient='vertical')  #垂直滚动条

        # 然后在进行重绘表格并填充数据
        # 绘制表格,容器：self.table_frame 滚动条？todo，水平滚动条没有作用。最后再解决。因为一般祝福语没有那么长。
        self.tree_contacts = Treeview(self.table_frame,columns=('index','contacts','honornames','blesswords'),show='headings',
                            height=15,
                             yscrollcommand=ybar.set,xscrollcommand=xbar.set) # height指行数，不是行高。
        ybar['command'] = self.tree_contacts.yview
        xbar['command'] = self.tree_contacts.xview
        self.tree_contacts.grid(row=0, column=0,columnspan=2,sticky=NSEW) # columnspan=2合并单元格，横跨2列
        ybar.grid(row=0, column=2, sticky='ns')   #上下对齐
        xbar.grid(row=1, column=0,columnspan=2,sticky='we')  #左右对齐
        
        
        # 设置标题
        self.tree_contacts.heading('#1',text='序号')
        self.tree_contacts.heading('#2',text='好友')
        self.tree_contacts.heading('#3',text='尊称')
        self.tree_contacts.heading('#4',text='祝福词')
        # 格式化栏位.注意它的对齐方式，中间是CENTER，其他按照上N下S左W右E，好玩。
        self.tree_contacts.column("#1",anchor=E,width=50)     
        self.tree_contacts.column("#2",anchor=CENTER,width=200)
        self.tree_contacts.column("#3",anchor=CENTER,width=200)
        self.tree_contacts.column("#4",anchor=W,width=self.table_width-450) # 注意宽度=table总宽度减去前面散列的宽度
        
        # 新建情况，默认填充1行数据
        if open_csv_file =='':# 如果打开文件选项为空，表示是新建立文件，默认一行即可
            self.tree_contacts.insert("",index=END,values=(1,self.default_nickname,"尊称","祝福词"))
            self.opened_file_name=''
            self.win.title(self.inti_title + '——未命名文件')
        else:# 打开指定文件情况
            self.opened_file_name = open_csv_file
            self.win.title(self.inti_title + '——' + self.opened_file_name)
            nicknames = []
            honor_names=[]
            blesswords=[]
            with open(self.opened_file_name, 'r',encoding='utf-8') as f:
                csv_reader = csv.reader(f,delimiter='%')

                for L in csv_reader:# 获得本行的字符串列表
                    # process each L:L是行数据列表，比如['雷斯林', '亲爱的家人', '在新春佳节之际，越胜祝福大家身体健康，家庭幸福！/::)']
                    if L[0][0] != '#' : #行数据列表的第一个字符串的第一个字母是#的时候，是注释。不是#的进行处理。    
                        nicknames.append(rlstrip(L[0]))   # 获取用户昵称,去除昵称前后的空格
                        honor_names.append(rlstrip(L[1])) # 尊称
                        blesswords.append(rlstrip(L[2]))  # 祝福语
            row_count = len(nicknames)
            for i in range(row_count): # 倒序遍历
                self.tree_contacts.insert("",index=END,values=(i+1,nicknames[i],honor_names[i],blesswords[i]))
        
        # 定义鼠标右键菜单
        self.right_menu = Menu(self.tree_contacts,tearoff=False)
        self.right_menu.add_command(label='添加记录',command=self.treeContactsAdd)
        self.right_menu.add_command(label='删除记录',command=self.treeContactsDelete)
        self.right_menu.add_separator()
        self.right_menu.add_command(label='发送选中',command=self.sendSelectedMsg)
        self.right_menu.add_command(label='发送全部',command=self.sendAllMsg)
        self.right_menu.add_separator()

        self.right_menu.add_command(label='选中全部',command=self.selectAll)
        self.right_menu.add_command(label='清空选择',command=self.clearSelected) # 全选，反选，<todo>
        self.right_menu.add_separator()

        self.right_menu.add_command(label='批量修改',command=self.batchModify) # 

        self.tree_contacts.bind('<Button-3>',self.onTreeContactsRightClick)

        # 定义双击事件，为编辑状态,参考文章：https://blog.csdn.net/falwat/article/details/127494533
        # 创建`self.delegate_var`用于绑定代理部件
        self.delegate_var = StringVar()
        # 将`self.tree_item_edit()`函数绑定到`self.tree`的`<Double-1>`事件上.
        self.tree_contacts.bind('<Double-1>', func=self.tree_item_edit)

    # 双击编辑相关方法
    def tree_item_edit(self, e: Event):
        # 获取选中的项目
        self.selected_item = self.tree_contacts.selection()[0]

        for i, col in enumerate(self.tree_contacts['columns']):
            # 获取选中项目中每列的x偏移, y偏移, 宽度 和 高度
            
            x, y, w, h =  self.tree_contacts.bbox(self.selected_item, col)
            # 判断鼠标点击事件的位置, 是否在选中项目的某列中
            if x < e.x < x + w and y < e.y < y + h:
                self.selected_column = col
                text = self.tree_contacts.item(self.selected_item, 'values')[i]
                break
            else:
                # 鼠标点击位置不在数据列中
                self.selected_column = None
                # 获取该项目的x偏移, y偏移, 宽度 和 高度
                x, y, w, h =  self.tree_contacts.bbox(self.selected_item)
                text = self.tree_contacts.item(self.selected_item, 'text')
        
        # 设置 `self.delegate_var` 的值
        self.delegate_var.set(text)
        # 如果选中列不是index列：
        if self.selected_column !='index':
            self.delegate_widget = Entry(self.tree_contacts, width=w // 10, textvariable=self.delegate_var)
            
            self.delegate_widget.bind('<FocusOut>', func=self.tree_item_edit_done) # 失去焦点
            self.delegate_widget.bind('<Return>', func=self.tree_item_edit_done)# 或者敲入回车
            self.delegate_widget.selection_range(0, END) # 设置文本为选中状态
            # 将 self.delegate_widget 放到 `self.tree` 的网格布局中, `padx` 和 `pady` 分别给选中单元格的 `x` 偏移 和 `y` 偏移
            self.delegate_widget.place(x=x,y=y) #  这样方位置才是对的。用grid不行
            self.delegate_widget.focus()
    
    # 编辑完成
    def tree_item_edit_done(self, e):
        # 将`self.delegate_widget`从`self.tree`的网格布局中移除
        self.delegate_widget.place_forget()
        self.tree_contacts.set(self.selected_item, self.selected_column, self.delegate_var.get())
    
    # 读取微信配置
    def readWxConfig(self):
        csv_lines = []
        # 从文件读取微信配置
        with open(self.wxconfig_csv, 'r', encoding='utf-8') as f:
            csv_reader = csv.reader(f)
            for L in csv_reader:
                csv_lines.append(L[0])
        
        self.wechat_exe_dir = csv_lines[0]
        self.class_name=csv_lines[1]
        self.window_name=csv_lines[2]
        self.default_nickname = csv_lines[3]
    
    # 发送选中消息,在这里直接运行的话，焦点问题已经搞定。
    def senddMsg(self,all = False):
        interval = 0.5
        list_contacts = self.tree_contacts.selection()
        if all:# 如果发送全部
            list_contacts = self.tree_contacts.get_children()
        for itemid in list_contacts: #一旦开始发送，无法终止。
            
            nickname = self.tree_contacts.item(itemid,'values')[1]
            honorname = self.tree_contacts.item(itemid,'values')[2]
            blessword = self.tree_contacts.item(itemid,'values')[3]
            msg = honorname + ',' + blessword
            self.wxapp=subprocess.Popen(self.wechat_exe_dir) # 使用这个就能搞定失去焦点问题了。2023-1-30日good
        
            time.sleep(interval)
            auto.SendKeys('{Ctrl}F')

            time.sleep(interval)
            pyperclip.copy(nickname)
            auto.SendKeys('{Ctrl}V')


            time.sleep(interval)
            auto.SendKeys('{Enter}')

            time.sleep(interval)
            pyperclip.copy(msg)
            auto.SendKeys('{Ctrl}V')

            time.sleep(interval)
            auto.SendKeys('{Enter}')
        tkinter.messagebox.showinfo('提示','信息发送完毕！')
        # 隐藏微信窗口
        subprocess.Popen.terminate(self.wxapp)
        
    # 发送选中记录
    def sendSelectedMsg(self):
        self.senddMsg()

    # 发送全部记录
    def sendAllMsg(self):
        self.senddMsg(all=True)
    
    # 全部选中
    def selectAll(self):
        self.clearSelected()
        for item in self.tree_contacts.get_children():
            self.tree_contacts.selection_add(item) 
        
    # 在当前行上面添加一行,运行在右键之后。如果不在右键之后的话，没法取得index
    def treeContactsAdd(self):
        current_index = self.tree_contacts.index(self.current_itemID) # 根据itemID取得 index
        self.tree_contacts.insert("",index=current_index,values=('1',self.default_nickname,"尊称","祝福词"))
        # 修改后面的全部 把序号都加1
        # 从头全部刷新序号
        self.refreshIndex()
    
    # 通过菜单命令添加,默认添加在第一行
    '''
    def contactsAdd(self):
        self.tree_contacts.insert("",index=0,values=('1',self.default_nickname,"请填入尊称","请在这里填入祝福词"))
        # 修改后面的全部 把序号都加1
        # 从头全部刷新序号
        self.refreshIndex()
    '''

    # 删除选中的行
    def treeContactsDelete(self):
        for itemid in self.tree_contacts.selection():
            self.tree_contacts.delete(itemid)
        # 从头全部刷新序号
        self.refreshIndex()

    # 清空选择
    def clearSelected(self):
        for item in self.tree_contacts.selection():
            self.tree_contacts.selection_remove(item) # clear 方法不对，要用remove来清除选择
    
    # 批量修改选中的祝福语
    def batchModify(self):
        if self.tree_blesswords.selection()==():
            tkinter.messagebox.showinfo('提示','请先单击选择一条祝福词，然后再执行该操作！')
        else:#
            selected_item_id1 = self.tree_blesswords.selection()[0]
            blessword = self.tree_blesswords.item(selected_item_id1,'values')[1]
            # 对选中或者所有进行批量修改，如果没有选中默认全部修改
            target_items = self.tree_contacts.selection()
            if target_items ==():
                target_items = self.tree_contacts.get_children()

            for item in target_items:
                self.tree_contacts.set(item,column='blesswords',value=blessword)
    
    # 刷新序号
    def refreshIndex(self):
        # 
        for item in self.tree_contacts.get_children():
            index_item = self.tree_contacts.index(item)
            self.tree_contacts.set(item, column='index', value=index_item+1)
        
    
    # 表格上鼠标右键事件
    def onTreeContactsRightClick(self,event):
        
        self.current_itemID = self.tree_contacts.identify_row(event.y) # 事件所在行ID
        selected_itemids = self.tree_contacts.selection() # ('I001','I002'),返回是一个数组
        if self.current_itemID not in selected_itemids: # 如果当前行不是被选中状态，则设定当前行为选中状态，
            self.tree_contacts.selection_set(self.current_itemID)
        
        self.right_menu.tk_popup(event.x_root, event.y_root, 0)
    
    # 保存文件
    def saveFileAs(self):
        # 清空指定文件内容
        with open(self.opened_file_name, 'w', encoding='utf-8') as f:
            f.write('') # 清空
        # 写入全部数据
        Line_space_list =[]# 记录空行的行号
        
        with open(self.opened_file_name, 'a', encoding='utf-8') as f:
            tail_str = '\n'
            for item in self.tree_contacts.get_children():
                item_values = self.tree_contacts.item(item,'values')
                nickname = item_values[1]
                honorname = item_values[2]
                blessword = item_values[3]
                if  nickname.isspace() or honorname.isspace() or blessword.isspace() or nickname=='' or honorname=='' or blessword=='': 
                    #数据为空的时候记录,不写入
                    item_index = self.tree_contacts.index(item)
                    Line_space_list.append(item_index)
                else:# 写入
                    f.write(nickname + '%' + honorname + '%' + blessword + tail_str) # 同时把昵称当作尊称写入。为后续管理做准备。
                #
        # 表格中空行的数量，如果有空行，空行不保存
        line_space_num = len(Line_space_list)
        if line_space_num > 0 :
            strmeg = '第'
            for i in range(line_space_num):
                strlineindex = Line_space_list[i]
                strmeg += str(strlineindex)+ ','
            strmeg += '行内容没有被存储，有空值！' 
            tkinter.messagebox.showinfo('发送提示',strmeg)

    # 绘制祝福词管理框架和表格###########################################################
    # 
    def drawBlessworsFrame(self):
        self.bless_frame = LabelFrame(self.win,width=self.frame_width,height=50,text='祝福词列表')
        self.bless_frame.pack()
    
    # 绘制祝福词表格，为方便起见，本设计不在对祝福词的数据库进行多份管理。就一份就可以了。没有必要做多个。默认名字就是__blesswords.csv_
    # 格式：每行就是一个祝福词就可以了。
    def drawBlessworsTable(self):
        xbar = Scrollbar(self.bless_frame, orient='horizontal')  # 水平滚动条
        ybar = Scrollbar(self.bless_frame,orient='vertical')  #垂直滚动条

        # 然后在进行重绘表格并填充数据
        # 绘制表格,容器：self.table_frame 滚动条？todo，水平滚动条没有作用。最后再解决。因为一般祝福语没有那么长。
        self.tree_blesswords = Treeview(self.bless_frame,columns=('index','blesswords'),show='headings',
                            height=5,
                             yscrollcommand=ybar.set,xscrollcommand=xbar.set) # height指行数，不是行高。
        ybar['command'] = self.tree_blesswords.yview
        xbar['command'] = self.tree_blesswords.xview
        self.tree_blesswords.grid(row=0, column=0,columnspan=2,sticky=NSEW) # columnspan=2合并单元格，横跨2列
        ybar.grid(row=0, column=2, sticky='ns')   #上下对齐
        xbar.grid(row=1, column=0,columnspan=2,sticky='we')  #左右对齐

        # 设置标题
        self.tree_blesswords.heading('#1',text='序号')
        self.tree_blesswords.heading('#2',text='祝福词')
        
        # 格式化栏位.注意它的对齐方式，中间是CENTER，其他按照上N下S左W右E，好玩。
        self.tree_blesswords.column("#1",anchor=E,width=50)     
        self.tree_blesswords.column("#2",anchor=W,width=self.table_width-50) # 注意宽度=table总宽度减去前面散列的宽度

        # 打开文件
        self.belsswords_csv = '__blesswords__.csv'
        with open(self.belsswords_csv, 'r',encoding='utf-8') as f:
            csv_reader = csv.reader(f)
            i = 1
            for L in csv_reader:# 获得本行的字符串列表
                self.tree_blesswords.insert("",index=END,values=(i,L[0]))
                i += 1
        # 右键菜单
        # 定义鼠标右键菜单
        self.right_menu_bless = Menu(self.tree_blesswords,tearoff=False)
        self.right_menu_bless.add_command(label='添加记录',command=self.treeBlessAdd)
        self.right_menu_bless.add_command(label='删除记录',command=self.treeBlessDelete)
        self.right_menu_bless.add_separator()
        self.right_menu_bless.add_command(label='清空选择',command=self.treeBlessCLearSelected) # 全选，反选，<todo>
        self.right_menu_bless.add_separator()

        self.right_menu_bless.add_command(label='批量修改',command=self.batchModify)
        self.right_menu_bless.add_separator()

        self.right_menu_bless.add_command(label='保存文件',command=self.treeBlessSave)
        
        # 绑定
        self.tree_blesswords.bind('<Button-3>',self.onTreeBlessRightCLick)
        
        # 定义双击事件，为编辑状态,参考文章：https://blog.csdn.net/falwat/article/details/127494533
        # 创建`self.delegate_var`用于绑定代理部件
        self.delegate_var_bless = StringVar()
        # 将`self.tree_item_edit()`函数绑定到`self.tree`的`<Double-1>`事件上.
        self.tree_blesswords.bind('<Double-1>', func=self.tree_item_edit_bless)

    # 处理双击事件
    def tree_item_edit_bless(self, e: Event):
        # 获取选中的项目
        self.selected_item = self.tree_blesswords.selection()[0]

        for i, col in enumerate(self.tree_blesswords['columns']):
            # 获取选中项目中每列的x偏移, y偏移, 宽度 和 高度
            
            x, y, w, h =  self.tree_blesswords.bbox(self.selected_item, col)
            # 判断鼠标点击事件的位置, 是否在选中项目的某列中
            if x < e.x < x + w and y < e.y < y + h:
                self.selected_column = col
                text = self.tree_blesswords.item(self.selected_item, 'values')[i]
                break
            else:
                # 鼠标点击位置不在数据列中
                self.selected_column = None
                # 获取该项目的x偏移, y偏移, 宽度 和 高度
                x, y, w, h =  self.tree_blesswords.bbox(self.selected_item)
                text = self.tree_blesswords.item(self.selected_item, 'text')
        
        # 设置 `self.delegate_var` 的值
        self.delegate_var_bless.set(text)
        # 如果选中列不是index列：
        if self.selected_column !='index':
            self.delegate_widget = Entry(self.tree_blesswords, width=w // 10, textvariable=self.delegate_var_bless)
            
            self.delegate_widget.bind('<FocusOut>', func=self.tree_item_edit_done_bless) # 失去焦点
            self.delegate_widget.bind('<Return>', func=self.tree_item_edit_done_bless)# 或者敲入回车
            self.delegate_widget.selection_range(0, END)
            # 将 self.delegate_widget 放到 `self.tree` 的网格布局中, `padx` 和 `pady` 分别给选中单元格的 `x` 偏移 和 `y` 偏移
            self.delegate_widget.place(x=x,y=y) #  这样方位置才是对的。用grid不行
            self.delegate_widget.focus()
    
    # 编辑完成
    def tree_item_edit_done_bless(self, e):
        # 将`self.delegate_widget`从`self.tree`的网格布局中移除
        self.delegate_widget.place_forget()
        self.tree_blesswords.set(self.selected_item, self.selected_column, self.delegate_var_bless.get())
    
    # 祝福词的右键菜单
    def onTreeBlessRightCLick(self,event):
    # 表格上鼠标右键事件
        self.current_itemID = self.tree_blesswords.identify_row(event.y) # 事件所在行ID
        selected_itemids = self.tree_blesswords.selection() # ('I001','I002'),返回是一个数组
        if self.current_itemID not in selected_itemids: # 如果当前行不是被选中状态，则设定当前行为选中状态，
            self.tree_blesswords.selection_set(self.current_itemID)
        
        self.right_menu_bless.tk_popup(event.x_root, event.y_root, 0)
        
    # 祝福词右键菜单的添加
    def treeBlessAdd(self):
        current_index = self.tree_blesswords.index(self.current_itemID) # 根据itemID取得 index
        self.tree_blesswords.insert("",index=current_index,values=('1','请双击此处，添加新的祝福词！'))
        # 修改后面的全部 把序号都加1
        # 从头全部刷新序号
        self.refreshBlessIndex()
    
    # 祝福词右键菜单的删除
    def treeBlessDelete(self):
        for itemid in self.tree_blesswords.selection():
            self.tree_blesswords.delete(itemid)
        # 从头全部刷新序号
        self.refreshBlessIndex()
    
    # 清空选择
    def treeBlessCLearSelected(self):
        for item in self.tree_blesswords.selection():
            self.tree_blesswords.selection_remove(item) # clear 方法不对，要用remove来清除选择
    
    # 祝福词右键菜单的保存
    def treeBlessSave(self):
        with open(self.belsswords_csv, 'w',encoding='utf-8') as f:
            f.write('')
        with open(self.belsswords_csv, 'a',encoding='utf-8') as f:
            for item in self.tree_blesswords.get_children():
                item_values = self.tree_blesswords.item(item,'values')
                blessword = item_values[1]
                f.write(blessword + '\n')
    
    # 刷新祝福词序号
    def refreshBlessIndex(self):
        for item in self.tree_blesswords.get_children():
            index_item = self.tree_blesswords.index(item)
            self.tree_blesswords.set(item, column='index', value=index_item+1)
    ################################################################################## 
    # mainloop 函数
    def mainLoop(self):
        self.win.mainloop()

# 过滤掉字符串前后的空格
def rlstrip(str):
    return str.rstrip().lstrip()

# 主程序
def main():
    appGUI = WxGroupMsgGUI()
    
    
    
    # 本行必须放在最后
    appGUI.mainLoop()
    

if __name__ == '__main__':
    main()
     
