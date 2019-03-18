# encoding: utf-8
'''
@author: weiyang_tang
@contact: weiyang_tang@126.com
@file: getWeixinFriendList_excel.py
@time: 2019-02-04 22:05
@desc: 获取微信好友的信息，使用前在当前路径下新建一个名为data文件夹,在文件夹里创建一个excel 文件记住是以 .xls,不是.xlsx
'''

import itchat
import xlwt
import xlrd
from xlutils.copy import copy


def getFriends():
    itchat.auto_login(hotReload=True)
    # itchat.run()
    friendList = itchat.get_friends()
    i = 0
    for each in friendList:
        if len(each['RemarkName']) < 2:  # 如果 备注名为空,则用微信昵称
            print(i)
            print(each['NickName'])
            execlWrite(i, 0, each['NickName'])
        else:
            print(i)
            print(each['RemarkName'])
            execlWrite(i, 0, each['RemarkName'])
        i = i + 1


def sendMessage(name, message):
    itchat.auto_login(hotReload=True)
    sender = itchat.search_friends(name)[0]['UserName']
    itchat.send(message, toUserName=sender)
    itchat.send(message, toUserName='filehelper')


def execlWrite(row, col, content):
    old_excel = xlrd.open_workbook('data/weixinFridendList.xls', formatting_info=True)
    # 将操作文件对象拷贝，变成可写的workbook对象
    new_excel = copy(old_excel)
    # 获得第一个sheet的对象
    ws = new_excel.get_sheet(0)
    # 写入数据
    ws.write(row, col, content)
    new_excel.save('data/weixinFridendList.xls')
    # 另存为excel文件，并将文件命名
    # new_excel.save('new_fileName.xls')


def excelRead():
    # 文件位置

    ExcelFile = xlrd.open_workbook('data/weixinFridendList.xls')  # 微信好友名单记录到data/weixinFridendList.xls

    print(ExcelFile.sheet_names())

    sheet = ExcelFile.sheet_by_index(0)

    # sheet = ExcelFile.sheet_by_name('Sheet1')

    # 打印sheet的名称，行数，列数

    print(sheet.name, sheet.nrows, sheet.ncols)
    for i in range(0, sheet.nrows - 1):
        if len(sheet.cell(i, 1).value) > 0:  # 发送消息的内容是否为空
            print(sheet.cell(i, 0).value + "\t" + sheet.cell(i, 1).value + " 已经发送了")


if __name__ == '__main__':
    # friendList.excelRead()
    getFriends()
    # excelRead()
