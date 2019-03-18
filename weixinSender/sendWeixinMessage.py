# encoding: utf-8
'''
@author: weiyang_tang
@contact: weiyang_tang@126.com
@file: sendWeixinMessage.py
@time: 2019-02-05 16:38
@desc: 获取好友备注或者昵称后，写入excel，在excel后一列的消息，写入所要发送的消息
'''

import itchat
import xlrd
import datetime
import time
from threading import Timer


def TimerSender():
    send()
    t = Timer(24 * 3600, TimerSender)  # 每隔24*3600（一天）秒再次发送消息
    t.start()


def send():
    # itchat.auto_login(hotReload=True)
    ExcelFile = xlrd.open_workbook('data/weixinFridendList.xls')
    print(ExcelFile.sheet_names())
    sheet = ExcelFile.sheet_by_index(0)
    print(sheet.name, sheet.nrows, sheet.ncols)
    for i in range(0, sheet.nrows - 1):
        if len(sheet.cell(i, 1).value) > 0:
            sender = itchat.search_friends(sheet.cell(i, 0).value)[0]['UserName']
            message = sheet.cell(i, 1).value
            itchat.send(message, toUserName=sender)
            itchat.send(message, toUserName='filehelper')
            print(sheet.cell(i, 0).value + "\t" + sheet.cell(i, 1).value + " 已经发送了")


if __name__ == '__main__':
    itchat.auto_login(hotReload=True)
    sendTime = "2019-3-18 11:30:00"  # 发送时间 格式为年月日时分秒毫秒

    targetTime = datetime.datetime.strptime(sendTime, "%Y-%m-%d %H:%M:%S")
    now = time.time()
    delay_time = targetTime.timestamp() - now  # 发送时间与现在时间的差值

    t = Timer(delay_time, TimerSender)
    t.start()
