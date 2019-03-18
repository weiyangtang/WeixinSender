# WeixinSender
## 利用itchat实现微信批量每天定时发送消息
### 前期准备工作
cmd 命令依次输入
```python
pip install itchat
pip install xlrd
pip install xlwt
pip install xlutils
```
### 从GitHub下载项目
```
git clone https://github.com/weiyangtang/WeixinSender.git
```
或者直接[GitHub - weiyangtang/WeixinSender: 利用itchat实现微信批量每天定时发送消息](https://github.com/weiyangtang/WeixinSender)点击下载
![](https://ws1.sinaimg.cn/large/007EXT0fgy1g16tfi27bjj31hc0p40u6)
### 文件结构
![](https://ws1.sinaimg.cn/large/007EXT0fgy1g16tdcdw2hj308i04o0sl)
### 使用说明
1. 运行 getWeixinFriendList_excel.py 获取微信好友信息写入到Excel data\weixinFridendList.xl.xls，在Excel文件中第一列为好友备注名，
第二列填入你所要发送的消息
![](https://ws1.sinaimg.cn/large/007EXT0fgy1g16tdcihftj30l20grmxl)
2. 修改 sendWeixinMessage.py中的发送时间,改成你想要发送的时间
```python
sendTime = "2019-3-18 11:30:00"  # 发送时间 格式为年月日时分秒毫秒
```
![](https://ws1.sinaimg.cn/large/007EXT0fgy1g16tdc9uuej30p406z74g)
3. 运行 sendWeixinMessage.py，微信扫描二维码登录

