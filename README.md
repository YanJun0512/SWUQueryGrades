# SWUQueryGrades
python+selenium从西南大学教务处获取个人成绩。
***
环境
```
win10
python3.6
pycharm 2017.3.2
Edge浏览器
```

依赖库
```
import requests
from selenium import webdriver
from multiprocessing import Process
import time
import win32con
import win32api
```

学校教务系统只有校园内网IP可以访问，因此这个实验也只有连入校园网的电脑可以成功运行。


登录校内门户（教务系统）post的数据全部加密处理，且每次密文都不一样，所以实验选择```selenium```配合```webdrive```，自动输入账号密码，从跳转网页中获取```cookie```值，再用这个```cookie```值访问。


登录页面属于**校内门户**，而**教务系统**————也就是成绩查询的入口又和校内门户有不同网址，抓包发现他们的```cookie```值也不一样，就需要先从校内门户登录，再利用webdrive点击**教务系统**按钮进入教务系统，再获取```cookie```值。


这里采用Edge浏览器驱动程序作为webdrive，原因是chrome浏览器无法通过学校教务系统的证书，原因不明。
