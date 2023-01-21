# -*- codeing = utf-8 -*-
# @Time: 2021/6/17 1:11
# @Author: Foxhuty
# @File: test_01.py
# @Software: PyCharm
# @Based on python 3.9
import requests

headers = {
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) '
                  'AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36'}
r = requests.get('https://www.baidu.com/s',
                 params={
                     'wd': '毛泽东'
                 },
                 headers=headers)

print(r.status_code)
r.encoding = 'utf-8'
print(r.encoding)
print('*' * 30)
print(r.text)

print(r.headers['Content-Type'])

r=requests.get('http://httpbin.org/user-agent')

print(r.headers['Content-Type'])
d=r.json()
print(d)
t=(1,2,3)
l=list(t)
print(l)


