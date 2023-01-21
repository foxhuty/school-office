# -*- codeing = utf-8 -*-
# @Time: 2021/6/19 1:14
# @Author: Foxhuty
# @File: test_02.py
# @Software: PyCharm
# @Based on python 3.9


import re
re_pattern = r'^([1-9]\d{5}[12]\d{3}(0[1-9]|1[012])(0[1-9]|[12][0-9]|3[01])\d{3}[0-9xX])$'
res=re.match(re_pattern,'510181200612040039')
print(res)
