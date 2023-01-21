# -*- codeing = utf-8 -*-
# @Time: 2021/11/29 2:14
# @Author: Foxhuty
# @File: test_data.py
# @Software: PyCharm
# @Based on python 3.9

import os
import pandas as pd
from datetime import datetime
from scores_source.scores_process import GetInfoFromId

root_dir = r'D:\成绩统计结果'

# file = r'D:\work documents\高2020级\高2020级半期考试.xlsx'
# df = pd.read_excel(file, sheet_name="文科")


def del_files(file_dir):
    n = 0
    if len(os.listdir(file_dir)) >= 3:
        for f in os.listdir(file_dir):
            os.remove(file_dir + os.sep + f)
        print(f'files have been deleted')
        n += 1
    print(f'第{n}次清理完毕')


def split_into_class(df_data):
    class_numbers = df_data['班级'].unique()
    writer = pd.ExcelWriter(r'D:\成绩统计结果\按班级拆分表.xlsx')
    df_data.to_excel(writer, sheet_name='按班拆分表', index=False)
    whole_df = pd.DataFrame()
    for i in class_numbers:
        class_number = df_data[df_data['班级'] == i].reset_index(drop=True)
        class_number['序号'] = [k + 1 for k in class_number.index]
        whole_df = pd.concat([whole_df, class_number])
    whole_df.to_excel(writer, sheet_name='按班拆分表', index=False)
    # print(whole_df)
    final_file = os.path.basename(writer)
    print(final_file)
    writer.close()
    return final_file


def get_date_weekday():
    date = datetime.now().strftime('%Y年%m月%d日')
    week = datetime.weekday(datetime.now())
    week_name = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
    week_day = week_name[week]
    date_week = f'{date} {week_day}'
    return date_week


if __name__ == '__main__':
    file_location = r'D:\年级管理数据\高2021级\高二上\成都十一中学高2021级学生信息表.xls'
    weekday = get_date_weekday()
    get_info=GetInfoFromId(file_location)
    get_info.get_sex_birth_age()

