# -*- codeing = utf-8 -*-
# @Time: 2021/5/27 23:32
# @Author: Foxhuty
# @File: scores_process.py
# @Software: PyCharm
# @Based on python 3.9
import pandas as pd
import numpy as np
import os
import random
import time
import warnings
import re
import datetime

warnings.simplefilter('ignore', FutureWarning)


class ScoreAnalysis(object):
    """
    用高中成绩分析，计算平均分，有效分，有效分人数，错位生人数，成绩对比，学科评定，考室安排，学生个人成绩单等。
    简单高效地完成考试成绩分析。
    """
    arts_scores = []
    science_scores = []

    def __init__(self, path):
        self.path = path
        self.df_arts = pd.read_excel(self.path, sheet_name='文科', index_col='序号',
                                     dtype={
                                         '班级': str,
                                         '序号': str,
                                         '名次': str,
                                         '考生号': str,
                                         '考号': str

                                     }
                                     )
        self.df_science = pd.read_excel(self.path, sheet_name='理科', index_col='序号',
                                        dtype={
                                            '班级': 'str',
                                            '序号': 'str',
                                            '名次': 'str',
                                            '考生号': 'str',
                                            '考号': 'str'

                                        }
                                        )

    def __str__(self):
        return f'对{os.path.basename(self.path)}进行成绩数据处理，生成excel表.'

    def get_av(self):
        """
        计算各科平均分
        :return: 文科，理科各班各科平均分
        """
        av_class_arts = self.df_arts.groupby(['班级'])[
            ['语文', '数学', '英语', '政治', '历史', '地理', '总分']].mean().round(2)
        av_general_arts = self.df_arts[['语文', '数学', '英语', '政治', '历史', '地理', '总分']].apply(np.nanmean,
                                                                                                       axis=0).round(2)
        av_general_arts.name = '年级'
        av_arts = av_class_arts.append(av_general_arts)
        student_number = self.get_student_number_class(self.df_arts)
        av_arts = av_arts.join(student_number)
        order = ['参考人数', '语文', '数学', '英语', '政治', '历史', '地理', '总分']
        av_arts = av_arts[order]

        av_class_science = self.df_science.groupby(['班级'])[
            ['语文', '数学', '英语', '物理', '化学', '生物', '总分']].mean().round(2)
        av_general_science = self.df_science[['语文', '数学', '英语', '物理', '化学', '生物', '总分']].apply(np.nanmean,
                                                                                                             axis=0).round(
            2)
        av_general_science.name = '年级'
        av_science = av_class_science.append(av_general_science)
        student_number_science = self.get_student_number_class(self.df_science)
        av_science = av_science.join(student_number_science)
        order_science = ['参考人数', '语文', '数学', '英语', '物理', '化学', '生物', '总分']
        av_science = av_science[order_science]
        return av_arts, av_science

    @staticmethod
    def get_student_number_class(df_data):
        student_number_class = df_data['班级'].value_counts()
        student_number_class.name = '参考人数'
        student_number_class['年级'] = student_number_class.sum()
        return student_number_class

    def get_goodscores_arts(self, goodtotal_arts):
        """
        计算文科各科有效分
        goodtotal:划线总分，高线，中线，低线
        """
        chn = self.get_subject_good_score(self.df_arts, '语文', goodtotal_arts)
        math = self.get_subject_good_score(self.df_arts, '数学', goodtotal_arts)
        eng = self.get_subject_good_score(self.df_arts, '英语', goodtotal_arts)
        pol = self.get_subject_good_score(self.df_arts, '政治', goodtotal_arts)
        his = self.get_subject_good_score(self.df_arts, '历史', goodtotal_arts)
        geo = self.get_subject_good_score(self.df_arts, '地理', goodtotal_arts)
        if (chn + math + eng + pol + his + geo) > goodtotal_arts:
            math = math - 1
        if (chn + math + eng + pol + his + geo) < goodtotal_arts:
            eng = eng + 1

        return chn, math, eng, pol, his, geo, goodtotal_arts

    def get_goodscores_science(self, goodtotal_science):
        """
        计算理科各科有效分
        goodtotal:划线总分，高线，中线，低线
        """
        chn = self.get_subject_good_score(self.df_science, '语文', goodtotal_science)
        math = self.get_subject_good_score(self.df_science, '数学', goodtotal_science)
        eng = self.get_subject_good_score(self.df_science, '英语', goodtotal_science)
        phys = self.get_subject_good_score(self.df_science, '物理', goodtotal_science)
        chem = self.get_subject_good_score(self.df_science, '化学', goodtotal_science)
        bio = self.get_subject_good_score(self.df_science, '生物', goodtotal_science)

        if (chn + math + eng + phys + chem + bio) > goodtotal_science:
            math = math - 1
        if (chn + math + eng + phys + chem + bio) < goodtotal_science:
            eng = eng + 1

        return chn, math, eng, phys, chem, bio, goodtotal_science

    def single_double_arts(self, chn, math, eng, pol, his, geo, total):
        """
        计算各科各班单有效和双有效人数
        :param chn:
        :param math:
        :param eng:
        :param pol:
        :param his:
        :param geo:
        :param total:
        :return: result_single,result_double
        """
        single_chn_arts, double_chn_arts = self.get_single_double_score(self.df_arts, '语文', chn, total)
        single_math_arts, double_math_arts = self.get_single_double_score(self.df_arts, '数学', math, total)
        single_eng_arts, double_eng_arts = self.get_single_double_score(self.df_arts, '英语', eng, total)
        single_pol_arts, double_pol_arts = self.get_single_double_score(self.df_arts, '政治', pol, total)
        single_his_arts, double_his_arts = self.get_single_double_score(self.df_arts, '历史', his, total)
        single_geo_arts, double_geo_arts = self.get_single_double_score(self.df_arts, '地理', geo, total)
        single_total_arts, double_total_arts = self.get_single_double_score(self.df_arts, '总分', total, total)
        name_num = self.df_arts.groupby(['班级'])['姓名'].count()
        name_num.name = '参考人数'
        arts_single = pd.concat([name_num, single_chn_arts, single_math_arts, single_eng_arts,
                                 single_pol_arts, single_his_arts, single_geo_arts, single_total_arts],
                                axis=1)
        arts_double = pd.concat(
            [name_num, double_chn_arts, double_math_arts, double_eng_arts,
             double_pol_arts, double_his_arts, double_geo_arts, double_total_arts], axis=1)

        arts_single.loc['文科共计'] = [arts_single['参考人数'].sum(),
                                       arts_single['语文'].sum(),
                                       arts_single['数学'].sum(),
                                       arts_single['英语'].sum(),
                                       arts_single['政治'].sum(),
                                       arts_single['历史'].sum(),
                                       arts_single['地理'].sum(),
                                       arts_single['总分'].sum()
                                       ]
        # 新增一列上线率
        arts_single['上线率'] = arts_single['总分'] / arts_single['参考人数']
        arts_single['上线率'] = arts_single['上线率'].apply(lambda x: format(x, '.2%'))
        arts_double.loc['文科共计'] = [arts_double['参考人数'].sum(),
                                       arts_double['语文'].sum(),
                                       arts_double['数学'].sum(),
                                       arts_double['英语'].sum(),
                                       arts_double['政治'].sum(),
                                       arts_double['历史'].sum(),
                                       arts_double['地理'].sum(),
                                       arts_double['总分'].sum()]
        good_scores_dict = {'参考人数': ' ', '语文': chn, '数学': math, '英语': eng,
                            '政治': pol, '历史': his, '地理': geo, '总分': total}
        good_scores = pd.DataFrame(good_scores_dict, index=['班级'])
        return arts_single, arts_double, good_scores

    def single_double_science(self, chn, math, eng, phys, chem, bio, total):
        """
               计算理科各科各班上单有效和双有效分人数
               """

        single_chn_science, double_chn_science = self.get_single_double_score(self.df_science, '语文', chn, total)
        single_math_science, double_math_science = self.get_single_double_score(self.df_science, '数学', math, total)
        single_eng_science, double_eng_science = self.get_single_double_score(self.df_science, '英语', eng, total)
        single_phys_science, double_phys_science = self.get_single_double_score(self.df_science, '物理', phys, total)
        single_chem_science, double_chem_science = self.get_single_double_score(self.df_science, '化学', chem, total)
        single_bio_science, double_bio_science = self.get_single_double_score(self.df_science, '生物', bio, total)
        single_total_science, double_total_science = self.get_single_double_score(self.df_science, '总分', total, total)

        name_num = self.df_science.groupby(['班级'])['姓名'].count()
        name_num.name = '参考人数'

        science_single = pd.concat([name_num, single_chn_science, single_math_science, single_eng_science,
                                    single_phys_science, single_chem_science, single_bio_science, single_total_science],
                                   axis=1)
        science_double = pd.concat(
            [name_num, double_chn_science, double_math_science, double_eng_science, double_phys_science,
             double_chem_science, double_bio_science, double_total_science], axis=1)

        science_single.loc['理科共计'] = [science_single['参考人数'].sum(),
                                          science_single['语文'].sum(),
                                          science_single['数学'].sum(),
                                          science_single['英语'].sum(),
                                          science_single['物理'].sum(),
                                          science_single['化学'].sum(),
                                          science_single['生物'].sum(),
                                          science_single['总分'].sum()
                                          ]
        # 新增一列上线率
        science_single['上线率'] = science_single['总分'] / science_single['参考人数']
        science_single['上线率'] = science_single['上线率'].apply(lambda x: format(x, '.2%'))
        science_double.loc['理科共计'] = [science_double['参考人数'].sum(),
                                          science_double['语文'].sum(),
                                          science_double['数学'].sum(),
                                          science_double['英语'].sum(),
                                          science_double['物理'].sum(),
                                          science_double['化学'].sum(),
                                          science_double['生物'].sum(),
                                          science_double['总分'].sum()]
        good_scores_dict = {'参考人数': ' ', '语文': chn, '数学': math, '英语': eng,
                            '物理': phys, '化学': chem, '生物': bio, '总分': total}
        good_scores_science = pd.DataFrame(good_scores_dict, index=['班级'])

        return science_single, science_double, good_scores_science

    def contribution_shoot_arts(self):
        result_single, result_double, good_scores = self.single_double_arts(*ScoreAnalysis.arts_scores)
        shoot_dict = {'语文': result_double['语文'] / result_single['语文'],
                      '数学': result_double['数学'] / result_single['数学'],
                      '英语': result_double['英语'] / result_single['英语'],
                      '政治': result_double['政治'] / result_single['政治'],
                      '历史': result_double['历史'] / result_single['历史'],
                      '地理': result_double['地理'] / result_single['地理']}
        shoot_df = pd.DataFrame(shoot_dict)
        contribution_dict = {'语文': result_double['语文'] / result_double['总分'],
                             '数学': result_double['数学'] / result_double['总分'],
                             '英语': result_double['英语'] / result_double['总分'],
                             '政治': result_double['政治'] / result_double['总分'],
                             '历史': result_double['历史'] / result_double['总分'],
                             '地理': result_double['地理'] / result_double['总分']}
        contribution_df = pd.DataFrame(contribution_dict)
        return shoot_df, contribution_df

    def contribution_shoot_science(self):
        result_single, result_double, good_score_science = self.single_double_science(*ScoreAnalysis.science_scores)
        shoot_dict = {'语文': result_double['语文'] / result_single['语文'],
                      '数学': result_double['数学'] / result_single['数学'],
                      '英语': result_double['英语'] / result_single['英语'],
                      '物理': result_double['物理'] / result_single['物理'],
                      '化学': result_double['化学'] / result_single['化学'],
                      '生物': result_double['生物'] / result_single['生物'],
                      }
        shoot_df = pd.DataFrame(shoot_dict)
        contribution_dict = {'语文': result_double['语文'] / result_double['总分'],
                             '数学': result_double['数学'] / result_double['总分'],
                             '英语': result_double['英语'] / result_double['总分'],
                             '物理': result_double['物理'] / result_double['总分'],
                             '化学': result_double['化学'] / result_double['总分'],
                             '生物': result_double['生物'] / result_double['总分'],
                             }
        contribution_df = pd.DataFrame(contribution_dict)
        return shoot_df, contribution_df

    def subject_grade_arts(self):

        result_single, result_double, good_scores = self.single_double_arts(*ScoreAnalysis.arts_scores)
        shoot_df, contribution_df = self.contribution_shoot_arts()
        result_single.fillna(0, inplace=True)
        result_double.fillna(0, inplace=True)
        shoot_df.fillna(0, inplace=True)
        contribution_df.fillna(0, inplace=True)
        grade = pd.DataFrame(columns=shoot_df.columns, index=shoot_df.index)

        def grade_assess(subject):
            for i in shoot_df.index:
                if result_single['总分'].at[i] != 0:
                    if (result_single[subject].at[i]) >= (result_single['总分'].at[i]) * 0.8:
                        if (contribution_df[subject].at[i] >= 0.7) & (shoot_df[subject].at[i] >= 0.6):
                            grade[subject].at[i] = 'A'
                        elif (contribution_df[subject].at[i] >= 0.7) & (shoot_df[subject].at[i] < 0.6):
                            grade[subject].at[i] = 'B'
                        elif (contribution_df[subject].at[i] < 0.7) & (shoot_df[subject].at[i] >= 0.6):
                            grade[subject].at[i] = 'C'
                        else:
                            grade[subject].at[i] = 'D'
                    else:
                        grade[subject].at[i] = 'E'
                else:
                    grade[subject].at[i] = 'F'

        grade_assess('语文')
        grade_assess('数学')
        grade_assess('英语')
        grade_assess('政治')
        grade_assess('历史')
        grade_assess('地理')

        contribution_df = contribution_df.applymap(lambda x: format(x, '.2%'))
        shoot_df = shoot_df.applymap(lambda x: format(x, '.2%'))
        final_grade = pd.concat([good_scores, result_single, result_double, contribution_df, shoot_df, grade],
                                keys=['有效分', '单有效', '双有效', '贡献率', '命中率', '等级'])
        return final_grade

    def subject_grade_science(self):
        result_single, result_double, good_score_science = self.single_double_science(*ScoreAnalysis.science_scores)
        shoot_df, contribution_df = self.contribution_shoot_science()
        result_single.fillna(0, inplace=True)
        result_double.fillna(0, inplace=True)
        shoot_df.fillna(0, inplace=True)
        contribution_df.fillna(0, inplace=True)
        grade = pd.DataFrame(columns=shoot_df.columns, index=shoot_df.index)

        def grade_assess(subject):
            for i in shoot_df.index:
                if result_single['总分'].at[i] != 0:
                    if (result_single[subject].at[i]) >= (result_single['总分'].at[i]) * 0.8:
                        if (contribution_df[subject].at[i] >= 0.7) & (shoot_df[subject].at[i] >= 0.6):
                            grade[subject].at[i] = 'A'
                        elif (contribution_df[subject].at[i] >= 0.7) & (shoot_df[subject].at[i] < 0.6):
                            grade[subject].at[i] = 'B'
                        elif (contribution_df[subject].at[i] < 0.7) & (shoot_df[subject].at[i] >= 0.6):
                            grade[subject].at[i] = 'C'
                        else:
                            grade[subject].at[i] = 'D'
                    else:
                        grade[subject].at[i] = 'E'
                else:
                    grade[subject].at[i] = 'F'

        grade_assess('语文')
        grade_assess('数学')
        grade_assess('英语')
        grade_assess('物理')
        grade_assess('化学')
        grade_assess('生物')

        contribution_df = contribution_df.applymap(lambda x: format(x, '.2%'))
        shoot_df = shoot_df.applymap(lambda x: format(x, '.2%'))

        final_grade = pd.concat([good_score_science, result_single, result_double, contribution_df, shoot_df, grade],
                                keys=['有效分', '单有效', '双有效', '贡献率', '命中率', '等级'])
        return final_grade

    def goodscore_arts(self, chn, math, eng, pol, his, geo, total):
        """
        计算文科各科各班单有效和双有效人数
        """
        single_chn_arts, double_chn_arts = self.get_single_double_score(self.df_arts, '语文', chn, total)
        single_math_arts, double_math_arts = self.get_single_double_score(self.df_arts, '数学', math, total)
        single_eng_arts, double_eng_arts = self.get_single_double_score(self.df_arts, '英语', eng, total)
        single_pol_arts, double_pol_arts = self.get_single_double_score(self.df_arts, '政治', pol, total)
        single_his_arts, double_his_arts = self.get_single_double_score(self.df_arts, '历史', his, total)
        single_geo_arts, double_geo_arts = self.get_single_double_score(self.df_arts, '地理', geo, total)
        single_total_arts, double_total_arts = self.get_single_double_score(self.df_arts, '总分', total, total)

        name_num = self.df_arts.groupby(['班级'])['姓名'].count()
        name_num.name = '参考人数'

        goodscore_dict = {'参考人数': ' ', '语文': chn, '数学': math, '英语': eng, '政治': pol, '历史': his,
                          '地理': geo, '总分': total}
        goodscore_df = pd.DataFrame(goodscore_dict, index=['班级'])

        result_single = pd.concat([name_num, single_chn_arts, single_math_arts, single_eng_arts,
                                   single_pol_arts, single_his_arts, single_geo_arts, single_total_arts],
                                  axis=1)
        result_double = pd.concat(
            [name_num, double_chn_arts, double_math_arts, double_eng_arts,
             double_pol_arts, double_his_arts, double_geo_arts, double_total_arts], axis=1)

        result_single.loc['文科共计'] = [result_single['参考人数'].sum(),
                                         result_single['语文'].sum(),
                                         result_single['数学'].sum(),
                                         result_single['英语'].sum(),
                                         result_single['政治'].sum(),
                                         result_single['历史'].sum(),
                                         result_single['地理'].sum(),
                                         result_single['总分'].sum()
                                         ]
        # 增加一列上线率
        # result_single['上线率'] = result_single['总分'] / result_single['参考人数']
        # result_single['上线率'] = result_single['上线率'].apply(lambda x: format(x, '.2%'))
        result_single = self.good_scores_ratio(result_single, '总分')
        result_double.loc['文科共计'] = [result_double['参考人数'].sum(),
                                         result_double['语文'].sum(),
                                         result_double['数学'].sum(),
                                         result_double['英语'].sum(),
                                         result_double['政治'].sum(),
                                         result_double['历史'].sum(),
                                         result_double['地理'].sum(),
                                         result_double['总分'].sum()]

        # 计算错位生人数
        unmatched_dict = {'参考人数': name_num, '语文': single_total_arts - double_chn_arts,
                          '数学': single_total_arts - double_math_arts, '英语': single_total_arts - double_eng_arts,
                          '政治': single_total_arts - double_pol_arts, '历史': single_total_arts - double_his_arts,
                          '地理': single_total_arts - double_geo_arts, '总分': single_total_arts - double_total_arts}
        unmatched_df = pd.DataFrame(unmatched_dict)
        # 新增一行共计
        unmatched_df.loc['共计'] = [unmatched_df['参考人数'].sum(),
                                    unmatched_df['语文'].sum(),
                                    unmatched_df['数学'].sum(),
                                    unmatched_df['英语'].sum(),
                                    unmatched_df['政治'].sum(),
                                    unmatched_df['历史'].sum(),
                                    unmatched_df['地理'].sum(),
                                    unmatched_df['总分'].sum()
                                    ]
        result_final_arts = pd.concat([goodscore_df, result_single, result_double, unmatched_df], axis=0,
                                      keys=['有效分数', '单有效', '双有效', '错位数'])
        # 计算学科贡献率，命中率和等级评定
        shoot_df, contribution_df = self.contribution_shoot_arts()
        result_single.fillna(0, inplace=True)
        result_double.fillna(0, inplace=True)
        shoot_df.fillna(0, inplace=True)
        contribution_df.fillna(0, inplace=True)
        grade = pd.DataFrame(columns=['语文', '数学', '英语', '政治', '历史', '地理'], index=shoot_df.index)

        def grade_assess(subject):
            for i in shoot_df.index:
                if result_single['总分'].at[i] != 0:
                    if (result_single[subject].at[i]) >= (result_single['总分'].at[i]) * 0.8:
                        if (contribution_df[subject].at[i] >= 0.7) & (shoot_df[subject].at[i] >= 0.6):
                            grade[subject].at[i] = 'A'
                        elif (contribution_df[subject].at[i] >= 0.7) & (shoot_df[subject].at[i] < 0.6):
                            grade[subject].at[i] = 'B'
                        elif (contribution_df[subject].at[i] < 0.7) & (shoot_df[subject].at[i] >= 0.6):
                            grade[subject].at[i] = 'C'
                        else:
                            grade[subject].at[i] = 'D'
                    else:
                        grade[subject].at[i] = 'E'
                else:
                    grade[subject].at[i] = 'F'

        grade_assess('语文')
        grade_assess('数学')
        grade_assess('英语')
        grade_assess('政治')
        grade_assess('历史')
        grade_assess('地理')
        contribution_df = contribution_df.applymap(lambda x: format(x, '.2%'))
        shoot_df = shoot_df.applymap(lambda x: format(x, '.2%'))
        final_grade = pd.concat(
            [goodscore_df, result_single, result_double, unmatched_df, contribution_df, shoot_df, grade],
            keys=['有效分数', '单有效', '双有效', '错位数', '贡献率', '命中率', '等级'])

        return result_final_arts, final_grade

    def goodscore_science(self, chn, math, eng, phys, chem, bio, total):
        """
        计算理科各科各班上单有效和双有效分人数
        """
        single_chn_science, double_chn_science = self.get_single_double_score(self.df_science, '语文', chn, total)
        single_math_science, double_math_science = self.get_single_double_score(self.df_science, '数学', math, total)
        single_eng_science, double_eng_science = self.get_single_double_score(self.df_science, '英语', eng, total)
        single_phys_science, double_phys_science = self.get_single_double_score(self.df_science, '物理', phys, total)
        single_chem_science, double_chem_science = self.get_single_double_score(self.df_science, '化学', chem, total)
        single_bio_science, double_bio_science = self.get_single_double_score(self.df_science, '生物', bio, total)
        single_total_science, double_total_science = self.get_single_double_score(self.df_science, '总分', total, total)

        name_num = self.df_science.groupby(['班级'])['姓名'].count()
        name_num.name = '参考人数'

        goodscore_dict = {'参考人数': ' ', '语文': chn, '数学': math, '英语': eng, '物理': phys,
                          '化学': chem, '生物': bio, '总分': total}
        goodscore_df = pd.DataFrame(goodscore_dict, index=['有效分数'])

        result_single = pd.concat([name_num, single_chn_science, single_math_science, single_eng_science,
                                   single_phys_science, single_chem_science, single_bio_science, single_total_science],
                                  axis=1)
        result_double = pd.concat(
            [name_num, double_chn_science, double_math_science, double_eng_science, double_phys_science,
             double_chem_science, double_bio_science, double_total_science], axis=1)

        result_single.loc['理科共计'] = [result_single['参考人数'].sum(),
                                         result_single['语文'].sum(),
                                         result_single['数学'].sum(),
                                         result_single['英语'].sum(),
                                         result_single['物理'].sum(),
                                         result_single['化学'].sum(),
                                         result_single['生物'].sum(),
                                         result_single['总分'].sum()
                                         ]
        # 新增一列上线率
        result_single = self.good_scores_ratio(result_single, '总分')

        result_double.loc['理科共计'] = [result_double['参考人数'].sum(),
                                         result_double['语文'].sum(),
                                         result_double['数学'].sum(),
                                         result_double['英语'].sum(),
                                         result_double['物理'].sum(),
                                         result_double['化学'].sum(),
                                         result_double['生物'].sum(),
                                         result_double['总分'].sum()]

        # 计算错位生人数
        unmatched_dict = {'参考人数': name_num, '语文': single_total_science - double_chn_science,
                          '数学': single_total_science - double_math_science,
                          '英语': single_total_science - double_eng_science,
                          '物理': single_total_science - double_phys_science,
                          '化学': single_total_science - double_chem_science,
                          '生物': single_total_science - double_bio_science, '总分': single_total_science}
        unmatched_df = pd.DataFrame(unmatched_dict)
        unmatched_df.loc['共计'] = [unmatched_df['参考人数'].sum(),
                                    unmatched_df['语文'].sum(),
                                    unmatched_df['数学'].sum(),
                                    unmatched_df['英语'].sum(),
                                    unmatched_df['物理'].sum(),
                                    unmatched_df['化学'].sum(),
                                    unmatched_df['生物'].sum(),
                                    unmatched_df['总分'].sum()]
        result_final_science = pd.concat([goodscore_df, result_single, result_double, unmatched_df], axis=0,
                                         keys=['有效分数', '单有效', '双有效', '错位数'])
        result_final_science.fillna(0, inplace=True)

        # 计算学科贡献率，命中率和等级评定
        shoot_dict = {'语文': result_double['语文'] / result_single['语文'],
                      '数学': result_double['数学'] / result_single['数学'],
                      '英语': result_double['英语'] / result_single['英语'],
                      '物理': result_double['物理'] / result_single['物理'],
                      '化学': result_double['化学'] / result_single['化学'],
                      '生物': result_double['生物'] / result_single['生物']}
        shoot_df = pd.DataFrame(shoot_dict)
        contribution_dict = {'语文': result_double['语文'] / result_double['总分'],
                             '数学': result_double['数学'] / result_double['总分'],
                             '英语': result_double['英语'] / result_double['总分'],
                             '物理': result_double['物理'] / result_double['总分'],
                             '化学': result_double['化学'] / result_double['总分'],
                             '生物': result_double['生物'] / result_double['总分']}
        contribution_df = pd.DataFrame(contribution_dict)
        result_single.fillna(0, inplace=True)
        result_double.fillna(0, inplace=True)
        shoot_df.fillna(0, inplace=True)
        contribution_df.fillna(0, inplace=True)
        grade = pd.DataFrame(columns=['语文', '数学', '英语', '物理', '化学', '生物'], index=shoot_df.index)

        def grade_assess(subject):
            for i in shoot_df.index:
                if result_single['总分'].at[i] != 0:
                    if (result_single[subject].at[i]) >= (result_single['总分'].at[i]) * 0.8:
                        if (contribution_df[subject].at[i] >= 0.7) & (shoot_df[subject].at[i] >= 0.6):
                            grade[subject].at[i] = 'A'
                        elif (contribution_df[subject].at[i] >= 0.7) & (shoot_df[subject].at[i] < 0.6):
                            grade[subject].at[i] = 'B'
                        elif (contribution_df[subject].at[i] < 0.7) & (shoot_df[subject].at[i] >= 0.6):
                            grade[subject].at[i] = 'C'
                        else:
                            grade[subject].at[i] = 'D'
                    else:
                        grade[subject].at[i] = 'E'
                else:
                    grade[subject].at[i] = 'F'

        grade_assess('语文')
        grade_assess('数学')
        grade_assess('英语')
        grade_assess('物理')
        grade_assess('化学')
        grade_assess('生物')
        contribution_df = contribution_df.applymap(lambda x: format(x, '.2%'))
        shoot_df = shoot_df.applymap(lambda x: format(x, '.2%'))
        final_grade = pd.concat(
            [goodscore_df, result_single, result_double, unmatched_df, contribution_df, shoot_df, grade],
            keys=['有效分数', '单有效', '双有效', '错位数', '贡献率', '命中率', '等级'])
        return result_final_science, final_grade

    def get_unmatched_arts(self, chn, math, eng, pol, his, geo, total):

        df_chn = self.get_unmatched_students(self.df_arts, '语文', chn, total)
        df_math = self.get_unmatched_students(self.df_arts, '数学', math, total)
        df_eng = self.get_unmatched_students(self.df_arts, '英语', eng, total)
        df_pol = self.get_unmatched_students(self.df_arts, '政治', pol, total)
        df_his = self.get_unmatched_students(self.df_arts, '历史', his, total)
        df_geo = self.get_unmatched_students(self.df_arts, '地理', geo, total)
        df_unmatched_list = pd.concat([df_chn, df_math, df_eng, df_pol, df_his, df_geo], axis=1)

        df_cnh_num = df_chn.groupby('班级')[['语文']].count()
        df_math_num = df_math.groupby('班级')[['数学']].count()
        df_eng_num = df_eng.groupby('班级')[['英语']].count()
        df_pol_num = df_pol.groupby('班级')[['政治']].count()
        df_his_num = df_his.groupby('班级')[['历史']].count()
        df_geo_num = df_geo.groupby('班级')[['地理']].count()
        df_unmatched_num = pd.concat([df_cnh_num, df_math_num, df_eng_num, df_pol_num, df_his_num, df_geo_num], axis=1)
        return df_unmatched_num, df_unmatched_list

    def get_unmatched_science(self, chn, math, eng, phys, chem, bio, total):
        # 获取各科错位生名单
        df_chn = self.get_unmatched_students(self.df_science, '语文', chn, total)
        df_math = self.get_unmatched_students(self.df_science, '数学', math, total)
        df_eng = self.get_unmatched_students(self.df_science, '英语', eng, total)
        df_phys = self.get_unmatched_students(self.df_science, '物理', phys, total)
        df_chem = self.get_unmatched_students(self.df_science, '化学', chem, total)
        df_bio = self.get_unmatched_students(self.df_science, '生物', bio, total)
        df_unmatched_list = pd.concat([df_chn, df_math, df_eng, df_phys, df_chem, df_bio], axis=1)
        # 计算各科错位人数
        df_cnh_num = df_chn.groupby('班级')[['语文']].count()
        df_math_num = df_math.groupby('班级')[['数学']].count()
        df_eng_num = df_eng.groupby('班级')[['英语']].count()
        df_phys_num = df_phys.groupby('班级')[['物理']].count()
        df_chem_num = df_chem.groupby('班级')[['化学']].count()
        df_bio_num = df_bio.groupby('班级')[['生物']].count()
        df_unmatched_num = pd.concat([df_cnh_num, df_math_num, df_eng_num, df_phys_num, df_chem_num, df_bio_num],
                                     axis=1)
        return df_unmatched_num, df_unmatched_list

    def line_betweens(self, total=None, total_science=None):
        self.class_rank()
        line_condition = (self.df_arts['总分'] >= total - 20) & (self.df_arts['总分'] <= total + 20)
        line_condition_science = (self.df_science['总分'] >= total_science - 20) & (
                self.df_science['总分'] <= total_science + 20)
        df_line_arts = self.df_arts.loc[line_condition, :]
        df_line_science = self.df_science.loc[line_condition_science, :]
        writer = pd.ExcelWriter(r'D:\成绩统计结果\本次考试踩线生分班名单.xlsx')
        class_num = list(df_line_arts['班级'].drop_duplicates())
        class_num_science = list(df_line_science['班级'].drop_duplicates())
        for i in class_num:
            class_name = df_line_arts[df_line_arts['班级'] == i].reset_index(drop=True)
            class_name['序号'] = [k + 1 for k in class_name.index]
            class_name = class_name.loc[:, ['序号', '姓名', '班级', '语文', '数学', '英语',
                                            '政治', '历史', '地理', '总分', '排名']]
            class_name.to_excel(writer, sheet_name=i, index=False)

        for i in class_num_science:
            class_name = df_line_science[df_line_science['班级'] == i].reset_index(drop=True)
            class_name['序号'] = [k + 1 for k in class_name.index]
            class_name = class_name.loc[:, ['序号', '姓名', '班级', '语文', '数学', '英语',
                                            '物理', '化学', '生物', '总分', '排名']]
            class_name.to_excel(writer, sheet_name=i, index=False)

        writer.close()

    def class_divided(self):
        """
        计算获得文理科各班成绩表
        :return:
        """
        self.class_rank()
        class_No_arts = list(self.df_arts['班级'].drop_duplicates())
        class_NO_science = list(self.df_science['班级'].drop_duplicates())
        writer = pd.ExcelWriter(r'D:\成绩统计结果\本次考试文理各班成绩表.xlsx')
        for i in class_No_arts:
            class_arts = self.df_arts[self.df_arts['班级'] == i].reset_index(drop=True)
            class_arts['序号'] = [k + 1 for k in class_arts.index]
            class_arts['综合'] = class_arts['政治'] + class_arts['历史'] + class_arts['地理']
            class_arts = class_arts.loc[:, ['序号', '班级', '姓名', '语文', '数学', '英语', '综合',
                                            '政治', '历史', '地理', '总分', '排名']]
            class_arts.to_excel(writer, sheet_name=i, index=False)
        for i in class_NO_science:
            class_science = self.df_science[self.df_science['班级'] == i].reset_index(drop=True)
            class_science['序号'] = [k + 1 for k in class_science.index]
            class_science['综合'] = class_science['物理'] + class_science['化学'] + class_science['生物']
            class_science = class_science.loc[:, ['序号', '班级', '姓名', '语文', '数学', '英语', '综合',
                                                  '物理', '化学', '生物', '总分', '排名']]
            class_science.to_excel(writer, sheet_name=i, index=False)
        writer.close()

    def class_rank(self):
        """
        计算文理科学生排名
        :return:
        """
        self.df_arts['排名'] = self.df_arts['总分'].rank(method='min', ascending=False)
        # self.df_arts['排名'] = self.df_arts['排名'].apply(lambda x: format(int(x)))
        self.df_arts.sort_values(by='总分', ascending=False, inplace=True)

        self.df_science['排名'] = self.df_science['总分'].rank(method='min', ascending=False)
        # self.df_science['排名'] = self.df_science['排名'].apply(lambda x: format(int(x)))
        self.df_science.sort_values(by='总分', ascending=False, inplace=True)

    def score_label(self):
        """
        计算打印考生个人成绩单
        """
        self.class_rank()
        exam_arts = self.df_arts.loc[:,
                    ['班级', '姓名', '语文', '数学', '英语', '政治', '历史', '地理', '总分', '排名']]
        exam_science = self.df_science.loc[:,
                       ['班级', '姓名', '语文', '数学', '英语', '物理', '化学', '生物', '总分', '排名']]
        exam_arts.sort_values(by=['班级', '总分'], inplace=True, ascending=[True, False], ignore_index=True)
        exam_science.sort_values(by=['班级', '总分'], inplace=True, ascending=[True, False], ignore_index=True)

        for i in exam_arts.index:
            exam_arts.loc[i + 0.5] = exam_arts.columns
        exam_arts.sort_index(inplace=True, ignore_index=True)
        for i in exam_science.index:
            exam_science.loc[i + 0.5] = exam_science.columns
        exam_science.sort_index(inplace=True)
        with pd.ExcelWriter('本次考试学生个人成绩单.xlsx') as writer:
            exam_arts.to_excel(writer, sheet_name='文科成绩单', index=False)
            exam_science.to_excel(writer, sheet_name='理科成绩单', index=False)

    def arts_science_combined(self):
        arts_av, science_av = self.get_av()
        arts, grades_arts = self.goodscore_arts(*ScoreAnalysis.arts_scores)
        science, grades_science = self.goodscore_science(*ScoreAnalysis.science_scores)
        arts_av = self.write_open(arts_av)
        science_av = self.write_open(science_av)
        grades_arts = self.write_open(grades_arts)
        grades_science = self.write_open(grades_science)
        arts_science = pd.concat([grades_arts, grades_science, arts_av, science_av], ignore_index=True)
        # with pd.ExcelWriter(r'D:\成绩统计结果\文理有效分统计分析.xlsx') as writer:
        #     arts_science.to_excel(writer, sheet_name='文理有效分', index=False)
        result_file = self.file_name_by_time(arts_science)
        return result_file

    def arts_science_combined_school(self, goodtotal_arts=None, goodtotal_science=None):
        chn, math, eng, pol, his, geo, total = self.get_goodscores_arts(goodtotal_arts)
        chn_science, math_science, eng_science, phys, chem, bio, total_science = self.get_goodscores_science(
            goodtotal_science)
        arts_av, science_av = self.get_av()

        arts, grades_arts = self.goodscore_arts(chn, math, eng, pol, his, geo, total)
        science, grades_science = self.goodscore_science(chn_science, math_science, eng_science,
                                                         phys, chem, bio,
                                                         total_science)
        arts_av = self.write_open(arts_av)
        science_av = self.write_open(science_av)
        grades_arts = self.write_open(grades_arts)
        grades_science = self.write_open(grades_science)
        arts_science = pd.concat([grades_arts, grades_science, arts_av, science_av])

        # with pd.ExcelWriter(r'D:\成绩统计结果\文理有效分统计分析.xlsx') as writer:
        #     arts_science.to_excel(writer, sheet_name='文理有效分', index=False)
        # result_file = '文理有效分统计分析.xlsx'
        result_file = self.file_name_by_time(arts_science)
        return result_file

    def good_scores_arts_ratio(self, chn, math, eng, pol, his, geo, total):
        # 计算各班单有效学生人数
        single_chn_arts = self.df_arts[self.df_arts['语文'] >= chn].groupby(['班级'])['语文'].count()
        single_math_arts = self.df_arts[self.df_arts['数学'] >= math].groupby(['班级'])['数学'].count()
        single_eng_arts = self.df_arts[self.df_arts['英语'] >= eng].groupby(['班级'])['英语'].count()
        single_pol_arts = self.df_arts[self.df_arts['政治'] >= pol].groupby(['班级'])['政治'].count()
        single_his_arts = self.df_arts[self.df_arts['历史'] >= his].groupby(['班级'])['历史'].count()
        single_geo_arts = self.df_arts[self.df_arts['地理'] >= geo].groupby(['班级'])['地理'].count()
        single_total_arts = self.df_arts[self.df_arts['总分'] >= total].groupby(['班级'])['总分'].count()
        # 计算参考人数
        name_num = self.df_arts.groupby(['班级'])['姓名'].count()
        name_num.name = '参考人数'
        # 计算双有效各班学生人数
        df2 = self.df_arts[self.df_arts['总分'] >= total]
        double_chn_arts = df2[df2['语文'] >= chn].groupby(['班级'])['语文'].count()
        double_math_arts = df2[df2['数学'] >= math].groupby(['班级'])['数学'].count()
        double_eng_arts = df2[df2['英语'] >= eng].groupby(['班级'])['英语'].count()
        double_pol_arts = df2[df2['政治'] >= pol].groupby(['班级'])['政治'].count()
        double_his_arts = df2[df2['历史'] >= his].groupby(['班级'])['历史'].count()
        double_geo_arts = df2[df2['地理'] >= geo].groupby(['班级'])['地理'].count()
        double_total_arts = df2[df2['总分'] >= total].groupby(['班级'])['总分'].count()

        # goodscore_dict = {'参考人数': ' ', '语文': chn, '数学': math, '英语': eng, '政治': pol, '历史': his, '地理': geo, '总分': total}
        # goodscore_df = pd.DataFrame(goodscore_dict, index=['有效分数'])

        result_single = pd.concat([name_num, single_chn_arts, single_math_arts, single_eng_arts,
                                   single_pol_arts, single_his_arts, single_geo_arts, single_total_arts],
                                  axis=1)

        result_double = pd.concat(
            [name_num, double_chn_arts, double_math_arts, double_eng_arts,
             double_pol_arts, double_his_arts, double_geo_arts, double_total_arts], axis=1)

        result_single.loc['文科共计'] = [result_single['参考人数'].sum(),
                                         result_single['语文'].sum(),
                                         result_single['数学'].sum(),
                                         result_single['英语'].sum(),
                                         result_single['政治'].sum(),
                                         result_single['历史'].sum(),
                                         result_single['地理'].sum(),
                                         result_single['总分'].sum()
                                         ]
        # 新增上线率一列并用百分数表示
        # result_single['上线率'] = result_single['总分'] / result_single['参考人数']
        # result_single['上线率'] = result_single['上线率'].apply(lambda x: format(x, '.2%'))
        result_single = self.good_scores_ratio(result_single, '总分')
        result_single = self.good_scores_ratio(result_single, '语文')
        result_single = self.good_scores_ratio(result_single, '数学')
        result_single = self.good_scores_ratio(result_single, '英语')
        result_single = self.good_scores_ratio(result_single, '政治')
        result_single = self.good_scores_ratio(result_single, '历史')
        result_single = self.good_scores_ratio(result_single, '地理')
        order = ['参考人数', '语文', '语文上线率', '数学', '数学上线率', '英语', '英语上线率',
                 '政治', '政治上线率', '历史', '历史上线率', '地理', '地理上线率', '总分', '总分上线率']
        result_single = result_single[order]
        result_double.loc['文科共计'] = [result_double['参考人数'].sum(),
                                         result_double['语文'].sum(),
                                         result_double['数学'].sum(),
                                         result_double['英语'].sum(),
                                         result_double['政治'].sum(),
                                         result_double['历史'].sum(),
                                         result_double['地理'].sum(),
                                         result_double['总分'].sum()]
        result_double = self.good_scores_ratio(result_double, '总分')
        result_double = self.good_scores_ratio(result_double, '语文')
        result_double = self.good_scores_ratio(result_double, '数学')
        result_double = self.good_scores_ratio(result_double, '英语')
        result_double = self.good_scores_ratio(result_double, '政治')
        result_double = self.good_scores_ratio(result_double, '历史')
        result_double = self.good_scores_ratio(result_double, '地理')
        order = ['参考人数', '语文', '语文上线率', '数学', '数学上线率', '英语', '英语上线率',
                 '政治', '政治上线率', '历史', '历史上线率', '地理', '地理上线率', '总分', '总分上线率']
        result_double = result_double[order]
        single_double_concat = pd.concat([result_single, result_double], keys=['单有效', '双有效'])

        return single_double_concat

    def good_scores_science_ratio(self, chn, math, eng, pol, his, geo, total):

        single_chn_science = self.df_science[self.df_science['语文'] >= chn].groupby(['班级'])['语文'].count()
        single_math_science = self.df_science[self.df_science['数学'] >= math].groupby(['班级'])['数学'].count()
        single_eng_science = self.df_science[self.df_science['英语'] >= eng].groupby(['班级'])['英语'].count()
        single_phys_science = self.df_science[self.df_science['物理'] >= pol].groupby(['班级'])['物理'].count()
        single_chem_science = self.df_science[self.df_science['化学'] >= his].groupby(['班级'])['化学'].count()
        single_bio_science = self.df_science[self.df_science['生物'] >= geo].groupby(['班级'])['生物'].count()
        single_total_science = self.df_science[self.df_science['总分'] >= total].groupby(['班级'])['总分'].count()

        name_num = self.df_science.groupby(['班级'])['姓名'].count()
        name_num.name = '参考人数'

        df2 = self.df_science[self.df_science['总分'] >= total]
        double_chn_science = df2[df2['语文'] >= chn].groupby(['班级'])['语文'].count()
        double_math_science = df2[df2['数学'] >= math].groupby(['班级'])['数学'].count()
        double_eng_science = df2[df2['英语'] >= eng].groupby(['班级'])['英语'].count()
        double_phys_science = df2[df2['物理'] >= pol].groupby(['班级'])['物理'].count()
        double_chem_science = df2[df2['化学'] >= his].groupby(['班级'])['化学'].count()
        double_bio_science = df2[df2['生物'] >= geo].groupby(['班级'])['生物'].count()
        double_total__science = df2[df2['总分'] >= total].groupby(['班级'])['总分'].count()

        result_single = pd.concat([name_num, single_chn_science, single_math_science, single_eng_science,
                                   single_phys_science, single_chem_science, single_bio_science, single_total_science],
                                  axis=1)
        result_double = pd.concat(
            [name_num, double_chn_science, double_math_science, double_eng_science, double_phys_science,
             double_chem_science, double_bio_science, double_total__science], axis=1)

        result_single.loc['理科共计'] = [result_single['参考人数'].sum(),
                                         result_single['语文'].sum(),
                                         result_single['数学'].sum(),
                                         result_single['英语'].sum(),
                                         result_single['物理'].sum(),
                                         result_single['化学'].sum(),
                                         result_single['生物'].sum(),
                                         result_single['总分'].sum()
                                         ]
        # result_single['上线率'] = result_single['总分'] / result_single['参考人数']
        # result_single['上线率'] = result_single['上线率'].apply(lambda x: format(x, '.2%'))

        result_single = self.good_scores_ratio(result_single, '总分')
        result_single = self.good_scores_ratio(result_single, '语文')
        result_single = self.good_scores_ratio(result_single, '数学')
        result_single = self.good_scores_ratio(result_single, '英语')
        result_single = self.good_scores_ratio(result_single, '物理')
        result_single = self.good_scores_ratio(result_single, '化学')
        result_single = self.good_scores_ratio(result_single, '生物')
        order = ['参考人数', '语文', '语文上线率', '数学', '数学上线率', '英语', '英语上线率',
                 '物理', '物理上线率', '化学', '化学上线率', '生物', '生物上线率', '总分', '总分上线率']
        result_single = result_single[order]

        result_double.loc['理科共计'] = [result_double['参考人数'].sum(),
                                         result_double['语文'].sum(),
                                         result_double['数学'].sum(),
                                         result_double['英语'].sum(),
                                         result_double['物理'].sum(),
                                         result_double['化学'].sum(),
                                         result_double['生物'].sum(),
                                         result_double['总分'].sum()]
        result_double = self.good_scores_ratio(result_double, '总分')
        result_double = self.good_scores_ratio(result_double, '语文')
        result_double = self.good_scores_ratio(result_double, '数学')
        result_double = self.good_scores_ratio(result_double, '英语')
        result_double = self.good_scores_ratio(result_double, '物理')
        result_double = self.good_scores_ratio(result_double, '化学')
        result_double = self.good_scores_ratio(result_double, '生物')
        order = ['参考人数', '语文', '语文上线率', '数学', '数学上线率', '英语', '英语上线率',
                 '物理', '物理上线率', '化学', '化学上线率', '生物', '生物上线率', '总分', '总分上线率']
        result_double = result_double[order]
        single_double_concat_science = pd.concat([result_single, result_double], keys=['单有效', '双有效'])

        return single_double_concat_science

    @staticmethod
    def file_name_by_time(data_df):
        file_time = time.time()
        file_name = f'{file_time}.xlsx'
        data_df.to_excel('D:\\成绩统计结果\\' + file_name, sheet_name='成绩分析', index=False)
        return file_name

    @staticmethod
    def good_scores_ratio(data, subject):
        data[subject + '上线率'] = data[subject] / data['参考人数']
        data[subject + '上线率'] = data[subject + '上线率'].apply(lambda x: format(x, '.1%'))
        return data

    @staticmethod
    def get_subject_good_score(data, subject, total):
        """
        获取各科有效分
        :param data: df数据
        :param subject: 学科名
        :param total: 上线总分
        :return: 学科有效分
        """
        good_score_data = data.loc[data['总分'] >= total]
        subject_av = good_score_data[subject].mean()
        total_av = good_score_data['总分'].mean()
        subject_good_score = round(subject_av * total / total_av)
        return subject_good_score

    @staticmethod
    def get_single_double_score(data, subject, subject_score, total_score):
        """
        get the good scores of a subject in an exam
        :param data:
        :param subject:
        :param subject_score:
        :param total_score:
        :return:
        """
        single = data[data[subject] >= subject_score].groupby(['班级'])[subject].count()
        data_double = data[data['总分'] >= total_score]
        double = data_double[data_double[subject] >= subject_score].groupby(['班级'])[subject].count()
        return single, double

    @staticmethod
    def get_unmatched_students(data, subject, subject_score, total_score):
        """
        get the unmatched students in an exam
        :param data:
        :param subject:
        :param subject_score:
        :param total_score:
        :return:
        """
        df2 = data[data['总分'] >= total_score]
        df_unmatched_students = df2.loc[:, ['班级', '姓名', subject]].loc[df2[subject] < subject_score].sort_values(
            by=['班级', subject], ascending=[True, False]).reset_index(drop=True)
        return df_unmatched_students

    @staticmethod
    def write_open(df_data):
        df_data.to_excel('temp_data.xlsx')
        df_new = pd.read_excel('temp_data.xlsx', header=None)
        os.remove('temp_data.xlsx')
        return df_new

    @staticmethod
    def make_directory():
        if not os.path.exists('D:\\成绩统计结果'):
            os.makedirs('D:\\成绩统计结果')


class JuniorExam(object):

    def __init__(self, filepath):
        self.filepath = filepath
        self.df = pd.read_excel(self.filepath, sheet_name='总表')

    def junior_scores(self):
        """
        计算初中考试合格人数及合格率
        :return: 生成一个excel文件
        """
        if '化学' in self.df.columns:
            qualification_df = self.df.loc[:,
                               ['序号', '姓名', '班级', '语文A卷', '数学A卷', '英语A卷', '物理A卷', '化学', '总分']]
            # qualification_df.columns = ['序号', '姓名', '班级', '语文', '数学', '英语', '物理', '化学', '总分']

            # 计算合格人数
            single_chn = qualification_df[qualification_df['语文A卷'] >= 60].groupby(['班级'])['语文A卷'].count()
            single_math = qualification_df[qualification_df['数学A卷'] >= 60].groupby(['班级'])['数学A卷'].count()
            single_eng = qualification_df[qualification_df['英语A卷'] >= 60].groupby(['班级'])['英语A卷'].count()
            single_phys = qualification_df[qualification_df['物理A卷'] >= 60].groupby(['班级'])['物理A卷'].count()
            single_chem = qualification_df[qualification_df['化学'] >= 60].groupby(['班级'])['化学'].count()

            name_num = qualification_df.groupby(['班级'])['姓名'].count()
            name_num.name = '参考人数'

            qualified_data = pd.concat([name_num, single_chn, single_math, single_eng,
                                        single_phys, single_chem],
                                       axis=1)

            qualified_data.loc['年级'] = [qualified_data['参考人数'].sum(),
                                          qualified_data['语文A卷'].sum(),
                                          qualified_data['数学A卷'].sum(),
                                          qualified_data['英语A卷'].sum(),
                                          qualified_data['物理A卷'].sum(),
                                          qualified_data['化学'].sum()]
            # 计算合格率
            qualified_data['语文合格率'] = qualified_data['语文A卷'] / qualified_data['参考人数']
            qualified_data['语文合格率'] = qualified_data['语文合格率'].apply(lambda x: format(x, '.2%'))

            qualified_data['数学合格率'] = qualified_data['数学A卷'] / qualified_data['参考人数']
            qualified_data['数学合格率'] = qualified_data['数学合格率'].apply(lambda x: format(x, '.2%'))

            qualified_data['英语合格率'] = qualified_data['英语A卷'] / qualified_data['参考人数']
            qualified_data['英语合格率'] = qualified_data['英语合格率'].apply(lambda x: format(x, '.2%'))

            qualified_data['物理合格率'] = qualified_data['物理A卷'] / qualified_data['参考人数']
            qualified_data['物理合格率'] = qualified_data['物理合格率'].apply(lambda x: format(x, '.2%'))

            qualified_data['化学合格率'] = qualified_data['化学'] / qualified_data['参考人数']
            qualified_data['化学合格率'] = qualified_data['化学合格率'].apply(lambda x: format(x, '.2%'))

            qualified_data.reset_index(inplace=True)
            new_columns = ['班级', '参考人数', '语文A卷', '语文合格率', '数学A卷', '数学合格率', '英语A卷',
                           '英语合格率',
                           '物理A卷', '物理合格率', '化学', '化学合格率']
            qualified_data = qualified_data[new_columns]

            return qualified_data
        elif '物理A卷' not in self.df.columns:
            qualification_df = self.df.loc[:, ['序号', '姓名', '班级', '语文A卷', '数学A卷', '英语A卷', '总分']]
            # qualification_df.columns = ['序号', '姓名', '班级', '语文', '数学', '英语', '总分']

            # 计算合格人数
            single_chn = qualification_df[qualification_df['语文A卷'] >= 60].groupby(['班级'])['语文A卷'].count()
            single_math = qualification_df[qualification_df['数学A卷'] >= 60].groupby(['班级'])['数学A卷'].count()
            single_eng = qualification_df[qualification_df['英语A卷'] >= 60].groupby(['班级'])['英语A卷'].count()

            name_num = qualification_df.groupby(['班级'])['姓名'].count()
            name_num.name = '参考人数'

            qualified_data = pd.concat([name_num, single_chn, single_math, single_eng], axis=1)

            qualified_data.loc['年级'] = [qualified_data['参考人数'].sum(),
                                          qualified_data['语文A卷'].sum(),
                                          qualified_data['数学A卷'].sum(),
                                          qualified_data['英语A卷'].sum(),
                                          ]
            # 计算合格率
            qualified_data['语文合格率'] = qualified_data['语文A卷'] / qualified_data['参考人数']
            qualified_data['语文合格率'] = qualified_data['语文合格率'].apply(lambda x: format(x, '.2%'))

            qualified_data['数学合格率'] = qualified_data['数学A卷'] / qualified_data['参考人数']
            qualified_data['数学合格率'] = qualified_data['数学合格率'].apply(lambda x: format(x, '.2%'))

            qualified_data['英语合格率'] = qualified_data['英语A卷'] / qualified_data['参考人数']
            qualified_data['英语合格率'] = qualified_data['英语合格率'].apply(lambda x: format(x, '.2%'))

            qualified_data.reset_index(inplace=True)
            new_columns = ['班级', '参考人数', '语文A卷', '语文合格率', '数学A卷', '数学合格率', '英语A卷',
                           '英语合格率']
            qualified_data = qualified_data[new_columns]

            return qualified_data
        else:
            qualification_df = self.df.loc[:,
                               ['序号', '姓名', '班级', '语文A卷', '数学A卷', '英语A卷', '物理A卷', '总分']]
            # qualification_df.columns = ['序号', '姓名', '班级', '语文', '数学', '英语', '物理', '总分']

            # 计算合格人数
            single_chn = qualification_df[qualification_df['语文A卷'] >= 60].groupby(['班级'])['语文A卷'].count()
            single_math = qualification_df[qualification_df['数学A卷'] >= 60].groupby(['班级'])['数学A卷'].count()
            single_eng = qualification_df[qualification_df['英语A卷'] >= 60].groupby(['班级'])['英语A卷'].count()
            single_phys = qualification_df[qualification_df['物理A卷'] >= 60].groupby(['班级'])['物理A卷'].count()

            name_num = qualification_df.groupby(['班级'])['姓名'].count()
            name_num.name = '参考人数'

            qualified_data = pd.concat([name_num, single_chn, single_math, single_eng,
                                        single_phys], axis=1)

            qualified_data.loc['年级'] = [qualified_data['参考人数'].sum(),
                                          qualified_data['语文A卷'].sum(),
                                          qualified_data['数学A卷'].sum(),
                                          qualified_data['英语A卷'].sum(),
                                          qualified_data['物理A卷'].sum()]
            # 计算合格率
            qualified_data['语文合格率'] = qualified_data['语文A卷'] / qualified_data['参考人数']
            qualified_data['语文合格率'] = qualified_data['语文合格率'].apply(lambda x: format(x, '.2%'))

            qualified_data['数学合格率'] = qualified_data['数学A卷'] / qualified_data['参考人数']
            qualified_data['数学合格率'] = qualified_data['数学合格率'].apply(lambda x: format(x, '.2%'))

            qualified_data['英语合格率'] = qualified_data['英语A卷'] / qualified_data['参考人数']
            qualified_data['英语合格率'] = qualified_data['英语合格率'].apply(lambda x: format(x, '.2%'))

            qualified_data['物理合格率'] = qualified_data['物理A卷'] / qualified_data['参考人数']
            qualified_data['物理合格率'] = qualified_data['物理合格率'].apply(lambda x: format(x, '.2%'))

            qualified_data.reset_index(inplace=True)
            new_columns = ['班级', '参考人数', '语文A卷', '语文合格率', '数学A卷', '数学合格率', '英语A卷',
                           '英语合格率',
                           '物理A卷', '物理合格率']
            qualified_data = qualified_data[new_columns]

            return qualified_data

    def get_av(self):
        # 计算平均分
        if '化学' in self.df.columns:

            av_class = self.df.groupby(['班级'])[['语文A卷', '语文B卷', '语文合卷', '数学A卷', '数学B卷', '数学合卷',
                                                  '英语A卷', '英语B卷', '英语合卷', '物理A卷', '物理B卷', '物理合卷',
                                                  '化学', '总分']].mean().round(2)
            av_general = self.df[['语文A卷', '语文B卷', '语文合卷', '数学A卷', '数学B卷', '数学合卷',
                                  '英语A卷', '英语B卷', '英语合卷', '物理A卷', '物理B卷', '物理合卷',
                                  '化学', '总分']].apply(np.nanmean, axis=0).round(2)
            av_general.name = '年级平均分'
            av_scores = av_class.append(av_general)
            av_scores.reset_index(inplace=True)

            return av_scores
        elif '物理合卷' not in self.df.columns:
            av_class = self.df.groupby(['班级'])[['语文A卷', '语文B卷', '语文合卷', '数学A卷', '数学B卷', '数学合卷',
                                                  '英语A卷', '英语B卷', '英语合卷', '总分']].mean().round(2)
            av_general = self.df[['语文A卷', '语文B卷', '语文合卷', '数学A卷', '数学B卷', '数学合卷',
                                  '英语A卷', '英语B卷', '英语合卷', '总分']].apply(np.nanmean, axis=0).round(2)
            av_general.name = '年级平均分'
            av_scores = av_class.append(av_general)
            av_scores.reset_index(inplace=True)
            return av_scores
        else:
            av_class = self.df.groupby(['班级'])[['语文A卷', '语文B卷', '语文合卷', '数学A卷', '数学B卷', '数学合卷',
                                                  '英语A卷', '英语B卷', '英语合卷',
                                                  '物理A卷', '物理B卷', '物理合卷', '总分']].mean().round(2)
            av_general = self.df[['语文A卷', '语文B卷', '语文合卷', '数学A卷', '数学B卷', '数学合卷',
                                  '英语A卷', '英语B卷', '英语合卷', '物理A卷',
                                  '物理B卷', '物理合卷', '总分']].apply(np.nanmean, axis=0).round(2)
            av_general.name = '年级平均分'
            av_scores = av_class.append(av_general)
            av_scores.reset_index(inplace=True)
            return av_scores

    def get_goodscores(self, goodtotal):
        """
        计算各科有效分
        goodtotal:划线总分，高线，中线，低线
        """
        if '化学' in self.df.columns:
            good_score_df = self.df.loc[:,
                            ['序号', '姓名', '班级', '语文合卷', '数学合卷', '英语合卷', '物理合卷', '化学', '总分']]
            # good_score_df.columns = ['序号', '姓名', '班级', '语文', '数学', '英语', '物理', '化学', '总分']

            goodscoredata = good_score_df.loc[good_score_df['总分'] >= goodtotal]
            chnav = goodscoredata['语文合卷'].mean()
            mathav = goodscoredata['数学合卷'].mean()
            engav = goodscoredata['英语合卷'].mean()
            phyav = goodscoredata['物理合卷'].mean()
            chemav = goodscoredata['化学'].mean()
            totalav = goodscoredata['总分'].mean()
            factor = goodtotal / totalav
            chn = round(chnav * factor)
            math = round(mathav * factor)
            eng = round(engav * factor)
            phy = round(phyav * factor)
            chem = round(chemav * factor)
            if (chn + math + eng + phy + chem) > goodtotal:
                math -= 1
            if (chn + math + eng + phy + chem) < goodtotal:
                chn += 1
            return chn, math, eng, phy, chem, goodtotal

        elif '物理合卷' not in self.df.columns:
            good_score_df = self.df.loc[:, ['序号', '姓名', '班级', '语文合卷', '数学合卷', '英语合卷', '总分']]
            # good_score_df.columns = ['序号', '姓名', '班级', '语文', '数学', '英语', '总分']

            goodscoredata = good_score_df.loc[good_score_df['总分'] >= goodtotal]
            chnav = goodscoredata['语文合卷'].mean()
            mathav = goodscoredata['数学合卷'].mean()
            engav = goodscoredata['英语合卷'].mean()
            totalav = goodscoredata['总分'].mean()

            factor = goodtotal / totalav

            chn = round(chnav * factor)
            math = round(mathav * factor)
            eng = round(engav * factor)

            if (chn + math + eng) > goodtotal:
                math -= 1
            if (chn + math + eng) < goodtotal:
                chn += 1
            return chn, math, eng, goodtotal

        else:
            good_score_df = self.df.loc[:,
                            ['序号', '姓名', '班级', '语文合卷', '数学合卷', '英语合卷', '物理合卷', '总分']]
            # good_score_df.columns = ['序号', '姓名', '班级', '语文', '数学', '英语', '物理', '总分']
            goodscoredata = good_score_df.loc[good_score_df['总分'] >= goodtotal]
            chnav = goodscoredata['语文合卷'].mean()
            mathav = goodscoredata['数学合卷'].mean()
            engav = goodscoredata['英语合卷'].mean()
            phyav = goodscoredata['物理合卷'].mean()
            totalav = goodscoredata['总分'].mean()
            factor = goodtotal / totalav
            chn = round(chnav * factor)
            math = round(mathav * factor)
            eng = round(engav * factor)
            phy = round(phyav * factor)
            if (chn + math + eng + phy) > goodtotal:
                math -= 1
            if (chn + math + eng + phy) < goodtotal:
                chn += 1
            return chn, math, eng, phy, goodtotal

    def goodscore_process(self, goodtotal):
        """
        计算各科各班单有效和双有效人数
        """
        # 计算各班有效学生人数

        if '化学' in self.df.columns:
            good_score_df = self.df.loc[:,
                            ['序号', '姓名', '班级', '语文合卷', '数学合卷', '英语合卷', '物理合卷', '化学', '总分']]
            good_score_df.columns = ['序号', '姓名', '班级', '语文', '数学', '英语', '物理', '化学', '总分']

            chn, math, eng, phy, chem, total = self.get_goodscores(goodtotal)
            single_chn, double_chn = self.get_single_double_score(good_score_df, '语文', chn, total)
            single_math, double_math = self.get_single_double_score(good_score_df, '数学', math, total)
            single_eng, double_eng = self.get_single_double_score(good_score_df, '英语', eng, total)
            single_phy, double_phy = self.get_single_double_score(good_score_df, '物理', phy, total)
            single_chem, double_chem = self.get_single_double_score(good_score_df, '化学', chem, total)
            single_total, double_total = self.get_single_double_score(good_score_df, '总分', total, total)

            # 计算参考人数
            name_num = good_score_df.groupby(['班级'])['姓名'].count()
            name_num.name = '参考人数'

            goodscore_dict = {'参考人数': ' ', '语文': chn, '数学': math, '英语': eng,
                              '物理': phy, '化学': chem, '总分': total}
            goodscore_df = pd.DataFrame(goodscore_dict, index=['有效分数'])

            result_single = pd.concat([name_num, single_chn, single_math, single_eng,
                                       single_phy, single_chem, single_total],
                                      axis=1)

            result_double = pd.concat(
                [name_num, double_chn, double_math, double_eng,
                 double_phy, double_chem, double_total], axis=1)

            result_single.loc['共计'] = [result_single['参考人数'].sum(),
                                         result_single['语文'].sum(),
                                         result_single['数学'].sum(),
                                         result_single['英语'].sum(),
                                         result_single['物理'].sum(),
                                         result_single['化学'].sum(),
                                         result_single['总分'].sum()
                                         ]
            # 新增上线率一列并用百分数表示
            result_single['上线率'] = result_single['总分'] / result_single['参考人数']
            result_single['上线率'] = result_single['上线率'].apply(lambda x: format(x, '.2%'))
            # 新增一行文科共计。
            result_double.loc['共计'] = [result_double['参考人数'].sum(),
                                         result_double['语文'].sum(),
                                         result_double['数学'].sum(),
                                         result_double['英语'].sum(),
                                         result_double['物理'].sum(),
                                         result_double['化学'].sum(),
                                         result_double['总分'].sum()
                                         ]
            final_result = pd.concat([goodscore_df, result_single, result_double], keys=['有效分', '单有效', '双有效'])

            return final_result
        elif '物理合卷' not in self.df.columns:
            good_score_df = self.df.loc[:, ['序号', '姓名', '班级', '语文合卷', '数学合卷', '英语合卷', '总分']]
            good_score_df.columns = ['序号', '姓名', '班级', '语文', '数学', '英语', '总分']

            chn, math, eng, total = self.get_goodscores(goodtotal)
            single_chn, double_chn = self.get_single_double_score(good_score_df, '语文', chn, total)
            single_math, double_math = self.get_single_double_score(good_score_df, '数学', math, total)
            single_eng, double_eng = self.get_single_double_score(good_score_df, '英语', eng, total)
            single_total, double_total = self.get_single_double_score(good_score_df, '总分', total, total)

            # 计算参考人数
            name_num = good_score_df.groupby(['班级'])['姓名'].count()
            name_num.name = '参考人数'

            goodscore_dict = {'参考人数': ' ', '语文': chn, '数学': math, '英语': eng, '总分': total}
            goodscore_df = pd.DataFrame(goodscore_dict, index=['有效分数'])

            result_single = pd.concat([name_num, single_chn, single_math, single_eng, single_total], axis=1)

            result_double = pd.concat(
                [name_num, double_chn, double_math, double_eng, double_total], axis=1)

            result_single.loc['共计'] = [result_single['参考人数'].sum(),
                                         result_single['语文'].sum(),
                                         result_single['数学'].sum(),
                                         result_single['英语'].sum(),
                                         result_single['总分'].sum()]
            # 新增上线率一列并用百分数表示
            result_single['上线率'] = result_single['总分'] / result_single['参考人数']
            result_single['上线率'] = result_single['上线率'].apply(lambda x: format(x, '.2%'))
            # 新增一行文科共计。
            result_double.loc['共计'] = [result_double['参考人数'].sum(),
                                         result_double['语文'].sum(),
                                         result_double['数学'].sum(),
                                         result_double['英语'].sum(),
                                         result_double['总分'].sum()]
            final_result = pd.concat([goodscore_df, result_single, result_double], keys=['有效分', '单有效', '双有效'])

            return final_result

        else:
            good_score_df = self.df.loc[:,
                            ['序号', '姓名', '班级', '语文合卷', '数学合卷', '英语合卷', '物理合卷', '总分']]
            good_score_df.columns = ['序号', '姓名', '班级', '语文', '数学', '英语', '物理', '总分']
            chn, math, eng, phy, total = self.get_goodscores(goodtotal)
            single_chn, double_chn = self.get_single_double_score(good_score_df, '语文', chn, total)
            single_math, double_math = self.get_single_double_score(good_score_df, '数学', math, total)
            single_eng, double_eng = self.get_single_double_score(good_score_df, '英语', eng, total)
            single_phy, double_phy = self.get_single_double_score(good_score_df, '物理', phy, total)
            single_total, double_total = self.get_single_double_score(good_score_df, '总分', total, total)

            # 计算参考人数
            name_num = good_score_df.groupby(['班级'])['姓名'].count()
            name_num.name = '参考人数'

            goodscore_dict = {'参考人数': ' ', '语文': chn, '数学': math, '英语': eng,
                              '物理': phy, '总分': total}
            goodscore_df = pd.DataFrame(goodscore_dict, index=['有效分数'])

            result_single = pd.concat([name_num, single_chn, single_math, single_eng,
                                       single_phy, single_total],
                                      axis=1)

            result_double = pd.concat(
                [name_num, double_chn, double_math, double_eng,
                 double_phy, double_total], axis=1)

            result_single.loc['共计'] = [result_single['参考人数'].sum(),
                                         result_single['语文'].sum(),
                                         result_single['数学'].sum(),
                                         result_single['英语'].sum(),
                                         result_single['物理'].sum(),
                                         result_single['总分'].sum()
                                         ]
            # 新增上线率一列并用百分数表示
            result_single['上线率'] = result_single['总分'] / result_single['参考人数']
            result_single['上线率'] = result_single['上线率'].apply(lambda x: format(x, '.2%'))
            # 新增一行文科共计。
            result_double.loc['共计'] = [result_double['参考人数'].sum(),
                                         result_double['语文'].sum(),
                                         result_double['数学'].sum(),
                                         result_double['英语'].sum(),
                                         result_double['物理'].sum(),
                                         result_double['总分'].sum()
                                         ]
            final_result = pd.concat([goodscore_df, result_single, result_double], keys=['有效分', '单有效', '双有效'])

            return final_result

    def concat_results(self, goodtatal=None):
        good_score_results = self.goodscore_process(goodtatal)
        junior_av = self.get_av()
        junior_qualification = self.junior_scores()
        good_score_results = ScoreAnalysis.write_open(good_score_results)
        junior_av = ScoreAnalysis.write_open(junior_av)
        junior_qualification = ScoreAnalysis.write_open(junior_qualification)
        final_results = pd.concat([good_score_results, junior_av, junior_qualification])
        # with pd.ExcelWriter(r'D:\成绩统计结果\初中成绩统计分析.xlsx') as writer:
        #     final_results.to_excel(writer, sheet_name='成绩统计', index=False)
        final_file = ScoreAnalysis.file_name_by_time(final_results)
        final_file = os.path.basename(final_file)
        return final_file

    @staticmethod
    def get_single_double_score(data, subject, subject_score, total_score):
        single = data[data[subject] >= subject_score].groupby(['班级'])[subject].count()
        data_double = data[data['总分'] >= total_score]
        double = data_double[data_double[subject] >= subject_score].groupby(['班级'])[subject].count()
        return single, double


class ExamRoom(object):
    room_num = None

    def __init__(self, path):
        self.path = path

    def exam_room_info(self):
        """
        计算生成文理科各考室学生名单
        :return:
        """
        df_arts = pd.read_excel(self.path, sheet_name='文科', index_col='序号',
                                dtype={'班级': str, '考号': str, '序号': str})
        df_science = pd.read_excel(self.path, sheet_name='理科', index_col='序号',
                                   dtype={'班级': str, '考号': str, '序号': str})
        df_arts.sort_values(by='总分', ascending=False, inplace=True)
        df_science.sort_values(by='总分', ascending=False, inplace=True)
        room_numbers, room_numbers_science = self.get_room_number_arts_science(df_arts, df_science)

        # df_arts = df_arts.copy()
        # df_science = df_science.copy()

        df_arts['考室号'] = None
        df_arts['座位号'] = None
        df_arts = df_arts.loc[:, ['班级', '姓名', '考号', '考室号', '座位号']]
        df_arts.reset_index(drop=True, inplace=True)
        df_science['考室号'] = None
        df_science['座位号'] = None
        df_science = df_science.loc[:, ['班级', '姓名', '考号', '考室号', '座位号']]
        df_science.reset_index(drop=True, inplace=True)

        file_name = self.filename_by_time()
        writer = pd.ExcelWriter('D:\\成绩统计结果\\' + file_name)
        df_room_students = []
        df_room_students_science = []
        arts = pd.DataFrame(columns=df_arts.columns)
        science = pd.DataFrame(columns=df_science.columns)
        for idx, room_number in enumerate(room_numbers):
            begin = idx * ExamRoom.room_num
            end = begin + ExamRoom.room_num
            df_room_student = df_arts.iloc[begin:end]
            df_room_students.append((idx, room_number, df_room_student))

        for idx, room_number, df_room_student in df_room_students:
            for i in df_room_student.index:
                df_room_student = df_room_student.copy()
                df_room_student['考室号'].at[i] = room_number
                df_room_student['座位号'].at[i] = i + 1 if i < ExamRoom.room_num else i - idx * ExamRoom.room_num + 1
            df_room_student.to_excel(writer, sheet_name=room_number, index=False)
            arts = pd.concat([arts, df_room_student])

        for idx, room_number in enumerate(room_numbers_science):
            begin = idx * ExamRoom.room_num
            end = begin + ExamRoom.room_num
            df_room_student_science = df_science.iloc[begin:end]
            df_room_students_science.append((idx, room_number, df_room_student_science))
        for idx, room_number, df_room_student_science in df_room_students_science:
            for i in df_room_student_science.index:
                df_room_student_science = df_room_student_science.copy()
                df_room_student_science['考室号'].at[i] = room_number
                df_room_student_science['座位号'].at[
                    i] = i + 1 if i < ExamRoom.room_num else i - idx * ExamRoom.room_num + 1
            df_room_student_science.to_excel(writer, sheet_name=room_number, index=False)
            science = pd.concat([science, df_room_student_science])

        df_arts_general = arts.sort_values(by=['班级', '考室号'], ascending=[True, True])
        df_science_general = science.sort_values(by=['班级', '考室号'], ascending=[True, True])
        df_arts_general.to_excel(writer, sheet_name='文科', index=False)
        df_science_general.to_excel(writer, sheet_name='理科', index=False)
        # # 计算考生座签
        df_seat_arts = arts.copy()
        df_seat_science = science.copy()
        for i in df_seat_arts.index:
            df_seat_arts.loc[i + 0.5] = ['班级', '姓名', '考号', '考室号', '座位号']
        df_seat_arts.sort_index(inplace=True, ignore_index=True)
        for i in df_seat_science.index:
            df_seat_science.loc[i + 0.5] = ['班级', '姓名', '考号', '考室号', '座位号']
        df_seat_science.sort_index(inplace=True, ignore_index=True)
        df_seat_arts.to_excel(writer, sheet_name='文科座签', index=False)
        df_seat_science.to_excel(writer, sheet_name='理科座签', index=False)
        writer.close()
        # result_file = 'D:\\成绩统计结果\\' + file_name
        # final_file = os.path.basename(result_file)
        # final_file = self.file_zip(os.path.basename(result_file))
        return file_name

    @staticmethod
    def get_room_number_arts_science(df_arts, df_science):
        if len(df_arts) % ExamRoom.room_num != 0:
            room_numbers = [f'文科第{str(i + 1)}考室' for i in list(range(len(df_arts) // ExamRoom.room_num + 1))]
        else:
            room_numbers = [f'文科第{str(i + 1)}考室' for i in list(range(len(df_arts) // ExamRoom.room_num))]
        if len(df_science) % ExamRoom.room_num != 0:
            room_numbers_science = [f'理科第{str(i + 1)}考室' for i in
                                    list(range(len(df_science) // ExamRoom.room_num + 1))]
        else:
            room_numbers_science = [f'理科第{str(i + 1)}考室' for i in
                                    list(range(len(df_science) // ExamRoom.room_num))]
        return room_numbers, room_numbers_science

    @staticmethod
    def filename_by_time():
        file_time = time.time()
        file_name = f'{file_time}.xlsx'
        return file_name

    def exam_room(self):
        """
        计算生成高一文理不分科及初中各考室学生名单
        :return:
        """
        df = pd.read_excel(self.path, sheet_name='总表', index_col='序号',
                           dtype={'班级': str, '考号': str, '序号': str})
        df.sort_values(by='总分', ascending=False, inplace=True)
        room_numbers = self.get_room_number(df)
        df['考室号'] = None
        df['座位号'] = None
        df = df.loc[:, ['班级', '姓名', '考号', '考室号', '座位号']]
        df.reset_index(drop=True, inplace=True)
        # print(room_numbers)
        df_room_students = []
        df_temp = pd.DataFrame(columns=df.columns)
        for idx, room_number in enumerate(room_numbers):
            begin = idx * ExamRoom.room_num
            end = begin + ExamRoom.room_num
            df_room_student = df.iloc[begin:end]
            df_room_students.append((idx, room_number, df_room_student))

        file_name = self.filename_by_time()
        writer = pd.ExcelWriter('D:\\成绩统计结果\\' + file_name)

        for idx, room_number, df_room_student in df_room_students:
            for i in df_room_student.index:
                df_room_student = df_room_student.copy()
                df_room_student['考室号'].at[i] = room_number
                df_room_student['座位号'].at[
                    i] = i + 1 if i < ExamRoom.room_num else i - idx * ExamRoom.room_num + 1
            df_room_student.to_excel(writer, sheet_name=room_number, index=False)
            df_temp = pd.concat([df_temp, df_room_student])
        df_temp.to_excel(writer, sheet_name='考室安排表', index=False)

        for i in df_temp.index:
            df_temp.loc[i + 0.5] = ['班级', '姓名', '考号', '考室号', '座位号']
        seat_label = df_temp.sort_index()
        seat_label.to_excel(writer, sheet_name='学生座签', index=False)
        writer.close()
        # result_file = 'D:\\成绩统计结果\\' + file_name
        # final_file = os.path.basename(result_file)
        # final_file = self.file_zip(os.path.basename(result_file))
        return file_name

    @staticmethod
    def get_room_number(df):
        if len(df) % ExamRoom.room_num != 0:
            room_numbers = [f'第{str(i + 1)}考室' for i in list(range(len(df) // ExamRoom.room_num + 1))]
        else:
            room_numbers = [f'第{str(i + 1)}考室' for i in list(range(len(df) // ExamRoom.room_num))]
        return room_numbers

    def exam_room_choice(self):
        try:
            result_file = self.exam_room_info()
            return result_file
        except:
            result_file = self.exam_room()
            return result_file


class ExamInvigilators(object):
    """
    获取监考安排表
    """
    exam_numbers = None
    room_numbers = None

    def __init__(self, filepath):
        self.filepath = filepath
        self.df = pd.read_excel(self.filepath, index_col='序号', dtype={'序号': str})

    def invigilation_table(self):

        df_teachers = self.df.copy()
        df_teachers = df_teachers.loc[:, '姓名']
        teacher_list = df_teachers.values
        random.seed(0)
        random.shuffle(teacher_list)
        df_room = pd.DataFrame(columns=['科目' + str(i + 1) for i in range(ExamInvigilators.exam_numbers)],
                               index=['第' + str(i + 1) + '考室' for i in range(ExamInvigilators.room_numbers)])
        df_room.index.rename('考室号', inplace=True)
        df_room.reset_index(inplace=True)
        nums = 0
        try:
            for i in df_room.columns[1:]:
                for k in df_room.index:
                    df_room[i].at[k] = teacher_list[nums]
                    if nums < len(teacher_list) - 1:
                        nums += 1
                    else:
                        nums = 0
        except KeyError:
            df_room.fillna(np.NaN)

        df_room.set_index('考室号', inplace=True)
        return df_room

    def exam_teachers(self):
        """
        计算分考室监考教师名单
        """
        df_room = self.invigilation_table()
        file_name = self.get_file_name()
        writer = pd.ExcelWriter('D:\\成绩统计结果\\' + file_name)
        df_room.to_excel(writer, sheet_name='监考安排表')
        teachers = df_room.values
        teacher_dict = {}
        for i in range(len(teachers)):
            for k in teachers[i]:
                teacher_dict[k] = teacher_dict.get(k, 0) + 1
        teacher_occupied = list(teacher_dict.items())
        teacher_occupied.sort(key=lambda x: x[1], reverse=True)
        invigilators = []
        invigilation_times = []
        for i in range(len(teacher_occupied)):
            teacher, times = teacher_occupied[i]
            invigilators.append(teacher)
            invigilation_times.append(times)
        #     print('{0:>6}老师：    监考{1:<}节'.format(teacher, times))
        # print(f'共有{len(teacher_occupied)}个老师参与本次监考')
        invigilation_df = pd.DataFrame({'序号': [i + 1 for i in range(len(invigilators))], '监考老师': invigilators,
                                        '监考次数': invigilation_times})
        invigilation_df.set_index('序号', inplace=True)
        invigilation_df.loc['共计'] = [invigilation_df['监考老师'].count(),
                                       invigilation_df['监考次数'].sum()]
        invigilation_df.to_excel(writer, sheet_name='监考次数统计')
        writer.close()
        result_file = 'D:\\成绩统计结果\\' + file_name
        final_file = os.path.basename(result_file)

        return final_file

    @staticmethod
    def get_file_name():
        file_time = time.time()
        file_name = f'{file_time}.xlsx'
        return file_name


class SplitClass(object):
    def __init__(self, path):
        self.path = path
        self.df = pd.read_excel(self.path, sheet_name=0, dtype={
            '考号': str,
            '考生号': str
        })

    def split_into_class(self):
        class_numbers = self.df['班级'].unique()
        whole_df = pd.DataFrame()
        for i in class_numbers:
            class_number = self.df[self.df['班级'] == i].reset_index(drop=True)
            class_number['序号'] = [k + 1 for k in class_number.index]
            whole_df = pd.concat([whole_df, class_number])
        filename = ExamRoom.filename_by_time()
        writer = pd.ExcelWriter('D:\\成绩统计结果\\' + filename)
        whole_df.to_excel(writer, sheet_name='班级序号表', index=False)
        # final_file = os.path.basename(writer)
        writer.close()
        return filename


class CatalogueCourses(object):

    def __init__(self, file_path):
        self.file_path = file_path

    def file_open(self):
        df = pd.read_excel(self.file_path, sheet_name=0)
        return df

    def statistics_for_courses(self):
        data = self.file_open()
        class_count = self.get_course_data(data)
        # writer = pd.ExcelWriter(r'D:\成绩统计结果\汇总统计表.xlsx')
        # class_count.to_excel(writer, sheet_name='统计表', na_rep=None)
        # writer.close()
        return class_count

    def get_course_data(self, data):
        courses = data.columns[3:].tolist()
        class_count = data.groupby('班级')[courses].count()
        general_count = data[courses].count()
        class_count.loc['年级'] = general_count
        class_count['合计'] = class_count[courses].apply(np.sum, axis=1)
        class_count['班级人数'] = data['班级'].value_counts()
        class_count['班级人数'].at['年级'] = class_count['班级人数'].sum()
        class_count = class_count[['班级人数'] + courses + ['合计']]
        class_count.reset_index(inplace=True)
        class_count['班级人数'] = class_count['班级人数'].astype(int)
        return class_count

    def split_by_subject(self):
        data = self.file_open()
        class_count = self.get_course_data(data)
        courses = data.columns[3:].tolist()
        filename = ExamRoom.filename_by_time()
        writer = pd.ExcelWriter('D:\\成绩统计结果\\' + filename)
        class_count.to_excel(writer, sheet_name='汇总统计表', index=False)
        for i in range(len(courses)):
            new_data = data.loc[data[courses[i]].notnull(),
                                ['序号', '姓名', '班级', courses[i]]]
            new_data.reset_index(inplace=True, drop=True)
            new_data['序号'] = [k + 1 for k in new_data.index]
            new_data.to_excel(writer, sheet_name=courses[i], index=False)
        writer.close()
        return filename


class GetInfoFromId(object):
    def __init__(self, filename):
        self.filename = filename

    def make_regions_dict(self, regionNo, region):
        """
        生成身份证号前6位地区字典
        :param regionNo:
        :param region:
        :return:
        """
        filename = r'./static/id_region.xlsx'
        data = pd.read_excel(filename, sheet_name=0, dtype={'cityNo': str,
                                                            'districtNo': str,
                                                            'provinceNo': str})
        code_list = data[regionNo].tolist()
        district_list = data[region].tolist()
        region_dict = dict(zip(code_list, district_list))
        return region_dict

    def open_file(self):
        data = pd.read_excel(self.filename, dtype={'身份证号': str}, sheet_name=0)
        return data

    def data_remove_string(self, data, remove_string, data_column='身份证号'):
        """
        去除身份证号的多余字符,主要用于把学籍号转换成身份证号
        :param data:
        :param data_column: 在excel表上新建身份证号一列，把学籍号一列复制到这一列。
        :param remove_string: 要去除的字符。
        :return: 返回整理好的DataFrame数据
        """
        data[data_column] = data[data_column].str.replace(remove_string, '')
        return data

    # 清洗身份证号数据格式
    def get_clean_id(self):
        data = self.open_file()
        re_pattern = r'^([1-9]\d{5}[12]\d{3}(0[1-9]|1[012])(0[1-9]|[12][0-9]|3[01])\d{3}[0-9xX])$'

        # data = data.loc[data['身份证号'].notna()].reset_index()
        bad_df = pd.DataFrame(columns=data.columns)

        id_nums = []
        id_nums_bad = []
        for i in data.index:
            data['身份证号'] = data['身份证号'].astype(str)
            id_num = data['身份证号'].at[i].strip('\t').strip()
            if re.match(re_pattern, id_num):
                id_nums.append(id_num)
            else:
                id_nums_bad.append(id_num)
                bad_df.loc[i] = data.loc[i]
                data.drop(index=[i], inplace=True)
        data['身份证号'] = id_nums
        bad_df['身份证号'] = id_nums_bad
        data.reset_index(inplace=True, drop=True)
        if len(id_nums_bad) != 0:
            # print(f'错误的身份证号：{id_nums_bad},共计有{len(id_nums_bad)}个错误。请修正后再次运行。')
            bad_df = bad_df.loc[:, ['姓名', '身份证号']]
            # print(bad_df)
        # else:
        #     print("太棒了！经电脑检测，所有身份证号无误。信息提取完整！谢谢使用！")
        return data, bad_df, id_nums_bad

    def get_sex_birth_age(self):
        data, bad_df, id_bad = self.get_clean_id()
        regions = ['华北', '东北', '华南', '中南', '西南', '西北', '港澳台']
        province_dict = {
            '11': '北京市', '12': '天津市', '13': '河北省', '14': '山西省', '15': '内蒙古自治区',
            '21': '辽宁省', '22': '吉林省', '23': '黑龙江省',
            '31': '上海市', '32': '江苏省', '33': '浙江省', '34': '安徽省', '35': '福建省', '36': '江西省',
            '37': '山东省',
            '41': '河南省', '42': '湖北省', '43': '湖南省', '44': '广东省', '45': '广西壮族自治区', '46': '海南省',
            '51': '四川省', '52': '贵州省', '53': '云南省', '54': '西藏自治区', '50': '重庆市',
            '61': '陕西省', '62': '甘肃省', '63': '青海省', '64': '宁夏回族自治区', '65': '新疆维吾尔自治区',
            '83': '台湾地区', '81': '香港特别行政区', '82': '澳门特别行政区'
        }
        district_dict = self.make_regions_dict('districtNo', 'city_district')
        birth_date = []
        ages = []
        data['性别'] = ['女' if int(data['身份证号'].at[i][-2]) % 2 == 0 else '男' for i in data.index]
        region_num = [data['身份证号'].at[i][0] for i in data.index]
        province_num = [data['身份证号'].at[i][0:2] for i in data.index]
        district_num = [data['身份证号'].at[i][0:6] for i in data.index]
        df_year = [data['身份证号'].at[i][6:10] for i in data.index]
        df_month = [data['身份证号'].at[i][10:12] for i in data.index]
        df_day = [data['身份证号'].at[i][12:14] for i in data.index]
        data['所在区域'] = None
        for i in data.index:
            birth = datetime.date(int(df_year[i]), int(df_month[i]), int(df_day[i]))
            birth = birth.strftime('%Y年%m月%d日')
            age = datetime.datetime.today().year - int(df_year[i])
            ages.append(age)
            birth_date.append(birth)
        data['出生日期'] = birth_date
        data['年龄'] = ages
        region = [regions[int(num) - 1] for num in region_num]
        province = [province_dict.get(num, 'nodata') for num in province_num]
        district = [district_dict.get(num, 'nodata') for num in district_num]
        data['所在区域'] = region
        data['所属省份'] = province
        data['地区'] = district
        # print(data['地区'])
        if '序号' not in data.columns.tolist():
            data.insert(loc=0, column='序号', value=1)
        data['序号'] = [str(i + 1) for i in data.index]
        filename = ExamRoom.filename_by_time()
        data.to_excel('D:\\成绩统计结果\\' + filename, index=False)
        show_data = data.loc[:,
                    ['序号', '姓名', '身份证号', '性别', '出生日期', '年龄', '所属省份', '地区', '所在区域']]
        return filename, bad_df, id_bad, show_data
