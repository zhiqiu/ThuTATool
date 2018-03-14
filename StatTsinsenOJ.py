# -*- coding: utf-8 -*-

"""
合并清橙oj上的成绩统计结果
使用pandas处理excel

Author: chenql
"""

import sys
import os
import xlwt, xlrd
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.action_chains import *
import time

# TODO 自动下载清澄上的结果
def download_excel():

    driver = webdriver.Chrome()
    search_url = 'http://oj.tsinsen.com/'
    driver.get(search_url)

    # 登录
    driver.find_element_by_css_selector('#username').send_keys('chenql')
    driver.find_element_by_css_selector('#password').send_keys('chenqiuliang')
    driver.find_element_by_css_selector('#cuinfo > input:nth-child(3)').click()

    # 切换课程
    chain = ActionChains(driver)
    menu = driver.find_element_by_xpath('//*[@id="mainm"]/ul')
    chain.move_to_element(menu).perform()
    submenu = menu.find_element_by_xpath('./li[4]')
    chain.move_to_element(submenu).perform()


    time.sleep(10)
    lessons = driver.find_elements_by_xpath('//*[@id="domainList"]/div/table/tbody/tr')
    for lesson in lessons:
        lesson.find_element_by_xpath('./td[2]/div/a')
        lesson.click()


    time.sleep(5)

    # 元素不可点击
    # Element is not clickable at point
    # button = driver.find_element_by_css_selector("#holdings-cnt > ul > li.trade-history ")
    # button.click()

    # 用js实现点击
    #driver.execute_script('var q=document.getElementsByClassName("trade-history");q[0].click();')

    print(driver.page_source)
    # driver.quit()


def stat_score(score_path, same_path, threshold=0.8, except_name=['']):
    # 处理查重文件
    files = os.listdir(same_path)
    res_same = pd.DataFrame()

    d = dict()  # 每一个人雷同的题目
    for idx, file in enumerate(files):
        filepath = same_path + '/' + file
        if idx < 2:
            continue
        if os.path.isfile(filepath):
            # 读入查重excel表格
            # 用户名	姓名	 雷同数	题目1 题目2 ...
            # xxx   xxx  xx     xx    xx
            #
            # 这里的雷同数有重复，所以应该用后面的小题重复信息去重来计算
            dict_same = pd.read_excel(filepath, sheet_name=None, skiprows=range(4))  # header=None,
            #data = pd.read_excel(filepath, sheet_name=0, skiprows=range(5))  # header=None,['雷同统计']

            for key in dict_same:
                if key != '雷同统计':
                    df = dict_same[key][dict_same[key]['相似程度'] > threshold]     # 按相似度过滤
                    df = df[~df['用户1'].str.contains('|'.join(except_name))]       # 过滤助教和小教员的查重
                    df = df[~df['用户2'].str.contains('|'.join(except_name))]       # 过滤助教和小教员的查重

                    students = df['用户1'].values.tolist() + df['用户2'].values.tolist()
                    students = list(set(students))
                    for student in students:
                        sid = student.split('(')[0]
                        if sid not in d.keys():
                            d[sid] = [key]
                        else:
                            d[sid] += [key]
        # d = {2017011xxx : [xx, xx], ....}  #每个同学重复的题目

    # 处理成绩文件
    files = os.listdir(score_path)
    res_score = pd.DataFrame()
    for idx, file in enumerate(files):
        filepath = score_path + '/' + file
        if os.path.isfile(filepath):
            # 读入excel表格
            # 用户名	姓名	总分	题目1 题目2 ...
            # xxx   xxx  xx xx    xx
            data = pd.read_excel(filepath, sheet_name=0, skiprows=range(4))  # header=None,

            # 先减去雷同题目的得分
            for key in d:
                for pro_name in d[key]:
                    if pro_name in list(data.columns):

                        row_index = list(data['用户名']).index(key)
                        col_index = data.columns.get_loc(pro_name)
                        #print(row_index, col_index)
                        data.iat[row_index, col_index] = 0
            #print(data)
            # 平均分 = 总分(第3列到最后一列相加)/ 题目数(列数-3)
            # 不能使用第二列总分，因为里面有查重扣分
            data[str(idx + 1)] = data.iloc[:, 3:].apply(lambda x: x.sum(), axis=1) / data.iloc[:, 3:].shape[1]
            # print(data)

            if res_score.empty:
                res_score = data.iloc[:, [0, 1, -1]]
            else:
                res_score = pd.merge(res_score, data.iloc[:, [0, -1]], how='outer', on='用户名')

        #     # 第二列就是重复次数，但是有重复
        #     data[str(idx + 1)] = data.iloc[:, 2]
        #     data[str(idx + 1)] = data.iloc[:, 2]
        #     #data[str(idx + 1) + '_sum'] = data.iloc[:, 2]
        #     #data[str(idx + 1) + '_title'] = data.iloc[:, 2]
        #     # print(data)
        #
        #     if res_same.empty:
        #         res_same = data.iloc[:, [0, 1, -1]]
        #     else:
        #         res_same = pd.merge(res_same, data.iloc[:, [0, -1]], on='用户名')
        # #print(res_same)

    # 总雷同次数从第三次开始统计
    #print(res_score.iloc[:, 2:].shape[1])
    res_score['平均分'] = res_score.iloc[:, 2:].apply(lambda x: x.sum(), axis=1) / res_score.iloc[:, 2:].shape[1]
    res_score['总雷同次数'] = 0
    res_score['雷同题目'] = ''

    for key in d:
        index = list(res_score['用户名']).index(key)
        #print(index)
        res_score.iat[index, -1] = ','.join(d[key])    # 雷同题目
        res_score.iat[index, -2] = len(d[key])         # 总雷同次数
    #print(res_score)

    #print(res)
    writer = pd.ExcelWriter(score_path + '/res.xlsx')
    res_score.to_excel(writer, 'Sheet1', index=False)
    writer.save()


def main():
    # 分数文件路径
    score_path = 'C:/Users/chenql/Desktop/cs'
    # 查重文件路径
    same_path = 'C:/Users/chenql/Desktop/cs/same'
    except_name = ['chenql', 'wangyuping', '2017011349', '2017011326', '2017011325', '2017011346',
                   '2017011484', '2017011366', '2017011369', '2017011387', '2017011362', '2017011348']
    stat_score(score_path, same_path, threshold=0.9, except_name=except_name)

if __name__ == '__main__':
    main()
