# -*- coding: utf-8 -*-

"""
助教小工具

1. 上传成绩到网络学堂
2. 合并青橙oj的成绩单

Author: chenql
"""

import sys
import os
import xlwt, xlrd
import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time


'''
上传成绩到网络学堂
'''
def upload_excel_to_web_learning(excel_path, course_name, homework_name, col_id, col_score, col_comment, is_update):
    # excel 的格式
    # 用户名 xxx xxx xxx 平均分

    data = pd.read_excel(excel_path, sheet_name=0)  # header=None,
    data.fillna({col_score: 0, col_comment: ''}, inplace=True)  #空白分数填充0，空白评语填充''

    score_dict = pd.Series(data[col_score].values, index=data[col_id]).to_dict()
    comment_dict = pd.Series(data[col_comment].values, index=data[col_id]).to_dict()
    print(score_dict)
    print(comment_dict)
    upload_score_to_web_learning(score_dict, comment_dict, course_name, homework_name, is_update)


def upload_score_to_web_learning(score_dict, comment_dict, course_name, homework_name, is_update):
    # format
    # student_id : score
    driver = webdriver.Chrome()
    search_url = 'http://learn.tsinghua.edu.cn/index.jsp'
    driver.get(search_url)

    # 登录
    driver.find_element_by_xpath('/html/body/table[2]/tbody/tr/td[1]/table[2]/tbody/tr[1]/td/table/tbody/tr[1]/td/input').send_keys('chenql16')
    driver.find_element_by_xpath('/html/body/table[2]/tbody/tr/td[1]/table[2]/tbody/tr[1]/td/table/tbody/tr[2]/td/input[1]').send_keys('in333div)idually')
    driver.find_element_by_xpath('/html/body/table[2]/tbody/tr/td[1]/table[2]/tbody/tr[1]/td/table/tbody/tr[2]/td/input[2]').click()
    # 打开新标签页
    # js = "window.open('http://learn.tsinghua.edu.cn/MultiLanguage/lesson/teacher/course_locate.jsp?course_id=147693')"
    # driver.execute_script(js)
    # time.sleep(5)

    # 选择课程
    driver.switch_to.frame('content_frame')
    mycource = driver.find_element_by_xpath('//a[contains(text(), "%s")]' % course_name)
    mycource.click()

    # 选择作业
    driver.switch_to.window(driver.window_handles[-1])
    driver.find_element_by_xpath('//*[@id="left_menu"]/tbody/tr[6]/td/a').click()
    driver.switch_to.frame('content_frame')

    # 进入评阅
    myhomework = driver.find_element_by_xpath('//a[text()="%s"]/../../td[7]/table/tbody/tr/td[1]/input'
                                              % homework_name).click()
    # 如果已经登记过分数，此次为修改
    #
    if is_update:
        student_rows = driver.find_elements_by_xpath('//*[@id="table_box"]/tbody/tr/td[3]/a')
        for idx in range(1, len(student_rows)):
            row = driver.find_element_by_xpath('//*[@id="table_box"]/tbody/tr[%d]/td[3]/a' % idx)
            print(row.text)
            if not row.text.isdigit():
                continue
            student_id = eval(row.text)
            row.click()
            score = driver.find_element_by_xpath('//*[@id="post_rec_mark"]')
            comment = driver.find_element_by_xpath('//*[@id="post_rec_reply_detail"]')

            if student_id in score_dict.keys():
                print(student_id, score_dict[student_id], comment_dict[student_id])
                score.clear()
                score.send_keys(str(score_dict[student_id]))
                comment.clear()
                comment.send_keys('雷同题目: ' + comment_dict[student_id] if comment_dict[student_id] != ''
                                  else comment_dict[student_id])
            driver.find_element_by_xpath('//*[@id="F1"]/table[2]/tbody/tr/td/input[1]').click()
            time.sleep(0.2)

    # 没有登记过分数，登记一个同学的分数后学堂会自动跳转到下一个同学
    else:
        # 进入第一个同学的批改页面
        student = driver.find_element_by_xpath('//*[@id="table_box"]/tbody/tr[7]/td[3]/a')
        student.click()
        while True:
            student_id = driver.find_element_by_xpath('//*[@id="table_box"]/tbody/tr[1]/td[2]').text
            if not student_id:
                break
            # 清空输入框 ，输入分数，点击提交
            score = driver.find_element_by_xpath('//*[@id="post_rec_mark"]')
            comment = driver.find_element_by_xpath('//*[@id="post_rec_reply_detail"]')
            if eval(student_id) in score_dict.keys():
                score.clear()
                score.send_keys(str(score_dict[eval(student_id)]))
                comment.clear()
                comment.send_keys(comment_dict[eval(student_id)])
            else:
                score.send_keys('0')
            driver.find_element_by_xpath('//*[@id="F1"]/table[2]/tbody/tr/td/input[1]').click()

    print('ok')
    driver.quit()
    return




def main():
    # 上传成绩到网络学堂
    filepath = 'C:/Users/chenql/Desktop/cs/res.xlsx'
    upload_excel_to_web_learning(filepath, course_name="程序设计基础(2)(2017-2018秋季学期)",
                                 homework_name="大作业", col_id="用户名", col_score="大作业", col_comment='评语', is_update=0)

if __name__ == '__main__':
    main()
