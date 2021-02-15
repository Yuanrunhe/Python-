# -*- coding: utf-8 -*-
# Author : YRH
# Data : 
# Project : 
# Tool : PyCharm

import requests
from fake_useragent import UserAgent
import random
import json
import xlwt
import time
import threading
import time

ua = UserAgent()


def join_url(year):
    da = []
    # 项目类型代码
    projectType = ["218", "220", "339", "579", "630", "631", "649"]
    # 提取申请代码
    response = requests.get("http://kd.nsfc.gov.cn/common/data/fieldCode")
    response.encoding = response.apparent_encoding
    data = eval(response.text.replace("\xa0", "").replace("\u200d", ""))
    data = data["data"]
    for y in year:
        for p in projectType:
            for code in data:
                name = code["name"]
                code = code["code"]
                if len(code) <= 1:
                    continue
                da.append({"year": str(y), "name": name, "code": code, "project": p})

    return da


def spider(payload, data_list):
    try:
        url = "http://kd.nsfc.gov.cn/baseQuery/data/supportQueryResultsData"
        headers = {
            "User-Agent": ua.random,
            "Content-Type": "application/json"
        }
        data = {
            "code": payload["code"], "conclusionYear": payload["year"], "projectType": payload["project"],
            "queryType": "input",
            "complete": "true", "pageNum": 0,
        }
        response = requests.post(url, data=json.dumps(data), headers=headers)
        response.encoding = response.apparent_encoding
        text = response.text
        text = text.replace("\ue06d", "").replace("\u2022", "").replace("\xf6", "")
        da = eval(text)
        jiexi(da, data_list)
    except Exception as E:
        print(E)


def jiexi(resultsData, data_list):
    try:
        result = resultsData["data"]["resultsData"]
        for data in result:
            projectName = data[1]
            approvalN = data[2]
            projectType = data[3]
            projectLeader = data[5]
            relyingUnit = data[4]
            year = data[7]
            keyword = data[8]
            data_dict = {"项目名称": projectName,
                         "项目批准号": approvalN,
                         "项目类型": projectType,
                         "项目负责人": projectLeader,
                         "依托单位": relyingUnit,
                         "批准年度": year,
                         "关键词": keyword}
            print(data_dict)
            data_list.append(data_dict)
    except Exception as E:
        print(E)


def run(payload, data_list):
    print(payload)
    # time.sleep(0.1)


if __name__ == '__main__':

    all_data = []
    # 爬取的项目年份
    year = [2019]
    data = join_url(year)

    # 单任务执行
    # for d in data:
    #     try:
    #         spider(d, all_data)
    #     except Exception as e:
    #         print(e)

    # 多任务执行
    for d in data:
        try:
            spider_thread = threading.Thread(target=spider, kwargs={"payload": d, "data_list": all_data})
            spider_thread.start()
        except Exception as E:
            print(E)
        spider_thread.join()

    # 数据保存
    workbook = xlwt.Workbook(encoding="utf-8")
    worksheet = workbook.add_sheet("sheet1")
    head = ["项目名称", "项目类型", "批准年度", "项目批准号", "项目负责人", "依托单位", "关键词"]
    for i in range(7):
        worksheet.write(0, i, head[i])
    row = 1
    for i in all_data:
        for j in range(0, len(head)):
            worksheet.write(row, j, i[head[j]])
        row += 1
    workbook.save("结题项目数据.xls")
