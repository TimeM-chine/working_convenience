# -*- coding: utf-8 -*-
import requests
import json
import time
import pprint
import openpyxl
from datetime import datetime

with open("id.txt", encoding="GBK") as f:
    s = f.read()
ID_list = s.split('\n')
for i in range(len(ID_list)):
    ID_list[i] = int(ID_list[i])

# 准备部分
# cookie = '_ga=GA1.2.2057943407.1595661913; SL_C_23361dd035530_KEY=be556a167e74fcde3a3444e29b25f8e99fb0c59f; __guid=111581274.4233868734248029700.1610877138168.5208; admin-authorization=Bearer+eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJ1c2VyX3R5cGUiOiJhZG1pbiIsInVzZXJfaWQiOjQyODgsImlhdCI6MTYxMzU4Nzg2NSwianRpIjoiMjRkOGZhMjctNzE5NC0xMWViLTgwYmMtYzNhMjZkMDBkMGUwIn0.7IAjLJ06RdmTuqfyFL7GiCFkaOSxI9ZbAUY5r0_VYD0; internal_account_token=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJUb2tlbiIsImF1dGgiOiJST0xFX0FETUlOIiwibmFtZSI6IumCueaxleaWjCIsImVuaWQiOjQ4NDgsImlhdCI6MTYxMzYxNjY2NSwianRpIjoiMTExNGVmN2UtNmM5My00MjEwLWFhM2YtOTZjNWEzZTljZTJmIn0.SSKRITtr0HSPeYPXOBNqIx_TtUQ8QgOMYPBSjIGJx7c; authorization=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJDb2RlbWFvIEF1dGgiLCJ1c2VyX3R5cGUiOiJzdHVkZW50IiwidXNlcl9pZCI6NDc1MDEzMiwiaXNzIjoiQXV0aCBTZXJ2aWNlIiwicGlkIjoiNjVlZENUeWciLCJleHAiOjE2MTc2ODI4MTAsImlhdCI6MTYxMzc5NDgxMCwianRpIjoiNTBkMmZkM2QtNTc1Mi00YzcwLTg1ZGUtMTQzYWRhMDBlZGRkIn0.Jw_cyKu7_c-8kdocFw1hTZpq3o0URPrE0VsgLDzvyPk; __ca_uid_key__=6908428c-cf8a-450c-a8bc-77fbd0bd2e36; refresh-token=MTo0NzUwMTMyOndlYjpBQUFCZUVMbVRCSjdJdVJBS3pwQzRCLWdBVEpSNE9BcTo4YzRiMzhjMC04N2ZhLTQ0ZGUtYTg1MC1iMWZiNjFkMTNlMjA=; _gid=GA1.2.1847212740.1616202350; SERVERID=7eadfc7a9d5ed6727c515ba9042221d8|1616412828|1616400554'
cookie = '_ga=GA1.2.2057943407.1595661913; SL_C_23361dd035530_KEY=be556a167e74fcde3a3444e29b25f8e99fb0c59f; __guid=111581274.4233868734248029700.1610877138168.5208; authorization=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJDb2RlbWFvIEF1dGgiLCJ1c2VyX3R5cGUiOiJzdHVkZW50IiwidXNlcl9pZCI6NDc1MDEzMiwiaXNzIjoiQXV0aCBTZXJ2aWNlIiwicGlkIjoiNjVlZENUeWciLCJleHAiOjE2MTc2ODI4MTAsImlhdCI6MTYxMzc5NDgxMCwianRpIjoiNTBkMmZkM2QtNTc1Mi00YzcwLTg1ZGUtMTQzYWRhMDBlZGRkIn0.Jw_cyKu7_c-8kdocFw1hTZpq3o0URPrE0VsgLDzvyPk; refresh-token=MTo0NzUwMTMyOndlYjpBQUFCZUVMbVRCSjdJdVJBS3pwQzRCLWdBVEpSNE9BcTo4YzRiMzhjMC04N2ZhLTQ0ZGUtYTg1MC1iMWZiNjFkMTNlMjA=; __ca_uid_key__=6c91b293-34d7-447b-8bdd-9c3358f6413c; admin-authorization=Bearer+eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJ1c2VyX3R5cGUiOiJhZG1pbiIsInVzZXJfaWQiOjQyODgsImlhdCI6MTYxNjk3NjI1MCwianRpIjoiNTc5YTA0MDItOTA2NS0xMWViLTkxNTctYTkzNjFkNzU0MWY5In0.2-isjD2N0Vss11SNG01rfE_JEvqSTUguJ4y4ltSEjLE; internal_account_token=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJUb2tlbiIsImF1dGgiOiJST0xFX0FETUlOIiwibmFtZSI6IumCueaxleaWjCIsImVuaWQiOjQ4NDgsImlhdCI6MTYxNzAwNTA1MCwianRpIjoiMjNhNmE1NmEtN2E5Ny00NmQzLWEwNTYtMTgwNmFjYmVkZDdhIn0.ddsR2Ts2vc65t3ZAuU2f295QlFtz0kOLBM7PB9PuT4o; SERVERID=7eadfc7a9d5ed6727c515ba9042221d8|1617102898|1617102894'
headers = {
    "Accept": "application/json, text/plain, */*",
    "authorization_type": "3",
    "cookie": cookie,
    "content-type": "application/json;charset=UTF-8",
    "Origin": "https://crm.codemao.cn",
    "Referer": "https://crm.codemao.cn/",
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.86 Safari/537.36',
}

def num2time(num):
    lt = time.localtime(num)
    last_time = time.strftime("%Y-%m-%d %H:%M", lt)
    return last_time


def get_name(x):
    user_id = x
    tik = get_tikets(user_id)
    url = f"https://cloud-gateway.codemao.cn/api-crm-web/admin/users/{user_id}"
    r = requests.get(url, headers=headers)
    # print("标签查询状态码:", r.status_code)
    r = r.json()
    # pprint.pprint(r)
    name = r['full_name']
    gap = r['days_not_attend_class']
    return name,  gap, tik


def get_order_time(user_id):
    user_id = user_id
    url = f"https://cloud-gateway.codemao.cn/api-crm-web/admin/users/{user_id}/attendances?page=1"
    r = requests.get(url, headers=headers)
    order_time = []
    teacher = []
    for i in r.json()['attendance_details']:
        if i['attendanceState'] == "BEFORE_CLASS" and i['cancelled']==0:
            day = num2time(i['time_slot']) # 格式为 2021-04-10 19:30
            week = datetime.strptime(day, "%Y-%m-%d %H:%M").weekday() + 1  #结果为数字
            _ = day.split(" ")
            __ = (' 周' + str(week) + ' ').join(_)
            order_time.append(__)
            if len(i['teacher_username']) != 0:
                teacher.append(i['teacher_username'])
            else:
                teacher.append('未指定')
    order_time = '\n'.join(order_time)
    teacher = '\n'.join(teacher)
    if len(order_time)==0:
        order_time, teacher = ['未约课','未约课']
    return order_time, teacher


def get_tikets(x):
    url = r'https://cloud-gateway.codemao.cn/api-crm-web/admin/users/search'
    data ={
        'page': 1,
        'user_id': str(x)
    }
    r = requests.post(url, headers=headers, data=json.dumps(data))
    return r.json()['items'][0]['remainTickets']


# print(get_order_time(329920))
wb = openpyxl.Workbook()
sheet = wb.active
first_list = ['id', '学生姓名', 'N天未上课','剩余课时券','最近约课时间','约课老师']
sheet.append(first_list)
count = 0

for i in range(len(ID_list)):
    data = []
    info_1 = list(get_name(ID_list[i]))
    info_2 = list(get_order_time(ID_list[i]))
    data.append(ID_list[i])
    data = data + info_1 + info_2
    sheet.append(data)
    count += 1
    print(count)

wb.save('约课状态'+str(time.time())+'.xlsx')
