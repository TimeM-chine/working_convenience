# -*- coding: utf-8 -*-
import requests
import json
import time
import pprint
import openpyxl
from datetime import datetime

with open("test.txt", encoding="GBK") as f:
    s = f.read()
ID_list = s.split('\n')
for i in range(len(ID_list)):
    ID_list[i] = int(ID_list[i])

# ćć€éšć
# cookie = '_ga=GA1.2.2057943407.1595661913; SL_C_23361dd035530_KEY=be556a167e74fcde3a3444e29b25f8e99fb0c59f; __guid=111581274.4233868734248029700.1610877138168.5208; admin-authorization=Bearer+eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJ1c2VyX3R5cGUiOiJhZG1pbiIsInVzZXJfaWQiOjQyODgsImlhdCI6MTYxMzU4Nzg2NSwianRpIjoiMjRkOGZhMjctNzE5NC0xMWViLTgwYmMtYzNhMjZkMDBkMGUwIn0.7IAjLJ06RdmTuqfyFL7GiCFkaOSxI9ZbAUY5r0_VYD0; internal_account_token=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJUb2tlbiIsImF1dGgiOiJST0xFX0FETUlOIiwibmFtZSI6IumCueaxleaWjCIsImVuaWQiOjQ4NDgsImlhdCI6MTYxMzYxNjY2NSwianRpIjoiMTExNGVmN2UtNmM5My00MjEwLWFhM2YtOTZjNWEzZTljZTJmIn0.SSKRITtr0HSPeYPXOBNqIx_TtUQ8QgOMYPBSjIGJx7c; authorization=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJDb2RlbWFvIEF1dGgiLCJ1c2VyX3R5cGUiOiJzdHVkZW50IiwidXNlcl9pZCI6NDc1MDEzMiwiaXNzIjoiQXV0aCBTZXJ2aWNlIiwicGlkIjoiNjVlZENUeWciLCJleHAiOjE2MTc2ODI4MTAsImlhdCI6MTYxMzc5NDgxMCwianRpIjoiNTBkMmZkM2QtNTc1Mi00YzcwLTg1ZGUtMTQzYWRhMDBlZGRkIn0.Jw_cyKu7_c-8kdocFw1hTZpq3o0URPrE0VsgLDzvyPk; __ca_uid_key__=6908428c-cf8a-450c-a8bc-77fbd0bd2e36; refresh-token=MTo0NzUwMTMyOndlYjpBQUFCZUVMbVRCSjdJdVJBS3pwQzRCLWdBVEpSNE9BcTo4YzRiMzhjMC04N2ZhLTQ0ZGUtYTg1MC1iMWZiNjFkMTNlMjA=; _gid=GA1.2.1847212740.1616202350; SERVERID=7eadfc7a9d5ed6727c515ba9042221d8|1616412828|1616400554'
cookie = '_ga=GA1.2.2057943407.1595661913; SL_C_23361dd035530_KEY=be556a167e74fcde3a3444e29b25f8e99fb0c59f; __guid=111581274.4233868734248029700.1610877138168.5208; authorization=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJDb2RlbWFvIEF1dGgiLCJ1c2VyX3R5cGUiOiJzdHVkZW50IiwidXNlcl9pZCI6NDc1MDEzMiwiaXNzIjoiQXV0aCBTZXJ2aWNlIiwicGlkIjoiNjVlZENUeWciLCJleHAiOjE2MTc2ODI4MTAsImlhdCI6MTYxMzc5NDgxMCwianRpIjoiNTBkMmZkM2QtNTc1Mi00YzcwLTg1ZGUtMTQzYWRhMDBlZGRkIn0.Jw_cyKu7_c-8kdocFw1hTZpq3o0URPrE0VsgLDzvyPk; refresh-token=MTo0NzUwMTMyOndlYjpBQUFCZUVMbVRCSjdJdVJBS3pwQzRCLWdBVEpSNE9BcTo4YzRiMzhjMC04N2ZhLTQ0ZGUtYTg1MC1iMWZiNjFkMTNlMjA=; __ca_uid_key__=6c91b293-34d7-447b-8bdd-9c3358f6413c; admin-authorization=Bearer+eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJ1c2VyX3R5cGUiOiJhZG1pbiIsInVzZXJfaWQiOjQyODgsImlhdCI6MTYxNjk3NjI1MCwianRpIjoiNTc5YTA0MDItOTA2NS0xMWViLTkxNTctYTkzNjFkNzU0MWY5In0.2-isjD2N0Vss11SNG01rfE_JEvqSTUguJ4y4ltSEjLE; internal_account_token=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJUb2tlbiIsImF1dGgiOiJST0xFX0FETUlOIiwibmFtZSI6IumCueaxleaWjCIsImVuaWQiOjQ4NDgsImlhdCI6MTYxNzAwNTA1MCwianRpIjoiMjNhNmE1NmEtN2E5Ny00NmQzLWEwNTYtMTgwNmFjYmVkZDdhIn0.ddsR2Ts2vc65t3ZAuU2f295QlFtz0kOLBM7PB9PuT4o; SERVERID=2f59be74b3ab04ebbb5eb794875b917c|1617861123|1617859713'
headers = {
    "Accept": "application/json, text/plain, */*",
    "authorization_type": "3",
    "cookie": cookie,
    "content-type": "application/json;charset=UTF-8",
    "Origin": "https://crm.codemao.cn",
    "Referer": "https://crm.codemao.cn/",
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.86 Safari/537.36',
}
want_dic = {
    '0': 'æȘćĄ«',
    "101": "éèżć­Šäč äșè§ŁçŒçšç„èŻ",
    "102": "ćć ç«è”ć ć",
    "103": "ćŽè¶Łç±ć„œćčć»",
    "104": "ćčć»ć­©ć­çé»èŸæç»Ž",
    "105": "é»çŒć­©ć­çèĄšèŸŸæČéèœć",
    "106": "ćčć»ć­©ć­ć„œçć­Šäč äč æŻ",
    "107": "ćȘæŻćŹäžćŹèŻèŻç",
    "108": "æé€æžžæ",
    "109": "èŻæ±çčäžæžæ°",
    "99": "ć¶ä»",
    "110": "ćčć»ćæéźéąăè§Łćłéźéąçèœć",
    "111": "äșè§ŁäžäžïŒéĄșćșæ¶ä»ŁéèŠ"
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
    # print("æ ç­Ÿæ„èŻąç¶æç :", r.status_code)
    r = r.json()
    # pprint.pprint(r)
    name = r['full_name']
    if r['applicant'] == 1:
        parent = 'çžçž'
    elif r['applicant'] == 2:
        parent = 'ćŠćŠ'
    else:
        parent = 'æȘç„'
    code = r['applicant_phone_number'][7:]
    # last_time = r['last_attend_class_date']
    # lt = time.localtime(last_time)
    # last_time = time.strftime("%Y-%m-%d %H:%M", lt)
    # number = get_phone(code)
    # head_teacer = r['head_teacher_name']
    info = r['province_name'] + r['city_name'] + str(r['age']) + 'ćČ'
    # par_want = ''
    # for i in want_dic:
    #     if str(r['skill']['study_purpose']) == i:
    #         par_want = want_dic[i]
    us = r['user_tags']
    user_tags = []
    for i in range(len(us)):
        user_tags.append(us[i]['name'])
    user_tags = '\n'.join(user_tags)
    # gap = r['days_not_attend_class']
    # return name, info, parent, number, par_want, last_time, gap, user_tags, tik
    return name, info, parent, user_tags, tik

def get_phone(code):
    # url = r"https://open-service.codemao.cn/decode/phone_number"
    # data = {
    #     'cipher_text': code
    # }
    # r = requests.post(url, headers=headers, data=json.dumps(data))
    return "fail"


def get_tikets(x):
    url = r'https://cloud-gateway.codemao.cn/api-crm-web/admin/users/search'
    data ={
        'page': 1,
        'user_id': str(x)
    }
    r = requests.post(url, headers=headers, data=json.dumps(data))
    return r.json()['items'][0]['remainTickets']

def get_class_info(x):
    user_id = x
    url = f"https://cloud-gateway.codemao.cn/api-crm-web/admin/crm/users/{user_id}/lesson/records?page=1"
    r = requests.get(url, headers=headers).json()
    d = r['work_record_details']
    class_name = d[0]['point_description']
    begin_time = {}
    count = 0
    comments = []
    for i in range(len(d)):
        begin_time[str(d[i]['attendance_id'])] = d[i]['begin_time']
        if len(d[i]['comments']) > 0 and len(d[i]['comments'][0]['content']) > 6:
            comments.append(d[i]['comments'][0]['content'])
        if d[i]['tob_status'] == 'ć·Čćźæ':
            count += 1
    begin_time = list(begin_time.values())
    t = []
    for i in range(len(begin_time)):
        begin_time[i] = num2time(begin_time[i])
        week = datetime.strptime(begin_time[i], "%Y-%m-%d %H:%M").weekday() + 1
        t.append(week)
    final = []
    for i in range(len(begin_time)):
        ll = begin_time[i].split(' ')
        kk = (' ćš' + str(t[i]) + ' ').join(ll)
        final.append(kk)
    count = count *10
    count = str(count)+'%'
    final = '\n'.join(final)
    comments = '\n'.join(comments)
    return count, final, comments,class_name


# print(get_tikets(3481630))
wb = openpyxl.Workbook()
sheet = wb.active
# first_list = ['id', 'ć­Šçć§ć', 'ćșæŹäżĄæŻ', 'èŽèŽŁäșș',  'äžæŹĄäžèŻŸæ¶éŽ','Nć€©æȘäžèŻŸ',
#               'æ ç­Ÿ', 'ć©äœèŻŸæ¶ćž', 'éŻćłćźæç', 'æèżäžèŻŸèź°ćœ', 'æèżææćéŠ', 'èŻŸçšèżćșŠ']
first_list = ['id', 'ć­Šçć§ć', 'ćșæŹäżĄæŻ', 'èŽèŽŁäșș',  'æ ç­Ÿ', 'ć©äœèŻŸæ¶ćž', 'éŻćłćźæç', 'æèżäžèŻŸèź°ćœ', 'æèżææćéŠ', 'èŻŸçšèżćșŠ']
sheet.append(first_list)
count = 0

for i in range(len(ID_list)):
    data = []
    info_1 = list(get_name(ID_list[i]))
    # info_1 = [str(k) for k in info_1]
    info_2 = list(get_class_info(ID_list[i]))
    # info_2 = [str(p) for p in info_2]
    data.append(ID_list[i])
    data = data + info_1 + info_2
    sheet.append(data)
    count += 1
    print(count)
wb.save('ć­Šçć­Šæ'+str(time.time())+'.xlsx')
