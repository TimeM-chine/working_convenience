import requests
import json
import pprint
import time

with open("id.txt", encoding="GBK") as f:
    s = f.read()
ID_list = s.split('\n')
for i in range(len(ID_list)):
    ID_list[i] = int(ID_list[i])

# 准备部分
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
a =[]
data ={
    'page': 1,
    'limit': 50
}
url = f"https://cloud-gateway.codemao.cn/api-crm-web/admin/crm/customers/all"
r = requests.post(url, headers=headers,data=json.dumps(data))
# print("标签查询状态码:", r.status_code)
r = r.json()['items']
for i in r :
    a.append(i['user_id'])
data = {
    'page': 2,
    'limit': 50
}
r = requests.post(url, headers=headers, data=json.dumps(data))
r = r.json()['items']
for i in range(len(r)):
    a.append(r[i]['user_id'])
print('crm系统中学生数为',len(a))
print('id.txt中学生数为', len(ID_list))
file_name = "学生对比"+str(time.time())+".txt"
for i in ID_list:
    if i not in a:
        print(i,'未划入crm系统')
        with open(file_name, 'a') as f:
            f.write(str(i)+"  "+"未划入crm系统\n")


for i in a:
    if i not in ID_list:
        print(i, '在crm系统中多余')
        with open(file_name, 'a') as f:
            f.write(str(i)+"  "+"在crm系统中多余\n")
