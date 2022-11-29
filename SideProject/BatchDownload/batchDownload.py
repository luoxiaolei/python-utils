# 1. 登录获取token。
# 2. 通过遍历日期拼接要下载的文件请求地址。
# 3. 访问下载地址的地址，下载文件。


import requests
import time
from datetime import datetime, date, timedelta


def days_cur_month():
    m = datetime.now().month
    y = datetime.now().year
    ndays = (date(y, m+1, 1) - date(y, m, 1)).days
    d1 = date(y, m, 1)
    d2 = date(y, m, ndays)
    delta = d2 - d1

    return [(d1 + timedelta(days=i)).strftime('%Y-%m-%d') for i in range(delta.days + 1)]

url = "http://127.0.0.1/tenant/api/v1/user/login"
data = '{"passwd": "password", "email": "admin"}'
response = requests.post(url, data=data,
                    headers={'Content-Type':'application/json'})
cookie_raw = requests.utils.dict_from_cookiejar(response.cookies)
cookie = "token="+cookie_raw['token']
print(cookie)

headers = {'Cookie':cookie}

days = days_cur_month()
pre = 'http://127.0.0.1/videomon/api/v1/app/report/check/online/detail/export?arealayerno=2202&statTime='
end = '&dynamicCondition=%7B%22videoType%22%3A%7B%22type%22%3A%22MULTISELECT%22%2C%22value%22%3A%5B%22yd%22%5D%7D%7D&queryType=online'

toDay = time.strftime("%Y-%m-%d", time.localtime())

for day in days:
    if day == toDay:
        break 
    currUrl = pre+day+end
    r = requests.get(currUrl,headers=headers)
    with open('/Users/luoxiaolei/Desktop/test/'+day+'.xls','wb') as f:
        f.write(r.content)
        f.flush()


