import requests
import json

url = 'https://oas.epro.com.cn/api/blade-auth/oauth/token'

header = {
    "Accept": "application/json, text/plain, */*",
    "Authorization": "Basic QWRtaW46QWRtaW5fUENfMjAyMGVQUk8=",
    "Accept-Language": "zh-CN,zh-Hans;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Host": "oas.epro.com.cn",
    "Origin": "https://oas.epro.com.cn",
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/15.5 Safari/605.1.15",
    "Referer": "https://oas.epro.com.cn/",
    "Content-Length": "0",
    "Connection": "keep-alive",
    "appId": "Admin",
    "Tenant-Id": "epro",
    "Captcha-Code": "53832",  # 验证码
    "Captcha-Key": "8876ba0c903783219649e8c488e41a2c"
}

param = {
    "tenantId": "epro",
    "account": "",
    "username": "",  # 账号
    "password": "",  # 密码(加密后的密文)
    "grant_type": "captcha",
    "scope": "all",
    "type": "account",
    "user_type": "staff"
}


def request_token_Action(user, password, code, key):
    param["username"] = user
    param["password"] = password
    header["Captcha-Code"] = code
    header["Captcha-Key"] = key
    reps = requests.request("get", url, params=param, headers=header)
    print(reps.text)
    return reps.text

