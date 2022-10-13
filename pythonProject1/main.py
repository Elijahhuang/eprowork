# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

import requests
import json
import xlrd
import xlwt
import os
import xlsxwriter
import inputview
import requestUtil
import savedata


def analysisData(text):
    data = json.loads(text)
    return data


if __name__ == '__main__':
    text = requestUtil.request_token_Action('011534', 'e5ad9faea70154d97c60fe3572bf58dc', '94425', '6e159e66d156a48bca50a11b28b9b125')
    # text = "{\"access_token\":\".-.kp-YQb13rLlt5vOvy8r-Ow-qsyNbEk_B-jU3qYNIaJY\",\"token_type\":\"bearer\",\"refresh_token\":\".-.GCA-\",\"expires_in\":86399,\"scope\":\"all\",\"tenant_id\":\"epro\",\"biz_map\":{\"1\":\"1\",\"2\":\"1390512153638735873\",\"3\":\"1391651278186090498\",\"4\":\"1391658974918479874\",\"5\":\"1402873030710923265\"},\"user_name\":\"007943\",\"openId\":null,\"work_no\":\"007943\",\"phoneNum\":\"18795922952\",\"main_org_id\":\"82\",\"client_id\":\"Admin\",\"role_id\":\"2\",\"project_id\":\"1421324562124849153\",\"tenant_admin\":\"0\",\"is_authorize\":1,\"company_id\":\"42\",\"sex\":1,\"role_name\":\"\",\"license\":\"powered by epro\",\"odc_place_id\":\"0\",\"avatar_url\":\"\",\"user_id\":\"1421408474545078274\",\"org_id\":\"82\",\"job_id\":\"1390599290774097921\",\"nick_name\":\"黄广斌\",\"company_name\":\"南京分公司\",\"biz_id\":\"1402873030710923265\",\"is_original_password\":\"false\",\"account\":\"\",\"position_id\":\"\",\"jti\":\"44d61acb-1be9-4095-9bcc-0df46a61e5ea\"}"
    data = analysisData(text)
    access_token = ""
    if data:
        if 'error' in data:
            print(data['error_description'])
            exit(1)
        elif 'access_token' in data:
            access_token = data['access_token']
            savedata.save_access_token(access_token)
    get_access_token = savedata.get_access_token()
    print("get_access_token = " + get_access_token)
    if len(access_token) != 0:
        print(access_token)
    savedata.save_user_to_excel(data)


















