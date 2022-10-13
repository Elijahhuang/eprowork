import xlsxwriter
import os


# "biz_map"
biz_map_title = [
    "1",
    "2",
    "3",
    "4",
    "5"
]


user_info_title = [
    # "access_token",
    "token_type",
    # "refresh_token",
    "expires_in",
    "scope",
    "tenant_id",
    "user_name",
    "openId",
    "work_no",
    "phoneNum",
    "main_org_id",
    "client_id",
    "role_id",
    "project_id",
    "tenant_admin",
    "is_authorize",
    "company_id",
    "sex",
    "role_name",
    "license",
    "odc_place_id",
    "avatar_url",
    "user_id",
    "org_id",
    "job_id",
    "nick_name",
    "company_name",
    "biz_id",
    "is_original_password",
    "account",
    "position_id",
    "jti"
]


def save_user_to_excel(data):
    create_info_excel(str(get_work_path() + "/info.xlsx"), "table1", user_info_title, [data])


def save_access_token(access_token_text):
    file = open(get_access_token_path(), "w")
    file.write(access_token_text)


def get_access_token():
    file = open(get_access_token_path())
    return file.read()


def get_access_token_path():
    return str(get_work_path() + "/access_token.txt")


def get_work_path():
    path = str(os.path.join(os.path.expanduser('~'), "Desktop") + "/work")
    if not os.path.exists(path):
        os.makekdirs(path)
    return path


def create_info_excel(excel_path, table_name, title_name, content_list):
    if os.path.exists(excel_path):
        os.remove(excel_path)
    workbook = xlsxwriter.Workbook(excel_path)
    worksheet = workbook.add_worksheet(table_name)
    for i in range(0, len(title_name)):
        worksheet.write(0, i, title_name[i])
    for j in range(0, len(content_list)):
        obj = content_list[j]
        for k in range(0, len(title_name)):
            key = title_name[k]
            value = obj[key]
            worksheet.write(j+1, k, value)
    workbook.close()

