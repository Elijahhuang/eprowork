import xlwt
import os


def createUserInfo_sheet(f, table_name, try_count):
    if try_count <= 1:
        return None
    try:
        sheet = f.add_sheet(table_name + '-' + str(11 - try_count), cell_overwrite_ok=True)
        return sheet
    except Exception as e:
        return createUserInfo_sheet(f, table_name, try_count - 1)
    return None


def create_excel(excel_path, table_name, title_name, content_list):
    table1_invalid_start_x = 1
    table1_invalid_start_y = 2
    max_buf_len = []
    if os.path.exists(excel_path):
        os.remove(excel_path)
    f = xlwt.Workbook(encoding='utf-8')
    font = xlwt.Font()
    font.bold = True
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_TOP
    style1 = xlwt.XFStyle()
    style1.font = font
    style1.alignment = alignment
    style2 = xlwt.XFStyle()
    style2.alignment.wrap = 1

    try:
        sheet1 = f.add_sheet(table_name, cell_overwrite_ok=True)
    except Exception as e:
        sheet1 = createUserInfo_sheet(f, table_name, 10)
    for item in range(0, len(title_name)):
        sheet1.write(1, item+table1_invalid_start_x, title_name[item], style=style1)
        max_buf_len.append(len(title_name[item]))
        sheet1.col(1).width = 256 * (max_buf_len[item])

    item = 0
    column_len = len(title_name)
    for user_list in content_list:
        for column_index in range(0, column_len):
            title = title_name[column_index]
            value = user_list[title]
            if isinstance(value, str):
                if len(value) > max_buf_len[column_index]:
                    max_buf_len[column_index] = len(user_list[title_name[column_index]])
                if max_buf_len[column_index] > 150:
                    max_buf_len[column_index] = 150
            else:
                value = str(value)
            sheet1.col(column_index+table1_invalid_start_x).width = 256 * (max_buf_len[column_index] + 3)
            sheet1.write(table1_invalid_start_y + item, column_index+table1_invalid_start_x, value, style2)
        item += 1
    f.save(excel_path)


