import os
import traceback

import xlrd
import xlwt

from const import normal_data_row


def get_social_security():
    social_security_dict = {}
    read_social_security_excel = xlrd.open_workbook(r'../test.xlsx')
    social_security_sheet_names = read_social_security_excel.sheet_names()
    for sheet_name in social_security_sheet_names:
        sheet = read_social_security_excel.sheet_by_name(sheet_name)
        for row in sheet.get_rows():
            if 0 == len(str(row[0].value)):
                continue
            else:
                if isinstance(row[1].value, str) or 0 == len(str(row[1].value)):
                    row[1].value = float(0.0)
                social_security_dict[row[0].value] = row[1].value
    return social_security_dict


def read_excel():
    if os.path.exists('../奉贤coco2018年12月代发工资表_输出.xls'):
        os.remove('../奉贤coco2018年12月代发工资表_输出.xls')
    write_excel = xlwt.Workbook()
    read_excel = xlrd.open_workbook(r'../奉贤coco2018年12月代发工资表.xlsx')
    write_sheet = write_excel.add_sheet('奉贤coco2018年12月代发工资表')
    sheet_names = read_excel.sheet_names()
    social_security_dict = get_social_security()
    new_row_count = 0
    first_sheet = True
    # 遍历所有表
    for sheet_name in sheet_names:
        sheet = read_excel.sheet_by_name(sheet_name)
        for row in sheet.get_rows():  # 遍历所有行
            try:
                if not isinstance(row[0].value, float):
                    if first_sheet:
                        # 标题
                        if new_row_count == 0:
                            write_sheet.write(new_row_count, 0, str(row[0].value))
                            new_row_count += 1
                            continue
                        else:
                            # 列名
                            title_column_count = 0
                            for title_cell in row:
                                write_sheet.write(new_row_count, title_column_count, str(title_cell.value))
                                title_column_count += 1
                            new_row_count += 1
                            first_sheet = False
                            continue
                    else:
                        continue
                new_row_count = normal_data_row(row, write_sheet, new_row_count, social_security_dict)
            except Exception as e:
                traceback.print_exc(str(e) + str(row))
    write_excel.save('../奉贤coco2018年12月代发工资表_输出.xls')


if __name__ == '__main__':
    read_excel()
