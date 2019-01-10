service_fee = 75
ticket_point = 0.06


def get_two_float(float_num, n=2):
    f_str = str(float_num)  # f_str = '{}'.format(f_str) 也可以转换为字符串
    a, b, c = f_str.partition('.')
    c = (c + "0" * n)[:n]  # 如论传入的函数有几位小数，在字符串后面都添加n为小数0
    return float(".".join([a, c]))


def calculate_tax(money):
    money = int(money)
    if money < 5000:
        return 0.0
    elif money <= 8000:
        tax = (money - 5000) * 3 / 100
    elif money <= 17000:
        tax = 3000 * 3 / 100 + (money - 5000 - 3000) * 0.1
    elif money <= 30000:
        tax = 3000 * 3 / 100 + 9000 * 0.1 + (money - 12000 - 5000) * 0.2
    elif money <= 40000:
        tax = 3000 * 3 / 100 + 9000 * 0.1 + 13000 * 0.2 + (money - 25000 - 5000) * 0.25
    elif money <= 60000:
        tax = 3000 * 3 / 100 + 9000 * 0.1 + 13000 * 0.2 + 10000 * 0.25 + (money - 35000 - 5000) * 0.3
    elif money <= 85000:
        tax = 3000 * 3 / 100 + 9000 * 0.1 + 13000 * 0.2 + 10000 * 0.25 + 20000 * 0.3 + (money - 55000 - 5000) * 0.35
    else:
        tax = 3000 * 3 / 100 + 9000 * 0.1 + 13000 * 0.2 + 10000 * 0.25 + 20000 * 0.3 + 25000 * 0.35 + (money - 85000) * 0.45
    return get_two_float(tax)


def calculate_ticket_point(should_pay, tax):
    return get_two_float((should_pay - tax + service_fee) * ticket_point)


def calculate_actual_pay(should_pay, social_security, tax):
    return get_two_float(should_pay - social_security - tax)


def calculate_transfer_money(actual_pay, ticket_point):
    return get_two_float(actual_pay + service_fee + ticket_point)


def normal_data_row(row, write_sheet, new_row_count, social_security_dict):
    row[0].value = int(row[0].value)
    row[9].value = social_security_dict[row[4].value]
    row[10].value = calculate_tax(row[8].value)
    row[12].value = service_fee
    row[13].value = calculate_ticket_point(row[8].value, row[10].value)
    row[11].value = calculate_actual_pay(row[8].value, row[9].value, row[10].value)
    row[14].value = calculate_transfer_money(row[11].value, row[13].value)
    new_column_count = 0
    for cell in row:
        write_sheet.write(new_row_count, new_column_count, str(cell.value))
        new_column_count += 1
    new_row_count += 1
    return new_row_count
