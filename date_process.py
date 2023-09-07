def convert_to_month_day(value):
    # 获取月份和日期，并转换成字符串输出
    month = str(int(value[:2]))
    day = str(int(value[2:]))

    # 构造输出字符串
    output_string = f"{month}月{day}日"
    return output_string


def period_start(value):
    month = int(value[:2])
    day = int(value[2:]) - 7

    if day <= 0:
        month -= 1
        day += 30
    
    month = str(month)
    day = str(day)
    
    # 构造输出字符串
    output_string = f"{month}月{day}日"
    return output_string