from datetime import datetime



def convert_to_month_day(value):
    # 将四位数日期转换成日期对象
    date_string = f"{value:04d}"
    date_object = datetime.strptime(date_string, "%m%d")

    # 获取月份和日期，并转换成字符串输出
    month = date_object.strftime("%-m")
    day = date_object.strftime("%-d")

    # 构造输出字符串
    output_string = f"{month} 月 {day} 号"
    return output_string

convert_to_month_day(int('0630'))