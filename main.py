"""爬取 hfutxc 寝室卫生检查系统——查询宿舍床铺评分的数据"""
import time
import csv
import requests
from bs4 import BeautifulSoup

URL = "http://39.106.82.121/query/getStudentScore"  # 请求地址
GET = False  # 是否请求，True：请求数据并保存然后处理，False：读取保存的数据然后处理
DELAY = 0.1  # 每次请求间隔时间，单位秒

BUILDING = "9N"  # 寝室楼栋，1~10+N/S/#，不区分大小写
FLOOR = range(1, 7)  # 层号范围，默认为range(1, 7)
ROOM = range(1, 37)  # 房间号范围，默认为range(1, 37)

YEAR_TERM_INDEX = ((23, 1),)  # 目标学期，格式为(年,学期序号)，单个学期需在tuple后打,
WEEK_NUM = 20  # 学期的周数，默认为20


def year_term_get(date: str) -> tuple:
    """使用日期计算学期"""
    year = int(date[2:4])
    month = int(date[5:7])
    if month <= 2:
        year -= 1
        term = 1
    elif month <= 8:
        year -= 1
        term = 2
    else:
        term = 1
    return (year, term)


def dorm_req(dorm: str) -> bool:
    """请求并保存获取的数据表格为htm文件，返回请求是否有效"""
    try:
        response = requests.get(URL, params={"student_code": dorm}, timeout=10)
        response.encoding = "UTF-8"
        if response.status_code != 200:  # http状态码异常
            print(f"Request {dorm} failed! Code: {response.status_code}")
            return False
        text = response.text
        if text == '<tr><td colspan="5" align="center">无数据</td></tr>\n':  # 无数据则不保存
            return False
        with open(f"{dorm}.htm", 'w+', encoding='UTF-8') as f:
            f.write(text)
        return True
    except requests.exceptions.Timeout:  # 超时
        print(f"Request {dorm} timed out!")
        return False
    except requests.RequestException as e:  # 请求错误
        print(f"Request {dorm} failed! Error: {e}")
        return False


def dorm_dec(dorm: str) -> list:
    """解码保存的指定寝室的数据，返回指定日期以前的成绩"""
    date_index = ["-1"]*WEEK_NUM*len(YEAR_TERM_INDEX)
    try:
        with open(f"{dorm}.htm", 'r', encoding='UTF-8') as f:
            html_content = f.read()
    except FileNotFoundError:
        return date_index
    soup = BeautifulSoup(html_content, 'html.parser')
    rows = soup.find_all('tr')
    for row in rows:
        cols = row.find_all('td')
        if len(cols) != 5:  # 排除一次成绩（四行）中的非首行
            continue

        if cols[4].text != "--":  # 处理日期填入错误情况
            date = cols[4].text
        else:
            date = cols[2].text[:10]

        term = year_term_get(date)
        if term in YEAR_TERM_INDEX:  # 判断是否为目标学期
            pos = (int(cols[0].text)-1)+WEEK_NUM*YEAR_TERM_INDEX.index(term)
            date_index[pos] = cols[1].text
    return date_index


def remove_empty_weeks(table: list) -> list:
    """清理无数据周"""
    columns_to_delete = []
    for col_index in range(1, len(table[0])):
        if all(row[col_index] == "-1" for row in table[1:]):
            columns_to_delete.append(col_index)

    for row in table:
        for col_index in reversed(columns_to_delete):
            del row[col_index]

    return table


# 生成表头
head = ["寝室"]
for column in range(len(YEAR_TERM_INDEX)):
    for num in range(WEEK_NUM):
        head.append(str(num+1))
head.append("平均成绩")
output = [head]

# 处理数据
for floor in FLOOR:
    for room in ROOM:
        dorm_name = f"{BUILDING}{floor}{room:02d}"
        if GET:
            time.sleep(DELAY)
            if not dorm_req(dorm_name):  # 这一步会请求成绩
                continue
        score = dorm_dec(dorm_name)
        # 跳过空寝室
        if all(item == "-1" for item in score):
            continue
        # 计算有效平均成绩
        score_int = [int(num) for num in score if int(num) != -1]
        average = sum(score_int) / len(score_int)
        # 添加首尾列
        score.insert(0, dorm_name)
        score.append(f"{average:.2f}")
        print(",".join(score))
        output.append(score)

# 清理无数据周
output = remove_empty_weeks(output)
WEEK_COUNT = len(output[0])-2

# 从旧到新按照成绩排序
head = output.pop(0)
for i in range(0, WEEK_COUNT+1):
    output.sort(key=lambda x, i=i: float(x[i+1]), reverse=True)
output.insert(0, head)

# 添加序号列
for i, output_row in enumerate(output[1:], start=1):
    output_row.insert(0, i)
output[0].insert(0, "序号")

# 生成输出文件名
output_filename = f"{BUILDING}-"
for year_term in YEAR_TERM_INDEX:
    for item in year_term:
        output_filename += f"{str(item)}-"
output_filename += str(WEEK_COUNT)

# 保存csv
with open(f"{output_filename}.csv", 'w', newline='', encoding='GBK') as output_file:
    writer = csv.writer(output_file)
    writer.writerows(output)

print(f"Done! #validate dorm:{len(output)-1} week:{WEEK_COUNT}")
