import requests
import time
import csv
from bs4 import BeautifulSoup

URL = 'http://39.106.82.121/query/getStudentScore'  # 请求地址
GET = False  # 是否请求，True：请求数据并保存然后处理，False：读取保存的数据然后处理
DELAY = 0.1  # 每次请求间隔时间

BUILDING = "9N"  # 寝室楼栋，1~10+N/S/#
FLOOR = range(1, 7)  # 层号范围，默认为range(1, 7)
ROOM = range(1, 37)  # 房间号范围，默认为range(1, 37)

TERM_INDEX = ((2023, 1),)  # 目标学期，单个学期需在tuple后打,
WEEK_NUM = 20  # 学期的周数，默认为20


def term_get(date: str) -> tuple:
    year = int(date[0:4])
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


def dorm_req(dorm: str) -> bool:  # 请求并保存获取的数据表格为htm文件，返回请求是否有效
    try:
        response = requests.get(URL, params={'student_code': dorm})
        response.encoding = "UTF-8"
        if response.status_code == 200:
            text = response.text
            if text == '<tr><td colspan="5" align="center">无数据</td></tr>\n':
                return False
            with open(f"{dorm}.htm", 'w+') as f:
                f.write(text)
            return True
        else:
            print(f"Request {dorm} failed! Code: {response.status_code}")
    except requests.RequestException as e:
        print(f"Request {dorm} failed! Error: {e}")
    return False


def dorm_dec(dorm: str) -> list:  # 解码保存的指定寝室的数据，返回指定日期以前的成绩
    date_index = ['0']*WEEK_NUM*len(TERM_INDEX)
    try:
        with open(f"{dorm}.htm", 'r') as f:
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

        term = term_get(date)
        if term in TERM_INDEX:  # 判断是否为目标学期
            pos = (int(cols[0].text)-1)+WEEK_NUM*TERM_INDEX.index(term)
            date_index[pos] = cols[1].text
    return date_index


head = ['寝室']
for column in range(len(TERM_INDEX)):
    for num in range(WEEK_NUM):
        head.append(str(num+1))
head.append('平均成绩')
output = [head]


for floor in FLOOR:
    for room in ROOM:
        dorm = f"{BUILDING}{floor}{room:02d}"
        if GET:
            time.sleep(DELAY)
            if not dorm_req(dorm):  # 这一步会请求成绩
                continue
        score = dorm_dec(dorm)
        # 跳过空寝室
        if all(item == '0' for item in score):
            continue
        # 计算有效平均成绩
        score_int = [int(num) for num in score if int(num)]
        average = sum(score_int) / len(score_int)
        # 添加首尾列
        score.insert(0, dorm)
        score.append(f'{average:.2f}')
        print(','.join(score))
        output.append(score)

# 检查哪些列需要删除
columns_to_delete = []
for col_index in range(1, len(output[0])):
    if all(row[col_index] == '0' for row in output[1:]):
        columns_to_delete.append(col_index)

# 删除需要删除的列
for col_index in reversed(columns_to_delete):
    for row in output:
        del row[col_index]

# 保存csv
with open('output.csv', 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerows(output)

print(f"Done! #validate dorm:{len(output)-1} week:{len(output[0])-2}")
