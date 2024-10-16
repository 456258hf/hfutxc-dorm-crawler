"""请求数据相关代码"""
import os
import time
import requests

URL = "http://39.106.82.121/query/getStudentScore"  # 请求地址
DELAY = 0.01  # 每次请求间隔时间，单位秒
TIMEOUT = 10  # 请求超时时间，单位秒


def dorm_req(dorm: str, timeout: float) -> bool:
    """请求指定寝室的成绩数据并保存为htm文件，返回请求是否有效"""
    try:
        response = requests.get(
            URL, params={"student_code": dorm}, timeout=timeout)
        response.encoding = "UTF-8"
        # http状态码异常
        if response.status_code != 200:
            print(f"Request {dorm} failed! Code: {response.status_code}")
            return False
        text = response.text
        # 无数据则不保存
        if text.startswith('<tr><td colspan="5" align="center">无数据</td></tr>\n'):
            return True
        # 匹配楼栋号
        building = dorm[:-3]
        if not os.path.exists(building):
            os.makedirs(building)
        with open(f"{building}\\{dorm}.htm", 'w+', encoding='UTF-8') as f:
            f.write(text)
        return True
    except requests.exceptions.Timeout:  # 超时
        print(f"Request {dorm} timed out!")
        return False
    except requests.RequestException as e:  # 请求错误
        print(f"Request {dorm} failed! Error: {e}")
        return False

def dorms_req(dorms: list) -> list:
    """请求指定寝室们的成绩数据并保存为htm文件，返回获取失败的寝室列表"""
    unsuccessful_dorms = []
    for dorm_index, dorm_name in enumerate(dorms, start=1):
        print(
            f"\rGetting dorm {dorm_name} , {dorm_index}/{len(dorms)}", end="")
        time.sleep(DELAY)
        if not dorm_req(dorm_name, TIMEOUT):
            unsuccessful_dorms.append(dorm_name)
    return unsuccessful_dorms
