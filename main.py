"""爬取 hfutxc 寝室卫生检查系统——查询宿舍床铺评分的数据"""
import dorm_request
import dorm_decode
from dorm_dict import DORM_DICT
from config import BUILDINGS

# 寝室楼栋，范围如下：1N, 1S, 2N, 2S, 3S, 4N, 4S, 5N, 5S, 6N, 6S, 7#, 8#, 9N, 9S, 10N, 10S
BUILDING = "9N"

FLOORS = range(1, 7)  # 层号范围，默认为range(1, 7)
ROOMS = range(1, 41)  # 房间号范围，默认为range(1, 41)

YEAR_TERM_INDEX = ((24, 1),)  # 目标学期，格式为(年,学期序号)，单个学期需在tuple后打,

EXTENSION = ["csv", "xlsx"]

"""
以下为单楼栋
"""
# 遍历求法
dorms = []
for floor in FLOORS:
    for room in ROOMS:
        dorms.append(f"{BUILDING}{floor}{room:02d}")
# 字典求法
dorms = DORM_DICT[BUILDING]

"""
以下为多楼栋
"""
# 遍历求法
dorms = []
for buingding in BUILDINGS:
    for floor in FLOORS:
        for room in ROOMS:
            dorms.append(f"{buingding}{floor}{room:02d}")
# 字典求法
dorms = []
for string_list in DORM_DICT.values():
    dorms.extend(string_list)

# 常用：单楼栋字典求法，生成csv与xlsx
dorms = DORM_DICT[BUILDING]
unsuccessful_dorms = dorm_request.dorms_req(dorms)
while unsuccessful_dorms:  # 循环获取直至无失败寝室
    unsuccessful_dorms = dorm_request.dorms_req(unsuccessful_dorms)
dorm_decode.dorms_dec(BUILDING, dorms, YEAR_TERM_INDEX, EXTENSION)

# 常用：多楼栋字典求法，同时生成每栋的csv与xlsx
# for building in BUILDINGS:
#     dorms = DORM_DICT[building]
# unsuccessful_dorms = dorm_request.dorms_req(dorms)
# while unsuccessful_dorms:  # 循环获取直至无失败寝室
#     unsuccessful_dorms = dorm_request.dorms_req(unsuccessful_dorms)
# dorm_decode.dorms_dec(building, dorms, YEAR_TERM_INDEX, EXTENSION)

# 多楼栋遍历求法
# dorms = []
# for buingding in BUILDINGS:
#     for floor in FLOORS:
#         for room in ROOMS:
#             dorms.append(f"{buingding}{floor}{room:02d}")
# unsuccessful_dorms = dorm_request.dorms_req(dorms)
# while unsuccessful_dorms:  # 循环获取直至无失败寝室
#     unsuccessful_dorms = dorm_request.dorms_req(unsuccessful_dorms)

# 多楼栋生成csv与xlsx
# for building in BUILDINGS:
#     dorms = DORM_DICT[building]
#     dorm_decode.dorms_dec(building, dorms, YEAR_TERM_INDEX, EXTENSION)
