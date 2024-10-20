"""爬取 hfutxc 寝室卫生检查系统——查询宿舍床铺评分的数据"""
import dorm_request
import dorm_decode
from dorm_dict import DORM_DICT
from faculty_dict import FACULTY_DICT

"""const config"""
# 寝室楼栋
BUILDINGS = ["1N", "1S", "2N", "2S", "3N", "3S", "4N", "4S", "5N",
             "5S", "6N", "6S", "7#", "8#", "9N", "9S", "10N", "10S"]

# 院系
FACULTIES = ["生态", "文法", "英语", "城市", "电气", "计算",
             "经济", "能源", "食品", "物流", "材料", "机械"]

# 年级
GRADES = ["21", "22", "23", "24"]

# 层号范围，默认为range(1, 7)
FLOORS = range(1, 7)

# 房间号范围，默认为range(1, 41)
ROOMS = range(1, 41)
"""const config"""

# 寝室楼栋，范围如下：1N, 1S, 2N, 2S, 3N, 3S, 4N, 4S, 5N, 5S, 6N, 6S, 7#, 8#, 9N, 9S, 10N, 10S
BUILDING = "9N"

# 院系年级，例如：机械22
FACULTY = "机械22"

YEAR_TERM_INDEX = ((2024, 1),)  # 目标学期，格式为(年,学期序号)，单个学期需在tuple后打,

IF_CSV = True
IF_XLSX = True

REQ = False

# 常用：单楼栋字典，生成csv与xlsx
dorms = DORM_DICT[BUILDING]
if REQ:
    unsuccessful_dorms = dorm_request.dorms_req(dorms)
    while unsuccessful_dorms:  # 循环获取直至无失败寝室
        unsuccessful_dorms = dorm_request.dorms_req(unsuccessful_dorms)
dorm_decode.dorms_dec(BUILDING, dorms, YEAR_TERM_INDEX, IF_CSV, IF_XLSX)

# 常用：单院系年级字典，生成csv与xlsx
dorms = FACULTY_DICT[FACULTY]
if REQ:
    unsuccessful_dorms = dorm_request.dorms_req(dorms)
    while unsuccessful_dorms:  # 循环获取直至无失败寝室
        unsuccessful_dorms = dorm_request.dorms_req(unsuccessful_dorms)
dorm_decode.dorms_dec(FACULTY, dorms, YEAR_TERM_INDEX, IF_CSV, IF_XLSX)

# 遍历字典2732，以楼栋各自生成csv与xlsx
# for building in BUILDINGS:
#     dorms = DORM_DICT[building]
#     if REQ:
#         unsuccessful_dorms = dorm_request.dorms_req(dorms)
#         while unsuccessful_dorms:  # 循环获取直至无失败寝室
#             unsuccessful_dorms = dorm_request.dorms_req(unsuccessful_dorms)
#     dorm_decode.dorms_dec(building, dorms, YEAR_TERM_INDEX, IF_CSV, IF_XLSX)

# 遍历字典2732，以院系年级各自生成csv与xlsx
# for faculty in FACULTIES:
#     for grade in GRADES:
#         faculty_grade = faculty+grade
#         dorms = FACULTY_DICT[faculty_grade]
#         if REQ:
#             unsuccessful_dorms = dorm_request.dorms_req(dorms)
#             while unsuccessful_dorms:  # 循环获取直至无失败寝室
#                 unsuccessful_dorms = dorm_request.dorms_req(unsuccessful_dorms)
#         dorm_decode.dorms_dec(faculty_grade, dorms,
#                               YEAR_TERM_INDEX, IF_CSV, IF_XLSX)

# 遍历字典2732，合并生成csv与xlsx
# dorms = []
# for building in BUILDINGS:
#     dorms.extend(DORM_DICT[building])
# if REQ:
#     unsuccessful_dorms = dorm_request.dorms_req(dorms)
#     while unsuccessful_dorms:  # 循环获取直至无失败寝室
#         unsuccessful_dorms = dorm_request.dorms_req(unsuccessful_dorms)
# dorm_decode.dorms_dec("all", dorms, YEAR_TERM_INDEX, IF_CSV, IF_XLSX)

# 遍历全校4320，不生成，用于后续生成字典
# dorms = []
# for buingding in BUILDINGS:
#     for floor in FLOORS:
#         for room in ROOMS:
#             dorms.append(f"{buingding}{floor}{room:02d}")
# if REQ:
#     unsuccessful_dorms = dorm_request.dorms_req(dorms)
#     while unsuccessful_dorms:  # 循环获取直至无失败寝室
#         unsuccessful_dorms = dorm_request.dorms_req(unsuccessful_dorms)
