# hfutxc-dorm-crawler

爬取 hfutxc [寝室卫生检查系统](http://39.106.82.121/query)——查询宿舍床铺评分的数据

---

## Requirements:

```bash
pip install -r requirements.txt
```

## Usage:

0. 当前寝室字典生成于 2023-11-17 ，如果发生寝室调动，导致一个曾经从未有过分数的寝室拥有分数，可以手动在 `dorm_dict.py` 内添加这个寝室，或使用 `dorm_dict_gen.py` 来自动生成寝室字典，请仅在确认需要时使用它
   - 编辑 `dorm_dict_gen.py` 首部的配置
   - run it
1. 编辑 `main.py` 首部的配置
2. run

## Todo:

1. ☑️ 联合查询多学期成绩
2. ☑️ 给表格添加表头
3. ☑️ 智能处理无分数周情况
4. ☑️ 智能处理无数据情况
5. ☑️ 优化运行速度
6. ☑️ 直接生成排版好的 csv 与 xlsx 文件
7. GUI

## Bug:

1. ☑️ 单学期的处理无效
2. ☑️ 成绩为 0 的平均数处理异常
3. ☑️ 错误的周数处理异常
4. ☑️ 未新建数据存储文件夹
