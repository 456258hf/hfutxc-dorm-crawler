# hfutxc-dorm-crawler

爬取 hfutxc [寝室卫生检查系统](http://39.106.82.121/query)——查询宿舍床铺评分的数据
GUI 绝赞制作中

---

## Requirements:

```bash
pip install -r requirements.txt
```

## Usage:

当前寝室字典生成于 2024-10-18

```bash
py main.py
```

## Todo:

### 优化：

1. 合并两个 dict_gen
2. frame 类化
3. decode 函数包装
4. 优化学期输入方式

### 添加功能：

1. ☑️ 进度显示与结果输出采用 progressbar
2. ☑️ 完成添加 delay 与 timeout 配置
3. 完成寝室数量显示
4. 添加 decode 进度条
5. 按键添加暂停、停止
6. 添加配置保存按键
7. 添加切换后自动保存默认值功能
8. 添加生成 xlsx 后自动打开选项
9. 添加 request decode const 等配置的 GUI 编辑功能、
10. 添加最大重试次数

### 远期想法：

1. 使用数据库管理
2. 使用 PyQt 再写一份

## Bug:

1. 切换界面时不会自动停止、清除已显示内容
